# forms_app/views/form8_view.py

from decimal import Decimal
import pandas as pd
import re
from datetime import datetime
from django.shortcuts import render, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required

from ..forms import Form8UploadForm
from ..models import Form8Report


@login_required
def form8_upload(request):
    if request.method == "POST":
        form = Form8UploadForm(request.POST)
        files = request.FILES.getlist("files")

        if not files:
            messages.error(request, "❌ Не выбрано ни одного файла.")
            form = Form8UploadForm()
        else:
            success_count = 0
            for f in files:
                try:
                    df = pd.read_excel(f)

                    required_cols = [
                        "Прибыль",
                        "Чистые продажи Наши",
                        "% СПП",
                        "Наша цена Средняя",
                        "Прибыль на 1 Юбку",
                        "Заказы",
                        "%Выкупа",
                    ]
                    missing = [col for col in required_cols if col not in df.columns]
                    if missing:
                        messages.warning(
                            request,
                            f"Файл '{f.name}' — нет колонок: {', '.join(missing)}",
                        )
                        continue

                    profit = Decimal(str(df["Прибыль"].sum()))
                    clean_sales = Decimal(str(df["Чистые продажи Наши"].sum()))
                    orders = int(df["Заказы"].sum())

                    spp_series = df["% СПП"][(df["% СПП"] > 0) & (df["% СПП"].notna())]
                    spp = (
                        Decimal(str(spp_series.mean())) if len(spp_series) > 0 else None
                    )

                    avg_price_series = df["Наша цена Средняя"][
                        (df["Наша цена Средняя"] > 0)
                    ]
                    avg_price = (
                        Decimal(str(avg_price_series.mean()))
                        if len(avg_price_series) > 0
                        else None
                    )

                    profit_per_skirt_series = df["Прибыль на 1 Юбку"][
                        (df["Прибыль на 1 Юбку"] > 0)
                    ]
                    profit_per_skirt = (
                        Decimal(str(profit_per_skirt_series.mean()))
                        if len(profit_per_skirt_series) > 0
                        else None
                    )

                    pickup_rate_series = df["%Выкупа"][(df["%Выкупа"] > 0)]
                    pickup_rate = (
                        Decimal(str(pickup_rate_series.mean()))
                        if len(pickup_rate_series) > 0
                        else None
                    )

                    filename = f.name
                    match = re.search(r"(\d{2}\.\d{2}\.\d{4})", filename)
                    date_extracted = None
                    if match:
                        try:
                            date_extracted = datetime.strptime(
                                match.group(1), "%d.%m.%Y"
                            ).date()
                        except ValueError:
                            pass

                    week_name = filename.replace(".xlsx", "")

                    Form8Report.objects.update_or_create(
                        user=request.user,
                        week_name=week_name,
                        defaults={
                            "date_extracted": date_extracted,
                            "profit": profit if pd.notna(profit) else None,
                            "clean_sales_ours": (
                                clean_sales if pd.notna(clean_sales) else None
                            ),
                            "spp_percent": spp,
                            "avg_price": avg_price,
                            "profit_per_skirt": profit_per_skirt,
                            "orders": orders,
                            "pickup_rate": pickup_rate,
                        },
                    )
                    success_count += 1
                except Exception as e:
                    messages.error(request, f"Ошибка при обработке {f.name}: {e}")

            if success_count > 0:
                messages.success(
                    request, f"✅ Успешно обработано: {success_count} файлов"
                )
            else:
                messages.warning(request, "❌ Ни один файл не был успешно обработан.")
            return redirect("forms_app:form8_upload")

    else:
        form = Form8UploadForm()

    # === ФИЛЬТРАЦИЯ ПО ДАТАМ ===
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    reports = Form8Report.objects.filter(user=request.user)

    if start_date:
        try:
            start_date_parsed = datetime.strptime(start_date, "%Y-%m-%d").date()
            reports = reports.filter(date_extracted__gte=start_date_parsed)
        except ValueError:
            start_date = None

    if end_date:
        try:
            end_date_parsed = datetime.strptime(end_date, "%Y-%m-%d").date()
            reports = reports.filter(date_extracted__lte=end_date_parsed)
        except ValueError:
            end_date = None

    # Сортировка
    reports = reports.order_by("date_extracted") or reports.order_by("-uploaded_at")

    # Подготовка данных для графиков
    chart_data = {
        "labels": [r.week_name for r in reports],
        "profit": [float(r.profit) if r.profit is not None else 0 for r in reports],
        "sales": [
            float(r.clean_sales_ours) if r.clean_sales_ours is not None else 0
            for r in reports
        ],
        "spp": [
            float(r.spp_percent) if r.spp_percent is not None else 0 for r in reports
        ],
        "price": [
            float(r.avg_price) if r.avg_price is not None else 0 for r in reports
        ],
        "profit_per_skirt": [
            float(r.profit_per_skirt) if r.profit_per_skirt is not None else 0
            for r in reports
        ],
        "orders": [r.orders or 0 for r in reports],
        "pickup": [
            float(r.pickup_rate) if r.pickup_rate is not None else 0 for r in reports
        ],
    }

    context = {
        "form": form,
        "reports": reports,
        "chart_data": chart_data,
        "start_date": start_date,
        "end_date": end_date,
    }

    return render(request, "forms_app/form8_upload.html", context)


@login_required
def form8_clear(request):
    if request.method == "POST":
        deleted_count = Form8Report.objects.filter(user=request.user).count()
        Form8Report.objects.filter(user=request.user).delete()
        messages.success(request, f"✅ Удалено {deleted_count} записей формы 8.")
    return redirect("forms_app:form8_upload")


@login_required
def form8_export(request):
    import pandas as pd
    from django.http import HttpResponse
    from io import BytesIO
    from datetime import timezone as datetime_timezone

    # Данные для экспорта
    reports = Form8Report.objects.filter(user=request.user).values(
        "week_name",
        "date_extracted",
        "profit",
        "clean_sales_ours",
        "spp_percent",
        "avg_price",
        "profit_per_skirt",
        "orders",
        "pickup_rate",
        "uploaded_at",
    )
    df = pd.DataFrame(reports)

    # Исправление: убираем timezone у datetime
    if "uploaded_at" in df.columns and not df.empty:
        df["uploaded_at"] = df["uploaded_at"].apply(
            lambda x: (
                x.astimezone(datetime_timezone.utc).replace(tzinfo=None)
                if x.tzinfo
                else x
            )
        )

    # Экспорт в Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Форма 8")

    output.seek(0)
    response = HttpResponse(
        output,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="form8_reports.xlsx"'
    return response
