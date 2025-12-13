import re
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, JsonResponse
from django.contrib import messages
from forms_app.forms import UploadFileForm14
from forms_app.models import Form14Data
from openpyxl.styles import Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter


@login_required
def upload_file14(request):
    """Загрузка файлов для формы 14 - агрегация по дням"""
    if request.method == "POST":
        print("🔹 Form14: POST-данные:", request.POST)
        print("🔹 Form14: FILES:", request.FILES)

        form = UploadFileForm14(request.POST)
        uploaded_files = request.FILES.getlist("file")
        print(f"🔹 Form14: Загружено файлов: {len(uploaded_files)}")

        if not uploaded_files:
            messages.error(request, "❌ Ни одного файла не было загружено.")
            return render(request, "forms_app/form14_upload.html", {"form": form})

        total_uploaded = 0
        total_skipped = 0

        for uploaded_file in uploaded_files:
            print(f"📄 Form14: Обработка файла: {uploaded_file.name}")

            if not uploaded_file.name.lower().endswith(".xlsx"):
                messages.error(request, f"❌ {uploaded_file.name} — не .xlsx")
                total_skipped += 1
                continue

            try:
                file_data = BytesIO(uploaded_file.read())

                # Читаем исходный файл (такой же как в Form12)
                df_raw = pd.read_excel(file_data, header=1)
                df_raw = df_raw.reset_index(drop=True)

                print(
                    f"   ✅ Form14: Прочитано строк из исходного файла: {len(df_raw)}"
                )
                print(f"   📊 Form14: Колонки в исходном файле: {list(df_raw.columns)}")

                # Проверяем наличие необходимых колонок
                required_columns = [
                    "Артикул WB",
                    "шт.",
                    "Сумма заказов минус комиссия WB, руб.",
                    "Выкупили, шт.",
                    "К перечислению за товар, руб.",
                    "Текущий остаток, шт.",
                ]

                missing_columns = [
                    col for col in required_columns if col not in df_raw.columns
                ]

                if missing_columns:
                    print(f"   ❌ Form14: Отсутствуют колонки: {missing_columns}")
                    messages.error(
                        request,
                        f"❌ В файле {uploaded_file.name} отсутствуют колонки: {', '.join(missing_columns)}",
                    )
                    total_skipped += 1
                    continue

                # СУММИРУЕМ ВСЕ ЗНАЧЕНИЯ ПО ВСЕМ АРТИКУЛАМ И РАЗМЕРАМ
                # Без группировки по артикулам, просто сумма по всем строкам
                total_orders = df_raw["шт."].sum()
                total_order_amount = df_raw[
                    "Сумма заказов минус комиссия WB, руб."
                ].sum()
                total_sold = df_raw["Выкупили, шт."].sum()
                total_transfer = df_raw["К перечислению за товар, руб."].sum()
                total_stock = df_raw["Текущий остаток, шт."].sum()

                print(f"   📊 Form14: Итоговые суммы:")
                print(f"     • Заказы, шт.: {total_orders}")
                print(f"     • Сумма заказов: {total_order_amount}")
                print(f"     • Выкуплено: {total_sold}")
                print(f"     • К перечислению: {total_transfer}")
                print(f"     • Остаток: {total_stock}")

                # Извлечение даты из имени файла
                match = re.search(r"(\d{4}-\d{2}-\d{2})", uploaded_file.name)
                if match:
                    file_date = datetime.strptime(match.group(1), "%Y-%m-%d").date()
                else:
                    # Если дата не найдена, используем текущую
                    file_date = datetime.now().date()
                print(f"   📅 Form14: Извлечена дата: {file_date}")

                # Проверяем, существует ли уже запись за эту дату
                existing_record = Form14Data.objects.filter(
                    user=request.user, date=file_date
                ).first()

                if existing_record:
                    # Обновляем существующую запись
                    existing_record.total_orders_qty = int(total_orders)
                    existing_record.total_order_amount_net = float(total_order_amount)
                    existing_record.total_sold_qty = int(total_sold)
                    existing_record.total_transfer_amount = float(total_transfer)
                    existing_record.total_current_stock = int(total_stock)
                    existing_record.save()
                    print(f"   🔄 Form14: Обновлена запись за {file_date}")
                    total_uploaded += 1
                else:
                    # Создаем новую запись
                    new_record = Form14Data(
                        user=request.user,
                        date=file_date,
                        total_orders_qty=int(total_orders),
                        total_order_amount_net=float(total_order_amount),
                        total_sold_qty=int(total_sold),
                        total_transfer_amount=float(total_transfer),
                        total_current_stock=int(total_stock),
                    )
                    new_record.save()
                    print(f"   ✅ Form14: Создана новая запись за {file_date}")
                    total_uploaded += 1

            except Exception as e:
                print(f"   ❌ Form14: Ошибка обработки: {e}")
                messages.error(
                    request, f"❌ Ошибка при обработке {uploaded_file.name}: {e}"
                )
                total_skipped += 1
                continue

        # 📢 Итоговые сообщения
        if total_uploaded:
            messages.success(
                request,
                f"✅ Form14: Успешно обработано {total_uploaded} файлов.",
            )
        if total_skipped:
            messages.warning(request, f"⚠️ Form14: Пропущено {total_skipped} файлов.")

        return redirect("forms_app:form14_list")

    else:
        form = UploadFileForm14()

    # Получаем количество существующих записей для отображения
    records_count = Form14Data.objects.filter(user=request.user).count()

    return render(
        request,
        "forms_app/form14_upload.html",
        {"form": form, "records_count": records_count},
    )


@login_required
def form14_list(request):
    """Список всех дней с агрегированными данными"""
    records = Form14Data.objects.filter(user=request.user).order_by("-date")

    # Рассчитываем общие итоги
    total_stats = {
        "total_orders": sum(r.total_orders_qty or 0 for r in records),
        "total_order_amount": sum(r.total_order_amount_net or 0 for r in records),
        "total_sold": sum(r.total_sold_qty or 0 for r in records),
        "total_transfer": sum(r.total_transfer_amount or 0 for r in records),
        "current_stock": records.first().total_current_stock if records.exists() else 0,
    }

    return render(
        request,
        "forms_app/form14_list.html",
        {
            "records": records,
            "total_stats": total_stats,
        },
    )


@login_required
def form14_chart(request, chart_type=None):
    """График агрегированных данных по дням"""
    if chart_type is None:
        chart_type = "orders"

    records = Form14Data.objects.filter(user=request.user).order_by("date")

    if not records.exists():
        messages.warning(request, "Нет данных для построения графика.")
        return redirect("forms_app:form14_list")

    # === Фильтрация по датам ===
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    if start_date:
        try:
            start_date_parsed = datetime.strptime(start_date, "%Y-%m-%d").date()
            records = records.filter(date__gte=start_date_parsed)
        except ValueError:
            start_date = None

    if end_date:
        try:
            end_date_parsed = datetime.strptime(end_date, "%Y-%m-%d").date()
            records = records.filter(date__lte=end_date_parsed)
        except ValueError:
            end_date = None

    # Форматируем даты и данные
    dates = [r.date.strftime("%d.%m.%Y") for r in records]

    # Выбираем данные в зависимости от типа графика
    if chart_type == "orders":
        data = [r.total_orders_qty or 0 for r in records]
        label = "Общие заказы, шт."
        color = "rgb(54, 162, 235)"
        y_axis_label = "Количество, шт."
    elif chart_type == "order_amount":
        data = [round(float(r.total_order_amount_net or 0), 1) for r in records]
        label = "Общая сумма заказов (минус комиссия), руб."
        color = "rgb(255, 159, 64)"
        y_axis_label = "Сумма, руб."
    elif chart_type == "sold":
        data = [r.total_sold_qty or 0 for r in records]
        label = "Всего выкуплено, шт."
        color = "rgb(255, 99, 132)"
        y_axis_label = "Количество, шт."
    elif chart_type == "transfer":
        data = [round(float(r.total_transfer_amount or 0), 1) for r in records]
        label = "Общая сумма к перечислению, руб."
        color = "rgb(75, 192, 192)"
        y_axis_label = "Сумма, руб."
    elif chart_type == "stock":
        data = [r.total_current_stock or 0 for r in records]
        label = "Общий остаток на складе, шт."
        color = "rgb(153, 102, 255)"
        y_axis_label = "Количество, шт."
    else:  # default: orders
        data = [r.total_orders_qty or 0 for r in records]
        label = "Общие заказы, шт."
        color = "rgb(54, 162, 235)"
        y_axis_label = "Количество, шт."

    return render(
        request,
        "forms_app/form14_chart.html",
        {
            "dates": dates,
            "data": data,
            "label": label,
            "color": color,
            "chart_type": chart_type,
            "y_axis_label": y_axis_label,
            "start_date": start_date,
            "end_date": end_date,
            "total_records": records.count(),
        },
    )


@login_required
def clear_form14_data(request):
    """Очистка всех данных формы 14"""
    if request.method == "POST":
        deleted, _ = Form14Data.objects.filter(user=request.user).delete()
        messages.success(
            request, f"✅ Удалено {deleted} записей. Данные формы 14 обнулены."
        )
        return redirect("forms_app:form14_list")

    return render(
        request,
        "forms_app/form14_confirm_clear.html",
        {"count": Form14Data.objects.filter(user=request.user).count()},
    )


@login_required
def form14_delete_by_date(request):
    """Удаление данных за определенную дату"""
    if request.method == "POST":
        date_str = request.POST.get("date")
        if not date_str:
            messages.error(request, "❌ Не указана дата для удаления.")
            return redirect("forms_app:form14_list")

        try:
            delete_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messages.error(request, "❌ Неверный формат даты. Используйте ГГГГ-ММ-ДД.")
            return redirect("forms_app:form14_list")

        # Удаляем запись за указанную дату
        deleted_count = Form14Data.objects.filter(
            user=request.user, date=delete_date
        ).delete()[0]

        if deleted_count:
            messages.success(
                request,
                f"✅ Удалены данные за {delete_date.strftime('%d.%m.%Y')}",
            )
        else:
            messages.warning(
                request,
                f"ℹ️ Нет данных для удаления за {delete_date.strftime('%d.%m.%Y')}",
            )

        return redirect("forms_app:form14_list")

    # GET запрос - показываем форму выбора даты
    available_dates = (
        Form14Data.objects.filter(user=request.user)
        .values_list("date", flat=True)
        .distinct()
        .order_by("-date")
    )

    return render(
        request,
        "forms_app/form14_delete_by_date.html",
        {
            "available_dates": available_dates,
            "records_count": Form14Data.objects.filter(user=request.user).count(),
        },
    )


@login_required
def export_form14_excel(request):
    """Экспорт данных формы 14 в Excel"""
    data = Form14Data.objects.filter(user=request.user).order_by("-date")
    if not data.exists():
        messages.warning(request, "Нет данных для экспорта.")
        return redirect("forms_app:form14_list")

    # Создаем DataFrame
    rows = []
    for item in data:
        rows.append(
            {
                "Дата": item.date.strftime("%d.%m.%Y"),
                "Общие заказы, шт.": item.total_orders_qty,
                "Общая сумма заказов (минус комиссия), руб.": item.total_order_amount_net,
                "Всего выкуплено, шт.": item.total_sold_qty,
                "Общая сумма к перечислению, руб.": item.total_transfer_amount,
                "Общий остаток на складе, шт.": item.total_current_stock,
            }
        )

    df = pd.DataFrame(rows)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        workbook = writer.book

        # Стиль для заголовков
        if "header_style" not in workbook.named_styles:
            header_style = NamedStyle(
                name="header_style",
                font=Font(bold=True),
                alignment=Alignment(
                    wrap_text=True, horizontal="center", vertical="center"
                ),
            )
            workbook.add_named_style(header_style)

        sheet_name = "Form14_Агрегированные_данные"
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        worksheet = writer.sheets[sheet_name]
        for cell in worksheet[1]:
            cell.style = "header_style"

        # Автоподбор ширины колонок
        for column in worksheet.columns:
            max_length = max(
                (len(str(cell.value)) if cell.value else 0 for cell in column),
                default=0,
            )
            adjusted_width = min(max_length + 2, 65)
            worksheet.column_dimensions[get_column_letter(column[0].column)].width = (
                adjusted_width
            )

    buffer.seek(0)
    filename = f"form14_data_{request.user.username}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


@login_required
def form14_api_data(request, chart_type):
    """API для получения данных для графиков (для AJAX запросов)"""
    records = Form14Data.objects.filter(user=request.user).order_by("date")

    dates = [r.date.strftime("%d.%m.%Y") for r in records]

    if chart_type == "orders":
        data = [r.total_orders_qty or 0 for r in records]
        label = "Общие заказы, шт."
    elif chart_type == "order_amount":
        data = [round(float(r.total_order_amount_net or 0), 1) for r in records]
        label = "Общая сумма заказов, руб."
    elif chart_type == "sold":
        data = [r.total_sold_qty or 0 for r in records]
        label = "Всего выкуплено, шт."
    elif chart_type == "transfer":
        data = [round(float(r.total_transfer_amount or 0), 1) for r in records]
        label = "Общая сумма к перечислению, руб."
    elif chart_type == "stock":
        data = [r.total_current_stock or 0 for r in records]
        label = "Общий остаток на складе, шт."
    else:
        return JsonResponse({"error": "Invalid chart type"}, status=400)

    return JsonResponse({"dates": dates, "data": data, "label": label})
