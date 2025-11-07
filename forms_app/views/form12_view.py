# forms_app/views/form12_view.py

import re
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.contrib import messages
from forms_app.forms import UploadFileForm12, Form12DataForm
from forms_app.models import Form12Data
from openpyxl.styles import Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter


@login_required
def upload_file12(request):
    if request.method == "POST":
        print("🔹 POST-данные:", request.POST)
        print("🔹 FILES:", request.FILES)

        form = UploadFileForm12(request.POST)
        uploaded_files = request.FILES.getlist("file")
        print(f"🔹 Загружено файлов: {len(uploaded_files)}")

        if not uploaded_files:
            messages.error(request, "❌ Ни одного файла не было загружено.")
            return render(request, "forms_app/form12_upload.html", {"form": form})

        total_uploaded = 0
        total_skipped = 0

        for uploaded_file in uploaded_files:
            print(f"📄 Обработка файла: {uploaded_file.name}")

            if not uploaded_file.name.lower().endswith(".xlsx"):
                messages.error(request, f"❌ {uploaded_file.name} — не .xlsx")
                total_skipped += 1
                continue

            try:
                file_data = BytesIO(uploaded_file.read())

                # === ОБРАБОТКА КАК В ФОРМЕ 10 ===
                # Читаем исходный файл (как в Form10)
                df_raw = pd.read_excel(file_data, header=1)
                df_raw = df_raw.reset_index(drop=True)

                print(f"   ✅ Прочитано строк из исходного файла: {len(df_raw)}")
                print(f"   📊 Колонки в исходном файле: {list(df_raw.columns)}")

                # Проверяем наличие необходимых колонок в исходном файле
                required_columns = ["Артикул WB", "шт.", "Размер"]
                missing_columns = [
                    col for col in required_columns if col not in df_raw.columns
                ]

                if missing_columns:
                    print(
                        f"   ❌ Отсутствуют колонки в исходном файле: {missing_columns}"
                    )
                    messages.error(
                        request,
                        f"❌ В файле {uploaded_file.name} отсутствуют колонки: {', '.join(missing_columns)}",
                    )
                    total_skipped += 1
                    continue

                # Группируем по артикулам (как в Form10 - лист 2)
                df_processed = (
                    df_raw.groupby(
                        ["Артикул WB", "Артикул продавца"],
                        as_index=False,
                    )
                    .agg(
                        {
                            "шт.": "sum",
                            "Сумма заказов минус комиссия WB, руб.": "sum",
                            "Выкупили, шт.": "sum",
                            "К перечислению за товар, руб.": "sum",
                            "Текущий остаток, шт.": "sum",
                        }
                    )
                    .round(0)
                )

                # Переименовываем колонку как в Form10
                df_processed = df_processed.rename(columns={"шт.": "Заказы, шт."})

                print(
                    f"   ✅ Обработано записей после группировки: {len(df_processed)}"
                )

            except Exception as e:
                print(f"   ❌ Ошибка чтения/обработки: {e}")
                messages.error(
                    request, f"❌ Ошибка при обработке {uploaded_file.name}: {e}"
                )
                total_skipped += 1
                continue

            # Извлечение даты из имени файла
            match = re.search(r"(\d{4}-\d{2}-\d{2})", uploaded_file.name)
            if match:
                file_date = datetime.strptime(match.group(1), "%Y-%m-%d").date()
            else:
                # Если дата не найдена в имени, используем текущую дату
                file_date = datetime.now().date()
            print(f"   📅 Извлечена дата: {file_date}")

            # Подготовка записей для сохранения в БД
            new_records = []
            for idx, row in df_processed.iterrows():
                wb_article = str(row["Артикул WB"]).strip()
                if not wb_article or wb_article == "0":
                    print(f"   ⚠️ Пропущен Артикул WB: '{wb_article}' (строка {idx})")
                    continue

                # Логируем первую валидную строку
                if len(new_records) == 0:
                    seller_article_sample = row.get("Артикул продавца", "")
                    print(
                        f"   ✅ Первый валидный Артикул WB: {wb_article}, Артикул продавца: {seller_article_sample}"
                    )

                def safe_float(val):
                    try:
                        return float(val) if pd.notna(val) else None
                    except:
                        return None

                def safe_int(val):
                    try:
                        return int(val) if pd.notna(val) else None
                    except:
                        return None

                new_records.append(
                    Form12Data(
                        user=request.user,
                        wb_article=wb_article,
                        barcode=None,  # Не сохраняем баркод при группировке по артикулам
                        seller_article=str(row.get("Артикул продавца", "")).strip()
                        or None,
                        size=None,  # Не сохраняем размер при группировке по артикулам
                        orders_qty=safe_int(row.get("Заказы, шт.")),
                        order_amount_net=safe_float(
                            row.get("Сумма заказов минус комиссия WB, руб.")
                        ),
                        sold_qty=safe_int(row.get("Выкупили, шт.")),
                        transfer_amount=safe_float(
                            row.get("К перечислению за товар, руб.")
                        ),
                        current_stock=safe_int(row.get("Текущий остаток, шт.")),
                        date=file_date,
                    )
                )

            # Сохраняем в БД
            try:
                created = Form12Data.objects.bulk_create(new_records)
                print(f"   ✅ Сохранено записей в БД: {len(created)}")
                total_uploaded += len(created)
            except Exception as e:
                print(f"   ❌ Ошибка сохранения в БД: {e}")
                # Пробуем сохранить по одной записи
                created_count = 0
                for record in new_records:
                    try:
                        record.save()
                        created_count += 1
                    except Exception as e2:
                        print(f"      ❌ Ошибка сохранения записи: {e2}")
                print(f"   ✅ Сохранено записей (по одной): {created_count}")
                total_uploaded += created_count

        # 📢 Итоговые сообщения
        if total_uploaded:
            messages.success(
                request,
                f"✅ Успешно загружено {total_uploaded} записей из {len(uploaded_files)} файлов.",
            )
        if total_skipped:
            messages.warning(request, f"⚠️ Пропущено {total_skipped} файлов.")
        if not total_uploaded and not total_skipped:
            messages.info(
                request, "ℹ️ Файлы были, но ни одной валидной строки не найдено."
            )

        return redirect("forms_app:form12_list")

    else:
        form = UploadFileForm12()

    # Получаем количество существующих записей для отображения в шаблоне
    articles_count = Form12Data.objects.filter(user=request.user).count()

    return render(
        request,
        "forms_app/form12_upload.html",
        {"form": form, "articles_count": articles_count},
    )


# === СПИСОК ВСЕХ АРТИКУЛОВ WB ===
@login_required
def form12_list(request):
    queryset = Form12Data.objects.filter(user=request.user).order_by(
        "wb_article", "-date"
    )
    seen_articles = {}
    for item in queryset:
        if item.wb_article not in seen_articles:
            seen_articles[item.wb_article] = item.seller_article or "—"

    articles_with_seller = [
        {
            "wb_article": code,
            "seller_article": article,
        }
        for code, article in seen_articles.items()
    ]

    # Сортировка по артикулу (как строка)
    articles_with_seller.sort(key=lambda x: x["wb_article"])

    return render(
        request,
        "forms_app/form12_list.html",
        {"articles_with_seller": articles_with_seller},
    )


# === ДЕТАЛИ ПО АРТИКУЛУ WB ===
@login_required
def form12_detail(request, wb_article):
    records = (
        Form12Data.objects.filter(user=request.user, wb_article=wb_article)
        .select_related("user")
        .order_by("-date")
    )

    if not records.exists():
        messages.warning(request, f"Нет данных для артикула WB: {wb_article}")
        return redirect("forms_app:form12_list")

    # Берём артикул продавца из самой свежей записи
    latest_record = records.first()
    seller_article = (
        latest_record.seller_article
        if latest_record and latest_record.seller_article
        else "—"
    )

    # Рассчитываем статистику
    total_orders = sum(r.orders_qty or 0 for r in records)
    total_sold = sum(r.sold_qty or 0 for r in records)
    total_transfer = sum(r.transfer_amount or 0 for r in records)
    current_stock = latest_record.current_stock or 0

    return render(
        request,
        "forms_app/form12_detail.html",
        {
            "records": records,
            "wb_article": wb_article,
            "seller_article": seller_article,
            "total_orders": total_orders,
            "total_sold": total_sold,
            "total_transfer": total_transfer,
            "current_stock": current_stock,
        },
    )


# === РЕДАКТИРОВАНИЕ ЗАПИСИ ===
@login_required
def form12_edit(request, pk):
    record = get_object_or_404(Form12Data, pk=pk, user=request.user)
    if request.method == "POST":
        form = Form12DataForm(request.POST, instance=record)
        if form.is_valid():
            # Сохраняем форму, но не коммитим в БД сразу
            form_instance = form.save(commit=False)
            # Автоматически устанавливаем текущего пользователя
            form_instance.user = request.user
            form_instance.save()
            messages.success(request, "Запись обновлена!")
            return redirect("forms_app:form12_detail", wb_article=record.wb_article)
        else:
            messages.error(request, "❌ Пожалуйста, исправьте ошибки в форме.")
    else:
        form = Form12DataForm(instance=record)
        # Автоматически устанавливаем текущего пользователя в начальных данных
        form.initial["user"] = request.user

    return render(
        request, "forms_app/form12_edit.html", {"form": form, "record": record}
    )


# === ЭКСПОРТ В EXCEL ===
@login_required
def export_form12_excel(request):
    data = Form12Data.objects.filter(user=request.user).order_by("wb_article", "date")
    if not data.exists():
        messages.warning(request, "Нет данных для экспорта.")
        return redirect("forms_app:form12_list")

    # Группируем по артикулу WB
    df_dict = {}
    for item in data:
        wb_article = item.wb_article
        if wb_article not in df_dict:
            df_dict[wb_article] = []
        df_dict[wb_article].append(
            {
                "Дата": item.date.strftime("%d.%m.%Y"),
                "Артикул WB": item.wb_article,
                "Баркод": item.barcode or "",
                "Артикул продавца": item.seller_article or "",
                "Размер": item.size or "",
                "Заказы, шт.": item.orders_qty,
                "Сумма заказов минус комиссия WB, руб.": item.order_amount_net,
                "Выкупили, шт.": item.sold_qty,
                "К перечислению за товар, руб.": item.transfer_amount,
                "Текущий остаток, шт.": item.current_stock,
            }
        )

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

        for wb_article, rows in df_dict.items():
            df = pd.DataFrame(rows)
            sheet_name = str(wb_article)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            for cell in worksheet[1]:
                cell.style = "header_style"

            for column in worksheet.columns:
                max_length = max(
                    (len(str(cell.value)) if cell.value else 0 for cell in column),
                    default=0,
                )
                adjusted_width = min(max_length + 2, 65)
                worksheet.column_dimensions[
                    get_column_letter(column[0].column)
                ].width = adjusted_width

    buffer.seek(0)
    filename = f"form12_data_{request.user.username}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


# === ГРАФИК ПО АРТИКУЛУ WB ===
@login_required
def form12_chart(request, wb_article, chart_type=None):
    if chart_type is None:
        chart_type = "orders"

    records = Form12Data.objects.filter(
        user=request.user, wb_article=wb_article
    ).order_by("date")
    if not records.exists():
        messages.warning(
            request, f"Нет данных для построения графика по артикулу WB: {wb_article}"
        )
        return redirect("forms_app:form12_list")

    # Берём артикул продавца из самой свежей записи
    latest_record = records.first()
    seller_article = (
        latest_record.seller_article
        if latest_record and latest_record.seller_article
        else "—"
    )

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

    if chart_type == "orders":
        data = [r.orders_qty or 0 for r in records]
        label = "Заказы, шт."
        color = "rgb(54, 162, 235)"
    elif chart_type == "sold":
        data = [r.sold_qty or 0 for r in records]
        label = "Выкупили, шт."
        color = "rgb(255, 99, 132)"
    elif chart_type == "transfer":
        data = [round(float(r.transfer_amount or 0), 1) for r in records]
        label = "К перечислению за товар, руб."
        color = "rgb(75, 192, 192)"
    elif chart_type == "stock":
        data = [r.current_stock or 0 for r in records]
        label = "Текущий остаток, шт."
        color = "rgb(153, 102, 255)"
    else:  # default: orders
        data = [r.orders_qty or 0 for r in records]
        label = "Заказы, шт."
        color = "rgb(54, 162, 235)"

    return render(
        request,
        "forms_app/form12_chart.html",
        {
            "wb_article": wb_article,
            "seller_article": seller_article,
            "dates": dates,
            "data": data,
            "label": label,
            "color": color,
            "chart_type": chart_type,
            "start_date": start_date,
            "end_date": end_date,
        },
    )


# === ОБНУЛЕНИЕ ВСЕХ ДАННЫХ ФОРМЫ 12 ===
@login_required
def clear_form12_data(request):
    if request.method == "POST":
        deleted, _ = Form12Data.objects.filter(user=request.user).delete()
        messages.success(
            request, f"✅ Удалено {deleted} записей. Данные формы 12 обнулены."
        )
        return redirect("forms_app:form12_list")

    return render(
        request,
        "forms_app/form12_confirm_clear.html",
        {"count": Form12Data.objects.filter(user=request.user).count()},
    )


# forms_app/views/form12_view.py
@login_required
def form12_delete(request, pk):
    record = get_object_or_404(Form12Data, pk=pk, user=request.user)
    wb_article = record.wb_article

    if request.method == "POST":
        record.delete()
        messages.success(request, "✅ Запись успешно удалена!")
        return redirect("forms_app:form12_detail", wb_article=wb_article)

    return render(request, "forms_app/form12_confirm_delete.html", {"record": record})


# forms_app/views/form12_view.py
@login_required
def form12_delete_all(request):
    """Удаление ВСЕХ данных формы 12 для текущего пользователя"""
    records = Form12Data.objects.filter(user=request.user)

    if request.method == "POST":
        count = records.count()
        records.delete()
        messages.success(request, f"✅ Удалены ВСЕ данные формы 12: {count} записей!")
        return redirect("forms_app:form12_list")

    return render(
        request,
        "forms_app/form12_confirm_delete_all.html",
        {
            "records_count": records.count(),
            "articles_count": records.values("wb_article").distinct().count(),
        },
    )


@login_required
def form12_delete_article(request, wb_article):
    """Удаление всех данных по ОДНОМУ артикулу"""
    records = Form12Data.objects.filter(user=request.user, wb_article=wb_article)

    if request.method == "POST":
        count = records.count()
        records.delete()
        messages.success(
            request, f"✅ Удалено {count} записей по артикулу {wb_article}!"
        )
        return redirect("forms_app:form12_list")

    return render(
        request,
        "forms_app/form12_confirm_delete_article.html",
        {
            "wb_article": wb_article,
            "records_count": records.count(),
            "seller_article": (
                records.first().seller_article if records.exists() else "—"
            ),
        },
    )
