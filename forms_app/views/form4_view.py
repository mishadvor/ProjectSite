# forms_app/views/form4_view.py

import re
import pandas as pd
from datetime import datetime
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.contrib import messages
from forms_app.forms import UploadFileForm, Form4DataForm
from forms_app.models import Form4Data  # Убедись, что модель добавлена
from django.db.models import Q
from openpyxl.styles import Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter


@login_required
def upload_file(request):
    if request.method == "POST":
        print("🔹 POST-данные:", request.POST)
        print("🔹 FILES:", request.FILES)
        print("🔹 FILES keys:", request.FILES.keys())

        # 📌 Создаём форму ТОЛЬКО с POST-данными (без FILES!)
        form = UploadFileForm(request.POST)

        # 📌 Получаем файлы вручную
        uploaded_files = request.FILES.getlist("file")
        print(f"🔹 Загружено файлов: {len(uploaded_files)}")

        # ❌ Проверяем, есть ли файлы
        if not uploaded_files:
            messages.error(request, "❌ Ни одного файла не было загружено.")
            return render(request, "forms_app/form4_upload.html", {"form": form})

        total_uploaded = 0
        total_skipped = 0

        # ✅ Обрабатываем каждый файл
        for uploaded_file in uploaded_files:
            print(f"📄 Обработка файла: {uploaded_file.name}")

            # Проверка расширения
            if not uploaded_file.name.lower().endswith(".xlsx"):
                messages.error(request, f"❌ {uploaded_file.name} — не .xlsx")
                total_skipped += 1
                continue

            try:
                file_data = BytesIO(uploaded_file.read())
                df_input = pd.read_excel(file_data, sheet_name=0).head(150)
                print(f"   ✅ Прочитано строк: {len(df_input)}")
            except Exception as e:
                print(f"   ❌ Ошибка чтения: {e}")
                messages.error(
                    request, f"❌ Ошибка при чтении {uploaded_file.name}: {e}"
                )
                total_skipped += 1
                continue

            # Проверка обязательных колонок
            required_columns = ["Код номенклатуры"]
            missing_columns = [
                col for col in required_columns if col not in df_input.columns
            ]
            if missing_columns:
                print(f"   ❌ Нет колонок: {missing_columns}")
                messages.error(
                    request,
                    f"❌ В файле {uploaded_file.name} отсутствуют колонки: {', '.join(missing_columns)}",
                )
                total_skipped += 1
                continue

            # Извлечение даты из имени файла
            match = re.search(r"(\d{2}\.\d{2}\.\d{4})\.xlsx", uploaded_file.name)
            file_date = (
                datetime.strptime(match.group(1), "%d.%m.%Y").date()
                if match
                else datetime.now().date()
            )
            print(f"   📅 Извлечена дата: {file_date}")

            # Подготовка записей
            new_records = []
            for idx, row in df_input.iterrows():
                code = str(row["Код номенклатуры"]).strip()
                if not code or code in {"0", "000", "000000000"}:
                    print(f"   ⚠️ Пропущен код: '{code}' (строка {idx})")
                    continue

                # Логируем первую валидную строку
                if len(new_records) == 0:
                    article_sample = row.get("Артикул поставщика", "")
                    print(
                        f"   ✅ Первый валидный код: {code}, Артикул: {article_sample}"
                    )

                article = str(row.get("Артикул поставщика", "")).strip() or None

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
                    Form4Data(
                        user=request.user,
                        code=code,
                        article=article,
                        date=file_date,
                        clear_sales_our=safe_float(row.get("Чистые продажи Наши")),
                        clear_sales_vb=safe_float(row.get("Чистая реализация ВБ")),
                        clear_transfer=safe_float(row.get("Чистое Перечисление")),
                        clear_transfer_without_log=safe_float(
                            row.get("Чистое Перечисление без Логистики")
                        ),
                        our_price_mid=safe_float(row.get("Наша цена Средняя")),
                        vb_selling_mid=safe_float(row.get("Реализация ВБ Средняя")),
                        transfer_mid=safe_float(row.get("К перечислению Среднее")),
                        transfer_without_log_mid=safe_float(
                            row.get("К Перечислению без Логистики Средняя")
                        ),
                        qentity_sale=safe_int(row.get("Чистые продажи, шт")),
                        sebes_sale=safe_float(row.get("Себес Продаж (600р)")),
                        profit_1=safe_float(row.get("Прибыль на 1 Юбку")),
                        percent_sell=safe_float(row.get("%Выкупа")),
                        profit=safe_float(row.get("Прибыль")),
                        orders=safe_int(row.get("Заказы")),
                        percent_log_price=safe_float(row.get("% Лог/Реализация ВБ")),
                        spp_percent=safe_float(row.get("% СПП")),
                    )
                )

            # Сохраняем в БД
            created = Form4Data.objects.bulk_create(new_records, ignore_conflicts=True)
            print(f"   ✅ Сохранено записей: {len(created)}")
            total_uploaded += len(created)

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

        # ✅ Редирект на список
        return redirect("forms_app:form4_list")

    else:
        form = UploadFileForm()

    return render(request, "forms_app/form4_upload.html", {"form": form})


# === СПИСОК ВСЕХ КОДОВ (как "листы") ===
@login_required
def form4_list(request):
    # ✅ Получаем объекты, сортируем: сначала по коду, потом свежие данные сверху
    queryset = Form4Data.objects.filter(user=request.user).order_by("code", "-date")

    seen_codes = {}
    for item in queryset:
        if item.code not in seen_codes:
            seen_codes[item.code] = item.article

    # Формируем список для шаблона
    codes_with_articles = [
        {
            "code": code,
            "article": article or "—",
        }
        for code, article in seen_codes.items()
    ]

    # Сортируем по коду
    try:
        codes_with_articles.sort(key=lambda x: int(x["code"]))
    except ValueError:
        codes_with_articles.sort(key=lambda x: x["code"])

    # Получаем уникальные даты для пользователя (для формы удаления по дате)
    user_dates = (
        Form4Data.objects.filter(user=request.user)
        .values_list("date", flat=True)
        .distinct()
        .order_by("-date")
    )

    # Преобразуем в список строк
    dates_list = [date.strftime("%Y-%m-%d") for date in user_dates]

    return render(
        request,
        "forms_app/form4_list.html",
        {
            "codes_with_articles": codes_with_articles,
            "available_dates": dates_list,  # Добавляем даты в контекст
        },
    )


# === ПРОСМОТР ДАННЫХ ПО КОНКРЕТНОМУ КОДУ ===
@login_required
def form4_detail(request, code):
    records = (
        Form4Data.objects.filter(user=request.user, code=code)
        .select_related("user")
        .order_by("date")
    )

    if not records.exists():
        messages.warning(request, f"Нет данных для кода: {code}")
        return redirect("forms_app:form4_list")

    # Берём артикул из самой свежей записи
    latest_record = records.first()
    article = latest_record.article if latest_record and latest_record.article else "—"

    return render(
        request,
        "forms_app/form4_detail.html",
        {"records": records, "code": code, "article": article},
    )


# === РЕДАКТИРОВАНИЕ ЗАПИСИ ===
@login_required
def form4_edit(request, pk):
    record = get_object_or_404(Form4Data, pk=pk, user=request.user)
    if request.method == "POST":
        form = Form4DataForm(request.POST, instance=record)
        if form.is_valid():
            form.save()
            messages.success(request, "Запись обновлена!")
            return redirect("forms_app:form4_detail", code=record.code)
    else:
        form = Form4DataForm(instance=record)
    return render(
        request, "forms_app/form4_edit.html", {"form": form, "record": record}
    )


@login_required
def export_form4_excel(request):
    data = Form4Data.objects.filter(user=request.user).order_by("code", "date")
    if not data.exists():
        messages.warning(request, "Нет данных для экспорта.")
        return redirect("forms_app:form4_list")

    # Группируем по коду
    df_dict = {}
    for item in data:
        code = item.code
        if code not in df_dict:
            df_dict[code] = []
        df_dict[code].append(
            {
                "Дата": item.date.strftime("%d.%m.%Y"),
                "Код номенклатуры": item.code,
                "Артикул": item.article or "",
                "Чистые продажи Наши": item.clear_sales_our,
                "Чистая реализация ВБ": item.clear_sales_vb,
                "Чистое Перечисление": item.clear_transfer,
                "Чистое Перечисление без Логистики": item.clear_transfer_without_log,
                "Наша цена Средняя": item.our_price_mid,
                "Реализация ВБ Средняя": item.vb_selling_mid,
                "К перечислению Среднее": item.transfer_mid,
                "К Перечислению без Логистики Средняя": item.transfer_without_log_mid,
                "Чистые продажи, шт": item.qentity_sale,
                "Себес Продаж (600р)": item.sebes_sale,
                "Прибыль на 1 Юбку": item.profit_1,
                "%Выкупа": item.percent_sell,
                "Прибыль": item.profit,
                "Заказы": item.orders,
                "% Лог/Реализация ВБ": item.percent_log_price,
                "% СПП": item.spp_percent,
            }
        )

    # Создаём Excel в памяти
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        workbook = writer.book

        # === Стиль для заголовков ===
        if "header_style" not in workbook.named_styles:
            header_style = NamedStyle(
                name="header_style",
                font=Font(bold=True),
                alignment=Alignment(
                    wrap_text=True, horizontal="center", vertical="center"
                ),
            )
            workbook.add_named_style(header_style)

        for code, rows in df_dict.items():
            df = pd.DataFrame(rows)
            sheet_name = str(code)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Получаем лист
            worksheet = writer.sheets[sheet_name]

            # Применяем стиль к первой строке (заголовкам)
            for cell in worksheet[1]:
                cell.style = "header_style"

            # Автоподбор ширины столбцов
            for column in worksheet.columns:
                max_length = max(
                    (len(str(cell.value)) if cell.value else 0 for cell in column),
                    default=0,
                )
                # Ограничиваем ширину (макс. 65 символов)
                adjusted_width = min(max_length + 2, 65)
                worksheet.column_dimensions[
                    get_column_letter(column[0].column)
                ].width = adjusted_width

    buffer.seek(0)
    filename = f"form4_data_{request.user.username}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


# === ГРАФИК ПО ПРИБЫЛИ С ФИЛЬТРОМ ПО ДАТАМ ===
@login_required
def form4_chart(request, code, chart_type=None):
    if chart_type is None:
        chart_type = "profit"

    # Получаем записи, упорядоченные по дате
    records = Form4Data.objects.filter(user=request.user, code=code).order_by("date")
    if not records.exists():
        messages.warning(request, f"Нет данных для построения графика по коду: {code}")
        return redirect("forms_app:form4_list")

    # Берём артикул из самой свежей записи
    latest_record = records.first()
    article = latest_record.article if latest_record and latest_record.article else "—"

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

    # Форматируем даты
    dates = [r.date.strftime("%d.%m.%Y") for r in records]

    # Инициализируем переменные
    data_values = []
    label = ""
    color = ""

    # Получаем данные для графика
    if chart_type == "sales":
        data_values = [float(r.clear_sales_our or 0) for r in records]
        label = "Чистые продажи Наши"
        color = "rgb(54, 162, 235)"
    elif chart_type == "orders":
        data_values = [r.orders or 0 for r in records]
        label = "Заказы"
        color = "rgb(153, 102, 255)"
    elif chart_type == "percent":
        data_values = [float(r.percent_sell or 0) for r in records]
        label = "% Выкупа"
        color = "rgb(255, 159, 64)"
    elif chart_type == "price":
        data_values = [float(r.our_price_mid or 0) for r in records]
        label = "Наша цена Средняя"
        color = "rgb(255, 99, 132)"
    elif chart_type == "log_price_percent":
        data_values = [
            float(r.percent_log_price if r.percent_log_price is not None else 0)
            for r in records
        ]
        label = "% Лог/Реализация ВБ"
        color = "rgb(255, 205, 86)"
    elif chart_type == "qentity_sale":
        data_values = [r.qentity_sale or 0 for r in records]
        label = "Чистые продажи, шт"
        color = "rgb(40, 167, 69)"
    elif chart_type == "spp_percent":
        data_values = [float(r.spp_percent or 0) for r in records]
        label = "% СПП"
        color = "rgb(111, 66, 193)"
    else:  # profit
        data_values = [float(r.profit or 0) for r in records]
        label = "Прибыль"
        color = "rgb(75, 192, 192)"

    # Вычисляем медианное значение
    def calculate_median(values):
        if not values:
            return 0
        # Фильтруем None значения
        filtered_values = [v for v in values if v is not None]
        if not filtered_values:
            return 0
        sorted_values = sorted(filtered_values)
        n = len(sorted_values)
        if n % 2 == 1:
            return sorted_values[n // 2]
        else:
            return (sorted_values[n // 2 - 1] + sorted_values[n // 2]) / 2

    median_value = calculate_median(data_values)

    # Вычисляем сумму прибыли и среднюю прибыль в день (только для графика прибыли)
    total_profit = 0
    avg_profit_per_day = 0
    if chart_type == "profit":
        total_profit = sum([float(r.profit or 0) for r in records])
        if len(dates) > 0:
            avg_profit_per_day = total_profit / len(dates)

    # Подготавливаем данные для графика (заменяем None на 0)
    data = []
    for val in data_values:
        if val is None:
            data.append(0)
        try:
            if chart_type in ["orders", "qentity_sale"]:
                # Приводим к float, округляем, затем к int — защищаемся от 14.000000000000002
                data.append(int(round(float(val))))
            else:
                data.append(round(float(val), 1))
        except (ValueError, TypeError):
            data.append(0)

    return render(
        request,
        "forms_app/form4_chart.html",
        {
            "code": code,
            "article": article,
            "dates": dates,
            "data": data,
            "median_value": median_value,
            "total_profit": total_profit,
            "avg_profit_per_day": avg_profit_per_day,  # <-- Добавляем среднюю прибыль в день
            "label": label,
            "color": color,
            "chart_type": chart_type,
            "start_date": start_date,
            "end_date": end_date,
        },
    )


# === ОБНУЛЕНИЕ ВСЕХ ДАННЫХ ФОРМЫ 4 ===
@login_required
def clear_form4_data(request):
    if request.method == "POST":
        deleted, _ = Form4Data.objects.filter(user=request.user).delete()
        messages.success(
            request, f"✅ Удалено {deleted} записей. Данные формы 4 обнулены."
        )
        return redirect("forms_app:form4_list")

    # Если GET — показываем страницу подтверждения
    return render(
        request,
        "forms_app/form4_confirm_clear.html",
        {"count": Form4Data.objects.filter(user=request.user).count()},
    )


@login_required
def clear_form4_by_date(request):
    """
    Удаление всех данных за определенную дату
    """
    if request.method == "POST":
        date_str = request.POST.get("date")

        if not date_str:
            messages.error(request, "❌ Не указана дата для удаления.")
            return redirect("forms_app:form4_clear_by_date")

        try:
            # Парсим дату из строки
            date_to_delete = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messages.error(request, "❌ Неверный формат даты. Используйте YYYY-MM-DD.")
            return redirect("forms_app:form4_clear_by_date")

        # Удаляем записи пользователя за указанную дату
        deleted_count, _ = Form4Data.objects.filter(
            user=request.user, date=date_to_delete
        ).delete()

        if deleted_count > 0:
            messages.success(
                request,
                f"✅ Удалено {deleted_count} записей за {date_to_delete.strftime('%d.%m.%Y')}",
            )
        else:
            messages.info(
                request,
                f"ℹ️ Нет данных для удаления за {date_to_delete.strftime('%d.%m.%Y')}",
            )

        return redirect("forms_app:form4_list")

    # Если GET запрос - показываем форму выбора даты
    # Получаем все уникальные даты у пользователя
    user_dates = (
        Form4Data.objects.filter(user=request.user)
        .values_list("date", flat=True)
        .distinct()
        .order_by("-date")
    )

    # Преобразуем в список для шаблона
    dates_list = [date.strftime("%Y-%m-%d") for date in user_dates]

    return render(
        request,
        "forms_app/form4_clear_by_date.html",
        {"available_dates": dates_list, "dates_count": len(dates_list)},
    )
