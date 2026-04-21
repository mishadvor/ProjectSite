# forms_app/views/form20_view.py
"""
Форма 20: Ежедневные данные по артикулам
Отличие от Формы 4: не суммирует, а показывает ежедневные изменения по каждому артикулу
"""
import re
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, JsonResponse
from django.contrib import messages
from django.db.models import Q
from forms_app.forms import UploadFileForm, Form20DataForm
from forms_app.models import Form20Data
from openpyxl.styles import Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter
import json


# ============================================================================
# 📤 ЗАГРУЗКА ФАЙЛОВ
# ============================================================================
@login_required
def upload_file20(request):
    """Загрузка ежедневных файлов (Форма 20)"""

    # 🔥 Вспомогательная функция для извлечения даты
    def extract_date_from_filename(filename):
        match = re.search(r"(\d{2}\.\d{2}\.\d{4})", filename)
        if match:
            try:
                return datetime.strptime(match.group(1), "%d.%m.%Y").date()
            except ValueError:
                pass
        return None

    if request.method == "POST":
        form = UploadFileForm(request.POST)
        uploaded_files = request.FILES.getlist("file")

        if not uploaded_files:
            messages.error(request, "❌ Ни одного файла не было загружено.")
            return render(request, "forms_app/form20_upload.html", {"form": form})

        total_uploaded = 0
        total_skipped = 0

        for uploaded_file in uploaded_files:
            if not uploaded_file.name.lower().endswith(".xlsx"):
                messages.error(request, f"❌ {uploaded_file.name} — не .xlsx")
                total_skipped += 1
                continue

            try:
                file_data = BytesIO(uploaded_file.read())
                df_input = pd.read_excel(file_data, sheet_name=0).head(150)
            except Exception as e:
                messages.error(
                    request, f"❌ Ошибка при чтении {uploaded_file.name}: {e}"
                )
                total_skipped += 1
                continue

            required_columns = ["Код номенклатуры"]
            missing_columns = [
                col for col in required_columns if col not in df_input.columns
            ]
            if missing_columns:
                messages.error(
                    request,
                    f"❌ В файле {uploaded_file.name} отсутствуют колонки: {', '.join(missing_columns)}",
                )
                total_skipped += 1
                continue

            # 🔥 ИЗВЛЕЧЕНИЕ ДАТЫ — ИСПРАВЛЕНО
            file_date = extract_date_from_filename(uploaded_file.name)
            if file_date is None:
                file_date = datetime.now().date()
                messages.warning(
                    request,
                    f"⚠️ Не найдена дата в '{uploaded_file.name}', использована {file_date}",
                )

            new_records = []
            for idx, row in df_input.iterrows():
                code = str(row["Код номенклатуры"]).strip()
                if not code or code in {"0", "000", "000000000"}:
                    continue

                article = str(row.get("Артикул поставщика", "")).strip() or None

                def safe_float(val):
                    try:
                        return float(val) if pd.notna(val) else None
                    except (ValueError, TypeError):
                        return None

                def safe_int(val):
                    try:
                        return int(val) if pd.notna(val) else None
                    except (ValueError, TypeError):
                        return None

                new_records.append(
                    Form20Data(
                        user=request.user,
                        code=code,
                        article=article,
                        date=file_date,  # 🔥 Теперь правильная дата из имени файла!
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
                        percent_log_price=safe_float(row.get("% Лог/Наша Цена")),
                        spp_percent=safe_float(row.get("% СПП")),
                    )
                )

            created = Form20Data.objects.bulk_create(new_records, ignore_conflicts=True)
            total_uploaded += len(created)
            print(
                f"   ✅ {uploaded_file.name}: сохранено {len(created)} записей за {file_date}"
            )

        if total_uploaded:
            messages.success(
                request, f"✅ Загружено {total_uploaded} ежедневных записей."
            )
        if total_skipped:
            messages.warning(request, f"⚠️ Пропущено {total_skipped} файлов.")

        return redirect("forms_app:form20_list")

    else:
        form = UploadFileForm()

    return render(request, "forms_app/form20_upload.html", {"form": form})


# ============================================================================
# 📋 СПИСОК АРТИКУЛОВ
# ============================================================================
@login_required
def form20_list(request):
    """Список уникальных кодов (артикулов) для Формы 20"""
    # Получаем все записи пользователя, сортируем: сначала код, потом свежие даты
    queryset = Form20Data.objects.filter(user=request.user).order_by("code", "-date")

    # Собираем уникальные коды с последним артикулом
    seen_codes = {}
    for item in queryset:
        if item.code not in seen_codes:
            seen_codes[item.code] = item.article

    # Формируем список для шаблона
    codes_with_articles = [
        {"code": code, "article": article or "—"}
        for code, article in seen_codes.items()
    ]

    # Сортировка: сначала числовые коды, потом строковые
    try:
        codes_with_articles.sort(key=lambda x: int(x["code"]))
    except ValueError:
        codes_with_articles.sort(key=lambda x: x["code"])

    # Уникальные даты для фильтрации
    user_dates = (
        Form20Data.objects.filter(user=request.user)
        .values_list("date", flat=True)
        .distinct()
        .order_by("-date")
    )
    dates_list = [date.strftime("%Y-%m-%d") for date in user_dates]

    # Статистика
    total_records = Form20Data.objects.filter(user=request.user).count()
    total_codes = len(seen_codes)
    date_range = (
        f"{user_dates.last().strftime('%d.%m.%Y')} — {user_dates.first().strftime('%d.%m.%Y')}"
        if user_dates
        else "—"
    )

    return render(
        request,
        "forms_app/form20_list.html",
        {
            "codes_with_articles": codes_with_articles,
            "available_dates": dates_list,
            "form_name": "Форма 20 (Ежедневные данные)",
            "total_records": total_records,
            "total_codes": total_codes,
            "date_range": date_range,
        },
    )


# ============================================================================
# 🔍 ДЕТАЛИ ПО АРТИКУЛУ — ТАБЛИЦА С ЕЖЕДНЕВНЫМИ ИЗМЕНЕНИЯМИ
# ============================================================================
@login_required
def form20_detail(request, code):
    """
    Детальный просмотр данных по конкретному коду.
    Показывает ежедневные значения и изменения к предыдущему дню.
    """
    records = (
        Form20Data.objects.filter(user=request.user, code=code)
        .select_related("user")
        .order_by("date")
    )

    if not records.exists():
        messages.warning(request, f"Нет ежедневных данных для кода: {code}")
        return redirect("forms_app:form20_list")

    latest_record = records.first()
    article = latest_record.article if latest_record and latest_record.article else "—"

    # 🔥 Рассчитываем изменения к предыдущему дню для ключевых метрик
    records_with_changes = []
    prev_record = None

    # Поля для сравнения: (атрибут модели, человекочитаемое название, формат)
    compare_fields = [
        ("profit", "Прибыль", "{:.2f}"),
        ("clear_sales_our", "Продажи Наши", "{:.2f}"),
        ("orders", "Заказы", "{:.0f}"),
        ("percent_sell", "% Выкупа", "{:.1f}"),
        ("our_price_mid", "Цена Наша", "{:.2f}"),
        ("qentity_sale", "Продажи шт", "{:.0f}"),
    ]

    for record in records:
        row_data = {"record": record, "changes": {}}

        if prev_record:
            for field, label, fmt in compare_fields:
                curr_val = getattr(record, field)
                prev_val = getattr(prev_record, field)

                # Рассчитываем изменение только если оба значения есть
                if curr_val is not None and prev_val is not None and prev_val != 0:
                    diff = curr_val - prev_val
                    pct_change = (diff / abs(prev_val)) * 100 if prev_val != 0 else 0

                    row_data["changes"][field] = {
                        "diff": diff,
                        "pct": pct_change,
                        "diff_fmt": fmt.format(diff),
                        "pct_fmt": f"{pct_change:+.1f}%",
                        "trend": "up" if diff > 0 else "down" if diff < 0 else "same",
                        "label": label,
                    }

        records_with_changes.append(row_data)
        prev_record = (
            record  # Текущая запись становится "предыдущей" для следующей итерации
        )

    # Доступные типы графиков для этого артикула
    chart_types = [
        {"key": "profit", "label": "💰 Прибыль", "icon": "📈"},
        {"key": "sales", "label": "🛒 Продажи", "icon": "💵"},
        {"key": "orders", "label": "📦 Заказы", "icon": "📋"},
        {"key": "percent", "label": "% Выкупа", "icon": "🎯"},
        {"key": "price", "label": "🏷️ Цена", "icon": "💲"},
        {"key": "qentity_sale", "label": "📊 Продажи, шт", "icon": "🔢"},
    ]

    return render(
        request,
        "forms_app/form20_detail.html",
        {
            "records": records_with_changes,  # 🔥 С изменениями!
            "code": code,
            "article": article,
            "form_name": "Форма 20",
            "chart_types": chart_types,
        },
    )


# ============================================================================
# ✏️ РЕДАКТИРОВАНИЕ ЗАПИСИ
# ============================================================================
@login_required
def form20_edit(request, pk):
    """Редактирование отдельной записи"""
    record = get_object_or_404(Form20Data, pk=pk, user=request.user)

    if request.method == "POST":
        form = Form20DataForm(request.POST, instance=record)
        if form.is_valid():
            form.save()
            messages.success(request, "✅ Запись обновлена!")
            return redirect("forms_app:form20_detail", code=record.code)
        else:
            messages.error(request, "❌ Ошибка валидации формы")
    else:
        form = Form20DataForm(instance=record)

    return render(
        request,
        "forms_app/form20_edit.html",
        {"form": form, "record": record, "form_name": "Форма 20"},
    )


# ============================================================================
# 📈 ГРАФИК ПО АРТИКУЛУ — ЕЖЕДНЕВНЫЕ ТОЧКИ (без агрегации!)
# ============================================================================
@login_required
def form20_chart(request, code, chart_type=None):
    """
    График по конкретному артикулу.
    ВАЖНО: Показывает ежедневные значения КАК ЕСТЬ, без суммирования или усреднения.
    """
    if chart_type is None:
        chart_type = "profit"

    records = Form20Data.objects.filter(user=request.user, code=code).order_by("date")

    if not records.exists():
        messages.warning(request, f"Нет данных для построения графика по коду: {code}")
        return redirect("forms_app:form20_list")

    latest_record = records.first()
    article = latest_record.article if latest_record and latest_record.article else "—"

    # === Фильтрация по датам (через GET-параметры) ===
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")
    filtered_records = records

    if start_date:
        try:
            start_parsed = datetime.strptime(start_date, "%Y-%m-%d").date()
            filtered_records = filtered_records.filter(date__gte=start_parsed)
        except ValueError:
            start_date = None  # Сбрасываем некорректную дату

    if end_date:
        try:
            end_parsed = datetime.strptime(end_date, "%Y-%m-%d").date()
            filtered_records = filtered_records.filter(date__lte=end_parsed)
        except ValueError:
            end_date = None

    # Форматируем даты для оси X
    dates = [r.date.strftime("%d.%m.%Y") for r in filtered_records]

    # === Конфигурация типов графиков ===
    chart_configs = {
        "profit": {
            "field": "profit",
            "label": "Прибыль",
            "color": "rgb(75, 192, 192)",
            "unit": "₽",
            "decimals": 2,
        },
        "sales": {
            "field": "clear_sales_our",
            "label": "Чистые продажи Наши",
            "color": "rgb(54, 162, 235)",
            "unit": "₽",
            "decimals": 2,
        },
        "orders": {
            "field": "orders",
            "label": "Заказы",
            "color": "rgb(153, 102, 255)",
            "unit": "шт",
            "decimals": 0,
        },
        "percent": {
            "field": "percent_sell",
            "label": "% Выкупа",
            "color": "rgb(255, 159, 64)",
            "unit": "%",
            "decimals": 1,
        },
        "price": {
            "field": "our_price_mid",
            "label": "Наша цена Средняя",
            "color": "rgb(255, 99, 132)",
            "unit": "₽",
            "decimals": 2,
        },
        "log_price_percent": {
            "field": "percent_log_price",
            "label": "% Лог/Наша Цена",
            "color": "rgb(255, 205, 86)",
            "unit": "%",
            "decimals": 1,
        },
        "qentity_sale": {
            "field": "qentity_sale",
            "label": "Чистые продажи, шт",
            "color": "rgb(40, 167, 69)",
            "unit": "шт",
            "decimals": 0,
        },
        "spp_percent": {
            "field": "spp_percent",
            "label": "% СПП",
            "color": "rgb(111, 66, 193)",
            "unit": "%",
            "decimals": 1,
        },
    }

    config = chart_configs.get(chart_type, chart_configs["profit"])
    field_name = config["field"]

    # 🔥 Извлекаем значения КАК ЕСТЬ — без суммирования!
    raw_values = [getattr(r, field_name) for r in filtered_records]

    # Подготовка данных для Chart.js (замена None на 0, округление)
    data = []
    for val in raw_values:
        if val is None:
            data.append(0)
        else:
            try:
                float_val = float(val)
                if config["decimals"] == 0:
                    data.append(int(round(float_val)))
                else:
                    data.append(round(float_val, config["decimals"]))
            except (ValueError, TypeError):
                data.append(0)

    # === Статистика для отображения ===
    # Медиана
    def calc_median(values):
        filtered = [v for v in values if v is not None]
        if not filtered:
            return 0
        s = sorted(filtered)
        n = len(s)
        return s[n // 2] if n % 2 == 1 else (s[n // 2 - 1] + s[n // 2]) / 2

    median_value = calc_median(raw_values)

    # Мин/Макс/Среднее
    numeric_values = [v for v in raw_values if v is not None]
    stats = {}
    if numeric_values:
        stats = {
            "min": min(numeric_values),
            "max": max(numeric_values),
            "avg": sum(numeric_values) / len(numeric_values),
        }

    # 🔥 Сумма и средняя прибыль в день — ИСПРАВЛЕНО: правильные имена переменных
    total_profit = 0
    avg_profit_per_day = 0
    if chart_type == "profit" and numeric_values:
        total_profit = sum(numeric_values)  # ← сумма прибыли
        avg_profit_per_day = (
            total_profit / len(dates) if dates else 0
        )  # ← средняя в день

    # === Доступные типы графиков ===
    available_charts = [
        {"key": "profit", "label": "💰 Прибыль"},
        {"key": "sales", "label": "🛒 Продажи"},
        {"key": "orders", "label": "📦 Заказы"},
        {"key": "percent", "label": "% Выкупа"},
        {"key": "price", "label": "🏷️ Цена"},
        {"key": "qentity_sale", "label": "📊 Шт"},
    ]

    return render(
        request,
        "forms_app/form20_chart.html",
        {
            "code": code,
            "article": article,
            "dates": dates,
            "data": data,
            "median_value": median_value,
            "stats": stats,
            # 🔥 ИСПРАВЛЕНО: передаём именно те имена, которые ждёт шаблон
            "total_profit": total_profit,
            "avg_profit_per_day": avg_profit_per_day,
            "label": config["label"],
            "color": config["color"],
            "unit": config["unit"],
            "chart_type": chart_type,
            "available_charts": available_charts,
            "start_date": start_date,
            "end_date": end_date,
            "form_name": "Форма 20",
            "records_count": len(dates),
        },
    )


# ============================================================================
# 📥 ЭКСПОРТ В EXCEL
# ============================================================================
@login_required
def export_form20_excel(request):
    """
    Экспорт ежедневных данных в Excel.
    Каждый артикул — на отдельном листе, данные в хронологическом порядке.
    """
    data = Form20Data.objects.filter(user=request.user).order_by("code", "date")

    if not data.exists():
        messages.warning(request, "Нет данных для экспорта.")
        return redirect("forms_app:form20_list")

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
                "% Лог/Наша Цена": item.percent_log_price,
                "% СПП": item.spp_percent,
            }
        )

    # Создаём Excel в памяти
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        workbook = writer.book

        # Стиль для заголовков
        if "header_style" not in workbook.named_styles:
            header_style = NamedStyle(
                name="header_style",
                font=Font(bold=True, size=10),
                alignment=Alignment(
                    wrap_text=True, horizontal="center", vertical="center"
                ),
            )
            workbook.add_named_style(header_style)

        for code, rows in df_dict.items():
            df = pd.DataFrame(rows)
            # Название листа: код (макс. 31 символ для Excel)
            sheet_name = str(code)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]

            # Применяем стиль к заголовкам
            for cell in worksheet[1]:
                cell.style = "header_style"

            # Автоподбор ширины столбцов
            for column in worksheet.columns:
                max_length = max(
                    (len(str(cell.value)) if cell.value else 0 for cell in column),
                    default=0,
                )
                adjusted_width = min(max_length + 2, 65)  # Ограничение 65 символов
                col_letter = get_column_letter(column[0].column)
                worksheet.column_dimensions[col_letter].width = adjusted_width

    buffer.seek(0)
    filename = f"form20_daily_{request.user.username}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


# ============================================================================
# 🗑️ ОЧИСТКА ДАННЫХ
# ============================================================================
@login_required
def clear_form20_data(request):
    """Удаление ВСЕХ ежедневных данных пользователя"""
    if request.method == "POST":
        count = Form20Data.objects.filter(user=request.user).count()
        Form20Data.objects.filter(user=request.user).delete()
        messages.success(
            request, f"✅ Удалено {count} записей. Данные Формы 20 обнулены."
        )
        return redirect("forms_app:form20_list")

    # GET: показываем страницу подтверждения
    count = Form20Data.objects.filter(user=request.user).count()
    return render(
        request,
        "forms_app/form20_confirm_clear.html",
        {
            "count": count,
            "form_name": "Форма 20",
            "warning_text": "Будут удалены ВСЕ ежедневные данные. Это действие нельзя отменить!",
        },
    )


@login_required
def clear_form20_by_date(request):
    """Удаление данных за конкретную дату"""
    if request.method == "POST":
        date_str = request.POST.get("date")

        if not date_str:
            messages.error(request, "❌ Не указана дата для удаления.")
            return redirect("forms_app:form20_clear_by_date")

        try:
            date_to_delete = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messages.error(request, "❌ Неверный формат даты. Используйте ГГГГ-ММ-ДД.")
            return redirect("forms_app:form20_clear_by_date")

        deleted_count, _ = Form20Data.objects.filter(
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

        return redirect("forms_app:form20_list")

    # GET: показываем форму выбора даты
    user_dates = (
        Form20Data.objects.filter(user=request.user)
        .values_list("date", flat=True)
        .distinct()
        .order_by("-date")
    )
    dates_list = [date.strftime("%Y-%m-%d") for date in user_dates]

    return render(
        request,
        "forms_app/form20_clear_by_date.html",
        {
            "available_dates": dates_list,
            "dates_count": len(dates_list),
            "form_name": "Форма 20",
        },
    )


# ============================================================================
# 📊 ДОП: Быстрая статистика по артикулу (для AJAX)
# ============================================================================
@login_required
def form20_stats_api(request, code):
    """
    API endpoint для получения быстрой статистики по артикулу.
    Используется для динамического обновления на странице.
    """
    records = Form20Data.objects.filter(user=request.user, code=code).order_by("-date")

    if not records.exists():
        return JsonResponse({"error": "Нет данных"}, status=404)

    records_list = list(records.values())
    latest = records.first()

    # Простая статистика
    profits = [r["profit"] for r in records if r["profit"] is not None]

    response_data = {
        "code": code,
        "article": latest.article,
        "latest_date": latest.date.strftime("%d.%m.%Y") if latest.date else None,
        "latest_profit": latest.profit,
        "total_days": records.count(),
        "avg_profit": sum(profits) / len(profits) if profits else None,
        "min_profit": min(profits) if profits else None,
        "max_profit": max(profits) if profits else None,
    }

    return JsonResponse(response_data)


# ============================================================================
# 📋 ДОП: Сравнение двух дат для артикула
# ============================================================================
@login_required
def form20_compare_dates(request, code):
    """
    Сравнение показателей артикула за две выбранные даты.
    """
    if request.method == "POST":
        date1_str = request.POST.get("date1")
        date2_str = request.POST.get("date2")

        if not date1_str or not date2_str:
            messages.error(request, "❌ Укажите обе даты для сравнения")
            return redirect("forms_app:form20_detail", code=code)

        try:
            date1 = datetime.strptime(date1_str, "%Y-%m-%d").date()
            date2 = datetime.strptime(date2_str, "%Y-%m-%d").date()
        except ValueError:
            messages.error(request, "❌ Неверный формат даты")
            return redirect("forms_app:form20_detail", code=code)

        rec1 = Form20Data.objects.filter(
            user=request.user, code=code, date=date1
        ).first()
        rec2 = Form20Data.objects.filter(
            user=request.user, code=code, date=date2
        ).first()

        if not rec1 or not rec2:
            messages.warning(request, "⚠️ Нет данных для одной из выбранных дат")
            return redirect("forms_app:form20_detail", code=code)

        # Сравниваемые поля
        fields = [
            ("profit", "Прибыль", "{:.2f} ₽"),
            ("clear_sales_our", "Продажи Наши", "{:.2f} ₽"),
            ("orders", "Заказы", "{} шт"),
            ("percent_sell", "% Выкупа", "{:.1f}%"),
            ("our_price_mid", "Цена", "{:.2f} ₽"),
        ]

        comparison = []
        for field, label, fmt in fields:
            val1 = getattr(rec1, field)
            val2 = getattr(rec2, field)

            if val1 is not None and val2 is not None and val1 != 0:
                diff = val2 - val1
                pct = (diff / abs(val1)) * 100
                comparison.append(
                    {
                        "field": field,
                        "label": label,
                        "val1": (
                            fmt.format(val1) if isinstance(val1, (int, float)) else val1
                        ),
                        "val2": (
                            fmt.format(val2) if isinstance(val2, (int, float)) else val2
                        ),
                        "diff": (
                            fmt.format(diff) if isinstance(diff, (int, float)) else diff
                        ),
                        "pct": f"{pct:+.1f}%",
                        "trend": "up" if diff > 0 else "down" if diff < 0 else "same",
                    }
                )

        return render(
            request,
            "forms_app/form20_compare.html",
            {
                "code": code,
                "article": rec1.article or "—",
                "date1": date1,
                "date2": date2,
                "comparison": comparison,
                "form_name": "Форма 20",
            },
        )

    # GET: перенаправляем на детали с подсказкой
    messages.info(request, "💡 Выберите две даты в таблице для сравнения")
    return redirect("forms_app:form20_detail", code=code)
