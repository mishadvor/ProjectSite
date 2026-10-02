# forms_app/views/form21_view.py
import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
import re

# Значения по умолчанию (если поля на форме не заполнены)
UNIT_COST_DEFAULT = 900          # руб. за 1 выкуп
TAX_RATE_DEFAULT = 0.08          # 8%

# Имя листа в файле себестоимости
SEBES_SHEET = "Sebes"
SEBES_PREFIX_COL = "Префикс_группы"
SEBES_COST_COL = "Себестоимость"

# Типы начислений, которые относятся к общескладским расходам Ozon
# (не привязаны ни к одному артикулу, идут со знаком минус)
SERVICE_TYPES = [
    "Страхование товара от массовых повреждений",
    "Подписка Premium",
    "Ускоренный сбор отзывов",
    "Отгрузка в нерекомендованный слот",
]

# Типы начислений, которые относятся к логистике (включая основную "Логистика")
LOGISTICS_TYPES = [
    "Логистика",
    "Дополнительная упаковка товара на складе",
    "Доставка до места выдачи",
    "Доставка до места выдачи силами Ozon",
    "Обеспечение материалами для упаковки товара",
    "Обработка возвратов, отмен и невыкупов партнёрами",
    "Обработка отправления Drop-off партнёрами (ПВЗ)",
    "Обратная логистика",
    "Упаковка товара партнёрами",
]

# Типы начислений, которые являются бонусами от Ozon к выручке
# (включаются в расчёт "полной цены продажи" для процентных показателей)
REVENUE_BONUS_TYPES = [
    "Баллы за скидки",
    "Программы партнёров",
]


def extract_prefix(article):
    if pd.isna(article) or article == "":
        return "unknown"
    parts = str(article).split("_")
    if len(parts) >= 1:
        return parts[0]
    return str(article)[:3]


def calculate_purchase_percentage(revenue_count, logistics_count):
    if logistics_count == 0:
        return 0.0
    return round((revenue_count / logistics_count) * 100, 1)


def extract_date_range(filename):
    if not filename:
        return None
    match = re.search(r"(\d{2}\.\d{2}\.\d{4}-\d{2}\.\d{2}\.\d{4})", filename)
    if match:
        return match.group(1)
    return None


def load_cost_map(sebes_file):
    """Читает xlsx с листом 'Sebes' и колонками 'Префикс_группы', 'Себестоимость'."""
    if not sebes_file:
        return {}

    xls = pd.ExcelFile(sebes_file)
    sheet = None
    if SEBES_SHEET in xls.sheet_names:
        sheet = SEBES_SHEET
    else:
        for name in xls.sheet_names:
            if str(name).strip().lower().startswith("sebes"):
                sheet = name
                break
        if sheet is None:
            sheet = xls.sheet_names[0]

    df = pd.read_excel(xls, sheet_name=sheet)

    def find_col(df, target):
        target_norm = target.strip().lower().replace("_", " ")
        for c in df.columns:
            if str(c).strip().lower().replace("_", " ") == target_norm:
                return c
        return None

    col_prefix = find_col(df, SEBES_PREFIX_COL)
    col_cost = find_col(df, SEBES_COST_COL)

    if col_prefix is None or col_cost is None:
        raise ValueError(
            f"В файле себестоимости должны быть колонки "
            f"'{SEBES_PREFIX_COL}' и '{SEBES_COST_COL}'. "
            f"Найдены: {list(df.columns)}"
        )

    df = df.dropna(subset=[col_prefix, col_cost])
    df[col_prefix] = df[col_prefix].astype(str).str.strip()
    df[col_cost] = pd.to_numeric(df[col_cost], errors="coerce")

    return dict(zip(df[col_prefix], df[col_cost]))


def parse_float(value, default):
    """Пытается превратить значение из формы в float. Если не получилось — default."""
    if value is None:
        return default
    try:
        cleaned = str(value).strip().replace(" ", "").replace(",", ".")
        if cleaned == "":
            return default
        return float(cleaned)
    except (TypeError, ValueError):
        return default


def form21(request):
    if request.method == "POST":
        excel_file = request.FILES.get("excel_file")
        sebes_file = request.FILES.get("sebes_file")

        # --- Ручные параметры из формы ---
        unit_cost_default = parse_float(
            request.POST.get("unit_cost"), UNIT_COST_DEFAULT
        )
        tax_rate_percent = parse_float(
            request.POST.get("tax_rate"), TAX_RATE_DEFAULT * 100
        )
        tax_rate = tax_rate_percent / 100.0

        if not excel_file:
            return render(
                request,
                "forms_app/form21.html",
                {
                    "error": "Пожалуйста, выберите файл отчёта продаж.",
                    "unit_cost_default": unit_cost_default,
                    "tax_rate_percent": tax_rate_percent,
                },
            )

        try:
            # --- Себестоимость из файла (если есть) ---
            cost_map = {}
            if sebes_file:
                try:
                    cost_map = load_cost_map(sebes_file)
                except Exception as e:
                    return render(
                        request,
                        "forms_app/form21.html",
                        {
                            "error": f"Ошибка чтения файла себестоимости: {e}",
                            "unit_cost_default": unit_cost_default,
                            "tax_rate_percent": tax_rate_percent,
                        },
                    )

            # --- Продажи ---
            df = pd.read_excel(excel_file, skiprows=1, header=0)
            df["Префикс_артикула"] = df["Артикул"].apply(extract_prefix)

            ad_types = ["Оплата за клик"]
            df["Реклама"] = df["Тип начисления"].isin(ad_types).astype(int)
            df_ad = df[df["Реклама"] == 1].copy()
            df_non_ad = df[df["Реклама"] == 0].copy()

            total_ad_cost = df_ad["Сумма итого, руб."].sum() if len(df_ad) > 0 else 0

            # ---------- ОБЩИЕ РАСХОДЫ OZON (вне групп артикулов) ----------
            # Это строки без артикула (или с префиксом "unknown"):
            # Страхование, Premium, Ускоренный сбор отзывов, Отгрузка в нерекомендованный слот.
            # Они уже входят в исходную "Общую сумму", но не относятся ни к одной группе.
            service_total = 0.0
            for stype in SERVICE_TYPES:
                mask = df_non_ad["Тип начисления"] == stype
                if mask.any():
                    service_total += float(df_non_ad.loc[mask, "Сумма итого, руб."].sum())
            service_total = round(service_total, 2)

            # ---------- Группировка по артикулам ----------
            detailed_stats = []
            for article in df_non_ad["Артикул"].unique():
                a = df_non_ad[df_non_ad["Артикул"] == article]
                revenue_count = len(a[a["Тип начисления"] == "Выручка"])
                logistics_count = len(a[a["Тип начисления"] == "Логистика"])
                detailed_stats.append({
                    "Артикул": article,
                    "Префикс": extract_prefix(article),
                    "Общая сумма, руб": a["Сумма итого, руб."].sum(),
                    "Выручка, руб": a[a["Тип начисления"] == "Выручка"]["Сумма итого, руб."].sum(),
                    "Логистика, руб": a[a["Тип начисления"] == "Логистика"]["Сумма итого, руб."].sum(),
                    "Количество выкупов": revenue_count,
                    "Количество заказов": logistics_count,
                    "Процент выкупа, %": calculate_purchase_percentage(revenue_count, logistics_count),
                })
            detailed_df = pd.DataFrame(detailed_stats).sort_values("Общая сумма, руб", ascending=False)

            # Убираем из детального отчёта служебные строки без артикула / с unknown
            detailed_df = detailed_df[
                detailed_df["Артикул"].astype(str).str.strip().ne("") &
                detailed_df["Артикул"].astype(str).str.strip().str.lower().ne("nan") &
                detailed_df["Артикул"].astype(str).str.strip().ne("unknown")
            ].copy()

            # ---------- Группировка по префиксам ----------
            group_stats = []
            for prefix in df_non_ad["Префикс_артикула"].unique():
                g = df_non_ad[df_non_ad["Префикс_артикула"] == prefix]
                revenue_count = len(g[g["Тип начисления"] == "Выручка"])
                logistics_count = len(g[g["Тип начисления"] == "Логистика"])
                group_stats.append({
                    "Префикс_группы": prefix,
                    "Общая сумма, руб": g["Сумма итого, руб."].sum(),
                    "Выручка, руб": g[g["Тип начисления"] == "Выручка"]["Сумма итого, руб."].sum(),
                    "Логистика, руб": g[g["Тип начисления"] == "Логистика"]["Сумма итого, руб."].sum(),
                    "Количество выкупов": revenue_count,
                    "Количество заказов": logistics_count,
                    "Процент выкупа, %": calculate_purchase_percentage(revenue_count, logistics_count),
                    "Количество артикулов в группе": g["Артикул"].nunique(),
                })
            group_df_result = pd.DataFrame(group_stats).sort_values("Общая сумма, руб", ascending=False)

            prefix_pivot = pd.pivot_table(
                df_non_ad,
                values="Сумма итого, руб.",
                index="Префикс_артикула",
                columns="Тип начисления",
                aggfunc="sum",
                fill_value=0,
            )
            prefix_pivot_reset = prefix_pivot.reset_index().rename(
                columns={"Префикс_артикула": "Префикс_группы"}
            )
            merged_df = pd.merge(group_df_result, prefix_pivot_reset, on="Префикс_группы", how="left")

            total_revenue = merged_df["Выручка, руб"].sum()
            if total_revenue > 0 and total_ad_cost != 0:
                merged_df["Рекламные расходы, руб"] = (
                    merged_df["Выручка, руб"] / total_revenue * total_ad_cost
                ).round(2)
                merged_df["Чистая прибыль, руб"] = (
                    merged_df["Общая сумма, руб"] + merged_df["Рекламные расходы, руб"]
                ).round(2)
            else:
                merged_df["Рекламные расходы, руб"] = 0
                merged_df["Чистая прибыль, руб"] = merged_df["Общая сумма, руб"]

            # ===== УДАЛЯЕМ служебную группу "unknown" из отчёта по группам =====
            # Это не группа артикулов, а общескладские расходы Ozon
            # (Страхование, Подписка Premium, Ускоренный сбор отзывов, Отгрузка в нерекомендованный слот).
            # Они уже отражены отдельными строками в "0_Финансовая_сводка" и учтены
            # в строке "Общие расходы Ozon (вне групп артикулов)".
            # Чтобы не засорять лист "1_Группы_объединенная" и не искажать
            # "Положительную" и "Отрицательную маржу" — исключаем её.
            merged_df = merged_df[merged_df["Префикс_группы"] != "unknown"].copy()
            merged_df = merged_df[merged_df["Префикс_группы"].astype(str).str.strip() != ""].copy()

            # ---------- Себестоимость: из файла, иначе — значение из формы ----------
            merged_df["Себестоимость"] = (
                merged_df["Префикс_группы"].astype(str).str.strip()
                .map(cost_map)
                .fillna(unit_cost_default)
                .astype(float)
            )
            used_fallback = (
                merged_df["Префикс_группы"].astype(str).str.strip().isin(cost_map.keys()) == False
            )
            fallback_groups = merged_df.loc[used_fallback, "Префикс_группы"].tolist()

            merged_df["Кол-во*себес"] = (
                merged_df["Количество выкупов"] * merged_df["Себестоимость"]
            ).round(2)
            merged_df["Чистая прибыль - себес"] = (
                merged_df["Чистая прибыль, руб"] - merged_df["Кол-во*себес"]
            ).round(2)

            # ---------- Налог (ставка из формы) ----------
            merged_df["Налог"] = (merged_df["Выручка, руб"] * tax_rate).round(2)

            merged_df["Маржа"] = (
                merged_df["Чистая прибыль - себес"] - merged_df["Налог"]
            ).round(2)

            # Учитываем все логистические расходы (расширенная логистика)
            def calc_log_pct(row):
                revenue = row["Выручка, руб"]
                balli = row.get("Баллы за скидки", 0)
                if pd.isna(balli):
                    balli = 0
                prog = row.get("Программы партнёров", 0)
                if pd.isna(prog):
                    prog = 0
                
                # Полная цена продажи = Выручка + Баллы + Программы партнёров
                full_price = revenue + balli + prog
                
                # ⚠️ ВАЖНО: вычисляем сумму всех логистических расходов по группе
                logistics_sum = 0.0
                for log_type in LOGISTICS_TYPES:
                    if log_type in row.index:
                        val = row[log_type]
                        if not pd.isna(val):
                            logistics_sum += abs(float(val))
                
                if full_price > 0:
                    return round(logistics_sum / full_price * 100, 1)
                return 100.0

            merged_df["% Лог/(Выручка+Баллы)"] = merged_df.apply(calc_log_pct, axis=1)

            merged_df["Средняя цена выкупа"] = merged_df.apply(
                lambda row: round(row["Выручка, руб"] / row["Количество выкупов"], 2)
                if row["Количество выкупов"] > 0 else 0.0,
                axis=1,
            )

            # ---------- Порядок колонок ----------
            base_columns = [
                "Префикс_группы",
                "Выручка, руб",
                "Общая сумма, руб",
                "Чистая прибыль, руб",
                "Себестоимость",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог",
                "Маржа",
                "Логистика, руб",
                "% Лог/(Выручка+Баллы)",
                "Количество выкупов",
                "Количество заказов",
                "Процент выкупа, %",
                "Средняя цена выкупа",
                "Количество артикулов в группе",
                "Рекламные расходы, руб",
            ]
            other_columns = [
                col for col in merged_df.columns
                if col not in base_columns and col != "Префикс_группы"
            ]
            merged_df = merged_df[[c for c in base_columns if c in merged_df.columns] + other_columns]
            merged_df = merged_df.sort_values("Маржа", ascending=False)

            # ================= УДАЛЕНИЕ ДУБЛЕЙ КОЛОНОК =================
            # После merge в merged_df могут быть колонки "Выручка" и "Выручка, руб"
            # (а также "Логистика" и "Логистика, руб"). Удаляем версии без ", руб"
            cols_to_drop = []
            for col in merged_df.columns:
                # Если есть колонка с таким же именем + ", руб" — удаляем текущую
                if f"{col}, руб" in merged_df.columns:
                    cols_to_drop.append(col)
            
            if cols_to_drop:
                merged_df = merged_df.drop(columns=cols_to_drop, errors='ignore')
                detailed_df = detailed_df.drop(columns=[c for c in cols_to_drop if c in detailed_df.columns], errors='ignore')

            # ================= ФИНАНСОВАЯ СВОДКА =================
            
            # 1. Сначала считаем абсолютные суммы по всему исходному файлу
            total_credited = round(df[df["Сумма итого, руб."] > 0]["Сумма итого, руб."].sum(), 2)
            total_deducted = round(df[df["Сумма итого, руб."] < 0]["Сумма итого, руб."].sum(), 2)

            formulas = {
                "Начислено": "Сумма всех положительных значений в колонке 'Сумма итого, руб.' исходного файла",
                "Удержано": "Сумма всех отрицательных значений в колонке 'Сумма итого, руб.' исходного файла",
                "Общая сумма, руб": "Сумма всех операций (Выручка - Логистика - Прочие начисления по каждому Артикулу)",
                "Выручка, руб": "Сумма операций с типом 'Выручка'",
                "Логистика, руб": (
                    "Сумма всех логистических расходов: "
                    "Логистика + Обратная логистика + Доставка до места выдачи + "
                    "Доставка до места выдачи силами Ozon + Обработка возвратов, отмен и невыкупов партнёрами + "
                    "Обработка отправления Drop-off партнёрами (ПВЗ) + Упаковка товара партнёрами + "
                    "Обеспечение материалами для упаковки товара + Дополнительная упаковка товара на складе"
                ),
                "Количество выкупов": "Количество операций с типом 'Выручка'",
                "Количество заказов": "Количество операций с типом 'Логистика'",
                "Количество артикулов в группе": "Количество уникальных артикулов в группе",
                "Рекламные расходы, руб": "Расходы на рекламу (тип 'Оплата за клик'), распределённые пропорционально выручке",
                "Чистая прибыль, руб": "Общая сумма - Общие расходы - Рекламные расходы",
                "Себестоимость": "Себестоимость за 1 выкуп: из Sebes.xlsx по префиксу, иначе — значение с формы",
                "Кол-во*себес": "Количество выкупов × Себестоимость группы",
                "Чистая прибыль - себес": "Чистая прибыль, руб - Кол-во*себес",
                "Налог": f"Выручка, руб × {tax_rate_percent:.2f}%",
                "Маржа": "Чистая прибыль - себес - Налог (по группам артикулов, без общих расходов)",
            }

            

            # 2. Инициализируем список СРАЗУ с двумя новыми строками ПЕРВЫМИ
            summary_data = [
                {
                    "Показатель": "Начислено",
                    "Итог": total_credited,
                    "Тип начисления в расчете": formulas["Начислено"]
                },
                {
                    "Показатель": "Удержано",
                    "Итог": total_deducted,
                    "Тип начисления в расчете": formulas["Удержано"]
                }
            ]

            numeric_columns = [
                "Выручка, руб",
                "Общая сумма, руб",
                "Логистика, руб",
                "Количество выкупов",
                "Количество заказов",
                "Количество артикулов в группе",
                "Рекламные расходы, руб",
                "Чистая прибыль, руб",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог",
                "Маржа",
            ]

            # 3. СНАЧАЛА: собираем все numeric_columns из pivot
            for col in prefix_pivot_reset.columns:
                if col in ["Префикс_группы"]:
                    continue
                
                # Пропускаем, если такая колонка уже есть в списке
                if col in numeric_columns:
                    continue
                
                # Пропускаем, если есть вариант с ", руб" на конце
                if f"{col}, руб" in numeric_columns:
                    continue
                
                # Пропускаем типы, которые уже учтены в "Общескладских расходах Ozon"
                if col in SERVICE_TYPES:
                    continue
                
                numeric_columns.append(col)
                formulas[col] = f"Сумма операций с типом '{col}'"

            # 4. ПОТОМ: ОТДЕЛЬНЫМ циклом обрабатываем все numeric_columns 
            # (они добавятся в список ПОСЛЕ "Начислено" и "Удержано")
            for col in numeric_columns:
                if col == "Логистика, руб":
                    # Расширенная логистика: сумма всех логистических типов из исходного файла
                    log_mask = df_non_ad["Тип начисления"].isin(LOGISTICS_TYPES)
                    col_sum = round(df_non_ad.loc[log_mask, "Сумма итого, руб."].sum(), 2)
                elif col == "Общая сумма, руб":
                    # Берем сумму по всем операциям без рекламы (включая "unknown")
                    col_sum = round(df_non_ad["Сумма итого, руб."].sum(), 2)
                elif col == "Чистая прибыль, руб":
                    # Прибавляем общескладские расходы к прибыли по группам
                    col_sum = round(merged_df[col].sum() + service_total, 2)
                elif col in merged_df.columns:
                    # Для остальных показателей (выручка, маржа и т.д.) 
                    # сумма по группам и так корректна
                    col_sum = round(merged_df[col].sum(), 2)
                else:
                    # Если колонки нет в merged_df (например, доп. типы из pivot)
                    col_sum = 0.0

                summary_data.append({
                    "Показатель": col,
                    "Итог": col_sum,
                    "Тип начисления в расчете": formulas.get(
                        col, "Сумма всех операций по данному типу"
                    ),
                })

            financial_summary = pd.DataFrame(summary_data)

            # ============= ПОЛОЖИТЕЛЬНАЯ / ОТРИЦАТЕЛЬНАЯ МАРЖА (по группам) =============

            margin_all = round(float(merged_df["Маржа"].sum()), 2) if "Маржа" in merged_df.columns else 0.0

            if "Маржа" in merged_df.columns:
                margin_col = pd.to_numeric(merged_df["Маржа"], errors="coerce").fillna(0)
                positive_margin = round(margin_col[margin_col > 0].sum(), 2)
                negative_margin = round(margin_col[margin_col < 0].sum(), 2)
            else:
                positive_margin = 0.0
                negative_margin = 0.0

            margin_total_with_service = round(margin_all + service_total, 2)
            control_value = round(margin_total_with_service - (positive_margin + negative_margin), 2)

            positive_row = pd.DataFrame({
                "Показатель": ["Положительная маржа (по группам)"],
                "Итог": [positive_margin],
                "Тип начисления в расчете": [
                    "Сумма значений Маржа > 0 по группам артикулов (без общих расходов)"
                ],
            })
            negative_row = pd.DataFrame({
                "Показатель": ["Отрицательная маржа (по группам)"],
                "Итог": [negative_margin],
                "Тип начисления в расчете": [
                    "Сумма значений Маржа < 0 по группам артикулов (без общих расходов)"
                ],
            })
            service_row = pd.DataFrame({
                "Показатель": ["Общие расходы Ozon (вне групп артикулов)"],
                "Итог": [service_total],
                "Тип начисления в расчете": [
                    "Страхование + Подписка Premium + Ускоренный сбор отзывов + "
                    "Отгрузка в нерекомендованный слот. Уже входят в исходную 'Общую сумму', "
                    "но не относятся ни к одной группе артикулов."
                ],
            })
            margin_total_row = pd.DataFrame({
                "Показатель": ["Маржа (за минусом общих)"],
                "Итог": [margin_total_with_service],
                "Тип начисления в расчете": [
                    "Маржа по группам - Общие расходы. Это реальная прибыль за период."
                ],
            })
            margin_groups_row = pd.DataFrame({
                "Показатель": ["Маржа (по группам артикулов)"],
                "Итог": [margin_all],
                "Тип начисления в расчете": [
                    "Сумма значений Маржа по всем группам артикулов (без общих расходов)"
                ],
            })
            control_row = pd.DataFrame({
                "Показатель": ["Контроль: Маржа (итого) − (Полож. + Отриц.)"],
                "Итог": [control_value],
                "Тип начисления в расчете": [
                    "Должно равняться 'Общие расходы Ozon (вне групп артикулов)'. "
                    "Если совпадает — расчёт корректен."
                ],
            })

            # Перестраиваем порядок: вставляем новые строки сразу после "Маржа"
            idx_margin = financial_summary.index[
                financial_summary["Показатель"] == "Маржа"
            ].tolist()

            if idx_margin:
                pos = idx_margin[0]
                head = financial_summary.iloc[: pos + 1]
                tail = financial_summary.iloc[pos + 1 :]
                financial_summary = pd.concat(
                    [
                        head,
                        margin_groups_row,
                        service_row,
                        margin_total_row,
                        positive_row,
                        negative_row,
                        control_row,
                        tail,
                    ],
                    ignore_index=True,
                )
            else:
                financial_summary = pd.concat(
                    [
                        financial_summary,
                        margin_groups_row,
                        service_row,
                        margin_total_row,
                        positive_row,
                        negative_row,
                        control_row,
                    ],
                    ignore_index=True,
                )

            # ================= НОВЫЕ ПОКАЗАТЕЛИ =================
            
            # 1. Вознаграждение Озон = Вознаграждение за продажу + Возврат вознаграждения
            vozvrat_voznagrazhdeniya = 0.0
            vozvrat_voznagrazhdeniya_row = financial_summary[
                financial_summary["Показатель"] == "Возврат вознаграждения"
            ]
            if not vozvrat_voznagrazhdeniya_row.empty:
                vozvrat_voznagrazhdeniya = float(vozvrat_voznagrazhdeniya_row["Итог"].values[0])

            vozvrat_za_prodazhu = 0.0
            vozvrat_za_prodazhu_row = financial_summary[
                financial_summary["Показатель"] == "Вознаграждение за продажу"
            ]
            if not vozvrat_za_prodazhu_row.empty:
                vozvrat_za_prodazhu = float(vozvrat_za_prodazhu_row["Итог"].values[0])

            vozvrat_ozon_total = round(vozvrat_voznagrazhdeniya + vozvrat_za_prodazhu, 2)

            # 2. Выручка для расчета %
            vyruchka = 0.0
            vyruchka_row = financial_summary[
                financial_summary["Показатель"] == "Выручка, руб"
            ]
            if not vyruchka_row.empty:
                vyruchka = float(vyruchka_row["Итог"].values[0])

            # 3. Бонусы от Ozon для расчета полной цены продажи
            balli_za_skidki = 0.0
            balli_row = financial_summary[
                financial_summary["Показатель"] == "Баллы за скидки"
            ]
            if not balli_row.empty:
                balli_za_skidki = float(balli_row["Итог"].values[0])

            programmy_partnerov = 0.0
            prog_row = financial_summary[
                financial_summary["Показатель"] == "Программы партнёров"
            ]
            if not prog_row.empty:
                programmy_partnerov = float(prog_row["Итог"].values[0])

            # Полная цена продажи = Выручка + Баллы за скидки + Программы партнёров
            polnaya_tsena_prodazhi = vyruchka + balli_za_skidki + programmy_partnerov

            # 4. % Озон = |Вознаграждение Озон| / Полная цена продажи * 100
            # Берем абсолютное значение, т.к. вознаграждение отрицательное
            procent_ozon = round((abs(vozvrat_ozon_total) / polnaya_tsena_prodazhi * 100), 1) if polnaya_tsena_prodazhi != 0 else 0.0

            # Создаем новые строки для Вознаграждение Озон и % Озон
            vozvrat_ozon_row = pd.DataFrame({
                "Показатель": ["Вознаграждение Озон"],
                "Итог": [vozvrat_ozon_total],
                "Тип начисления в расчете": [
                    "Вознаграждение за продажу - Возврат вознаграждения"
                ],
            })
            procent_ozon_row = pd.DataFrame({
                "Показатель": ["% Озон"],
                "Итог": [procent_ozon],
                "Тип начисления в расчете": [
                    "|Вознаграждение Озон| / (Выручка + Баллы + Программы партнёров) * 100"
                ],
            })

            # Вставляем новые строки после "Выручка, руб"
            idx_vyruchka = financial_summary.index[
                financial_summary["Показатель"] == "Выручка, руб"
            ].tolist()

            if idx_vyruchka:
                pos = idx_vyruchka[0]
                head = financial_summary.iloc[: pos + 1]
                tail = financial_summary.iloc[pos + 1 :]
                financial_summary = pd.concat(
                    [
                        head,
                        vozvrat_ozon_row,
                        procent_ozon_row,
                        tail,
                    ],
                    ignore_index=True,
                )
            else:
                financial_summary = pd.concat(
                    [
                        financial_summary,
                        vozvrat_ozon_row,
                        procent_ozon_row,
                    ],
                    ignore_index=True,
                )

            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Количество групп"],
                "Итог": [len(merged_df)],
                "Тип начисления в расчете": ["Количество уникальных префиксов артикулов"],
            })], ignore_index=True)

            # Средневзвешенная себестоимость
            total_purchases_for_cost = merged_df["Количество выкупов"].sum()
            weighted_unit_cost = (
                round(merged_df["Кол-во*себес"].sum() / total_purchases_for_cost, 2)
                if total_purchases_for_cost > 0 else 0.0
            )
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Средневзвешенная себестоимость за 1 выкуп"],
                "Итог": [weighted_unit_cost],
                "Тип начисления в расчете": ["Сумма(Кол-во*себес) / Сумма(Количество выкупов)"],
            })], ignore_index=True)

            # Источник себестоимости
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Источник себестоимости"],
                "Итог": [
                    f"Файл Sebes.xlsx ({len(cost_map)} префиксов)"
                    if cost_map else
                    f"Ручной ввод — {unit_cost_default} руб. за 1 выкуп"
                ],
                "Тип начисления в расчете": [
                    "Если для префикса нет значения — берётся значение по умолчанию"
                ],
            })], ignore_index=True)

            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Себестоимость по умолчанию (руб.)"],
                "Итог": [unit_cost_default],
                "Тип начисления в расчете": [
                    "Применяется к группам, которых нет в Sebes.xlsx (или ко всем, если файл не загружен)"
                ],
            })], ignore_index=True)

            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Групп без значения себестоимости (fallback)"],
                "Итог": [len(fallback_groups)],
                "Тип начисления в расчете": [
                    ", ".join(fallback_groups[:20]) + ("..." if len(fallback_groups) > 20 else "")
                    if fallback_groups else "нет"
                ],
            })], ignore_index=True)

            # НАЛОГОВАЯ СТАВКА
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Налоговая ставка"],
                "Итог": [f"{tax_rate_percent:.2f} %"],
                "Тип начисления в расчете": ["Ставка налога, применённая к Выручке"],
            })], ignore_index=True)

            # Процент выкупа
            total_purchases = 0
            total_orders = 0
            for _, row in financial_summary.iterrows():
                if row["Показатель"] == "Количество выкупов":
                    total_purchases = row["Итог"]
                elif row["Показатель"] == "Количество заказов":
                    total_orders = row["Итог"]
            purchase_percentage_total = (
                round((total_purchases / total_orders) * 100, 1)
                if total_orders > 0 else 0.0
            )
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Процент Выкупа"],
                "Итог": [purchase_percentage_total],
                "Тип начисления в расчете": ["Количество выкупов / Количество заказов * 100"],
            })], ignore_index=True)

            # расширенная логистика уже посчитана выше
            # Берём её значение из financial_summary
            logistics_total = 0.0
            logistics_row = financial_summary[
                financial_summary["Показатель"] == "Логистика, руб"
            ]
            if not logistics_row.empty:
                logistics_total = abs(float(logistics_row["Итог"].values[0]))

            total_revenue_for_log = polnaya_tsena_prodazhi
            weighted_log_revenue = (
                round(logistics_total / total_revenue_for_log * 100, 1)
                if total_revenue_for_log != 0 else 100.0
            )

            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["% Лог/(Выручка + Баллы + Прог) (взвешенный)"],
                "Итог": [weighted_log_revenue],
                "Тип начисления в расчете": ["Сумма(все логистические расходы) / Сумма(Выручка+Баллы) * 100"],
            })], ignore_index=True)

            
            median_log_revenue = (
                round(merged_df["% Лог/(Выручка+Баллы)"].median(), 1)
                if len(merged_df) > 0 else 0.0
            )
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["% Лог/(Выручка+Баллы+Прог) (медиана по группам)"],
                "Итог": [median_log_revenue],
                "Тип начисления в расчете": [
                    "Медиана значений % Лог/(Выручка+Баллы+Прог.партнёров) по группам"
                ],
            })], ignore_index=True)

                        # Считаем % Озон + Лог
            procent_ozon_log = round(procent_ozon + weighted_log_revenue, 1)

            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["% Озон + Лог"],
                "Итог": [procent_ozon_log],
                "Тип начисления в расчете": [
                    "% Озон + % Лог/(Выручка + Баллы + Прог) (взвешенный)"
                ],
            })], ignore_index=True)

            zero_revenue_groups = int((merged_df["Выручка, руб"] == 0).sum())
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Групп с Выручка = 0"],
                "Итог": [zero_revenue_groups],
                "Тип начисления в расчете": [
                    "Количество групп, у которых Выручка, руб = 0 (только логистика)"
                ],
            })], ignore_index=True)

            # Задаём нужный порядок показателей
            desired_order = [
                "Начислено",
                "Удержано",
                "Выручка, руб",
                "Общая сумма, руб",
                "Чистая прибыль, руб",
                "Логистика, руб",
                "% Лог/(Выручка + Баллы) (взвешенный)",
                "% Лог/(Выручка + Баллы) (медиана по группам)",
                "Баллы за скидки",
                "Вознаграждение Озон",
                "% Озон",
                "% Озон + Лог",
                "Количество выкупов",
                "Количество заказов",
                "Процент Выкупа",
                "Количество артикулов в группе",
                "Рекламные расходы, руб",
                "Программы партнёров",
                "Себестоимость",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог",
                "Маржа",
                "Маржа (по группам артикулов)",
                "Общие расходы Ozon (вне групп артикулов)",
                "Маржа (за минусом общих)",
                "Положительная маржа (по группам)",
                "Отрицательная маржа (по группам)",
                "Контроль: Маржа (итого) − (Полож. + Отриц.)",
                "Количество групп",
                "Средневзвешенная себестоимость за 1 выкуп",
                "Источник себестоимости",
                "Себестоимость по умолчанию (руб.)",
                "Групп без значения себестоимости (fallback)",
                "Налоговая ставка",
                "% Лог/Выручка (медиана по группам)",  # если есть
                "Групп с Выручка = 0",
            ]

            # Пересобираем DataFrame в нужном порядке
            # (строки, которых нет в desired_order, добавятся в конец)
            financial_summary = financial_summary.set_index("Показатель")
            financial_summary = financial_summary.loc[
                [p for p in desired_order if p in financial_summary.index] +
                [p for p in financial_summary.index if p not in desired_order]
            ].reset_index()

            # ================= EXCEL =================
            output = BytesIO()
            from openpyxl.styles import PatternFill
            from openpyxl.utils import get_column_letter

            fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

            def paint_purchase_percentage(ws, df):
                if "Процент выкупа, %" not in df.columns:
                    return
                col_idx = df.columns.get_loc("Процент выкупа, %") + 1
                for r in range(2, len(df) + 2):
                    cell = ws.cell(row=r, column=col_idx)
                    v = cell.value
                    if v is None:
                        continue
                    cell.fill = fill_red if v < 20 else (fill_yellow if v < 30 else fill_green)

            def paint_margin(ws, df):
                if "Маржа" not in df.columns:
                    return
                col_idx = df.columns.get_loc("Маржа") + 1
                for r in range(2, len(df) + 2):
                    cell = ws.cell(row=r, column=col_idx)
                    v = cell.value
                    if v is None:
                        continue
                    cell.fill = fill_green if v > 0 else fill_red

            def paint_summary_margin_rows(ws, df):
                """Красит итоговые строки маржи и общескладских расходов в сводке."""
                if "Показатель" not in df.columns:
                    return
                for r in range(2, len(df) + 2):
                    name = ws.cell(row=r, column=1).value
                    if name in ("Положительная маржа (по группам)",):
                        ws.cell(row=r, column=1).fill = fill_green
                        ws.cell(row=r, column=2).fill = fill_green
                    elif name in ("Отрицательная маржа (по группам)",
                                  "Общие расходы Ozon (вне групп артикулов)"):
                        ws.cell(row=r, column=1).fill = fill_red
                        ws.cell(row=r, column=2).fill = fill_red
                    elif name in ("Маржа (за минусом общих)",
                                  "Маржа (по группам артикулов)"):
                        # выделяем жирным в Excel без заливки
                        ws.cell(row=r, column=1).font = ws.cell(row=r, column=1).font.copy(bold=True)
                        ws.cell(row=r, column=2).font = ws.cell(row=r, column=2).font.copy(bold=True)

            def autofit_columns(ws, df, min_width=8, max_width=60):
                for idx, col_name in enumerate(df.columns, start=1):
                    header_len = len(str(col_name))
                    try:
                        data_len = df[col_name].astype(str).map(len).max()
                    except Exception:
                        data_len = 0
                    if pd.isna(data_len):
                        data_len = 0
                    width = min(max(header_len, int(data_len), min_width), max_width)
                    ws.column_dimensions[get_column_letter(idx)].width = width + 2

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                financial_summary.to_excel(writer, sheet_name="0_Финансовая_сводка", index=False)
                merged_df.to_excel(writer, sheet_name="1_Группы_объединенная", index=False)
                detailed_df.to_excel(writer, sheet_name="3_Детально_по_артикулам", index=False)

                if cost_map:
                    pd.DataFrame(
                        list(cost_map.items()),
                        columns=["Префикс_группы", "Себестоимость"],
                    ).to_excel(writer, sheet_name="Sebes (использовано)", index=False)

                paint_purchase_percentage(writer.sheets["1_Группы_объединенная"], merged_df)
                paint_purchase_percentage(writer.sheets["3_Детально_по_артикулам"], detailed_df)
                paint_margin(writer.sheets["1_Группы_объединенная"], merged_df)
                paint_summary_margin_rows(writer.sheets["0_Финансовая_сводка"], financial_summary)

                autofit_columns(writer.sheets["0_Финансовая_сводка"], financial_summary)
                autofit_columns(writer.sheets["1_Группы_объединенная"], merged_df)
                autofit_columns(writer.sheets["3_Детально_по_артикулам"], detailed_df)

            output.seek(0)
            date_range = extract_date_range(excel_file.name)
            result_filename = (
                f"ozon_analysis_result_{date_range}.xlsx"
                if date_range else "ozon_analysis_result.xlsx"
            )
            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = f"attachment; filename={result_filename}"
            return response

        except Exception as e:
            return render(
                request,
                "forms_app/form21.html",
                {
                    "error": f"Ошибка при обработке файла: {str(e)}",
                    "unit_cost_default": unit_cost_default,
                    "tax_rate_percent": tax_rate_percent,
                },
            )

    # GET — значения по умолчанию
    return render(
        request,
        "forms_app/form21.html",
        {
            "unit_cost_default": UNIT_COST_DEFAULT,
            "tax_rate_percent": TAX_RATE_DEFAULT * 100,
        },
    )