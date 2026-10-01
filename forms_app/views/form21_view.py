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
        # убираем возможные пробелы и запятые (на случай "900,50")
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
        tax_rate = tax_rate_percent / 100.0  # 8 -> 0.08

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

            merged_df["% Лог/Выручка"] = merged_df.apply(
                lambda row: round(abs(row["Логистика, руб"]) / row["Выручка, руб"] * 100, 1)
                if row["Выручка, руб"] > 0 else 100.0,
                axis=1,
            )
            merged_df["Средняя цена выкупа"] = merged_df.apply(
                lambda row: round(row["Выручка, руб"] / row["Количество выкупов"], 2)
                if row["Количество выкупов"] > 0 else 0.0,
                axis=1,
            )

            # ---------- Порядок колонок ----------
            base_columns = [
                "Префикс_группы",
                "Общая сумма, руб",
                "Выручка, руб",
                "Чистая прибыль, руб",
                "Себестоимость",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог",
                "Маржа",
                "Логистика, руб",
                "% Лог/Выручка",
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

            # ================= ФИНАНСОВАЯ СВОДКА =================
            formulas = {
                "Общая сумма, руб": "Сумма всех операций (Выручка + Логистика + Прочие начисления)",
                "Выручка, руб": "Сумма операций с типом 'Выручка'",
                "Логистика, руб": "Сумма операций с типом 'Логистика'",
                "Количество выкупов": "Количество операций с типом 'Выручка'",
                "Количество заказов": "Количество операций с типом 'Логистика'",
                "Количество артикулов в группе": "Количество уникальных артикулов в группе",
                "Рекламные расходы, руб": "Расходы на рекламу (тип 'Оплата за клик'), распределённые пропорционально выручке",
                "Чистая прибыль, руб": "Общая сумма + Рекламные расходы",
                "Себестоимость": "Себестоимость за 1 выкуп: из Sebes.xlsx по префиксу, иначе — значение с формы",
                "Кол-во*себес": "Количество выкупов × Себестоимость группы",
                "Чистая прибыль - себес": "Чистая прибыль, руб - Кол-во*себес",
                "Налог": f"Выручка, руб × {tax_rate_percent:.2f}%",
                "Маржа": "Чистая прибыль - себес - Налог",
            }

            summary_data = []
            numeric_columns = [
                "Общая сумма, руб",
                "Выручка, руб",
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
            for col in prefix_pivot_reset.columns:
                if col not in ["Префикс_группы"] and col not in numeric_columns:
                    numeric_columns.append(col)
                    formulas[col] = f"Сумма операций с типом '{col}'"

            for col in numeric_columns:
                if col in merged_df.columns:
                    summary_data.append({
                        "Показатель": col,
                        "Итог": merged_df[col].sum(),
                        "Тип начисления в расчете": formulas.get(
                            col, "Сумма всех операций по данному типу"
                        ),
                    })
            financial_summary = pd.DataFrame(summary_data)

            # ============= ПОЛОЖИТЕЛЬНАЯ / ОТРИЦАТЕЛЬНАЯ МАРЖА =============
            if "Маржа" in merged_df.columns:
                margin_col = pd.to_numeric(merged_df["Маржа"], errors="coerce").fillna(0)

                positive_margin = round(margin_col[margin_col > 0].sum(), 2)
                negative_margin = round(margin_col[margin_col < 0].sum(), 2)

                # Положительная маржа — сумма групп с Маржа > 0
                # Отрицательная маржа — сумма групп с Маржа < 0 (со знаком минус)
                positive_row = pd.DataFrame({
                    "Показатель": ["Положительная маржа"],
                    "Итог": [positive_margin],
                    "Тип начисления в расчете": [
                        "Сумма значений Маржа > 0 (прибыльные группы)"
                    ],
                })
                negative_row = pd.DataFrame({
                    "Показатель": ["Отрицательная маржа"],
                    "Итог": [negative_margin],
                    "Тип начисления в расчете": [
                        "Сумма значений Маржа < 0 (убыточные группы)"
                    ],
                })

                # Перестраиваем порядок: вставляем эти строки сразу после "Маржа"
                before_margin = financial_summary[
                    financial_summary["Показатель"] != "Маржа"
                ]
                # Определяем позицию строки "Маржа" в текущей сводке
                # Проще: разделим на "до Маржа" и "после Маржа"
                idx_margin = financial_summary.index[
                    financial_summary["Показатель"] == "Маржа"
                ].tolist()

                if idx_margin:
                    pos = idx_margin[0]  # индекс строки "Маржа"
                    head = financial_summary.iloc[: pos + 1]      # ... включая Маржа
                    tail = financial_summary.iloc[pos + 1 :]      # всё, что после
                    financial_summary = pd.concat(
                        [head, positive_row, negative_row, tail],
                        ignore_index=True,
                    )
                else:
                    # если "Маржа" вдруг нет — просто добавляем в конец
                    financial_summary = pd.concat(
                        [financial_summary, positive_row, negative_row],
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

            # НАЛОГОВАЯ СТАВКА — отдельная строка
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

            # % Лог/Выручка
            total_logistics = abs(merged_df["Логистика, руб"].sum())
            total_revenue_for_log = merged_df["Выручка, руб"].sum()
            weighted_log_revenue = (
                round(total_logistics / total_revenue_for_log * 100, 1)
                if total_revenue_for_log != 0 else 100.0
            )
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["% Лог/Выручка (взвешенный)"],
                "Итог": [weighted_log_revenue],
                "Тип начисления в расчете": ["Сумма(Логистика) / Сумма(Выручка) * 100"],
            })], ignore_index=True)

            median_log_revenue = (
                round(merged_df["% Лог/Выручка"].median(), 1)
                if len(merged_df) > 0 else 0.0
            )
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["% Лог/Выручка (медиана по группам)"],
                "Итог": [median_log_revenue],
                "Тип начисления в расчете": ["Медиана значений % Лог/Выручка по группам"],
            })], ignore_index=True)

            zero_revenue_groups = int((merged_df["Выручка, руб"] == 0).sum())
            financial_summary = pd.concat([financial_summary, pd.DataFrame({
                "Показатель": ["Групп с Выручка = 0"],
                "Итог": [zero_revenue_groups],
                "Тип начисления в расчете": [
                    "Количество групп, у которых Выручка, руб = 0 (только логистика)"
                ],
            })], ignore_index=True)

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