# forms_app/views/form21_view.py
import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
import re

# Константа себестоимости (руб. за 1 выкуп)
UNIT_COST = 900
TAX_RATE = 0.08


def extract_prefix(article):
    """Извлекает префикс артикула (первые 3 знака до _)"""
    if pd.isna(article) or article == "":
        return "unknown"
    parts = str(article).split("_")
    if len(parts) >= 1:
        return parts[0]
    return str(article)[:3]


def calculate_purchase_percentage(revenue_count, logistics_count):
    """Расчет процента выкупа"""
    if logistics_count == 0:
        return 0.0
    return round((revenue_count / logistics_count) * 100, 1)


def extract_date_range(filename):
    """Извлекает диапазон дат из имени файла вида '..._02.02.2026-08.02.2026.xlsx'."""
    if not filename:
        return None
    match = re.search(r"(\d{2}\.\d{2}\.\d{4}-\d{2}\.\d{2}\.\d{4})", filename)
    if match:
        return match.group(1)
    return None


def form21(request):
    """Загрузка файла и скачивание обработанного результата"""
    if request.method == "POST":
        excel_file = request.FILES.get("excel_file")

        if not excel_file:
            return render(
                request,
                "forms_app/form21.html",
                {"error": "Пожалуйста, выберите файл для загрузки."},
            )

        try:
            # Читаем файл
            df = pd.read_excel(excel_file, skiprows=1, header=0)

            # Добавляем префикс
            df["Префикс_артикула"] = df["Артикул"].apply(extract_prefix)

            # Отделяем рекламу
            ad_types = ["Оплата за клик"]
            df["Реклама"] = df["Тип начисления"].isin(ad_types).astype(int)
            df_ad = df[df["Реклама"] == 1].copy()
            df_non_ad = df[df["Реклама"] == 0].copy()

            total_ad_cost = df_ad["Сумма итого, руб."].sum() if len(df_ad) > 0 else 0

            # ============= ГРУППИРОВКА 1: ПО ПОЛНЫМ АРТИКУЛАМ =============
            detailed_stats = []

            for article in df_non_ad["Артикул"].unique():
                article_df = df_non_ad[df_non_ad["Артикул"] == article]

                total_sum = article_df["Сумма итого, руб."].sum()
                revenue_count = len(
                    article_df[article_df["Тип начисления"] == "Выручка"]
                )
                logistics_count = len(
                    article_df[article_df["Тип начисления"] == "Логистика"]
                )
                purchase_percentage = calculate_purchase_percentage(
                    revenue_count, logistics_count
                )
                revenue_sum = article_df[article_df["Тип начисления"] == "Выручка"][
                    "Сумма итого, руб."
                ].sum()
                logistics_sum = article_df[article_df["Тип начисления"] == "Логистика"][
                    "Сумма итого, руб."
                ].sum()

                detailed_stats.append(
                    {
                        "Артикул": article,
                        "Префикс": extract_prefix(article),
                        "Общая сумма, руб": total_sum,
                        "Выручка, руб": revenue_sum,
                        "Логистика, руб": logistics_sum,
                        "Количество выкупов": revenue_count,
                        "Количество заказов": logistics_count,
                        "Процент выкупа, %": purchase_percentage,
                    }
                )

            detailed_df = pd.DataFrame(detailed_stats)
            detailed_df = detailed_df.sort_values("Общая сумма, руб", ascending=False)

            # ============= ГРУППИРОВКА 2: ПО ПРЕФИКСАМ =============
            group_stats = []
            for prefix in df_non_ad["Префикс_артикула"].unique():
                group_df = df_non_ad[df_non_ad["Префикс_артикула"] == prefix]
                total_sum = group_df["Сумма итого, руб."].sum()
                revenue_count = len(group_df[group_df["Тип начисления"] == "Выручка"])
                logistics_count = len(
                    group_df[group_df["Тип начисления"] == "Логистика"]
                )
                purchase_percentage = calculate_purchase_percentage(
                    revenue_count, logistics_count
                )
                revenue_sum = group_df[group_df["Тип начисления"] == "Выручка"][
                    "Сумма итого, руб."
                ].sum()
                logistics_sum = group_df[group_df["Тип начисления"] == "Логистика"][
                    "Сумма итого, руб."
                ].sum()
                unique_articles = group_df["Артикул"].nunique()

                group_stats.append(
                    {
                        "Префикс_группы": prefix,
                        "Общая сумма, руб": total_sum,
                        "Выручка, руб": revenue_sum,
                        "Логистика, руб": logistics_sum,
                        "Количество выкупов": revenue_count,
                        "Количество заказов": logistics_count,
                        "Процент выкупа, %": purchase_percentage,
                        "Количество артикулов в группе": unique_articles,
                    }
                )

            group_df_result = pd.DataFrame(group_stats)
            group_df_result = group_df_result.sort_values(
                "Общая сумма, руб", ascending=False
            )

            # Сводка по типам
            prefix_pivot = pd.pivot_table(
                df_non_ad,
                values="Сумма итого, руб.",
                index="Префикс_артикула",
                columns="Тип начисления",
                aggfunc="sum",
                fill_value=0,
            )

            # Объединенная таблица
            prefix_pivot_reset = prefix_pivot.reset_index()
            prefix_pivot_reset = prefix_pivot_reset.rename(
                columns={"Префикс_артикула": "Префикс_группы"}
            )
            merged_df = pd.merge(
                group_df_result, prefix_pivot_reset, on="Префикс_группы", how="left"
            )

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

            # ============= НОВЫЕ КОЛОНКИ (пункт 2) =============
            # Себестоимость — константа
            merged_df["Себестоимость"] = UNIT_COST
            # Кол-во*себес = Количество выкупов * 900
            merged_df["Кол-во*себес"] = (
                merged_df["Количество выкупов"] * UNIT_COST
            ).round(2)
            # Чистая прибыль - себес = Чистая прибыль, руб - Кол-во*себес
            merged_df["Чистая прибыль - себес"] = (
                merged_df["Чистая прибыль, руб"] - merged_df["Кол-во*себес"]
            ).round(2)
            # Налог 8% = Выручка, руб * 0,08
            merged_df["Налог 8%"] = (
                merged_df["Выручка, руб"] * TAX_RATE
            ).round(2)
            # Маржа = Чистая прибыль - себес - Налог 8%
            merged_df["Маржа"] = (
                merged_df["Чистая прибыль - себес"] - merged_df["Налог 8%"]
            ).round(2)

            # ============= КОЛОНКА: % Лог/Выручка =============
            # Если Выручка = 0, значит есть только логистика => % = 100
            # Логистика в отчёте со знаком "-", берём модуль.
            merged_df["% Лог/Выручка"] = merged_df.apply(
                lambda row: round(abs(row["Логистика, руб"]) / row["Выручка, руб"] * 100, 1)
                if row["Выручка, руб"] > 0
                else 100.0,
                axis=1,
            )

            # ============= КОЛОНКА: Средняя цена выкупа =============
            # Выручка, руб / Количество выкупов
            # Если выкупов нет — оставляем 0, чтобы не делить на ноль
            merged_df["Средняя цена выкупа"] = merged_df.apply(
                lambda row: round(row["Выручка, руб"] / row["Количество выкупов"], 2)
                if row["Количество выкупов"] > 0
                else 0.0,
                axis=1,
            )

            # ============= ПОРЯДОК КОЛОНОК НА ЛИСТЕ "1_Группы_объединенная" =============
            # После "Выручка, руб" идут: Чистая прибыль, руб, Себестоимость,
            # Кол-во*себес, Чистая прибыль - себес, Налог 8%, Маржа
            base_columns = [
                "Префикс_группы",
                "Общая сумма, руб",
                "Выручка, руб",
                "Чистая прибыль, руб",
                "Себестоимость",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог 8%",
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
            # Добавляем оставшиеся колонки из сводки по типам (кроме префикса)
            other_columns = [
                col
                for col in merged_df.columns
                if col not in base_columns and col != "Префикс_группы"
            ]
            final_columns = [c for c in base_columns if c in merged_df.columns] + other_columns
            merged_df = merged_df[final_columns]

            merged_df = merged_df.sort_values("Маржа", ascending=False)

            # ============= ФИНАНСОВАЯ СВОДКА С ПОЯСНЕНИЯМИ =============

            formulas = {
                "Общая сумма, руб": "Сумма всех операций (Выручка + Логистика + Прочие начисления)",
                "Выручка, руб": "Сумма операций с типом 'Выручка'",
                "Логистика, руб": "Сумма операций с типом 'Логистика'",
                "Количество выкупов": "Количество операций с типом 'Выручка'",
                "Количество заказов": "Количество операций с типом 'Логистика'",
                "Количество артикулов в группе": "Количество уникальных артикулов в группе",
                "Рекламные расходы, руб": "Расходы на рекламу (тип 'Оплата за клик'), распределенные пропорционально выручке",
                "Чистая прибыль, руб": "Общая сумма + Рекламные расходы",
                "Себестоимость": "Фиксированная себестоимость за 1 выкуп (900 руб.)",
                "Кол-во*себес": "Количество выкупов * 900",
                "Чистая прибыль - себес": "Чистая прибыль, руб - Кол-во*себес",
                "Налог 8%": "Выручка, руб * 0,08",
                "Маржа": "Чистая прибыль - себес - Налог 8%",
            }

            # Собираем итоги по всем числовым колонкам из merged_df
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
                "Себестоимость",
                "Кол-во*себес",
                "Чистая прибыль - себес",
                "Налог 8%",
                "Маржа",
            ]

            # Добавляем колонки из сводки по типам (кроме префикса)
            for col in prefix_pivot_reset.columns:
                if col not in ["Префикс_группы"] and col not in numeric_columns:
                    numeric_columns.append(col)
                    formulas[col] = f"Сумма операций с типом '{col}'"

            for col in numeric_columns:
                if col in merged_df.columns:
                    total_value = merged_df[col].sum()
                    summary_data.append(
                        {
                            "Показатель": col,
                            "Итог": total_value,
                            "Тип начисления в расчете": formulas.get(
                                col, "Сумма всех операций по данному типу"
                            ),
                        }
                    )

            financial_summary = pd.DataFrame(summary_data)

            # Количество групп
            groups_count_row = pd.DataFrame(
                {
                    "Показатель": ["Количество групп"],
                    "Итог": [len(merged_df)],
                    "Тип начисления в расчете": [
                        "Количество уникальных префиксов артикулов"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, groups_count_row], ignore_index=True
            )

            # ============= ПРОЦЕНТ ВЫКУПА (пункт 1) =============
            # Берем готовые итоги из financial_summary
            total_purchases = 0
            total_orders = 0
            for _, row in financial_summary.iterrows():
                if row["Показатель"] == "Количество выкупов":
                    total_purchases = row["Итог"]
                elif row["Показатель"] == "Количество заказов":
                    total_orders = row["Итог"]

            purchase_percentage_total = (
                round((total_purchases / total_orders) * 100, 1)
                if total_orders > 0
                else 0.0
            )

            purchase_row = pd.DataFrame(
                {
                    "Показатель": ["Процент Выкупа"],
                    "Итог": [purchase_percentage_total],
                    "Тип начисления в расчете": [
                        "Количество выкупов / Количество заказов * 100"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, purchase_row], ignore_index=True
            )

            # ============= % Лог/Выручка: взвешенный и медианный =============
            total_logistics = merged_df["Логистика, руб"].sum()
            total_revenue_for_log = merged_df["Выручка, руб"].sum()

            # ============= % Лог/Выручка: взвешенный и медианный =============
            # Логистика со знаком "-", поэтому берём модуль
            total_logistics = abs(merged_df["Логистика, руб"].sum())
            total_revenue_for_log = merged_df["Выручка, руб"].sum()

            weighted_log_revenue = (
                round(total_logistics / total_revenue_for_log * 100, 1)
                if total_revenue_for_log != 0
                else 100.0
            )
            weighted_row = pd.DataFrame(
                {
                    "Показатель": ["% Лог/Выручка (взвешенный)"],
                    "Итог": [weighted_log_revenue],
                    "Тип начисления в расчете": [
                        "Сумма(Логистика) / Сумма(Выручка) * 100"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, weighted_row], ignore_index=True
            )

            median_log_revenue = (
                round(merged_df["% Лог/Выручка"].median(), 1)
                if len(merged_df) > 0
                else 0.0
            )
            median_row = pd.DataFrame(
                {
                    "Показатель": ["% Лог/Выручка (медиана по группам)"],
                    "Итог": [median_log_revenue],
                    "Тип начисления в расчете": [
                        "Медиана значений % Лог/Выручка по группам"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, median_row], ignore_index=True
            )

            zero_revenue_groups = int((merged_df["Выручка, руб"] == 0).sum())
            zero_rev_row = pd.DataFrame(
                {
                    "Показатель": ["Групп с Выручка = 0"],
                    "Итог": [zero_revenue_groups],
                    "Тип начисления в расчете": [
                        "Количество групп, у которых Выручка, руб = 0 (только логистика)"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, zero_rev_row], ignore_index=True
            )


            # ============= МАРЖА В СВОДКЕ (пункт 3) =============
            # Показатель "Маржа" уже добавлен в numeric_columns выше,
            # поэтому его сумма уже присутствует в financial_summary.
            # Ничего дополнительно добавлять не нужно — колонка "Маржа"
            # в merged_df агрегируется в цикле по numeric_columns.

            # Создание Excel файла
            output = BytesIO()

            from openpyxl.styles import PatternFill
            from openpyxl.utils import get_column_letter

            fill_red = PatternFill(
                start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
            )
            fill_yellow = PatternFill(
                start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
            )
            fill_green = PatternFill(
                start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
            )

            def paint_purchase_percentage(worksheet, df):
                """Красит колонку 'Процент выкупа, %' по диапазонам:
                0-20 — бледно-красный, 20-30 — бледно-жёлтый, 30-100 — бледно-зелёный.
                """
                if "Процент выкупа, %" not in df.columns:
                    return
                col_idx = df.columns.get_loc("Процент выкупа, %") + 1
                for row_idx in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    value = cell.value
                    if value is None:
                        continue
                    if value < 20:
                        cell.fill = fill_red
                    elif value < 30:
                        cell.fill = fill_yellow
                    else:
                        cell.fill = fill_green
            
            def paint_margin(worksheet, df):
                """Красит колонку 'Маржа': >0 — бледно-зелёный, <=0 — бледно-красный."""
                if "Маржа" not in df.columns:
                    return
                col_idx = df.columns.get_loc("Маржа") + 1
                for row_idx in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    value = cell.value
                    if value is None:
                        continue
                    if value > 0:
                        cell.fill = fill_green
                    else:
                        cell.fill = fill_red

            def autofit_columns(worksheet, df, min_width=8, max_width=60):
                """Ширина колонок по максимуму из длины заголовка и данных."""
                for idx, col_name in enumerate(df.columns, start=1):
                    header_len = len(str(col_name))
                    try:
                        data_len = df[col_name].astype(str).map(len).max()
                    except Exception:
                        data_len = 0
                    if pd.isna(data_len):
                        data_len = 0
                    width = max(header_len, int(data_len), min_width)
                    width = min(width, max_width)
                    worksheet.column_dimensions[get_column_letter(idx)].width = width + 2

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                financial_summary.to_excel(
                    writer, sheet_name="0_Финансовая_сводка", index=False
                )
                merged_df.to_excel(
                    writer, sheet_name="1_Группы_объединенная", index=False
                )
                detailed_df.to_excel(
                    writer, sheet_name="3_Детально_по_артикулам", index=False
                )

                paint_purchase_percentage(
                    writer.sheets["1_Группы_объединенная"], merged_df
                )
                paint_purchase_percentage(
                    writer.sheets["3_Детально_по_артикулам"], detailed_df
                )
                paint_margin(
                    writer.sheets["1_Группы_объединенная"], merged_df
                )

                autofit_columns(
                    writer.sheets["0_Финансовая_сводка"], financial_summary
                )
                autofit_columns(
                    writer.sheets["1_Группы_объединенная"], merged_df
                )
                autofit_columns(
                    writer.sheets["3_Детально_по_артикулам"], detailed_df
                )

            output.seek(0)

            date_range = extract_date_range(excel_file.name)
            if date_range:
                result_filename = f"ozon_analysis_result_{date_range}.xlsx"
            else:
                result_filename = "ozon_analysis_result.xlsx"

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
                {"error": f"Ошибка при обработке файла: {str(e)}"},
            )

    return render(request, "forms_app/form21.html")