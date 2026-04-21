# forms_app/views/form21_view.py
import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO


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

            merged_df = merged_df.sort_values("Общая сумма, руб", ascending=False)

            # ============= ФИНАНСОВАЯ СВОДКА С ПОЯСНЕНИЯМИ =============

            # Словарь с описанием формул расчета для каждого показателя
            formulas = {
                "Общая сумма, руб": "Сумма всех операций (Выручка + Логистика + Прочие начисления)",
                "Выручка, руб": "Сумма операций с типом 'Выручка'",
                "Логистика, руб": "Сумма операций с типом 'Логистика'",
                "Количество выкупов": "Количество операций с типом 'Выручка'",
                "Количество заказов": "Количество операций с типом 'Логистика'",
                "Количество артикулов в группе": "Количество уникальных артикулов в группе",
                "Рекламные расходы, руб": "Расходы на рекламу (тип 'Оплата за клик'), распределенные пропорционально выручке",
                "Чистая прибыль, руб": "Общая сумма + Рекламные расходы",
            }

            # Собираем итоги по всем числовым колонкам из merged_df
            summary_data = []

            # Список колонок для агрегации (все числовые колонки)
            numeric_columns = [
                "Общая сумма, руб",
                "Выручка, руб",
                "Логистика, руб",
                "Количество выкупов",
                "Количество заказов",
                "Количество артикулов в группе",
                "Рекламные расходы, руб",
                "Чистая прибыль, руб",
            ]

            # Добавляем колонки из сводки по типам (кроме префикса)
            for col in prefix_pivot_reset.columns:
                if col not in ["Префикс_группы"] and col not in numeric_columns:
                    numeric_columns.append(col)
                    formulas[col] = f"Сумма операций с типом '{col}'"

            # Рассчитываем итоги по каждой колонке
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

            # Добавляем строку с количеством групп
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

            # Добавляем строку с процентом выкупа (средний)
            avg_purchase_percentage = (
                merged_df["Процент выкупа, %"].mean() if len(merged_df) > 0 else 0
            )
            purchase_row = pd.DataFrame(
                {
                    "Показатель": ["Средний процент выкупа, %"],
                    "Итог": [round(avg_purchase_percentage, 1)],
                    "Тип начисления в расчете": [
                        "(Количество выкупов / Количество заказов) * 100, усредненный по группам"
                    ],
                }
            )
            financial_summary = pd.concat(
                [financial_summary, purchase_row], ignore_index=True
            )

            # Создание Excel файла (3 страницы, без форматирования)
            output = BytesIO()

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

            output.seek(0)

            # Возвращаем файл
            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                "attachment; filename=ozon_analysis_result.xlsx"
            )
            return response

        except Exception as e:
            return render(
                request,
                "forms_app/form21.html",
                {"error": f"Ошибка при обработке файла: {str(e)}"},
            )

    return render(request, "forms_app/form21.html")
