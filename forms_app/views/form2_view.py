# forms_app/views/form2_view.py
import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl import Workbook


def form2(request):
    if request.method == "POST":
        mode = request.POST.get("mode")
        try:
            # Загрузка файлов
            if mode == "single":
                file = request.FILES.get("file_single")
                if not file:
                    return render(
                        request,
                        "forms_app/form2.html",
                        {"error": "Необходимо загрузить файл."},
                    )
                df = pd.read_excel(file, dtype={"Баркод": str, "Размер": str})
            elif mode == "combined":
                file_russia = request.FILES.get("file_russia")
                file_cis = request.FILES.get("file_cis")
                if not file_russia or not file_cis:
                    return render(
                        request,
                        "forms_app/form2.html",
                        {"error": "Пожалуйста, загрузите оба файла."},
                    )
                df_russia = pd.read_excel(
                    file_russia, dtype={"Баркод": str, "Размер": str}
                )
                df_cis = pd.read_excel(file_cis, dtype={"Баркод": str, "Размер": str})
                df = pd.concat([df_russia, df_cis], ignore_index=True)
            else:
                return render(
                    request, "forms_app/form2.html", {"error": "Неизвестный режим."}
                )

            # Начало блока ОБРАБОТКА ДАННЫХ =================================
            # Суммирование и агрегация данных
            sums1_per_category = (
                df.groupby("Артикул поставщика")
                .agg(
                    {
                        "Цена розничная": "sum",
                        "Вайлдберриз реализовал Товар (Пр)": "sum",
                        "К перечислению Продавцу за реализованный Товар": "sum",
                        "Услуги по доставке товара покупателю": "sum",
                    }
                )
                .astype(int)
                .reset_index()
            )
            sums1_per_category["К Перечислению без Логистики"] = (
                sums1_per_category["К перечислению Продавцу за реализованный Товар"]
                - sums1_per_category["Услуги по доставке товара покупателю"]
            ).astype(int)
            sums1_per_category["Сумма СПП"] = (
                sums1_per_category["Цена розничная"]
                - sums1_per_category["Вайлдберриз реализовал Товар (Пр)"]
            ).astype(int)
            sums1_per_category["% Лог/рс"] = (
                (
                    (
                        sums1_per_category["Услуги по доставке товара покупателю"]
                        / sums1_per_category[
                            "К перечислению Продавцу за реализованный Товар"
                        ]
                    )
                    * 100
                )
                .replace(np.inf, 100.0)
                .round(1)
            )
            sums1_per_category["% Лог/Наша Цена"] = (
                (
                    (
                        sums1_per_category["Услуги по доставке товара покупателю"]
                        / sums1_per_category["Цена розничная"]
                    )
                    * 100
                )
                .replace(np.inf, 100.0)
                .round(1)
            )

            # Фильтруем только возвраты
            returns_by_article = (
                df[df["Тип документа"] == "Возврат"]
                .groupby("Артикул поставщика")[
                    [
                        "Цена розничная",
                        "Вайлдберриз реализовал Товар (Пр)",
                        "К перечислению Продавцу за реализованный Товар",
                    ]
                ]
                .sum()
                .fillna(0)
                .reset_index()
            )
            returns_by_article = returns_by_article.rename(
                columns={
                    "Цена розничная": "Возвраты Наша цена",
                    "Вайлдберриз реализовал Товар (Пр)": "Возвраты реализация ВБ",
                    "К перечислению Продавцу за реализованный Товар": "Возврты к перечислению",
                }
            )

            # Объединяем таблицы
            first_merged = sums1_per_category.merge(
                returns_by_article, on="Артикул поставщика", how="left"
            ).fillna(0)

            # Добавляем колонки Чистые продажи
            first_merged["Чистые продажи Наши"] = (
                first_merged["Цена розничная"] - first_merged["Возвраты Наша цена"]
            )
            first_merged["Чистая реализацич ВБ"] = (
                first_merged["Вайлдберриз реализовал Товар (Пр)"]
                - first_merged["Возвраты реализация ВБ"]
            )
            first_merged["Чистое Перечисление"] = (
                first_merged["К перечислению Продавцу за реализованный Товар"]
                - first_merged["Возврты к перечислению"]
            )
            first_merged["Чистое Перечисление без Логистики"] = (
                first_merged["Чистое Перечисление"]
                - first_merged["Услуги по доставке товара покупателю"]
            )

            # Средние значения
            cost_per_category = (
                df.groupby("Артикул поставщика")
                .agg(
                    {
                        "Цена розничная": lambda x: (
                            x[x != 0].mean() if any(x != 0) else 0
                        ),
                        "Вайлдберриз реализовал Товар (Пр)": lambda x: (
                            x[x != 0].mean() if any(x != 0) else 0
                        ),
                        "К перечислению Продавцу за реализованный Товар": lambda x: (
                            x[x != 0].mean() if any(x != 0) else 0
                        ),
                        "Услуги по доставке товара покупателю": lambda x: x.mean() * 2,
                    }
                )
                .astype(int)
                .reset_index()
            )
            cost_per_category["СПП Средняя"] = (
                cost_per_category["Цена розничная"]
                - cost_per_category["Вайлдберриз реализовал Товар (Пр)"]
            ).round(1)
            cost_per_category["К Перечислению без Логистики Средняя"] = (
                cost_per_category["К перечислению Продавцу за реализованный Товар"]
                - cost_per_category["Услуги по доставке товара покупателю"]
            ).round(1)
            cost_per_category["% Лог/Перечисление с Лог Средний"] = (
                (
                    (
                        cost_per_category["Услуги по доставке товара покупателю"]
                        / cost_per_category[
                            "К перечислению Продавцу за реализованный Товар"
                        ]
                    )
                    * 100
                )
                .replace(np.inf, 100.0)
                .round(1)
            )
            cost_per_category["% Лог/Наша цена Средний"] = (
                (
                    (
                        cost_per_category["Услуги по доставке товара покупателю"]
                        / cost_per_category["Цена розничная"]
                    )
                    * 100
                )
                .replace(np.inf, 100.0)
                .round(1)
            )

            # Объединяем таблицы
            second_merged = first_merged.merge(
                cost_per_category, on="Артикул поставщика", how="left"
            ).fillna(0)
            second_merged = second_merged.rename(
                columns={
                    "Цена розничная_y": "Наша цена Средняя",
                    "Вайлдберриз реализовал Товар (Пр)_y": "Реализация ВБ Средняя",
                    "К перечислению Продавцу за реализованный Товар_y": "К перечислению Среднее",
                    "Услуги по доставке товара покупателю_y": "Логистика Одной Юбки Средняя",
                }
            )
            """
            # Обработка логистики
            df_exploded = df.explode("Виды логистики, штрафов и корректировок ВВ")
            df_exploded["Виды логистики, штрафов и корректировок ВВ"] = df_exploded[
                "Виды логистики, штрафов и корректировок ВВ"
            ].fillna("Не указано")
            status_log = (
                df_exploded.groupby("Артикул поставщика")[
                    "Виды логистики, штрафов и корректировок ВВ"
                ]
                .value_counts()
                .unstack(fill_value=0)
                .reset_index()
            )
            """

            log_col = next((col for col in df.columns if "Виды логистики" in col), None)

            if log_col:
                # Обработка логистики
                df_exploded = df.explode(log_col)
                df_exploded[log_col] = df_exploded[log_col].fillna("Не указано")

                status_log = (
                    df_exploded.groupby("Артикул поставщика")[log_col]
                    .value_counts()
                    .unstack(fill_value=0)
                    .reset_index()
                )

            for col in [
                "К клиенту при продаже",
                "От клиента при возврате",
                "От клиента при отмене",
            ]:
                status_log[col] = status_log.get(col, pd.Series(0)).fillna(0)

            numerator = (
                status_log["К клиенту при продаже"]
                - status_log["От клиента при возврате"]
            )
            denominator = (
                status_log["От клиента при отмене"]
                + status_log["К клиенту при продаже"]
            )
            status_log["%Выкупа"] = np.where(
                (numerator == 0) & (denominator == 0),
                0,
                np.where(
                    (numerator == 0) & (denominator > 0),
                    -100,
                    np.where(
                        denominator == 0, 0, (numerator / denominator) * 100
                    ).astype(int),
                ),
            )
            status_log["Себес Продаж (600р)"] = (numerator * 600).round(0)
            status_log["Чистые продажи, шт"] = numerator
            status_log["Заказы"] = (
                status_log["От клиента при отмене"]
                + status_log["К клиенту при продаже"]
            )

            # Финальное объединение
            third_merged = second_merged.merge(
                status_log[
                    [
                        "Артикул поставщика",
                        "К клиенту при продаже",
                        "%Выкупа",
                        "Себес Продаж (600р)",
                        "Чистые продажи, шт",
                        "Заказы",
                    ]
                ],
                on="Артикул поставщика",
                how="left",
            ).fillna(0)

            # Переименование столбцов
            third_merged = third_merged.rename(
                columns={
                    "К клиенту при продаже": "Кол-во Продаж",
                    "Цена розничная_x": "Сумма Продаж Наша Цена",
                    "Вайлдберриз реализовал Товар (Пр)_x": "Сумма Продаж по цене ВБ",
                    "К перечислению Продавцу за реализованный Товар_x": "Сумма Продаж Перечислени С Лог",
                    "Услуги по доставке товара покупателю_x": "Логистика",
                }
            )

            # Расчеты
            third_merged["Маржа"] = (
                third_merged["Чистое Перечисление без Логистики"]
                - third_merged["Себес Продаж (600р)"]
            ).round(1)
            third_merged["Налоги"] = (
                third_merged["Чистая реализацич ВБ"] * 0.07
            ).round(1)
            third_merged["Прибыль"] = (
                third_merged["Маржа"] - third_merged["Налоги"]
            ).round(1)
            third_merged.sort_values(
                by="Сумма Продаж Наша Цена", ascending=False, inplace=True
            )
            third_merged["Прибыль на 1 Юбку"] = (
                (
                    third_merged["Маржа"] / third_merged["Чистые продажи, шт"]
                )  # !!!!!! Изменение
                .replace(np.inf, 0)
                .round(1)
            )

            # Формирование итоговой таблицы
            all_add_log = (
                df.groupby("Обоснование для оплаты")
                .agg(
                    {
                        "Услуги по доставке товара покупателю": "sum",
                        "Общая сумма штрафов": "sum",
                        "Хранение": "sum",
                        "Удержания": "sum",
                        "Платная приемка": "sum",
                    }
                )
                .reset_index()
            )
            totall_summary = pd.DataFrame(
                {
                    "Колонка": [
                        "Логистика",
                        "Сумма СПП",
                        "Сумма Чистых продаж без Возвратов и Логистики",
                        "Чистые продажи, шт",  # !!!!! изменение
                        "Заказы",
                        "Себестоимость продаж",
                        "Прибыль без налога",
                        "Штрафы",
                        "Хранение",
                        "Удержания",
                        "Платная приемка",
                        "Итого: прибыль минус доп. удержания",
                    ],
                    "Общая сумма": [
                        third_merged["Логистика"].sum(),
                        third_merged["Сумма СПП"].sum(),
                        third_merged["Чистое Перечисление без Логистики"].sum(),
                        third_merged["Чистые продажи, шт"].sum(),  # !!!! изменение
                        third_merged["Заказы"].sum(),
                        third_merged["Себес Продаж (600р)"].sum(),
                        third_merged["Прибыль"].sum(),
                        all_add_log["Общая сумма штрафов"].sum(),
                        all_add_log["Хранение"].sum(),
                        all_add_log["Удержания"].sum(),
                        all_add_log["Платная приемка"].sum(),
                        third_merged["Прибыль"].sum()
                        - (
                            all_add_log["Общая сумма штрафов"].sum()
                            + all_add_log["Хранение"].sum()
                            + all_add_log["Удержания"].sum()
                            + all_add_log["Платная приемка"].sum()
                        ),
                    ],
                }
            )

            # Обработка "Софт" товаров
            summary_soft = (
                df[df["Артикул поставщика"].str.contains("Софт", case=False, na=False)]
                .groupby("Артикул поставщика", as_index=False)
                .agg(
                    {
                        "Цена розничная": [
                            ("Сумма продаж наша цена (Софт)", "sum"),
                            ("Средняя Наша цена (Софт)", lambda x: x[x != 0].mean()),
                        ]
                    }
                )
                .round(0)
                .astype(int, errors="ignore")
            )
            summary_soft.columns = [
                "Артикул поставщика",
                "Сумма продаж (Софт)",
                "Цена средняя (Софт)",
            ]
            summary_soft.sort_values(
                by="Сумма продаж (Софт)", ascending=False, inplace=True
            )

            # Генерация Excel-файла
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                third_merged.to_excel(
                    writer, sheet_name="Summary_Table_by_Art", index=False
                )
                totall_summary.to_excel(
                    writer, sheet_name="Totall_Summary", index=False
                )
                summary_soft.to_excel(writer, sheet_name="Soft_Summary", index=False)
            output.seek(0)

            # Отправляем файл пользователю
            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                'attachment; filename="wildberries_report_form2.xlsx"'
            )
            return response

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return render(
                request,
                "forms_app/form2.html",
                {"error": f"Ошибка при обработке: {str(e)}"},
            )

    return render(request, "forms_app/form2.html")
