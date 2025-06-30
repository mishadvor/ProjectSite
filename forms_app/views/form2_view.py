# forms_app/views/form2_view.py
import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
from openpyxl import Workbook


def safe_convert_to_int(value):
    """Безопасное преобразование в целые числа"""
    try:
        if isinstance(value, (pd.Series, pd.DataFrame)):
            return pd.to_numeric(value, errors="coerce").fillna(0).astype(int)
        return int(float(value))
    except (ValueError, TypeError):
        return 0


def safe_convert_to_float(value):
    """Безопасное преобразование в числа с плавающей точкой"""
    try:
        if isinstance(value, (pd.Series, pd.DataFrame)):
            return pd.to_numeric(value, errors="coerce").fillna(0.0)
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def safe_mean_calculation(x):
    """Безопасный расчет среднего значения"""
    try:
        filtered = x[x != 0]
        if len(filtered) > 0:
            return float(filtered.mean())
        return 0.0
    except Exception:
        return 0.0


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

            # Список числовых колонок для обработки
            numeric_cols = [
                "Цена розничная",
                "Вайлдберриз реализовал Товар (Пр)",
                "К перечислению Продавцу за реализованный Товар",
                "Услуги по доставке товара покупателю",
            ]

            # Предварительная обработка числовых колонок
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(safe_convert_to_float)

            # Основная агрегация по коду номенклатуры
            sums1_per_category = (
                df.groupby("Код номенклатуры")
                .agg(
                    {
                        "Артикул поставщика": "first",
                        "Цена розничная": "sum",
                        "Вайлдберриз реализовал Товар (Пр)": "sum",
                        "К перечислению Продавцу за реализованный Товар": "sum",
                        "Услуги по доставке товара покупателю": "sum",
                    }
                )
                .reset_index()
            )

            # Безопасное преобразование числовых колонок
            for col in numeric_cols:
                sums1_per_category[col] = sums1_per_category[col].apply(
                    safe_convert_to_int
                )

            # Дополнительные расчеты
            sums1_per_category["К Перечислению без Логистики"] = (
                sums1_per_category["К перечислению Продавцу за реализованный Товар"]
                - sums1_per_category["Услуги по доставке товара покупателю"]
            ).apply(safe_convert_to_int)

            sums1_per_category["Сумма СПП"] = (
                sums1_per_category["Цена розничная"]
                - sums1_per_category["Вайлдберриз реализовал Товар (Пр)"]
            ).apply(safe_convert_to_int)

            # Расчет процентов с обработкой деления на 0
            sums1_per_category["% Лог/рс"] = (
                np.where(
                    sums1_per_category["К перечислению Продавцу за реализованный Товар"]
                    == 0,
                    0,
                    (
                        sums1_per_category["Услуги по доставке товара покупателю"]
                        / sums1_per_category[
                            "К перечислению Продавцу за реализованный Товар"
                        ]
                    )
                    * 100,
                )
            ).round(1)

            sums1_per_category["% Лог/Наша Цена"] = (
                np.where(
                    sums1_per_category["Цена розничная"] == 0,
                    0,
                    (
                        sums1_per_category["Услуги по доставке товара покупателю"]
                        / sums1_per_category["Цена розничная"]
                    )
                    * 100,
                )
            ).round(1)

            # Агрегация возвратов по коду номенклатуры
            returns_by_code = (
                df[df["Тип документа"] == "Возврат"]
                .groupby("Код номенклатуры")
                .agg(
                    {
                        "Цена розничная": "sum",
                        "Вайлдберриз реализовал Товар (Пр)": "sum",
                        "К перечислению Продавцу за реализованный Товар": "sum",
                    }
                )
                .reset_index()
                .rename(
                    columns={
                        "Цена розничная": "Возвраты Наша цена",
                        "Вайлдберриз реализовал Товар (Пр)": "Возвраты реализация ВБ",
                        "К перечислению Продавцу за реализованный Товар": "Возвраты к перечислению",
                    }
                )
            )

            # Безопасное преобразование возвратов
            for col in [
                "Возвраты Наша цена",
                "Возвраты реализация ВБ",
                "Возвраты к перечислению",
            ]:
                returns_by_code[col] = returns_by_code[col].apply(safe_convert_to_int)

            # Объединение данных
            first_merged = sums1_per_category.merge(
                returns_by_code, on="Код номенклатуры", how="left"
            ).fillna(0)

            # Расчет чистых продаж
            first_merged["Чистые продажи Наши"] = (
                first_merged["Цена розничная"] - first_merged["Возвраты Наша цена"]
            ).apply(safe_convert_to_int)

            first_merged["Чистая реализация ВБ"] = (
                first_merged["Вайлдберриз реализовал Товар (Пр)"]
                - first_merged["Возвраты реализация ВБ"]
            ).apply(safe_convert_to_int)

            first_merged["Чистое Перечисление"] = (
                first_merged["К перечислению Продавцу за реализованный Товар"]
                - first_merged["Возвраты к перечислению"]
            ).apply(safe_convert_to_int)

            first_merged["Чистое Перечисление без Логистики"] = (
                first_merged["Чистое Перечисление"]
                - first_merged["Услуги по доставке товара покупателю"]
            ).apply(safe_convert_to_int)

            # Расчет средних значений по коду номенклатуры
            # =====================================================
            cost_per_category = (
                df.groupby("Код номенклатуры")
                .agg(
                    {
                        "Артикул поставщика": "first",
                        "Цена розничная": lambda x: safe_mean_calculation(x),
                        "Вайлдберриз реализовал Товар (Пр)": lambda x: safe_mean_calculation(
                            x
                        ),
                        "К перечислению Продавцу за реализованный Товар": lambda x: safe_mean_calculation(
                            x
                        ),
                        "Услуги по доставке товара покупателю": lambda x: safe_convert_to_float(
                            x.mean() * 2
                        ),
                    }
                )
                .reset_index()
            )

            # Дополнительные расчеты для средних значений
            cost_per_category["СПП Средняя"] = (
                cost_per_category["Цена розничная"]
                - cost_per_category["Вайлдберриз реализовал Товар (Пр)"]
            ).round(1)

            cost_per_category["К Перечислению без Логистики Средняя"] = (
                cost_per_category["К перечислению Продавцу за реализованный Товар"]
                - cost_per_category["Услуги по доставке товара покупателю"]
            ).round(1)

            cost_per_category["% Лог/Перечисление с Лог Средний"] = (
                np.where(
                    cost_per_category["К перечислению Продавцу за реализованный Товар"]
                    == 0,
                    0,
                    (
                        cost_per_category["Услуги по доставке товара покупателю"]
                        / cost_per_category[
                            "К перечислению Продавцу за реализованный Товар"
                        ]
                    )
                    * 100,
                )
            ).round(1)

            cost_per_category["% Лог/Наша цена Средний"] = (
                np.where(
                    cost_per_category["Цена розничная"] == 0,
                    0,
                    (
                        cost_per_category["Услуги по доставке товара покупателю"]
                        / cost_per_category["Цена розничная"]
                    )
                    * 100,
                )
            ).round(1)

            # Объединение со средними значениями
            second_merged = first_merged.merge(
                cost_per_category,
                on="Код номенклатуры",
                how="left",
                suffixes=("", "_Среднее"),
            ).fillna(0)

            # Обработка логистики (если есть соответствующая колонка)
            log_col = next((col for col in df.columns if "Виды логистики" in col), None)
            if log_col:
                df_exploded = df.explode(log_col)
                df_exploded[log_col] = df_exploded[log_col].fillna("Не указано")

                status_log = (
                    df_exploded.groupby("Код номенклатуры")
                    .agg(
                        {
                            log_col: lambda x: x.value_counts().to_dict(),
                            "Артикул поставщика": "first",
                        }
                    )
                    .reset_index()
                )

                # Раскрываем словарь в колонки
                status_log = pd.concat(
                    [
                        status_log.drop(log_col, axis=1),
                        status_log[log_col].apply(pd.Series).fillna(0),
                    ],
                    axis=1,
                )

                # Расчет показателей логистики
                for col in [
                    "К клиенту при продаже",
                    "От клиента при возврате",
                    "От клиента при отмене",
                ]:
                    status_log[col] = status_log.get(col, pd.Series(0)).fillna(0)

                numerator = status_log["К клиенту при продаже"]
                denominator = (
                    status_log["От клиента при отмене"]
                    + status_log["К клиенту при продаже"]
                    + status_log["От клиента при возврате"]
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
                status_log["Заказы"] = denominator
                status_log["От клиента при возврате"]
                status_log["От клиента при отмене"]

                # Объединение с данными логистики
                third_merged = second_merged.merge(
                    status_log[
                        [
                            "Код номенклатуры",
                            "%Выкупа",
                            "Себес Продаж (600р)",
                            "Чистые продажи, шт",
                            "Заказы",
                            "От клиента при возврате",
                            "От клиента при отмене",
                        ]
                    ],
                    on="Код номенклатуры",
                    how="left",
                ).fillna(0)
            else:
                third_merged = second_merged

            # Переименование столбцов
            third_merged = third_merged.rename(
                columns={
                    "Цена розничная": "Сумма Продаж Наша Цена",
                    "Вайлдберриз реализовал Товар (Пр)": "Сумма Продаж по цене ВБ",
                    "К перечислению Продавцу за реализованный Товар": "Сумма Продаж Перечисление С Лог",
                    "Услуги по доставке товара покупателю": "Логистика",
                    "От клиента при возврате": "Возвраты, шт",
                    "От клиента при отмене": "Отмена",
                }
            )

            # Финальные расчеты
            third_merged["Маржа"] = (
                third_merged["Чистое Перечисление без Логистики"]
                - third_merged["Себес Продаж (600р)"]
            ).round(1)

            third_merged["Налоги"] = (
                third_merged["Чистая реализация ВБ"] * 0.07
            ).round(1)

            third_merged["Прибыль"] = (
                third_merged["Маржа"] - third_merged["Налоги"]
            ).round(1)

            third_merged["Прибыль на 1 Юбку"] = (
                (third_merged["Прибыль"] / third_merged["Чистые продажи, шт"])
                .replace(np.inf, 0)
                .round(1)
            )

            # Добавляем пустые колонки перед определением порядка
            third_merged["План на неделю"] = ""  # Пустая строка
            third_merged["План по доходу"] = ""  # Пустая строка

            third_merged = third_merged.rename(
                columns={
                    "Цена розничная": "Сумма Продаж Наша Цена",
                    "Вайлдберриз реализовал Товар (Пр)": "Сумма Продаж по цене ВБ",
                    "К перечислению Продавцу за реализованный Товар": "Сумма Продаж Перечисление С Лог",
                    "Услуги по доставке товара покупателю": "Логистика",
                    "От клиента при возврате": "Возвраты, шт",
                    "От клиента при отмене": "Отмена",
                    "Цена розничная_Среднее": "Наша цена Средняя",
                    "Вайлдберриз реализовал Товар (Пр)_Среднее": "Реализация ВБ Средняя",
                    "К перечислению Продавцу за реализованный Товар_Среднее": "К перечислению Среднее",
                    "Услуги по доставке товара покупателю_Среднее": "Логистика Средняя",
                }
            )

            # Определяем желаемый порядок колонок
            desired_columns_order = [
                "Код номенклатуры",
                "Артикул поставщика",
                "Чистые продажи Наши",
                "Чистая реализация ВБ",
                "Чистое Перечисление",
                "Чистое Перечисление без Логистики",
                "Себес Продаж (600р)",
                "Прибыль",
                "Наша цена Средняя",  # Наша цена Средняя
                "Реализация ВБ Средняя",
                "К перечислению Среднее",
                "Прибыль на 1 Юбку",
                "Заказы",
                "Чистые продажи, шт",
                "%Выкупа",
                "СПП Средняя",
                "План на неделю",
                "План по доходу",
                "Сумма Продаж Наша Цена",
                "Сумма Продаж по цене ВБ",
                "Сумма Продаж Перечисление С Лог",
                "Логистика",
                "К Перечислению без Логистики",
                "Сумма СПП",
                "% Лог/рс",
                "% Лог/Наша Цена",
                "Возвраты Наша цена",
                "Возвраты реализация ВБ",
                "Возвраты к перечислению",
                "Услуги по доставке товара покупателю_Среднее",
                "К Перечислению без Логистики Средняя",
                "% Лог/Перечисление с Лог Средний",
                "% Лог/Наша цена Средний",
                "Возвраты, шт",
                "Отмена",
                "Маржа",
                "Налоги",
            ]

            # Фильтруем только существующие колонки
            existing_columns = [
                col for col in desired_columns_order if col in third_merged.columns
            ]
            third_merged = third_merged[existing_columns]

            third_merged.sort_values(
                by="Чистое Перечисление без Логистики", ascending=False, inplace=True
            )

            # Итоговая сводка
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
                        "Чистые продажи, шт",
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
                        third_merged["Чистые продажи, шт"].sum(),
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
                .groupby("Код номенклатуры", as_index=False)
                .agg(
                    {
                        "Артикул поставщика": "first",
                        "Цена розничная": [
                            ("Сумма продаж (Софт)", "sum"),
                            (
                                "Цена средняя (Софт)",
                                lambda x: safe_convert_to_float(
                                    x[x != 0].mean() if any(x != 0) else 0
                                ),
                            ),
                        ],
                    }
                )
            )

            # Переименование колонок для "Софт" товаров
            summary_soft.columns = [
                "Код номенклатуры",
                "Артикул поставщика",
                "Сумма продаж (Софт)",
                "Цена средняя (Софт)",
            ]

            # Безопасное преобразование числовых колонок
            summary_soft["Сумма продаж (Софт)"] = safe_convert_to_int(
                summary_soft["Сумма продаж (Софт)"]
            )
            summary_soft["Цена средняя (Софт)"] = safe_convert_to_float(
                summary_soft["Цена средняя (Софт)"]
            ).round(0)

            summary_soft.sort_values(
                by="Сумма продаж (Софт)", ascending=False, inplace=True
            )

            # Генерация Excel файла
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                third_merged.to_excel(
                    writer,
                    sheet_name="Основные данные",
                    index=False,
                    columns=existing_columns,
                )

                totall_summary.to_excel(
                    writer, sheet_name="Итоговая сводка", index=False
                )
                summary_soft.to_excel(writer, sheet_name="Софт товары", index=False)

            output.seek(0)

            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                'attachment; filename="wildberries_report.xlsx"'
            )
            return response

        except Exception as e:
            import traceback

            traceback.print_exc()
            return render(
                request,
                "forms_app/form2.html",
                {"error": f"Ошибка при обработке: {str(e)}"},
            )

    return render(request, "forms_app/form2.html")
