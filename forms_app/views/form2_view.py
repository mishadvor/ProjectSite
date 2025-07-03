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

            third_merged.sort_values(by="Прибыль", ascending=False, inplace=True)
            # === Начало: Группировка по "Чистое Перечисление без Логистики" ===
            conditions = [
                third_merged["Прибыль"] > 10000,
                (third_merged["Прибыль"] >= 5000) & (third_merged["Прибыль"] <= 10000),
                (third_merged["Прибыль"] > 0) & (third_merged["Прибыль"] < 5000),
                third_merged["Прибыль"] < 0,
            ]

            categories_profit = [
                "1. >10 000",
                "2. 5 000 - 10 000",
                "3. 0 - 5 000",
                "4. <0 (убытки)",
            ]

            third_merged["Группа по прибыли"] = np.select(
                conditions, categories_profit, default="Не попал"
            )
            # === Конец: Группировка по "Чистое Перечисление без Логистики" ===

            # === Начало: Код для группировки по категориям Артикулов===
            def get_prefix(article):
                """Извлечение первых трёх символов артикула"""
                try:
                    return str(article).split("_")[0][:3]
                except:
                    return ""

            # Добавляем временный столбец с префиксом
            third_merged["Префикс"] = third_merged["Артикул поставщика"].apply(
                get_prefix
            )

            # Определяем категории и соответствующие им префиксы артикула
            categories = {
                # fmt: off
                "Экокожа черная": ["051", "054", "072", "079", "085"],
                "Джерси черная": ["001", "002", "003", "004", "005", "050", "122"],
                "Экокожа цветная": ["052", "053", "056", "057", "058", "059", "060", "061", "062", "063", "064",
                                    "065", "066", "067", "068", "069", "070", "071", "073", "074", "075", "076",
                                    "077", "078", "080", "081", "082", "083", "084", "093", "100", "101", "102",
                                    "103", "123"],
                "Джерси цветная": ["006", "007", "008", "009", "010", "011", "012", "013", "014", "015", "016",
                                "017", "018", "019", "020", "021", "022", "023", "024", "025", "026", "027",
                                "028", "029", "030", "031", "032", "033", "034", "035", "036", "037", "038",
                                "039", "040", "041", "042", "043", "044", "045", "046", "047", "048", "049",
                                "055", "086", "087", "088", "089", "090", "091", "092", "094", "095", "096",
                                "097", "098", "099", "104", "105", "106", "107", "108", "109", "110", "111",
                                "112", "113", "114", "115", "116", "117", "118", "119", "120", "121", "131",
                                "132", "133", "281", "341", "342", "343", "344", "345", "346", "347", "348",
                                "349", "350", "351", "354", "355", "356", "357", "358", "387"],
                "Джерси Короткая Черная (40,50)": ["352", "353"],
                "Софт лето": ["203", "197", "206", "168", "169", "170", "171", "172", "173", "174", "175", "176",
                            "177", "178", "179", "180", "181", "182", "183", "184", "185", "186", "187", "188",
                            "189", "190", "191", "192", "193", "194", "195", "196", "197", "198", "199", "200",
                            "201", "202", "203", "204", "205", "206", "207", "208", "209", "210", "211", "212",
                            "213", "225", "226", "227", "228", "229", "230", "232", "233", "234", "235", "236",
                            "237", "238", "239", "240", "241", "242", "243", "244", "245", "246", "247", "248",
                            "249", "250", "251", "252", "253", "254", "255", "256", "258", "259", "260", "261",
                            "262", "263", "264", "265", "266", "267", "268", "269", "270", "271", "272", "273",
                            "274", "275", "276", "277", "278", "279", "280", "282", "283", "284", "285", "286",
                            "287", "288", "289", "290", "291", "292", "293", "294", "295", "296", "297", "298",
                            "299", "300", "301", "302", "303", "304", "305", "306", "307", "308", "309", "310",
                            "311", "312", "313", "314", "315", "317", "318", "319", "320", "321", "322", "323",
                            "324", "325", "326", "328", "329", "330", "331", "332", "333", "334", "335", "336",
                            "337", "338", "340", "359", "360", "361", "362", "363", "364", "365", "366", "367",
                            "368", "369", "370", "371", "372", "373", "374", "375", "376", "377", "378", "379",
                            "380", "381", "382", "383", "385", "386"],
                "Другое": ["124", "125", "126", "127", "128", "129", "130", "134", "135", "136", "137", "138", 
                        "139", "140", "141", "142", "143", "144", "145", "146", "147", "148", "149", "150",
                        "151", "152", "153", "154", "155", "156", "157", "158", "159", "160", "161", "162",
                        "163", "164", "165", "166", "167", "214", "215", "216", "217", "218", "219", "220",
                        "221", "222", "223", "224", "339"],
                # fmt: on
                # Можно добавлять новые категории здесь
            }
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

            # Генерация Excel файла
            output = BytesIO()

            # summary_soft.to_excel(writer, sheet_name="Софт товары", index=False)
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # Лист "Основные данные"
                third_merged.to_excel(
                    writer,
                    sheet_name="Основные данные",
                    index=False,
                    columns=existing_columns,
                )

                # Лист "Итоговая сводка"
                totall_summary.to_excel(
                    writer, sheet_name="Итоговая сводка", index=False
                )

                # Листы с категориями
                for category, prefixes in categories.items():
                    # Фильтруем строки, где Префикс находится в списке для этой категории
                    filtered = third_merged[third_merged["Префикс"].isin(prefixes)]

                    # Удаляем временный столбец перед записью
                    filtered = filtered.drop(columns=["Префикс"])

                    # Проверяем, не превышает ли имя листа 31 символ (ограничение Excel)
                    safe_sheet_name = category[:31]

                    # Записываем на отдельный лист
                    filtered.to_excel(writer, sheet_name=safe_sheet_name, index=False)

                # Если нужно — можно добавить и другие листы (например, summary_soft)

                # Группы по прибыли

                for category in categories_profit:
                    filtered = third_merged[
                        third_merged["Группа по прибыли"] == category
                    ]
                    safe_sheet_name = category[
                        :31
                    ]  # Ограничение Excel на длину имени листа
                    filtered.to_excel(writer, sheet_name=safe_sheet_name, index=False)

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
