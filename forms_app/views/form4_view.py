# forms_app/views/form4_view.py

import os
import re
import pandas as pd
from datetime import datetime, timedelta
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.core.files.storage import default_storage
from forms_app.forms import UploadFileForm
from forms_app.models import UserReport
from io import BytesIO
from django.conf import settings


@login_required
def upload_file(request):
    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES["file"]

            # Читаем файл в памяти без сохранения на диск
            file_data = BytesIO(uploaded_file.read())
            df_input = pd.read_excel(file_data, sheet_name=0)  # Чтение из памяти

            # === Получаем дату из имени файла или текущую ===
            def extract_date_from_filename(filename):
                match = re.search(r"отчет_за_(\d{2}\.\d{2}\.\d{4})\.xlsx", filename)
                if match:
                    return datetime.strptime(match.group(1), "%d.%m.%Y")
                return None

            file_date = extract_date_from_filename(uploaded_file.name)
            if not file_date:
                print("⚠️ Дата не найдена в имени файла. Используем сегодняшнюю дату.")
                file_date = datetime.now()

            # Вычисляем воскресенье недели
            # Для воскресенья прошлой недели:
            sunday_of_previous_week = file_date - timedelta(
                days=(file_date.weekday() + 1)
            )
            week_date = sunday_of_week.strftime("%d.%m.%Y")

            # Возвращаем курсор в начало (на случай повторного чтения)
            file_data.seek(0)
            df_input = pd.read_excel(file_data, sheet_name=0).head(150)

            if df_input.empty:
                raise ValueError("❌ Входной файл пустой — нечего записывать.")

            # === Путь к файлу ===
            user_id = request.user.id
            user_folder = f"user_reports/{user_id}"
            output_file_name = "Separated_Art_Rep.xlsx"
            output_file_path = os.path.join(
                settings.MEDIA_ROOT, user_folder, output_file_name
            )

            # Создаём папку пользователя, если её нет
            output_dir = os.path.join(settings.MEDIA_ROOT, user_folder)
            os.makedirs(output_dir, exist_ok=True)

            # === Функция для очистки названия листа от запрещённых символов ===
            def sanitize_sheet_name(name):
                invalid_chars = r"[\\/*?:\[\]]"
                return re.sub(invalid_chars, "", str(name).strip())[:31]

            # === Проверяем, есть ли уже такой файл ===
            existing_sheets = []
            mode = "w"
            if_sheet_exists = None

            if os.path.exists(output_file_path):
                try:
                    with pd.ExcelFile(output_file_path) as xls:
                        existing_sheets = xls.sheet_names
                    mode = "a"
                    if_sheet_exists = "overlay"
                except Exception as e:
                    print(f"⚠️ Целевой файл повреждён или нечитаем: {e}. Создаём новый.")

            print(f"Записываем в файл: {output_file_path} (режим: {mode})")

            # === Обработка и запись данных ===
            with pd.ExcelWriter(
                output_file_path,
                engine="openpyxl",
                mode=mode,
                if_sheet_exists=if_sheet_exists,
            ) as writer:

                for _, row in df_input.iterrows():
                    code = row["Код номенклатуры"]
                    sheet_name = sanitize_sheet_name(code)

                    # Собираем данные
                    article = row["Артикул поставщика"]
                    clear_sales_our = row["Чистые продажи Наши"]
                    clear_sales_vb = row["Чистая реализацич ВБ"]
                    clear_transfer = row["Чистое Перечисление"]
                    clear_transfer_without_log = row[
                        "Чистое Перечисление без Логистики"
                    ]
                    our_price_mid = row["Наша цена Средняя"]
                    vb_selling_mid = row["Реализация ВБ Средняя"]
                    transfer_mid = row["К перечислению Среднее"]
                    transfer_without_log_mid = row[
                        "К Перечислению без Логистики Средняя"
                    ]
                    qentity_sale = row["Чистые продажи, шт"]
                    sebes_sale = row["Себес Продаж (600р)"]
                    profit_1 = row["Прибыль на 1 Юбку"]
                    percent_sell = row["%Выкупа"]
                    profit = row["Прибыль"]
                    orders = row["Заказы"]

                    new_row = pd.DataFrame(
                        [
                            {
                                "Дата": week_date,
                                "Код номенклатуры": code,
                                "Артикул": article,
                                "Чистые продажи Наши": clear_sales_our,
                                "Чистая реализацич ВБ": clear_sales_vb,
                                "Чистое Перечисление": clear_transfer,
                                "Чистое Перечисление без Логистики": clear_transfer_without_log,
                                "Наша цена Средняя": our_price_mid,
                                "Реализация ВБ Средняя": vb_selling_mid,
                                "К перечислению Среднее": transfer_mid,
                                "К Перечислению без Логистики Средняя": transfer_without_log_mid,
                                "Чистые продажи, шт": qentity_sale,
                                "Себес Продаж (600р)": sebes_sale,
                                "Прибыль на 1 Юбку": profit_1,
                                "%Выкупа": percent_sell,
                                "Прибыль": profit,
                                "Заказы": orders,
                            }
                        ]
                    )

                    # Если лист существует — читаем и дополняем
                    if sheet_name in existing_sheets:
                        try:
                            df_existing = pd.read_excel(writer, sheet_name=sheet_name)
                            df_updated = pd.concat(
                                [df_existing, new_row], ignore_index=True
                            )
                        except Exception as e:
                            print(
                                f"⚠️ Ошибка при чтении листа '{sheet_name}': {e}. Создаём новый."
                            )
                            df_updated = new_row
                    else:
                        df_updated = new_row

                    df_updated.to_excel(writer, sheet_name=sheet_name, index=False)

                # Защита от пустого файла
                if len(writer.sheets) == 0:
                    pd.DataFrame().to_excel(writer, sheet_name="Шаблон", index=False)

            # === Сохраняем информацию о файле в БД ===
            UserReport.objects.update_or_create(
                user=request.user,
                file_name="Separated_Art_Rep.xlsx",
                defaults={
                    "file_path": f"{user_folder}/{output_file_name}",
                    "report_type": "form4",
                },
            )

            return redirect("forms_app:success_page")  # Стало (правильно)

    else:
        form = UploadFileForm()

    return render(request, "forms_app/upload.html", {"form": form})
