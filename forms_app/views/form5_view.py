# forms_app/views/form5_view.py
import os
import pandas as pd
from django.shortcuts import render, HttpResponse
from django.contrib.auth.decorators import login_required
from django.conf import settings
from forms_app.models import UserReport
from io import BytesIO


def extract_first_3(article):
    """Возвращает первые 3 символа артикула"""
    return str(article)[:3]


def prepare_df(df):
    """Подготавливает DataFrame к обработке"""
    if "Размер" in df.columns:
        df["Размер"] = df["Размер"].astype(str).str.replace(r"\.0$", "", regex=True)

    if "Артикул поставщика" in df.columns:
        df["Группа артикула"] = df["Артикул поставщика"].apply(extract_first_3)

    return df.groupby(["Группа артикула", "Размер"], as_index=False)["Количество"].sum()


@login_required
def form5(request):
    if request.method == "POST":
        # Инициализация всех переменных DataFrame
        df_stock_raw = pd.DataFrame()
        df_input1_raw = pd.DataFrame()
        df_input2_raw = pd.DataFrame()
        df_input3_raw = pd.DataFrame()

        user_id = request.user.id
        base_dir = os.path.join("user_stock", str(user_id))
        output_path = os.path.join(base_dir, "output_stock.xlsx")
        full_output_path = os.path.join(settings.MEDIA_ROOT, output_path)
        os.makedirs(os.path.dirname(full_output_path), exist_ok=True)

        input1 = request.FILES.get("input1")
        input2 = request.FILES.get("input2")
        input3 = request.FILES.get("input3")
        input_stock = request.FILES.get("input_stock")

        # Обработка input_stock
        df_stock = pd.DataFrame(columns=["Группа артикула", "Размер", "Количество"])
        if input_stock:
            try:
                temp_path = full_output_path + ".tmp"
                with open(temp_path, "wb+") as destination:
                    for chunk in input_stock.chunks():
                        destination.write(chunk)

                if os.path.exists(full_output_path):
                    os.remove(full_output_path)
                os.rename(temp_path, full_output_path)

                df_stock_raw = pd.read_excel(full_output_path)
                df_stock = prepare_df(df_stock_raw)
            except Exception as e:
                print(f"Ошибка при обработке input_stock: {e}")

        elif os.path.exists(full_output_path):
            try:
                df_stock_raw = pd.read_excel(full_output_path)
                df_stock = prepare_df(df_stock_raw)
            except Exception as e:
                print(f"Ошибка при чтении старого файла: {e}")

        # Обработка input1
        df_input1 = pd.DataFrame(columns=["Группа артикула", "Размер", "Количество"])
        if input1:
            try:
                df_input1_raw = pd.read_excel(BytesIO(input1.read()))
                df_input1 = prepare_df(df_input1_raw)
            except Exception as e:
                print(f"Ошибка при чтении input1: {e}")

        # Обработка input2
        COLUMN_MAPPING = {"Артикул продавца": "Артикул поставщика"}
        df_input2 = pd.DataFrame(columns=["Группа артикула", "Размер", "Количество"])
        if input2:
            try:
                df_input2_raw = pd.read_excel(BytesIO(input2.read()))
                df_input2_raw.rename(columns=COLUMN_MAPPING, inplace=True)
                df_input2_raw["Размер"] = (
                    df_input2_raw["Размер"]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                )
                if "Количество" not in df_input2_raw.columns:
                    df_input2_raw["Количество"] = 1
                df_input2 = prepare_df(df_input2_raw)
            except Exception as e:
                print(f"Ошибка при чтении input2: {e}")

        # Обработка input3
        df_input3 = pd.DataFrame(columns=["Группа артикула", "Размер", "Количество"])
        if input3:
            try:
                df_input3_raw = pd.read_excel(BytesIO(input3.read()))
                df_input3_raw.rename(columns=COLUMN_MAPPING, inplace=True)
                if "Количество, шт." in df_input3_raw.columns:
                    df_input3_raw.rename(
                        columns={"Количество, шт.": "Количество"}, inplace=True
                    )
                df_input3 = prepare_df(df_input3_raw)
            except Exception as e:
                print(f"Ошибка при чтении input3: {e}")

        # Формирование изменений
        changes = pd.concat(
            [
                df_input1.assign(change=df_input1["Количество"]),
                df_input2.assign(change=-df_input2["Количество"]),
                df_input3.assign(change=-df_input3["Количество"]),
            ]
        )

        changes_grouped = changes.groupby(
            ["Группа артикула", "Размер"], as_index=False
        )["change"].sum()

        # Обновление остатков
        df_stock = df_stock.set_index(["Группа артикула", "Размер"])
        changes_grouped = changes_grouped.set_index(["Группа артикула", "Размер"])

        updated_stock = df_stock.add(
            changes_grouped[["change"]].rename(columns={"change": "Количество"}),
            fill_value=0,
        )
        updated_stock["Количество"] = updated_stock["Количество"].fillna(0).astype(int)
        updated_stock = updated_stock.reset_index()

        # Сбор уникальных артикулов
        dfs_to_concat = []
        if not df_stock_raw.empty and "Артикул поставщика" in df_stock_raw.columns:
            dfs_to_concat.append(df_stock_raw[["Артикул поставщика"]])
        if not df_input1_raw.empty and "Артикул поставщика" in df_input1_raw.columns:
            dfs_to_concat.append(df_input1_raw[["Артикул поставщика"]])
        if not df_input2_raw.empty and "Артикул поставщика" in df_input2_raw.columns:
            dfs_to_concat.append(df_input2_raw[["Артикул поставщика"]])
        if not df_input3_raw.empty and "Артикул поставщика" in df_input3_raw.columns:
            dfs_to_concat.append(df_input3_raw[["Артикул поставщика"]])

        if dfs_to_concat:
            all_artifacts = pd.concat(dfs_to_concat).drop_duplicates().dropna()
            group_to_full = dict(
                zip(
                    all_artifacts["Артикул поставщика"].apply(extract_first_3),
                    all_artifacts["Артикул поставщика"],
                )
            )
            updated_stock["Артикул поставщика"] = updated_stock["Группа артикула"].map(
                group_to_full
            )
        else:
            updated_stock["Артикул поставщика"] = updated_stock["Группа артикула"]
            print("Предупреждение: не удалось восстановить полные артикулы")

        # Финализация DataFrame
        if "Группа артикула" in updated_stock.columns:
            updated_stock.drop(columns=["Группа артикула"], inplace=True)
        updated_stock = updated_stock[["Артикул поставщика", "Размер", "Количество"]]

        # Сохранение результата
        updated_stock.to_excel(full_output_path, index=False)

        # Сохранение информации об отчете
        UserReport.objects.update_or_create(
            user=request.user,
            file_name="output_stock.xlsx",
            defaults={
                "file_path": output_path,
                "report_type": "form5",
            },
        )

        # Отправка файла пользователю
        with open(full_output_path, "rb") as f:
            response = HttpResponse(
                f.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = 'attachment; filename="output_stock.xlsx"'
            return response

    return render(request, "forms_app/form5.html")
