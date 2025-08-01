# forms_app/views/form6_view.py

import pandas as pd
import os
from django.shortcuts import render, HttpResponse
from django.contrib.auth.decorators import login_required
from django.conf import settings
from io import BytesIO
from forms_app.models import StockRecord


def extract_first_3(article):
    """Возвращает первые 3 символа артикула"""
    return str(article)[:3]


def prepare_df(df):
    """Подготавливает DataFrame к обработке"""

    # Удаляем .0 у размеров (например, 46.0 → 46)
    if "Размер" in df.columns:
        df["Размер"] = df["Размер"].astype(str).str.replace(r"\.0$", "", regex=True)

    # Если колонок нет — создаём их
    for col in ["Место", "Примечание"]:
        if col not in df.columns:
            df[col] = ""

    # Группировка по первым 3 символам артикула + размер
    if "Артикул поставщика" in df.columns:
        df["Группа артикула"] = df["Артикул поставщика"].apply(extract_first_3)
    else:
        df["Группа артикула"] = ""

    # Группируем данные
    if not df.empty and "Количество" in df.columns:
        grouped = df.groupby(["Группа артикула", "Размер"], as_index=False).agg(
            {
                "Артикул поставщика": "first",
                "Место": lambda x: x.dropna().iloc[0] if not x.isna().all() else "",
                "Примечание": lambda x: (
                    x.dropna().iloc[0] if not x.isna().all() else ""
                ),
                "Количество": "sum",
            }
        )
    else:
        grouped = pd.DataFrame(
            columns=[
                "Группа артикула",
                "Размер",
                "Артикул поставщика",
                "Количество",
                "Место",
                "Примечание",
            ]
        )

    return grouped


@login_required
def form6(request):
    if request.method == "POST":
        user = request.user

        input1 = request.FILES.get("input1")
        input2 = request.FILES.get("input2")
        input3 = request.FILES.get("input3")
        input_stock = request.FILES.get("input_stock")

        base_dir = os.path.join("user_stock", str(user.id))
        output_path = os.path.join(base_dir, "output_stock_form6.xlsx")
        full_output_path = os.path.join(settings.MEDIA_ROOT, output_path)
        os.makedirs(os.path.dirname(full_output_path), exist_ok=True)

        # --- Подготовка начального остатка ---
        df_stock = pd.DataFrame(
            columns=["Группа артикула", "Размер", "Количество", "Место", "Примечание"]
        )

        if input_stock:
            try:
                df_stock_raw = pd.read_excel(BytesIO(input_stock.read()), sheet_name=0)

                # Переименовываем при необходимости
                if "Артикул" in df_stock_raw.columns:
                    df_stock_raw.rename(
                        columns={"Артикул": "Артикул поставщика"},
                        inplace=True,
                        errors="ignore",
                    )

                # Проверяем/добавляем нужные колонки
                for col in [
                    "Артикул поставщика",
                    "Размер",
                    "Количество",
                    "Место",
                    "Примечание",
                ]:
                    if col not in df_stock_raw.columns:
                        df_stock_raw[col] = (
                            ""
                            if col in ["Артикул поставщика", "Место", "Примечание"]
                            else 0
                        )

                df_stock = prepare_df(df_stock_raw)
            except Exception as e:
                print(f"❌ Ошибка при чтении input_stock: {e}")

        else:
            records = StockRecord.objects.filter(user=user).values(
                "article_full_name", "size", "quantity", "location", "note"
            )
            if records:
                df_stock = pd.DataFrame(records)
                df_stock.rename(
                    columns={
                        "article_full_name": "Артикул поставщика",
                        "size": "Размер",
                        "quantity": "Количество",
                        "location": "Место",
                        "note": "Примечание",
                    },
                    inplace=True,
                )
                df_stock = prepare_df(df_stock)

        # --- Обработка input1 (поступления) ---
        df_input1 = pd.DataFrame()
        if input1:
            try:
                df_raw = pd.read_excel(BytesIO(input1.read()), sheet_name=0)
                df_raw.rename(
                    columns={"Артикул продавца": "Артикул поставщика"},
                    inplace=True,
                    errors="ignore",
                )
                if "Количество, шт." in df_raw.columns:
                    df_raw.rename(
                        columns={"Количество, шт.": "Количество"}, inplace=True
                    )

                df_processed = prepare_df(df_raw)
                df_input1 = df_processed.assign(change=df_processed["Количество"])
            except Exception as e:
                print(f"❌ Ошибка при чтении input1: {e}")

        # --- Обработка input2 (FBS списание) ---
        df_input2 = pd.DataFrame()
        if input2:
            try:
                df_raw = pd.read_excel(BytesIO(input2.read()), sheet_name=0)
                df_raw.rename(
                    columns={"Артикул продавца": "Артикул поставщика"},
                    inplace=True,
                    errors="ignore",
                )
                df_processed = prepare_df(df_raw)
                df_input2 = df_processed.assign(change=-df_processed["Количество"])
            except Exception as e:
                print(f"❌ Ошибка при чтении input2: {e}")

        # --- Обработка input3 (FBO списание) ---
        df_input3 = pd.DataFrame()
        if input3:
            try:
                df_raw = pd.read_excel(BytesIO(input3.read()), sheet_name=0)
                df_raw.rename(
                    columns={"Артикул продавца": "Артикул поставщика"},
                    inplace=True,
                    errors="ignore",
                )
                if "Количество, шт." in df_raw.columns:
                    df_raw.rename(
                        columns={"Количество, шт.": "Количество"}, inplace=True
                    )

                df_processed = prepare_df(df_raw)
                df_input3 = df_processed.assign(change=-df_processed["Количество"])
            except Exception as e:
                print(f"❌ Ошибка при чтении input3: {e}")

        # --- Применение изменений ---
        changes = pd.concat([df_input1, df_input2, df_input3], ignore_index=True)

        if not changes.empty:
            changes_grouped = changes.groupby(
                ["Группа артикула", "Размер"], as_index=False
            ).agg(
                {
                    "Артикул поставщика": "first",
                    "change": "sum",
                    "Место": "first",
                    "Примечание": "first",
                }
            )
        else:
            changes_grouped = pd.DataFrame(
                columns=[
                    "Группа артикула",
                    "Размер",
                    "change",
                    "Артикул поставщика",
                    "Место",
                    "Примечание",
                ]
            )

        # --- Объединение остатков и изменений ---
        df_stock.set_index(["Группа артикула", "Размер"], inplace=True)
        changes_grouped.set_index(["Группа артикула", "Размер"], inplace=True)

        # Объединяем по индексу
        merged = df_stock.combine_first(
            changes_grouped.reset_index().set_index(["Группа артикула", "Размер"])
        )
        merged["change"] = merged["change"].fillna(0)

        # Обновляем количество
        merged["Количество"] = merged["Количество"].fillna(0) + merged["change"]

        merged = merged.reset_index()
        merged["Количество"] = merged["Количество"].astype(int)

        # --- Обновление данных в базе ---
        StockRecord.objects.filter(user=user).delete()

        updated_records = []
        for _, row in merged.iterrows():
            article = row["Артикул поставщика"]
            size = row["Размер"]
            quantity = row["Количество"]
            location = row.get("Место", "") or "Не указано"
            note = row.get("Примечание", "")

            updated_records.append(
                StockRecord(
                    user=user,
                    article_full_name=article,
                    size=size,
                    quantity=quantity,
                    location=location,
                    note=note,
                )
            )

        StockRecord.objects.bulk_create(updated_records)

        # --- Генерация выходного файла ---
        updated_df = pd.DataFrame(
            list(
                StockRecord.objects.filter(user=user).values(
                    "article_full_name", "size", "quantity", "location", "note"
                )
            )
        )
        updated_df.rename(
            columns={
                "article_full_name": "Артикул поставщика",
                "size": "Размер",
                "quantity": "Количество",
                "location": "Место",
                "note": "Примечание",
            },
            inplace=True,
        )
        updated_df = updated_df[
            ["Артикул поставщика", "Размер", "Количество", "Место", "Примечание"]
        ]

        updated_df.to_excel(full_output_path, index=False)

        # --- Отправка файла пользователю ---
        with open(full_output_path, "rb") as f:
            response = HttpResponse(
                f.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                'attachment; filename="output_stock_form6.xlsx"'
            )
            return response

    return render(request, "forms_app/form6.html")
