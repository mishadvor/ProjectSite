# forms_app/views/form5_view.py

import os
import pandas as pd
from django.shortcuts import render, HttpResponse
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.conf import settings
from forms_app.models import UserReport
from io import BytesIO


def prepare_df(df):
    """Подготавливает DataFrame к обработке"""
    df["Размер"] = df["Размер"].astype(str).str.replace(r"\.0$", "", regex=True)
    return df.groupby(["Артикул поставщика", "Размер"], as_index=False)[
        "Количество"
    ].sum()


@login_required
def form5(request):
    if request.method == "POST":
        user_id = request.user.id
        base_dir = os.path.join("user_stock", str(user_id))
        output_path = os.path.join(base_dir, "output_stock.xlsx")

        # Папка пользователя
        full_output_path = os.path.join(settings.MEDIA_ROOT, output_path)
        os.makedirs(os.path.dirname(full_output_path), exist_ok=True)

        # Получаем файлы из формы
        input1 = request.FILES.get("input1")
        input2 = request.FILES.get("input2")
        input3 = request.FILES.get("input3")
        input_stock = request.FILES.get("input_stock")  # ✅ Файл начальных остатков

        # --- Подготовка начального DataFrame ---
        df_stock = pd.DataFrame(columns=["Артикул поставщика", "Размер", "Количество"])

        if input_stock:
            try:
                df_stock_raw = pd.read_excel(BytesIO(input_stock.read()))
                df_stock = prepare_df(df_stock_raw)
            except Exception as e:
                print(f"❌ Ошибка при чтении input_stock: {e}")
                df_stock = pd.DataFrame(
                    columns=["Артикул поставщика", "Размер", "Количество"]
                )
        elif os.path.exists(full_output_path):
            try:
                df_stock_raw = pd.read_excel(full_output_path)
                df_stock = prepare_df(df_stock_raw)
            except Exception as e:
                print(f"⚠️ Старый файл повреждён: {e}. Используем пустой.")
                df_stock = pd.DataFrame(
                    columns=["Артикул поставщика", "Размер", "Количество"]
                )
        else:
            df_stock = pd.DataFrame(
                columns=["Артикул поставщика", "Размер", "Количество"]
            )

        # --- Чтение входных данных ---
        df_input1 = pd.DataFrame(columns=["Артикул поставщика", "Размер", "Количество"])
        if input1:
            try:
                df_input1_raw = pd.read_excel(BytesIO(input1.read()))
                df_input1 = prepare_df(df_input1_raw)
            except Exception as e:
                print(f"❌ Ошибка при чтении input1: {e}")

        COLUMN_MAPPING = {"Артикул продавца": "Артикул поставщика"}

        df_input2 = pd.DataFrame(columns=["Артикул поставщика", "Размер", "Количество"])
        if input2:
            try:
                df_input2_raw = pd.read_excel(BytesIO(input2.read()))
                df_input2_raw.rename(columns=COLUMN_MAPPING, inplace=True)
                df_input2 = prepare_df(df_input2_raw)
            except Exception as e:
                print(f"❌ Ошибка при чтении input2: {e}")

        df_input3 = pd.DataFrame(columns=["Артикул поставщика", "Размер", "Количество"])
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
                print(f"❌ Ошибка при чтении input3: {e}")

        # --- Формируем изменения ---
        changes = pd.concat(
            [
                df_input1.assign(change=df_input1["Количество"]),
                df_input2.assign(change=-df_input2["Количество"]),
                df_input3.assign(change=-df_input3["Количество"]),
            ]
        )

        changes_grouped = changes.groupby(
            ["Артикул поставщика", "Размер"], as_index=False
        )["change"].sum()

        # --- Обновляем остатки ---
        df_stock = df_stock.set_index(["Артикул поставщика", "Размер"])
        changes_grouped = changes_grouped.set_index(["Артикул поставщика", "Размер"])

        updated_stock = df_stock.add(
            changes_grouped[["change"]].rename(columns={"change": "Количество"}),
            fill_value=0,
        )
        updated_stock["Количество"] = updated_stock["Количество"].fillna(0).astype(int)
        updated_stock = updated_stock.reset_index()

        # --- Сохраняем обратно в Excel ---
        updated_stock.to_excel(full_output_path, index=False)

        # === Сохраняем информацию об отчёте ===
        UserReport.objects.update_or_create(
            user=request.user,
            file_name="output_stock.xlsx",
            defaults={
                "file_path": os.path.join(base_dir, "output_stock.xlsx"),
                "report_type": "form5",
            },
        )

        # Отправляем результат пользователю
        with open(full_output_path, "rb") as f:
            response = HttpResponse(
                f.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = 'attachment; filename="output_stock.xlsx"'
            return response

    return render(request, "forms_app/form5.html")
