# forms_app/views/form3_view.py

import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from io import BytesIO
from django.shortcuts import render, HttpResponse
from django import forms
from openpyxl.drawing.image import Image as XLImage
from forms_app.forms import UploadFileForm  # ✅ Так тоже работает


# --- Форма прямо здесь ---
class ExcelUploadForm(forms.Form):
    excel_file = forms.FileField(label="Загрузите Excel-файл (.xlsx)")


def form3(request):
    if request.method == "POST":
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES["excel_file"]
            df = pd.read_excel(uploaded_file)

            # --- Первый отчет: по Области ---
            area_local = (
                df[df["Область"].notna() & (df["Область"] != "")]
                .groupby(["Область"])
                .agg({"Выкупили, шт.": "sum", "К перечислению за товар, руб.": "sum"})
                .astype(int)
                .reset_index()
            )
            area_local.sort_values(
                by="К перечислению за товар, руб.", ascending=False, inplace=True
            )

            # График для "Local_area"
            plt.figure(figsize=(12, 6))
            plt.bar(
                area_local["Область"].head(20),
                area_local["К перечислению за товар, руб."].head(20),
                color="skyblue",
            )
            plt.title("Сумма к перечислению по регионам (топ-20)")
            plt.xlabel("Регион")
            plt.ylabel("Сумма, руб.")
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            image_data_local = BytesIO()
            plt.savefig(image_data_local, format="png")
            plt.close()
            image_data_local.seek(0)

            # --- Второй отчет: по Федеральному округу ---
            area_federal = (
                df[df["Федеральный округ"].notna() & (df["Федеральный округ"] != "")]
                .groupby(["Федеральный округ"])
                .agg({"Выкупили, шт.": "sum", "К перечислению за товар, руб.": "sum"})
                .astype(int)
                .reset_index()
            )
            area_federal.sort_values(
                by="К перечислению за товар, руб.", ascending=False, inplace=True
            )

            # График для "Federal_area"
            plt.figure(figsize=(12, 6))
            plt.bar(
                area_federal["Федеральный округ"].head(10),
                area_federal["К перечислению за товар, руб."].head(10),
                color="lightgreen",
            )
            plt.title("Сумма к перечислению по федеральным округам")
            plt.xlabel("Федеральные округа")
            plt.ylabel("Сумма, руб.")
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            image_data_federal = BytesIO()
            plt.savefig(image_data_federal, format="png")
            plt.close()
            image_data_federal.seek(0)

            # --- Генерация Excel-файла в памяти ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # Лист 1: Local_area
                area_local.to_excel(writer, sheet_name="Local_area", index=False)
                ws1 = writer.sheets["Local_area"]
                img1 = XLImage(image_data_local)
                ws1.add_image(img1, "F2")

                # Лист 2: Federal_area
                area_federal.to_excel(writer, sheet_name="Federal_area", index=False)
                ws2 = writer.sheets["Federal_area"]
                img2 = XLImage(image_data_federal)
                ws2.add_image(img2, "F2")

            # --- Отправка файла пользователю ---
            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = "attachment; filename=Sums_Area.xlsx"
            return response

    else:
        form = ExcelUploadForm()

    return render(request, "forms_app/form3.html", {"form": form})
