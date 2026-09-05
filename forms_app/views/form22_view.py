# forms_app/views/form22_view.py

import pandas as pd
import numpy as np
from io import BytesIO
from django.shortcuts import render, HttpResponse
from django import forms
from django.contrib.auth.decorators import login_required
from openpyxl.drawing.image import Image as XLImage
import re
import os
from datetime import datetime


# Форма для загрузки файлов на Шаге 1
class Form22UploadForm(forms.Form):
    """Форма для загрузки файлов на Шаге 1"""

    file = forms.FileField(
        label="Выберите файлы Excel (.xlsx)",
        widget=forms.ClearableFileInput(),
        required=True,
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["file"].widget.attrs.update({"multiple": True, "accept": ".xlsx"})


class Form22Step2Form(forms.Form):
    """Форма для загрузки файлов на Шаге 2"""

    orders_file = forms.FileField(
        label="Файл с заказами (содержит 'Form22_Step1_Result' в имени)",
        widget=forms.ClearableFileInput(attrs={"accept": ".xlsx"}),
        required=True,
    )
    stock_file = forms.FileField(
        label="Файл с остатками (output_stock_form6)",
        widget=forms.ClearableFileInput(attrs={"accept": ".xlsx"}),
        required=True,
    )


def process_single_file(file_content, filename):
    """
    Обрабатывает один Excel файл и возвращает DataFrame с нужными колонками
    """
    try:
        df = pd.read_excel(BytesIO(file_content), header=1)

        required_columns = ["Артикул продавца", "Размер", "шт."]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            return (
                None,
                f"В файле {filename} отсутствуют колонки: {', '.join(missing_columns)}",
            )

        df_selected = df[["Артикул продавца", "Размер", "шт."]].copy()
        df_selected = df_selected.rename(columns={"шт.": "Заказано шт."})

        # Очищаем данные
        df_selected = df_selected.dropna(
            subset=["Артикул продавца", "Размер"], how="all"
        )
        df_selected["Артикул продавца"] = (
            df_selected["Артикул продавца"].astype(str).str.strip()
        )
        df_selected["Размер"] = df_selected["Размер"].astype(str).str.strip()
        df_selected["Артикул продавца"] = df_selected["Артикул продавца"].replace(
            "", "nan"
        )
        df_selected["Размер"] = df_selected["Размер"].replace("", "nan")
        df_selected["Заказано шт."] = df_selected["Заказано шт."].fillna(0)
        df_selected["Заказано шт."] = pd.to_numeric(
            df_selected["Заказано шт."], errors="coerce"
        ).fillna(0)

        return df_selected, None

    except Exception as e:
        return None, f"Ошибка при обработке {filename}: {str(e)}"


def sum_data_from_files(file_data_dict):
    """
    Суммирует данные из всех загруженных файлов
    """
    all_data = []
    errors = []
    processed_count = 0

    for filename, content in file_data_dict.items():
        df_processed, error = process_single_file(content, filename)
        if df_processed is not None:
            all_data.append(df_processed)
            processed_count += 1
        else:
            errors.append(error)

    if processed_count == 0:
        return None, errors

    combined_df = pd.concat(all_data, ignore_index=True)

    # Группировка по Артикулу и Размеру
    result_by_size = combined_df.groupby(
        ["Артикул продавца", "Размер"], as_index=False
    )["Заказано шт."].sum()
    result_by_size = result_by_size.sort_values(
        "Заказано шт.", ascending=False
    ).reset_index(drop=True)
    result_by_size["Заказано шт."] = result_by_size["Заказано шт."].round().astype(int)

    # Группировка только по Артикулу
    result_by_article = combined_df.groupby(["Артикул продавца"], as_index=False)[
        "Заказано шт."
    ].sum()
    result_by_article = result_by_article.sort_values(
        "Заказано шт.", ascending=False
    ).reset_index(drop=True)
    result_by_article["Заказано шт."] = (
        result_by_article["Заказано шт."].round().astype(int)
    )

    return (result_by_size, result_by_article), errors


def create_step1_result_file(result_by_size, result_by_article):
    """
    Создает Excel файл с результатами Шага 1
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Страница 1: По размерам
        result_by_size.to_excel(writer, sheet_name="По размерам", index=False)

        # Страница 2: По артикулам
        result_by_article.to_excel(writer, sheet_name="По артикулам", index=False)

        # Настраиваем ширину колонок
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

    output.seek(0)
    return output.getvalue()


def process_orders_file_step2(file_content):
    """
    Обрабатывает файл с заказами для Шага 2
    Читает первую страницу (По размерам)
    """
    try:
        xls = pd.ExcelFile(BytesIO(file_content))
        sheet_names = xls.sheet_names

        if not sheet_names:
            return None, "Файл не содержит страниц"

        # Берем первую страницу
        first_sheet = sheet_names[0]
        df_orders = pd.read_excel(BytesIO(file_content), sheet_name=first_sheet)

        required_columns = ["Артикул продавца", "Размер", "Заказано шт."]
        missing_columns = [
            col for col in required_columns if col not in df_orders.columns
        ]

        if missing_columns:
            return (
                None,
                f"В файле заказов отсутствуют колонки: {', '.join(missing_columns)}",
            )

        df_orders = df_orders[required_columns].copy()
        df_orders["Артикул продавца"] = (
            df_orders["Артикул продавца"].astype(str).str.strip()
        )
        df_orders["Размер"] = df_orders["Размер"].astype(str).str.strip()
        df_orders["Артикул продавца"] = df_orders["Артикул продавца"].replace("", "nan")
        df_orders["Размер"] = df_orders["Размер"].replace("", "nan")
        df_orders["Заказано шт."] = pd.to_numeric(
            df_orders["Заказано шт."], errors="coerce"
        ).fillna(0)

        return df_orders, None

    except Exception as e:
        return None, f"Ошибка при обработке файла заказов: {str(e)}"


def process_stock_file_step2(file_content):
    """
    Обрабатывает файл с остатками для Шага 2
    """
    try:
        df_stock = pd.read_excel(BytesIO(file_content))

        required_columns = ["Артикул поставщика", "Размер", "Количество"]
        missing_columns = [
            col for col in required_columns if col not in df_stock.columns
        ]

        if missing_columns:
            return (
                None,
                f"В файле остатков отсутствуют колонки: {', '.join(missing_columns)}",
            )

        df_stock = df_stock[required_columns].copy()
        df_stock["Артикул поставщика"] = (
            df_stock["Артикул поставщика"].astype(str).str.strip()
        )
        df_stock["Размер"] = df_stock["Размер"].astype(str).str.strip()
        df_stock["Артикул поставщика"] = df_stock["Артикул поставщика"].replace(
            "", "nan"
        )
        df_stock["Размер"] = df_stock["Размер"].replace("", "nan")
        df_stock["Количество"] = pd.to_numeric(
            df_stock["Количество"], errors="coerce"
        ).fillna(0)

        return df_stock, None

    except Exception as e:
        return None, f"Ошибка при обработке файла остатков: {str(e)}"


def merge_orders_with_stock_step2(df_orders, df_stock):
    """
    Объединяет данные заказов с остатками и вычисляет разницу
    """
    merged_df = pd.merge(
        df_orders,
        df_stock,
        left_on=["Артикул продавца", "Размер"],
        right_on=["Артикул поставщика", "Размер"],
        how="right",
    )

    merged_df["Заказано шт."] = merged_df["Заказано шт."].fillna(0)
    merged_df["Склад минус заказы"] = (
        merged_df["Количество"] - merged_df["Заказано шт."]
    )

    merged_df["Количество"] = merged_df["Количество"].round().astype(int)
    merged_df["Заказано шт."] = merged_df["Заказано шт."].round().astype(int)
    merged_df["Склад минус заказы"] = (
        merged_df["Склад минус заказы"].round().astype(int)
    )

    merged_df = merged_df.rename(columns={"Артикул поставщика": "Артикул"})
    merged_df = merged_df.rename(columns={"Количество": "Количество на складе"})

    if "Артикул продавца" in merged_df.columns:
        merged_df = merged_df.drop(columns=["Артикул продавца"])

    final_columns = [
        "Артикул",
        "Размер",
        "Заказано шт.",
        "Количество на складе",
        "Склад минус заказы",
    ]
    merged_df = merged_df[final_columns]
    merged_df = merged_df.sort_values(["Артикул", "Размер"]).reset_index(drop=True)

    return merged_df


def create_step2_result_file(result_df):
    """
    Создает Excel файл с результатами Шага 2
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Страница 1: Все товары
        result_df.to_excel(writer, sheet_name="Все товары", index=False)

        # Страница 2: Анализ продаж
        sales_df = result_df[result_df["Заказано шт."] > 0].copy()
        sales_df = sales_df.sort_values(
            "Склад минус заказы", ascending=True
        ).reset_index(drop=True)

        if len(sales_df) > 0:
            sales_df.to_excel(writer, sheet_name="Анализ продаж", index=False)

        # Настраиваем ширину колонок
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

    output.seek(0)
    return output.getvalue()


@login_required
def form22(request):
    """
    Главная страница Формы 22
    """
    step = request.GET.get("step", "1")

    context = {
        "step": step,
        "step1_form": Form22UploadForm(),
        "step2_form": Form22Step2Form(),
    }

    return render(request, "forms_app/form22.html", context)


@login_required
def form22_step1(request):
    """
    Шаг 1: Суммирование заказов из трех файлов
    """
    if request.method == "POST":
        form = Form22UploadForm(request.POST, request.FILES)

        if form.is_valid():
            uploaded_files = request.FILES.getlist("file")

            if not uploaded_files:
                return render(
                    request,
                    "forms_app/form22.html",
                    {
                        "step": "1",
                        "step1_form": form,
                        "error": "Ни одного файла не было загружено.",
                    },
                )

            # Проверяем, что загружены Excel файлы
            file_data = {}
            skipped_files = []

            for uploaded_file in uploaded_files:
                if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
                    file_data[uploaded_file.name] = uploaded_file.read()
                else:
                    skipped_files.append(uploaded_file.name)

            if len(file_data) < 3:
                error_msg = (
                    f"Загружено только {len(file_data)} Excel файлов. Нужно минимум 3."
                )
                if skipped_files:
                    error_msg += (
                        f" Пропущены не Excel файлы: {', '.join(skipped_files)}"
                    )
                return render(
                    request,
                    "forms_app/form22.html",
                    {"step": "1", "step1_form": form, "error": error_msg},
                )

            # Обрабатываем данные
            result, errors = sum_data_from_files(file_data)

            if result is None:
                error_msg = "Не удалось обработать файлы."
                if errors:
                    error_msg += " " + " ".join(errors)
                return render(
                    request,
                    "forms_app/form22.html",
                    {"step": "1", "step1_form": form, "error": error_msg},
                )

            result_by_size, result_by_article = result

            # Создаем Excel файл
            excel_file_bytes = create_step1_result_file(
                result_by_size, result_by_article
            )

            # ✅ Отправка файла - ИМЕНА НА ЛАТИНИЦЕ (как в Форме 3)
            response = HttpResponse(
                excel_file_bytes,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                "attachment; filename=Form22_Step1_Result.xlsx"
            )
            return response
        else:
            return render(
                request,
                "forms_app/form22.html",
                {
                    "step": "1",
                    "step1_form": form,
                    "error": "Пожалуйста, проверьте форму.",
                },
            )

    return render(
        request,
        "forms_app/form22.html",
        {"step": "1", "step1_form": Form22UploadForm()},
    )


@login_required
def form22_step2(request):
    """
    Шаг 2: Совмещение с остатками на складе
    """
    if request.method == "POST":
        form = Form22Step2Form(request.POST, request.FILES)

        if form.is_valid():
            orders_file = request.FILES.get("orders_file")
            stock_file = request.FILES.get("stock_file")

            # Проверяем, что файлы загружены
            if not orders_file or not stock_file:
                return render(
                    request,
                    "forms_app/form22.html",
                    {
                        "step": "2",
                        "step2_form": form,
                        "error": "Не все файлы загружены.",
                    },
                )

            # Проверяем расширения
            if not orders_file.name.lower().endswith((".xlsx", ".xls")):
                return render(
                    request,
                    "forms_app/form22.html",
                    {
                        "step": "2",
                        "step2_form": form,
                        "error": f'Файл "{orders_file.name}" не является Excel файлом.',
                    },
                )

            if not stock_file.name.lower().endswith((".xlsx", ".xls")):
                return render(
                    request,
                    "forms_app/form22.html",
                    {
                        "step": "2",
                        "step2_form": form,
                        "error": f'Файл "{stock_file.name}" не является Excel файлом.',
                    },
                )

            # Обрабатываем файл с заказами
            orders_content = orders_file.read()
            df_orders, error_orders = process_orders_file_step2(orders_content)

            if df_orders is None:
                return render(
                    request,
                    "forms_app/form22.html",
                    {"step": "2", "step2_form": form, "error": error_orders},
                )

            # Обрабатываем файл с остатками
            stock_content = stock_file.read()
            df_stock, error_stock = process_stock_file_step2(stock_content)

            if df_stock is None:
                return render(
                    request,
                    "forms_app/form22.html",
                    {"step": "2", "step2_form": form, "error": error_stock},
                )

            # Совмещаем данные
            merged_df = merge_orders_with_stock_step2(df_orders, df_stock)

            # Создаем Excel файл
            excel_file_bytes = create_step2_result_file(merged_df)

            # ✅ Отправка файла - ИМЕНА НА ЛАТИНИЦЕ (как в Форме 3)
            response = HttpResponse(
                excel_file_bytes,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                "attachment; filename=Form22_Step2_Result.xlsx"
            )
            return response
        else:
            return render(
                request,
                "forms_app/form22.html",
                {
                    "step": "2",
                    "step2_form": form,
                    "error": "Пожалуйста, проверьте форму.",
                },
            )

    return render(
        request, "forms_app/form22.html", {"step": "2", "step2_form": Form22Step2Form()}
    )
