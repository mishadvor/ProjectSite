# forms_app/views/form10_view.py

import pandas as pd
import numpy as np
from io import BytesIO
from django.shortcuts import render
from django.http import HttpResponse
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill
from openpyxl.utils import get_column_letter


def form10_view(request):
    result_df = None
    error_message = None

    if request.method == "POST" and request.FILES.get("excel_file"):
        uploaded_file = request.FILES["excel_file"]

        # Проверка расширения файла (игнорируем регистр)
        if not uploaded_file.name.lower().endswith(".xlsx"):
            error_message = "Пожалуйста, загрузите файл в формате .xlsx (поддерживаются .xlsx и .XLSX от Wildberries)."
        else:
            try:
                # Чтение Excel-файла
                df_raw = pd.read_excel(uploaded_file, header=1)
                df_raw = df_raw.reset_index(drop=True)

                # Обработка данных — аналогично твоему коду
                df1 = (
                    df_raw.groupby(
                        ["Артикул WB", "Баркод", "Артикул продавца", "Размер"],
                        as_index=False,
                    )
                    .agg(
                        {
                            "шт.": "sum",
                            "Сумма заказов минус комиссия WB, руб.": "sum",
                            "Выкупили, шт.": "sum",
                            "К перечислению за товар, руб.": "sum",
                            "Текущий остаток, шт.": "sum",
                        }
                    )
                    .round(0)
                )

                df1 = df1.rename(columns={"шт.": "Заказы, шт."})

                # Сортировка по доходу
                result_df = df1.sort_values(
                    by=["Сумма заказов минус комиссия WB, руб."], ascending=False
                ).reset_index(drop=True)

                # Экспорт в Excel "на лету"
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl", mode="w") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="Стат_продаж")
                    workbook = writer.book
                    worksheet = writer.sheets["Стат_продаж"]

                    # Стиль заголовков
                    style_name = "header_style"
                    if style_name not in workbook.named_styles:
                        header_style = NamedStyle(name=style_name)
                        header_style.font = Font(bold=True)
                        header_style.alignment = Alignment(
                            wrap_text=True, horizontal="center", vertical="center"
                        )
                        workbook.add_named_style(header_style)

                    for cell in worksheet[1]:
                        cell.style = style_name

                    # Автоподбор ширины столбцов
                    for column in worksheet.columns:
                        max_length = 0
                        col_letter = get_column_letter(column[0].column)
                        for cell in column:
                            try:
                                if cell.value not in [None, ""]:
                                    max_length = max(max_length, len(str(cell.value)))
                            except:
                                continue
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[col_letter].width = adjusted_width

                output.seek(0)

                # Возвращаем файл как ответ
                response = HttpResponse(
                    output.read(),
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                response["Content-Disposition"] = (
                    'attachment; filename="form10_result.xlsx"'
                )
                return response

            except Exception as e:
                error_message = f"Ошибка при обработке файла: {str(e)}"

    # Преобразуем результат в HTML-таблицу для отображения
    table_html = (
        result_df.to_html(classes="table table-striped table-bordered", index=False)
        if result_df is not None
        else None
    )

    context = {
        "table_html": table_html,
        "error_message": error_message,
    }
    return render(request, "forms_app/form10.html", context)
