import pandas as pd
import numpy as np
from django.http import HttpResponse
from django.shortcuts import render
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment,
    Font,
    Border,
    Side,
    PatternFill,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.styles import NamedStyle, Alignment, Font, Border, Side


def form2(request):
    if request.method == "POST":
        mode = request.POST.get("mode")

        try:
            file = request.FILES.get("file_single")
            if not file:
                return render(
                    request,
                    "forms_app/form2.html",
                    {"error": "Необходимо загрузить файл."},
                )

            print(f"=== FORM2 DEBUG: File: {file.name}, Size: {file.size} ===")

            import tempfile
            import os
            from openpyxl import load_workbook

            # Сохраняем файл временно
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                for chunk in file.chunks():
                    tmp_file.write(chunk)
                tmp_path = tmp_file.name

            print(f"=== FORM2 DEBUG: File saved to {tmp_path} ===")

            # Пробуем openpyxl напрямую (более легковесный)
            print("=== FORM2 DEBUG: Trying openpyxl directly ===")
            try:
                # Только чтение данных, без всего лишнего
                wb = load_workbook(tmp_path, read_only=True, data_only=True)
                sheet = wb.active

                # Просто считаем строки для теста
                row_count = 0
                for row in sheet.iter_rows(values_only=True):
                    row_count += 1
                    if row_count % 1000 == 0:  # Логируем каждые 1000 строк
                        print(f"=== FORM2 DEBUG: Read {row_count} rows ===")

                wb.close()
                print(
                    f"=== FORM2 DEBUG: Successfully read {row_count} rows with openpyxl ==="
                )

                # Удаляем временный файл
                os.unlink(tmp_path)

                return render(
                    request,
                    "forms_app/form2.html",
                    {"success": f"Файл прочитан успешно! Строк: {row_count}"},
                )

            except Exception as openpyxl_error:
                print(f"=== FORM2 OPENPYXL ERROR: {str(openpyxl_error)} ===")
                import traceback

                print(f"=== FORM2 OPENPYXL TRACEBACK: {traceback.format_exc()} ===")

                # Удаляем временный файл
                os.unlink(tmp_path)
                raise openpyxl_error

        except Exception as e:
            import traceback

            error_details = traceback.format_exc()
            print("=== FORM2 FULL ERROR ===")
            print(error_details)
            print("=== END ERROR ===")
            return render(
                request, "forms_app/form2.html", {"error": f"Ошибка: {str(e)}"}
            )

    return render(request, "forms_app/form2.html")
