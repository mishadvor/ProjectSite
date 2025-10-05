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
            import signal

            # Сохраняем файл временно
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                for chunk in file.chunks():
                    tmp_file.write(chunk)
                tmp_path = tmp_file.name

            print(f"=== FORM2 DEBUG: File saved to {tmp_path} ===")

            # Функция с таймаутом
            def read_excel_with_timeout(path, timeout=30):
                def timeout_handler(signum, frame):
                    raise TimeoutError("Чтение Excel превысило время ожидания")

                # Устанавливаем таймаут
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(timeout)

                try:
                    result = pd.read_excel(path, engine="openpyxl")
                    signal.alarm(0)  # Отключаем таймаут
                    return result
                except TimeoutError:
                    print("=== FORM2 DEBUG: TIMEOUT ERROR ===")
                    raise
                except Exception as e:
                    signal.alarm(0)  # Отключаем таймаут при других ошибках
                    raise

            # Пробуем прочитать с таймаутом
            print("=== FORM2 DEBUG: Before pd.read_excel with timeout ===")
            try:
                df = read_excel_with_timeout(tmp_path, timeout=60)
                print(f"=== FORM2 DEBUG: DataFrame shape: {df.shape} ===")
            except TimeoutError:
                print("=== FORM2 DEBUG: Excel reading timed out ===")
                # Пробуем альтернативный метод - чтение по частям
                try:
                    print("=== FORM2 DEBUG: Trying chunk reading ===")
                    chunks = []
                    for chunk in pd.read_excel(
                        tmp_path, engine="openpyxl", chunksize=1000
                    ):
                        chunks.append(chunk)
                    df = pd.concat(chunks, ignore_index=True)
                    print(
                        f"=== FORM2 DEBUG: Chunk reading success, shape: {df.shape} ==="
                    )
                except Exception as chunk_error:
                    print(f"=== FORM2 CHUNK ERROR: {str(chunk_error)} ===")
                    raise chunk_error
            except Exception as read_error:
                print(f"=== FORM2 READ ERROR: {str(read_error)} ===")
                import traceback

                print(f"=== FORM2 READ TRACEBACK: {traceback.format_exc()} ===")
                raise read_error

            # Удаляем временный файл
            os.unlink(tmp_path)
            print("=== FORM2 DEBUG: Temp file deleted ===")

            return render(
                request,
                "forms_app/form2.html",
                {"success": f"Файл прочитан успешно! Строк: {len(df)}"},
            )

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
