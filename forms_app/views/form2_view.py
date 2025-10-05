def form2(request):
    if request.method == "POST":
        mode = request.POST.get("mode")

        # ВРЕМЕННО УПРОЩЕННАЯ ВЕРСИЯ ДЛЯ ДИАГНОСТИКИ
        try:
            file = request.FILES.get("file_single")
            if not file:
                return render(
                    request,
                    "forms_app/form2.html",
                    {"error": "Необходимо загрузить файл."},
                )

            print(f"=== FORM2 DEBUG: File: {file.name}, Size: {file.size} ===")

            # Простейшее чтение файла
            import tempfile
            import os

            # Сохраняем файл временно для диагностики
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                for chunk in file.chunks():
                    tmp_file.write(chunk)
                tmp_path = tmp_file.name

            print(f"=== FORM2 DEBUG: File saved to {tmp_path} ===")

            # Пробуем прочитать
            df = pd.read_excel(tmp_path, engine="openpyxl")
            print(f"=== FORM2 DEBUG: DataFrame shape: {df.shape} ===")

            # Удаляем временный файл
            os.unlink(tmp_path)

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
