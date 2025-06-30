# stock_replace_view.py

import os
import pandas as pd
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from forms_app.models import StockRecord
from django.core.exceptions import PermissionDenied
from django.conf import settings
from django.contrib import messages
from forms_app.models import UserReport
from io import BytesIO


@login_required
def replace_stock(request):
    """
    Представление для полной замены output_stock.xlsx новым файлом.
    Пользователь загружает файл через форму с кнопкой "replace_stock"
    """

    user_id = request.user.id
    base_dir = os.path.join("user_stock", str(user_id))
    output_path = os.path.join(base_dir, "output_stock.xlsx")
    full_output_path = os.path.join(settings.MEDIA_ROOT, output_path)

    # Создаем папку пользователя, если её нет
    os.makedirs(os.path.dirname(full_output_path), exist_ok=True)

    if request.method == "POST":
        replace_file = request.FILES.get("replace_stock")

        if not replace_file:
            messages.error(request, "❌ Файл не выбран для замены.")
            return redirect("forms_app:form5")

        try:
            print(f"🔄 Начинаем замену файла: {full_output_path}")

            # Временный путь для безопасного сохранения
            temp_path = full_output_path + ".tmp"

            # Сохраняем загруженный файл во временный путь
            with open(temp_path, "wb+") as destination:
                for chunk in BytesIO(replace_file.read()):
                    destination.write(chunk)

            # Удаляем старый файл, если он существует
            if os.path.exists(full_output_path):
                os.remove(full_output_path)
                print("🗑️ Старый файл удален")

            # Переименовываем временный файл в основной
            os.rename(temp_path, full_output_path)
            print("✅ Файл успешно заменён")

            # Обновляем запись в БД (если нужно)
            UserReport.objects.update_or_create(
                user=request.user,
                file_name="output_stock.xlsx",
                defaults={
                    "file_path": output_path,
                    "report_type": "form5",
                },
            )

            # Перенаправляем обратно на форму 5 с сообщением об успехе
            messages.success(request, "✅ Файл остатков успешно заменён")
            return redirect("forms_app:form5")

        except Exception as e:
            print(f"❌ Ошибка при замене файла: {e}")
            messages.error(request, f"❌ Ошибка при замене файла: {e}")
            return redirect("forms_app:form5")

    # GET-запрос (для тестирования или ошибок)
    messages.warning(request, "⚠️ Неверный метод запроса")
    return redirect("forms_app:form5")


@login_required
def preview_output_stock(request):
    """
    Представление для предпросмотра текущего файла output_stock.xlsx с поддержкой поиска
    """

    user_id = request.user.id
    base_dir = os.path.join("user_stock", str(user_id))
    full_output_path = os.path.join(settings.MEDIA_ROOT, base_dir, "output_stock.xlsx")

    if not os.path.exists(full_output_path):
        return render(
            request,
            "forms_app/preview.html",
            {"error": "❌ Файл output_stock.xlsx не найден для этого пользователя"},
        )

    try:
        # Чтение файла
        df = pd.read_excel(full_output_path)

        # Убираем лишние колонки
        if "Полный артикул" in df.columns:
            df = df.drop(columns=["Полный артикул"])

        # Поиск по запросу
        query = request.GET.get("q")
        if query:
            # Ищем по всем строкам и столбцам
            df = df[
                df.astype(str)
                .apply(lambda row: row.str.contains(query, case=False, na=False))
                .any(axis=1)
            ]

        # Конвертация в HTML
        table_html = df.to_html(
            classes="table table-bordered table-striped", index=False
        )

        return render(
            request,
            "forms_app/preview.html",
            {"table": table_html, "query": query or ""},
        )

    except Exception as e:
        print(f"❌ Ошибка при чтении файла: {e}")
        return render(
            request,
            "forms_app/preview.html",
            {"error": f"Ошибка при чтении файла: {e}"},
        )


@login_required
def replace_sql_stock(request):
    """
    Полная замена данных через загрузку Excel-файла в SQL
    """
    user = request.user

    if request.method == "POST":
        uploaded_file = request.FILES.get("replace_sql_stock")
        if not uploaded_file:
            messages.error(request, "❌ Файл не выбран")
            return redirect("forms_app:editable_preview_sql")

        try:
            df = pd.read_excel(BytesIO(uploaded_file.read()))

            required_columns = [
                "Артикул поставщика",
                "Размер",
                "Количество",
                "Место",
                "Примечание",
            ]
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"❌ В файле отсутствует колонка '{col}'")

            # Удаляем старые записи пользователя
            StockRecord.objects.filter(user=user).delete()

            # Создаем новые
            records = []
            for _, row in df.iterrows():
                records.append(
                    StockRecord(
                        user=user,
                        article_full_name=row["Артикул поставщика"],
                        size=row["Размер"],
                        quantity=row["Количество"],
                        location=row.get("Место", "Не указано"),
                        note=row.get("Примечание", ""),
                    )
                )

            StockRecord.objects.bulk_create(records)
            messages.success(request, "✅ Данные успешно заменены")
            return redirect("forms_app:editable_preview_sql")

        except Exception as e:
            messages.error(request, f"❌ Ошибка: {e}")
            return redirect("forms_app:editable_preview_sql")

    return render(request, "forms_app/replace_sql.html")
