# forms_app/views/stock_replace_view.py

from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.conf import settings
from django.contrib import messages
from .models import StockRecord
import pandas as pd
import os
from io import BytesIO


@login_required
def replace_sql_stock(request):
    """
    Представление: замена всех записей в StockRecord данными из загруженного Excel
    """

    user = request.user

    if request.method == "POST":
        uploaded_file = request.FILES.get("replace_sql_stock")

        if not uploaded_file:
            messages.error(request, "❌ Файл не выбран")
            return redirect("forms_app:editable_preview_sql")

        try:
            # Чтение файла
            df = pd.read_excel(BytesIO(uploaded_file.read()))

            # Проверка наличия нужных колонок
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

            # Удаление старых данных
            StockRecord.objects.filter(user=user).delete()

            # Подготовка новых записей
            records_to_create = []
            for _, row in df.iterrows():
                records_to_create.append(
                    StockRecord(
                        user=user,
                        article_full_name=row["Артикул поставщика"],
                        size=row["Размер"],
                        quantity=row["Количество"],
                        location=row.get("Место", "Не указано"),
                        note=row.get("Примечание", ""),
                    )
                )

            # Массовое сохранение
            StockRecord.objects.bulk_create(records_to_create)

            messages.success(request, "✅ Данные успешно обновлены через SQL")
            return redirect("forms_app:editable_preview_sql")

        except Exception as e:
            messages.error(request, f"❌ Ошибка при обработке файла: {e}")
            return redirect("forms_app:editable_preview_sql")

    # GET-запрос — показываем страницу с формой загрузки
    return render(request, "forms_app/replace_sql.html")
