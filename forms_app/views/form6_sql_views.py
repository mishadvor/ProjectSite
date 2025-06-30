# forms_app/views/form6_sql_views.py

import pandas as pd
from django.shortcuts import render, HttpResponse
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.core.exceptions import PermissionDenied
from forms_app.models import StockRecord
import os
from django.shortcuts import render, redirect
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger


@login_required
def preview_sql(request):
    """
    Представление: предпросмотр текущих остатков через SQL с пагинацией
    """
    user = request.user
    records = StockRecord.objects.filter(user=user).values(
        "article_full_name", "size", "quantity", "location", "note"
    )

    if not records.exists():
        return render(
            request,
            "forms_app/preview_sql.html",
            {"error": "❌ Нет данных остатков для отображения"},
        )

    df = pd.DataFrame(records)
    df.rename(
        columns={
            "article_full_name": "Артикул поставщика",
            "size": "Размер",
            "quantity": "Количество",
            "location": "Место",
            "note": "Примечание",
        },
        inplace=True,
    )

    query = request.GET.get("q")
    if query:
        df = df[
            df.astype(str)
            .apply(lambda row: row.str.contains(query, case=False, na=False))
            .any(axis=1)
        ]

    # Пагинация
    paginator = Paginator(
        df.to_dict(orient="records"), per_page=20
    )  # 20 записей на странице
    page_number = request.GET.get("page")

    try:
        page_obj = paginator.page(page_number)
    except PageNotAnInteger:
        page_obj = paginator.page(1)
    except EmptyPage:
        page_obj = paginator.page(paginator.num_pages)

    updated_df = pd.DataFrame(page_obj.object_list)
    table_html = updated_df.to_html(
        classes="table table-bordered table-striped", index=False
    )

    return render(
        request,
        "forms_app/preview_sql.html",
        {
            "table": table_html,
            "query": query or "",
            "page_obj": page_obj,
            "is_paginated": True,
        },
    )


@login_required
def editable_preview_sql(request):
    user = request.user
    query = request.GET.get("q", "")

    # Фильтруем записи через БД
    records = StockRecord.objects.filter(user=user)

    if query:
        from django.db.models import Q

        records = records.filter(
            Q(article_full_name__icontains=query)
            | Q(size__icontains=query)
            | Q(location__icontains=query)
        )

    # Пагинация
    paginator = Paginator(records.values(), 20)  # 20 записей на странице
    page_number = request.GET.get("page")

    try:
        page_obj = paginator.page(page_number)
    except PageNotAnInteger:
        page_obj = paginator.page(1)
    except EmptyPage:
        page_obj = paginator.page(paginator.num_pages)

    return render(
        request,
        "forms_app/preview_sql_editable.html",
        {
            "data": page_obj,
            "query": query,
            "is_paginated": True,
            "page_obj": page_obj,
            "paginator": paginator,
        },
    )


@login_required
def save_stock_sql(request):
    """
    Представление: сохраняет редактированные данные обратно в SQL
    """
    user = request.user
    records = StockRecord.objects.filter(user=user)

    if not records.exists():
        raise PermissionDenied("❌ Нет записей для сохранения")

    updated_records = []
    for record in records:
        quantity_key = f"quantity_{record.id}"
        location_key = f"location_{record.id}"
        note_key = f"note_{record.id}"

        try:
            new_quantity = int(request.POST.get(quantity_key, record.quantity))
        except ValueError:
            new_quantity = record.quantity

        new_location = request.POST.get(location_key, record.location or "")
        new_note = request.POST.get(note_key, record.note or "")

        record.quantity = new_quantity
        record.location = new_location
        record.note = new_note
        updated_records.append(record)

    StockRecord.objects.bulk_update(updated_records, ["quantity", "location", "note"])
    return redirect("forms_app:editable_preview_sql")


@login_required
def download_sql(request):
    """
    Представление: скачивание текущих остатков как Excel-файла
    """
    user = request.user
    records = StockRecord.objects.filter(user=user)

    if not records.exists():
        raise PermissionDenied("❌ Нет данных для загрузки")

    df = pd.DataFrame(
        list(
            records.values("article_full_name", "size", "quantity", "location", "note")
        )
    )
    df.rename(
        columns={
            "article_full_name": "Артикул поставщика",
            "size": "Размер",
            "quantity": "Количество",
            "location": "Место",
            "note": "Примечание",
        },
        inplace=True,
    )
    df = df[["Артикул поставщика", "Размер", "Количество", "Место", "Примечание"]]

    base_dir = os.path.join("user_stock", str(user.id))
    full_output_path = os.path.join(
        settings.MEDIA_ROOT, base_dir, "output_stock_form6.xlsx"
    )
    os.makedirs(os.path.dirname(full_output_path), exist_ok=True)

    df.to_excel(full_output_path, index=False)

    with open(full_output_path, "rb") as f:
        response = HttpResponse(
            f.read(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = (
            'attachment; filename="output_stock_form6.xlsx"'
        )
        return response


@login_required
def reset_stock_sql(request):
    user = request.user
    # Очищаем все записи пользователя
    StockRecord.objects.filter(user=user).delete()

    return redirect("forms_app:preview_sql")
