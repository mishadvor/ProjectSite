# forms_app/views/form16_view.py

import pandas as pd
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, FileResponse
from django.contrib import messages
from django.db import transaction
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import io
import os
import tempfile

# Импортируем модель (она теперь определена в models.py)
from forms_app.models import Form16Article
from forms_app.forms import Form16UploadForm


@login_required
def form16_main(request):
    """Главная страница Формы 16"""
    articles_count = Form16Article.objects.filter(
        user=request.user, is_active=True
    ).count()

    context = {"articles_count": articles_count, "has_articles": articles_count > 0}
    return render(request, "forms_app/form16_main.html", context)


@login_required
def form16_edit_table(request):
    """Редактирование таблицы артикулов"""
    # Получаем существующие артикулы пользователя
    existing_articles = {
        article.position: article
        for article in Form16Article.objects.filter(user=request.user)
    }

    if request.method == "POST":
        try:
            with transaction.atomic():
                # Обрабатываем все 15 позиций
                for position in range(1, 31):
                    article_wb = request.POST.get(f"article_wb_{position}", "").strip()
                    our_article = request.POST.get(
                        f"our_article_{position}", ""
                    ).strip()
                    comments = request.POST.get(f"comments_{position}", "").strip()
                    is_active = request.POST.get(f"active_{position}") == "on"

                    if position in existing_articles:
                        # Обновляем существующий артикул
                        article = existing_articles[position]
                        if article_wb:
                            article.article_wb = article_wb
                            article.our_article = our_article
                            article.comments = comments
                            article.is_active = is_active
                            article.save()
                        else:
                            # Если поле пустое - удаляем артикул
                            article.delete()
                    else:
                        # Создаем новый артикул
                        if article_wb:
                            Form16Article.objects.create(
                                user=request.user,
                                position=position,
                                article_wb=article_wb,
                                our_article=our_article,
                                comments=comments,
                                is_active=is_active,
                            )

            messages.success(request, "Таблица артикулов успешно сохранена!")
            return redirect("forms_app:form16_edit_table")

        except Exception as e:
            messages.error(request, f"Ошибка при сохранении: {str(e)}")

    # Для GET запроса - готовим данные для формы
    articles_list = Form16Article.objects.filter(user=request.user).order_by("position")

    # Создаем список данных для каждой позиции
    table_data = []
    for position in range(1, 31):
        if position in existing_articles:
            article = existing_articles[position]
            table_data.append(
                {
                    "position": position,
                    "article_wb": article.article_wb,
                    "our_article": article.our_article,
                    "comments": article.comments,
                    "is_active": article.is_active,
                    "exists": True,
                }
            )
        else:
            table_data.append(
                {
                    "position": position,
                    "article_wb": "",
                    "our_article": "",
                    "comments": "",
                    "is_active": False,
                    "exists": False,
                }
            )

    context = {"table_data": table_data, "articles_list": articles_list}
    return render(request, "forms_app/form16_edit_table.html", context)


@login_required
def form16_generate_report(request):
    """Генерация отчета на основе загруженного файла и сохраненных артикулов"""
    # Получаем активные артикулы пользователя
    articles_qs = Form16Article.objects.filter(
        user=request.user, is_active=True
    ).order_by("position")

    # Получаем артикулы как строки и очищаем от лишних пробелов
    articles = [str(article.article_wb).strip() for article in articles_qs]

    if not articles:
        messages.error(
            request, "У вас нет сохраненных артикулов. Заполните таблицу сначала."
        )
        return redirect("forms_app:form16_edit_table")

    # Инициализируем форму
    form = None

    if request.method == "POST":
        # Проверяем наличие файла
        if "file" not in request.FILES:
            messages.error(request, "Файл не загружен")
            return redirect("forms_app:form16_generate_report")

        file = request.FILES["file"]

        # Проверяем расширение файла
        if not file.name.endswith(".xlsx"):
            messages.error(request, "Файл должен быть в формате .xlsx")
            return redirect("forms_app:form16_generate_report")

        try:
            # Сначала проверим, какие страницы есть в файле
            xls = pd.ExcelFile(file)
            sheet_names = xls.sheet_names

            # Ищем нужную страницу
            target_sheet = None
            for sheet in sheet_names:
                if "детальная" in sheet.lower():
                    target_sheet = sheet
                    break

            if not target_sheet:
                target_sheet = sheet_names[0]
                messages.warning(
                    request,
                    f"Страница 'Детальная информация' не найдена. Используется первая страница: '{target_sheet}'",
                )

            # Читаем файл
            df = pd.read_excel(
                file,
                sheet_name=target_sheet,
                header=1,
            )

            # Преобразуем колонку 'Артикул WB' в строки
            df["Артикул WB_clean"] = df["Артикул WB"].astype(str).str.strip()

            # Очищаем наши артикулы
            articles_clean = [str(article).strip() for article in articles]

            # Получаем все уникальные артикулы из файла
            all_articles_in_file = df["Артикул WB_clean"].unique().tolist()

            # Диагностика
            diagnostic_info = {
                "file_name": file.name,
                "sheet_name": target_sheet,
                "total_rows_in_file": len(df),
                "unique_articles_in_file": len(all_articles_in_file),
                "our_articles_count": len(articles_clean),
                "found_articles": [],
                "not_found_articles": [],
                "file_articles_sample": all_articles_in_file[:20],
            }

            # Проверяем каждый артикул
            for article in articles_clean:
                if article in all_articles_in_file:
                    diagnostic_info["found_articles"].append(article)
                else:
                    diagnostic_info["not_found_articles"].append(article)

            # Если найдено мало артикулов - предупредить
            found_count = len(diagnostic_info["found_articles"])

            if found_count == 0:
                error_msg = f"❌ Ни один артикул не найден в файле!\n\n"
                error_msg += f"Искали: {articles_clean}\n\n"
                error_msg += f"Артикулы в файле (первые 30):\n"
                for i, art in enumerate(all_articles_in_file[:30], 1):
                    error_msg += f"{i}. {art}\n"

                messages.error(request, error_msg)
                return redirect("forms_app:form16_generate_report")

            # Фильтруем найденные артикулы
            df_filtered = df[
                df["Артикул WB_clean"].isin(diagnostic_info["found_articles"])
            ].copy()

            # Создаем порядок сортировки (по нашему списку артикулов)
            article_order = {article: i for i, article in enumerate(articles_clean)}

            # Добавляем колонку для сортировки
            df_filtered["Порядок_артикула"] = df_filtered["Артикул WB_clean"].map(
                article_order
            )
            df_filtered["Порядок_артикула"] = df_filtered["Порядок_артикула"].fillna(
                999
            )

            # Преобразуем размер для сортировки
            def try_convert_to_numeric(x):
                try:
                    return float(x)
                except:
                    return float("inf")

            df_filtered["Размер_число"] = (
                df_filtered["Размер"].astype(str).apply(try_convert_to_numeric)
            )

            # Сортируем
            df_sorted = df_filtered.sort_values(
                ["Порядок_артикула", "Артикул продавца", "Размер_число"],
                ascending=[True, True, True],
            )

            # Выбираем нужные колонки
            available_columns = []
            desired_columns = [
                "Артикул продавца",
                "Артикул WB",
                "Размер",
                "Доступность",
                "Комментарии",
                "Заказали, шт",
                "Выкупили, шт",
                "Процент выкупа",
                "Остатки на текущий день, шт",
            ]

            for col in desired_columns:
                if col in df_sorted.columns:
                    available_columns.append(col)

            df_result = df_sorted[available_columns].copy()

            # Добавляем комментарии из нашей базы данных
            df_result = df_sorted[available_columns].copy()

            # Удаляем дубликаты
            df_result = df_result.drop_duplicates()

            # === ДОБАВЛЯЕМ КОЛОНКУ С КОММЕНТАРИЯМИ ===
            # Создаем словарь комментариев по артикулам из нашей базы
            comments_dict = {}
            for article in Form16Article.objects.filter(
                user=request.user, is_active=True
            ):
                if article.comments:  # Только если есть комментарии
                    comments_dict[str(article.article_wb).strip()] = article.comments

            # Добавляем колонку "Комментарии" после "Доступность"
            # Находим индекс колонки "Доступность"
            if "Доступность" in df_result.columns:
                availability_idx = df_result.columns.get_loc("Доступность")
                # Вставляем колонку "Комментарии" после "Доступность"
                df_result.insert(availability_idx + 1, "Комментарии", "")
            else:
                # Если нет колонки "Доступность", добавляем в конец
                df_result["Комментарии"] = ""

            # Заполняем комментарии из нашей базы
            def get_comments(row):
                article_wb = (
                    str(row["Артикул WB"]).strip() if "Артикул WB" in row else ""
                )
                return comments_dict.get(article_wb, "")

            df_result["Комментарии"] = df_result.apply(get_comments, axis=1)

            # === ИСПРАВЛЕНИЕ 1: Добавляем пустую колонку "Отгрузка ФБО" ===
            df_result["Отгрузка ФБО"] = ""

            # Переупорядочиваем колонки: новая колонка в конце
            columns_order = [
                "Артикул продавца",
                "Артикул WB",
                "Размер",
                "Доступность",
                "Комментарии",
                "Заказали, шт",
                "Выкупили, шт",
                "Процент выкупа",
                "Остатки на текущий день, шт",
                "Отгрузка ФБО",
            ]

            # Оставляем только существующие колонки в нужном порядке
            final_columns = [col for col in columns_order if col in df_result.columns]
            df_result = df_result[final_columns]

            # Создаем Excel файл
            wb = Workbook()
            ws = wb.active
            ws.title = f"ФОРМИРОВАНИЕ_ФБО"

            # === РАСЧЕТ СТРОК ДЛЯ ФОРМАТИРОВАНИЯ ===

            # 1. Заголовок отчета (1 строка)
            ws.append(["ОТЧЕТ 15 ЛОКО ПОРАЗМЕРНЫЙ ПО ОСТАТКАМ (Форма 16)"])
            title_row = 1

            # 2. Информация (2 строки)
            ws.append(
                [f"Дата формирования: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}"]
            )
            ws.append([f"Пользователь: {request.user.username}"])
            ws.append([])  # Пустая строка

            # 3. Статистика
            ws.append(["СТАТИСТИКА:"])
            ws.append([f"Файл: {file.name}"])
            ws.append([f"Страница: {target_sheet}"])
            ws.append([f"Искали артикулов: {diagnostic_info['our_articles_count']}"])
            ws.append([f"Найдено артикулов: {found_count}"])
            ws.append([f"Всего строк: {len(df_result)}"])
            ws.append([])  # Пустая строка

            # 4. Предупреждения о ненайденных артикулах
            if diagnostic_info["not_found_articles"]:
                ws.append(["ВНИМАНИЕ: Следующие артикулы не найдены в файле:"])
                for art in diagnostic_info["not_found_articles"]:
                    ws.append([f"• {art}"])
                ws.append([])  # Пустая строка

            # 5. Еще одна пустая строка перед таблицей
            ws.append([])

            # 6. ЗАГОЛОВКИ ТАБЛИЦЫ - запоминаем номер этой строки!
            headers = list(df_result.columns)
            ws.append(headers)
            HEADER_ROW = ws.max_row  # Это строка с заголовками таблицы!
            print(f"Заголовки таблицы на строке: {HEADER_ROW}")

            # 7. ДАННЫЕ ТАБЛИЦЫ
            for _, row in df_result.iterrows():
                ws.append(row.tolist())

            # Начало данных (первая строка после заголовков)
            DATA_START_ROW = HEADER_ROW + 1
            print(f"Данные начинаются со строки: {DATA_START_ROW}")

            # === ФОРМАТИРОВАНИЕ ===

            # 1. Заголовок отчета (строка 1)
            ws.merge_cells(
                start_row=1, start_column=1, end_row=1, end_column=len(headers)
            )
            title_fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            title_font = Font(color="FFFFFF", bold=True, size=14)

            ws.cell(row=title_row, column=1).fill = title_fill
            ws.cell(row=title_row, column=1).font = title_font
            ws.cell(row=title_row, column=1).alignment = Alignment(
                horizontal="center", vertical="center"
            )

            # 2. Заголовки таблицы - ФОРМАТИРУЕМ ИМЕННО HEADER_ROW
            header_fill = PatternFill(
                start_color="4F81BD", end_color="4F81BD", fill_type="solid"
            )
            header_font = Font(color="FFFFFF", bold=True, size=11)
            header_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            print(f"Форматируем заголовки таблицы на строке {HEADER_ROW}: {headers}")

            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=HEADER_ROW, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = header_border
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

            # 3. Данные таблицы (цветовое кодирование по артикулам)
            colors = [
                PatternFill(
                    start_color="DAE8FC", end_color="DAE8FC", fill_type="solid"
                ),
                PatternFill(
                    start_color="D5E8D4", end_color="D5E8D4", fill_type="solid"
                ),
                PatternFill(
                    start_color="FFE6CC", end_color="FFE6CC", fill_type="solid"
                ),
                PatternFill(
                    start_color="F8CECC", end_color="F8CECC", fill_type="solid"
                ),
                PatternFill(
                    start_color="E1D5E7", end_color="E1D5E7", fill_type="solid"
                ),
                PatternFill(
                    start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
                ),
                PatternFill(
                    start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"
                ),
                PatternFill(
                    start_color="D0E0E3", end_color="D0E0E3", fill_type="solid"
                ),
                PatternFill(
                    start_color="E8D5E7", end_color="E8D5E7", fill_type="solid"
                ),
                PatternFill(
                    start_color="D4E6F1", end_color="D4E6F1", fill_type="solid"
                ),
                PatternFill(
                    start_color="E8F5E8", end_color="E8F5E8", fill_type="solid"
                ),
                PatternFill(
                    start_color="FFF8E1", end_color="FFF8E1", fill_type="solid"
                ),
                PatternFill(
                    start_color="F3E5F5", end_color="F3E5F5", fill_type="solid"
                ),
                PatternFill(
                    start_color="E0F2F1", end_color="E0F2F1", fill_type="solid"
                ),
                PatternFill(
                    start_color="FFEBEE", end_color="FFEBEE", fill_type="solid"
                ),
            ]

            # Раскрашиваем строки данных по артикулам
            current_article = None
            color_index = 0
            article_colors = {}

            print(f"Форматируем данные с строки {DATA_START_ROW} до {ws.max_row}")

            for row in range(DATA_START_ROW, ws.max_row + 1):
                article = ws.cell(
                    row=row, column=2
                ).value  # Колонка 'Артикул WB' (вторая колонка)
                if article:
                    article_str = str(article).strip()
                    if article_str != current_article:
                        current_article = article_str
                        if current_article not in article_colors:
                            article_colors[current_article] = colors[
                                color_index % len(colors)
                            ]
                            color_index += 1

                    fill_color = article_colors.get(current_article, colors[0])

                    # Форматируем всю строку данных ВКЛЮЧАЯ колонку "Отгрузка ФБО"
                    for col in range(
                        1, len(headers) + 1
                    ):  # Это ВКЛЮЧАЕТ последнюю колонку
                        cell = ws.cell(row=row, column=col)
                        cell.fill = fill_color
                        cell.border = Border(
                            left=Side(style="thin"),
                            right=Side(style="thin"),
                            top=Side(style="thin"),
                            bottom=Side(style="thin"),
                        )

                        # Выравнивание
                        if (
                            col >= 5 and col <= 8
                        ):  # Числовые колонки (Заказали, Выкупили, Процент, Остатки)
                            try:
                                if cell.value is not None and str(cell.value).strip():
                                    cell.alignment = Alignment(
                                        horizontal="right", vertical="center"
                                    )
                            except:
                                cell.alignment = Alignment(
                                    horizontal="left", vertical="center"
                                )
                        else:
                            cell.alignment = Alignment(
                                horizontal="left", vertical="center"
                            )

            # 4. Устанавливаем ширину для всех колонок (включая "Отгрузка ФБО")
            for col_idx in range(1, len(headers) + 1):
                max_length = 0
                column_letter = get_column_letter(col_idx)

                # Проверяем заголовок
                header_cell = ws.cell(row=HEADER_ROW, column=col_idx)
                max_length = max(max_length, len(str(header_cell.value or "")))

                # Проверяем данные
                for row in range(DATA_START_ROW, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    try:
                        cell_value = str(cell.value or "")
                        max_length = max(max_length, len(cell_value))
                    except:
                        pass

                # Устанавливаем ширину
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            # 5. Особенная ширина для колонки "Отгрузка ФБО" (немного шире для удобства ввода)
            last_col_letter = get_column_letter(len(headers))
            ws.column_dimensions[last_col_letter].width = (
                18  # Немного шире для удобства
            )

            # Сохраняем в буфер
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            # Возвращаем файл
            filename = f"ОБОРАЧИВАЕМОСТЬ_Форма16_{found_count}_артикулов.xlsx"
            response = FileResponse(
                buffer,
                as_attachment=True,
                filename=filename,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            return response

        except Exception as e:
            messages.error(request, f"Ошибка при обработке файла: {str(e)}")
            import traceback

            traceback.print_exc()

    # Для GET запроса
    from forms_app.forms import Form16UploadForm

    form = Form16UploadForm()

    context = {
        "form": form,
        "articles": articles_qs,
        "articles_count": len(articles),
    }

    return render(request, "forms_app/form16_generate_report.html", context)


@login_required
def form16_delete_all(request):
    """Удаление всех артикулов пользователя"""
    if request.method == "POST":
        count = Form16Article.objects.filter(user=request.user).count()
        Form16Article.objects.filter(user=request.user).delete()
        messages.success(request, f"Удалено {count} артикулов")
        return redirect("forms_app:form16_edit_table")

    return render(
        request,
        "forms_app/form16_delete_all.html",
        {"articles_count": Form16Article.objects.filter(user=request.user).count()},
    )
