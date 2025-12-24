# views.py - СОВМЕСТНАЯ ВЕРСИЯ (старые функции + новый алгоритм)

import os
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, FileResponse
from django.contrib import messages
from django.conf import settings
from forms_app.forms import PatternForm, CuttingForm
from forms_app.models import Pattern15

# ==================== ИМПОРТ ДЛЯ НОВОГО АЛГОРИТМА ====================
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from ortools.sat.python import cp_model
from io import BytesIO
import time
import random

# ==================== СТАРЫЕ ФУНКЦИИ (БЕЗ ИЗМЕНЕНИЙ) ====================


@login_required
def form15_view(request):
    """Основная страница формы 15 - СТАРАЯ ВЕРСИЯ БЕЗ ИЗМЕНЕНИЙ"""
    patterns = Pattern15.objects.filter(user=request.user)

    if request.method == "POST":
        # Если добавляем новое лекало
        if "add_pattern" in request.POST:
            pattern_form = PatternForm(request.POST)
            if pattern_form.is_valid():
                pattern = pattern_form.save(commit=False)
                pattern.user = request.user
                pattern.save()
                messages.success(request, f"Лекало '{pattern.name}' добавлено")
                return redirect("forms_app:form15_view")

        # Если запускаем расчет раскроя
        elif "calculate" in request.POST:
            cutting_form = CuttingForm(request.POST)
            if cutting_form.is_valid():
                fabric_width = cutting_form.cleaned_data["fabric_width"]
                num_sets = cutting_form.cleaned_data["num_sets"]
                output_format = cutting_form.cleaned_data["output_format"]

                # Проверяем, есть ли лекала
                if not patterns.exists():
                    messages.error(request, "Добавьте хотя бы одно лекало")
                    return redirect("forms_app:form15_view")

                # Сохраняем параметры в сессии
                request.session["fabric_width"] = fabric_width
                request.session["num_sets"] = num_sets
                request.session["output_format"] = output_format

                # Перенаправляем на страницу расчета
                return redirect("forms_app:form15_calculate")

    else:
        pattern_form = PatternForm()
        cutting_form = CuttingForm(
            initial={"fabric_width": 1500, "num_sets": 1, "output_format": "pdf"}
        )

    return render(
        request,
        "forms_app/form15_view.html",
        {
            "patterns": patterns,
            "pattern_form": pattern_form,
            "cutting_form": cutting_form,
            "patterns_count": patterns.count(),
        },
    )


@login_required
def form15_edit_pattern(request, pk):
    """Редактирование лекала - СТАРАЯ ВЕРСИЯ БЕЗ ИЗМЕНЕНИЙ"""
    pattern = get_object_or_404(Pattern15, pk=pk, user=request.user)

    if request.method == "POST":
        form = PatternForm(request.POST, instance=pattern)
        if form.is_valid():
            form.save()
            messages.success(request, f"Лекало '{pattern.name}' обновлено")
            return redirect("forms_app:form15_view")
    else:
        form = PatternForm(instance=pattern)

    return render(
        request,
        "forms_app/form15_edit_pattern.html",
        {"form": form, "pattern": pattern},
    )


@login_required
def form15_delete_pattern(request, pk):
    """Удаление лекала - СТАРАЯ ВЕРСИЯ БЕЗ ИЗМЕНЕНИЙ"""
    pattern = get_object_or_404(Pattern15, pk=pk, user=request.user)

    if request.method == "POST":
        name = pattern.name
        pattern.delete()
        messages.success(request, f"Лекало '{name}' удалено")
        return redirect("forms_app:form15_view")

    return render(request, "forms_app/form15_delete_pattern.html", {"pattern": pattern})


@login_required
def form15_clear_all(request):
    """Очистка всех лекал пользователя - СТАРАЯ ВЕРСИЯ БЕЗ ИЗМЕНЕНИЙ"""
    if request.method == "POST":
        count = Pattern15.objects.filter(user=request.user).count()
        Pattern15.objects.filter(user=request.user).delete()
        messages.success(request, f"Удалено {count} лекал")
        return redirect("forms_app:form15_view")

    return render(
        request,
        "forms_app/form15_clear_all.html",
        {"patterns_count": Pattern15.objects.filter(user=request.user).count()},
    )


@login_required
def form15_import_excel(request):
    """Импорт лекал из Excel - СТАРАЯ ВЕРСИЯ БЕЗ ИЗМЕНЕНИЙ"""
    if request.method == "POST" and request.FILES.get("excel_file"):
        excel_file = request.FILES["excel_file"]

        try:
            # Читаем Excel
            df = pd.read_excel(excel_file)

            # Проверяем обязательные колонки
            required_cols = ["имя", "ширина", "высота"]
            for col in required_cols:
                if col not in df.columns:
                    messages.error(request, f"В файле отсутствует колонка '{col}'")
                    return redirect("forms_app:form15_view")

            imported_count = 0
            for _, row in df.iterrows():
                try:
                    name = str(row["имя"]).strip()
                    width = int(float(row["ширина"]))
                    height = int(float(row["высота"]))

                    if name and width > 0 and height > 0:
                        # Проверяем, нет ли уже такого лекала
                        if not Pattern15.objects.filter(
                            user=request.user, name=name, width=width, height=height
                        ).exists():
                            Pattern15.objects.create(
                                user=request.user,
                                name=name,
                                width=width,
                                height=height,
                            )
                            imported_count += 1
                except (ValueError, TypeError):
                    continue

            messages.success(request, f"Импортировано {imported_count} лекал")

        except Exception as e:
            messages.error(request, f"Ошибка при импорте: {str(e)}")

    return redirect("forms_app:form15_view")


# ==================== НОВЫЙ АЛГОРИТМ (ТОЛЬКО РАСЧЕТ) ====================


def optimize_packing(patterns, fabric_width, time_limit=180):
    """
    Оптимальная упаковка прямоугольников с фиксированной ориентацией
    patterns: список кортежей (ширина, высота)
    fabric_width: ширина полотна
    time_limit: ограничение по времени в секундах
    """
    n = len(patterns)

    # Максимальная возможная длина (сумма всех высот)
    max_length = sum(h for _, h in patterns)

    # Создаем модель
    model = cp_model.CpModel()

    # Переменные: координаты левого нижнего угла
    x = [model.new_int_var(0, fabric_width, f"x_{i}") for i in range(n)]
    y = [model.new_int_var(0, max_length, f"y_{i}") for i in range(n)]

    # Переменная для длины полотна
    length = model.new_int_var(0, max_length, "length")

    # Добавляем ограничения на размещение
    for i in range(n):
        w_i, h_i = patterns[i]

        # Лекало должно помещаться по ширине
        model.add(x[i] + w_i <= fabric_width)

        # Лекало должно помещаться по длине
        model.add(y[i] + h_i <= length)

        # Ограничения непересечения для каждой пары лекал
        for j in range(i + 1, n):
            w_j, h_j = patterns[j]

            # Создаем булевы переменные для 4 возможных положений
            left = model.new_bool_var(f"left_{i}_{j}")
            right = model.new_bool_var(f"right_{i}_{j}")
            below = model.new_bool_var(f"below_{i}_{j}")
            above = model.new_bool_var(f"above_{i}_{j}")

            # Определяем условия
            # i левее j
            model.add(x[i] + w_i <= x[j]).only_enforce_if(left)
            # j левее i
            model.add(x[j] + w_j <= x[i]).only_enforce_if(right)
            # i ниже j
            model.add(y[i] + h_i <= y[j]).only_enforce_if(below)
            # j ниже i
            model.add(y[j] + h_j <= y[i]).only_enforce_if(above)

            # Хотя бы одно условие должно выполняться
            model.add_bool_or([left, right, below, above])

    # Длина полотна должна быть не меньше максимальной координаты Y+высота
    for i in range(n):
        _, h_i = patterns[i]
        model.add(length >= y[i] + h_i)

    # Минимизируем длину полотна
    model.minimize(length)

    # Решаем
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = 8  # Используем все ядра

    status = solver.solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, None

    # Собираем результаты
    placements = []
    min_length = solver.value(length)

    for i in range(n):
        placements.append(
            {
                "id": i,
                "x": solver.value(x[i]),
                "y": solver.value(y[i]),
                "width": patterns[i][0],
                "height": patterns[i][1],
                "rotated": False,
            }
        )

    # Сортируем по Y (снизу вверх), затем по X (слева направо)
    placements.sort(key=lambda p: (p["y"], p["x"]))

    return placements, min_length


def create_visualization(
    placements,
    fabric_width,
    min_length,
    num_sets,
    display_numbers,
    legend_numbers,
    legend_names,
    legend_info,
    # Параметры настройки
    font_coefficient=0.07,  # коэффициент размера шрифта (от меньшей стороны)
    min_font_size=8,  # минимальный размер шрифта
    bbox_padding=0.10,  # внутренний отступ фона
    bbox_linewidth=0.3,  # толщина рамки
    bbox_alpha=0.85,  # прозрачность фона
):
    """
    Создание визуализации:
    - На рисунке: номера из первого комплекта (повторяются для второго)
    - Для зеркальных лекал добавляется буква "З" рядом с номером
    - В легенде: только уникальные лекала первого комплекта
    - Цвета: разные для разных комплектов
    """
    fig = plt.figure(figsize=(20, 14))
    gs = fig.add_gridspec(1, 2, width_ratios=[0.25, 0.75], wspace=0.05)
    ax_legend = fig.add_subplot(gs[0])
    ax_graph = fig.add_subplot(gs[1])

    # ===== ГРАФИК С РАСКЛАДКОЙ =====
    padding = 50
    ax_graph.add_patch(
        plt.Rectangle(
            (0, 0),
            fabric_width,
            min_length,
            linewidth=2,
            edgecolor="black",
            facecolor="lightgray",
            alpha=0.3,
        )
    )

    # ===== ЦВЕТА ДЛЯ КОМПЛЕКТОВ =====
    # Разные цвета для разных комплектов
    if num_sets == 1:
        set_colors = ["#FF9999"]  # один цвет для одного комплекта
    elif num_sets == 2:
        set_colors = ["#FF9999", "#99FF99"]  # разные цвета для двух комплектов
    elif num_sets == 3:
        set_colors = ["#FF9999", "#99FF99", "#9999FF"]  # для трех комплектов
    else:
        # Генерируем цвета для большего количества комплектов
        colors = plt.cm.Set3(np.linspace(0, 1, num_sets))
        set_colors = [plt.colors.to_hex(c) for c in colors]

    # ===== РИСУЕМ ЛЕКАЛА =====
    for idx, p in enumerate(placements):
        # Получаем номер для отображения (из display_numbers)
        if idx < len(display_numbers):
            pattern_number = display_numbers[idx]
        else:
            pattern_number = idx + 1

        # Определяем цвет в зависимости от номера комплекта
        set_num = p.get("set_number", 1) - 1  # 0-based индекс
        color = set_colors[set_num % len(set_colors)]

        # Определяем, является ли лекало зеркальным
        is_mirrored = p.get("is_mirrored", False)

        # Прямоугольник
        rect = plt.Rectangle(
            (p["x"], p["y"]),
            p["width"],
            p["height"],
            linewidth=0.8,
            edgecolor="black",
            facecolor=color,
            alpha=0.7,
        )
        ax_graph.add_patch(rect)

        # ===== НАСТРОЙКА ШРИФТА И ФОНА =====
        # Рассчитываем размер шрифта
        min_side = min(p["height"], p["width"])
        fontsize = max(min_font_size, min_side * font_coefficient)

        # Текст для отображения: добавляем "З" для зеркальных лекал
        if is_mirrored:
            display_text = f"{pattern_number}з"  # Добавляем букву "З"
        else:
            display_text = f"{pattern_number}"  # Обычный номер

        # Адаптируем параметры для маленьких лекал
        if min_side < 50:
            bbox_padding_adj = 0.05
            bbox_linewidth_adj = 0.2
            # Для очень маленьких лекал уменьшаем шрифт для буквы "З"
            if is_mirrored and min_side < 40:
                display_text = (
                    f"{pattern_number}*"  # Заменяем "З" на "*" для очень маленьких
                )
        elif min_side < 100:
            bbox_padding_adj = bbox_padding * 0.7
            bbox_linewidth_adj = bbox_linewidth * 0.7
        else:
            bbox_padding_adj = bbox_padding
            bbox_linewidth_adj = bbox_linewidth

        # Цвет текста: красный для зеркальных лекал
        text_color = "red" if is_mirrored else "black"

        # Отображаем текст с фоном
        ax_graph.text(
            p["x"] + p["width"] / 2,
            p["y"] + p["height"] / 2,
            display_text,
            ha="center",
            va="center",
            fontsize=fontsize,
            fontweight="bold",
            color=text_color,
            bbox=dict(
                boxstyle=f"round,pad={bbox_padding_adj}",
                facecolor="white",
                edgecolor=(
                    "red" if is_mirrored else "black"
                ),  # Красная рамка для зеркальных
                alpha=bbox_alpha,
                linewidth=bbox_linewidth_adj * (1.5 if is_mirrored else 1.0),
            ),
        )

    # Настройки графика
    ax_graph.set_xlim(-padding, fabric_width + padding)
    ax_graph.set_ylim(-padding, min_length + padding)
    ax_graph.set_aspect("equal")
    ax_graph.invert_yaxis()
    ax_graph.set_xlabel("Ширина (мм)", fontsize=12)
    ax_graph.set_ylabel("Длина (мм)", fontsize=12)

    # ===== РАССЧИТЫВАЕМ СТАТИСТИКУ ДЛЯ ЗАГОЛОВКА =====
    total_patterns = len(placements)
    total_area = sum(p["width"] * p["height"] for p in placements)
    fabric_area = fabric_width * min_length
    utilization = (total_area / fabric_area * 100) if fabric_area > 0 else 0
    waste_area = fabric_area - total_area
    waste_m2 = waste_area / 1000000

    # Считаем зеркальные лекала
    mirrored_count = sum(1 for p in placements if p.get("is_mirrored", False))

    # Заголовок с информацией о комплектах
    if num_sets > 1:
        sets_info = (
            f"РАСКЛАДКА ДЛЯ {num_sets} КОМПЛЕКТОВ (цвета обозначают разные комплекты)"
        )
    else:
        sets_info = "РАСКЛАДКА ДЛЯ 1 КОМПЛЕКТА"

    ax_graph.set_title(
        f"ОПТИМАЛЬНАЯ РАСКЛАДКА ЛЕКАЛ\n"
        f"{sets_info}\n"
        f"Ширина ткани: {fabric_width} мм | Длина раскроя: {min_length} мм | "
        f"Площадь ткани: {fabric_area/1000000:.3f} м²\n"
        f"Всего лекал: {total_patterns} | Зеркальных: {mirrored_count} | "
        f"Использование ткани: {utilization:.1f}% | Отходы: {waste_m2:.3f} м²",
        fontsize=13,
        pad=20,
        fontweight="bold",
        linespacing=1.5,
    )

    ax_graph.grid(True, linestyle=":", alpha=0.3, linewidth=0.5)

    # ===== ЛЕГЕНДА (ТОЛЬКО УНИКАЛЬНЫЕ ЛЕКАЛА ПЕРВОГО КОМПЛЕКТА) =====
    ax_legend.axis("off")

    # Заголовок легенды
    unique_count = len(legend_numbers)
    legend_title = f"ЛЕГЕНДА ({unique_count} уникальных лекал)\n" + "═" * 30
    ax_legend.text(
        0.02,
        0.97,
        legend_title,
        transform=ax_legend.transAxes,
        fontsize=14,
        fontweight="bold",
        verticalalignment="top",
        color="darkblue",
    )

    # Подзаголовок с пояснением
    explanation_lines = []
    if num_sets > 1:
        explanation_lines.append(
            f"Примечание: В раскладке {num_sets} комплекта одинаковых лекал"
        )
        explanation_lines.append(
            "На рисунке одинаковые номера повторяются для каждого комплекта"
        )
        explanation_lines.append("Цвета показывают принадлежность к разным комплектам")

    # Добавляем пояснение про зеркальные лекала
    if any(info["is_mirrored"] for info in legend_info.values()):
        explanation_lines.append(
            "Лекала с буквой 'З' или '*' - зеркальные (требуют переворота)"
        )
        explanation_lines.append(
            "Красный цвет рамки/текста также указывает на зеркальное лекало"
        )

    if explanation_lines:
        explanation = "\n".join(explanation_lines)
        ax_legend.text(
            0.02,
            0.93 if len(explanation_lines) <= 3 else 0.30,
            explanation,
            transform=ax_legend.transAxes,
            fontsize=12,
            style="italic",
            color="red",
            verticalalignment="top",
        )
        y_pos_list = 0.87 if len(explanation_lines) <= 3 else 0.84
    else:
        y_pos_list = 0.92

    line_height = 0.030

    ax_legend.text(
        0.02,
        y_pos_list,
        "УНИКАЛЬНЫЕ ЛЕКАЛА (ПЕРВЫЙ КОМПЛЕКТ):",
        transform=ax_legend.transAxes,
        fontsize=12,
        fontweight="bold",
        color="darkgreen",
    )

    y_pos_list -= line_height * 1.5

    # Отображаем уникальные лекала первого комплекта
    displayed_items = 0

    # Сортируем по номерам
    sorted_pairs = sorted(zip(legend_numbers, legend_names))

    for legend_num, legend_name in sorted_pairs:
        if legend_num not in legend_info:
            continue

        info = legend_info[legend_num]
        displayed_items += 1

        # Форматируем имя
        display_name = legend_name
        if len(display_name) > 25:
            display_name = display_name[:22] + "..."

        dimensions = f" ({info['width']}×{info['height']} мм)"

        # Добавляем пометку "ЗЕРКАЛО" для зеркальных лекал
        mirrored_mark = " [ЗЕРКАЛО]" if info["is_mirrored"] else ""

        # Отображаем
        legend_item = f"#{legend_num:02d} — {display_name}{dimensions}{mirrored_mark}"

        # Чередуем цвета фона для читаемости
        # Красный фон для зеркальных лекал
        if info["is_mirrored"]:
            bg_color = "mistyrose"
            border_color = "red"
        else:
            bg_color = "lightyellow" if legend_num % 2 == 0 else "honeydew"
            border_color = "gold" if legend_num % 2 == 0 else "lightgreen"

        ax_legend.text(
            0.02,
            y_pos_list,
            legend_item,
            transform=ax_legend.transAxes,
            fontsize=10,
            verticalalignment="top",
            bbox=dict(
                boxstyle="round,pad=0.15",
                facecolor=bg_color,
                alpha=0.8,
                edgecolor=border_color,
                linewidth=0.6,
            ),
        )

        y_pos_list -= line_height * 1.1

        # Проверяем, помещаются ли еще строки
        if y_pos_list < 0.05:
            remaining = len(sorted_pairs) - displayed_items
            if remaining > 0:
                ax_legend.text(
                    0.02,
                    y_pos_list - line_height * 0.5,
                    f"... и еще {remaining} лекал",
                    transform=ax_legend.transAxes,
                    fontsize=9,
                    style="italic",
                    color="gray",
                )
            break

    # Информация о дате генерации
    footer_text = f"Сгенерировано: {time.strftime('%d.%m.%Y %H:%M')}"
    ax_legend.text(
        0.02,
        0.02,
        footer_text,
        transform=ax_legend.transAxes,
        fontsize=10,
        style="italic",
        color="black",
    )

    plt.tight_layout()
    return fig


def generate_pdf_response_new(
    placements,
    fabric_width,
    min_length,
    num_sets,
    display_numbers,
    legend_numbers,
    legend_names,
    legend_info,
):
    """
    Генерация PDF с новой системой нумерации
    """
    buffer = BytesIO()

    # Создаем визуализацию
    fig = create_visualization(
        placements,
        fabric_width,
        min_length,
        num_sets,
        display_numbers,  # номера для рисунка
        legend_numbers,  # номера для легенды
        legend_names,  # имена для легенды
        legend_info,  # информация о лекалах
    )

    # Сохраняем в PDF
    with PdfPages(buffer) as pdf:
        pdf.savefig(fig, bbox_inches="tight", dpi=300)

    plt.close(fig)
    buffer.seek(0)

    filename = f"cutting_layout_{num_sets}_sets.pdf"
    return FileResponse(
        buffer,
        as_attachment=True,
        filename=filename,
        content_type="application/pdf",
    )


def generate_excel_response_new(
    placements, display_numbers, all_names, fabric_width, min_length, num_sets
):
    """
    Генерация Excel файла для нового алгоритма
    """
    # Собираем данные
    data = []
    for i, p in enumerate(placements):
        if i < len(display_numbers) and i < len(all_names):
            pattern_num = display_numbers[i]
            name = all_names[i]
        else:
            pattern_num = i + 1
            name = f"Лекало_{i+1}"

        # Определяем, является ли зеркальным
        is_mirrored = p.get("is_mirrored", False)

        data.append(
            {
                "Pattern Number": pattern_num,
                "Pattern Name": name,
                "Width (mm)": p["width"],
                "Height (mm)": p["height"],
                "Position X (mm)": p["x"],
                "Position Y (mm)": p["y"],
                "Is Mirrored": "Да" if is_mirrored else "Нет",
            }
        )

    df = pd.DataFrame(data)

    # Рассчитываем статистику
    total_area = sum(p["width"] * p["height"] for p in placements)
    fabric_area = fabric_width * min_length
    utilization = round(total_area / fabric_area * 100, 1) if fabric_area > 0 else 0
    mirrored_count = sum(1 for p in placements if p.get("is_mirrored", False))

    # Создаем Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Layout", index=False)

        summary_data = {
            "Parameter": [
                "Fabric Width",
                "Fabric Length",
                "Number of Sets",
                "Total Patterns",
                "Mirrored Patterns",
                "Area Utilization",
            ],
            "Value": [
                f"{fabric_width} mm",
                f"{min_length} mm",
                num_sets,
                len(data),
                mirrored_count,
                f"{utilization}%",
            ],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

    output.seek(0)

    filename = f"cutting_layout_{num_sets}_sets.xlsx"
    response = FileResponse(
        output,
        as_attachment=True,
        filename=filename,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return response


@login_required
def form15_calculate(request):
    """Расчет раскроя с новым алгоритмом"""
    print("=== form15_calculate: Начало работы (новый алгоритм) ===")

    # Получаем параметры
    try:
        fabric_width = int(request.POST.get("fabric_width", 1500))
        num_sets = int(request.POST.get("num_sets", 1))
        output_format = request.POST.get("output_format", "pdf")

        print(
            f"Параметры: ширина={fabric_width}, комплектов={num_sets}, формат={output_format}"
        )
    except (ValueError, TypeError) as e:
        print(f"Ошибка получения параметров: {e}")
        messages.error(request, "Неверные параметры расчета")
        return redirect("forms_app:form15_view")

    # Получаем лекала пользователя с сортировкой по номеру
    patterns = Pattern15.objects.filter(user=request.user).order_by("pattern_number")
    print(f"Найдено лекал в базе: {patterns.count()}")

    if not patterns.exists():
        messages.error(request, "Добавьте хотя бы одно лекало для расчета")
        return redirect("forms_app:form15_view")

    try:
        # Подготавливаем данные лекал
        base_patterns = []
        base_numbers = []
        base_names = []
        is_mirrored_list = []  # Новый список: является ли лекало зеркальным

        for pattern in patterns:
            base_patterns.append((pattern.width, pattern.height))
            base_numbers.append(pattern.pattern_number)
            base_names.append(pattern.name)

            # Проверяем, является ли лекало зеркальным
            is_mirrored = any(
                keyword.lower() in pattern.name.lower()
                for keyword in [
                    "зеркало",
                    "зеркальное",
                    "зеркальная",
                    "mirror",
                    "mirrored",
                ]
            )
            is_mirrored_list.append(is_mirrored)

            print(
                f"Лекало #{pattern.pattern_number}: {pattern.name} - {pattern.width}x{pattern.height} мм"
                f"{' [ЗЕРКАЛО]' if is_mirrored else ''}"
            )

        # Формируем полный список лекал для всех комплектов
        all_patterns = base_patterns * num_sets
        all_numbers = base_numbers * num_sets  # Номера будут повторяться
        all_names = base_names * num_sets
        all_mirrored = (
            is_mirrored_list * num_sets
        )  # Признаки зеркальности тоже повторяются

        print(f"Всего лекал для раскроя: {len(all_patterns)}")
        print(f"Уникальных лекал: {len(base_patterns)}")
        print(f"Номера для первого комплекта: {base_numbers}")
        print(
            f"Зеркальные лекала: {[name for name, is_mirrored in zip(base_names, is_mirrored_list) if is_mirrored]}"
        )

        # Проверяем ширину
        for i, (w, h) in enumerate(all_patterns):
            if w > fabric_width:
                error_msg = f"Лекало '{all_names[i]}' ({w} мм) шире полотна ({fabric_width} мм)!"
                print(f"ОШИБКА: {error_msg}")
                messages.error(request, error_msg)
                return redirect("forms_app:form15_view")

        # ===== ИСПОЛЬЗУЕМ НОВЫЙ АЛГОРИТМ =====
        print("Начинаем оптимизацию с OR-Tools...")
        placements, min_length = optimize_packing(
            all_patterns, fabric_width, time_limit=30
        )

        if not placements:
            print("Не удалось найти решение")
            messages.error(request, "Не удалось найти решение за отведенное время")
            return redirect("forms_app:form15_view")

        print(f"Упаковка завершена. Минимальная длина: {min_length} мм")
        print(f"Количество упакованных элементов: {len(placements)}")

        # ===== ПРОСТАЯ СИСТЕМА: НОМЕРА ТОЛЬКО ИЗ ПЕРВОГО КОМПЛЕКТА =====
        # Каждому размещенному лекалу назначаем номер из первого комплекта
        for p in placements:
            idx = p["id"]  # индекс в all_patterns
            if idx < len(all_numbers):
                # Берем номер из первого комплекта (модульная арифметика)
                base_idx = idx % len(base_numbers)
                p["display_number"] = base_numbers[
                    base_idx
                ]  # номер из первого комплекта
                p["base_name"] = base_names[base_idx]  # имя из первого комплекта
                p["is_mirrored"] = is_mirrored_list[base_idx]  # является ли зеркальным
                p["set_number"] = (
                    idx // len(base_numbers)
                ) + 1  # номер комплекта (1, 2, ...)
            else:
                p["display_number"] = idx + 1
                p["base_name"] = f"Лекало_{idx+1}"
                p["is_mirrored"] = False
                p["set_number"] = 1

        # ===== ПОДГОТОВКА ДАННЫХ ДЛЯ ЛЕГЕНДЫ =====
        # В легенде показываем ТОЛЬКО уникальные лекала первого комплекта
        legend_numbers = base_numbers.copy()  # Номера из первого комплекта
        legend_names = base_names.copy()  # Имена из первого комплекта

        # Создаем информацию для легенды
        legend_info = {}
        for i, (num, name) in enumerate(zip(base_numbers, base_names)):
            legend_info[num] = {
                "name": name,
                "display_name": name,  # Без указания комплекта
                "width": base_patterns[i][0],
                "height": base_patterns[i][1],
                "is_mirrored": is_mirrored_list[
                    i
                ],  # Добавляем информацию о зеркальности
            }

        print(f"Номера в легенде: {legend_numbers}")
        print(f"Всего записей в легенде: {len(legend_numbers)}")
        print(
            f"Зеркальные лекала в легенде: {[num for num, info in legend_info.items() if info['is_mirrored']]}"
        )

        # ===== СОБИРАЕМ ДАННЫЕ ДЛЯ РИСУНКА =====
        # На рисунке будут те же номера, что и в легенде (повторяющиеся)
        display_numbers_on_chart = [p["display_number"] for p in placements]
        print(f"Пример номеров на рисунке: {display_numbers_on_chart[:15]}")

        # Генерация файла
        if output_format == "pdf":
            print("Генерация PDF файла...")
            response = generate_pdf_response_new(
                placements,
                fabric_width,
                min_length,
                num_sets,
                display_numbers_on_chart,  # номера для рисунка (из первого комплекта)
                legend_numbers,  # номера для легенды (только первый комплект)
                legend_names,  # имена для легенды (только первый комплект)
                legend_info,  # информация о лекалах
            )
        else:
            print("Генерация Excel файла...")
            # Для Excel можно сохранить полную информацию
            response = generate_excel_response_new(
                placements,
                [p["display_number"] for p in placements],
                [p["base_name"] for p in placements],
                fabric_width,
                min_length,
                num_sets,
            )

        print("=== form15_calculate: Файл успешно сгенерирован ===")
        return response

    except Exception as e:
        print(f"=== КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        import traceback

        traceback.print_exc()
        messages.error(request, f"Ошибка при расчете: {str(e)}")
        return redirect("forms_app:form15_view")
