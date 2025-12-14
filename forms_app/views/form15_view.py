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
from rectpack import newPacker, PackingMode, PackingBin
import random
import json
from io import BytesIO


@login_required
def form15_view(request):
    """Основная страница формы 15"""
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
    """Редактирование лекала"""
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
    """Удаление лекала"""
    pattern = get_object_or_404(Pattern15, pk=pk, user=request.user)

    if request.method == "POST":
        name = pattern.name
        pattern.delete()
        messages.success(request, f"Лекало '{name}' удалено")
        return redirect("forms_app:form15_view")

    return render(request, "forms_app/form15_delete_pattern.html", {"pattern": pattern})


def generate_pdf_response(
    packer, all_names, fabric_width, min_length, num_sets, base_names=None
):
    """
    Генерация PDF с нумерацией лекал и легендой.

    base_names — список имён одного комплекта (для легенды).
    Если не передан — используем уникальные имена из all_names.
    """
    from io import BytesIO
    from django.http import FileResponse
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    import random

    # Определяем базовые имена и нумерацию
    if base_names is None:
        # Извлекаем уникальные имена в порядке первого появления
        seen = set()
        base_names = []
        for name in all_names:
            if name not in seen:
                base_names.append(name)
                seen.add(name)

    # Создаём словарь: имя → номер
    name_to_number = {name: i + 1 for i, name in enumerate(base_names)}

    buffer = BytesIO()
    fig, ax = plt.subplots(figsize=(11.7, 8.3))

    # Рисуем полотно
    ax.add_patch(
        patches.Rectangle(
            (0, 0),
            fabric_width,
            min_length,
            linewidth=2,
            edgecolor="black",
            facecolor="none",
        )
    )

    # Рисуем лекала с нумерацией
    random.seed(42)
    for abin in packer:
        for rect in abin:
            name = all_names[rect.rid]
            num = name_to_number[name]
            width, height = rect.width, rect.height
            x, y = rect.x, rect.y

            color = [random.random() * 0.7 + 0.15 for _ in range(3)]
            ax.add_patch(
                patches.Rectangle(
                    (x, y),
                    width,
                    height,
                    linewidth=1,
                    edgecolor="black",
                    facecolor=color,
                    alpha=0.7,
                )
            )

            # Показываем только номер, крупно и чётко
            fontsize = min(12, height * 0.6, width * 0.6)
            fontsize = max(6, fontsize)  # не меньше 6
            ax.text(
                x + width / 2,
                y + height / 2,
                f"#{num}",
                ha="center",
                va="center",
                fontsize=fontsize,
                fontweight="bold",
                color="black",
            )

    # Настройка осей
    ax.set_xlim(0, fabric_width)
    ax.set_ylim(0, min_length)
    # ax.set_aspect("equal")
    ax.invert_yaxis()
    ax.set_xlabel("Ширина (мм)")
    ax.set_ylabel("Длина (мм)")
    # Настраиваем график на заполнение всей области
    ax.margins(0.02)  # небольшие отступы от краёв

    total_patterns = sum(1 for abin in packer for _ in abin)
    total_area = sum(rect.width * rect.height for abin in packer for rect in abin)
    fabric_area = fabric_width * min_length
    utilization = round(total_area / fabric_area * 100, 1) if fabric_area > 0 else 0

    ax.set_title(
        f"Раскладка {num_sets} комплектов ({total_patterns} лекал) | "
        f"Длина: {min_length} мм | Использование: {utilization}%"
    )
    ax.grid(True, linestyle=":", alpha=0.3)

    # === ДОБАВЛЯЕМ ЛЕГЕНДУ ВНИЗУ ===
    legend_text = "Легенда:\n"
    for i, name in enumerate(base_names, 1):
        # Найдём размеры этого лекала (берём первое вхождение)
        idx = all_names.index(name)
        # Чтобы не ломать логику, просто покажем имя
        legend_text += f"  #{i} — {name}\n"

    # Добавляем текст внизу графика
    fig.text(
        0.02,  # x (слева)
        0.02,  # y (снизу)
        legend_text,
        fontsize=9,
        verticalalignment="bottom",
        horizontalalignment="left",
        bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", alpha=0.8),
    )

    ax.set_position([0.3, 0.10, 0.5, 0.87])

    # === ЛЕГЕНДА (только один раз!) ===
    legend_text = "Легенда:\n"
    for i, name in enumerate(base_names, 1):
        legend_text += f"  #{i} — {name}\n"

    fig.text(
        0.02,
        0.02,
        legend_text,
        fontsize=8.5,
        verticalalignment="bottom",
        horizontalalignment="left",
        bbox=dict(boxstyle="round,pad=0.25", facecolor="lightyellow", alpha=0.85),
    )

    # Сохраняем БЕЗ bbox_inches='tight' — это ключевое!
    plt.savefig(buffer, format="pdf", bbox_inches=None, pad_inches=0)
    plt.close(fig)
    buffer.seek(0)

    filename = f"cutting_layout_{num_sets}_sets.pdf"
    return FileResponse(
        buffer,
        as_attachment=True,
        filename=filename,
        content_type="application/pdf",
    )


def generate_excel_response(packer, all_names, fabric_width, min_length, num_sets):
    """Генерация Excel файла для скачивания"""
    from io import BytesIO
    from django.http import FileResponse

    # Собираем данные
    data = []
    for abin in packer:
        for rect in abin:
            data.append(
                {
                    "Pattern Name": all_names[rect.rid],  # Латинские названия колонок
                    "Width (mm)": rect.width,
                    "Height (mm)": rect.height,
                    "Position X (mm)": rect.x,
                    "Position Y (mm)": rect.y,
                }
            )

    df = pd.DataFrame(data)

    # Рассчитываем статистику
    total_area = sum(rect.width * rect.height for abin in packer for rect in abin)
    fabric_area = fabric_width * min_length
    utilization = round(total_area / fabric_area * 100, 1) if fabric_area > 0 else 0

    # Создаем Excel в памяти
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Лист с раскладкой
        df.to_excel(writer, sheet_name="Layout", index=False)

        # Лист со сводкой
        summary_data = {
            "Parameter": [
                "Fabric Width",
                "Fabric Length",
                "Number of Sets",
                "Total Patterns",
                "Area Utilization",
            ],
            "Value": [
                f"{fabric_width} mm",
                f"{min_length} mm",
                num_sets,
                len(data),
                f"{utilization}%",
            ],
            "Unit": ["mm", "mm", "sets", "pcs", "%"],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

    output.seek(0)

    # ✅ Тоже используем FileResponse с латинским именем
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
    """Расчет раскроя и генерация файла"""
    print("=== form15_calculate: Начало работы ===")
    print(f"Метод: {request.method}")
    print(f"POST данные: {dict(request.POST)}")

    # Получаем параметры из POST
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

    # Получаем лекала пользователя
    patterns = Pattern15.objects.filter(user=request.user)
    print(f"Найдено лекал в базе: {patterns.count()}")

    if not patterns.exists():
        print("ОШИБКА: Нет лекал для расчета")
        messages.error(request, "Добавьте хотя бы одно лекало для расчета")
        return redirect("forms_app:form15_view")

    try:
        # Подготавливаем данные лекал
        base_patterns = []
        base_names = []

        for pattern in patterns:
            base_patterns.append((pattern.width, pattern.height))
            base_names.append(pattern.name)
            print(f"Лекало: {pattern.name} - {pattern.width}x{pattern.height} мм")

        # Формируем полный список лекал
        all_patterns = base_patterns * num_sets
        all_names = base_names * num_sets
        print(f"Всего лекал для раскроя: {len(all_patterns)}")

        # Проверяем, чтобы ни одно лекало не было шире ткани
        for i, (w, h) in enumerate(all_patterns):
            if w > fabric_width:
                error_msg = f"Лекало '{all_names[i]}' ({w} мм) шире полотна ({fabric_width} мм)!"
                print(f"ОШИБКА: {error_msg}")
                messages.error(request, error_msg)
                return redirect("forms_app:form15_view")

        # Упаковка лекал
        print("Начинаем упаковку лекал...")
        packer = newPacker(
            mode=PackingMode.Offline, bin_algo=PackingBin.BBF, rotation=False
        )

        # Добавляем "рулон" ткани
        packer.add_bin(fabric_width, float("inf"))

        # Добавляем все лекала
        for i, (w, h) in enumerate(all_patterns):
            packer.add_rect(w, h, rid=i)

        # Выполняем упаковку
        packer.pack()

        # Определяем минимальную длину
        min_length = 0
        for abin in packer:
            for rect in abin:
                min_length = max(min_length, rect.y + rect.height)

        print(f"Упаковка завершена. Минимальная длина: {min_length} мм")
        print(
            f"Количество упакованных элементов: {sum(1 for abin in packer for _ in abin)}"
        )

        # Генерация файла в зависимости от формата
        if output_format == "pdf":
            print("Генерация PDF файла...")
            response = generate_pdf_response(
                packer, all_names, fabric_width, min_length, num_sets
            )
        else:
            print("Генерация Excel файла...")
            response = generate_excel_response(
                packer, all_names, fabric_width, min_length, num_sets
            )

        print("=== form15_calculate: Файл успешно сгенерирован ===")
        return response

    except Exception as e:
        print(f"=== КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        import traceback

        traceback.print_exc()  # Печатаем полный трейсбэк
        messages.error(request, f"Ошибка при расчете: {str(e)}")
        return redirect("forms_app:form15_view")


@login_required
def form15_clear_all(request):
    """Очистка всех лекал пользователя"""
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
    """Импорт лекал из Excel"""
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
