import threading

from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import render, redirect, get_object_or_404

from .models import GenerationTask, ImageAsset, GenerationResult, PromptPreset
from .forms import PromptPresetForm
from .tasks import run_task


@login_required
def generator(request):
    """Главная страница: выбор режима, поз, кнопка «Сгенерировать»."""
    if request.method == "POST":
        mode = request.POST.get("mode")
        size = request.POST.get("size", "2K")

        selected = {
            "model":  request.POST.getlist("model_ids"),
            "skirt":  request.POST.getlist("skirt_ids"),
            "top":    request.POST.getlist("top_ids"),
            "bag":    request.POST.getlist("bag_ids"),
            "shoes":  request.POST.getlist("shoes_ids"),
            "master": request.POST.getlist("master_ids"),
            "location": request.POST.getlist("location_ids"),
        }
                
        if mode == "outfit":
            task = GenerationTask.objects.create(
                user=request.user,
                mode="outfit",
                model_name="flux-3-image",
                size=size,
                selected_assets=selected,
            )
        elif mode == "pose":
            poses = request.POST.getlist("poses")
            if not poses:
                messages.error(request, "Выберите хотя бы одну позу")
                return redirect("photogen:generator")
            task = GenerationTask.objects.create(
                user=request.user,
                mode="pose",
                model_name="muse-image",
                poses=poses,
                size=size,
                selected_assets=selected,
            )
        else:
            messages.error(request, "Неверный режим генерации")
            return redirect("photogen:generator")

        # Запускаем в фоне, чтобы не держать HTTP-запрос
        threading.Thread(target=run_task, args=(task.id,), daemon=True).start()

        return redirect("photogen:result", task_id=task.id)

    # Счётчики для интерфейса
    counts = []
    for cat, label in ImageAsset.CATEGORY_CHOICES:
        n = ImageAsset.objects.filter(category=cat, uploaded_by=request.user).count()
        counts.append((label, n))

    # Референсы для выбора (Flux)
    refs_outfit = []
    for cat, label in [("model", "Модель"), ("skirt", "Юбка"), ("top", "Топ"),
                       ("bag", "Сумка"), ("shoes", "Обувь")]:
        assets = ImageAsset.objects.filter(category=cat, uploaded_by=request.user)
        refs_outfit.append((cat, label, assets))

    # Референсы для выбора (Muse)
    refs_pose = []
    for cat, label in [("master", "Мастер модель"), ("location", "Локация")]:
        assets = ImageAsset.objects.filter(category=cat, uploaded_by=request.user)
        refs_pose.append((cat, label, assets))

    poses = PromptPreset.objects.filter(
        kind="pose", is_active=True
    ).order_by("order", "title")

    return render(request, "photogen/generator.html", {
        "counts": counts,
        "refs_outfit": refs_outfit,
        "refs_pose": refs_pose,
        "poses": poses,
    })


@login_required
def result(request, task_id):
    """Страница результата генерации (с polling статуса)."""
    task = get_object_or_404(GenerationTask, pk=task_id, user=request.user)
    return render(request, "photogen/result.html", {"task": task})


@login_required
def task_status(request, task_id):
    """JSON-статус задачи для polling из JS."""
    task = get_object_or_404(GenerationTask, pk=task_id, user=request.user)
    results = [
        {"url": r.image.url, "pose": r.pose}
        for r in task.results.all()
    ]

    # Для pose-режима — прогресс по позам
    selected_poses = task.poses or []
    done_poses = [r.pose for r in task.results.all()]
    # Текущая поза — первая из выбранных, которая ещё не готова
    current_pose = None
    if task.mode == "pose" and task.status in ("pending", "running"):
        for code in selected_poses:
            if code not in done_poses:
                current_pose = code
                break

    # Названия поз (для красивого отображения на клиенте)
    pose_titles = {
        p.code: p.title
        for p in PromptPreset.objects.filter(kind="pose")
    }

    return JsonResponse({
        "status": task.status,
        "error": task.error_message,
        "results": results,
        "cost_rub": str(task.cost_rub),
        "mode": task.mode,
        "total": len(selected_poses),
        "done": len(done_poses),
        "current_pose": current_pose,
        "current_pose_title": pose_titles.get(current_pose, "") if current_pose else "",
        "poses": [
            {"code": c, "title": pose_titles.get(c, c),
             "done": c in done_poses}
            for c in selected_poses
        ],
    })


@login_required
def upload_asset(request):
    """Загрузка одного или нескольких файлов в библиотеку."""
    if request.method == "POST":
        category = request.POST.get("category")
        images = request.FILES.getlist("images")

        if not category or not images:
            messages.error(request, "Выберите категорию и файлы")
            return redirect("photogen:library")

        for img in images:
            ImageAsset.objects.create(
                category=category,
                image=img,
                uploaded_by=request.user,
            )
        messages.success(request, f"Загружено файлов: {len(images)}")
    return redirect("photogen:library")


@login_required
def library(request):
    """Страница библиотеки — все загруженные референсы с фильтрами."""
    category = request.GET.get("category")
    qs = ImageAsset.objects.filter(uploaded_by=request.user)
    if category:
        qs = qs.filter(category=category)

    counts = {
        cat: ImageAsset.objects.filter(category=cat, uploaded_by=request.user).count()
        for cat, _ in ImageAsset.CATEGORY_CHOICES
    }
    return render(request, "photogen/library.html", {
        "assets": qs,
        "active_category": category,
        "categories": ImageAsset.CATEGORY_CHOICES,
        "counts": counts,
    })


@login_required
def delete_asset(request, asset_id):
    """Удаление одного референса."""
    asset = get_object_or_404(ImageAsset, pk=asset_id, uploaded_by=request.user)
    if request.method == "POST":
        asset.delete()
        messages.success(request, "Референс удалён")
    return redirect("photogen:library")

@login_required
def history(request):
    tasks = GenerationTask.objects.filter(user=request.user).order_by("-created_at")[:100]
    return render(request, "photogen/history.html", {"tasks": tasks})

@login_required
def task_detail(request, task_id):
    """Детали одной задачи: параметры, использованные референсы, все результаты."""
    task = get_object_or_404(GenerationTask, pk=task_id, user=request.user)
    return render(request, "photogen/task_detail.html", {"task": task})

@login_required
def delete_task(request, task_id):
    """Удаление задачи вместе с её результатами и физическими файлами."""
    task = get_object_or_404(GenerationTask, pk=task_id, user=request.user)
    if request.method == "POST":
        for r in task.results.all():
            if r.image:
                r.image.delete(save=False)
        task.delete()
        messages.success(request, f"Задача #{task_id} удалена")
    return redirect("photogen:history")


# ============================================================
# Управление промтами (только для staff)
# ============================================================

@staff_member_required
def prompt_list(request):
    kind = request.GET.get("kind")
    qs = PromptPreset.objects.all()
    if kind:
        qs = qs.filter(kind=kind)
    return render(request, "photogen/prompt_list.html", {
        "prompts": qs,
        "kinds": PromptPreset.KIND_CHOICES,
        "active_kind": kind,
    })


@staff_member_required
def prompt_edit(request, pk=None):
    obj = get_object_or_404(PromptPreset, pk=pk) if pk else None
    if request.method == "POST":
        form = PromptPresetForm(request.POST, instance=obj)
        if form.is_valid():
            saved = form.save()
            messages.success(request, f"Промт «{saved.title}» сохранён")
            return redirect("photogen:prompt_list")
    else:
        form = PromptPresetForm(instance=obj)
    return render(request, "photogen/prompt_edit.html", {"form": form, "obj": obj})


@staff_member_required
def prompt_delete(request, pk):
    obj = get_object_or_404(PromptPreset, pk=pk)
    if request.method == "POST":
        obj.delete()
        messages.success(request, "Промт удалён")
    return redirect("photogen:prompt_list")


