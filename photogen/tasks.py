"""
Логика генерации изображений через AITUNNEL.
Перенесено из рабочих скриптов пользователя.
"""
import base64
import random
from datetime import datetime

from django.core.cache import cache
from django.core.files.base import ContentFile
from django.utils import timezone

from config import client
from .models import ImageAsset, GenerationTask, GenerationResult, PromptPreset


CACHE_TTL = 300  # 5 минут

PROMPT_OUTFIT = (
    "Image 1 is the model. Keep her face, hairstyle, and body shape unchanged. "
    "Dress her in the skirt from Image 2, the top from Image 3, "
    "the bag from Image 4, and the shoes from Image 5. "
    "Keep the exact design, color, and material of each item from its source image. "
    "Use the background and lighting from Image 1."
)

PROMPT_POSE = """Используй @Image1 как единственный источник внешности девушки, юбки и всего образа.
Используй @Image2 как источник СТИЛИЗАЦИИ локации: материалы, цвета, свет, атмосфера.
НЕ копируй @Image2 целиком. Возьми из файла общую атмосферу, вид помещения, только визуальный язык:
мягкий тёплый свет, современный интерьер.

СОХРАНИ АБСОЛЮТНО БЕЗ ИЗМЕНЕНИЙ с @Image1:
- Лицо, причёску, макияж, украшения.
- Одежду: юбку, верх, обувь, сумку — всё как на @Image1.
- НЕ МЕНЯЙ фасон, цвет, ткань и посадку одежды.

Помести девушку из @Image1 в СТИЛИЗАЦИЮ @Image2.
Каждый кадр — другой фрагмент локации.

Создай фотографию этой девушки в ПОЛНЫЙ РОСТ:
{pose}

Лицо и одежда — как на @Image1. Стиль локации — как на @Image2.
"""

def _get_prompt(kind: str, code: str | None, default: str = "") -> str:
    """Возвращает текст промта из БД с кэшем. Если нет — default."""
    key = f"photogen_prompt_{kind}_{code or ''}"
    text = cache.get(key)
    if text is None:
        qs = PromptPreset.objects.filter(kind=kind, is_active=True)
        if code is not None:
            qs = qs.filter(code=code)
        obj = qs.first()
        text = obj.template if obj else default
        cache.set(key, text, CACHE_TTL)
    return text

# ============================================================
# Вспомогательные функции
# ============================================================

def _pick_asset(category: str, selected: dict = None):
    """Возвращает случайный файл либо конкретный (по id из selected).
    selected — словарь вида {"model": ["3","7"], "skirt": [], ...}.
    """
    ids = (selected or {}).get(category) or []
    if ids:
        assets = list(ImageAsset.objects.filter(id__in=ids))
    else:
        assets = list(ImageAsset.objects.filter(category=category))
    if not assets:
        raise FileNotFoundError(
            f"В библиотеке нет изображений для категории «{category}». "
            f"Загрузите хотя бы одно фото на странице генератора."
        )
    asset = random.choice(assets)
    return open(asset.image.path, "rb"), asset



def _save_one_result(task, order_idx: int, pose_key: str, item):
    """Сохраняет один GenerationResult сразу — чтобы UI видел прогресс."""
    if getattr(item, "b64_json", None):
        image_data = base64.b64decode(item.b64_json)
    elif getattr(item, "url", None):
        import requests
        resp = requests.get(item.url, timeout=60)
        resp.raise_for_status()
        image_data = resp.content
    else:
        return  # битый результат

    filename = f"result_{task.id}_{order_idx+1}_{datetime.now():%Y%m%d_%H%M%S}.png"
    result = GenerationResult(task=task, order=order_idx, pose=pose_key)
    result.image.save(filename, ContentFile(image_data), save=True)

def _save_results(task, results):
    """Сохраняет ВСЕ результаты в GenerationResult."""
    if not results:
        raise ValueError("Модель вернула пустой результат")

    for idx, entry in enumerate(results):
        # Разбираем элемент: либо кортеж (pose, item), либо просто item
        if isinstance(entry, tuple) and len(entry) == 2:
            pose_key, item = entry
        else:
            pose_key, item = "", entry

        # Дальше работаем с item (это объект ответа от OpenAI)
        if getattr(item, "b64_json", None):
            image_data = base64.b64decode(item.b64_json)
        elif getattr(item, "url", None):
            import requests
            resp = requests.get(item.url, timeout=60)
            resp.raise_for_status()
            image_data = resp.content
        else:
            continue  # пропускаем битый результат

        filename = f"result_{task.id}_{idx+1}_{datetime.now():%Y%m%d_%H%M%S}.png"
        result = GenerationResult(task=task, order=idx, pose=pose_key)
        result.image.save(filename, ContentFile(image_data), save=True)


# ============================================================
# Генерация: два режима
# ============================================================

def generate_outfit(task: GenerationTask):
    """Режим «Сборка образа» (flux-3-image). 5 референсов."""
    files = []
    used_assets = []
    try:
        for cat in ["model", "skirt", "top", "bag", "shoes"]:
            f, asset = _pick_asset(cat, task.selected_assets)
            files.append(f)
            used_assets.append(asset)

        prompt = _get_prompt("outfit", "outfit", default=PROMPT_OUTFIT)
        response = client.images.edit(
            model="flux-3-image",
            image=files,
            prompt=prompt,
            size=task.size,
            extra_body={"safety_tolerance": 4},
        )
        task.assets.set(used_assets)
        return response.data, 3.69
    finally:
        for f in files:
            f.close()


def generate_pose(task: GenerationTask, poses_selected):
    """Режим «Мастер-модель + локация» (muse-image). По одному запросу на позу.
    Сохраняет каждый результат сразу после получения — чтобы UI видел прогресс.
    """
    master_file, master_asset = _pick_asset("master", task.selected_assets)
    location_file, location_asset = _pick_asset("location", task.selected_assets)

    total_cost = 0
    order_idx = 0
    try:
        wrapper = _get_prompt("pose_wrapper", "main", default=PROMPT_POSE)
        for pose_key in poses_selected:
            pose_text = _get_prompt("pose", pose_key)
            if not pose_text:
                continue
            prompt = wrapper.format(pose=pose_text)

            response = client.images.edit(
                model="muse-image",
                image=[master_file, location_file],
                prompt=prompt,
                size="1024x1024",
                extra_body={"sequential_image_generation": "disabled"},
            )

            # Сохраняем результаты этой позы сразу
            for item in response.data:
                _save_one_result(task, order_idx, pose_key, item)
                order_idx += 1
            total_cost += 1.8

            master_file.seek(0)
            location_file.seek(0)

        task.assets.set([master_asset, location_asset])
        return total_cost
    finally:
        master_file.close()
        location_file.close()

# ============================================================
# Главная точка входа для воркера
# ============================================================

def run_task(task_id: int):
    """Выполняет задачу генерации и обновляет её статус."""
    task = GenerationTask.objects.get(pk=task_id)
    task.status = "running"
    task.save(update_fields=["status"])

    try:
        if task.mode == "outfit":
            results, cost = generate_outfit(task)
            _save_results(task, results)
        else:
            cost = generate_pose(task, task.poses or [])

        task.cost_rub = cost
        task.status = "done"
        task.finished_at = timezone.now()
        task.save()

    except Exception as e:
        task.status = "failed"
        task.error_message = str(e)
        task.finished_at = timezone.now()
        task.save()