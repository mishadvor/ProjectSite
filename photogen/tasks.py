"""
Логика генерации изображений через AITUNNEL.
Перенесено из рабочих скриптов пользователя.
"""
import base64
import random
from datetime import datetime

from django.core.files.base import ContentFile
from django.utils import timezone

from config import client
from .models import ImageAsset, GenerationTask, GenerationResult


# ============================================================
# Промпты (жёстко зашиты — как в ваших скриптах)
# ============================================================

PROMPT_OUTFIT = (
    "Image 1 is the model. Keep her face, hairstyle, and body shape unchanged. "
    "Dress her in the skirt from Image 2, the top from Image 3, "
    "the bag from Image 4, and the shoes from Image 5. "
    "Keep the exact design, color, and material of each item from its source image. "
    "Use the background and lighting from Image 1."
)

POSES = {
    "sitting": (
        "Поза: Девушка сидит на скамейке, корпус слегка развёрнут в три четверти к камере. "
        "Спина прямая, плечи расслаблены, лёгкий наклон корпуса вперёд. "
        "Правая рука опущена вниз и опирается ладонью на скамейку справа от бедра, пальцы расслаблены. "
        "Левая рука лежит на правом бедре ближе к колену, пальцы расслаблены, кисть слегка развёрнута. "
        "Ноги: правая нога закинута на левую — колено направлено в сторону, голень опущена, стопа расслаблена. "
        "Из-за позы нога на ногу разрез оливковой кожаной юбки обязательно слегка приоткрыт спереди, "
        "но фасон, длина и посадка юбки не меняются. "
        "Голова слегка наклонена к правому плечу, взгляд прямо в камеру, мягкая тёплая улыбка. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
    "profile": (
        "Поза: Девушка стоит в профиль, опершись рукой на колонну. "
        "Корпус развёрнут боком к камере, вес тела на одну ногу, вторая чуть согнута. "
        "Одна рука лежит на колонне, вторая свободно опущена вдоль тела. "
        "Голова повёрнута в сторону камеры, взгляд в объектив. "
        "Юбка видна сбоку, силуэт и посадка читаются чётко. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
    "walking": (
        "Поза: Девушка идёт спокойным шагом вперёд, корпус чуть наклонён вперёд, "
        "одна нога впереди, вторая сзади. Юбка в лёгком движении, разрез слегка открыт. "
        "Руки свободно опущены вдоль тела, одна рука чуть отведена назад. "
        "Взгляд направлен вперёд, волосы в лёгком движении. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
        "back": (
        "Поза: Девушка стоит спиной к камере, слегка повернув голову через плечо. "
        "Корпус прямой, вес на обе ноги, ноги на ширине плеч. "
        "Руки свободно опущены вдоль тела. "
        "Взгляд через плечо в камеру, мягкая улыбка. "
        "Юбка видна сзади: посадка на бёдрах, длина, БЕЗ разреза сзади. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
    "standing": (
        "Поза: Девушка стоит лицом к камере, ноги чуть расставлены, вес равномерно на обеих. "
        "Одна рука свободно опущена, вторая чуть согнута в локте и лежит на бедре. "
        "Корпус прямой, плечи расслаблены, лёгкий поворот головы в три четверти. "
        "Взгляд в камеру, спокойная уверенная улыбка. "
        "Юбка с разрезом, виден разрез, силуэт и посадка читаются чётко. "
        "Кадр в полный рост: от макушки до каблуков, юбка и обувь видны целиком, силуэт читается чётко."
    ),
    "leaning_wall": (
        "Поза: Девушка стоит, прислонившись спиной к стене, одна нога согнута в колене "
        "и упирается стопой в стену, вторая стоит прямо. "
        "Руки скрещены на груди или одна опущена вдоль тела. "
        "Голова чуть наклонена, взгляд в сторону или в камеру. Лёгкая непринуждённая поза. "
        "Юбка, виден сбоку немного разрез, силуэт и посадка читаются чётко. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
    "hands_pockets": (
        "Поза: Девушка стоит прямо, обе руки слегка опущены в карманы юбки. "
        "Ноги слегка скрещены (одна перед другой), корпус чуть развёрнут в три четверти. "
        "Плечи расслаблены, лёгкая улыбка, взгляд прямо в камеру. "
        "Юбка, виден немного разрез, силуэт и посадка читаются чётко. "
        "Акцент на посадке и силуэте юбки. Кадр в полный рост: от макушки до каблуков."
    ),
    "holding_bag": (
        "Поза: Девушка стоит в три четверти к камере, одна рука держит сумку "
        "на плече или в согнутой руке у бедра. Вторая рука свободно опущена. "
        "Вес тела чуть смещён на одну ногу. Взгляд в камеру, мягкая улыбка. Сумка видна целиком. "
        "Юбка с разредом спереди, виден немного разрез, силуэт и посадка читаются чётко. "
        "Кадр в полный рост: от макушки до каблуков, юбка и обувь видны целиком."
    ),
    "twist": (
        "Поза: Девушка стоит лицом к камере, корпус развёрнут чуть в сторону, голова повёрнута к камере. "
        "Руки слегка отведены от тела, пальцы расслаблены. Вес на одной ноге, вторая чуть согнута. "
        "Поза в движении — момент поворота. Юбка в лёгком колыхании, но фасон и посадка сохранены. "
        "Юбка с разредом спереди, виден немного разрез, силуэт и посадка читаются чётко. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
    "sitting_chair": (
        "Поза: Девушка сидит на стуле или табурете, спина прямая, обе ноги опущены на пол, стопы вместе. "
        "Руки свободно лежат на коленях или одна рука на бедре, вторая на спинке стула. "
        "Корпус в три четверти к камере, голова повёрнута в камеру, мягкая улыбка. "
        "Юбка аккуратно ложится по ногам, разрез виден спереди. "
        "Кадр в полный рост: от макушки до каблуков, юбка и сапоги видны целиком."
    ),
}


PROMPT_POSE = """
Используй @Image1 как единственный источник внешности девушки, юбки и всего образа.
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


# photogen/tasks.py — заменить _save_first_result на _save_results

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

        response = client.images.edit(
            model="flux-3-image",
            image=files,
            prompt=PROMPT_OUTFIT,
            size=task.size,
            extra_body={"safety_tolerance": 4},
        )
        task.assets.set(used_assets)
        return response.data, 3.69
    finally:
        for f in files:
            f.close()


def generate_pose(task: GenerationTask, poses_selected):
    """Режим «Мастер-модель + локация» (muse-image). По одному запросу на позу."""
    master_file, master_asset = _pick_asset("master", task.selected_assets)
    location_file, location_asset = _pick_asset("location", task.selected_assets)

    results_with_pose = []  
    total_cost = 0
    try:
        for pose_key in poses_selected:
            pose_text = POSES.get(pose_key, "")
            if not pose_text:
                continue
            prompt = PROMPT_POSE.format(pose=pose_text)

            response = client.images.edit(
                model="muse-image",
                image=[master_file, location_file],
                prompt=prompt,
                size="1024x1024",
                extra_body={"sequential_image_generation": "disabled"},
            )
             # каждую картинку помечаем её позой
            for item in response.data:
                results_with_pose.append((pose_key, item))
            total_cost += 1.8
      
            # Сброс указателя для повторного чтения
            master_file.seek(0)
            location_file.seek(0)

        task.assets.set([master_asset, location_asset])
        return results_with_pose, total_cost
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
        else:
            results, cost = generate_pose(task, task.poses or [])

        _save_results(task, results)
        task.cost_rub = cost
        task.status = "done"
        task.finished_at = timezone.now()
        task.save()

    except Exception as e:
        task.status = "failed"
        task.error_message = str(e)
        task.finished_at = timezone.now()
        task.save()