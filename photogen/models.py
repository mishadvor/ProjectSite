from django.db import models
from django.conf import settings
from io import BytesIO
from PIL import Image as PilImage
from django.core.files.base import ContentFile


class ImageAsset(models.Model):
    """Изображение в библиотеке референсов."""

    CATEGORY_CHOICES = [
        ("model", "Модель"),
        ("skirt", "Юбка"),
        ("top", "Топ"),
        ("bag", "Сумка"),
        ("shoes", "Обувь"),
        ("master", "Мастер модель"),
        ("location", "Локация"),
    ]
    MAX_SIZE = (1500, 1500)

    

    category = models.CharField("Категория", max_length=20, choices=CATEGORY_CHOICES)
    image = models.ImageField("Изображение", upload_to="assets/%Y/%m/")
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="photogen_assets",
        verbose_name="Загрузил",
    )
    created_at = models.DateTimeField("Загружено", auto_now_add=True)

    class Meta:
        verbose_name = "Референс"
        verbose_name_plural = "Референсы"
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.get_category_display()} — {self.image.name}"

    def save(self, *args, **kwargs):
        if self.pk is None and self.image:
            try:
                img = PilImage.open(self.image)
                if img.width > self.MAX_SIZE[0] or img.height > self.MAX_SIZE[1]:
                    img.thumbnail(self.MAX_SIZE, PilImage.LANCZOS)
                    output = BytesIO()
                    fmt = img.format if img.format in ("JPEG", "PNG", "WEBP") else "JPEG"
                    img.save(output, format=fmt, quality=88, optimize=True)
                    output.seek(0)
                    self.image.save(self.image.name, ContentFile(output.read()), save=False)
            except Exception:
                pass
        super().save(*args, **kwargs)    


class GenerationTask(models.Model):
    """Одна задача генерации."""

    MODE_CHOICES = [
        ("outfit", "Сборка образа (Flux 3)"),
        ("pose", "Мастер-модель + локация (Muse)"),
    ]

    STATUS_CHOICES = [
        ("pending", "Ожидает"),
        ("running", "Выполняется"),
        ("done", "Готово"),
        ("failed", "Ошибка"),
    ]

    POSE_CHOICES = [
        ("sitting", "Сидя на скамейке"),
        ("profile", "Профиль у колонны"),
        ("walking", "Быстрый шаг"),
        ("back", "Спиной через плечо"),
        ("standing", "Стоя в полный рост"),
        ("leaning_wall", "Прислонившись к стене"),
        ("hands_pockets", "Руки в карманах"),
        ("holding_bag", "С сумкой на плече"),
        ("twist", "Лёгкий разворот"),
        ("sitting_chair", "Сидя на стуле"),
    ]

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="photogen_tasks",
        verbose_name="Пользователь",
    )
    mode = models.CharField("Режим", max_length=10, choices=MODE_CHOICES)
    model_name = models.CharField("Модель AI", max_length=50)
    size = models.CharField("Разрешение", max_length=16, default="2K")
    poses = models.JSONField("Выбранные позы", default=list, blank=True)
    selected_assets = models.JSONField("Выбранные референсы", default=dict, blank=True)
    status = models.CharField(
        "Статус", max_length=10, choices=STATUS_CHOICES, default="pending"
    )
    
    error_message = models.TextField("Ошибка", blank=True)
    cost_rub = models.DecimalField(
        "Стоимость, ₽", max_digits=6, decimal_places=2, default=0
    )
    created_at = models.DateTimeField("Создано", auto_now_add=True)
    finished_at = models.DateTimeField("Завершено", null=True, blank=True)

    assets = models.ManyToManyField(
        ImageAsset, blank=True, related_name="tasks", verbose_name="Использованные референсы"
    )

    class Meta:
        verbose_name = "Задача генерации"
        verbose_name_plural = "Задачи генерации"
        ordering = ["-created_at"]

    def __str__(self):
        return f"#{self.pk} {self.get_mode_display()} — {self.get_status_display()}"

# photogen/models.py — добавить после GenerationTask

class GenerationResult(models.Model):
    """Одно сгенерированное изображение внутри задачи."""
    task = models.ForeignKey(
        GenerationTask,
        on_delete=models.CASCADE,
        related_name="results",
        verbose_name="Задача",
    )
    image = models.ImageField("Изображение", upload_to="results/%Y/%m/")
    pose = models.CharField("Поза", max_length=20, blank=True)
    created_at = models.DateTimeField("Создано", auto_now_add=True)
    order = models.PositiveSmallIntegerField("Порядок", default=0)

    class Meta:
        verbose_name = "Результат"
        verbose_name_plural = "Результаты"
        ordering = ["order", "id"]

    def __str__(self):
        return f"Результат #{self.pk} задачи #{self.task_id}"
        
    def pose_display(self):
        """Человекочитаемое название позы из GenerationTask.POSE_CHOICES."""
        if not self.pose:
            return "—"
        mapping = dict(GenerationTask.POSE_CHOICES)
        return mapping.get(self.pose, self.pose)
