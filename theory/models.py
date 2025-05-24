# theory/models.py

from django.db import models
from ckeditor.fields import RichTextField  # ← добавь это
from ckeditor_uploader.fields import (
    RichTextUploadingField,
)  # ← и это, если используются загрузки


class StatisticsArticle(models.Model):
    title = models.CharField("Заголовок", max_length=200)
    slug = models.SlugField("URL-адрес", unique=True)
    content = RichTextUploadingField("Содержание")  # ← работает после импорта
    video_url = models.URLField("Видео (ВК / Rutube)", blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.title

    class Meta:
        verbose_name = "Статья: Чтение статистик"
        verbose_name_plural = "Статьи: Чтение статистик"


class GlossaryTerm(models.Model):
    term = models.CharField("Термин", max_length=150)
    definition = RichTextField("Определение")  # ← теперь работает
    slug = models.SlugField("URL-адрес", unique=True)
    category = models.CharField("Категория", max_length=100, blank=True, null=True)

    def __str__(self):
        return self.term

    class Meta:
        verbose_name = "Термин глоссария"
        verbose_name_plural = "Термины глоссария"
