# theory/admin.py

from django.contrib import admin
from .models import StatisticsArticle
from .models import StatisticsArticle, GlossaryTerm


@admin.register(StatisticsArticle)
class StatisticsArticleAdmin(admin.ModelAdmin):
    list_display = ("title", "slug")
    prepopulated_fields = {"slug": ("title",)}


@admin.register(GlossaryTerm)
class GlossaryTermAdmin(admin.ModelAdmin):
    list_display = ("term", "definition", "category")
    search_fields = ("term",)
    prepopulated_fields = {"slug": ("term",)}
    list_filter = ("category",)
