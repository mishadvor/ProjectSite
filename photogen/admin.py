from django.contrib import admin
from .models import ImageAsset, GenerationTask, PromptPreset


@admin.register(ImageAsset)
class ImageAssetAdmin(admin.ModelAdmin):
    list_display = ("id", "category", "uploaded_by", "created_at")
    list_filter = ("category", "created_at")
    search_fields = ("image",)
    date_hierarchy = "created_at"


@admin.register(GenerationTask)
class GenerationTaskAdmin(admin.ModelAdmin):
    list_display = ("id", "user", "mode", "model_name", "status", "cost_rub", "created_at", "finished_at")
    list_filter = ("mode", "status", "created_at")
    search_fields = ("user__username", "error_message")
    readonly_fields = ("created_at", "finished_at")
    date_hierarchy = "created_at"


@admin.register(PromptPreset)
class PromptPresetAdmin(admin.ModelAdmin):
    list_display = ("kind", "code", "title", "order", "is_active", "updated_at")
    list_filter = ("kind", "is_active")
    search_fields = ("code", "title", "template")
    list_editable = ("order", "is_active")
    ordering = ("kind", "order")