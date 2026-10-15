from django.contrib import admin
from .models import ImageAsset, GenerationTask


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
