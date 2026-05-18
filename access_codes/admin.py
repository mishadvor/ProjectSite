from django.contrib import admin
from .models import MasterAccessCode


@admin.register(MasterAccessCode)
class MasterAccessCodeAdmin(admin.ModelAdmin):
    list_display = ["code", "is_active", "updated_at", "updated_by"]
    list_editable = ["is_active"]
    readonly_fields = ["updated_at"]
