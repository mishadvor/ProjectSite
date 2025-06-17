from django.contrib import admin
from .models import UserReport


@admin.register(UserReport)
class UserReportAdmin(admin.ModelAdmin):
    list_display = ("user", "file_name", "report_type", "last_updated")
    list_filter = ("report_type",)
    search_fields = ("user__username", "file_name")
