from django.contrib import admin
from .models import UserReport
from .models import WeeklyReport


@admin.register(UserReport)
class UserReportAdmin(admin.ModelAdmin):
    list_display = ("user", "file_name", "report_type", "last_updated")
    list_filter = ("report_type",)
    search_fields = ("user__username", "file_name")


@admin.register(WeeklyReport)
class WeeklyReportAdmin(admin.ModelAdmin):
    list_display = ("week_name", "art_group", "profit", "created_at")
    search_fields = ("week_name", "art_group")
    list_filter = ("week_name", "art_group")
