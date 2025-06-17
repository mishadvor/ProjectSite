from django.db import models
from django.contrib.auth.models import User


class UserReport(models.Model):  # ✅ Теперь models определён
    REPORT_TYPES = (("form4", "Накопительный отчёт"), ("form5", "Остатки"))

    user = models.ForeignKey(User, on_delete=models.CASCADE)
    file_name = models.CharField(max_length=255)
    file_path = models.CharField(max_length=500)
    report_type = models.CharField(max_length=10, choices=REPORT_TYPES)
    last_updated = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.file_name} — {self.user.username}"
