from django.db import models

# Create your models here.
# forms_app/models.py

from django.contrib.auth.models import User
from django.db import models


class UserReport(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    output_file = models.FileField(upload_to="user_reports/")
    last_updated = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Отчет {self.user.username} — {self.last_updated}"
