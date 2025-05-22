# forms_app/urls.py

from django.urls import path

# Импортируем представления
from .views.form1_view import form1
from .views.form2_view import form2
from .views.form3_view import form3

app_name = "forms_app"

urlpatterns = [
    path("form1/", form1, name="form1"),  # ← Так правильно
    path("form2/", form2, name="form2"),  #   Вызываем функции напрямую
    path("form3/", form3, name="form3"),  #   без views.
]
