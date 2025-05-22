from django.urls import path
from . import views

app_name = "home"  # Добавьте это

urlpatterns = [
    path("", views.home, name="home"),
]
