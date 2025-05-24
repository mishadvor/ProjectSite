# theory/urls.py
from django.urls import path
from . import views

app_name = "theory"  # ← Эта строка важна!

urlpatterns = [
    path("", views.theory_index, name="theory_index"),
    path("reading-stats/", views.reading_stats, name="reading_stats"),
    path("glossary/", views.glossary, name="glossary"),
    # Статьи для "Чтение статистик"
    path("articles/", views.article_list, name="article_list"),
    path("articles/<slug:slug>/", views.article_detail, name="article_detail"),
]
