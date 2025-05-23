# theory/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path("", views.theory_index, name="theory_index"),
    path("reading-stats/", views.reading_stats, name="reading_stats"),
    path("glossary/", views.glossary, name="glossary"),
]
