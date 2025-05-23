# theory/views.py
from django.shortcuts import render


def theory_index(request):
    return render(request, "theory/index.html")


def reading_stats(request):
    return render(request, "theory/reading_stats.html")


def glossary(request):
    return render(request, "theory/glossary.html")
