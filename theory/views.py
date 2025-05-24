# theory/views.py

from django.shortcuts import render, get_object_or_404
from .models import StatisticsArticle, GlossaryTerm


def theory_index(request):
    return render(request, "theory/index.html")


def reading_stats(request):
    articles = StatisticsArticle.objects.all()
    return render(request, "theory/reading_stats.html", {"articles": articles})


def glossary(request):
    return render(request, "theory/glossary.html")


def article_list(request):
    articles = StatisticsArticle.objects.all()
    return render(request, "theory/article_list.html", {"articles": articles})


def article_detail(request, slug):
    article = get_object_or_404(StatisticsArticle, slug=slug)
    return render(request, "theory/article_detail.html", {"article": article})


def glossary(request):
    terms = GlossaryTerm.objects.all()  # ← эта строка важна
    return render(
        request, "theory/glossary.html", {"terms": terms}
    )  # передаём в шаблон
