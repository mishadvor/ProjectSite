from django.shortcuts import render


def form1(request):
    return render(request, "forms_app/form1.html")


def form2(request):
    return render(request, "forms_app/form2.html")


def form3(request):
    return render(request, "forms_app/form3.html")
