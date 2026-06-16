from django.shortcuts import render
from django.contrib.auth.decorators import login_required


@login_required
def teste(request):
    return render(request, "core/teste.html")