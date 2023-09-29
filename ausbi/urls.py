from django.urls import path
from .views import baixar_planilhas

urlpatterns = [
    path('baixar_planilhas/', baixar_planilhas, name='baixar_planilhas'),
]