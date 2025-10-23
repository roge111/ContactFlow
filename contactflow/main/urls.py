from django.contrib import admin
from django.urls import path

from . import views

urlpatterns = [
    path('', views.main, name='main'),
    path('import/', views.import_file, name='import_file'),  # API для импорта
    path('export/', views.export_file, name='export_file'),  # API для экспорта
    path('import-page/', views.import_page, name='import_page'),  # Страница импорта
    path('export-page/', views.export_page, name='export_page'),  # Страница экспорта
]