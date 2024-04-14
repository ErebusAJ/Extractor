from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_document, name='home'),
    path('about', views.about, name='about'),
    path('excel/', views.generate_excel, name='excel'),
]