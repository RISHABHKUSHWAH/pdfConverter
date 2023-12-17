from django.urls import path
from . import views

urlpatterns = [
    path('pdf_to_excel',views.pdfToExcel,name='pdfToExcel'),
    path('pdf_to_word',views.pdfToWord,name='pdfToWord'),
    path('pdf_to_jpg',views.pdfToJpg,name='pdfToJpg'),
    path('pdf_to_powerpoint',views.pdfToPowerpoint,name='pdfToPowerpoint'),
]
