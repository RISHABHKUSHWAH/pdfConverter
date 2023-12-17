from django.urls import path
from . import views
from django.conf import settings 
from django.conf.urls.static import static
urlpatterns = [
    path('excel_to_pdf/',views.excelToPdf,name='excel_to_pdf'),
    path('word_to_pdf/',views.wordToPdf,name='word_to_pdf'),
    path('jpg_to_pdf/',views.jpgToPdf,name='jpg_to_pdf'),
    path('powerpoint_to_pdf/',views.powerpointToPdf,name='powerpoint_to_pdf'),
    path('html_to_pdf/',views.htmlToPdf,name='html_to_pdf'),
]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

