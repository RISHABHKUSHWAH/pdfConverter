from django.shortcuts import render 
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
import os
from PyPDF2 import PdfMerger
def home(request):
    return render(request,'index.html')
    
def mergeFile(request):
    if request.method == "POST":
        pdf_files  = request.FILES.getlist('select_file')
        output_pdf = r'C:\Users\kushwah\OneDrive\Desktop\convert\pdfConverter\output.pdf'
        try:
            pdf_merger = PdfMerger()
            for pdf_file in pdf_files:
                pdf_merger.append(pdf_file)
            pdf_merger.write(output_pdf)
            pdf_merger.close()
        except Exception as e:
            print(f"Error: {str(e)}")
            return HttpResponse(f"Error: {str(e)}")
        return render(request,'merge.html')
    return render(request,'merge.html')