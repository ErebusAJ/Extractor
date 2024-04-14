from django.shortcuts import render, redirect
from .forms import DocumentForm
from .utils import extract_info_and_generate_excel
from django.http import HttpResponse
from .models import Document


def upload_document(request):
    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return redirect('excel')
    else:
        form = DocumentForm()
    return render(request, 'main/home.html', {'form': form})

def generate_excel(request):
    document = Document.objects.last()
    if document:
        excel_file_path = extract_info_and_generate_excel(document.file.path)
        if excel_file_path:
            with open(excel_file_path, 'rb') as file:
                response = HttpResponse(file.read(), content_type='application/vnd.ms-excel')
                response['Content-Disposition'] = 'attachment; filename="CV_Info.xlsx"'
                return response
    return render(request, 'main/generate_excel.html')

def about(request):
    return render(request, 'main/about.html')