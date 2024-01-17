from django.shortcuts import render, redirect
from django.http import HttpResponse

from .forms import UploadFileForm
from .utils import process_excel


def upload_file(request):
    """
    Обработчик загрузки файла.

    param:
        request (HttpRequest): Запрос от клиента.
    return:
        HttpResponse: Возвращает страницу загрузки
        файла или перенаправляет на
        страницу скачивания обработанного файла.
    """
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            input_file = request.FILES['file']
            output_file_path = 'processed_file.xlsx'
            process_excel(input_file, output_file_path)
            return redirect('download_file', file_name='processed_file.xlsx')
    else:
        form = UploadFileForm()
    return render(request, 'files/upload_file.html', {'form': form})


def download_file(request, file_name):
    """
    Обработчик скачивания файла.

    param:
        request (HttpRequest): Запрос от клиента.
        file_name (str): Имя файла для скачивания.

    return:
        HttpResponse: Возвращает файл для скачивания.
    """
    file_path = 'processed_file.xlsx'
    with open(file_path, 'rb') as file:
        response = HttpResponse(
            file.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="{file_name}"'
        return response
