from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render
from .forms import UploadFileForm
from .file_manager import interpreter
# Create your views here.


def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        #print(request.FILES)
        if form.is_valid():
            interpreter(request.FILES['file'])
            return render(request, 'success.html')
    else:
        form = UploadFileForm()
    return render(request, 'input.html', {'form': form})


def index(request):
    return HttpResponse("py-kompas-control")
