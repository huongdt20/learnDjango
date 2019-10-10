from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from demo_media import  models
from django.conf import settings
import os
from xlrd import open_workbook
# Create your views here.

def upfile(request):
    if request.method == 'POST' and request.FILES['myfile']:
        name = request.POST.get('carname')
        price = request.POST.get('price')
        myfile = request.FILES['myfile']
        fileStorage = FileSystemStorage()
        fileDir = 'cars/'+ name + '/' + myfile.name # vd: a up ảnh tên acb.jpg, xe tên là honda -> /media/cars/honda/abc.jpg
        fileStorage.save(fileDir,myfile) # giống lưu file trong java hoặc C#

        car = models.Car(name= name, price= price, photo= fileDir)
        car.save()
    return render(request,'upload_without_delete.html')

def cars(request):
    car = models.Car.objects.all()
    content = {'car': car}
    return render(request,'cars.html', content)


def upload_excel(request):
    if request.method =='POST' and request.FILES['myexcel']:
        myfile = request.FILES['myexcel']
        fs = FileSystemStorage()
        file = fs.save('excel/'+myfile.name, myfile)
        upload_url = fs.url(file)
        url_file = settings.BASE_DIR.replace('\\', '//') + str(upload_url).replace('/', '//').replace('%20', ' ')

        ############## OPEN FILE ####################
        workbook = open_workbook(url_file)
        sheet_names = workbook.sheet_names()

        for name in sheet_names:
            print('-----------------------------------------------\n')
            print(name)
            work_sheet = workbook.sheet_by_name(name)
            nrows = work_sheet.nrows
            ncols = work_sheet.ncols

            # for col in range(work_sheet.ncols):
            for row in range(1,work_sheet.nrows): # cho row chay tu 1 den so cot( de bo hang dau tien trong cot)

                  a = models.upload( row1=work_sheet.cell_value(row,0),row2=work_sheet.cell_value(row,1),
                                      row3=work_sheet.cell_value(row,2),row4=work_sheet.cell_value(row,3))
                  a.save()


        ############### Xoa file de giam dung luong he thong #################
        try:
            os.remove(url_file)
        except os.error:
            pass
    return   render(request, 'upload_and_read_excel.html')