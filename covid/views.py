from django.shortcuts import redirect, render
def home(request):
    import xlrd
    workbook = xlrd.open_workbook("D:\\indhu college\\miniproject-2\\covid\\tamilnadu.xls")
    sheet = workbook.sheet_by_index(0)
    col=sheet.cell_value(0, 0)
    rows=sheet.nrows
    print(sheet.cell_value(22, 6))
    print(rows)
    cured=sheet.cell_value(rows-1,3)
    death=sheet.cell_value(rows-1,4)
    total_cases=sheet.cell_value(rows-1,5)
    cured=int(cured)
    death=int(death)
    total_cases=int(total_cases)
    print(cured)
    # Extracting number of columns
    print(sheet.ncols)
     #vaccine
    workbook1 = xlrd.open_workbook("D:\\indhu college\\miniproject-2\\covid\\vaccine.xls")
    sheet1=workbook1.sheet_by_index(0)
    n=sheet1.nrows
    vaccinated=sheet1.cell_value(n-1,20)
    vaccinated=int(vaccinated)
    return render(request,'home.html',{'cured':cured,'death':death,'total_cases':total_cases,'vaccinated':vaccinated}) 

def charts(request):
    import xlrd
    workbook = xlrd.open_workbook("D:\\indhu college\\miniproject-2\\covid\\tamilnadu.xls")
    workbook1 = xlrd.open_workbook("D:\\indhu college\\miniproject-2\\covid\\vaccine.xls")
    sheet = workbook.sheet_by_index(0)
    sheet1=workbook1.sheet_by_index(0)
    val=[]
    months=[]
    m=[]
    j=0
    #vaccine
    n=sheet1.nrows
    vaccinated=sheet1.cell_value(n-1,17)
    for i in range(0,sheet.nrows):
        if j!=0:
            val.append(int(sheet.cell_value(i,6)))
            xl_date=sheet.cell_value(i,0)
            datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
            date_object = datetime_date.date()
            string_date = date_object.isoformat()
            #months.append(string_date[8:])
            m.append(string_date)
        j=j+1  
    #print(months)
    val=list(map(str,val))
    val=' '.join(val)
    months=list(map(str,months))
    months=' '.join(months)
    return render(request,"charts.html",{'x':val,'y':m})


def covidend(request):  
    import xlrd 
    workbook1 = xlrd.open_workbook("D:\\indhu college\\miniproject-2\\covid\\vaccine.xls")
    sheet1=workbook1.sheet_by_index(0)
    n=sheet1.nrows
    noofpeopletoachieveheardimmunity=940000000*0.7
    vaccinated=sheet1.cell_value(n-1,19)
    totalvaccinated=int(vaccinated)
    vaccinatedpeople=int(totalvaccinated/2)
    noofpeoplevaccinatedinpercent=(vaccinatedpeople*100)/920000000
    balancepeople=(940000000*0.7)-vaccinatedpeople
    herdimmuntiyachievedsofar=(noofpeoplevaccinatedinpercent*90)/100
    
    return render(request,"covidend.html",{'vaccinatedpeople':vaccinatedpeople,'noofpeoplevaccinatedinpercent':noofpeoplevaccinatedinpercent,'balancepeople':balancepeople,'herdimmuntiyachievedsofar':herdimmuntiyachievedsofar,'herdimmunityneeded':63})
