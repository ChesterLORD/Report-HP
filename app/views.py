from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.urls import reverse
from .views import *
from .forms import SearchForm ,ReportForm
from .models import HospitalMRD, ClinicalPharmacist,AddNote
from .export.excel import WriteToExcel
from .export.pdf import WriteToPdf






def index(request):
    data = ClinicalPharmacist.objects.filter(MRD=HospitalMRD.objects.first())
    data_note = AddNote.objects.all()


    if request.method == 'POST':
        form = SearchForm(request.POST)
        if form.is_valid():
            data = form.search()
    else:
        form = SearchForm()
    if 'excel' in request.POST:
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=Report.xlsx'
        xlsx_data = WriteToExcel(data)
        response.write(xlsx_data)
        return response
    if 'pdf' in request.POST:
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachement; filename=Report.pdf'
        pdf_data = WriteToPdf(data)
        response.write(pdf_data)
        return response
    context = {"form": form, "data": data , "data_note":data_note}
    return render(request, 'app/index.html', context)





def report(request):
    flag = 0
    #data = ClinicalPharmacist.objects.filter(station=WeatherStation.objects.first())
    form = ReportForm(request.POST or None)
    if request.method == 'POST':
        if form.is_valid():

            MRD = form.cleaned_data.get('MRD')
            date = form.cleaned_data.get('date')
            roomNum = form.cleaned_data.get('roomNum')
            GFR = form.cleaned_data.get('GFR')
            renal_impairment = form.cleaned_data.get('renal_impairment')
            Liver_Imparrenenty = form.cleaned_data.get('Liver_Imparrenenty')
            Dose_adjustment = form.cleaned_data.get('Dose_adjustment')
            Dose_adjustment_LI = form.cleaned_data.get('Dose_adjustment_LI')
            Urine_output = form.cleaned_data.get('Urine_output')
            Balance = form.cleaned_data.get('Balance')
            intervention = form.cleaned_data.get('intervention')
            Child_pugh_score = form.cleaned_data.get('Child_pugh_score')
            Feeding = form.cleaned_data.get('Feeding')
            Bowel_motion = form.cleaned_data.get('Bowel_motion')
            Electrolytes_imbalance = form.cleaned_data.get('Electrolytes_imbalance')
            Hyper_Hypo = form.cleaned_data.get('Hyper_Hypo')
            ABI = form.cleaned_data.get('ABI')
            Metabolic_num = form.cleaned_data.get('Metabolic_num')
            Respiratory_num = form.cleaned_data.get('Respiratory_num')
            Metabolic = form.cleaned_data.get('Metabolic')
            Respiratory = form.cleaned_data.get('Respiratory')
            QT_C = form.cleaned_data.get('QT_C')
            QT_C_num = form.cleaned_data.get('QT_C_num')
            VITALS = form.cleaned_data.get('VITALS')
            Analgesic_management = form.cleaned_data.get('Analgesic_management')
            Sedation = form.cleaned_data.get('Sedation')
            Thromboembolic_Prophylaxis = form.cleaned_data.get('Thromboembolic_Prophylaxis')
            Stress_Ulcer_Pophylaxis = form.cleaned_data.get('Stress_Ulcer_Pophylaxis')
            Glycemic_control_target_BG = form.cleaned_data.get('Glycemic_control_target_BG')
            T_BG = form.cleaned_data.get('T_BG')
            Infection = form.cleaned_data.get('Infection')
            Treatment = form.cleaned_data.get('Treatment')
            AB_INTERVENTION = form.cleaned_data.get('AB_INTERVENTION')
            MP_LIST = form.cleaned_data.get('MP_LIST')



            report = ClinicalPharmacist(MRD=MRD, date = date , GFR = GFR, renal_impairment = renal_impairment,
                Liver_Imparrenenty = Liver_Imparrenenty, Dose_adjustment = Dose_adjustment,  Urine_output = Urine_output, 
                Balance = Balance, intervention= intervention ,Child_pugh_score = Child_pugh_score, Feeding = Feeding, 
                Bowel_motion=Bowel_motion ,Dose_adjustment_LI = Dose_adjustment_LI   ,
                Electrolytes_imbalance = Electrolytes_imbalance, Hyper_Hypo = Hyper_Hypo,
                ABI = ABI, Metabolic_num = Metabolic_num, 
                Respiratory_num=Respiratory_num ,Metabolic = Metabolic   ,
                Respiratory = Respiratory, QT_C = QT_C,
                QT_C_num = QT_C_num   ,
                VITALS = VITALS, Analgesic_management = Analgesic_management,
                Sedation = Sedation, Thromboembolic_Prophylaxis = Thromboembolic_Prophylaxis, 
                Stress_Ulcer_Pophylaxis=Stress_Ulcer_Pophylaxis ,
                Glycemic_control_target_BG = Glycemic_control_target_BG, T_BG = T_BG,
                Infection = Infection, Treatment = Treatment,AB_INTERVENTION = AB_INTERVENTION, MP_LIST = MP_LIST, status=False)
            report.save()

            flag = 1
    context = {
        'form': form, 
        'flag': flag,
    }
    html_template = loader.get_template('app/report.html')


    return HttpResponse(html_template.render(context, request))