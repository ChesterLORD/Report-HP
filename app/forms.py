from django import forms

from .models import HospitalMRD, ClinicalPharmacist, AddNote


def available_stations():
    return HospitalMRD.objects.values_list("id", "MRD")


class SearchForm(forms.Form):
    MRD = forms.ChoiceField(
        label="MRD",
        choices=available_stations
    )
    date_from = forms.DateField(
        label="From",
        input_formats=['%Y-%m-%d',],
        widget=forms.TextInput(attrs={'size': 10}),
        required=False
    )
    date_to = forms.DateField(
        label="To",
        input_formats=['%Y-%m-%d',],
        widget=forms.TextInput(attrs={'size': 10}),
        required=False
    )

    def search(self):
        MRD = self.cleaned_data.get('MRD', None)
        date_from = self.cleaned_data.get('date_from', None)
        date_to = self.cleaned_data.get('date_to', None)

        data = ClinicalPharmacist.objects.filter()


        if MRD:
            data = data.filter(MRD=MRD)
        if date_from:
            data = data.filter(date__gte=date_from)
        if date_to:
            data = data.filter(date__lte=date_to)

        return data


class ReportForm(forms.Form):

    URINE_OUTPUT_CHOICES = [
    ('N', 'Normal'),
    ('O', 'Oliguria'),
    ('P', 'Polyuria'),
    ]

    BALANCE_CHOICES = [
    ('-V', '+Ve'),
    ('+V', '-Ve'),
    
    ]

    FEEDING_CHOICES = [
    ('NPO', 'NPO'),
    ('Oral', 'Oral'),
    ('NGT/OGT', 'NGT/OGT'),
    ('PEG', 'PEG'),
    ('TPN', 'TPN'),
    ]

   
    ELECTROLYTES_CHOICES = [
    ('Na', 'Na'),
    ('K', 'K'),
    ('Mg', 'Mg'),
    ('Po4', 'Po4'),
    ('Ca', 'Ca'),
    ]

    ABI_CHOICES = [
    ('Yes', 'Yes'),
    ('Normal', 'Normal'),
    ]



    HYPER_HYPO_CHOICES = [
    ('hyper', 'hyper'),
    ('hypo', 'hypo'),
    
    ]
    METABOLIC_CHOICES = [
    ('Acidosis', 'Acidosis'),
    ('Alkalosis', 'Alkalosis'),
    ]  

    RESPIRATORY_CHOICES = [
    ('Acidosis', 'Acidosis'),
    ('Alkalosis', 'Alkalosis'),
    ]            
          
    QT_C_CHOICES = [
    ('>500', '>500'),
    ('<500', '<500'),
    ]

    VITALS_CHOICES = [
    ('stable', 'stable'),
    ('tech carbon', 'techcarbon'),
    ('tachypnea'  , 'tachypnea'),
    ('htn', 'HTN'),
    ('bradycardia', 'bradycardia'),
    ('bradypnea', 'bradypnea'),
    ('hypotension', 'hypotension'),
    ('Fever', 'Fever'),
    ] 

    TP_CHOICES = [
    ('Normal', 'Normal'),
    ('Mechanical', 'Mechanical'),
    ('Pharmacology', 'Pharmacology'),

    ]

    SUP_CHOICES =[

            ('Normal','Normal'),
            ('Yes',(
                ('Pantoprazole-40-mg-IV-OD','Pantoprazole-40-mg-IV-OD'),
                ),
            ),

    ]

    
    T_BG_CHOICES = [
    
        ('G1',(
                ('Hypoglycomic', 'Hypoglycomic'),
                ('Glucagon', 'Glucagon'),
                ('D50%', 'D50%'),
                ('D5%/D10%', 'D5%/D10%')
                
                )
            ),
        ('G2',(
                ('koton', 'koton'),
                ('Hydration', 'Hydration'),
                )
            ),
        ('G3',(
                ('Hyperglycemia', 'Hyperglycemia'),
                ('Insulin Infusion', 'Insulin Infusion'),
                ('DKA protocol', 'DKA protocol'),
                ('SQ Sliding Scale', 'SQ Sliding Scale'),
                ('Long acting insulin', 'Long acting insulin'),

                
                
                
                )
            ),
    

    ]
    

    INFECTION_CHOICES = [
    ('UA', 'UA'),
    ('Markers', 'Markers'),
    ('Fever', 'Fever'),
    ('Culture', 'Culture'),
    ('X-Ray', 'X-Ray'),
    ('Swab', 'Swab'),


    ]   

    UA_CHOICES = [
    ('+Ve', '+Ve'),
    ('-Ve', '-Ve'),
    ] 



  

    TREATMENT_CHOICES = [
    ('Antifungal', 'Antifungal'),
    ('Antiviral', 'Antiviral'),
    ('Antimonial', 'Antimonial'),
    ('AB-A',(
                ('Ampicillin', 'Ampicillin'),
                ('Amoxiclav', 'Amoxiclav'),
                ('Amikacin', 'Amikacin'),
                ('Azithromycin','Azithromycin'),
                ('Anidulafungin','Anidulafungin'),
                ('Albendazole','Albendazole'),
                ('Acyclovir','Acyclovir'),
                
                )
    ),
    ('AB-C',(
                ('Cefazoline', 'Cefazoline'),
                ('Ceftriaxone', 'Ceftriaxone'),
                ('Cefotaxime','Cefotaxime'),
                ('Cefepim','Cefepim'),
                ('Ceftarolin','Ceftarolin'),
                ('Ceftazidim','Ceftazidim'),
                ('Ceftazidim/avibactam','Ceftazidim/avibactam'),
                ('Cftobiprole','Cftobiprole'),
                )
            ),
    ('AB-E',(
                ('Ertapenem', 'Ertapenem'),
                ('Tigecycline', 'Tigecycline'),
                
                
                
                )
            ),

             
    ('AB-F',(
                ('Flucloxacillin ', 'Flucloxacillin'),
                ('Fluconazole', 'Fluconazole'),
                
                
                )
            ),
    ('AB-G',(
                ('Gentamycin', 'Gentamycin'),
                ('Ganciclovir', 'Ganciclovir'),
                
                
                
                )
            ),
    ('AB-P',(
                ('Penicillin G', 'Penicillin G'),
                ('Piperacillin/tazobactam', 'Piperacillin/tazobactam'),
                ('Pentamidine', 'Pentamidine'),
                
                
                )
            ),

    ('AB-V',(
                ('Vancomycin', 'Vancomycin'),
                ('Voriconazole', 'Voriconazole'),
                
                
                
                )
            ),
    ('AB-T',(
                ('Teicoplanin', 'Teicoplanin'),
                ('Tigecycline', 'Tigecycline'),
                
                
                
                )
            ),
    ('AB-M',(
                ('Meropenem', 'Meropenem'),
                ('Metronidazole', 'Metronidazole'),
                
                
                
                )
            ),
    ('AB-I',(
                ('Imipenem cilastatin', 'Imipenem cilastatin'),
                ('Isavuconazole', 'Isavuconazole'),
                
                
                
                )
            ),
             
    ('AB-L',(
                ('Levofloxacin', 'Levofloxacin'),
                ('Linezolid', 'Linezolid'),
                
                
                
                )
            ),
              
    ('AB-V',(
                ('Voriconazole', 'Voriconazole'),
                ('Ganciclovir', 'Ganciclovir'),
                
                
                
                )
            ),
    ('AB-R',(
                ('Remdesivir', 'Remdesivir'),
                
                
                
                
                )
            ),
   
]
    AB_INTERVENTION_CHOICES = [

            ('Intervention',(
                ('Escalution-need','Escalution'),
                ('De-escalution-need','De-escalution'),
                ('Renal-dose-adjustment','Renal-dose-adjustment'),
                ('hepatic-dose-adjustment','hepatic-dose-adjustment'),
                ('Renal-dose-adjustment','Renal-dose-adjustment'),
                ('Discontinue','Discontinue'),
                )
            ),
    ]

    AB_MP_CHOICES = [

        
            ('Monitoring Parameter',(
                ('Markers','Markers' ),
                ('Platelets','Platelets' ),
                ('QT-c','QT-c' ),
                ('Level','Level' ),
                ('Interaction','Interaction' ),
                ('Candida_Score','Candida_Score' ),
                ('Lactate','Lactate' ),
                ),
            ),

       ]

    

    MRD = forms.ChoiceField(label="MRD",choices=available_stations)    
    date = forms.DateField()
    GFR = forms.IntegerField(widget=forms.TextInput(attrs={'placeholder':'GFR'}))
    Child_pugh_score = forms.IntegerField(widget=forms.TextInput(attrs={'placeholder':'CPS'}))
    Liver_Imparrenenty = forms.BooleanField(required=False)
    Dose_adjustment = forms.BooleanField(required=False)
    Urine_output = forms.ChoiceField(choices=URINE_OUTPUT_CHOICES)
    Balance = forms.ChoiceField(choices=BALANCE_CHOICES)
    intervention =forms.CharField(required=False)
    Feeding = forms.ChoiceField(choices=FEEDING_CHOICES)
    Bowel_motion = forms.CharField(required=False)    
    Electrolytes_imbalance = forms.ChoiceField(choices=ELECTROLYTES_CHOICES)    
    Hyper_Hypo = forms.ChoiceField(choices=HYPER_HYPO_CHOICES)
    ABI = forms.ChoiceField(choices=ABI_CHOICES) 
    Metabolic_num = forms.IntegerField(widget=forms.TextInput(attrs={'placeholder':'Metabolic'}))
    Respiratory_num = forms.IntegerField(widget=forms.TextInput(attrs={'placeholder':'Respiratory'}))
    Metabolic = forms.ChoiceField(choices=METABOLIC_CHOICES)
    Respiratory = forms.ChoiceField(choices=RESPIRATORY_CHOICES)
    QT_C_num  = forms.IntegerField(widget=forms.TextInput(attrs={'placeholder':'Respiratory'}))
    QT_C= forms.ChoiceField(choices=QT_C_CHOICES)
    VITALS= forms.ChoiceField(choices=VITALS_CHOICES)
    Analgesic_management  = forms.BooleanField(required=False)
    Sedation = forms.BooleanField(required=False)
    Thromboembolic_Prophylaxis = forms.ChoiceField(choices=TP_CHOICES)
    Stress_Ulcer_Pophylaxis = forms.ChoiceField(choices=SUP_CHOICES )
    Glycemic_control_target_BG  = forms.BooleanField(required=False)
    T_BG = forms.ChoiceField(  choices=T_BG_CHOICES )
    Infection  = forms.ChoiceField(  choices=INFECTION_CHOICES )
    Treatment  = forms.ChoiceField(  choices=TREATMENT_CHOICES )
    AB_INTERVENTION  = forms.ChoiceField( choices=AB_INTERVENTION_CHOICES )
    MP_LIST  = forms.ChoiceField( choices=AB_MP_CHOICES )









    