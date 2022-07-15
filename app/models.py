from django.db import models


class HospitalMRD(models.Model):
    MRD = models.CharField(max_length=256)



    class Meta:
        verbose_name = 'HospitalMRD'
        verbose_name_plural = ' Hospital MRD '

    def __str__(self):
        return self.MRD


class Temperature(models.Model):
    station = models.ForeignKey(HospitalMRD, on_delete=models.CASCADE)
    date = models.DateField()
    max = models.DecimalField(max_digits=4, decimal_places=1)
    mean = models.DecimalField(max_digits=4, decimal_places=1)
    min = models.DecimalField(max_digits=4, decimal_places=1)

    def __str__(self):
        return str(self.date)

class ClinicalPharmacist(models.Model):
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

    
    MRD = models.ForeignKey(HospitalMRD, on_delete=models.CASCADE)
    patient = models.CharField( max_length=512, default=False)
    date = models.DateField()
    GFR = models.IntegerField( null=True, blank=True)
    renal_impairment = models.BooleanField(null=True, blank=True)
    Liver_Imparrenenty = models.BooleanField(null=True, blank=True)
    Child_pugh_score = models.IntegerField( default=True)
    Dose_adjustment  = models.BooleanField(null=True, blank=True)
    Balance  = models.CharField( max_length=512, choices=BALANCE_CHOICES, default=False)
    intervention = models.BooleanField(null=True, blank=True)
    Dose_adjustment_LI = models.BooleanField(null=True, blank=True)
    Urine_output  = models.CharField( max_length=512, choices=URINE_OUTPUT_CHOICES, default=False)
    Feeding  = models.CharField( max_length=512, choices=FEEDING_CHOICES, default=False)
    Bowel_motion  = models.CharField( max_length=512, default=False)
    Electrolytes_imbalance = models.CharField(max_length=512,choices=ELECTROLYTES_CHOICES , default=False)    
    Hyper_Hypo  = models.CharField( max_length=512, choices=HYPER_HYPO_CHOICES, default=False)
    ABI  = models.CharField(max_length=512,choices=ABI_CHOICES , default=False)    
    Metabolic_num = models.IntegerField( null=True, blank=True)
    Respiratory_num = models.IntegerField(null=True, blank=True)
    Metabolic   = models.CharField( max_length=512, choices=METABOLIC_CHOICES, default=False)
    Respiratory  = models.CharField( max_length=512, choices=RESPIRATORY_CHOICES, default=False)
    QT_C  = models.CharField( max_length=512, choices=QT_C_CHOICES, default=False)
    QT_C_num = models.IntegerField( null=True, blank=True)
    VITALS  = models.CharField( max_length=512, choices=VITALS_CHOICES , default=False)
    Analgesic_management  = models.BooleanField(null=True, blank=True)
    Sedation = models.BooleanField(null=True, blank=True)
    Thromboembolic_Prophylaxis = models.CharField( max_length=512, choices=TP_CHOICES, default=False)
    Stress_Ulcer_Pophylaxis = models.CharField( max_length=512, choices=SUP_CHOICES , default=False)
    Glycemic_control_target_BG  = models.BooleanField(null=True, blank=True)
    T_BG = models.CharField( max_length=512, choices=T_BG_CHOICES , default=False)
    Infection  = models.CharField( max_length=512, choices=INFECTION_CHOICES , default=False)
    Treatment  = models.CharField( max_length=512, choices=TREATMENT_CHOICES , default=False)
    AB_INTERVENTION  = models.CharField( max_length=512, choices=AB_INTERVENTION_CHOICES , default=False)
    MP_LIST  = models.CharField( max_length=512, choices=AB_MP_CHOICES , default=False)


    status = models.BooleanField(default= False)

   
    # could change req on thge model :
    def save(self,*args, **kwargs):
        if self.GFR==0 or None:
            self.renal_impairment=False

        if self.Liver_Imparrenenty == True :
            self.Dose_adjustment    
            
        if self.Electrolytes_imbalance == True :
            self.Electrolytes_imbalance == True    
            
        super().save(*args, **kwargs)    
    




    def __str__(self):
        return '{}-{}-{}-{}'.format(self.MRD,self.patient,self.date ,self.status)    

class AddNote(models.Model):

    patient = models.ForeignKey(ClinicalPharmacist, on_delete=models.CASCADE,default=False)
    description = models.CharField(max_length=500 , null=True, blank=True)
    message = models.TextField()
    date = models.DateField(auto_now_add=True)
    

    
    class Meta:
        verbose_name = 'AddNote'
        verbose_name_plural = 'Add Notes'
    
    def __str__(self):
        
        return '{}-{}'.format(self.patient, self.date)    

