from io import BytesIO
import xlsxwriter
from django.utils.translation import ugettext

def WriteToExcel(weather_data, town=None):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)

    # define styles to use
    title = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
    })
    header = workbook.add_format({
            'bg_color': 'gray',
            'color': 'black',
            'align': 'center',
            'valign': 'top',
            'border': 1,
           
            
            
    })
    header_2 = workbook.add_format({
            'bg_color': 'green',
            'color': 'black',
            'align': 'center',
            'valign': 'top',
            'border': 2,
           
            
            
    })
    body = workbook.add_format({
            'bg_color': '#cfe7f5',
            'color': 'black',
            'align': 'left',
            'valign': 'top',
            
            
    })
   
         
    cell = workbook.add_format({})
    cell_center = workbook.add_format({'align': 'center'})

    # add a worksheet to work with
    worksheet_s = workbook.add_worksheet("Summary")
    worksheet_T = workbook.add_worksheet("Tables")


    # create title
    title_text = ("Hospital Report") % {'MRD': town}
    # add title to sheet, use merge_range to let title span over multiple columns
    worksheet_s.merge_range('B2:H2', title_text, title)

    # Add headers for data
    worksheet_s.write(3, 0, ugettext("Date"), header_2)
    worksheet_s.write(3, 1, ugettext("patient"), header_2)

    worksheet_s.write(7, 0, ugettext(u"GFR"), header)
    worksheet_s.write(7, 1, ugettext(u"Child_pugh_score"), header)    
    worksheet_s.write(7, 2, ugettext(u"Liver Imparrenenty"), header)
    worksheet_s.write(7, 3, ugettext(u"Dose_adjustment"), header)

    worksheet_s.write(9, 0, ugettext(u"Balance"), header)
    worksheet_s.write(9, 1, ugettext(u"intervention"), header)
    worksheet_s.write(9, 2, ugettext(u"Urine output"), header)
    worksheet_s.write(9, 3, ugettext(u"Feeding"), header)

    worksheet_s.write(11, 0, ugettext(u"Hyper Hypo"), header)
    worksheet_s.write(11, 1, ugettext(u"ABI"), header)
    worksheet_s.write(11, 2, ugettext(u"Metabolic num"), header)
    worksheet_s.write(11, 3, ugettext(u"Respiratory num"), header)

    worksheet_s.write(13, 0, ugettext(u"Metabolic"), header)
    worksheet_s.write(13, 1, ugettext(u"Respiratory"), header)
    worksheet_s.write(13, 2, ugettext(u"QT_C"), header)
    worksheet_s.write(13, 3, ugettext(u"QT_C_num"), header)

    worksheet_s.write(15, 0, ugettext(u"Analgesic management"), header)
    worksheet_s.write(15, 1, ugettext(u"Sedation"), header)
    worksheet_s.write(15, 2, ugettext(u"Thromboembolic Prophylaxis"), header)
    worksheet_s.write(15, 3, ugettext(u"Stress Ulcer Pophylaxis"), header)

    worksheet_s.write(17, 0, ugettext(u"Glycemic control target_BG"), header)
    worksheet_s.write(17, 1, ugettext(u"Bowel motion"), header)


    # write data to cells
    station_name_width = 26
    for idx, data in enumerate(weather_data):
        row_index = 4 + idx
        row_index_2 = 8 + idx
        row_index_3 = 10 + idx
        row_index_4 = 12 + idx
        row_index_5 = 14 + idx
        row_index_6 = 16 + idx
        row_index_7 = 18 + idx



       
        worksheet_s.write(row_index, 0, data.date.strftime("%Y-%m-%d"), cell_center)
        worksheet_s.write_string(row_index, 1, data.patient, cell_center)

        worksheet_s.write(row_index_2,0, data.GFR, cell_center)
        worksheet_s.write(row_index_2,1, data.Child_pugh_score, cell_center)
        worksheet_s.write(row_index_2,2, data.Liver_Imparrenenty, cell_center)
        worksheet_s.write(row_index_2,3, data.Dose_adjustment, cell_center)

        worksheet_s.write(row_index_3,0, data.Balance, cell_center)
        worksheet_s.write(row_index_3,1, data.intervention, cell_center)
        worksheet_s.write(row_index_3,2, data.Urine_output, cell_center)
        worksheet_s.write(row_index_3,3, data.Feeding, cell_center)

        worksheet_s.write(row_index_4,0, data.Hyper_Hypo, cell_center)
        worksheet_s.write(row_index_4,1, data.ABI, cell_center)
        worksheet_s.write(row_index_4,2, data.Metabolic_num, cell_center)
        worksheet_s.write(row_index_4,3, data.Respiratory_num, cell_center)

        worksheet_s.write(row_index_5,0, data.Metabolic, cell_center)
        worksheet_s.write(row_index_5,1, data.Respiratory, cell_center)
        worksheet_s.write(row_index_5,2, data.QT_C, cell_center)
        worksheet_s.write(row_index_5,3, data.QT_C_num, cell_center)

        worksheet_s.write(row_index_6,0, data.Analgesic_management, cell_center)
        worksheet_s.write(row_index_6,1, data.Sedation, cell_center)
        worksheet_s.write(row_index_6,2, data.Thromboembolic_Prophylaxis, cell_center)
        worksheet_s.write(row_index_6,3, data.Stress_Ulcer_Pophylaxis, cell_center)

        worksheet_s.write(row_index_7,0, data.Glycemic_control_target_BG, cell_center)
        worksheet_s.write(row_index_7,1, data.Bowel_motion, cell_center)




        data = [
               ['Electrolytes_imbalance', 'Na','K','Mg','Po4','Ca'],
               
               ['VITALS',    'stable','tech carbon','tachypnea','htn','bradycardia','bradypnea','hypotension','Fever'],
               
               ['T_BG',   'Hypoglycomic','Glucagon','D50%','D5%/D10%','koton','Hydration','Hyperglycemia','Insulin Infusion','DKA protocol','SQ Sliding Scale','Long acting insulin'],
               
               ['Infection','UA','Markers','Fever','Culture','X-Ray','Swab'],

               ['Treatment',   'Ampicillin' ,'Amoxiclav','Amikacin','Azithromycin','Anidulafungin','Albendazole','Acyclovir',
               'Cefazoline','Ceftriaxone','Cefotaxime','Cefepim','Ceftarolin','Ceftazidim','Ceftazidim','Cftobiprole','Ertapenem',
               'Tigecycline','Flucloxacillin','Fluconazole','Gentamycin','Ganciclovir','Penicillin G','Piperacillin/tazobactam','Pentamidine',
               'Vancomycin','Voriconazole','Teicoplanin','Tigecycline','Meropenem','Metronidazole','Imipenem', 'cilastatin','Isavuconazole',
               'Levofloxacin','Linezolid','Voriconazole','Ganciclovir','Remdesivir'],
               
               ['AB_INTERVENTION',   'Escalution-need','De-escalution-need','Renal-dose-adjustment','hepatic-dose-adjustment','Renal-dose-adjustment','Discontinue'],
               
               ['Monitoring Parameter', 'Markers','Platelets','QT-c','Level','Interaction','Candida_Score','Lactate',],
            ]

        worksheet_T.add_table("A1:AM8", {'data':data ,'columns': [
               {'header': 'Name'},
               {'header': 'Group1'},
               {'header': 'Group2'},
               {'header': 'Group3'},
               {'header': 'Group4'},
               {'header': 'Group5'},
               {'header': 'Group6'},
               {'header': 'Group7'}
               ]
            })

       
        

        
        worksheet_s.set_column('A:A', 26)
        worksheet_s.set_column('B:B', 30)
        worksheet_s.set_column('C:C', 30)
        worksheet_s.set_column('D:D', 26)
        worksheet_s.set_column('E:E', 30)
        worksheet_s.set_column('F:F', 30)
        worksheet_s.set_column('G:G', 26)
        worksheet_s.set_column('H:H', 30)
        worksheet_s.set_column('I:I', 30)
        worksheet_s.set_column('J:J', 26)
        worksheet_s.set_column('K:K', 30)
        worksheet_s.set_column('L:L', 30)
        worksheet_s.set_column('M:M', 26)
        worksheet_s.set_column('N:N', 30)
        worksheet_s.set_column('O:O', 30)
        worksheet_s.set_column('P:P', 26)
        worksheet_s.set_column('Q:Q', 30)
        worksheet_s.set_column('R:R', 30)
        worksheet_s.set_column('S:S', 26)
        worksheet_s.set_column('T:T', 30)
        worksheet_s.set_column('U:U', 30)
        worksheet_s.set_column('V:V', 26)
        worksheet_s.set_column('W:W', 30)
        worksheet_s.set_column('X:X', 30)
        worksheet_s.set_column('Y:Y', 30)
        worksheet_s.set_column('Z:Z', 26)
        worksheet_s.set_column('AA:AA', 30)
        worksheet_s.set_column('AB:AB', 30)
        worksheet_s.set_column('AC:AC', 26)
        worksheet_s.set_column('AD:AD', 30)
        worksheet_s.set_column('AE:AE', 30)



        worksheet_T.set_column('A:A', 26)
        worksheet_T.set_column('B:B', 30)
        worksheet_T.set_column('C:C', 30)
        worksheet_T.set_column('D:D', 26)
        worksheet_T.set_column('E:E', 30)
        worksheet_T.set_column('F:F', 30)
        worksheet_T.set_column('G:G', 26)
        worksheet_T.set_column('H:H', 30)
        worksheet_T.set_column('I:I', 30)
        worksheet_T.set_column('J:J', 26)
        worksheet_T.set_column('K:K', 30)
        worksheet_T.set_column('L:L', 30)
        worksheet_T.set_column('M:M', 26)
        worksheet_T.set_column('N:N', 30)
        worksheet_T.set_column('O:O', 30)
        worksheet_T.set_column('P:P', 26)
        worksheet_T.set_column('Q:Q', 30)
        worksheet_T.set_column('R:R', 30)
        worksheet_T.set_column('S:S', 26)
        worksheet_T.set_column('T:T', 30)
        worksheet_T.set_column('U:U', 30)
        worksheet_T.set_column('V:V', 26)
        worksheet_T.set_column('W:W', 30)
        worksheet_T.set_column('X:X', 30)
        worksheet_T.set_column('Y:Y', 30)
        worksheet_T.set_column('Z:Z', 26)
        worksheet_T.set_column('AA:AA', 30)
        worksheet_T.set_column('AB:AB', 30)
        worksheet_T.set_column('AC:AC', 26)
        worksheet_T.set_column('AD:AD', 30)
        worksheet_T.set_column('AE:AE', 30)
        worksheet_T.set_column('AF:AF', 30)
        worksheet_T.set_column('AG:AG', 30)
        worksheet_T.set_column('AI:AI', 26)
        worksheet_T.set_column('AJ:AJ', 30)
        worksheet_T.set_column('AK:AK', 30)
        worksheet_T.set_column('AL:AL', 26)
        worksheet_T.set_column('AM:AM', 30)
        worksheet_T.set_column('AF:AF', 30)

    # add Report
    #worksheet_x = workbook.add_worksheet("Report")

    # Report data
    #for row_index, data in enumerate(weather_data):
        #worksheet_x.write(row_index, 0, data.date.strftime("%Y-%m-%d"))
        #worksheet_x.write(row_index, 1, data.Electrolytes_imbalance)
        #worksheet_x.write(row_index, 2, data.VITALS)
        #worksheet_x.write(row_index, 3, data.T_BG)
    
    # add formulas to calc avg. over time
    worksheet_s.write_formula(row_index, 5, '=AVERAGE({0}{1}:{0}{2})'.format('C', 6, row_index+1))
    worksheet_s.write_formula(row_index, 6, '=AVERAGE({0}{1}:{0}{2})'.format('D', 6, row_index+1))
    worksheet_s.write_formula(row_index, 7, '=AVERAGE({0}{1}:{0}{2})'.format('E', 6, row_index+1))

    

   
    worksheet_c = workbook.add_worksheet("Charts")
    worksheet_d = workbook.add_worksheet("Chart data")
    # chart data
    for row_index, data in enumerate(weather_data):
        worksheet_d.write(row_index, 0, data.date.strftime("%Y-%m-%d"))
        worksheet_d.write(row_index, 1, data.GFR)
        worksheet_d.write(row_index, 2, data.Child_pugh_score)
        worksheet_d.write(row_index, 3, data.QT_C_num)

    # line chart
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
            'categories': '=Chart data!$A1:$A{0}'.format(len(weather_data)),
            'values': '=Chart data!$B1:$B{0}'.format(len(weather_data)),
            'marker': {'type': 'square'},
            'name': ugettext("GFR")
    })
    line_chart.add_series({
            'categories': '=Chart data!$A1:$A{0}'.format(len(weather_data)),
            'values': '=Chart data!$C1:$C{0}'.format(len(weather_data)),
            'marker': {'type': 'square'},
            'name': ugettext("Child_pugh_score")
    })
    line_chart.add_series({
            'categories': '=Chart data!$A1:$A{0}'.format(len(weather_data)),
            'values': '=Chart data!$D1:$D{0}'.format(len(weather_data)),
            'marker': {'type': 'square'},
            'name': ugettext("QT_C_num")
        })
    line_chart.set_title({'name': ugettext("Report Chart")})
    line_chart.set_x_axis({ 'text_axis': True, 'date_axis': False })
    line_chart.set_y_axis({ 'num_format': u'## °C' })
    worksheet_c.insert_chart('B2', line_chart, {'x_scale': 2, 'y_scale': 1})

    workbook.close()
    xlsx_data = output.getvalue()
    # xlsx_data contains the Excel file
    return xlsx_data

