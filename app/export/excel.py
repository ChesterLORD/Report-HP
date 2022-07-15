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
            'bg_color': 'green',
            'color': 'black',
            'align': 'center',
            'valign': 'top',
            'border': 1,
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

    # create title
    title_text = ("Hospital Report") % {'MRD': town}
    # add title to sheet, use merge_range to let title span over multiple columns
    worksheet_s.merge_range('B2:H2', title_text, title)

    # Add headers for data
    worksheet_s.write(3, 0, ugettext("Date"), header)
    worksheet_s.write(3, 1, ugettext("patient"), header)
    worksheet_s.write(3, 2, ugettext(u"GFR"), header)
    worksheet_s.write(3, 3, ugettext(u"Renal impairment"), header)    
    worksheet_s.write(3, 4, ugettext(u"Liver Imparrenenty"), header)
    worksheet_s.write(3, 5, ugettext(u"Child pugh_score"), header)
    worksheet_s.write(3, 6, ugettext(u"Dose adjustment"), header)
    worksheet_s.write(3, 7, ugettext(u"Balance"), header)
    worksheet_s.write(3, 8, ugettext(u"intervention"), header)
    worksheet_s.write(3, 9, ugettext(u"Dose adjustment_LI"), header)
    worksheet_s.write(3, 10, ugettext(u"Urine output"), header)
    worksheet_s.write(3, 11, ugettext(u"Feeding"), header)
    worksheet_s.write(3, 12, ugettext(u"Bowel motion"), header)
    worksheet_s.write(3, 13,ugettext(u"Electrolytes imbalance"), header)
    worksheet_s.write(3, 14, ugettext(u"Hyper Hypo"), header)
    worksheet_s.write(3, 15, ugettext(u"ABI"), header)
    worksheet_s.write(3, 16, ugettext(u"Metabolic num"), header)
    worksheet_s.write(3, 17, ugettext(u"Respiratory num"), header)
    worksheet_s.write(3, 18, ugettext(u"Metabolic"), header)
    worksheet_s.write(3, 19, ugettext(u"Respiratory"), header)
    worksheet_s.write(3, 20, ugettext(u"QT_C"), header)
    worksheet_s.write(3, 21, ugettext(u"QT_C_num"), header)
    worksheet_s.write(3, 22, ugettext(u"VITALS"), header)
    worksheet_s.write(3, 23, ugettext(u"Analgesic management"), header)
    worksheet_s.write(3, 24, ugettext(u"Sedation"), header)
    worksheet_s.write(3, 25, ugettext(u"Thromboembolic Prophylaxis"), header)
    worksheet_s.write(3, 26, ugettext(u"Stress Ulcer Pophylaxis"), header)
    worksheet_s.write(3, 27, ugettext(u"Glycemic control target_BG"), header)
    worksheet_s.write(3, 28, ugettext(u"T_BG"), header)
    worksheet_s.write(3, 29, ugettext(u"Infection"), header)
    worksheet_s.write(3, 30, ugettext(u"Treatment"), header)
    worksheet_s.write(3, 31, ugettext(u"AB INTERVENTION"), header)
    worksheet_s.write(3, 32, ugettext(u"MP LIST"), header)

    # write data to cells
    station_name_width = 26
    for idx, data in enumerate(weather_data):
        row_index = 4 + idx
        worksheet_s.write(row_index, 0, data.date.strftime("%Y-%m-%d"), cell_center)
        worksheet_s.write_string(row_index, 1, data.patient, cell_center)
        worksheet_s.write(row_index,2, data.GFR, cell_center)
        worksheet_s.write(row_index,3, data.renal_impairment, cell_center)
        worksheet_s.write(row_index,4, data.Liver_Imparrenenty, cell_center)
        worksheet_s.write(row_index,5, data.Child_pugh_score, cell_center)
        worksheet_s.write(row_index,6, data.Dose_adjustment, cell_center)
        worksheet_s.write(row_index,7, data.Balance, cell_center)
        worksheet_s.write(row_index,8, data.intervention, cell_center)
        worksheet_s.write(row_index,9, data.Dose_adjustment_LI, cell_center)
        worksheet_s.write(row_index,10, data.Urine_output, cell_center)
        worksheet_s.write(row_index,11, data.Feeding, cell_center)
        worksheet_s.write(row_index,12, data.Bowel_motion, cell_center)
        worksheet_s.write(row_index,13, data.Electrolytes_imbalance, cell_center)
        worksheet_s.write(row_index,14, data.Hyper_Hypo, cell_center)
        worksheet_s.write(row_index,15, data.ABI, cell_center)
        worksheet_s.write(row_index,16, data.Metabolic_num, cell_center)
        worksheet_s.write(row_index,17, data.Respiratory_num, cell_center)
        worksheet_s.write(row_index,18, data.Metabolic, cell_center)
        worksheet_s.write(row_index,19, data.Respiratory, cell_center)
        worksheet_s.write(row_index,20, data.QT_C, cell_center)
        worksheet_s.write(row_index,21, data.QT_C_num, cell_center)
        worksheet_s.write(row_index,22, data.VITALS, cell_center)
        worksheet_s.write(row_index,23, data.Analgesic_management, cell_center)
        worksheet_s.write(row_index,24, data.Sedation, cell_center)
        worksheet_s.write(row_index,25, data.Thromboembolic_Prophylaxis, cell_center)
        worksheet_s.write(row_index,26, data.Stress_Ulcer_Pophylaxis, cell_center)
        worksheet_s.write(row_index,27, data.Glycemic_control_target_BG, cell_center)
        worksheet_s.write(row_index,28, data.T_BG, cell_center)
        worksheet_s.write(row_index,29, data.Infection, cell_center)
        worksheet_s.write(row_index,30, data.Treatment, cell_center)
        worksheet_s.write(row_index,31, data.AB_INTERVENTION, cell_center)
        worksheet_s.write(row_index,32, data.MP_LIST, cell_center)
        

        
        

        # add formulas to calc avg. over time
        worksheet_s.write_formula(row_index, 5, '=AVERAGE({0}{1}:{0}{2})'.format('C', 6, row_index+1))
        worksheet_s.write_formula(row_index, 6, '=AVERAGE({0}{1}:{0}{2})'.format('D', 6, row_index+1))
        worksheet_s.write_formula(row_index, 7, '=AVERAGE({0}{1}:{0}{2})'.format('E', 6, row_index+1))

    # resize rows and columns
    worksheet_s.set_column('A:A', 26)
    worksheet_s.set_column('B:B', 12)
    worksheet_s.set_column('C:H', 12)

    # add chart
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

