from io import BytesIO
import reportlab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.charts.linecharts import SampleHorizontalLineChart
from reportlab.graphics.widgets.markers import makeMarker
from reportlab.graphics.shapes import Drawing

from django.conf import settings


def WriteToPdf(weather_data, town=None):
    buffer = BytesIO()
    report = PdfPrint(buffer)
    pdf = report.report(weather_data, "Hospital report")
    return pdf


class PdfPrint:
    def __init__(self, buffer):
        self.buffer = buffer
        self.pageSize = A4
        self.width, self.height = self.pageSize

    def report(self, weather_history, title):
        doc = SimpleDocTemplate(
            self.buffer, rightMargin=72, leftMargin=72,
            topMargin=30, bottomMargin=72,
            pagesize=self.pageSize
        )
        # register fonts
        freesans = settings.BASE_DIR + settings.STATIC_URL + "FreeSans.ttf"
        freesansbold = settings.BASE_DIR + settings.STATIC_URL + "FreeSansBold.ttf"
        pdfmetrics.registerFont(TTFont('FreeSans', freesans))
        pdfmetrics.registerFont(TTFont('FreeSansBold', freesansbold))
        # set up styles
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(
            name="TableHeader", fontSize=9, alignment=TA_CENTER,
            fontName="FreeSansBold"))
        styles.add(ParagraphStyle(
            name="ParagraphTitle", fontSize=9, alignment=TA_JUSTIFY,
            fontName="FreeSansBold"))
        styles.add(ParagraphStyle(
            name="Justify", alignment=TA_JUSTIFY, fontName="FreeSans"))

        data = []
        data.append(Paragraph(title, styles["Title"]))
        data.append(Spacer(1, 12))
        table_data = []
        # table header
        table_data.append([
            Paragraph('Date', styles['TableHeader']),
            Paragraph('Station', styles['TableHeader']),
        ])
        for wh in weather_history:
            # add a row to table
            table_data.append([
                wh.date,
                Paragraph(wh.patient, styles['Justify']),
            ])
        # create table1

        wh_table = Table(table_data, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))




       
        table_data_2 = []
         # table header
        table_data_2.append([
            Paragraph('GFR', styles['TableHeader']),
            Paragraph('Child_pugh_score', styles['TableHeader']),
            Paragraph('Liver_Imparrenenty', styles['TableHeader']),
        ])
        for wh in weather_history:
         # add a row to table
            table_data_2.append([
                u"{0} ".format(wh.GFR),
                u"{0} ".format(wh.Child_pugh_score),
                u"{0} ".format(wh.Liver_Imparrenenty),
                

            ])

        # create table2

        wh_table = Table(table_data_2,  colWidths=200, rowHeights=50)


        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))


#-----------------------------------------------


       
        table_data_3 = []
         # table header
        table_data_3.append([
            Paragraph('Dose_adjustment', styles['TableHeader']),
            Paragraph('Balance', styles['TableHeader']),
        ])
        for wh in weather_history:
         # add a row to table
            table_data_3.append([
                u"{0} ".format(wh.Dose_adjustment),
                u"{0} ".format(wh.Balance),
                

            ])

        # create table2

        wh_table = Table(table_data_3,  colWidths=283, rowHeights=50)


        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))

#-----------------------------------------------------------        

        table_data_4 = []
         # table header
        table_data_4.append([
            Paragraph('intervention', styles['TableHeader']),
            Paragraph('Urine_output', styles['TableHeader']),
            Paragraph('Feeding', styles['TableHeader']),
            

        ])
        for wh in weather_history:
         # add a row to table
            table_data_4.append([
                u"{0} ".format(wh.intervention),
                u"{0} ".format(wh.Urine_output),
                u"{0} ".format(wh.Feeding),
              
            ])


        # create table3

        wh_table = Table(table_data_4, colWidths=200, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))

#-----------------------------------------------------------
        table_data_5 = []
         # table header
        table_data_5.append([
            
            Paragraph('Bowel_motion', styles['TableHeader']),
            Paragraph('Electrolytes_imbalance', styles['TableHeader']),

        ])
        for wh in weather_history:
         # add a row to table
            table_data_5.append([
               
                u"{0} ".format(wh.Bowel_motion),
                u"{0} ".format(wh.Electrolytes_imbalance),
            ])


        # create table3

        wh_table = Table(table_data_5, colWidths=283, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))




#-----------------------------------------------------------        




        table_data_6= []
         # table header
        table_data_6.append([
            Paragraph('Hyper_Hypo', styles['TableHeader']),
            Paragraph('ABI', styles['TableHeader']),
            Paragraph('Metabolic_num', styles['TableHeader']),
            
        ])
        for wh in weather_history:
         # add a row to table
            table_data_6.append([
                u"{0} ".format(wh.Hyper_Hypo),
                u"{0} ".format(wh.ABI),
                u"{0} ".format(wh.Metabolic_num),
               
            ])

        # create table3

        wh_table = Table(table_data_6 , colWidths=200, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    


#---------------------------------------------------------------------
        table_data_7= []
         # table header
        table_data_7.append([
            
            Paragraph('Respiratory_num', styles['TableHeader']),
            Paragraph('Metabolic', styles['TableHeader']),
            Paragraph('Respiratory', styles['TableHeader'])
        ])
        for wh in weather_history:
         # add a row to table
            table_data_7.append([
                
                u"{0} ".format(wh.Respiratory_num),
                u"{0} ".format(wh.Metabolic),
                u"{0} ".format(wh.Respiratory)
            ])

        # create table3

        wh_table = Table(table_data_7 , colWidths=200, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#------------------------------------------------------------------------


        table_data_8 = []
         # table header
        table_data_8.append([
            Paragraph('QT_C', styles['TableHeader']),
            Paragraph('QT_C_num', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_8.append([
                u"{0} ".format(wh.QT_C),
                u"{0} ".format(wh.QT_C_num),
             
            ])



        # create table3

        wh_table = Table(table_data_8, colWidths=283, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))
#-----------------------------------------------------------------------
        table_data_9 = []
         # table header
        table_data_9.append([
            
            Paragraph('Analgesic_management', styles['TableHeader']),
            Paragraph('Sedation', styles['TableHeader'])
        ])
        for wh in weather_history:
         # add a row to table
            table_data_9.append([
               
                u"{0} ".format(wh.Analgesic_management),
                u"{0} ".format(wh.Sedation)
            ])



        # create table3

        wh_table = Table(table_data_9, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))

#---------------------------------------------------------------------------               
        


        table_data_10 = []
         # table header
        table_data_10.append([
            Paragraph('Thromboembolic_Prophylaxis', styles['TableHeader']),
            Paragraph('Stress_Ulcer_Pophylaxis', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_10.append([
                u"{0} ".format(wh.Thromboembolic_Prophylaxis),
                u"{0} ".format(wh.Stress_Ulcer_Pophylaxis),
               
            ])


        # create table3

        wh_table = Table(table_data_10, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))
#----------------------------------------------------------------------------

        table_data_11 = []
         # table header
        table_data_11.append([
            
            Paragraph('Glycemic_control_target_BG', styles['TableHeader']),
            Paragraph('Infection', styles['TableHeader'])
        ])
        for wh in weather_history:
         # add a row to table
            table_data_11.append([
                
                u"{0} ".format(wh.Glycemic_control_target_BG),
                u"{0} ".format(wh.Infection)
            ])


        # create table3

        wh_table = Table(table_data_11, colWidths=283, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))

#----------------------------------------------------------------------                
        

        table_data_12 = []
         # table header
        table_data_12.append([
            Paragraph('Treatment', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_12.append([
                u"{0} ".format(wh.Treatment),
               
           
            ])

        wh_table = Table(table_data_12,colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#---------------------------------------------------------------------
#----------------------------------------------------------------------                
        

        table_data_13 = []
         # table header
        table_data_13.append([
            
            Paragraph('AB_INTERVENTION', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_13.append([
               
                u"{0} ".format(wh.AB_INTERVENTION),
                
           
            ])

        wh_table = Table(table_data_13, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#--------------------------------------------------------------------- 
#----------------------------------------------------------------------                
        

        table_data_14 = []
         # table header
        table_data_14.append([
           
            Paragraph('MP_LIST', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_14.append([
              
                u"{0} ".format(wh.MP_LIST),
           
            ])

        wh_table = Table(table_data_14, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#---------------------------------------------------------------------
#----------------------------------------------------------------------                
        

        table_data_15 = []
         # table header
        table_data_15.append([
           
            Paragraph('VITALS', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_15.append([
              
                u"{0} ".format(wh.VITALS),
           
            ])

        wh_table = Table(table_data_15, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#--------------------------------------------------------------------- 
#----------------------------------------------------------------------                
        

        table_data_16 = []
         # table header
        table_data_16.append([
           
            Paragraph('T_BG', styles['TableHeader']),
           
        ])
        for wh in weather_history:
         # add a row to table
            table_data_16.append([
              
                u"{0} ".format(wh.T_BG),
           
            ])

        wh_table = Table(table_data_16, colWidths=280, rowHeights=50)
        wh_table.hAlign = 'CENTER'
        wh_table.setStyle(TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
             ('BACKGROUND', (0, 0), (-1, 0), colors.gray)]))
        data.append(wh_table)
        data.append(Spacer(1,1))    
#---------------------------------------------------------------------  









        
        
        # create line chart
        chart = SampleHorizontalLineChart()
        chart.width = 350
        chart.height = 135
        mins = [float(x.GFR) for x in weather_history]
        means = [float(x.Child_pugh_score) for x in weather_history]
        maxs = [float(x.QT_C_num) for x in weather_history]
        chart.data = [mins, means, maxs]
        chart.lineLabels.fontName = 'FreeSans'
        chart.strokeColor = colors.white
        chart.fillColor = colors.lightblue
        chart.lines[0].strokeColor = colors.red
        chart.lines[0].strokeWidth = 2
        chart.lines.symbol = makeMarker('Square')
        chart.lineLabelFormat = '%2.0f'
        chart.categoryAxis.joinAxisMode = 'bottom'
        chart.categoryAxis.labels.fontName = 'FreeSans'
        chart.categoryAxis.labels.angle = 45
        chart.categoryAxis.labels.boxAnchor = 'e'
        chart.categoryAxis.categoryNames = [str(x.date) for x in weather_history]
        chart.valueAxis.labelTextFormat = '%2.0f '
        chart.valueAxis.valueStep = 10

        # chart needs to be put in a drawing
        d = Drawing(0, 170)
        d.add(chart)
        # add drawing to data
        data.append(d)


        doc.build(data)
        pdf = self.buffer.getvalue()
        self.buffer.close()
        return pdf
