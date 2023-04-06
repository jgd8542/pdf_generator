from genericpath import isfile
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pdf_annotate import PdfAnnotator, Location, Appearance
from core.Models.Attendee import Attendee
from core.Models.Report import Report
from core.DB import DB
from core.Models.Attatchment import Attatchment
from flask import Flask, send_from_directory, url_for
import os
import time
import pikepdf
from pikepdf import Pdf
import fillpdf
from fillpdf import fillpdfs
from glob import glob 
import shutil
import pypdftk
    



class PDF():
    @staticmethod
    def generatePDF(path, report_id, path_out_file):


        try:   
            os.mkdir("static/pdf/downloaded/")
        except:
            print("Downloaded directory exists.")

        try:
            os.mkdir("static/pdf/temp/")
        except:
            print("Temp directory exists.")
            empty = PDF().isEmpty("static/pdf/temp")
            if not empty:
                print("Temp directory is not empty. Removing files....")
                PDF().clearFolder("static/pdf/temp")

        try:
            os.mkdir("static/pdf/temp2/")
        except:
            print("Temp2 directory exists.")
            empty = PDF().isEmpty("static/pdf/temp2")
            if not empty:
                print("Temp2 directory is not empty. Removing files....")
                PDF().clearFolder("static/pdf/temp2")



        if os.path.isfile(path_out_file):
            try:
                os.remove(path_out_file)
            except OSError as e:  # name the Exception `e`
                print("Failed with:", e)  # look what it says
                print("Error code:", e)

        # Gather Data
        multipleEAs = False
        extraNames = False
        i = 0
        data = {}
        # Attendees
        reportAttendees = Attendee().getAllAttendeesByReport_Id(report_id)
        for attendee in reportAttendees:
            if (len(reportAttendees) > 14):
                data['name0'] = "See Attached"
            else:
                data['name' + str(i)] = attendee['name']
            i = i+1
    #     # Projects
        if (i > 13):
            extraNames = True
        j = 0
        k = 0
        reportProjects = Report().getProjectsAssoc(report_id)
        projectStr = ""
        for project in reportProjects:
            projectStr = projectStr + str(project['project_name']) + ','
            k += 1
    #     # remove trailing comma from projectStr
        projectStr = projectStr[:-1]
    #     # check if there is more than one project, if so display 'See Attached' and send projects to new page 'project_eas.pdf'
        if (k > 1):
            attached = projectStr  # use to print all project names on separate page
            projectStr = "See Attached"
            multipleEAs = True

        data['project_number'] = projectStr

        # Other
        otherInfo = Report().getReportByReport_Id(report_id)[0]
        data['topics_discussed'] = otherInfo["topics_discussed"]
        data['suggestions'] = otherInfo['suggestions']
        data['supervisor_comment'] = otherInfo['supervisor_comment']
        # NEEDS REFORMATTING
        data['meeting_date'] = otherInfo['meeting_date']
        data['office_name'] = otherInfo['office_nickname']
        # data['Safe work habits'] = "On"
        # data['SlipTripFall hazards'] = "On"
        # data["Safety glasses"] = "On"

        # Checkboxes
        checkboxes = otherInfo['checkboxes'].split(',')
        if 'Field Related Activities' in checkboxes:
            data['Field Related Activities'] = "Yes"
        if 'Office Related Activities' in checkboxes:
            data['Office Related Activities'] = "Yes"
        if 'Safety Standdown' in checkboxes:
            data['Safety Standdown'] = "Yes"
        if 'Safe work habits' in checkboxes:
            data['Safe work habits'] = "Yes"
        if 'Safe work conditions' in checkboxes:
            data['Safe work conditions'] = "Yes"
        if 'SlipTripFall hazards' in checkboxes:
            data['SlipTripFall hazards'] = "Yes"
        if 'PoliciesProcedures' in checkboxes:
            data['PoliciesProcedures'] = "Yes"
        if 'Close Calls' in checkboxes:
            data['Close Calls'] = "Yes"
        if 'Maintenance Chapter8' in checkboxes:
            data['Maintenance Chapter8'] = "Yes"
        if 'Traffic controlFlagging' in checkboxes:
            data['Traffic controlFlagging'] = "Yes"
        if 'SlipTripFall hazards' in checkboxes:
            data['SlipTripFall hazards'] = "Yes"
        if 'First aid treatment' in checkboxes:
            data['First aid treatment'] = "Yes"
        if 'Respirator safety' in checkboxes:
            data['Respirator safety'] = "Yes"
        if 'Confined spaces' in checkboxes:
            data['Confined spaces'] = "Yes"
        if 'Hard hats' in checkboxes:
            data['Hard hats'] = "Yes"
        if 'Protective vehicles' in checkboxes:
            data['Protective vehicles'] = "Yes"
        if 'Safety vest' in checkboxes:
            data['Safety vest'] = "Yes"
        if 'Body protection' in checkboxes:
            data['Body protection'] = "Yes"
        if 'Foot protection' in checkboxes:
            data['Foot protection'] = "Yes"
        if 'Safety glasses' in checkboxes:
            data['Safety glasses'] = "On"


        # fillpdfs.write_fillable_pdf(path, path_out_file, data,flatten=True)
        pypdftk.fill_form(path, data, path_out_file,flatten=True)

        if multipleEAs:

            # set temp file name for extra project EAs
            eaFileName = "static/pdf/downloaded/project_eas_" + report_id + ".pdf"

            # call the canvas class using 'reportlab' library
            canvas = Canvas(eaFileName)

            # eas page headline
            headline = "Safety Project EAs"

            # set document title
            canvas.setTitle(headline)

            # separate the attached string into an array in order to display line by line
            line = 0
            col = 0
            for project in attached.rsplit(','):
                canvas.drawString(72 + col, (670-line), project)
                line += 20
                if line >= 600:
                    line = 0
                    col += 100

            # draw the headline on the page
            canvas.drawString(72, 710, headline)
            canvas.drawString(72, 700, '_______________')

            # logo for separate page
            logo = 'static/images/caltranslogo.png'

            # draw the logo on the page
            canvas.drawImage(logo, 20, 775, width=50, height=50, mask='auto')

            # assign the date and office name strings
            dateStr = str(data['meeting_date'])
            officeStr = "Office Name: " + str(data['office_name'])

            # draw the date and office name strings
            canvas.drawString(500, 805, dateStr)
            canvas.drawString(225, 805, officeStr)

            # save all that has been drawn to project_eas.pdf
            canvas.save()

            ####END SEPARATE PROJECT EA PAGE CODE####

            ####BEGIN EXTRA NAMES PAGE CODE####

        # Check if names are greater then 15
        if extraNames:

            # set temp file name for extra names file
            nameFileName = "static/pdf/downloaded/names_" + report_id + ".pdf"

            # call the canvas class using 'reportlab' library
            canvas = Canvas(nameFileName)

            # eas page headline
            headline = "Meeting Attendee Names"

            # set document title
            canvas.setTitle(headline)

            # separate the attached string into an array in order to display line by line
            line = 0
            col = 0
            for attendee in reportAttendees:
                canvas.drawString(72 + col, (670-line), attendee['name'])
                line += 20
                if line >= 600:
                    line = 0
                    col += 200

            # draw the headline on the page
            canvas.drawString(72, 710, headline)
            canvas.drawString(72, 700, '_______________')

            # logo for separate page
            logo = 'static/images/caltranslogo.png'

            # draw the logo on the page
            canvas.drawImage(logo, 20, 775, width=50, height=50, mask='auto')

            # assign the date and office name strings
            dateStr = str(data['meeting_date'])
            officeStr = "Office Name: " + str(data['office_name'])

            # draw the date and office name strings
            canvas.drawString(500, 805, dateStr)
            canvas.drawString(225, 805, officeStr)

            # save all that has been drawn to project_eas.pdf
            canvas.save()

            ####END EXTRA NAMES PAGE CODE####


      

        # check if there is multiple EAs or several attendees
        if (multipleEAs or extraNames):

            if multipleEAs:
                
                shutil.copy(eaFileName, "static/pdf/temp")   

            if extraNames:

                shutil.copy(nameFileName, "static/pdf/temp")


      #  shutil.copy(path_out_file, "static/pdf/temp")





   

        # check if there are attachments
        files = {}
        files = Attatchment().getAttatchmentsByReportId(report_id)
        idx = 0
        # if files:
        for i in files:

        #         # assigns the file from its place in the files (attachments) array
            attachedFile = files[idx]["path"]


            shutil.copy(attachedFile, "static/pdf/temp")


            idx += 1


        pdf = Pdf.new()

        #add the meeting report as the first page of the downloadable document
        src = Pdf.open(path_out_file)
        pdf.pages.extend(src.pages)


        #add the additional attachments
        for file in glob('static/pdf/temp/*.pdf'):
            try:
                src = Pdf.open(file)
                pdf.pages.extend(src.pages)
                
                Pdf.close(src) 
                os.remove(file)
            except: 
                print('Error appending file to output PDF.')
        
            
    
        pdf.save('static/pdf/temp2/merged.pdf')




        f = open("static/pdf/temp2/merged.pdf", "rb")
        inputpdf = PdfFileReader(f)

        #split multipage pdfs into individual pages (for writing page numbers)
        for j in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(j))
            if(j<10):
                with open("static/pdf/temp/document-page0%s.pdf" % j, "wb") as outputStream:
                    output.write(outputStream)
            else:
                with open("static/pdf/temp/document-page%s.pdf" % j, "wb") as outputStream:
                    output.write(outputStream)


        totalPages = inputpdf.getNumPages()

        
        pages = 0 
        # # write thepage numbers to each page
        for k in range(totalPages):
            pages += 1
            if(k<10):
                PDF().add_page_num("static/pdf/temp/document-page0%s.pdf" % k, pages, totalPages-1)
            else:
                PDF().add_page_num("static/pdf/temp/document-page%s.pdf" % k, pages, totalPages-1)


        pdf2 = Pdf.new()
        

        for file in glob('static/pdf/temp/*.pdf'):
            try:
                src = Pdf.open(file)
                pdf2.pages.extend(src.pages)
                Pdf.close(src)
                os.remove(file)
            except: 
                print('Error appending file to output PDF.')
            
    
        pdf2.save(path_out_file)



        f.close()
        if os.path.isfile("static/pdf/temp2/merged.pdf"):
            os.remove("static/pdf/temp2/merged.pdf")
        if os.path.isfile("static/pdf/downloaded/names_" + report_id + ".pdf"):
            os.remove("static/pdf/downloaded/names_" + report_id + ".pdf")
        if os.path.isfile("static/pdf/downloaded/project_eas_" + report_id + ".pdf"):
            os.remove("static/pdf/downloaded/project_eas_" + report_id + ".pdf")

        return data




    @ staticmethod
    def add_page_num(page, pageNum, totalPages=None):
        temp = PdfAnnotator(page)
        temp.set_page_dimensions((612, 792), 0)
        if not totalPages:
            temp.add_annotation(
                'text',
                Location(x1=490, y1=1, x2=5000, y2=15, page=0),
                Appearance(fill=[0, 0, 0],
                           stroke_width=1,
                           font_size=8, content='SAFETY MEETING Page ' + str(pageNum),),
            )
        else:
            temp.add_annotation(
                'text',
                Location(x1=490, y1=0, x2=5000, y2=15, page=0),
                Appearance(fill=[0, 0, 0],
                           stroke_width=1,
                           font_size=8, content='SAFETY MEETING Page ' + str(pageNum) + '/' + str(totalPages+1),),
            )
        temp.write(page)



    @ staticmethod
    def get_info(path):
        with open(path, 'rb') as f:
            pdf = PdfFileReader(f)
            info = pdf.getDocumentInfo()
            fields = pdf.getFields()
            number_of_pages = pdf.getNumPages()
            print(fields)
            return number_of_pages

    @ staticmethod
    def get_attachments(id):
        db = DB()
        sql = f"SELECT id, file_name, path from report_attachments_local WHERE report_id = '{id}'"
        result = db.manualQuery(sql)
        return result

    @staticmethod
    def isEmpty(path):
        if os.path.exists(path) and not os.path.isfile(path):
    
            # Checking if the directory is empty or not
            if not os.listdir(path):
                return True
            else:
                return False
        else:
            print("The path is either for a file or not valid")

    @staticmethod
    def clearFolder(path):
        for file in glob(path  + "/*"):
            if os.path.isfile(file):
                os.remove(file)
            elif os.path.isdir(file):
                shutil.rmtree(file)
            else:
                print("Error removing files from: " + path)

    @ staticmethod
    def delete_attachments(path):
        db = DB()
        sql = f"DELETE from report_attachments_local WHERE id = '{path}'"
        result = db.manualQuery(sql)
        return result

