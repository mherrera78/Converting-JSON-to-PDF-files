from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing #, _DrawingEditorMixin 
from reportlab.graphics.charts.piecharts import Pie, WedgeLabel
from reportlab.lib.colors import white
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.graphics.shapes import Rect

from datetime import datetime, timedelta
#from James - adjust label positions
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.graphics.charts.legends import Legend
import math 
import textwrap
import numpy as np
import pandas as pd

import json
from datetime import datetime
import sys
#import os

####################  PENDING  #######################
# - Make sure to include bonuses section as a list of attributes from JSON file

####################  FINAL PENDING  #################
JSON_FILE_NAME = "TRS_file.json"
STOCK_HISTORY_FILE = "Stockl-history.xlsx"

JSON_FILE_PATH = "..//statements//data//cleaned//"
OUTPUT_PATH = "Output_v2//"
SOURCE_PATH = "..//statements//data//raw//"

####################  CONSTANTS  #######################

Y_POSITION_NAME = 702    # 762 # 765
X_POSITION_NAME = 565

X_POSITION_LEFT = 30

Y_POSITION = 605  # 668  this is the coordenate for the "COMPENSATION" section for the first TRS. Change the line 289
#Y_POSITION = 750

Y_POSITION_FOOTER_TITLE = 105 # 150
Y_POSITION_FOOTER_TEXT = 110 # 135
Y_POSITION_FOOTER_RECTANGLE = 120 

X_POSITION_FOOTER_TITLE = 30
X_POSITION_FOOTER_TEXT = 30

Y_POSITION_FOOTNOTE = 25
Y_POSITION_QR_IMAGE = 45
X_POSITION_QR_IMAGE = 518  # 522 512

#X_POSITION_FINAL = 560
X_POSITION_COL0 = 30
X_POSITION_COL1 = 65
X_POSITION_COL2 = 320
X_POSITION_COL3 = 500
X_POSITION_COL4 = 565 # 550

""" X_POSITION_COL0 = 225
X_POSITION_COL1 = 240
X_POSITION_COL2 = 420
X_POSITION_COL3 = 520
X_POSITION_COL4 = 578 """

HEX_COMPENSATION = "#00467f"
HEX_BENEFITS = "#0078d2"
HEX_TRAVEL = "#c30019"
HEX_PIE_LABELS = "#415763"
HEX_RECTANGLE_COLOR = "#D0DAE0"  # "#F2F2F2"  ## rectanble
HEX_DOC_TITLE = "#FFFFFF"

HEX_ITEMS = "#3A4D66"
HEX_NAME = "#FFFFFF" #"#00B0F0"

FONT_SIZE_TITLE = 11 # 11
FONT_SIZE_TITLE_TOTAL = 9 # 9 this is the total for each section "Total Annual", "Total"
FONT_SIZE_TITLE_GRAND_TOTAL = 12  #this is the column for the total on the far left
FONT_SIZE_SUBTITLE = 10
FONT_SIZE_ITEM = 9

#Pie Labels
FONT_SIZE_PIE_LABELS = 6.5

#Pie chart size
PIE_WIDTH = 111
PIE_HEIGHT = 111

FONT_SIZE_NAME = 21 # 22
FONT_SIZE_TITLE = 15 # 11 this is the title for sections
FONT_SIZE_JOB_TITLE = 14 # 13
FONT_SIZE_TOTAL_LEFT = 15
FONT_TITLE_SIZE = 38 # 37 36

#Total annual
TOTAL_ANNUAL = 0

FONT_SIZE_FOOTNOTE = 8.2 # 6
FONT_SIZE_FOOTER_TEXT = 9  # 8.8

FONTAA_BOLD = "font-Bold"
FONTAA_REGULAR = "font-Regular"
FONTAA_REGULAR_ITALIC = "font-RegularItalic"

FONTAA_LIGHT = "font-Light"
FONTAA_LIGHT_ITALIC = "font-Italic"

####################  END CONSTANTS  #######################

####################  FONT REGISTRATION  #######################

pdfmetrics.registerFont(TTFont('font-Regular', './fonts/font-Regular.ttf'))
pdfmetrics.registerFont(TTFont('font-Bold', './fonts/font-Bold.ttf'))
pdfmetrics.registerFont(TTFont('font-RegularItalic', './fonts/font-RegularItalic.ttf'))

pdfmetrics.registerFont(TTFont('font-Light', './fonts/font-Light.ttf'))
pdfmetrics.registerFont(TTFont('font-Light-Italic', './fonts/font-Medium Italic.ttf'))


####################  END FONT REGISTRATION  #######################

####################  FUNCTIONS  #######################

#get's yesterday's date.  If yesterday is Saturday, Sunday or Monday, return the last friday's date. 
def get_last_valid_stock_date():
    today = datetime.today()
    yesterday = today - timedelta(days=1)
    
    # If yesterday was Saturday, Sunday or Monday
    if yesterday.weekday() in [5, 6, 0]:
        # Find the most recent Friday
        last_friday = yesterday - timedelta(days=(yesterday.weekday() + 3) % 7)
        return last_friday.strftime('%Y-%m-%d')
    else:
        return yesterday.strftime('%Y-%m-%d')


## Function to retrieve the statement date from AAL-History file
def find_price_by_statement_date(statementDate):

    # Read the excel file and find the date
    df = pd.read_excel(SOURCE_PATH + STOCK_HISTORY_FILE)
    
    df['Date'] = pd.to_datetime(df['Date'], unit='D', origin='1899-12-30')

    result = df[df['Date'] == statementDate]

    if result.shape[0] == 0:
        return -1.0
    else:
        res = result.iloc[0]['Close']
        return res

def split_string(input_string):
    if len(input_string) <= 32:  #28
        return [input_string]
    else:
        words = input_string.split(' ')
        result = []
        current_string = ''
        w = 0
        line = 1
        for word in words:
            
            if (len(current_string) + len(word) + 1 > 32) and line < 2:
                line += 1
                result.append(current_string)
                current_string = word
            else:
                w += 1
                
                if w > 1:
                    current_string += ' ' + word
                else:
                    current_string += word

        result.append(current_string)
        return result[0:]

def draw_wrapped_text(c, text, max_width, x, y, line_height):
    # Wrap the text
    lines = textwrap.wrap(text, max_width)

    # Draw each line
    for i, line in enumerate(lines):
        c.drawString(x, y - i * line_height, line)

def center_percent(label, strPerc):
# get the length of the input string
    n = len(label)
    n2 = len(strPerc)

# divide the length by 2 and round up
    if n > 20:
        m = math.ceil((n // 2) + 5)
    elif n >15:
        m = math.ceil((n // 2) + 1)
    else:
        m = math.ceil((n // 2))

    m2 = math.ceil((n2 // 2))

# create a string of m blank spaces
    spaces = (" " * m) + (" " * m2)
    
    # return the spaces string
    return spaces

def next_line(lineType):
    
    global Y_POSITION

    match lineType:
        case 1: 
            #this is a line after a SECTION TITLE
            Y_POSITION = Y_POSITION - 15 # 23
        
        case 2:
            #this is a line after a SECTION SUBTITLE
            Y_POSITION = Y_POSITION - 10 # 12
        
        case 3:
            #this is a line for second SECTION SUBTITLE
            Y_POSITION = Y_POSITION - 15 # 22

        case 4: 
            #this is a line after an item
            Y_POSITION = Y_POSITION - 12 #13

        case 5: 
            #this is a line after a SECTION TOTAL
            Y_POSITION = Y_POSITION - 25 # 30 35

def insert_blank_space(string):

    # initialize an empty list to store the spaced out letters
    spaced = []
    
    # loop through each letter in the string
    for letter in string:
        # add the letter and a blank space to the list
        spaced.append(letter.upper())
        spaced.append(" ")
    
    # join the list elements into a new string
    new_string = "".join(spaced)
    
    # return the new string
    return new_string

def format_usd(number):

# convert the number to a string with two decimal places
    number = "{:.2f}".format(number)

    # split the number into the integer and fractional parts
    integer, fraction = number.split(".")

    # add commas to the integer part every three digits
    integer = "{:,}".format(int(integer))

    # return the formatted string with a dollar sign and a fraction
    return "$" + integer 

def date_to_string(date):
    
    # convert the date string to a datetime object
    date_obj = datetime.strptime(date, "%Y-%m-%d")
    
    # format the date object to a string with month name, day and year
    date_str = date_obj.strftime("%B %d, %Y")
    
    # return the date string
    return date_str

def reset_Y():
    
    global Y_POSITION
    #Y_POSITION = 650
    Y_POSITION = 605  # 668  change the position here for where the "COMPENSATION" section begins.  Also change line for first TRS: 48

####################  END FUNCTIONS  #######################

####################  JSON FILE DECLARATION  #######################

# Open the json file in read mode

try:
    start_time = datetime.now()
    print(" ")
    print("****************")
    print ("Started at: " + str(start_time.strftime('%I:%M:%S %p')))
    counter = 0

    with open(JSON_FILE_PATH + JSON_FILE_NAME, "r") as f:
        # Load the json data into a variable
        data = json.load(f)

except ValueError:
    print("Error: Unable to open the source file.")
    sys.exit(1)
    

# validate there is at least 1 employee in json file
if len(data) < 1:
    print ("Error: there is no employees on the source file.")
else:

####################  GENERATING THE PDF FILE  #######################
    
    #iterate through all employees in JSON file
    for i in range(len(data)):

        # Create a new Drawing object
        d = Drawing(50, 0)

        #this section centers the percentage on the second line
        strCompTitle = insert_blank_space(data[i]["incomeData"]["totalRewards"]["elements"][0]["label"]).rstrip()
        strBenTitle = insert_blank_space(data[i]["incomeData"]["totalRewards"]["elements"][1]["label"]).rstrip()
        strTravTitle = insert_blank_space(data[i]["incomeData"]["totalRewards"]["elements"][2]["label"]).rstrip()

        # Create a new Canvas object
        strFileName = data[i]["empID"] + "_" + data[i]["empFirstName"] + "_" + data[i]["empLastName"]
        c = canvas.Canvas(OUTPUT_PATH + strFileName +".pdf")
        
        c.setPageSize((612, 792))  ## This sets the page size.  612 x 792 = 8.5" x 11" (conversion rate:  1" = 72 points)

       
        # Create an ImageReader object
        banner = ImageReader("images//statement-banner-1.png")

        qr_Image = ImageReader('images//qr_code.png')
        
        # Add image to Canvas #18 , 620, width = 72 * 7.8
        c.drawImage(banner, 18, 570, width = 72 * 8, mask='auto', preserveAspectRatio=True)
        
        ####################  LEFT SIDE SECTION  #######################

        c.setFillColor(HEX_NAME) 
        c.setFont(FONTAA_BOLD, FONT_SIZE_NAME)
        full_Name = data[i]["empFirstName"] + ' ' + data[i]["empLastName"]

        c.drawRightString(X_POSITION_NAME, Y_POSITION_NAME, full_Name)
        #c.drawString(X_POSITION_LEFT, Y_POSITION_NAME - 20, )

        c.setFillColor(HEX_NAME)
        c.setFont(FONTAA_LIGHT, FONT_SIZE_JOB_TITLE) ## title font style

        title = split_string(data[i]["empJobTitle"])
        y_Pos = Y_POSITION_NAME - 22

        for line in title:
            #print(line)
            c.drawRightString(X_POSITION_NAME, y_Pos , line)
            y_Pos -= 15

        ####################  TOP FOOTER #######################
        
        # QR Image
        c.drawImage( qr_Image, X_POSITION_QR_IMAGE, Y_POSITION_QR_IMAGE, width = 50, mask='auto', preserveAspectRatio=True) # 45

        c.setFillColor(HEX_ITEMS)
        c.setFont(FONTAA_BOLD, FONT_SIZE_FOOTER_TEXT)
        c.drawString(X_POSITION_FOOTER_TITLE, Y_POSITION_FOOTER_TITLE, "About Your Statement")
        
        dateYYYYMMDD = data[i]["statementDates"]["compDataAsOf"]

        dateAsOf = date_to_string(data[i]["statementDates"]["compDataAsOf"])
        travelAsOF = date_to_string(data[i]["statementDates"]["travelDataAsOf"]) 
        
        # DELETED: (travel as of " + travelAsOF + ")
        text_right = "Data is as of " + dateAsOf  + ". The rewards shown are a lookback over the past 12 months with the exception of annualized base salary."

        c.setFont(FONTAA_REGULAR, FONT_SIZE_FOOTER_TEXT)
        # 140
        draw_wrapped_text(c , text_right, 125, X_POSITION_FOOTER_TITLE ,Y_POSITION_FOOTER_TEXT - 17 , 10) # 140 139
        #draw_wrapped_text

        ####################  BOTTOM FOOTER #######################

        c.setFont(FONTAA_BOLD, FONT_SIZE_FOOTER_TEXT)
        c.setFillColor(HEX_ITEMS)

        c.drawString(X_POSITION_FOOTER_TITLE, Y_POSITION_FOOTER_TITLE - 55, "Benefits")
        c.setFont(FONTAA_REGULAR, FONT_SIZE_FOOTER_TEXT)
        text_left= "For additional information regarding your benefits, contact HR."

        # (object=c, string to be used, width of text, X coordinate, Y coordinate, spacing in between text )
        draw_wrapped_text(c, text_left, 140, X_POSITION_FOOTER_TITLE ,Y_POSITION_FOOTER_TEXT -70, 11)

        c.setFont(FONTAA_BOLD, 32)
        c.setFillColor(HEX_DOC_TITLE)
        
        ####################  HEADER #######################

        c.setFillColor(HEX_DOC_TITLE) 
        c.setFont(FONTAA_LIGHT, FONT_TITLE_SIZE) # c.setFont(FONTAA_BOLD, FONT_SIZE_NAME)
        c.drawRightString(X_POSITION_NAME, Y_POSITION_NAME + 25, "2023 Income Statement")

        ####################  RIGHT SIDE SECTION  #######################

        ####################  COMPENSATION  #######################

        c.setFillColor(HexColor(HEX_COMPENSATION))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE)
        c.drawString(X_POSITION_COL0, Y_POSITION, strCompTitle)


        c.setFillColor(HexColor(HEX_COMPENSATION))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_GRAND_TOTAL)  #FONT_SIZE_TITLE_TOTAL
        c.drawRightString(X_POSITION_COL4, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["totalComp"]["actualAmount"]))

        next_line(1)

        c.setFillColor(HexColor(HEX_ITEMS))
        c.setFont(FONTAA_BOLD, FONT_SIZE_SUBTITLE)
        c.drawRightString(X_POSITION_COL2, Y_POSITION, "Target")
        c.drawRightString(X_POSITION_COL3, Y_POSITION, "Actual")

        #next_line(2)
        #c.drawRightString(X_POSITION_COL2, Y_POSITION, "Compensation")
        #c.drawRightString(X_POSITION_COL3, Y_POSITION, "Compensation")


        next_line(3)
        c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
        c.drawString(X_POSITION_COL1, Y_POSITION, data[i]["incomeData"]["compensationData"]["baseComp"]["label"])
        c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["baseComp"]["targetAmount"]))
        c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["baseComp"]["actualAmount"]))
        #calculating the running total annual
        TOTAL_ANNUAL = TOTAL_ANNUAL + data[i]["incomeData"]["compensationData"]["baseComp"]["actualAmount"]

        strNotes = ""
        if data[i]["incomeData"]["compensationData"]["baseComp"]["notes"] is not None:
            strNotes = data[i]["incomeData"]["compensationData"]["baseComp"]["notes"]

        c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
        c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)

        next_line(4)
        c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
        c.drawString(X_POSITION_COL1, Y_POSITION, data[i]["incomeData"]["compensationData"]["sti"]["label"])

        if data[i]["incomeData"]["compensationData"]["sti"]["targetAmount"] is not None:
            c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["sti"]["targetAmount"]))
    
        if data[i]["incomeData"]["compensationData"]["sti"]["actualAmount"] is not None:
            c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["sti"]["actualAmount"]))
            TOTAL_ANNUAL = TOTAL_ANNUAL + data[i]["incomeData"]["compensationData"]["sti"]["actualAmount"]


        strNotes = ""
        if data[i]["incomeData"]["compensationData"]["sti"]["notes"] is not None:
            strNotes = data[i]["incomeData"]["compensationData"]["sti"]["notes"]

        c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
        c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)

        next_line(4)
        c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
        c.drawString(X_POSITION_COL1, Y_POSITION, "")

        strTargetPercent = str(int(float(data[i]["incomeData"]["compensationData"]["sti"]["targetPercent"])*100))
        
        #Capture when an employee does not have STI, therefore his/her actual is null
        if data[i]["incomeData"]["compensationData"]["sti"]["actualPercent"] != None: 
            strActualPercent = str(int(float(data[i]["incomeData"]["compensationData"]["sti"]["actualPercent"])*100))
        else:
            strActualPercent = "0"

        c.drawRightString(X_POSITION_COL2, Y_POSITION, "(" + strTargetPercent + "% of base)")
        if data[i]["incomeData"]["compensationData"]["sti"]["actualPercent"] != None:
            c.drawRightString(X_POSITION_COL3, Y_POSITION, "(" + strActualPercent + "% of base)")

        next_line(4)

        for item in data[i]["incomeData"]["compensationData"]["lti"]:

            if (item["targetAmount"] != None and item["actualAmount"] != None):
                c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
                c.drawString(X_POSITION_COL1, Y_POSITION, item["label"])
                if item["targetAmount"] is not None:
                    c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(item["targetAmount"]))
                
                if item["actualAmount"] is not None:
                    c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(item["actualAmount"]))
                    c.drawRightString(X_POSITION_COL3 + 4, Y_POSITION, "*")
                    TOTAL_ANNUAL = TOTAL_ANNUAL + item["actualAmount"]

                strNotes = ""
                if item["notes"] is not None:
                    strNotes = item["notes"]
                c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
                c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)
                next_line(4)

        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_TOTAL)
        #c.drawString(X_POSITION_COL1, Y_POSITION, data[i]["incomeData"]["compensationData"]["totalComp"]["label"])
        c.drawString(X_POSITION_COL1, Y_POSITION, "Total Annual")
        c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["totalComp"]["targetAmount"]))
        #c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(data[i]["incomeData"]["compensationData"]["totalComp"]["actualAmount"]))
        c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(TOTAL_ANNUAL))

        strNotes = ""
        if data[i]["incomeData"]["compensationData"]["totalComp"]["notes"] is not None:
            strNotes = data[i]["incomeData"]["compensationData"]["totalComp"]["notes"]
        c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
        c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)

        ##### This is the new section, will include Actual vs. Target, One-Time Compensation section with detail of all payments
        next_line(4)
        actualVsTarget = str(int(np.round(((TOTAL_ANNUAL / data[i]["incomeData"]["compensationData"]["totalComp"]["targetAmount"]) * 100))))
        actualVsTarget = actualVsTarget + "%"

        c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
        c.drawString(X_POSITION_COL1, Y_POSITION, "Actual vs. Target")
        c.drawRightString(X_POSITION_COL3, Y_POSITION, actualVsTarget)
        
        next_line(4)

        if (len(data[i]["incomeData"]["compensationData"]["bonus"]) > 0):
            
            next_line(4)
            c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
            c.drawString(X_POSITION_COL1, Y_POSITION, "One-Time Compensation")
        
            next_line(4)

            for item in data[i]["incomeData"]["compensationData"]["bonus"]:
                
                if (item["actualAmount"] != None):
                    lblText = str(item["label"])

                    c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
                    c.drawString(X_POSITION_COL1, Y_POSITION, lblText)
                    
                    if item["actualAmount"] is not None:
                        c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(item["actualAmount"]))
                        if lblText.startswith("LTI"):
                            c.drawRightString(X_POSITION_COL3 + 4, Y_POSITION, "*")

                    strNotes = ""
                    if item["notes"] is not None:
                        strNotes = item["notes"]
                    c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
                    c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)

                next_line(4)

        #next_line(4)

        c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_FOOTER_TEXT)
        
        date = get_last_valid_stock_date()
        dateMMDD = datetime.strptime(date, '%Y-%m-%d')
        
        c.drawRightString(X_POSITION_COL3, Y_POSITION, "* Valued at $" + str(find_price_by_statement_date(date)) + "/share (" + dateMMDD.strftime('%m/%d') +" close)")
        
        next_line(5)

        ####################  BENEFITS  #######################

        c.setFillColor(HexColor(HEX_BENEFITS))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE)
        c.drawString(X_POSITION_COL0, Y_POSITION, strBenTitle)

        c.setFillColor(HexColor(HEX_BENEFITS))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_GRAND_TOTAL)
        c.drawRightString(X_POSITION_COL4, Y_POSITION, format_usd(data[i]["incomeData"]["benefitsData"]["totalBenefits"]["erContribution"]))

        next_line(1)
        c.setFillColor(HexColor(HEX_ITEMS))
        c.setFont(FONTAA_BOLD, FONT_SIZE_SUBTITLE)
        c.drawRightString(X_POSITION_COL2, Y_POSITION, "Employee Contributions")
        c.drawRightString(X_POSITION_COL3, Y_POSITION, "Company Contributions")

        #next_line(2)
        #c.drawRightString(X_POSITION_COL2, Y_POSITION, "Contributions")
        #c.drawRightString(X_POSITION_COL3, Y_POSITION, "Compensation")

        next_line(3)

        for item in data[i]["incomeData"]["benefitsData"]["benefitsDetail"]:

            c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
            
            if item["eeContribution"] is not None or item["erContribution"] is not None:

                c.drawString(X_POSITION_COL1, Y_POSITION, item["label"])

                if item["eeContribution"] is not None:
                    c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(item["eeContribution"]))

                if item["erContribution"] is not None:
                    c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(item["erContribution"]))

                strNotes = ""
                if item["notes"] is not None:
                    strNotes = item["notes"]
                    
                c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
                c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)
            
                next_line(4)

        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_TOTAL)
        c.drawString(X_POSITION_COL1, Y_POSITION, data[i]["incomeData"]["benefitsData"]["totalBenefits"]["label"])
        c.drawRightString(X_POSITION_COL2, Y_POSITION, format_usd(data[i]["incomeData"]["benefitsData"]["totalBenefits"]["eeContribution"]))
        c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(data[i]["incomeData"]["benefitsData"]["totalBenefits"]["erContribution"]))

        ####################  TRAVEL  #######################

        next_line(5)
        c.setFillColor(HexColor(HEX_TRAVEL))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE)
        c.drawString(X_POSITION_COL0, Y_POSITION, strTravTitle)

        c.setFillColor(HexColor(HEX_TRAVEL))
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_GRAND_TOTAL)
        c.drawRightString(X_POSITION_COL4, Y_POSITION, format_usd(data[i]["incomeData"]["travelData"]["totalTravel"]["commercialValue"]))

        next_line(1)
        c.setFillColor(HexColor(HEX_ITEMS))
        c.setFont(FONTAA_BOLD, FONT_SIZE_SUBTITLE)
        c.drawRightString(X_POSITION_COL3, Y_POSITION, "Commercial Value")

        #next_line(2)
        #c.drawRightString(X_POSITION_COL3, Y_POSITION, "Value")

        next_line(3)

        for item in data[i]["incomeData"]["travelData"]["travelDetail"]:
                        
            if item["commercialValue"] is not None:
                c.setFont(FONTAA_REGULAR, FONT_SIZE_ITEM)
                c.drawString(X_POSITION_COL1, Y_POSITION, item["label"])
                c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(item["commercialValue"]))

                strNotes = ""
                if item["notes"] is not None:
                    strNotes = item["notes"]
                c.setFont(FONTAA_REGULAR_ITALIC, FONT_SIZE_ITEM)
                c.drawRightString(X_POSITION_COL4, Y_POSITION, strNotes)
                
                next_line(4)


        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_TOTAL)
        c.drawString(X_POSITION_COL1, Y_POSITION, data[i]["incomeData"]["travelData"]["totalTravel"]["label"])
        c.drawRightString(X_POSITION_COL3, Y_POSITION, format_usd(data[i]["incomeData"]["travelData"]["totalTravel"]["commercialValue"]))

        next_line(4)
        next_line(4)

        c.setFillColor(HEX_ITEMS)
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE)
        c.drawString(X_POSITION_COL0, Y_POSITION , "E A R N I N G S  S T A T E M E N T")
        c.setFont(FONTAA_BOLD, FONT_SIZE_TITLE_GRAND_TOTAL)
        c.drawRightString(X_POSITION_COL4, Y_POSITION , format_usd(data[i]["incomeData"]["totalRewards"]["totalAmount"]))

        #add a SILVER rectangle to devide top and bottom panels
        c.setFillColor(HEX_RECTANGLE_COLOR)
        c.setStrokeColor(HEX_RECTANGLE_COLOR)
        
        #c.rect(X,Y,width,height,fill=1) 
        c.rect(X_POSITION_COL0,Y_POSITION_FOOTER_RECTANGLE,535,5,fill=1)  # 525  this is the silver rectangle

        # Add the Drawing to the Canvas
        renderPDF.draw(d, c, 0, 250)#396

        ###  FOOTNOTE SECTION   ###
        c.setFillColor(HEX_ITEMS)

        c.setFont(FONTAA_REGULAR, FONT_SIZE_FOOTNOTE)
        c.drawRightString(X_POSITION_COL4, Y_POSITION_FOOTNOTE, "Employee ID # " + data[i]["empID"] )

        c.setFillColor(HexColor(HEX_TRAVEL))
        c.setFont(FONTAA_REGULAR, FONT_SIZE_FOOTNOTE)
        c.drawString(X_POSITION_COL0, Y_POSITION_FOOTNOTE, "PERSONAL AND CONFIDENTIAL")

        reset_Y()
        TOTAL_ANNUAL = 0
        
        # Finish up the page and save the file
        c.showPage()
        c.save()
        counter += 1

    end_time = datetime.now()
    print("****************")
    print ("Ended at: " + str(end_time.strftime('%I:%M:%S %p')))
    print("****************")
    print(" ")
    timeDif = end_time - start_time
    print ("Total Time: " + str(int(timeDif.total_seconds())) + " seconds. - " + str(counter) + " statements generated!")
    print(" ")
