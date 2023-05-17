import argparse
import win32com.client
import pywintypes
import win32api
import re
from collections import defaultdict


parser = argparse.ArgumentParser()
parser.add_argument(
    "-r",
    "--report",
    action="store_true",
    help="Returns the Report Groups and Sheet IDs of the template",
)
args = parser.parse_args()
report = args.report

ppt1 = win32com.client.Dispatch("PowerPoint.Application")
ppt1.Visible = True

ppt2 = win32com.client.Dispatch("PowerPoint.Application")
ppt2.Visible = True

# Open presentation
try:
    temp = ppt1.Presentations.Open(
        r"C:\Users\user\Desktop\Gfk_Ppt\GfK_Home_Appliances_Trends_Miele_JanDec22.pptx",
        ReadOnly=False,
    )
except pywintypes.com_error as err:
    print(err)
    info = err.excepinfo[5]
    errstring = win32api.FormatMessage(info)
    print("Message: ", errstring)


try:
    deck = ppt2.Presentations.Open(
        r"C:\Users\user\Desktop\Gfk_Ppt\GfK_Major_Domestic_Appliances_Trends_ETYT_JanDec22.pptx",
        ReadOnly=False,
    )
except pywintypes.com_error as err:
    print(err)
    info = err.excepinfo[5]
    errstring = win32api.FormatMessage(info)
    print("Message: ", errstring)

if report:
    RG_pattern = r"\bRG\b\s*(\S+)"
    ID_pattern = r"\bID\b\s*(\S+)"
    sheets_dict = defaultdict(set)
    for slide in temp.Slides:
        RG_found = False
        ID_found = False
        for shape in slide.Shapes:
            # Check if shape has text
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    text = text_frame.TextRange.Text
                    # Check if "ID" is present in the text
                    RG_match = re.search(RG_pattern, text)
                    ID_match = re.search(ID_pattern, text)
                    if RG_match:
                        RG_text = RG_match.group(1)
                        RG_found = True
                    if ID_match:
                        ID_text = ID_match.group(1)
                        ID_found = True
            if RG_found and ID_found:
                sheets_dict[RG_text].add(ID_text)

    def export_report(data, filename):
        with open(filename, "w") as file:
            for key, values in data.items():
                file.write(f"Report Group {key}:\n")
                for i, value in enumerate(values, 1):
                    file.write(f"Sheet ID {value}, ")
                    if i % 5 == 0:  # Insert a line break every 5 values
                        file.write("\n")
                file.write("\n")  # Add an additional line break between keys

    # Assuming your dictionary is named 'my_dict' and you want to export to 'report.txt'
    export_report(sheets_dict, "reports.txt")

# Close presentations
temp.Close()
ppt1.Quit()
deck.Close()
ppt2.Quit()
