import win32com.client
import pywintypes
import win32api
import re
from collections import defaultdict

class ReportListCreator:
    
    def create_report_list(source):

        ppt1 = win32com.client.Dispatch("PowerPoint.Application")
        ppt1.Visible = True

        # Open presentation
        try:
            temp = ppt1.Presentations.Open(
                source,
                ReadOnly=False,
            )
        except pywintypes.com_error as err:
            print(err)
            info = err.excepinfo[5]
            errstring = win32api.FormatMessage(info)
            print("Message: ", errstring)


        
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
        temp.Close()
        ppt1.Quit()
        output_string = ""

        for key, values in sheets_dict.items():
            output_string += f"Report Group <{key}>:\n"

            for i, value in enumerate(values, 1):
                output_string += f"Sheet <{value}>"

                # Add a comma and line break every 5 values
                if i % 5 == 0:
                    output_string += ",\n"
                else:
                    output_string += ", "

            output_string += "\n\n"

        # Remove the trailing line breaks
        output_string = output_string.rstrip("\n\n")

        return(output_string)

