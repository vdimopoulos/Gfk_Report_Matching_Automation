import os
import win32com.client as win32

def merge_ppts(template_path, deck_path, destination_path):
    # Create COM objects for PowerPoint application and presentation
    ppt_app = win32.gencache.EnsureDispatch("PowerPoint.Application")
    template = ppt_app.Presentations.Open(template_path)
    deck = ppt_app.Presentations.Open(deck_path)

    # Create a new presentation for the merged slides
    merged_ppt = ppt_app.Presentations.Add()

    # Copy slides from the template to the merged presentation
    for slide in template.Slides:
        # Get the ID text box shape from the template slide
        id_textbox = None
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    if "ID " in text_frame.TextRange.Text:
                        id_textbox = shape
                        break

        if id_textbox:
            # Extract the unique ID from the ID text box
            unique_id = id_textbox.TextFrame.TextRange.Text.split("ID ")[1].strip()

            # Find the slide with the matching ID in the deck
            for deck_slide in deck.Slides:
                for shape in deck_slide.Shapes:
                    if shape.HasTextFrame:
                        text_frame = shape.TextFrame
                        if text_frame.HasText:
                            if "ID " in text_frame.TextRange.Text:
                                if unique_id in text_frame.TextRange.Text:
                                    # Copy and paste the matching slide from the deck to the merged presentation
                                    deck_slide.Copy()
                                    merged_ppt.Slides.Paste()
                                    break

    # Copy remaining slides from the template that don't have an ID
    for slide in template.Slides:
        id_textbox = None
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    if "ID " in text_frame.TextRange.Text:
                        id_textbox = shape
                        break

        if not id_textbox:
            # Copy and paste the slide from the template to the merged presentation
            slide.Copy()
            merged_ppt.Slides.Paste()

    # Save and close the merged presentation
    merged_ppt.SaveAs(destination_path)
    merged_ppt.Close()

    # Close the template and deck presentations
    template.Close()
    deck.Close()

    # Close the PowerPoint application
    ppt_app.Quit()

# Usage example
template_path = r"C:\Users\vassilis.dimopoulos\OneDrive - GfK\Desktop\Gfk\VC&TWS\K_Workshop_Infoquest_2022 VC&TWS temp.pptx"
deck_path = r"C:\Users\vassilis.dimopoulos\OneDrive - GfK\Desktop\Gfk\VC&TWS\VC&TWS_Deck.pptx"
destination_path = r"C:\Users\vassilis.dimopoulos\OneDrive - GfK\Desktop\Gfk\VC&TWS\merged.pptx"
merge_ppts(template_path, deck_path, destination_path)