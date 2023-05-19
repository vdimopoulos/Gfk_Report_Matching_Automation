import customtkinter
import tkinter as tk
from report_list_creator_test import ReportListCreatorTest as rlc


class ReportListWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        icon_path = "gfk.ico"
        self.after(250, lambda: self.iconbitmap(icon_path))
        self.title("Report List Window")
        self.geometry("400x300")
        self.attributes('-topmost', False)
        self.label = customtkinter.CTkLabel(self, text="Report List Window")
        self.label.pack(padx=20, pady=20)
    
        def browse_file_option1():
            file_path = tk.filedialog.askopenfilename(parent=self,filetypes=[("PowerPoint files", "*.pptx")],initialdir="/")
            self.source_entry1.delete(0, tk.END)
            self.source_entry1.insert(0, file_path)


        def process_input():
            source = self.source_entry1.get()
            run_command(source)
            self.attributes('-topmost', False)

        def export_result(result,rs):
        # Open a save dialog to choose the destination file
            file_path = tk.filedialog.asksaveasfilename(parent=rs,
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],initialdir="/"
            )
            if file_path:
            # Save the result to the selected file
                with open(file_path, "w") as file:
                    file.write(result)


        def run_command(input_string):
        # Replace with your command or function that processes the input string
            result = rlc.extract_codes_from_pptx(input_string)
            
         # Create a new window for displaying the result
            icon_path = "gfk.ico"
            result_window = customtkinter.CTkToplevel(self)
            result_window.attributes('-topmost', True)
            result_window.title("Result")
            result_window.after(250, lambda: result_window.iconbitmap(icon_path))
            result_window.geometry("400x300")
            tk_textbox = customtkinter.CTkTextbox(result_window, activate_scrollbars=True)
            tk_textbox.pack(fill=customtkinter.BOTH, expand=True)
            tk_textbox.insert(customtkinter.END, result)
            result_window.close_button = customtkinter.CTkButton(result_window, text="Close",text_color='black',fg_color="orange",border_width=2,border_color="black", command=result_window.destroy)
            result_window.close_button.pack(side=customtkinter.RIGHT, padx=10)
            result_window.export_button = customtkinter.CTkButton(result_window, text="Export List",text_color='black',fg_color="orange",border_width=2,border_color="black",command=lambda: export_result(result,result_window))
            result_window.export_button.pack(side=customtkinter.LEFT, padx=10)
            result_window.attributes('-topmost', True)
            result_window.focus()
            


        self.source_label1 =customtkinter.CTkLabel(self, text="Template File Source:")
        self.source_label1.pack()
        self.source_entry1 = customtkinter.CTkEntry(self, width=350)
        self.source_entry1.pack(pady=20)
        self.browse_button1 = customtkinter.CTkButton(self, text="Browse",text_color='black',fg_color="orange",border_width=2,border_color="black", command=browse_file_option1)
        self.browse_button1.pack(pady=5)
        self.ok_button1 =customtkinter.CTkButton(self, text="OK",text_color='black',fg_color="orange",border_width=2,border_color="black",  command=process_input)
        self.ok_button1.pack(side=customtkinter.LEFT,padx=10)
        self.close_button = customtkinter.CTkButton(self, text="Close",text_color='black',fg_color="orange",border_width=2,border_color="black", command=self.destroy)
        self.close_button.pack(side=customtkinter.RIGHT, padx=10)
    

class PPCreatorWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        icon_path = "gfk.ico"
        self.after(250, lambda: self.iconbitmap(icon_path))
        self.title('Power Point Creator')
        self.geometry("400x500")
        
        self.label = customtkinter.CTkLabel(self, text="Power Point Creator Window")
        self.label.pack(padx=20, pady=20)

        def browse_file_option1():
            file_path = tk.filedialog.askopenfilename(parent=self,filetypes=[("PowerPoint files", "*.pptx")],initialdir="/")
            self.source_entry1.delete(0, tk.END)
            self.source_entry1.insert(0, file_path)

        self.source_label1 =customtkinter.CTkLabel(self, text="Template File Source:")
        self.source_label1.pack()
        self.source_entry1 = customtkinter.CTkEntry(self, width=350)
        self.source_entry1.pack(pady=5)
        self.browse_button1 = customtkinter.CTkButton(self, text="Browse",text_color='black',fg_color="orange",border_width=2,border_color="black", command=browse_file_option1)
        self.browse_button1.pack(pady=10)

        def browse_file_option2():
            file_path = tk.filedialog.askopenfilename(parent=self,filetypes=[("PowerPoint files", "*.pptx")],initialdir="/")
            self.source_entry2.delete(0, tk.END)
            self.source_entry2.insert(0, file_path)

        self.source_label2 =customtkinter.CTkLabel(self, text="Deck File Source:")
        self.source_label2.pack()
        self.source_entry2 = customtkinter.CTkEntry(self, width=350)
        self.source_entry2.pack(pady=5)
        self.browse_button2 = customtkinter.CTkButton(self, text="Browse",text_color='black',fg_color="orange",border_width=2,border_color="black", command=browse_file_option2)
        self.browse_button2.pack(pady=10)

        def browse_file_destination():
            file_path = tk.filedialog.askdirectory(parent=self,initialdir="/")
            self.destination_entry.delete(0, tk.END)
            self.destination_entry.insert(0, file_path)

        self.destination_label =customtkinter.CTkLabel(self, text="Destination File Location")
        self.destination_label.pack()
        self.destination_entry = customtkinter.CTkEntry(self, width=350)
        self.destination_entry.pack(pady=5)
        self.browse_button3 = customtkinter.CTkButton(self, text="Browse",text_color='black',fg_color="orange",border_width=2,border_color="black", command=browse_file_destination)
        self.browse_button3.pack(pady=10)

        self.ok_button1 =customtkinter.CTkButton(self, text="OK",text_color='black',fg_color="orange",border_width=2,border_color="black")
        self.ok_button1.pack(side=customtkinter.LEFT,padx=10)
        
        self.close_button = customtkinter.CTkButton(self, text="Close",text_color='black',fg_color="orange",border_width=2,border_color="black", command=self.destroy)
        self.close_button.pack(side=customtkinter.RIGHT, padx=10)

class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        icon_path = "gfk.ico"
        self.iconbitmap(icon_path)
        self.title("Presentation Manager")
        self.geometry("400x400")
        self.attributes('-topmost', False)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        

        self.button_1 = customtkinter.CTkButton(self, text="Get Template's Reports List",text_color='black',fg_color="orange",border_width=2,border_color="black", command=self.open_report_list_window)
        self.button_1.grid(row=0, column=0, sticky="nsew",padx=(50,50),pady=(90,50))

        self.button_2 = customtkinter.CTkButton(self, text="Create Presentation",text_color='black',fg_color="orange",border_width=2,border_color="black", command=self.open_pp_creator_window)
        self.button_2.grid(row=1, column=0, sticky="nsew",padx=(50,50),pady=(0,90))

        self.exit_button = customtkinter.CTkButton(self, text="Exit",text_color='black',fg_color="orange",border_width=2,border_color="black", command=self.destroy)
        self.exit_button.grid(pady=20)

        self.report_list_window = None
        self.pp_creator_window = None

    def open_report_list_window(self):
        if self.report_list_window is None or not self.report_list_window.winfo_exists():
            self.report_list_window = ReportListWindow(self)  # create window if its None or destroyed
            self.report_list_window.attributes('-topmost', 'true')
        else:
            self.report_list_window.focus()  # if window exists focus it
    
    def open_pp_creator_window(self):
        if self.pp_creator_window is None or not self.pp_creator_window.winfo_exists():
            self.pp_creator_window = PPCreatorWindow(self)  # create window if its None or destroyed
            self.pp_creator_window.attributes('-topmost', 'true')
        else:
            self.pp_creator_window.focus()  # if window exists focus it

app = App()
app.mainloop()