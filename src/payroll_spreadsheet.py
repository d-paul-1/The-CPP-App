import customtkinter as ctk
from tkinter import filedialog, messagebox  # Import filedialog to handle file uploads and messagebox for popups
import payroll_processing as pp
import pandas as pd
import os
import requests

class PayrollSpreadsheet(ctk.CTkFrame):


    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Variables to store main values
        self.has_unsaved_changes = False
        self.master_data_input_sheet_path = None  # Store the path for the master data input sheet
        self.pay_schedules_df = pd.DataFrame()  # Store the path for Year 1 of the pay schedule
        self.pay_periods = 0

        # Variable to dislay and proceesing status messages
        self.msg_label = None
        
        # Variable that tracks which option from the option menu was selected
        self.options_state =None
        

        # Create a header frame
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(side=ctk.TOP, fill="x")

        # Create a label for the frame
        label = ctk.CTkLabel(header_frame, text="Payroll Spreadsheet", font=("Arial", 24))
        label.pack(expand=True, anchor='center')

        # Back button to navigate to the previous frame
        back_button = ctk.CTkButton(header_frame, text="Back to Spreadsheet Generator",
                                     command=self.confirm_exit)
        back_button.pack(side=ctk.RIGHT, padx=10, pady=10)

        # Create a frame to hold inputs
        input_frame = ctk.CTkFrame(self, width=300, height=2000)
        input_frame.pack(side=ctk.LEFT, padx=20, pady=20)
        input_frame.pack_propagate(False)

        #Label and Input for the Master Data Input sheet
        input_label = ctk.CTkLabel(input_frame, text="Master Data Input Sheet Upload:")
        input_label.pack(pady=10)

        #Blank Label to Display the Master Data Input Sheet path
        self.master_data_input_sheet_label = ctk.CTkLabel(input_frame, text="")
        self.master_data_input_sheet_label.pack(pady=10)

        # Button to Upload Master Data Input Sheet
        self.master_data_input_sheet_button = ctk.CTkButton(input_frame,text="Upload master Data Input Sheet", command=self.upload_master_data_input)
        self.master_data_input_sheet_button.pack(pady=10)
    

        # Add content to the input_frame
        input_label = ctk.CTkLabel(input_frame, text="Select Payschedule Load Method:")
        input_label.pack(pady=10)

        # Define a variable for the option menu
        self.optionmenu_var = ctk.StringVar(value="Select an option")

        #   Option Menu for Payschedule Load mathods
        optionmenu = ctk.CTkOptionMenu(
            input_frame,
            variable=self.optionmenu_var,
            values=["Select an option", "Automatic", "URL", "Upload"],
            command=self.optionmenu_callback
        )
        optionmenu.pack(pady=10)

        # Entry Widgets for Automatic Option
        self.automatic_entry_year1 = ctk.CTkEntry(input_frame, placeholder_text="Enter Year 1 of FY (Jul-Dec):", width=300)
        self.automatic_entry_year1.pack(pady=10)
        self.automatic_entry_year1.pack_forget()  # Initially hide the entry widget

        self.automatic_entry_year2 = ctk.CTkEntry(input_frame, placeholder_text="Enter Year 2 of FY (Jan-Jun):", width=300)
        self.automatic_entry_year2.pack(pady=10)
        self.automatic_entry_year2.pack_forget()  # Initially hide the entry widget

        # Create entry widgets for URL in Option Menu
        self.entry_first_half = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.entry_first_half.pack(pady=10)
        self.entry_first_half.pack_forget()  # Initially hide the entry widget

    
        self.entry_second_half = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.entry_second_half.pack(pady=10)
        self.entry_second_half.pack_forget()  # Initially hide the entry widget


        # Labels to display Year 1 and 2 paths for Uplaod in Option Menu
        self.year1_label = ctk.CTkLabel(input_frame, text="Year 1 of FY (Jul-Dec):")
        self.year1_label.pack(pady=10)
        self.year1_label.pack_forget()

        self.year2_label = ctk.CTkLabel(input_frame, text="Year 2 of FY (Jan-Jun):")
        self.year2_label.pack(pady=10)
        self.year2_label.pack_forget()


        # Download Payschedule Button
        self.download_pay_button = ctk.CTkButton(input_frame,text="Load Pay Schedules", command= lambda: self.download_pay_schedules(self.options_state))
        self.download_pay_button.pack(pady=10)
        self.download_pay_button.pack_forget()

        #self.pay_periods=pp.get_payperiod(self.pay_schedules_df)


    def optionmenu_callback(self, choice):
        print("Option selected:", choice)

        # Hide both entries by default
        self.entry_first_half.pack_forget()
        self.entry_second_half.pack_forget()

        if choice == "Select an option":
            self.options_state=0
            self.entry_first_half.pack_forget()  # Hide entries
            self.entry_second_half.pack_forget()
            self.year1_label.pack_forget()
            self.year2_label.pack_forget()
            self.download_pay_button.pack_forget()

        elif choice == "Automatic":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=1

            self.automatic_entry_year1.configure(placeholder_text="Enter Year 1 of FY (Jul-Dec):")
            self.automatic_entry_year1.pack(pady=10)

            self.automatic_entry_year2.configure(placeholder_text="Enter Year 2 of FY (Jan-Jun):")
            self.automatic_entry_year2.pack(pady=10)

            self.download_pay_button.pack(pady=10)  

        elif choice == "URL":
            # Show both entries for the two halves of the year with appropriate placeholders
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=2

            self.entry_first_half.configure(placeholder_text="Enter URL for Year 1 of FY (Jul-Dec):")
            self.entry_first_half.pack(pady=10)

            self.entry_second_half.configure(placeholder_text="Enter URL for Year 2 of FY (Jan-Jun):")
            self.entry_second_half.pack(pady=10)

            

            self.download_pay_button.pack(pady=10) 


        elif choice == "Upload":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=3

            year1_pdf = self.upload_file("Select a pay schedule for Year 1 of FY (Jul-Dec):", self.year1_label) 
            year2_pdf = self.upload_file("Select a pay schedule for Year 2 of FY (Jan-Jun):", self.year2_label)  # Call the file upload function

            self.download_pay_button.pack(pady=10) 

    def upload_master_data_input(self):
        self.master_data_input_sheet_label.configure("Upload Master Data Input Sheet:")
        self.has_unsaved_changes = True
        self.master_data_input_sheet_path = self.upload_file("Select the Master Data Input Sheet", self.master_data_input_sheet_label)
        self.master_data_input_sheet_button.configure(text="Re-upload Master Data Input Sheet")


    def download_pay_schedules(self,option_state):
        if(option_state==1):
            year1 = self.automatic_entry_year1.get()
            year2 = self.automatic_entry_year2.get()

            self.pay_schedules_df=pp.merge_dfs(pp.process_automatic(year1),pp.process_automatic(year2))


        # Logic to handle the download process (e.g., save a file, request from server)
        messagebox.showinfo("Download", "Downloading pay schedules...")    

    def upload_file(self, text, label):
        # Open a file dialog to allow the user to upload a file
        file_path = filedialog.askopenfilename(
            title=text,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:  # If a file is selected
            label.configure(text=file_path)
            label.pack(pady=10)
            return file_path

    def confirm_exit(self):
        """Show confirmation dialog before navigating back."""
        if self.has_unsaved_changes:
            response = messagebox.askyesno("Unsaved Changes", "Any unsaved progress will be lost. Do you want to continue?")
            if response:  # User chose to continue
                self.reset_and_back()
        else:
            self.reset_and_back()

    #Function to reset all the option menu fields to its original state while choosing options
    def option_menu_selection_reset(self):

        # AUTOMATIC OPTION
        self.automatic_entry_year1.pack_forget()  # Hide the entry for the first half
        self.automatic_entry_year1.delete(0, 'end')  # Clear the entry text
        self.automatic_entry_year2.pack_forget()  # Hide the entry for the second half
        self.automatic_entry_year2.delete(0, 'end')  # Clear the entry text


        # URL OPTION
        self.entry_first_half.pack_forget()  # Hide the entry for the first half
        self.entry_first_half.delete(0, 'end')  # Clear the entry text
        self.entry_second_half.pack_forget()  # Hide the entry for the second half
        self.entry_second_half.delete(0, 'end')  # Clear the entry text
        

        # UPLOAD OPTION
        self.year1_label.pack_forget()  # Hide the Year 1 label
        self.year1_label.configure(text="Year 1 of FY (Jul-Dec):")  # Reset the label text
        
        self.year2_label.pack_forget()  # Hide the Year 2 label
        self.year2_label.configure(text="Year 2 of FY (Jan-Jun):")  # Reset the label text

        # Download button
        self.download_pay_button.pack_forget()


    # Reset page to its original state 
    def reset_and_back(self):
        # Reset all fields and labels

        # Master Data Input Field
        self.master_data_input_sheet_label.configure(text="")
        self.master_data_input_sheet_button.configure(text="Upload Master Data Input Sheet")# Upload button

        # Option Menu Field
        self.optionmenu_var.set("Select an option")  # Reset option menu
        self.entry_first_half.pack_forget()  # Hide the entry for the first half
        self.entry_first_half.delete(0, 'end')  # Clear the entry text
        self.entry_second_half.pack_forget()  # Hide the entry for the second half
        self.entry_second_half.delete(0, 'end')  # Clear the entry text
        
        self.year1_label.pack_forget()  # Hide the Year 1 label
        self.year1_label.configure(text="Year 1 of FY (Jul-Dec):")  # Reset the label text
        
        self.year2_label.pack_forget()  # Hide the Year 2 label
        self.year2_label.configure(text="Year 2 of FY (Jan-Jun):")  # Reset the label text
        
        self.controller.show_frame("SpreadsheetGenerator")  # Navigate back to the previous frame

    
    def display_msg(self,error_msg):
        self.msg_label = ctk.CTkLabel(self, text=error_msg, text_color="red")
        self.msg_label.pack(pady=10)

        # Remove the error message after 5 seconds
        self.after(5000, self.clear_error_message)

