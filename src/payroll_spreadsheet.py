import customtkinter as ctk
from tkinter import filedialog, messagebox  # Import filedialog to handle file uploads and messagebox for popups

class PayrollSpreadsheet(ctk.CTkFrame):


    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.has_unsaved_changes = False

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

        # Add content to the input_frame
        input_label = ctk.CTkLabel(input_frame, text="Select Payschedule Load Method:")
        input_label.pack(pady=10)

        # Define a variable for the option menu
        self.optionmenu_var = ctk.StringVar(value="Select an option")

        # Create the option menu
        optionmenu = ctk.CTkOptionMenu(
            input_frame,
            variable=self.optionmenu_var,
            values=["Select an option", "Automatic", "URL", "Upload"],
            command=self.optionmenu_callback
        )
        optionmenu.pack(pady=10)

        # Create entry widget for the 1st half of the year
        self.entry_first_half = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.entry_first_half.pack(pady=10)
        self.entry_first_half.pack_forget()  # Initially hide the entry widget

        # Create entry widget for the 2nd half of the year
        self.entry_second_half = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.entry_second_half.pack(pady=10)
        self.entry_second_half.pack_forget()  # Initially hide the entry widget

        self.year1_label = ctk.CTkLabel(input_frame, text="Year 1 of FY (Jul-Dec):")
        self.year1_label.pack(pady=10)
        self.year1_label.pack_forget()

        self.year2_label = ctk.CTkLabel(input_frame, text="Year 2 of FY (Jan-Jun):")
        self.year2_label.pack(pady=10)
        self.year2_label.pack_forget()


        self.download_pay_button = ctk.CTkButton(input_frame,text="Load Pay Schedules", command=self.download_pay_schedules)
        self.download_pay_button.pack(pady=10)
        self.download_pay_button.pack_forget()


    def optionmenu_callback(self, choice):
        print("Option selected:", choice)

        # Hide both entries by default
        self.entry_first_half.pack_forget()
        self.entry_second_half.pack_forget()

        if choice == "Select an option":
            self.entry_first_half.pack_forget()  # Hide entries
            self.entry_second_half.pack_forget()
            self.year1_label.pack_forget()
            self.year2_label.pack_forget()
            self.download_pay_button.pack_forget()

        elif choice == "Automatic":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.download_pay_button.pack(pady=10)  
            pass  # No need to show any entries

        elif choice == "URL":
            # Show both entries for the two halves of the year with appropriate placeholders
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.entry_first_half.configure(placeholder_text="Enter URL for Year 1 of FY (Jul-Dec):")
            self.entry_first_half.pack(pady=10)

            self.entry_second_half.configure(placeholder_text="Enter URL for Year 2 of FY (Jan-Jun):")
            self.entry_second_half.pack(pady=10)

            self.download_pay_button.pack(pady=10) 


        elif choice == "Upload":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.upload_file("Select a pay schedule for Year 1 of FY (Jul-Dec):", self.year1_label) 
            self.upload_file("Select a pay schedule for Year 2 of FY (Jan-Jun):", self.year2_label)  # Call the file upload function

            self.download_pay_button.pack(pady=10) 


    def download_pay_schedules(self):
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

    def confirm_exit(self):
        """Show confirmation dialog before navigating back."""
        if self.has_unsaved_changes:
            response = messagebox.askyesno("Unsaved Changes", "Any unsaved progress will be lost. Do you want to continue?")
            if response:  # User chose to continue
                self.reset_and_back()
        else:
            self.reset_and_back()

    
    def option_menu_selection_reset(self):
        self.entry_first_half.pack_forget()  # Hide the entry for the first half
        self.entry_first_half.delete(0, 'end')  # Clear the entry text
        self.entry_second_half.pack_forget()  # Hide the entry for the second half
        self.entry_second_half.delete(0, 'end')  # Clear the entry text
        
        self.year1_label.pack_forget()  # Hide the Year 1 label
        self.year1_label.configure(text="Year 1 of FY (Jul-Dec):")  # Reset the label text
        
        self.year2_label.pack_forget()  # Hide the Year 2 label
        self.year2_label.configure(text="Year 2 of FY (Jan-Jun):")  # Reset the label text

        self.download_pay_button.pack_forget()


    def reset_and_back(self):
        # Reset all fields and labels
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
