import customtkinter as ctk
from tkinter import filedialog, messagebox  # Import filedialog to handle file uploads and messagebox for popups

import utility.payschedule_processing as pp
import utility.generate_payroll_spreadsheet as gps
import utility.payroll_input_sheet_conversion as pisc

import pandas as pd
import os
import requests
import openpyxl
import subprocess
import shutil
import sys

class PayrollSpreadsheet(ctk.CTkFrame):


    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Variables to store main values
        self.has_unsaved_changes = False
        self.pay_schedules_df = pd.DataFrame()  # Store the path for Year 1 of the pay schedule
        self.pay_periods = 0
        self.options_state = None # Variable that tracks which option from the option menu was selected

        self.y1 = 0
        self.y2 = 0
    

        # Here are global variables for the important files and folders
        self.main_save_folder = None

        self.input_sheet_one_path = None
        self.input_sheet_two_path = None

        self.payroll_spreadsheet_path = None

        # Making the page
        self.build_page()


    def build_page(self):
        """This Function Builds all the Page frames
        """
        # Generating the Header Frame
        self.generate_header_frame()

        # Generating the Input Frame
        self.generate_input_frame()
        
        # Generating the Display Frame
        self.generate_display_frame()

        # Generating the Status Frame
        self.generate_status_frame()



    def generate_header_frame(self):
        """This Function Creates the Header Frame along with its relevant widgets
        """

        # Header Frame
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(side=ctk.TOP, fill="x")

        # Create a label for the frame
        label = ctk.CTkLabel(header_frame, text="Payroll Spreadsheet", font=("Arial", 24))
        label.pack(expand=True, anchor='center')

        # Back button to navigate to the previous frame
        back_button = ctk.CTkButton(header_frame, text="Back to Spreadsheet Generator", command=self.confirm_exit)
        back_button.pack(side=ctk.RIGHT, padx=10, pady=10)

        # Add to generate_input_frame method
        self.clear_display_button = ctk.CTkButton(
            header_frame,
            text="Clear Output",
            command=self.clear_output_display,
            fg_color="gray"  # Optional: give it a distinct color
        )
        self.clear_display_button.pack(side=ctk.RIGHT, padx=10, pady=10)

        # Button to store save directory
        self.main_save_folder_button = ctk.CTkButton(header_frame, text="1. Choose Save Destination", command= self.set_save_folder )
        self.main_save_folder_button.pack(side=ctk.LEFT, padx=25, pady=10)

        # Data Input sheet 1 Template download button
        data_input_sheet_one_template_download_button = ctk.CTkButton(header_frame, text="2. Download Input Sheet 1" , command=self.download_input_sheet_one)
        data_input_sheet_one_template_download_button.pack(side=ctk.LEFT, padx=10, pady=10)




    def set_save_folder(self):
        self.has_unsaved_changes = True
        self.main_save_folder = filedialog.askdirectory( title="Select Folder to Save Files" )
        self.main_save_folder_button.configure(text="1. Change Save Destination")  # Updating button text
        self.update_status(f"Save Directory set as {self.main_save_folder}" , "white")
        
    def confirm_exit(self):
        """ This function shows confirmation dialog before navigating back."""
        if self.has_unsaved_changes:
            response = messagebox.askyesno("Unsaved Changes", "Any unsaved progress will be lost. Do you want to continue?")
            if response:  # User chose to continue
                self.reset_and_back()
        else:
            self.reset_and_back()

    def enable_editing(self):
        """Enable editing in the text box."""
        self.text_box.configure(state="normal")  # Enable editing
        self.update_status("Editing enabled. Make changes and click 'Save Changes' when done.", "white")

        # Show the Save Changes button
        self.save_button.pack(side=ctk.LEFT, padx=5)

    


    def download_input_sheet_one(self):
        """
        Handles the download of the master data input sheet from the templates directory.
        
        This function checks for the existence of the input sheet template in the application assets,
        allows the user to select a save location, and copies the template to that location.
        """
        try:
            # Determine the base directory depending on the runtime environment
            if getattr(sys, 'frozen', False):  # Running as a PyInstaller bundle
                app_base_dir = sys._MEIPASS
            else:  # Running in a normal Python environment
                app_base_dir = os.path.dirname(os.path.abspath(__file__))

            # Construct the path to the template file
            template_file_path = os.path.join(app_base_dir, "assets", "templates", "input_sheet_one_template.xlsx")

            # Check if the template file exists
            if not os.path.exists(template_file_path):
                self.update_status("The Master Data Input Sheet template is missing.", "red")
                return

            # Determine the save location
            if not self.main_save_folder:
                # Prompt the user for a save location if main_save_folder is not set
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Master Data Input Sheet As"
                )
                if not file_path:
                    return  # User canceled the operation
            else:
                # Use the main_save_folder if set
                file_path = os.path.join(self.main_save_folder, "input_sheet_one_template.xlsx")

            # Copy the template file to the selected location
            shutil.copy(template_file_path, file_path)

            # Notify the user of successful download
            self.update_status(f"Download Complete. File saved successfully at {file_path}", "green")

            os.startfile(file_path)

        except Exception as e:
            # Handle errors and notify the user
            messagebox.showerror("Error", f"An error occurred while downloading the file:\n{e}")
            self.update_status("An error occurred during the download process.", "red")


    def save_changes(self):
        """Save changes made in the text box back to the DataFrame."""
        try:
            # Convert the edited text back into a DataFrame
            edited_df = self.get_dataframe_from_display()
            print(edited_df)


            if edited_df.empty:
                self.update_status("No valid data to save. Please check your edits.", "red")
                return

            # Update the pay_schedules_df variable
            self.pay_schedules_df = edited_df
            self.pay_periods = pp.get_payperiod(self.pay_schedules_df)  # Update the pay periods based on the new DataFrame
            self.pay_periods_label.configure(text=f"Number of pay periods: {self.pay_periods}")
            self.update_status("Changes saved successfully.", "green")

            # Disable editing again
            self.text_box.configure(state="disabled")
            self.save_button.pack_forget()  # Hide the Save Changes button
        except Exception as e:
            self.update_status(f"Failed to save changes: {e}", "red")
    


    def generate_input_frame(self):
        """This Function Creates the Input Frame along with its relevant widgets
        """

        # Create a container frame
        container_frame = ctk.CTkFrame(self, width=300, height=2000)
        container_frame.pack(side=ctk.LEFT, padx=20, pady=20)
        container_frame.pack_propagate(False)

        # Create scrollable frame inside container
        input_frame = ctk.CTkScrollableFrame(
            container_frame,
            width=280,
            height=1900
        )
        input_frame.pack(expand=True, fill="both", padx=5, pady=5)


        # Input sheet Label
        input_label = ctk.CTkLabel(input_frame, text="Input Sheets")
        input_label.pack(pady=10)

        # Blank Label to Display Input Sheet 1 path
        self.input_sheet_one_label = ctk.CTkLabel(input_frame, text="")
        self.input_sheet_one_label.pack(pady=10)

        # Button to Upload Input Sheet 1
        #self.input_sheet_one_button = ctk.CTkButton(input_frame,text="Upload Input Sheet 1", command=lambda: self.upload_input_sheet("input_sheet_one", "Input Sheet 1"))
        self.input_sheet_one_button = ctk.CTkButton(
        input_frame,
        text="3. Upload Input Sheet 1",
        command=lambda: setattr(self, "input_sheet_one_path", self.upload_input_sheet("input_sheet_one", "Input Sheet 1", "3."))
)
        self.input_sheet_one_button.pack(pady=10)

        # Button to Generate Input Sheet 2
        self.generate_input_sheet_two_button = ctk.CTkButton(input_frame,text="4. Generate Input Sheet 2", 
        command=lambda:[ #This lambda function calls the generate input sheet 2 command from payroll_input_sheet_conversion.py and change the button text to re-generate
            os.startfile(pisc.generate_input_sheet_two(save_directory=self.main_save_folder, input_sheet_one_path=self.input_sheet_one_path)),
            self.generate_input_sheet_two_button.configure(text="4. Re-generate Input Sheet 2")
        ] 
        )
        self.generate_input_sheet_two_button.pack(pady=10)

        # Blank Label to Display Input Sheet 2 path
        self.input_sheet_two_label = ctk.CTkLabel(input_frame, text="")
        self.input_sheet_two_label.pack(pady=10)

        # Button to Upload Input Sheet 2
        self.input_sheet_two_button = ctk.CTkButton(
            input_frame,
            text="5. Upload Input Sheet 2", 
            command=lambda: setattr(self, "input_sheet_two_path",self.upload_input_sheet("input_sheet_two", "Input Sheet 2", "5."))
        )
        self.input_sheet_two_button.pack(pady=10)


        # Payschedule Input Label
        input_label = ctk.CTkLabel(input_frame, text="Select Payschedule Load Method:")
        input_label.pack(pady=10)

        # Define a variable for the option menu
        self.optionmenu_var = ctk.StringVar(value="6. Select an option")

        # Option Menu for Payschedule Load mathods
        optionmenu = ctk.CTkOptionMenu(
            input_frame,
            variable=self.optionmenu_var,
            values=["6. Select an option", "Automatic", "URL", "Upload"],
            command=self.optionmenu_callback
        )
        optionmenu.pack(pady=10)

        # Entry Widgets for Automatic Option
        self.automatic_entry_year_1 = ctk.CTkEntry(input_frame, placeholder_text="Enter Year 1 of FY (Jul-Dec):", width=300)
        self.automatic_entry_year_1.pack(pady=10)
        self.automatic_entry_year_1.pack_forget()  # Initially hide the entry widget

        self.automatic_entry_year_2 = ctk.CTkEntry(input_frame, placeholder_text="Enter Year 2 of FY (Jan-Jun):", width=300)
        self.automatic_entry_year_2.pack(pady=10)
        self.automatic_entry_year_2.pack_forget()  # Initially hide the entry widget

        # Create entry widgets for URL in Option Menu
        self.url_entry_year_1 = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.url_entry_year_1.pack(pady=10)
        self.url_entry_year_1.pack_forget()  # Initially hide the entry widget

    
        self.url_entry_year_2 = ctk.CTkEntry(input_frame, placeholder_text="Enter your input here", width=300)
        self.url_entry_year_2.pack(pady=10)
        self.url_entry_year_2.pack_forget()  # Initially hide the entry widget


        # Labels to display Year 1 and 2 paths for Uplaod in Option Menu
        self.upload_label_year_1 = ctk.CTkLabel(input_frame, text="Year 1 of FY (Jul-Dec):")
        self.upload_label_year_1.pack(pady=10)
        self.upload_label_year_1.pack_forget()

        self.upload_label_year_2 = ctk.CTkLabel(input_frame, text="Year 2 of FY (Jan-Jun):")
        self.upload_label_year_2.pack(pady=10)
        self.upload_label_year_2.pack_forget()


        # Download Payschedule Button
        self.download_pay_button = ctk.CTkButton(input_frame,text="Load Pay Schedules", command= lambda: self.download_pay_schedules(self.options_state))
        self.download_pay_button.pack(pady=10)
        self.download_pay_button.pack_forget()

        # Create a frame for the button pair
        self.button_frame = ctk.CTkFrame(input_frame)
        self.button_frame.pack(pady=10)

        # Add Edit Button beside Load Pay Schedules
        self.edit_button = ctk.CTkButton(
            self.button_frame,
            text="Edit",
            command=self.enable_editing
        )
        self.edit_button.pack(side=ctk.LEFT, padx=5)
        self.edit_button.pack_forget()  # Initially hide the button

        # Add Save Changes Button
        self.save_button = ctk.CTkButton(
            self.button_frame,
            text="Save Changes",
            command=self.save_changes,
            fg_color="red"  # Set the button color to red
        )
        self.save_button.pack(side=ctk.LEFT, padx=5)
        self.save_button.pack_forget()  # Initially hide the button

        # Hide the button frame initially
        self.button_frame.pack_forget()



        self.pay_periods_label=ctk.CTkLabel(input_frame, text = "")
        self.pay_periods_label.pack(pady=10)
        self.pay_periods_label.pack_forget()

        self.generate_payroll_spreadsheet_button = ctk.CTkButton(input_frame, text="Generate Payroll Spreadsheet", command= lambda : self.generate_payroll_spreadsheet() )
        self.generate_payroll_spreadsheet_button.pack(pady=10)
        self.generate_payroll_spreadsheet_button.pack_forget()




    def generate_display_frame(self):
        """This Function Creates the Display Frame along with its relevant widgets
        """

        # Display Frame
        display_frame = ctk.CTkFrame(self, width= 1000, height=2000)
        display_frame.pack(side=ctk.LEFT,padx =20, pady=20)
        display_frame.pack_propagate(False)


        # Text-Box widget
        self.text_box_label = ctk.CTkLabel(display_frame,text="Output:")
        self.text_box_label.pack(anchor="nw", pady=10, padx = 10) 

        self.text_box = ctk.CTkTextbox(display_frame,wrap="none", activate_scrollbars=True)  # Wrapping disabled to support horizontal scroll
        self.text_box.pack(expand=True, fill="both", padx=10, pady=10)
        self.text_box.configure(state="disabled") # Set to READ ONLY mode as a default setting

    def generate_status_frame(self):
        """This Function Creates the Status Frame along with its relevant widgets
        """

        # Status Frame
        self.status_frame = ctk.CTkFrame(self, width=400,height=2000)
        self.status_frame.pack(side=ctk.LEFT, padx=20 , pady=20)
        self.status_frame.pack_propagate(False)

        # status frame title
        status_frame_title = ctk.CTkLabel(self.status_frame, text="Status:")
        status_frame_title.pack(anchor = "nw",pady = 10, padx=10)

        # Text Box
        self.status_box = ctk.CTkTextbox(self.status_frame,activate_scrollbars=True,wrap="word")
        self.status_box.pack(expand=True, fill="both", padx=10, pady=10)
        self.status_box.configure(state="disabled") # Set to READ ONLY mode as a default setting



    def update_status(self, message, color):
        """Funtion to add current status to the status box
        
        Args:
            message (_type_): status or error message
            color (_type_): color of the message
        """

        
        # Enable editing temporarily
        self.status_box.configure(state="normal")
        
        # Get the current line number
        self.current_line_number = int(self.status_box.index("insert").split('.')[0])
        
        # Configure the tag with the specified color for the current line
        tag_name = f"f{self.current_line_number}"  # Create a tag name based on the current line number
        self.status_box.tag_config(tag_name, foreground=color)  # Configure the tag for text color
        
        # Insert the message at the end of the textbox with the new tag
        self.status_box.insert("end", message + "\n\n", tag_name)  # Use the tag when inserting text
        
        # Scroll to the end so the latest message is visible
        self.status_box.see("end")
        
        # Set back to read-only mode
        self.status_box.configure(state="disabled")
        



    def optionmenu_callback(self, choice):
        """This Funtion handles the option menu selections

        Args:
            choice (_type_): choice selected
        """
        if choice != "Select an option":
            self.update_status(f"Option selected: {choice}", "white")

        # Hide both entries by default
        self.url_entry_year_1.pack_forget()
        self.url_entry_year_2.pack_forget()

        if choice == "6. Select an option":
            self.options_state=0
            self.option_menu_selection_reset()
            #self.url_entry_year_2.pack_forget()
            #self.upload_label_year_1.pack_forget()
            #self.upload_label_year_2.pack_forget()
            #self.download_pay_button.pack_forget()

        elif choice == "Automatic":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=1

            self.automatic_entry_year_1.configure(placeholder_text="Enter Year 1 of FY (Jul-Dec):")
            self.automatic_entry_year_1.pack(pady=10)

            self.automatic_entry_year_2.configure(placeholder_text="Enter Year 2 of FY (Jan-Jun):")
            self.automatic_entry_year_2.pack(pady=10)

            self.download_pay_button.pack(pady=10)

            self.button_frame.pack(pady=10)
            self.edit_button.pack(side=ctk.LEFT, padx=5)

        elif choice == "URL":
            # Show both entries for the two halves of the year with appropriate placeholders
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=2

            self.url_entry_year_1.configure(placeholder_text="Enter URL for Year 1 of FY (Jul-Dec):")
            self.url_entry_year_1.pack(pady=10)

            self.url_entry_year_2.configure(placeholder_text="Enter URL for Year 2 of FY (Jan-Jun):")
            self.url_entry_year_2.pack(pady=10)

            self.download_pay_button.pack(pady=10) 

            self.button_frame.pack(pady=10)
            self.edit_button.pack(side=ctk.LEFT, padx=5)


        elif choice == "Upload":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=3

            self.year1_pdf = self.upload_file("Select a pay schedule for Year 1 of FY (Jul-Dec):", self.upload_label_year_1,"pdf") 
            self.year2_pdf = self.upload_file("Select a pay schedule for Year 2 of FY (Jan-Jun):", self.upload_label_year_2 ,"pdf")  # Call the file upload function

            self.download_pay_button.pack(pady=10)

            self.button_frame.pack(pady=10) 
            self.edit_button.pack(side=ctk.LEFT, padx=5)

    def upload_input_sheet(self, sheetname, print_text, step_no):
        """
        Uploads an input sheet, dynamically updating associated UI elements.

        This function allows the user to upload an Excel file, validates the inputs, 
        updates the corresponding label and button dynamically, and processes the uploaded file.

        Args:
            sheetname (str): The name of the sheet, used to dynamically reference UI elements 
                            (e.g., "input_sheet_one").
            print_text (str): The descriptive text displayed on the label and button (e.g., "Input Sheet 1").

        Returns:
            str: The file path of the uploaded input sheet, or None if no file was selected.
        """

        # Skip everything if sheetname or print_text is None or empty
        if not sheetname or not print_text:
            return None

        # Retrieve label and button attributes dynamically
        label = getattr(self, f"{sheetname}_label", None)
        button = getattr(self, f"{sheetname}_button", None)

        # Update the label text if label exists
        if label:
            label.configure(text=f"{step_no} Upload {print_text}:")

        # Set the has_unsaved_changes flag to True
        self.has_unsaved_changes = True

        # Trigger file upload and handle empty file paths
        path = self.upload_file(f"Select {print_text}", label, "excel")
        if not path:
            if label:
                label.configure(text="")  # Clear label text if no file was selected
            return None  # Exit the function early

        # Load the Excel file and update the label/button
        self.load_excel(path)
        self.update_status(f"{print_text} Uploaded Successfully", "white")

        if button:
            button.configure(text=f"{step_no} Re-upload {print_text}")

        return path  # Return the file path




    def download_pay_schedules(self,option_state):
        """ Funcrion that calls the respective option menu process funtions and handles entry and par period display

        Args:
            option_state (_type_): option menu selection
        """
        successfull = False # variable to check if each option state task ran without errors

        if(option_state==1):
            year1 = self.automatic_entry_year_1.get()
            year2 = self.automatic_entry_year_2.get()

            self.pay_schedules_df , msg, color =pp.merge_dfs(self.process_automatic(year1),self.process_automatic(year2))
            self.update_status(msg, color)
            self.display_dataframe(self.pay_schedules_df)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text=f"Number of pay periods: {self.pay_periods}")
            self.pay_periods_label.pack(pady=10)
            successfull =True

        if(option_state==2):
            year1 = self.url_entry_year_1.get()
            year2 = self.url_entry_year_2.get()

            self.pay_schedules_df , msg, color =pp.merge_dfs(self.process_url(year1),self.process_url(year2))
            self.update_status(msg, color)
            self.display_dataframe(self.pay_schedules_df)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text=f"Number of pay periods: {self.pay_periods}")
            self.pay_periods_label.pack(pady=10)
            successfull =True

        if(option_state==3):

            self.pay_schedules_df , msg, color =pp.merge_dfs(self.process_upload(self.year1_pdf),self.process_upload(self.year2_pdf))
            self.update_status(msg, color)
            self.display_dataframe(self.pay_schedules_df)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text=f"Number of pay periods: {self.pay_periods}")
            self.pay_periods_label.pack(pady=10)
            successfull =True

        if successfull:
            self.generate_payroll_spreadsheet_button.pack(pady=10)



       

    def upload_file(self, text, label, type):
        """
        Opens a file dialog to allow the user to upload a file and updates the label with the file path.

        Parameters:
        text (str): The title text to display on the file dialog.
        label (tk.Label): The label widget that displays the selected file path.
        type (str): The type of file to filter in the dialog. Options are 'excel', 'pdf', or 'all'.

        Returns:
        str: The path of the selected file, or None if no file is selected.
        """
        # Determine file types based on the `type` parameter
        if type == 'excel':
            filetypes = [("Excel files", "*.xlsx *.xls")]
        elif type == 'pdf':
            filetypes = [("PDF files", "*.pdf")]
        else:  # For 'all' or any other input, allow all file types
            filetypes = [("All files", "*.*")]

        # Open a file dialog to allow the user to upload a file
        file_path = filedialog.askopenfilename(
            title=text,
            filetypes=filetypes
        )

        if file_path:  # If a file is selected
            label.configure(text=file_path)
            label.pack(pady=10)
            return file_path



    def process_automatic(self, year):
        """ Function to return df for the Automatic choice in Option Menu

        Args:
            year (_type_): year of pay schedule

        Returns:
            _type_: Pandas Df of payperiods
        """
        try:
            # Construct URL
            url, msg, color = pp.construct_url(year)
            self.update_status(msg, color)

            # Check for error in URL construction
            if color == "red":
                raise Exception("Error constructing URL. Please re-enter the year.")

            # Download PDF
            pdf, msg, color = pp.download_pdf(url)
            self.update_status(msg, color)

            # Check for error in downloading PDF
            if color == "red":
                raise Exception("Error downloading PDF. Please re-enter the year.")

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")

            return df  # Return the DataFrame if everything was successful

        except Exception as e:
            self.update_status(str(e), "red")  # Display the error message in red
            return None  # Return None to indicate failure
        

    def process_url(self, url):
        """ Function to return df for the URL choice in Option Menu

        Args:
            url (_type_): url of pay schedule

        Returns:
            _type_: Pandas Df of payperiods
        """
        try:
            # Download PDF
            pdf, msg, color = pp.download_pdf(url)
            self.update_status(msg, color)

            # Check for error in downloading PDF
            if color == "red":
                raise Exception("Error downloading PDF. Please re-enter the year.")

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")

            return df  # Return the DataFrame if everything was successful

        except Exception as e:
            self.update_status(str(e), "red")  # Display the error message in red
            return None  # Return None to indicate failure
        
    
    def process_upload(self, pdf):
        """ Function to return df for the URL choice in Option Menu

        Args:
            pdf (_type_): pdf of pay schedule

        Returns:
            _type_: Pandas Df of payperiods
        """
        try:

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")

            return df  # Return the DataFrame if everything was successful

        except Exception as e:
            self.update_status(str(e), "red")  # Display the error message in red
            return None  # Return None to indicate failure
        

    def display_dataframe(self, df):
        """This function displays a pandas data frame onto the output tetxbox

        Args:
            df (_type_): Pandas data frame to be displayed
        """
        if df is not None and not df.empty:
            df_string = df.to_string(index=False)  # Convert DataFrame to string without the index

            # Insert the DataFrame content into the textbox
            self.text_box.configure(state="normal")
            self.text_box.insert("1.0", df_string)  # "1.0" marks the start of the textbox
            # Make textbox read-only
            self.text_box.configure(state="disabled")
        else:
            messagebox.showinfo("DataFrame", "The DataFrame is empty or None.")  # Han

    def get_dataframe_from_display(self):
        """This function converts the text displayed in the output textbox back into a pandas DataFrame.

        Returns:
            pd.DataFrame: A DataFrame reconstructed from the text in the textbox.
        """
        # Enable the textbox temporarily to read its content
        self.text_box.configure(state="normal")
        text_content = self.text_box.get("1.0", "end").strip()
        self.text_box.configure(state="disabled")

        if not text_content:
            messagebox.showinfo("DataFrame", "The textbox is empty. Cannot convert to DataFrame.")
            return pd.DataFrame()

        try:
            # Split the text into lines
            lines = text_content.split("\n")
            
            # The first line contains headers
            headers = ["Pay Period", "Pay Period Dates"]
            
            # Process the data lines
            data = []
            for line in lines[1:]:
                # Find the position of the first date (looking for MM/DD/YYYY pattern)
                date_pos = line.find(next(d for d in line.split() if '/' in d))
                
                # Split line into Pay Period and Pay Period Dates
                pay_period = line[:date_pos].strip()
                pay_period_dates = line[date_pos:].strip()
                
                data.append([pay_period, pay_period_dates])

            # Convert to DataFrame
            df = pd.DataFrame(data, columns=headers)
            return df
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert text to DataFrame:\n{e}")
            return pd.DataFrame()

        


    def load_excel(self, file_path):
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        data = ""
        self.text_box.configure(state="normal")
        for row in sheet.iter_rows(values_only=True):
            data += "\t".join([str(cell) for cell in row]) + "\n"

        self.text_box.insert("1.0", data)
        self.text_box.configure(state="disabled")  # Make it read-only

    def clear_text_box(self, textbox):
        """This function deletes all the text in the CTKTextbox

        Args:
            textbox (_type_): Text box to be cleared
        """
        textbox.configure(state="normal")  # Enable editing if the box is in read-only mode
        textbox.delete("1.0", "end")       # Delete from the start to the end
        textbox.configure(state="disabled") # Optionally set it back to read-only
        



    def reset_and_back(self):
        """ This Funtion resets the page to its original state 
        """
        # Reset all fields and labels

        
        # Header Frame
        self.reset_header_frame()

        # input frame
        self.reset_input_frame()

        # Display Frame
        self.reset_display_frame()

        # Status Frame
        self.reset_status_frame()

        # Reset any global Variables
        self.reset_global_variables()
    
        self.controller.show_frame("SpreadsheetGenerator")  # Navigate back to the previous frame

    def reset_header_frame(self):
        """ This function resets the header frame to its original state
        """
        self.main_save_folder_button.configure(text="1. Choose Save Destination")

    def reset_input_frame(self):
        """ This function resets the input frame to its original state
        """

        # Master Data Input Field
        self.input_sheet_one_label.configure(text="") # Setting the path label clear
        self.input_sheet_one_button.configure(text="3. Upload Input Sheet 1")# Upload button



        self.input_sheet_two_label.configure(text="") # Setting the path label clear
        self.input_sheet_two_button.configure(text="5. Upload Input Sheet 2")# Upload button

        # Resetting the generate input sheet 2 button
        self.generate_input_sheet_two_button.configure(text="4. Generate Input Sheet 2")

        # Option Menu Reset
        self.optionmenu_var.set("6. Select an option")  # Reset option menu
        self.option_menu_selection_reset()

        self.download_pay_button.pack_forget() # hiding the load pay schedule button

        self.button_frame.pack_forget()
        self.edit_button.pack_forget()  # Hide the Edit button
        self.save_button.pack_forget()  # Hide the Save Changes button

    def reset_display_frame(self):
        """ This function resets the input frame to its original state
        """

        # clearing the display box
        self.clear_text_box(self.text_box)
        self.text_box_label.configure(text="Output:")

    def reset_status_frame(self):
        """ This function resets the status frame to its original state
        """

        # Clearing the status box
        self.clear_text_box(self.status_box)

    
    def reset_global_variables(self):
        """ This functions resets any global variables that need to be reset
        """

        self.main_save_folder = None
        self.has_unsaved_changes = False
        self.options_state = None # Variable that tracks which option from the option menu was selected
       



    def option_menu_selection_reset(self):
        """ This function resets all the option menu fields to its original state while choosing options
        """

        # AUTOMATIC OPTION
        self.automatic_entry_year_1.pack_forget()  # Hide the entry for the first half
        self.automatic_entry_year_1.delete(0, 'end')  # Clear the entry text
        self.automatic_entry_year_2.pack_forget()  # Hide the entry for the second half
        self.automatic_entry_year_2.delete(0, 'end')  # Clear the entry text


        # URL OPTION
        self.url_entry_year_1.pack_forget()  # Hide the entry for the first half
        self.url_entry_year_1.delete(0, 'end')  # Clear the entry text
        self.url_entry_year_2.pack_forget()  # Hide the entry for the second half
        self.url_entry_year_2.delete(0, 'end')  # Clear the entry text
        

        # UPLOAD OPTION
        self.upload_label_year_1.pack_forget()  # Hide the Year 1 label
        self.upload_label_year_1.configure(text="Year 1 of FY (Jul-Dec):")  # Reset the label text
        self.upload_label_year_2.pack_forget()  # Hide the Year 2 label
        self.upload_label_year_2.configure(text="Year 2 of FY (Jan-Jun):")  # Reset the label text

        # PAY PERIOD LABEL
        self.pay_periods_label.pack_forget()
        self.pay_periods_label.configure(text = "")

        # Clearing the Output Box
        self.clear_text_box(self.text_box)


        # Download button
        self.download_pay_button.pack_forget()

        self.button_frame.pack_forget()  # Hide the button frame
        self.edit_button.pack_forget()  # Hide the Edit button
        self.save_button.pack_forget()  # Hide the Save Changes button

        #Generate Payroll Spreadsheet button
        self.generate_payroll_spreadsheet_button.pack_forget()
 


    def open_excel(self,file_path):
        """
        Opens the specified Excel file in the default Excel application.

        Parameters:
        file_path (str): The full path of the Excel file to open. This should be a raw string or use double backslashes to avoid path issues on Windows.

        Raises:
        Exception: If the file cannot be opened, an exception message will be printed with the reason.
        """
        try:
            subprocess.Popen(['start', 'excel', file_path], shell=True)
        except Exception as e:
            self.update_status(f"Failed to open Excel file: {e}", "red")


    def generate_payroll_spreadsheet(self):

    
        try:

            if not self.main_save_folder:
            # Prompt the user to select a folder
                self.main_save_folder = filedialog.askdirectory(
                    title="Select Folder to Save Payroll Spreadsheet"
                )

            if not self.main_save_folder:
                # User canceled the selection
                self.update_status("Operation canceled. No folder selected.","White")
                return


            # Generate the payroll spreadsheet
            self.payroll_spreadsheet_path = gps.generate_payroll_spreadsheet(
                save_directory= self.main_save_folder,
                inputsheet_two_path= self.input_sheet_two_path,
                pay_schedule= self.pay_schedules_df,
                num_pay_periods= self.pay_periods
            )
            
            # Notify the user
            self.update_status(f"Payroll spreadsheet saved successfully at {self.payroll_spreadsheet_path}", "Green")

            os.startfile(self.payroll_spreadsheet_path)

        except Exception as e:
            # Handle any errors
            self.update_status(f"An error occurred while generating the payroll spreadsheet:\n{e}" , "Red")

    def clear_output_display(self):
        """
        Clears the output display text box and resets it to its initial state.
        """
        # Enable editing temporarily
        self.text_box.configure(state="normal")
        
        # Clear all content
        self.text_box.delete("1.0", "end")
        
        # Disable editing again
        self.text_box.configure(state="disabled")
        
        # Update status
        self.update_status("Output display cleared", "white")
