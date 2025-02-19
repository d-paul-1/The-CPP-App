import customtkinter as ctk
from tkinter import filedialog, messagebox  # Import filedialog to handle file uploads and messagebox for popups
import payroll_processing as pp
import salaries_tab_generation as stg
import pandas as pd
import os
import requests
import openpyxl
import subprocess
import shutil

class PayrollSpreadsheet(ctk.CTkFrame):


    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Variables to store main values
        self.has_unsaved_changes = False
        self.master_data_input_sheet_path = None  # Store the path for the master data input sheet
        self.pay_schedules_df = pd.DataFrame()  # Store the path for Year 1 of the pay schedule
        self.pay_periods = 0
        self.options_state =None # Variable that tracks which option from the option menu was selected

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

        # Master Data Input sheet download button
        master_data_input_sheet_download_button = ctk.CTkButton(header_frame, text="Download Master Data Input Sheet" , command=self.download_master_data_input_sheet)
        master_data_input_sheet_download_button.pack(side=ctk.RIGHT, padx=10, pady=10)
    
    def confirm_exit(self):
        """ This function shows confirmation dialog before navigating back."""
        if self.has_unsaved_changes:
            response = messagebox.askyesno("Unsaved Changes", "Any unsaved progress will be lost. Do you want to continue?")
            if response:  # User chose to continue
                self.reset_and_back()
        else:
            self.reset_and_back()

    def download_master_data_input_sheet(self):
        """Handles the download of the master data input sheet from the templates directory."""
        try:
            # Get the app's base directory
            app_base_dir = os.path.dirname(os.path.abspath(__file__))
            template_file_path = os.path.join(app_base_dir, "../templates", "master_data_input_sheet.xlsx")

            # Check if the template file exists
            if not os.path.exists(template_file_path):
                self.update_status( "The Master Data Input Sheet template is missing.","red")
                return

            # Prompt the user for a save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Master Data Input Sheet As"
            )
            if not file_path:
                return  # User canceled the operation

            # Copy the template file to the selected location
            shutil.copy(template_file_path, file_path)

            # Notify the user
            self.update_status(f"Download Complete, File saved successfully at {file_path}", "Green")

        except Exception as e:
            # Handle errors and notify the user
            messagebox.showerror("Error", f"An error occurred while downloading the file:\n{e}")



    def generate_input_frame(self):
        """This Function Creates the Input Frame along with its relevant widgets
        """

        # Input Frame
        input_frame = ctk.CTkFrame(self, width=300, height=2000)
        input_frame.pack(side=ctk.LEFT, padx=20, pady=20)
        input_frame.pack_propagate(False)


        # Master Data Input sheet
        input_label = ctk.CTkLabel(input_frame, text="Master Data Input Sheet Upload:")
        input_label.pack(pady=10)

        #Blank Label to Display the Master Data Input Sheet path
        self.master_data_input_sheet_label = ctk.CTkLabel(input_frame, text="")
        self.master_data_input_sheet_label.pack(pady=10)

        # Button to Upload Master Data Input Sheet
        self.master_data_input_sheet_button = ctk.CTkButton(input_frame,text="Upload master Data Input Sheet", command=self.upload_master_data_input)
        self.master_data_input_sheet_button.pack(pady=10)
        




        # Payschedule Input Label
        input_label = ctk.CTkLabel(input_frame, text="Select Payschedule Load Method:")
        input_label.pack(pady=10)

        # Define a variable for the option menu
        self.optionmenu_var = ctk.StringVar(value="Select an option")

        # Option Menu for Payschedule Load mathods
        optionmenu = ctk.CTkOptionMenu(
            input_frame,
            variable=self.optionmenu_var,
            values=["Select an option", "Automatic", "URL", "Upload"],
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

        self.pay_periods_label=ctk.CTkLabel(input_frame, text = "")
        self.pay_periods_label.pack(pady=10)
        self.pay_periods_label.pack_forget()

        self.generate_payroll_spreadsheet_button = ctk.CTkButton(input_frame, text="Generate Payroll Spreadsheet", command= lambda : self.generate_payroll_spreadsheet() )
        self.generate_payroll_spreadsheet_button.pack(pady=10)
        self.generate_payroll_spreadsheet_button.pack_forget()

        # Initialize progress bar
        self.progress_bar = ctk.CTkProgressBar(input_frame)
        self.progress_bar.pack(side=ctk.BOTTOM, fill=ctk.X, padx=10, pady=10)
        self.progress_bar.set(0)  # Start at 0%





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

        if choice == "Select an option":
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


        elif choice == "Upload":
            self.option_menu_selection_reset()
            self.has_unsaved_changes = True
            self.options_state=3

            self.year1_pdf = self.upload_file("Select a pay schedule for Year 1 of FY (Jul-Dec):", self.upload_label_year_1,"pdf") 
            self.year2_pdf = self.upload_file("Select a pay schedule for Year 2 of FY (Jan-Jun):", self.upload_label_year_2 ,"pdf")  # Call the file upload function

            self.download_pay_button.pack(pady=10) 

    def upload_master_data_input(self):
        self.master_data_input_sheet_label.configure("Upload Master Data Input Sheet:")
        self.has_unsaved_changes = True
        self.master_data_input_sheet_path = self.upload_file("Select the Master Data Input Sheet", self.master_data_input_sheet_label,"excel")
        self.load_excel(self.master_data_input_sheet_path)
        self.update_status("Master Data Input Sheet Uploaded Successfully","white")
        self.master_data_input_sheet_button.configure(text="Re-upload Master Data Input Sheet")


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
            self.progress_bar.set(1)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text= f"number of pay periods: {self.pay_periods}")
            self.pay_periods_label.pack(pady=10)
            successfull =True

        if(option_state==2):
            year1 = self.url_entry_year_1.get()
            year2 = self.url_entry_year_2.get()

            self.pay_schedules_df , msg, color =pp.merge_dfs(self.process_url(year1),self.process_url(year2))
            self.update_status(msg, color)
            self.display_dataframe(self.pay_schedules_df)
            self.progress_bar.set(1)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text= f"number of pay periods: {self.pay_periods}")
            self.pay_periods_label.pack(pady=10)
            successfull =True

        if(option_state==3):

            self.pay_schedules_df , msg, color =pp.merge_dfs(self.process_upload(self.year1_pdf),self.process_upload(self.year2_pdf))
            self.update_status(msg, color)
            self.display_dataframe(self.pay_schedules_df)
            self.progress_bar.set(1)
            self.pay_periods=pp.get_payperiod(self.pay_schedules_df)
            self.pay_periods_label.configure(text= f"number of pay periods: {self.pay_periods}")
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
        total_steps=5
        self.progress_bar.set(0)
        try:
            # Construct URL
            url, msg, color = pp.construct_url(year)
            self.update_status(msg, color)

            # Check for error in URL construction
            if color == "red":
                raise Exception("Error constructing URL. Please re-enter the year.")
            self.update_progress(1, total_steps)

            # Download PDF
            pdf, msg, color = pp.download_pdf(url)
            self.update_status(msg, color)

            # Check for error in downloading PDF
            if color == "red":
                raise Exception("Error downloading PDF. Please re-enter the year.")
            self.update_progress(2, total_steps)

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")
            self.update_progress(3, total_steps)

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")
            self.update_progress(4, total_steps)

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
        total_steps=5
        self.progress_bar.set(0)
        try:
            # Download PDF
            pdf, msg, color = pp.download_pdf(url)
            self.update_status(msg, color)

            # Check for error in downloading PDF
            if color == "red":
                raise Exception("Error downloading PDF. Please re-enter the year.")
            self.update_progress(2, total_steps)

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")
            self.update_progress(3, total_steps)

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")
            self.update_progress(4, total_steps)

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
        total_steps=5
        self.progress_bar.set(0)
        try:

            # Convert PDF to Word
            word, msg, color = pp.pdf_to_word(pdf)
            self.update_status(msg, color)

            # Check for error in converting PDF to Word
            if color == "red":
                raise Exception("Error converting PDF to Word. Please re-enter the year.")
            self.update_progress(3, total_steps)

            # Convert Word to DataFrame
            df, msg, color = pp.word_to_df(word)
            self.update_status(msg, color)

            # Check for error in converting Word to DataFrame
            if color == "red":
                raise Exception("Error converting Word to DataFrame. Please re-enter the year.")
            self.update_progress(4, total_steps)

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

        

    def update_progress(self, step, total_steps):
        """_summary_

        Args:
            step (_type_): _description_
            total_steps (_type_): _description_
        """        
        # Update the progress bar
        self.progress_bar.set(step / total_steps)
        self.update()  # Update the UI


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

        # input frame
        self.reset_input_frame()

        # Display Frame
        self.reset_display_frame()

        # Status Frame
        self.reset_status_frame()
    
        self.controller.show_frame("SpreadsheetGenerator")  # Navigate back to the previous frame
    
    def reset_input_frame(self):
        """ This function resets the input frame to its original state
        """

        # Master Data Input Field
        self.master_data_input_sheet_label.configure(text="") # Setting the path label clear
        self.master_data_input_sheet_button.configure(text="Upload Master Data Input Sheet")# Upload button

        # Option Menu Reset
        self.optionmenu_var.set("Select an option")  # Reset option menu
        self.option_menu_selection_reset()

        self.download_pay_button.pack_forget() # hiding the load pay schedule button

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

        #Generate Payroll Spreadsheet button
        self.generate_payroll_spreadsheet_button.pack_forget()

        #Reseting the Progress Bar 
        self.progress_bar.set(0) 


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
            # Prompt the user to select a folder
            save_folder = filedialog.askdirectory(
                title="Select Folder to Save Payroll Spreadsheet"
            )

            if not save_folder:
                # User canceled the selection
                self.update_status("Operation canceled. No folder selected.","White")
                return

            # Generate the payroll spreadsheet
            self.payroll_file_path = stg.generate_salaries_tab(master_data_input_sheet_path=self.master_data_input_sheet_path, 
                                                     save_path=save_folder,
                                                     pay_periods=self.pay_periods)

            # Notify the user
            self.update_status(f"Payroll spreadsheet saved successfully at {self.payroll_file_path}", "Green")

        except Exception as e:
            # Handle any errors
            self.update_status(f"An error occurred while generating the payroll spreadsheet:\n{e}" , "Red")


