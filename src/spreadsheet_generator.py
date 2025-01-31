import customtkinter as ctk

class SpreadsheetGenerator(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Financial Spreadsheets Generator", font=("Arial", 24))
        label.pack(pady=20)
        
        payroll_spreadsheet_button = ctk.CTkButton(self, text="Payroll Spreadsheet", command=lambda: controller.show_frame("PayrollSpreadsheet"))
        payroll_spreadsheet_button.pack(pady=10)


        funding_string_button = ctk.CTkButton(self, text="Funding String", command=lambda: controller.show_frame("AdminMenu"))
        funding_string_button.pack(pady=10)


        program_effort_button = ctk.CTkButton(self, text="Program Effort", command=lambda: controller.show_frame("AdminMenu"))
        program_effort_button.pack(pady=10)


        back_button = ctk.CTkButton(self, text="Back to Financial Tools", command=lambda: controller.show_frame("FinancialMenu"))
        back_button.pack(pady=10)
