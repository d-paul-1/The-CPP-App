import customtkinter as ctk

class AdminMenu(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Administrative Tools", font=("Arial", 24))
        label.pack(pady=20)

        spreadsheet_generator_button = ctk.CTkButton(self, text="Spreadsheet Generator", command=lambda: controller.show_frame("SpreadsheetGenerator"))
        spreadsheet_generator_button.pack(pady=10)

        back_button = ctk.CTkButton(self, text="Back to Main Menu", command=lambda: controller.show_frame("MainMenu"))
        back_button.pack(pady=10)

