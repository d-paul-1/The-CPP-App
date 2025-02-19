import customtkinter as ctk

class MainMenu(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Main Menu", font=("Arial", 24), fg_color="transparent")
        label.pack(pady=20)

        financial_menu_button = ctk.CTkButton(self, text="Financial Tools", command=lambda: controller.show_frame("FinancialMenu"))
        financial_menu_button.pack(pady=10)

        logout_button = ctk.CTkButton(self, text="Logout", command=lambda: controller.show_frame("LoginPage"))
        logout_button.pack(pady=10)

