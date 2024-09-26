import customtkinter as ctk

class PageOne(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Page One", font=("Arial", 24))
        label.pack(pady=20)

        back_button = ctk.CTkButton(self, text="Back to Start Page", command=lambda: controller.show_frame("StartPage"))
        back_button.pack(pady=10)
