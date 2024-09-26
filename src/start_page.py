import customtkinter as ctk

class StartPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Start Page", font=("Arial", 24))
        label.pack(pady=20)

        button1 = ctk.CTkButton(self, text="Go to Page One", command=lambda: controller.show_frame("PageOne"))
        button1.pack(pady=10)

        button2 = ctk.CTkButton(self, text="Go to Page Two", command=lambda: controller.show_frame("PageTwo"))
        button2.pack(pady=10)

        logout_button = ctk.CTkButton(self, text="Logout", command=lambda: controller.show_frame("LoginPage"))
        logout_button.pack(pady=10)
