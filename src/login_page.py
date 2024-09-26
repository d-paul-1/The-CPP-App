import customtkinter as ctk

class LoginPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = ctk.CTkLabel(self, text="Login Page", font=("Arial", 24))
        label.pack(pady=20)

        self.username_entry = ctk.CTkEntry(self, placeholder_text="Username")
        self.username_entry.pack(pady=10)

        self.password_entry = ctk.CTkEntry(self, placeholder_text="Password", show="*")
        self.password_entry.pack(pady=10)

        login_button = ctk.CTkButton(self, text="Login", command=self.login)
        login_button.pack(pady=10)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        # Simple login logic (replace with real authentication)
        if username == "admin" and password == "password":
            self.controller.show_frame("StartPage")
        else:
            error_label = ctk.CTkLabel(self, text="Invalid credentials. Try again.", text_color="red")
            error_label.pack(pady=10)
