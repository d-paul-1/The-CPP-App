import customtkinter as ctk
from login_page import LoginPage
from main_menu import MainMenu
from admin_menu import AdminMenu
from spreadsheet_generator import SpreadsheetGenerator
from payroll_spreadsheet import PayrollSpreadsheet

# Set the appearance mode to dark
ctk.set_appearance_mode("dark")  # Options: "light", "dark", "system"
ctk.set_default_color_theme("dark-blue")  # Optional: "blue", "green", "dark-blue"

class BaseApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.after(0, lambda:self.state('zoomed'))
        self.title("CPP APP")
        
        # Create a container to hold all pages
        self.container = ctk.CTkFrame(self, corner_radius=0)
        self.container.pack(fill="both", expand=True)

        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # Dictionary to store page frames
        self.frames = {}

        # Add pages to the container
        for PageClass in (LoginPage, MainMenu, AdminMenu, SpreadsheetGenerator,PayrollSpreadsheet):
            page_name = PageClass.__name__
            frame = PageClass(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Show the Login page at first
        self.show_frame("LoginPage")

    def show_frame(self, page_name):
        """Raise the page with the given name."""
        frame = self.frames[page_name]
        frame.tkraise()

if __name__ == "__main__":
    app = BaseApp()

    app.mainloop()
