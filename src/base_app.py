import os
import sys
import customtkinter as ctk
from pages.login_page import LoginPage
from pages.main_menu_page import MainMenu
from pages.financial_menu_page import FinancialMenu
from pages.spreadsheet_generator_page import SpreadsheetGenerator
from pages.payroll_spreadsheet_page import PayrollSpreadsheet

# Set the appearance mode to dark
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

def get_icon_path():
    """Dynamically set the path to the .ico file, compatible with PyInstaller."""
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS  # PyInstaller's temporary directory
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "assets", "cpp_heart_logo.ico")

class BaseApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Initial setup
        self.after(0, lambda: self.state('zoomed'))
        self.title("CPP APP")
        
        # Set the taskbar/dock icon
        try:
            icon_path = get_icon_path()
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Failed to set the icon: {e}")

        # Create a container to hold all pages
        self.container = ctk.CTkFrame(self, corner_radius=0)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # Dictionary to store page frames
        self.frames = {}

        # Add pages to the container
        for PageClass in (LoginPage, MainMenu, FinancialMenu, SpreadsheetGenerator, PayrollSpreadsheet):
            page_name = PageClass.__name__
            frame = PageClass(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Show the Login page at first
        self.show_frame("LoginPage")

        # Bind window resize event
        self.bind("<Configure>", self.on_resize)

    def show_frame(self, page_name):
        """Raise the page with the given name."""
        frame = self.frames[page_name]
        frame.tkraise()

    def on_resize(self, event):
        """Dynamically adjust widget scaling and layout on resize."""
        for frame in self.frames.values():
            if hasattr(frame, 'adjust_layout'):
                frame.adjust_layout(event.width, event.height)


# Base class for pages to inherit responsiveness utilities
class ResponsivePage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

    def adjust_layout(self, width, height):
        """Override this method in child pages to handle responsiveness."""
        pass

if __name__ == "__main__":
    app = BaseApp()
    app.mainloop()
