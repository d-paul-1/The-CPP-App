import customtkinter as ctk
from login_page import LoginPage
from start_page import StartPage
from page_one import PageOne
from page_two import PageTwo

# Set the appearance mode to dark
ctk.set_appearance_mode("dark")  # Options: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Optional: "blue", "green", "dark-blue"

class BaseApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("600x400")
        self.title("CustomTkinter Multi-Page App - Dark Theme")
        
        # Create a container to hold all pages
        self.container = ctk.CTkFrame(self, corner_radius=0)
        self.container.pack(fill="both", expand=True)

        # Dictionary to store page frames
        self.frames = {}

        # Add pages to the container
        for PageClass in (LoginPage, StartPage, PageOne, PageTwo):
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
