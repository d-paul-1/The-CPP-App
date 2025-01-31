import customtkinter as ctk
from PIL import Image  # Required for image handling
import os  # Required for dynamic path handling

class LoginPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Create a frame to hold all login elements and center them with padding
        login_frame = ctk.CTkFrame(self , fg_color="transparent")
        login_frame.pack(expand=True, padx=50, pady=50)  # Padding for the frame

        # Dynamically set the path to the logo image by going one level up
        current_dir = os.path.dirname(os.path.abspath(__file__))  # Get the current directory of the script
        logo_path = os.path.join(current_dir, os.pardir, "assets", "cpp_heart_logo.png")  # Go up one level, then to 'assets/logo.png'

        # Load and display the logo image
        logo_image = ctk.CTkImage(light_image=Image.open(logo_path), size=(150, 150))
        logo_label = ctk.CTkLabel(login_frame, image=logo_image, text="")  # Display logo, remove text
        logo_label.pack(pady=20)

        # Username entry field
        self.username_entry = ctk.CTkEntry(login_frame, placeholder_text="Username")
        self.username_entry.pack(pady=10)
        self.username_entry.bind("<Return>", self.focus_password)  # Bind Enter key to move focus to password entry

        # Password entry field
        self.password_entry = ctk.CTkEntry(login_frame, placeholder_text="Password", show="*")
        self.password_entry.pack(pady=10)
        self.password_entry.bind("<Return>", self.trigger_login)  # Bind Enter key to trigger login button

        # Login button
        self.login_button = ctk.CTkButton(login_frame, text="Login", command=self.login)
        self.login_button.pack(pady=10)

        # Error label placeholder (initially empty)
        self.error_label = None

    def focus_password(self, event):
        """Move focus to password entry when Enter is pressed on username entry."""
        self.password_entry.focus()

    def trigger_login(self, event):
        """Trigger the login action when Enter is pressed on password entry."""
        self.login()

    def clear_error_message(self):
        """Clear the error message after 5 seconds."""
        if self.error_label:
            self.error_label.destroy()
            self.error_label = None

    def login(self):

        
        valid_users = ["admin","admin",""]
        valid_passwords = ["CPP_FY25!","password",""]

        username = self.username_entry.get()
        password = self.password_entry.get()

        # Simple login logic (replace with real authentication)
        if (username,password) in zip (valid_users,valid_passwords):
            self.username_entry.delete(0, 'end')
            self.password_entry.delete(0, 'end')
            self.focus_set() 
            self.controller.show_frame("MainMenu")
        else:
            # Clear username and password fields
            self.username_entry.delete(0, 'end')
            self.password_entry.delete(0, 'end')
            
            # Set focus back to username entry
            self.username_entry.focus()
            
            # Display an error message
            if not self.error_label:
                self.error_label = ctk.CTkLabel(self, text="Invalid credentials. Try again.", text_color="red")
                self.error_label.pack(pady=10)

            # Remove the error message after 5 seconds
            self.after(5000, self.clear_error_message)
