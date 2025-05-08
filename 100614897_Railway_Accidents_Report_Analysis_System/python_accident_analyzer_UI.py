import openai
import os
import pypdf
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import threading
import time
from PIL import Image, ImageTk
#import io
#import base64

class AccidentAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Railway Accidents Report Analysis System")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Configure the railway theme color scheme
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use 'clam' as base theme
        
        # Define railway-themed colors
        self.primary_color = "#1a3a6e"  # Dark blue (railway uniform color)
        self.secondary_color = "#c93c20"  # Railway signal red
        self.bg_color = "#f7f7f7"  # Light grey (train station color)
        self.accent_color = "#2d5f8b"  # Railway logo blue
        self.track_color = "#484848"  # Railway track color
        self.train_color = "#ffd200"  # Train yellow warning color
        
        # Configure styles
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabelframe', background=self.bg_color)
        self.style.configure('TLabelframe.Label', background=self.bg_color, foreground=self.primary_color, font=('Arial', 10, 'bold'))
        self.style.configure('TLabel', background=self.bg_color, foreground=self.primary_color)
        self.style.configure('TButton', background=self.primary_color, foreground='white', font=('Arial', 9, 'bold'))
        self.style.map('TButton', 
            background=[('active', self.secondary_color), ('disabled', '#cccccc')],
            foreground=[('active', 'white'), ('disabled', '#999999')])
        
        # Special styles for railway elements
        self.style.configure('Signal.TButton', background=self.primary_color)
        self.style.configure('Track.TLabel', background=self.track_color, foreground='white')
        self.style.configure('Station.TFrame', background='#ffffff', relief='raised')
        
        # Loading icon variables
        self.loading = False
        self.loading_thread = None
        
        # Configuration variables
        self.excel_path = None
        self.pdf_directory = r"C:\Users\minia\OneDrive\Desktop\Final_Report_ID" #Path to Directory
        self.api_key = "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" #API Key
        self.logo_path = r"C:\Users\minia\OneDrive\Desktop\Final_Report_ID\Logo\DerbyLogo.jpg"
        self.logo_image = None  # To store the PhotoImage object
        self.logo_label = None
        
        # Data storage
        self.excel_data = None
        self.filtered_data = None
        
        # Railway icon data (base64 encoded)
        # This is a simplified, properly formatted railway icon
        self.railway_icon_data = """
        iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAOxAAADsQBlSsOGwAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAQ5SURBVFiFtZdPbBRlFMB/33e7W9rdthbBtgECqUYTSkhMOBgSezAVoieJiV5UPKCJCR5MjB715EEPHowHD8REY2JivFgCEmMETPyTQGpam+JBqW1Babu7s/s9D7Ptdrf7b3fR9yST+b755r3f+/vefDMLa4RU2i0DdwCHgYPAPmAvMARkVXUQQEQuAVdV9SKwAHwP/CAi58dGvWtPNRPJ3wRKpd0c8CJwHHgWGOgwleIc8DHwmYjMr0kApXLneeBN4JlOA7ZhDviwiHyYzybmOhJAqdzdwBTwSqcBWnBSVd8aG/WutoOYVdrNZJI3gY9YX+UALR8Gp2Lzo3gIQCrtZoH3gatAHrhWqcg/VoitRaYbhf0iExnkMWAh53AbOKEbC/5AcGKBBZAUcJz/8d3X5Vf5JRGJm4H2AW82e9Uk/0JgLNSp0m4BOI3hVq0u15aLomPj3qztDpoBmf2uFxxdOyKLCsyoH+QBfhGRfOIBpdJuAbiY8UPzl4uisRjzyR3B7Ikjfn84F9c6xOlL33rBvtY+Rc6rH+S3t3LIQGUCEHEOXyqUGF1Z1X0j6u0dVi/no3Pvv+p7tXr8jFyV8xRx/HqFoQVfa2DQUO6a+LQmMp1MYlLAwT/m0T1DcXuvZp0dyerVhRj1Q3P8uMHNBWieR+KVGCrXZFdjqgXmAEZaFNOyb1QX/fh0FVhakfuegXqTjIZHrU1AqbSbBw6QENIAXMfZlJPleHYoWUCk5TdPckKdTyRgDjvJm4yUw1gdoiOE5DqcLCAlYMkzBkMSgmMRtrtdnxEk12HiG1wUYUwzQCgWtxPgTvL2nQrwMHr7g7rHkigCYC5JgABLYAhU7Wa6AEIZCNbw0nXNQKfJN+HCk0w+VVokfQFrwQDYwC/eSvGkYXUFLA5GUQqGQvORXZLmewG0m4CxRhN4dK8ARqnmDENOc3/cEWHPcJgr0tLfZi8gSQLS0v9S0x5wFZuA9Y3ZdjiVD/cOaH1pZWnRwYs4xmPM+XQ+ZExTlmOjXqTCLGDPnTcauD2q9SShXvWC0Q2xRrXrTcQawI9LRl9/Qb2+fkXaF3b9kfeMDCQJmMsmfwOnmpVjF4vJb6MZd2/HPw7eHdD6SjHxzHCqkD4pUx8FkSIi1wN7/xgiMb7RaFbHJrznUz7a2iEzluGc63HiC7/a2GfxzGjxGmjSbO/d3xPVdjzlnLlYFN2SKt7oDj/aN4IuB2FwZUm5OB/j9aV8mUYXmmqI1x9bQANTQOsrTw++9oNMZ0QmlUq7jCm/CizWg/DK1aLoZ1/5FWaXGrLbfKERFZJLcDY2PxP/GSuVdrcDrwH70h5Np3m5EeWgcHzZudpMq+tAtGN0g5hV4DTwE/Ar8KcfZOYBSqXdrcBO4CngMPB4OznuA5eBj0TkRLuL/wXwM8IWsAFC7QAAAABJRU5ErkJggg==
        """
        
        # Create the UI
        self.create_menu()
        self.create_layout()
        
        # Try to load default files
        self.load_default_files()
        
        # Show welcome message
        self.root.after(500, self.show_welcome_message)
    
    def create_menu(self):
        """Create the application menu."""
        menu_bar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open Excel File", command=self.browse_excel_file)
        file_menu.add_command(label="Change PDF Directory", command=self.browse_pdf_directory)
        file_menu.add_separator()
        file_menu.add_command(label="Set API Key", command=self.set_api_key)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        
        # Help menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Help", command=self.show_help)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menu_bar)
    
    def create_layout(self):
        """Create the main application layout with railway theme."""
        # Configure root background
        self.root.configure(bg=self.bg_color)
        
        # Create a header with railway theme
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=60)
        header_frame.pack(fill=tk.X)
        
        # Add railway icon with error handling
            # Create a fallback colored box instead of an icon
        icon_label = tk.Label(header_frame, text="🚂", font=("Arial", 20), 
                              fg="white", bg=self.primary_color)
        icon_label.pack(side=tk.LEFT, padx=20)
        
        # Application title
        title_label = tk.Label(header_frame, text="Railway Accidents Report Analysis System", 
                              font=("Arial", 18, "bold"), fg="white", bg=self.primary_color)
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Create railway track decoration
        track_frame = tk.Frame(self.root, bg=self.track_color, height=10)
        track_frame.pack(fill=tk.X)

        # Create a frame for the logo in the top right corner
        self.logo_frame = tk.Frame(header_frame, bg=self.primary_color)
        self.logo_frame.pack(side=tk.RIGHT, padx=20)

        # Create placeholder for logo (empty label)
        self.logo_label = tk.Label(self.logo_frame, bg=self.primary_color)
        self.logo_label.pack(side=tk.RIGHT)
        
        # Add track sleepers
        for i in range(40):
            sleeper = tk.Frame(track_frame, bg="#8a8a8a", width=12, height=10)
            sleeper.pack(side=tk.LEFT, padx=18)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Filter frame (above paned window) styled as a station platform
        filter_frame = ttk.LabelFrame(main_frame, text="Search Options")
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Country filter
        ttk.Label(filter_frame, text="Filter by Country:").pack(side=tk.LEFT, padx=5)
        self.country_var = tk.StringVar()
        self.country_combo = ttk.Combobox(filter_frame, textvariable=self.country_var, width=30)
        self.country_combo.pack(side=tk.LEFT, padx=5)
        self.country_combo.bind("<<ComboboxSelected>>", self.filter_by_country)
        
        # Clear filter button styled as a railway signal
        clear_button = ttk.Button(filter_frame, text="Clear Filter", 
                                 style="Signal.TButton", command=self.clear_filter)
        clear_button.pack(side=tk.LEFT, padx=5)
        
        # Create a paned window to allow resizing sections
        paned_window = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned_window.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Top section - Excel data display (styled as a timetable)
        excel_frame = ttk.LabelFrame(paned_window, text="Accident Reports Database")
        paned_window.add(excel_frame, weight=1)
        
        # Create a tree view to display Excel data
        self.tree = ttk.Treeview(excel_frame)
        self.tree["columns"] = ("report_id", "title", "date", "type", "country")
        
        # Define columns
        self.tree.column("#0", width=0, stretch=tk.NO)  # Hidden column
        self.tree.column("report_id", anchor=tk.W, width=100)
        self.tree.column("title", anchor=tk.W, width=400)
        self.tree.column("date", anchor=tk.W, width=120)
        self.tree.column("type", anchor=tk.W, width=150)
        self.tree.column("country", anchor=tk.W, width=100)
        
        # Define headings
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading("report_id", text="Report ID", anchor=tk.W)
        self.tree.heading("title", text="Incident Title", anchor=tk.W)
        self.tree.heading("date", text="Date", anchor=tk.W)
        self.tree.heading("type", text="Occurrence Type", anchor=tk.W)
        self.tree.heading("country", text="Country", anchor=tk.W)
        
        # Style the treeview to look like a railway timetable
        self.style.configure("Treeview", 
                        background="#ffffff",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="#ffffff")
        self.style.configure("Treeview.Heading", 
                        background=self.primary_color, 
                        foreground="white", 
                        font=('Arial', 9, 'bold'))
        self.style.map('Treeview', background=[('selected', self.secondary_color)])

        self.style.map('Treeview.Heading',
              background=[('active', self.primary_color)],
              foreground=[('active', 'white')])
        
        # Add scrollbars
        excel_scrollbar_y = ttk.Scrollbar(excel_frame, orient=tk.VERTICAL, command=self.tree.yview)
        excel_scrollbar_x = ttk.Scrollbar(excel_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=excel_scrollbar_y.set, xscrollcommand=excel_scrollbar_x.set)
        
        # Grid layout for treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        excel_scrollbar_y.grid(row=0, column=1, sticky="ns")
        excel_scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # Make the treeview expandable
        excel_frame.columnconfigure(0, weight=1)
        excel_frame.rowconfigure(0, weight=1)
        
        # Bind select event
        self.tree.bind("<<TreeviewSelect>>", self.on_accident_selected)
        
        # Bottom section - Analysis area (styled as control panel)
        analysis_frame = ttk.LabelFrame(paned_window, text="Incident Report Analysis")
        paned_window.add(analysis_frame, weight=1)
        
        # Control buttons
        control_frame = ttk.Frame(analysis_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Style as railway control panel with dark background
        control_panel = tk.Frame(control_frame, bg=self.track_color, bd=2, relief=tk.RAISED)
        control_panel.pack(fill=tk.X, pady=5)
        
        # Selected report label with railway-style display
        tk.Label(control_panel, text="Selected Report:", 
                fg="white", bg=self.track_color, 
                font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=10, pady=5)
        
        self.selected_report_label = tk.Label(control_panel, text="None", 
                                            fg=self.train_color, bg=self.track_color,
                                            font=('Arial', 10))
        self.selected_report_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Frame for loading icon and analyze button
        button_frame = tk.Frame(control_panel, bg=self.track_color)
        button_frame.pack(side=tk.RIGHT, padx=10, pady=5)
        
        # Loading icon canvas styled as railway signal
        self.loading_canvas = tk.Canvas(button_frame, width=24, height=24, 
                                      bg=self.track_color, highlightthickness=0)
        self.loading_canvas.pack(side=tk.LEFT, padx=5)
        
        # Draw the loading icon as a railway signal
        self.loading_icon = self.loading_canvas.create_oval(2, 2, 22, 22, 
                                                       fill=self.train_color,
                                                       outline=self.train_color, width=2)
        self.loading_canvas.itemconfigure(self.loading_icon, state="hidden")
        
        # Analyze button styled as emergency signal
        self.analyze_button = tk.Button(button_frame, text="Analyse Report", 
                                      font=('Arial', 10, 'bold'),
                                      bg=self.primary_color, fg="white",
                                      activebackground="#ff5c40",
                                      activeforeground="white",
                                      command=self.analyze_report)
        self.analyze_button.pack(side=tk.LEFT, padx=5)
        self.analyze_button["state"] = "disabled"
        
        # Analysis output
        output_frame = ttk.Frame(analysis_frame)
        output_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create a tabbed interface with railway-themed tabs
        self.style.configure('TNotebook.Tab', background=self.bg_color, 
                           foreground=self.primary_color,
                           padding=[10, 2],
                           font=('Arial', 9, 'bold'))
        self.style.map('TNotebook.Tab', 
                     background=[('selected', self.primary_color)],
                     foreground=[('selected', 'white')])
        
        tab_control = ttk.Notebook(output_frame)
        
        # PDF text tab
        self.pdf_tab = ttk.Frame(tab_control)
        tab_control.add(self.pdf_tab, text="PDF Content")
        
        self.pdf_text = ScrolledText(self.pdf_tab, wrap=tk.WORD, bg="white", 
                                   font=('Consolas', 10))
        self.pdf_text.pack(fill=tk.BOTH, expand=True)
        
        # Analysis result tab
        self.analysis_tab = ttk.Frame(tab_control)
        tab_control.add(self.analysis_tab, text="Analysis Result")
        
        self.analysis_text = ScrolledText(self.analysis_tab, wrap=tk.WORD, bg="white", 
                                        font=('Consolas', 10))
        self.analysis_text.pack(fill=tk.BOTH, expand=True)
        
        # Add the notebook to the frame
        tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Status bar with railway style
        status_frame = tk.Frame(self.root, bg=self.primary_color, height=25)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Ready for analysis")
        status_bar = tk.Label(status_frame, textvariable=self.status_var, 
                            fg="white", bg=self.primary_color, 
                            anchor=tk.W, padx=10, pady=2)
        status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def set_logo_image(self, image_path):
        """Set the logo image in the top right corner."""
        try:
            # Open and resize the image (max height 50px)
            img = Image.open(image_path)
            
            # Calculate new dimensions (maintaining aspect ratio)
            aspect_ratio = img.width / img.height
            new_height = 40  # Fixed height for header
            new_width = int(new_height * aspect_ratio)
            
            img = img.resize((new_width, new_height), Image.LANCZOS)
            
            # Convert to PhotoImage
            photo_img = ImageTk.PhotoImage(img)
            
            # Update or create the logo label
            if self.logo_label:
                self.logo_label.config(image=photo_img, bg=self.primary_color)
            else:
                self.logo_label = tk.Label(self.logo_frame, image=photo_img, bg=self.primary_color)
                self.logo_label.pack(side=tk.RIGHT)
            
            # Keep a reference to prevent garbage collection
            self.logo_image = photo_img
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set logo image: {e}")
    
    def show_welcome_message(self):
        """Display a railway-themed welcome message with instructions when the app starts."""
        welcome = tk.Toplevel(self.root)
        welcome.title("Welcome to the Railway Accidents Report Analysis System")
        welcome.geometry("600x500")
        welcome.transient(self.root)
        welcome.grab_set()
        
        # Make window match railway style
        welcome.configure(bg=self.bg_color)
        
        # Add railway header
        header = tk.Frame(welcome, bg=self.primary_color, height=60)
        header.pack(fill=tk.X)
        
        # Railway icon with error handling
            # Create a fallback text label
        icon_label = tk.Label(header, text="🚂", font=("Arial", 20), 
                              fg="white", bg=self.primary_color)
        icon_label.pack(side=tk.LEFT, padx=20)
        
        # Welcome title
        title_label = tk.Label(header, text="Railway Accidents Report Analysis System", 
                             font=("Arial", 16, "bold"), fg="white", bg=self.primary_color)
        title_label.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Track decoration
        track = tk.Frame(welcome, bg=self.track_color, height=8)
        track.pack(fill=tk.X)
        
        # Instructions
        instructions_frame = ttk.Frame(welcome)
        instructions_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create scrollable text widget for instructions with railway styling
        instructions_text = ScrolledText(instructions_frame, wrap=tk.WORD, 
                                       bg="white", height=15, 
                                       font=("Arial", 10))
        instructions_text.pack(fill=tk.BOTH, expand=True)
        instructions_text.insert(tk.END, """
Welcome to the Railway Accidents Report Analysis System

This application helps you analyse accident reports from the ERAIL database using LLMs.

Basic Usage:
1. Load an Excel file with accident data (File → Open Excel File)
2. Set the directory containing PDF reports (File → Change PDF Directory)
3. Set your OpenAI API Key (File → Set API Key)
4. Select an accident from the list to analyse
5. Click "Analyse Report" to extract accident and decision insights

Features:
• Filter accidents by country using the dropdown at the top
• View full PDF content in the "PDF Content" tab
• View AI analysis in the "Analysis Result" tab
• The Analysis is saved automatically to an "Analysis" folder in a JSON format within the PDF Directory

Tips:
• Make sure your PDF directory contains the report files and ensure you have provided the correct path
• The PDF filename should contain the Final Report ID
• Results are saved as JSON files for future reference
• For large PDFs, the analysis might take longer or fail due to token limits
        """)
        instructions_text.configure(state="disabled")  # Make read-only
        
        # Close button with railway styling
        close_button = tk.Button(welcome, text="Start Journey", 
                               command=welcome.destroy,
                               bg=self.primary_color, fg="white",
                               font=("Arial", 11, "bold"),
                               padx=20, pady=5,
                               activebackground=self.secondary_color,
                               activeforeground="white")
        close_button.pack(pady=20)
        
        # Center the window
        welcome.update_idletasks()
        width = welcome.winfo_width()
        height = welcome.winfo_height()
        x = (welcome.winfo_screenwidth() // 2) - (width // 2)
        y = (welcome.winfo_screenheight() // 2) - (height // 2)
        welcome.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def show_help(self):
        """Show the help dialog with detailed instructions."""
        # Reuse the welcome message function
        self.show_welcome_message()
    
    def load_default_files(self):
        """Try to load default Excel file if available."""
        default_excel = os.path.join(os.getcwd(), "ERAIL DATABASE_1.xlsx")
        if os.path.exists(default_excel):
            self.excel_path = default_excel
            self.load_excel_data()

        # Load default logo
        default_logo =  r"C:\Users\minia\OneDrive\Desktop\Final_Report_ID\Logo\DerbyLogo.jpg"
        if os.path.exists(default_logo):
            self.logo_path = default_logo
            self.set_logo_image(default_logo)
        else:
            print(f"Default logo not found at: {default_logo}")
    
    def browse_excel_file(self):
        """Open a file dialog to select an Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.excel_path = file_path
            self.load_excel_data()
    
    def browse_pdf_directory(self):
        """Open a directory dialog to select the PDF folder."""
        dir_path = filedialog.askdirectory(
            title="Select PDF Directory"
        )
        
        if dir_path:
            self.pdf_directory = dir_path
            self.status_var.set(f"PDF directory set to: {self.pdf_directory}")
    
    def set_api_key(self):
        """Open a dialog to set the OpenAI API key."""
        api_key_window = tk.Toplevel(self.root)
        api_key_window.title("Set OpenAI API Key")
        api_key_window.geometry("400x200")
        api_key_window.resizable(False, False)
        api_key_window.transient(self.root)
        api_key_window.grab_set()
        
        # Style the window with railway theme
        api_key_window.configure(bg=self.bg_color)
        
        # Railway themed header
        header = tk.Frame(api_key_window, bg=self.primary_color, height=40)
        header.pack(fill=tk.X)
        
        tk.Label(header, text="API Key Authentication", 
               font=("Arial", 12, "bold"), fg="white", bg=self.primary_color).pack(pady=8)
        
        # Form area
        form_frame = tk.Frame(api_key_window, bg=self.bg_color, padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(form_frame, text="Enter your OpenAI API Key:", 
               bg=self.bg_color, fg=self.primary_color,
               font=("Arial", 10)).pack(pady=(10, 5))
        
        key_var = tk.StringVar(value=self.api_key)
        key_entry = tk.Entry(form_frame, textvariable=key_var, width=50, show="*",
                           font=("Consolas", 10), bd=2, relief=tk.SUNKEN)
        key_entry.pack(pady=5)
        
        def save_key():
            self.api_key = key_var.get().strip()
            api_key_window.destroy()
            if self.api_key:
                self.status_var.set("API Key has been set")
            else:
                self.status_var.set("API Key is empty")
        
        # Railway styled button
        save_button = tk.Button(form_frame, text="Save Key", 
                              command=save_key,
                              bg=self.secondary_color, fg="white",
                              font=("Arial", 10, "bold"),
                              padx=15, pady=4,
                              activebackground="#ff5c40",
                              activeforeground="white")
        save_button.pack(pady=15)
        
        # Center the window
        api_key_window.update_idletasks()
        width = api_key_window.winfo_width()
        height = api_key_window.winfo_height()
        x = (api_key_window.winfo_screenwidth() // 2) - (width // 2)
        y = (api_key_window.winfo_screenheight() // 2) - (height // 2)
        api_key_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def load_excel_data(self):
        """Load data from the Excel file into the treeview."""
        try:
            self.status_var.set(f"Loading Excel file: {self.excel_path}")
            self.root.update_idletasks()
            
            # Load the Excel file
            self.excel_data = pd.read_excel(self.excel_path)
            self.filtered_data = self.excel_data.copy()  # Create a copy for filtering
            
            # Debug: Print column names to console
            print("Excel columns:", self.excel_data.columns.tolist())
            
            # Populate the country combobox
            countries = ["All Countries"]
            if 'Country' in self.excel_data.columns:
                country_values = self.excel_data['Country'].dropna().unique()
                countries.extend(sorted([str(c) for c in country_values if str(c) != 'nan']))
            
            self.country_combo['values'] = countries
            self.country_combo.current(0)  # Select "All Countries" by default
            
            # Refresh the treeview
            self.refresh_treeview()
            
            self.status_var.set(f"Loaded {len(self.excel_data)} records from Excel file")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
            self.status_var.set("Error loading Excel file")
            print(f"Excel loading error: {e}")
    
    def refresh_treeview(self):
        """Refresh the treeview with current filtered data."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add data to treeview
        for i, row in self.filtered_data.reset_index().iterrows():
            # Get the report ID directly from the Final Report ID column
            report_id = ""
            if 'Final Report ID' in self.filtered_data.columns:
                report_id = str(row['Final Report ID']) if not pd.isna(row['Final Report ID']) else ""
            
            title = str(row.get('Title', '')) if not pd.isna(row.get('Title', '')) else ""
            date = str(row.get('Date of occurrence', '')) if not pd.isna(row.get('Date of occurrence', '')) else ""
            if isinstance(date, str) and len(date) > 10:
                date = date[:10]  # Format date
            occurrence_type = str(row.get('Occurrence type', '')) if not pd.isna(row.get('Occurrence type', '')) else ""
            country = str(row.get('Country', '')) if not pd.isna(row.get('Country', '')) else ""
            
            # Alternate row colors for railway timetable look
            if i % 2 == 0:
                self.tree.insert('', 'end', iid=i, values=(report_id, title, date, occurrence_type, country), tags=('evenrow',))
            else:
                self.tree.insert('', 'end', iid=i, values=(report_id, title, date, occurrence_type, country), tags=('oddrow',))
        
        # Configure row tags for alternating colors (railway timetable style)
        self.tree.tag_configure('evenrow', background='#f0f0f0')
        self.tree.tag_configure('oddrow', background='white')
    
    def filter_by_country(self, event=None):
        """Filter the accident reports by selected country."""
        selected_country = self.country_var.get()
        
        if selected_country == "All Countries":
            self.filtered_data = self.excel_data.copy()
        else:
            # Filter by selected country
            self.filtered_data = self.excel_data[self.excel_data['Country'] == selected_country]
        
        # Refresh the treeview with filtered data
        self.refresh_treeview()
        
        self.status_var.set(f"Filtered: Showing {len(self.filtered_data)} records for {selected_country}")
    
    def clear_filter(self):
        """Clear the country filter and show all records."""
        # Check if any filter is applied
        if self.country_var.get() == "All Countries":
            messagebox.showinfo("No Filter Applied", "No country filter is currently applied.")
            return
            
        self.country_combo.current(0)  # Set to "All Countries"
        self.filtered_data = self.excel_data.copy()
        self.refresh_treeview()
        self.status_var.set(f"Filter cleared: Showing all {len(self.excel_data)} records")
    
    def on_accident_selected(self, event):
        """Handle selection of an accident from the treeview."""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # Get the selected item
        item_id = selected_items[0]
        row_index = int(item_id)
        
        # Get the corresponding row from filtered data
        filtered_index = self.filtered_data.index[row_index]
        selected_row = self.filtered_data.loc[filtered_index]
        
        # Get the report ID directly from the Final Report ID column
        report_id = ""
        if 'Final Report ID' in self.filtered_data.columns:
            report_id = str(selected_row['Final Report ID']) 
            if pd.isna(report_id) or report_id == 'nan':
                report_id = ""
        
        title = str(selected_row.get('Title', ''))
        
        if not report_id:
            self.selected_report_label.config(text="No Report ID available")
            self.analyze_button["state"] = "disabled"
            return
        
        # Display report ID with railway departure board style
        self.selected_report_label.config(text=f"{report_id} - {title[:30]}...")
        self.analyze_button["state"] = "normal"
        
        # Clear previous content
        self.pdf_text.delete(1.0, tk.END)
        self.analysis_text.delete(1.0, tk.END)
        
        # Update status with railway announcement style
        self.status_var.set(f"Report {report_id} selected. Ready for analysis.")
    
    def find_pdf_by_id(self, report_id):
        """Find a PDF file that contains the report ID in its name."""
        if not report_id or report_id == '' or report_id == 'nan':
            return None
            
        report_id_str = str(report_id).strip()
        
        # Search for PDF files in the directory
        if os.path.exists(self.pdf_directory):
            for filename in os.listdir(self.pdf_directory):
                if filename.lower().endswith('.pdf') and report_id_str in filename:
                    return os.path.join(self.pdf_directory, filename)
        
        return None
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from a PDF file."""
        try:
            pdf_reader = pypdf.PdfReader(pdf_path)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text.strip()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text from PDF: {e}")
            return ""
    
    def ask_chatgpt(self, question, context):
        """Send a question along with PDF context to OpenAI's ChatGPT API."""
        if not self.api_key:
            messagebox.showerror("Error", "OpenAI API Key is not set. Please set it in the File menu.")
            return ""
            
        try:
            client = openai.OpenAI(api_key=self.api_key)
            
            self.status_var.set("Processing... Analyzing railway incident...")
            self.root.update_idletasks()
            
            # Check if context is too large (OpenAI has token limits)
            if len(context) > 50000:  # Approximate token limit check
                return "ERROR_TOKEN_LIMIT"
            
            response = client.chat.completions.create(
                model="gpt-4-turbo",
                messages=[
                    {"role": "system", "content": "You are an AI that analyses railway accident reports and returns information in a specific JSON format."},
                    {"role": "user", "content": f"""Given the following document, analyse the railway accident and provide a response in EXACTLY this JSON format:
                     
{{
    "cause_of_accident": "Brief explanation of accident cause in one sentence (max 20 words)",
    "decision_responses": {{
        "contact_to_signallers": [
            "Specific communications and actions taken by signallers, with precise details"
        ],
        "response_and_actions_taken": [
            "Concrete operational decisions, repair activities, and implemented recommendations"
        ]
    }}
}}

Document:
{context}

Question: {question}"""}
                ],
                max_tokens=2048
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to query OpenAI: {e}")
            return ""
    
    def start_loading_animation(self):
        """Start the loading animation with railway signal style."""
        self.loading = True
        self.loading_canvas.itemconfigure(self.loading_icon, state="normal")
        
        def animate():
            colors = [self.train_color, self.secondary_color, "#ffffff", self.secondary_color]
            color_index = 0
            while self.loading:
                color_index = (color_index + 1) % len(colors)
                self.loading_canvas.itemconfig(self.loading_icon, fill=colors[color_index])
                time.sleep(0.5)
                try:
                    self.root.update_idletasks()
                except:
                    break
        
        self.loading_thread = threading.Thread(target=animate)
        self.loading_thread.daemon = True
        self.loading_thread.start()
    
    def stop_loading_animation(self):
        """Stop the loading animation."""
        self.loading = False
        if self.loading_thread:
            if self.loading_thread.is_alive():
                self.loading_thread.join(0.1)
        self.loading_canvas.itemconfigure(self.loading_icon, state="hidden")
    
    def analyze_report(self):
        """Analyze the selected accident report."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select an accident report from the list to analyse.")
            return
                
        item_id = selected_items[0]
        row_index = int(item_id)
        
        # Get the corresponding row from filtered data
        filtered_index = self.filtered_data.index[row_index]
        selected_row = self.filtered_data.loc[filtered_index]
        
        # Get the report ID directly from the Final Report ID column
        report_id = ""
        if 'Final Report ID' in self.filtered_data.columns:
            report_id = str(selected_row['Final Report ID'])
            if pd.isna(report_id) or report_id == 'nan':
                report_id = ""
        
        if not report_id:
            messagebox.showerror("Error", "No Report ID available for this accident")
            return
                
        # Start loading animation
        self.start_loading_animation()
        self.analyze_button["state"] = "disabled"
        
        # Update status with railway announcement style
        self.status_var.set(f"ATTENTION: Analysis of incident {report_id} in progress...")
        
        # Run the analysis in a separate thread
        threading.Thread(target=self._run_analysis, args=(report_id,), daemon=True).start()
    
    def _run_analysis(self, report_id):
        """Run the analysis process in a background thread."""
        try:
            # Find the PDF file
            pdf_path = self.find_pdf_by_id(report_id)
            
            if not pdf_path:
                self.stop_loading_animation()
                messagebox.showerror("Error", f"Could not find PDF file for Report ID: {report_id}")
                self.root.after(0, lambda: self.status_var.set("Failed to find PDF file"))
                self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
                return
                
            self.root.after(0, lambda: self.status_var.set(f"Located report file: {os.path.basename(pdf_path)}"))
            
            # Extract text from PDF
            self.root.after(0, lambda: self.status_var.set("Extracting data from accident report..."))
            
            pdf_text = self.extract_text_from_pdf(pdf_path)
            
            if not pdf_text:
                self.stop_loading_animation()
                self.root.after(0, lambda: self.status_var.set("Failed to extract text from PDF"))
                self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
                return
                
            # Update the PDF text area in the main thread
            self.root.after(0, lambda: self.pdf_text.delete(1.0, tk.END))
            self.root.after(0, lambda: self.pdf_text.insert(tk.END, pdf_text))
            
            # Check if PDF is too large
            if len(pdf_text) > 50000:  # Approximate limit (adjust as needed)
                self.stop_loading_animation()
                messagebox.showwarning("Warning", "This PDF is too large for analysis. Please try another accident report with a smaller PDF file.")
                self.root.after(0, lambda: self.status_var.set("PDF too large for analysis"))
                self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
                return
            
            # Ask question to OpenAI
            question = """Analyze this railway accident report and identify:
            1. The cause of the accident by examining the 'Cause', 'Findings', or 'Investigation' sections.
            2. The decision responses by examining the 'Actions Taken', 'Recommendations', 'Response', or 'Following the Incident' sections."""
            
            self.root.after(0, lambda: self.status_var.set("Analyzing incident causes and responses..."))
            analysis = self.ask_chatgpt(question, pdf_text)
            
            if analysis == "ERROR_TOKEN_LIMIT":
                self.stop_loading_animation()
                messagebox.showwarning("Token Limit Exceeded", 
                                      "This PDF exceeds the token limit for analysis. Please try another accident report with a smaller PDF file.")
                self.root.after(0, lambda: self.status_var.set("Analysis failed - Token limit exceeded"))
                self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
                return
            
            if not analysis:
                self.stop_loading_animation()
                self.root.after(0, lambda: self.status_var.set("Analysis failed"))
                self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
                return
                
            # Update the analysis text area in the main thread
            self.root.after(0, lambda: self.analysis_text.delete(1.0, tk.END))
            self.root.after(0, lambda: self.analysis_text.insert(tk.END, analysis))
            
            # Try to format the JSON for better display
            try:
                json_data = json.loads(analysis)
                formatted_json = json.dumps(json_data, indent=4)
                self.root.after(0, lambda: self.analysis_text.delete(1.0, tk.END))
                self.root.after(0, lambda: self.analysis_text.insert(tk.END, formatted_json))
                
                # Apply syntax highlighting for JSON (simple version)
                self.root.after(0, lambda: self.highlight_json())
            except json.JSONDecodeError:
                # If not valid JSON, keep the original text
                pass
                
            # Save the analysis to a file
            output_dir = os.path.join(os.path.dirname(pdf_path), "Analysis")
            os.makedirs(output_dir, exist_ok=True)
            
            output_filename = f"{report_id}_analysis.json"
            output_path = os.path.join(output_dir, output_filename)
            
            with open(output_path, "w") as f:
                f.write(analysis)
                
            self.root.after(0, lambda: self.status_var.set(f"Analysis complete: Report saved to {output_path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis error: {str(e)}"))
            self.root.after(0, lambda: self.status_var.set(f"Analysis error: {str(e)}"))
        finally:
            # Stop loading animation and re-enable the analyze button
            self.stop_loading_animation()
            self.root.after(0, lambda: self.analyze_button.configure(state="normal"))
    
    def highlight_json(self):
        """Apply simple syntax highlighting to JSON in the analysis text widget."""
        content = self.analysis_text.get(1.0, tk.END)
        self.analysis_text.delete(1.0, tk.END)
        
        # Configure tags for different JSON elements
        self.analysis_text.tag_configure("key", foreground=self.primary_color, font=("Consolas", 10, "bold"))
        self.analysis_text.tag_configure("string", foreground=self.secondary_color)
        self.analysis_text.tag_configure("number", foreground="#008000")
        self.analysis_text.tag_configure("bracket", foreground=self.track_color, font=("Consolas", 10, "bold"))
        
        # Simple line-by-line highlighting
        lines = content.split('\n')
        for line in lines:
            pos = self.analysis_text.index(tk.END)
            self.analysis_text.insert(tk.END, line + '\n')
            
            # Highlight keys
            if ':' in line:
                key_end = line.find(':')
                key_start = line.find('"')
                if key_start >= 0 and key_start < key_end:
                    line_start = pos.split('.')[0]
                    self.analysis_text.tag_add("key", f"{line_start}.{key_start}", f"{line_start}.{key_end}")
            
            # Highlight brackets
            for bracket in ['{', '}', '[', ']']:
                bracket_pos = line.find(bracket)
                if bracket_pos >= 0:
                    line_start = pos.split('.')[0]
                    self.analysis_text.tag_add("bracket", f"{line_start}.{bracket_pos}", f"{line_start}.{bracket_pos+1}")
    
    def show_about(self):
        """Show the about dialog with railway theme."""
        about_window = tk.Toplevel(self.root)
        about_window.title("About Railway Accidents Report Analysis System")
        about_window.geometry("400x300")
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Style the window with railway theme
        about_window.configure(bg=self.bg_color)
        
        # Railway themed header
        header = tk.Frame(about_window, bg=self.primary_color, height=60)
        header.pack(fill=tk.X)
        
        # Railway icon with error handling
            # Create a fallback text label
        icon_label = tk.Label(header, text="🚂", font=("Arial", 20), 
                              fg="white", bg=self.primary_color)
        icon_label.pack(side=tk.LEFT, padx=20, pady=10)
        
        # Title
        title_label = tk.Label(header, text="About", 
                             font=("Arial", 14, "bold"), fg="white", bg=self.primary_color)
        title_label.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Content area with railway track decoration at top
        track = tk.Frame(about_window, bg=self.track_color, height=8)
        track.pack(fill=tk.X)
        
        content_frame = tk.Frame(about_window, bg=self.bg_color, padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # App name
        tk.Label(content_frame, text="Railway Accidents Report Analysis System", 
               font=("Arial", 12, "bold"), 
               fg=self.primary_color, bg=self.bg_color).pack(pady=(10, 5))
        
        # Version
        tk.Label(content_frame, text="Version 1.2", 
               font=("Arial", 10), 
               fg=self.primary_color, bg=self.bg_color).pack(pady=2)
        
        # Description
        tk.Label(content_frame, text="This application analyses railway accident reports\nfrom the ERAIL database using LLMs.", 
               font=("Arial", 10), 
               fg=self.track_color, bg=self.bg_color,
               justify=tk.CENTER).pack(pady=(15, 20))
        
        # Railway themed close button
        close_button = tk.Button(content_frame, text="Close", 
                               command=about_window.destroy,
                               bg=self.primary_color, fg="white",
                               font=("Arial", 10, "bold"),
                               padx=15, pady=5,
                               activebackground=self.secondary_color,
                               activeforeground="white")
        close_button.pack(pady=10)
        
        # Center the window
        about_window.update_idletasks()
        width = about_window.winfo_width()
        height = about_window.winfo_height()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

if __name__ == "__main__":
    root = tk.Tk()
    app = AccidentAnalyzerApp(root)
    root.mainloop()