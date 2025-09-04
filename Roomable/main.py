import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import re
from datetime import datetime
import os
from tkinter.font import Font

class RoomAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RoomAble - Available Rooms Finder")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Style configuration
        self.setup_styles()
        
        # Variables
        self.schedule_data = []
        self.all_rooms = set()
        self.selected_time = tk.StringVar(value="10:00 AM")
        self.selected_day = tk.StringVar(value="Tue")
        self.default_file = "class_schedule.xlsx"
        
        # UI Setup
        self.setup_ui()
        
        # Try loading default file automatically
        self.try_load_default_file()
    
    def setup_styles(self):
        """Configure modern styles for the application"""
        style = ttk.Style()
        
        # Theme settings
        style.theme_use('clam')
        
        # Colors
        bg_color = "#f5f7fa"
        primary_color = "#4f46e5"
        secondary_color = "#6366f1"
        accent_color = "#10b981"
        text_color = "#1e293b"
        
        # Configure styles
        style.configure('.', background=bg_color, foreground=text_color)
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), borderwidth=1)
        style.configure('Accent.TButton', foreground='white', background=primary_color, 
                       font=('Segoe UI', 10, 'bold'), borderwidth=0)
        style.configure('Secondary.TButton', foreground='white', background=secondary_color, 
                       borderwidth=0)
        style.configure('TEntry', fieldbackground='white', borderwidth=1)
        style.configure('TCombobox', fieldbackground='white')
        style.configure('Treeview', background='white', fieldbackground='white', 
                        rowheight=25, font=('Segoe UI', 10))
        style.configure('Treeview.Heading', background=primary_color, foreground='white', 
                       font=('Segoe UI', 10, 'bold'))
        style.map('Treeview.Heading', background=[('active', secondary_color)])
        style.map('Treeview', background=[('selected', secondary_color)])
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), foreground=primary_color)
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#64748b')
        
        # Notebook style
        style.configure('TNotebook', background=bg_color, borderwidth=0)
        style.configure('TNotebook.Tab', background=bg_color, padding=[10, 5], 
                        font=('Segoe UI', 10, 'bold'))
        style.map('TNotebook.Tab', background=[('selected', 'white')])
    
    def setup_ui(self):
        """Create the modern UI layout"""
        # Main container with padding
        main_container = ttk.Frame(self.root, padding=(20, 10))
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header section
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # App title
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        ttk.Label(title_frame, text="RoomAble", style='Header.TLabel', 
                 font=('Segoe UI', 18, 'bold')).pack(anchor='w')
        ttk.Label(title_frame, text="Available Rooms Finder", style='Status.TLabel').pack(anchor='w')
        
        # File selection section
        file_frame = ttk.LabelFrame(main_container, text=" Schedule File ", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # File entry with browse buttons
        file_controls = ttk.Frame(file_frame)
        file_controls.pack(fill=tk.X)
        
        ttk.Label(file_controls, text="Select schedule file:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.file_entry = ttk.Entry(file_controls, width=50)
        self.file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        button_frame = ttk.Frame(file_controls)
        button_frame.pack(side=tk.LEFT)
        
        ttk.Button(
            button_frame, 
            text="Load Default",
            command=self.load_default_file,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            button_frame,
            text="Browse...",
            command=self.load_other_file,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=2)
        
        # Time controls section
        time_frame = ttk.LabelFrame(main_container, text=" Time Selection ", padding=10)
        time_frame.pack(fill=tk.X, pady=(0, 15))
        
        time_controls = ttk.Frame(time_frame)
        time_controls.pack(fill=tk.X)
        
        # Current time button
        ttk.Button(
            time_controls,
            text="Use Current Time",
            command=self.use_current_time,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        # Time entry
        time_entry_frame = ttk.Frame(time_controls)
        time_entry_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(time_entry_frame, text="Time:").pack(anchor='w')
        self.time_entry = ttk.Entry(time_entry_frame, textvariable=self.selected_time, width=10)
        self.time_entry.pack()
        
        # Day dropdown
        day_frame = ttk.Frame(time_controls)
        day_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(day_frame, text="Day:").pack(anchor='w')
        self.day_combo = ttk.Combobox(
            day_frame, 
            textvariable=self.selected_day, 
            values=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
            state="readonly", 
            width=8
        )
        self.day_combo.pack()
        
        # Check availability button
        ttk.Button(
            time_controls,
            text="Find Available Rooms",
            command=self.check_availability,
            style='Accent.TButton'
        ).pack(side=tk.RIGHT)
        
        # Status bar
        self.status_bar = ttk.Frame(main_container, height=25)
        self.status_bar.pack(fill=tk.X, pady=(5, 0))
        self.status_label = ttk.Label(
            self.status_bar, 
            text="Ready", 
            style='Status.TLabel',
            anchor='w'
        )
        self.status_label.pack(fill=tk.X, padx=5)
        
        # Separator
        ttk.Separator(main_container).pack(fill=tk.X, pady=5)
        
        # Results section
        results_frame = ttk.Frame(main_container)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Building tabs
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.building_trees = {}
        buildings = ["PA", "PB", "PC", "PD"]
        for building in buildings:
            tab_frame = ttk.Frame(self.notebook, padding=5)
            self.notebook.add(tab_frame, text=f"{building} Building")
            
            # Treeview with scrollbars
            tree_container = ttk.Frame(tab_frame)
            tree_container.pack(fill=tk.BOTH, expand=True)
            
            # Vertical scrollbar
            y_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
            y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Horizontal scrollbar
            x_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
            x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
            
            # Treeview
            tree = ttk.Treeview(
                tree_container,
                columns=("Room", "Floor", "Status"),
                show="headings",
                yscrollcommand=y_scroll.set,
                xscrollcommand=x_scroll.set
            )
            tree.pack(fill=tk.BOTH, expand=True)
            
            # Configure scrollbars
            y_scroll.config(command=tree.yview)
            x_scroll.config(command=tree.xview)
            
            # Configure columns
            tree.heading("Room", text="Room")
            tree.heading("Floor", text="Floor")
            tree.heading("Status", text="Status")
            
            tree.column("Room", width=150, anchor=tk.CENTER)
            tree.column("Floor", width=80, anchor=tk.CENTER)
            tree.column("Status", width=100, anchor=tk.CENTER)
            
            # Tag configurations for alternating row colors
            tree.tag_configure('oddrow', background='#f8fafc')
            tree.tag_configure('evenrow', background='white')
            
            self.building_trees[building] = tree
    
    def use_current_time(self):
        """Set to current time and day, then check availability"""
        now = datetime.now()
        
        # Set time in 12-hour format with AM/PM
        current_time = now.strftime("%I:%M %p").lstrip("0")
        self.selected_time.set(current_time)
        
        # Set current day (first 3 letters)
        current_day = now.strftime("%a")
        self.selected_day.set(current_day)
        
        # Automatically check availability
        self.check_availability()
        self.status_label.config(text=f"Checked availability for current time: {current_time} {current_day}")
    
    def try_load_default_file(self):
        """Try to load default file if it exists in the same directory"""
        if os.path.exists(self.default_file):
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.default_file)
            self.process_file(self.default_file)
            self.status_label.config(text=f"Loaded default file: {self.default_file}")
        else:
            self.status_label.config(text="Default file not found. Please load a schedule file.")
    
    def load_default_file(self):
        """Explicitly load the default file"""
        if os.path.exists(self.default_file):
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.default_file)
            self.process_file(self.default_file)
            self.status_label.config(text=f"Loaded default file: {self.default_file}")
            self.check_availability()
        else:
            messagebox.showwarning("File Not Found", 
                                 f"Default file '{self.default_file}' not found in the program directory.")
    
    def load_other_file(self):
        """Browse and load another schedule file"""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
            title="Select Schedule File"
        )
        if filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filepath)
            self.process_file(filepath)
            self.status_label.config(text=f"Loaded file: {os.path.basename(filepath)}")
            self.check_availability()
    
    def parse_day(self, day_str):
        """Handle combined days like 'Sun &Tue' or 'Thu & SAT'"""
        day_str = str(day_str).strip()
        if "&" in day_str:
            return [d.strip()[:3] for d in day_str.split("&")]
        return [day_str[:3]]
    
    def extract_rooms(self, text):
        """Extract all room numbers from a cell's text"""
        room_pattern = re.compile(r'\b(PA|PB|PC|PD)\s*\d+[A-Z]?\b', re.IGNORECASE)
        return [match.group().upper() for match in room_pattern.finditer(text)]
    
    def parse_time(self, time_str):
        """Handle all time formats in the schedule"""
        # Clean inconsistent formatting
        time_str = re.sub(r'\s*:\s*', ':', time_str)  # Fix "1 : 40" -> "1:40"
        time_str = re.sub(r'\s+', ' ', time_str).strip()  # Normalize spaces
        
        # Handle time ranges
        if "-" in time_str:
            time_str = time_str.split("-")[0].strip()
        
        try:
            return datetime.strptime(time_str, "%I:%M %p").time()
        except ValueError:
            try:
                return datetime.strptime(time_str, "%H:%M").time()
            except ValueError:
                messagebox.showerror("Invalid Time Format", 
                                   f"Could not parse time: {time_str}\nPlease use HH:MM AM/PM format")
                raise
    
    def process_file(self, filepath):
        try:
            df = pd.read_excel(filepath, header=None)
            self.schedule_data = []
            self.all_rooms = set()
            
            # Find the header row with "Day /Time"
            header_row = None
            for i, row in df.iterrows():
                if "Day /Time" in str(row[0]):
                    header_row = i
                    break
            
            if header_row is None:
                messagebox.showerror("Error", "Could not find 'Day /Time' header row in the schedule file")
                return
            
            # Process schedule data
            current_days = None
            for i in range(header_row + 1, len(df)):
                row = df.iloc[i]
                
                # Get day(s) if available (first column)
                if pd.notna(row[0]):
                    current_days = self.parse_day(row[0])
                
                if not current_days:
                    continue
                
                # Process each time slot (columns 1-7)
                for j in range(1, 8):
                    if pd.isna(row[j]):
                        continue
                    
                    time_slot = df.iloc[header_row, j]
                    cell_content = str(row[j])
                    
                    # Extract all rooms from this cell
                    rooms = self.extract_rooms(cell_content)
                    for room in rooms:
                        self.all_rooms.add(room)
                        
                        # Create entries for each applicable day
                        for day in current_days:
                            self.schedule_data.append({
                                'day': day,
                                'time_slot': time_slot,
                                'room': room
                            })
            
            self.status_label.config(text=f"Loaded schedule with {len(self.all_rooms)} rooms")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{str(e)}")
            self.status_label.config(text="Error loading file")
    
    def check_availability(self):
        if not self.schedule_data:
            messagebox.showwarning("Warning", "No schedule data loaded")
            return
        
        try:
            check_time = self.parse_time(self.selected_time.get())
            check_day = self.selected_day.get()[:3]  # Use first 3 letters for matching
            
            # Clear all trees first
            for tree in self.building_trees.values():
                tree.delete(*tree.get_children())
            
            # Find occupied rooms at selected time
            occupied_rooms = set()
            for entry in self.schedule_data:
                if entry['day'] != check_day:
                    continue
                
                try:
                    time_slot = entry['time_slot']
                    if not isinstance(time_slot, str) or "-" not in time_slot:
                        continue
                    
                    # Extract and clean start/end times
                    start_str, end_str = [t.strip() for t in time_slot.split("-", 1)]
                    start_time = self.parse_time(start_str)
                    end_time = self.parse_time(end_str)
                    
                    if start_time <= check_time <= end_time:
                        occupied_rooms.add(entry['room'])
                except Exception as e:
                    print(f"Skipping invalid time slot: {time_slot} - {str(e)}")
            
            # Display only available rooms
            available_count = 0
            for building, tree in self.building_trees.items():
                # Get available rooms for this building
                available_rooms = [
                    room for room in self.all_rooms 
                    if room.startswith(building) and room not in occupied_rooms
                ]
                available_count += len(available_rooms)
                
                # Add to treeview sorted by room number with alternating colors
                for i, room in enumerate(sorted(available_rooms, key=lambda x: int(re.search(r'\d+', x).group()))):
                    floor = re.search(r'\d+', room).group()[0]  # First digit of room number
                    tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
                    tree.insert("", "end", values=(room, floor, "Available"), tags=tags)
            
            # Update status bar
            time_str = self.selected_time.get()
            day_str = self.selected_day.get()
            self.status_label.config(
                text=f"Found {available_count} available rooms at {time_str} on {day_str}"
            )
            
            # Show notification if no rooms available
            if available_count == 0:
                messagebox.showinfo("No Rooms Available", 
                                  f"No rooms available at {time_str} on {day_str}")
        
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid time format: {str(e)}\nPlease use HH:MM AM/PM")
            self.status_label.config(text="Error checking availability")

if __name__ == "__main__":
    root = tk.Tk()
    app = RoomAvailabilityApp(root)
    
    # Set window icon if available
    try:
        root.iconbitmap("icon.ico")  # Replace with your icon file if available
    except:
        pass
    
    root.mainloop()
