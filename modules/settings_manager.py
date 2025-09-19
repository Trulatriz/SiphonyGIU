import json
import os
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

class SettingsManager:
    """Manages user settings and preferences"""
    
    def __init__(self, settings_file="settings.json"):
        self.settings_file = settings_file
        self.default_settings = {
            "last_output_file": "",
            "last_doe_file": "",
            "last_density_file": "",
            "last_pdr_file": "",
            "last_histogram_folder": "",
            "last_dsc_folder": "",
            "last_oc_folder": "",
            "last_heatmap_file": "",
            "auto_save_enabled": True,
            "auto_save_interval": 10,
            "backup_enabled": True,
            "validation_enabled": True,
            "parallel_processing": False,
            "theme": "light",
            "window_geometry": "1200x900",
            "recent_files": [],
            "max_recent_files": 10
        }
        self.settings = self.load_settings()
    
    def load_settings(self):
        """Load settings from file"""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    loaded_settings = json.load(f)
                    # Merge with defaults to handle new settings
                    settings = self.default_settings.copy()
                    settings.update(loaded_settings)
                    return settings
            except:
                return self.default_settings.copy()
        return self.default_settings.copy()
    
    def save_settings(self):
        """Save settings to file"""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")
    
    def get(self, key, default=None):
        """Get a setting value"""
        return self.settings.get(key, default)
    
    def set(self, key, value):
        """Set a setting value"""
        self.settings[key] = value
        self.save_settings()
    
    def add_recent_file(self, filepath):
        """Add file to recent files list"""
        recent = self.settings.get("recent_files", [])
        if filepath in recent:
            recent.remove(filepath)
        recent.insert(0, filepath)
        recent = recent[:self.settings.get("max_recent_files", 10)]
        self.set("recent_files", recent)


class SettingsDialog:
    """Settings dialog window"""
    
    def __init__(self, parent, settings_manager):
        self.parent = parent
        self.settings = settings_manager
        self.window = tk.Toplevel(parent)
        self.window.title("Settings")
        self.window.geometry("500x600")
        self.window.grab_set()  # Make modal
        
        self.create_widgets()
        self.load_current_settings()
    
    def create_widgets(self):
        """Create the settings interface"""
        # Main frame with scrollbar
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create notebook for different setting categories
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # General settings
        general_frame = ttk.Frame(notebook, padding="10")
        notebook.add(general_frame, text="General")
        self.create_general_settings(general_frame)
        
        # Processing settings  
        processing_frame = ttk.Frame(notebook, padding="10")
        notebook.add(processing_frame, text="Processing")
        self.create_processing_settings(processing_frame)
        
        # File settings
        files_frame = ttk.Frame(notebook, padding="10")
        notebook.add(files_frame, text="Default Files")
        self.create_file_settings(files_frame)
        
        # Advanced settings
        advanced_frame = ttk.Frame(notebook, padding="10")
        notebook.add(advanced_frame, text="Advanced")
        self.create_advanced_settings(advanced_frame)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Save", command=self.save_settings).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.window.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_defaults).pack(side=tk.LEFT)
    
    def create_general_settings(self, parent):
        """Create general settings section"""
        # Auto-save
        self.auto_save_var = tk.BooleanVar()
        ttk.Checkbutton(parent, text="Enable auto-save during processing", 
                       variable=self.auto_save_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Auto-save interval
        ttk.Label(parent, text="Auto-save interval (labels):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.auto_save_interval_var = tk.StringVar()
        interval_spinner = ttk.Spinbox(parent, from_=5, to=50, textvariable=self.auto_save_interval_var, width=10)
        interval_spinner.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Backup
        self.backup_var = tk.BooleanVar()
        ttk.Checkbutton(parent, text="Create backup before processing", 
                       variable=self.backup_var).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Max recent files
        ttk.Label(parent, text="Max recent files:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.max_recent_var = tk.StringVar()
        recent_spinner = ttk.Spinbox(parent, from_=5, to=20, textvariable=self.max_recent_var, width=10)
        recent_spinner.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=5)
    
    def create_processing_settings(self, parent):
        """Create processing settings section"""
        # Validation
        self.validation_var = tk.BooleanVar()
        ttk.Checkbutton(parent, text="Enable data validation", 
                       variable=self.validation_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Parallel processing
        self.parallel_var = tk.BooleanVar()
        ttk.Checkbutton(parent, text="Enable parallel processing (experimental)", 
                       variable=self.parallel_var).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Theme
        ttk.Label(parent, text="Theme:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.theme_var = tk.StringVar()
        theme_combo = ttk.Combobox(parent, textvariable=self.theme_var, 
                                  values=["light", "dark"], state="readonly", width=15)
        theme_combo.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)
    
    def create_file_settings(self, parent):
        """Create default file settings section"""
        parent.columnconfigure(1, weight=1)
        
        # Default output file
        ttk.Label(parent, text="Default output file:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_file_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.output_file_var, state='readonly').grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(parent, text="Browse", command=self.browse_output_file).grid(row=0, column=2, pady=5)
        
        # Default DoE file
        ttk.Label(parent, text="Default DoE file:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.doe_file_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.doe_file_var, state='readonly').grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(parent, text="Browse", command=self.browse_doe_file).grid(row=1, column=2, pady=5)
        
        # Clear defaults button
        ttk.Button(parent, text="Clear All Defaults", command=self.clear_defaults).grid(row=2, column=0, pady=10)
    
    def create_advanced_settings(self, parent):
        """Create advanced settings section"""
        # Window geometry
        ttk.Label(parent, text="Default window size:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.geometry_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.geometry_var, width=15).grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Settings file location
        ttk.Label(parent, text="Settings file:").grid(row=1, column=0, sticky=tk.W, pady=5)
        settings_path = os.path.abspath(self.settings.settings_file)
        ttk.Label(parent, text=settings_path, foreground="gray").grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Export/Import settings
        ttk.Button(parent, text="Export Settings", command=self.export_settings).grid(row=2, column=0, pady=10)
        ttk.Button(parent, text="Import Settings", command=self.import_settings).grid(row=2, column=1, padx=(10, 0), pady=10)
    
    def load_current_settings(self):
        """Load current settings into the dialog"""
        self.auto_save_var.set(self.settings.get("auto_save_enabled", True))
        self.auto_save_interval_var.set(str(self.settings.get("auto_save_interval", 10)))
        self.backup_var.set(self.settings.get("backup_enabled", True))
        self.validation_var.set(self.settings.get("validation_enabled", True))
        self.parallel_var.set(self.settings.get("parallel_processing", False))
        self.theme_var.set(self.settings.get("theme", "light"))
        self.max_recent_var.set(str(self.settings.get("max_recent_files", 10)))
        self.geometry_var.set(self.settings.get("window_geometry", "1200x900"))
        self.output_file_var.set(self.settings.get("last_output_file", ""))
        self.doe_file_var.set(self.settings.get("last_doe_file", ""))
    
    def save_settings(self):
        """Save all settings"""
        try:
            self.settings.set("auto_save_enabled", self.auto_save_var.get())
            self.settings.set("auto_save_interval", int(self.auto_save_interval_var.get()))
            self.settings.set("backup_enabled", self.backup_var.get())
            self.settings.set("validation_enabled", self.validation_var.get())
            self.settings.set("parallel_processing", self.parallel_var.get())
            self.settings.set("theme", self.theme_var.get())
            self.settings.set("max_recent_files", int(self.max_recent_var.get()))
            self.settings.set("window_geometry", self.geometry_var.get())
            self.settings.set("last_output_file", self.output_file_var.get())
            self.settings.set("last_doe_file", self.doe_file_var.get())
            
            messagebox.showinfo("Success", "Settings saved successfully!")
            self.window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def reset_defaults(self):
        """Reset all settings to defaults"""
        result = messagebox.askyesno("Confirm Reset", "Reset all settings to defaults?")
        if result:
            self.settings.settings = self.settings.default_settings.copy()
            self.settings.save_settings()
            self.load_current_settings()
    
    def clear_defaults(self):
        """Clear default file paths"""
        self.output_file_var.set("")
        self.doe_file_var.set("")
    
    def browse_output_file(self):
        """Browse for default output file"""
        filename = filedialog.asksaveasfilename(
            title="Select default output file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file_var.set(filename)
    
    def browse_doe_file(self):
        """Browse for default DoE file"""
        filename = filedialog.askopenfilename(
            title="Select default DoE file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.doe_file_var.set(filename)
    
    def export_settings(self):
        """Export settings to file"""
        filename = filedialog.asksaveasfilename(
            title="Export settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(self.settings.settings, f, indent=2)
                messagebox.showinfo("Success", "Settings exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export settings: {str(e)}")
    
    def import_settings(self):
        """Import settings from file"""
        filename = filedialog.askopenfilename(
            title="Import settings",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    imported_settings = json.load(f)
                
                # Validate imported settings
                valid_keys = set(self.settings.default_settings.keys())
                imported_keys = set(imported_settings.keys())
                
                if not imported_keys.issubset(valid_keys):
                    messagebox.showwarning("Warning", "Some imported settings are not recognized and will be ignored.")
                
                # Update settings
                for key, value in imported_settings.items():
                    if key in valid_keys:
                        self.settings.set(key, value)
                
                self.load_current_settings()
                messagebox.showinfo("Success", "Settings imported successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import settings: {str(e)}")
