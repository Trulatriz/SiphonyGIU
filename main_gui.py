import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from pathlib import Path

# Import modules for each functionality
from modules.combine_module import CombineModule
from modules.dsc_module import DSCModule
from modules.sem_module import SEMImageEditor as SEMModule
from modules.oc_module import OCModule
from modules.pdr_module import PDRModule
from modules.plot_module import PlotModule
from modules.heatmap_module import HeatmapModule
from modules.settings_manager import SettingsManager, SettingsDialog
from modules.foam_type_manager import FoamTypeManager, FoamTypeDialog, PaperDialog, NewPaperDialog, ManageFoamsDialog, ManagePapersDialog

class PressTechGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PressTech - Siphony GUI")

        # Initialize settings and managers
        self.settings = SettingsManager()
        self.foam_manager = FoamTypeManager()

        # Apply saved geometry or default
        geometry = self.settings.get("window_geometry", "800x600")
        self.root.geometry(geometry)
        self.root.configure(bg='#f0f0f0')

        # Center the window
        self.center_window()

        # Ask for paper and foam type before enabling modules
        self.prompt_paper_and_foam_startup()

        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        # Create UI
        self.create_title()
        self.create_buttons()
        self.create_status_bar()
        self.create_menu_bar()

        # Bind close event to save settings
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
    
    def create_title(self):
        """Create the main title"""
        title_frame = ttk.Frame(self.main_frame)
        title_frame.grid(row=0, column=0, pady=(0, 30), sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(
            title_frame, 
            text="PressTech - Siphony GUI", 
            font=('Arial', 24, 'bold')
        )
        title_label.grid(row=0, column=0)
        
        subtitle_label = ttk.Label(
            title_frame, 
            text="Integrated Analysis Tools for Foam Processing", 
            font=('Arial', 12)
        )
        subtitle_label.grid(row=1, column=0, pady=(5, 0))

        # Current paper and foam type labels
        self.paper_label = ttk.Label(
            title_frame,
            text=f"Paper: {self.foam_manager.get_current_paper()}",
            font=('Arial', 10, 'italic')
        )
        self.paper_label.grid(row=2, column=0, pady=(8, 0), sticky=tk.W)

        self.foam_type_label = ttk.Label(
            title_frame,
            text=f"Foam Type: {self.foam_manager.get_current_foam_type()}",
            font=('Arial', 10, 'italic')
        )
        self.foam_type_label.grid(row=3, column=0, sticky=tk.W)
    
    def create_buttons(self):
        """Create the main functionality buttons"""
        buttons_frame = ttk.Frame(self.main_frame, padding=(0, 10))
        buttons_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        buttons_frame.columnconfigure(0, weight=1)

        button_style = {
            'width': 25,
            'padding': 15
        }

        extraction_frame = ttk.LabelFrame(buttons_frame, text='DATA EXTRACTION', padding=12)
        extraction_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        extraction_frame.columnconfigure((0, 1), weight=1)

        combine_btn = ttk.Button(
            extraction_frame,
            text='⚡ SMART COMBINE',
            command=self.open_combine,
            **button_style
        )
        combine_btn.grid(row=0, column=0, padx=10, pady=6, sticky=(tk.W, tk.E))

        specific_btn = ttk.Button(
            extraction_frame,
            text='🔬 FOAM-SPECIFIC ANALYSIS',
            command=self.show_foam_specific_dialog,
            **button_style
        )
        specific_btn.grid(row=0, column=1, padx=10, pady=6, sticky=(tk.W, tk.E))

        analysis_frame = ttk.LabelFrame(buttons_frame, text='DATA ANALYSIS', padding=12)
        analysis_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        analysis_frame.columnconfigure((0, 1), weight=1)

        analysis_btn = ttk.Button(
            analysis_frame,
            text='Publication Plots (Scatter)',
            command=self.open_publication_plots,
            **button_style
        )
        analysis_btn.grid(row=0, column=0, padx=10, pady=6, sticky=(tk.W, tk.E))

        heatmap_btn = ttk.Button(
            analysis_frame,
            text='🔥 HEATMAPS',
            command=self.open_heatmap,
            **button_style
        )
        heatmap_btn.grid(row=0, column=1, padx=10, pady=6, sticky=(tk.W, tk.E))

        org_frame = ttk.LabelFrame(buttons_frame, text='ORGANIZATION', padding=12)
        org_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        org_frame.columnconfigure((0, 1), weight=1)

        manage_papers_btn = ttk.Button(
            org_frame,
            text='📁 MANAGE PAPERS',
            command=self.manage_papers,
            **button_style
        )
        manage_papers_btn.grid(row=0, column=0, padx=10, pady=6, sticky=(tk.W, tk.E))

        manage_foams_btn = ttk.Button(
            org_frame,
            text='🧶 MANAGE FOAMS',
            command=self.manage_foams,
            **button_style
        )
        manage_foams_btn.grid(row=0, column=1, padx=10, pady=6, sticky=(tk.W, tk.E))

        self.add_tooltips(analysis_btn, combine_btn, specific_btn, manage_papers_btn, manage_foams_btn, heatmap_btn)

    def add_tooltips(self, *buttons):
        """Add tooltips to buttons"""
        tooltip_map = {}
        for b in buttons:
            txt = b.cget('text') if hasattr(b, 'cget') else None
            if not txt:
                continue
            if 'Publication Plots' in txt:
                tooltip_map[b] = "Plot All_Results_* with grouping and error bars"
            elif 'SMART COMBINE' in txt:
                tooltip_map[b] = "Smart combine with tracking and incremental updates"
            elif 'FOAM-SPECIFIC' in txt:
                tooltip_map[b] = "Open foam-specific analysis (DSC, SEM, OC, PDR)"
            elif 'MANAGE PAPERS' in txt:
                tooltip_map[b] = "Review papers, base folders, and delete or relocate"
            elif 'MANAGE FOAMS' in txt:
                tooltip_map[b] = "Edit foam types per paper and global list"
            elif 'HEATMAPS' in txt:
                tooltip_map[b] = "Spearman heatmaps with column selection"

        for button, text in tooltip_map.items():
            self.create_tooltip(button, text)

    def create_tooltip(self, widget, text):
        """Create and attach a simple tooltip to a widget"""
        def show_tooltip(event):
            tooltip = tk.Toplevel(self.root)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = ttk.Label(
                tooltip,
                text=text,
                background="lightyellow",
                relief="solid",
                borderwidth=1,
                font=('Arial', 9)
            )
            label.pack()
            widget.tooltip = tooltip

        def hide_tooltip(_event):
            if hasattr(widget, 'tooltip') and widget.tooltip:
                try:
                    widget.tooltip.destroy()
                except Exception:
                    pass
                widget.tooltip = None

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)
        """Create a tooltip for a widget"""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = ttk.Label(
                tooltip, 
                text=text, 
                background="lightyellow",
                relief="solid",
                borderwidth=1,
                font=('Arial', 9)
            )
            label.pack()
            
            widget.tooltip = tooltip
        
        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)
    
    def create_status_bar(self):
        """Create status bar at the bottom"""
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=2, column=0, pady=(30, 0), sticky=(tk.W, tk.E))
        
        status_label = ttk.Label(
            status_frame, 
            textvariable=self.status_var,
            font=('Arial', 9),
            foreground='gray'
        )
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Add version info
        version_label = ttk.Label(
            status_frame, 
            text="v1.0",
            font=('Arial', 9),
            foreground='gray'
        )
        version_label.grid(row=0, column=1, sticky=tk.E)
    
    def create_menu_bar(self):
        """Create the menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Settings", command=self.open_settings)
        file_menu.add_command(label="Select Paper...", command=self.switch_paper)
        file_menu.add_command(label="Manage Papers...", command=self.manage_papers)
        file_menu.add_command(label="Select Foam Type...", command=self.switch_foam_type)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Heatmaps", command=self.open_heatmap)

        # Plots menu
        plots_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Plots", menu=plots_menu)
        plots_menu.add_command(label="Publication Plots (Scatter)", command=self.open_publication_plots)
        tools_menu.add_command(label="📊 Publication Plots", command=self.open_analysis)
        tools_menu.add_command(label="⚡ Smart Combine", command=self.open_combine)
        tools_menu.add_command(label="🔬 Cell Size & Density", command=self.open_cell_analysis)
        tools_menu.add_separator()
        tools_menu.add_command(label="🌡️ DSC Analysis", command=self.open_dsc_with_foam_check)
        tools_menu.add_command(label="🔬 SEM Image Editor", command=self.open_sem_with_foam_check)
        tools_menu.add_command(label="🔓 Open-Cell Content", command=self.open_oc_with_foam_check)
        tools_menu.add_command(label="📊 Pressure Drop Rate", command=self.open_pdr_with_foam_check)
        tools_menu.add_separator()
        tools_menu.add_command(label="🔬 Foam-Specific Analysis…", command=self.show_foam_specific_dialog)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Instrucciones (HELP.md)", command=self.show_help)
        help_menu.add_command(label="Documentation", command=self.show_documentation)
        help_menu.add_command(label="About", command=self.show_about)

    def show_foam_specific_dialog(self):
        """Open foam-specific analysis dialog with Back to main screen."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Foam-Specific Analysis")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=20)
        frame.grid(row=0, column=0)

        ttk.Label(frame, text="Foam-Specific Analysis", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 15))

        dsc_btn = ttk.Button(frame, text="🌡️ DSC Analysis", command=lambda: (dialog.destroy(), self.open_dsc_with_foam_check()), width=25)
        dsc_btn.grid(row=1, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        sem_btn = ttk.Button(frame, text="🔬 SEM Image Editor", command=lambda: (dialog.destroy(), self.open_sem_with_foam_check()), width=25)
        sem_btn.grid(row=1, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
        oc_btn = ttk.Button(frame, text="🔓 Open-Cell Content", command=lambda: (dialog.destroy(), self.open_oc_with_foam_check()), width=25)
        oc_btn.grid(row=2, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        pdr_btn = ttk.Button(frame, text="📉 Pressure Drop Rate", command=lambda: (dialog.destroy(), self.open_pdr_with_foam_check()), width=25)
        pdr_btn.grid(row=2, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
        # SEM results workflow (direct access)
        obtain_hist_btn = ttk.Button(frame, text="Obtain SEM results", command=lambda: (dialog.destroy(), self.show_cell_analysis_instructions()), width=25)
        obtain_hist_btn.grid(row=3, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        combine_hist_btn = ttk.Button(frame, text="Combine SEM results", command=lambda: (dialog.destroy(), self.open_histogram_combiner()), width=25)
        combine_hist_btn.grid(row=3, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

        back_btn = ttk.Button(frame, text="⬅️ Back", command=dialog.destroy)
        back_btn.grid(row=4, column=0, columnspan=2, pady=(12, 0))

        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def open_analysis(self):
        """Legacy entry: route to Publication Plots."""
        return self.open_publication_plots()






    
    def open_combine(self):
        """Open Smart Combine Results module"""
        self.update_status("Opening Smart Combine Results module...")
        try:
            combine_window = tk.Toplevel(self.root)
            # Pass current paper path to CombineModule
            current_paper_path = getattr(self, 'current_paper_path', None)
            CombineModule(combine_window, current_paper_path)
            self.update_status("Smart Combine Results module opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Smart Combine module: {str(e)}")
            self.update_status("Error opening Smart Combine Results module")
    
    def open_cell_analysis(self):
        """Obtain SEM results: show a single entry to instructions with link"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Obtain SEM Results")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("450x240")
        dialog.resizable(False, False)
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Obtain SEM Results", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))

        # Description
        desc_label = ttk.Label(main_frame, 
                              text="Open the instructions (includes link to the web tool)",
                              font=("Arial", 10))
        desc_label.pack(pady=(0, 10))

        # Single action: open instructions (which contains link to webpage)
        action_frame = ttk.LabelFrame(main_frame, text="Instructions", padding=15)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        action_desc = ttk.Label(action_frame, 
                                text="Use the web tool to analyze SEM images and\ngenerate SEM result spreadsheets.",
                                font=("Arial", 9))
        action_desc.pack(pady=(0, 10))

        action_btn = ttk.Button(action_frame, text="Obtain SEM results (instructions)",
                                command=lambda: (dialog.destroy(), self.show_cell_analysis_instructions()))
        action_btn.pack()

        # Close button
        close_btn = ttk.Button(main_frame, text="Close", command=dialog.destroy)
        close_btn.pack(pady=(15, 0))
    
    def show_cell_analysis_instructions(self):
        """Show instructions for Cell Size and Cell Density analysis"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Cell Size & Cell Density - Instructions")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("700x600")
        
        # Main frame with scrollbar
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create canvas and scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Title
        title_label = ttk.Label(scrollable_frame, text="Cell Size & Cell Density Analysis", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions content
        instructions = [
            ("📁 Preparación de imágenes:", [
                "• Una muestra por vez (no mezclar diferentes muestras)",
                "• Nombres similares para réplicas de la misma muestra",
                "• Añadir '_1' a la primera réplica si las otras tienen '_2', '_3', etc.",
                "• Reemplazar espacios con guiones bajos (_)",
                "",
                "Ejemplo de nombres correctos:",
                "  PS_20250214_1_001.tif",
                "  PS_20250214_1_005.tif", 
                "  PS_20250214_2_001.tif",
                "  PS_20250214_2_002.tif"
            ]),
            ("⚙️ Datos de entrada:", [
                "• Densidad: Introducir densidad del material sólido (la del espumado no se usa)",
                "• Escala (preferible manual para evitar errores):",
                "  - Automática: se detecta de las líneas de referencia x10 = micrómetros",
                "  - Manual: introducir el número que aparece abajo a la derecha de la imagen",
                "    (en micrómetros), separados por comas en orden de las imágenes"
            ]),
            ("🔍 Tipos de análisis:", [
                "• Automático: Sube imágenes → detecta poros automáticamente → genera histograma",
                "• Con ROIs: Usa regiones predefinidas/editadas en ImageJ (archivo '_rois.zip')"
            ]),
            ("📊 Resultados:", [
                "• Histograma Excel con hoja por réplica + hoja combinada",
                "• Imágenes con poros detectados",
                "• ROIs editables en ImageJ si es necesario"
            ]),
            ("⚠️ Errores comunes:", [
                "• Error 500: Escala mal detectada → usar escala manual",
                "• 'Failed to fetch': Nombres de archivo incorrectos"
            ])
        ]
        
        for section_title, items in instructions:
            # Section title
            section_label = ttk.Label(scrollable_frame, text=section_title, 
                                    font=("Arial", 12, "bold"), foreground="blue")
            section_label.pack(anchor="w", pady=(10, 5))
            
            # Section items
            for item in items:
                item_label = ttk.Label(scrollable_frame, text=item, 
                                     font=("Arial", 10), wraplength=650)
                item_label.pack(anchor="w", padx=(20, 0), pady=1)
        
        # Button frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=(30, 10))
        
        # Open web tool button
        web_btn = ttk.Button(button_frame, text="🌐 Abrir Herramienta Web", 
                           command=lambda: self.open_web_tool())
        web_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Close button
        close_btn = ttk.Button(button_frame, text="❌ Cerrar", command=dialog.destroy)
        close_btn.pack(side=tk.LEFT)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Center dialog
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def open_dsc(self):
        """Open DSC Analysis module"""
        self.update_status("Opening DSC Analysis module...")
        try:
            dsc_window = tk.Toplevel(self.root)
            DSCModule(dsc_window)
            self.update_status("DSC Analysis module opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open DSC module: {str(e)}")
            self.update_status("Error opening DSC Analysis module")
    
    def open_sem(self):
        """Open SEM Image Editor module"""
        self.update_status("Opening SEM Image Editor module...")
        try:
            sem_window = tk.Toplevel(self.root)
            SEMModule(sem_window)
            self.update_status("SEM Image Editor module opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open SEM module: {str(e)}")
            self.update_status("Error opening SEM Image Editor module")
    
    def open_oc(self):
        """Open Open-Cell Content module"""
        self.update_status("Opening Open-Cell Content module...")
        try:
            oc_window = tk.Toplevel(self.root)
            # Pass current paper path and foam type to OCModule
            current_paper_path = getattr(self, 'current_paper_path', None)
            current_foam_type = getattr(self, 'current_foam_type', None)
            OCModule(oc_window, current_paper_path, current_foam_type)
            self.update_status("Open-Cell Content module opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open OC module: {str(e)}")
            self.update_status("Error opening Open-Cell Content module")
    
    def open_pdr(self):
        """Open Pressure Drop Rate module"""
        self.update_status("Opening Pressure Drop Rate module...")
        try:
            pdr_window = tk.Toplevel(self.root)
            PDRModule(pdr_window)
            self.update_status("Pressure Drop Rate module opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDR module: {str(e)}")
            self.update_status("Error opening Pressure Drop Rate module")

    def open_publication_plots(self):
        """Open publication-quality scatter plotter (All_Results_* Excel)."""
        self.update_status("Opening Publication Plots (Scatter)...")
        try:
            win = tk.Toplevel(self.root)
            base = None
            try:
                base = self.foam_manager.get_paper_root_path()
            except Exception:
                base = None
            default_glob = os.path.join(base, "Results", "All_Results_*.xlsx") if base else ""
            PlotModule(win, self.settings, default_all_results_glob=default_glob)
            self.update_status("Publication Plots opened")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open plotting module: {str(e)}")
            self.update_status("Error opening Publication Plots")
    
    def open_dsc_with_foam_check(self):
        """Open DSC Analysis module with foam type verification"""
        self.ensure_foam_type_selected()
        self.open_dsc()

    def open_sem_with_foam_check(self):
        """Open SEM Image Editor module with foam type verification"""
        self.ensure_foam_type_selected()
        self.open_sem()
    def open_oc_with_foam_check(self):
        """Open Open-Cell Content module with foam type verification"""
        self.ensure_foam_type_selected()
        self.open_oc()

    def open_pdr_with_foam_check(self):
        """Open Pressure Drop Rate module with foam type verification"""
        self.ensure_foam_type_selected()
        self.open_pdr()

    def ensure_foam_type_selected(self):
        """Ensure a foam type is selected before opening foam-specific modules"""
        current_foam = self.foam_manager.get_current_foam_type()
        available_foams = self.foam_manager.get_foam_types_for_paper()
        
        if current_foam not in available_foams:
            # Show foam type selection dialog
            fd = FoamTypeDialog(self.root, self.foam_manager)
            self.root.wait_window(fd.top)
            # Update the label
            if hasattr(self, 'foam_type_label'):
                self.foam_type_label.configure(text=f"Foam Type: {self.foam_manager.get_current_foam_type()}")
    
    def open_settings(self):
        """Open settings dialog"""
        try:
            SettingsDialog(self.root, self.settings)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open settings: {str(e)}")
    
    def show_documentation(self):
        """Show documentation"""
        try:
            import webbrowser
            doc_path = os.path.join(os.path.dirname(__file__), "OPTIMIZED_SYSTEM_DOCUMENTATION.md")
            if os.path.exists(doc_path):
                webbrowser.open(f"file://{doc_path}")
            else:
                messagebox.showinfo("Documentation", "Documentation files not found.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open documentation: {str(e)}")
    
    def show_help(self):
        """Open HELP.md with full usage instructions"""
        try:
            import webbrowser
            help_path = os.path.join(os.path.dirname(__file__), "HELP.md")
            if os.path.exists(help_path):
                webbrowser.open(f"file://{help_path}")
            else:
                messagebox.showinfo("Help", "HELP.md no encontrado en la carpeta de la aplicación.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la ayuda: {str(e)}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = """PressTech - Siphony GUI v1.0

An integrated analysis tool for foam processing research.

Features:
• Smart data combination with incremental updates
• Statistical analysis and visualization
• DSC thermal analysis
• SEM image editing
• Open-cell content calculation
• Pressure drop rate analysis
• Cell size and density analysis

Developed for advanced polymer foam research.
"""
        messagebox.showinfo("About PressTech Siphony GUI", about_text)
    
    def prompt_paper_and_foam_startup(self):
        """Show paper selector at startup only"""
        try:
            pd = PaperDialog(self.root, self.foam_manager)
            self.root.wait_window(pd.top)
        except Exception as e:
            print(f"Startup selection error: {e}")
    
    def show_workflow_selection(self):
        """Show first-level dialog: paper-level analysis and link to foam-specific"""
        import tkinter as tk
        from tkinter import ttk
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Analysis Type")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        frame = ttk.Frame(dialog, padding=20)
        frame.grid(row=0, column=0)
        
        ttk.Label(frame, text="Select Analysis Type", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Paper-level analysis
        ttk.Label(frame, text="Paper-Level Analysis:", font=("Arial", 11, "bold")).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        a_btn = ttk.Button(frame, text="📊 Publication Plots", command=lambda: (dialog.destroy(), self.open_analysis()), width=25)
        a_btn.grid(row=2, column=0, padx=(0, 10), pady=5, sticky=(tk.W, tk.E))
        c_btn = ttk.Button(frame, text="⚡ Smart Combine", command=lambda: (dialog.destroy(), self.open_combine()), width=25)
        c_btn.grid(row=2, column=1, padx=(10, 0), pady=5, sticky=(tk.W, tk.E))
        
        ttk.Separator(frame, orient='horizontal').grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)
        
        # Foam-specific link
        ttk.Label(frame, text="Foam-Specific:", font=("Arial", 11, "bold")).grid(row=4, column=0, columnspan=2, sticky=tk.W)
        s_btn = ttk.Button(frame, text="🔬 Specific Analysis", command=lambda: self._open_foam_specific_from(dialog), width=25)
        s_btn.grid(row=5, column=0, columnspan=2, pady=(8, 0))
        
        # Back to main
        back_btn = ttk.Button(frame, text="⬅️ Back", command=dialog.destroy)
        back_btn.grid(row=6, column=0, columnspan=2, pady=(15, 0))
        
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

    def _open_foam_specific_from(self, parent_dialog):
        """Open foam-specific dialog and allow going back to parent dialog."""
        import tkinter as tk
        from tkinter import ttk
        
        parent_dialog.withdraw()
        dialog = tk.Toplevel(self.root)
        dialog.title("Foam-Specific Analysis")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        frame = ttk.Frame(dialog, padding=20)
        frame.grid(row=0, column=0)
        
        ttk.Label(frame, text="Foam-Specific Analysis", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        dsc_btn = ttk.Button(frame, text="🌡️ DSC Analysis", command=lambda: (dialog.destroy(), parent_dialog.destroy(), self.open_dsc_with_foam_check()), width=25)
        dsc_btn.grid(row=1, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        sem_btn = ttk.Button(frame, text="🔬 SEM Image Editor", command=lambda: (dialog.destroy(), parent_dialog.destroy(), self.open_sem_with_foam_check()), width=25)
        sem_btn.grid(row=1, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
        oc_btn = ttk.Button(frame, text="🔓 Open-Cell Content", command=lambda: (dialog.destroy(), parent_dialog.destroy(), self.open_oc_with_foam_check()), width=25)
        oc_btn.grid(row=2, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        pdr_btn = ttk.Button(frame, text="📉 Pressure Drop Rate", command=lambda: (dialog.destroy(), parent_dialog.destroy(), self.open_pdr_with_foam_check()), width=25)
        pdr_btn.grid(row=2, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
        
        # Back: close this and restore parent dialog
        back_btn = ttk.Button(frame, text="⬅️ Back", command=lambda: (dialog.destroy(), parent_dialog.deiconify()))
        back_btn.grid(row=3, column=0, columnspan=2, pady=(12, 0))
        
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def select_workflow_and_close(self, dialog, workflow):
        """Deprecated: retained for compatibility, delegates to new dialog"""
        try:
            dialog.grab_release()
        except Exception:
            pass
        try:
            dialog.destroy()
        except Exception:
            pass
        if workflow == "specific":
            self.show_select_analysis_dialog()
        elif workflow == "analysis":
            self.open_analysis()
        elif workflow == "combine":
            self.open_combine()

    def switch_foam_type(self):
        """Allow switching foam type from the menu"""
        try:
            dlg = FoamTypeDialog(self.root, self.foam_manager)
            self.root.wait_window(dlg.top)
            if hasattr(self, 'foam_type_label'):
                self.foam_type_label.configure(text=f"Foam Type: {self.foam_manager.get_current_foam_type()}")
            self.update_status(f"Foam type set to {self.foam_manager.get_current_foam_type()}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to switch foam type: {str(e)}")

    def switch_paper(self):
        """Allow switching paper from the menu"""
        try:
            dlg = PaperDialog(self.root, self.foam_manager)
            self.root.wait_window(dlg.top)
            if hasattr(self, 'paper_label'):
                self.paper_label.configure(text=f"Paper: {self.foam_manager.get_current_paper()}")
            self.update_status(f"Paper set to {self.foam_manager.get_current_paper()}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to switch paper: {str(e)}")
    def manage_papers(self):
        """Manage the list of papers and their base folders."""
        try:
            dlg = ManagePapersDialog(self.root, self.foam_manager)
            self.root.wait_window(dlg.top)
            if hasattr(self, "paper_label"):
                self.paper_label.configure(text=f"Paper: {self.foam_manager.get_current_paper()}")
            if hasattr(self, "foam_type_label"):
                self.foam_type_label.configure(text=f"Foam Type: {self.foam_manager.get_current_foam_type()}")
            self.update_status("Papers updated")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to manage papers: {str(e)}")



    def open_heatmap(self):
        try:
            HeatmapModule(self.root, self.settings)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open heatmap module: {str(e)}")


    def manage_foams(self):
        """Open Manage Foams dialog from the main window."""
        try:
            current_paper = self.foam_manager.get_current_paper()
            dialog = ManageFoamsDialog(self.root, self.foam_manager, current_paper)
            self.root.wait_window(dialog.top)
            if hasattr(self, "foam_type_label"):
                self.foam_type_label.configure(text=f"Foam Type: {self.foam_manager.get_current_foam_type()}")
            self.update_status(f"Foams updated for {current_paper}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to manage foams: {str(e)}")

    def create_new_paper(self):
        """Create new paper with folder structure and templates"""
        try:
            dialog = NewPaperDialog(self.root, self.foam_manager)
            self.root.wait_window(dialog.top)
            
            if dialog.result and dialog.result.get('paper_name'):
                paper_name = dialog.result['paper_name']
                foam_types = dialog.result['foam_types']
                
                # Update current paper
                self.foam_manager.set_current_paper(paper_name)
                if hasattr(self, 'paper_label'):
                    self.paper_label.configure(text=f"Paper: {paper_name}")
                
                self.update_status(f"New paper '{paper_name}' created with {len(foam_types)} foam types")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create new paper: {str(e)}")
    
    def open_web_tool(self):
        """Open Cell Size and Cell Density web tool"""
        self.update_status("Opening Cell Size & Density analysis tool...")
        try:
            import webbrowser
            webbrowser.open("https://uid.tel.uva.es/poros")
            self.update_status("Cell Size & Density tool opened in browser")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open web tool: {str(e)}")
            self.update_status("Error opening Cell Size & Density tool")
    
    def open_histogram_combiner(self):
        """Open histogram combiner module"""
        try:
            self.update_status("Opening histogram combiner...")
            print("Attempting to import histogram_combiner_module...")  # Debug
            from modules.histogram_combiner_module import HistogramCombinerModule
            print("Import successful, creating combiner...")  # Debug
            combiner = HistogramCombinerModule(self.root, self.foam_manager)
            print("Combiner created, showing window...")  # Debug
            combiner.show()
            print("Window shown successfully")  # Debug
            self.update_status("Histogram combiner opened")
        except ImportError as e:
            error_msg = f"Failed to import histogram combiner module: {str(e)}"
            print(f"Import error: {error_msg}")  # Debug
            messagebox.showerror("Import Error", error_msg)
            self.update_status("Error importing histogram combiner")
        except Exception as e:
            error_msg = f"Failed to open histogram combiner: {str(e)}"
            print(f"General error: {error_msg}")  # Debug
            import traceback
            traceback.print_exc()  # Debug
            messagebox.showerror("Error", error_msg)
            self.update_status("Error opening histogram combiner")
    
    def on_closing(self):
        """Handle application closing"""
        # Save window geometry
        self.settings.set("window_geometry", self.root.geometry())
        self.root.destroy()
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = PressTechGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

