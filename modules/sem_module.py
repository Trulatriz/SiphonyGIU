import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
import numpy as np
import os

class SEMImageEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Editor de Imágenes SEM")
        # Tamaño optimizado para imágenes SEM de 1280x960
        # Ventana ligeramente más grande para incluir controles y márgenes
        self.root.geometry("1480x1100")
        self.root.minsize(1280, 960)

        # Variables de estado
        self.original_image = None
        self.original_filepath = None
        self.current_image = None
        self.unprocessed_image = None  # Imagen sin elementos gráficos añadidos
        self.processed_image = None
        self.display_image = None
        self.scale_factor = 1.0
        self.image_offset_x = 10  # Offset horizontal de la imagen en el canvas
        self.image_offset_y = 10  # Offset vertical de la imagen en el canvas
        self.pixels_per_micron = None
        self.calibration_line = None
        self.line_start = None
        self.line_end = None
        self.drawing_line = False
        self.selecting_region = False
        self.selection_rect = None
        self.selection_start = None

        # Configuración por defecto
        self.border_color = "#FF1493"  # Magenta
        self.border_width = 5  # Borde más fino
        self.scale_length_um = 100
        self.cell_size_enabled = False
        self.cell_size_value = ""
        self.density_overlay_enabled = False
        self.density_mode = "rho_f"  # Options: rho_f, rho_r, expansion
        self.density_value = ""

        # Estado del flujo de trabajo
        self.workflow_step = 0  # 0: cargar, 1: calibrar, 2: recortar, 3: configurar, 4: finalizar

        # Últimas carpetas usadas (se mantienen en memoria durante la sesión)
        self.last_open_dir = None
        self.last_save_dir = None

        # Sistema de deshacer/rehacer
        self.history = []  # Lista de estados (imagen, step, calibration)
        self.history_index = -1  # Índice actual en el historial
        self.max_history = 20  # Máximo número de estados a recordar

        self.setup_ui()
        self.update_workflow_instructions()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Panel de instrucciones
        self.instruction_frame = ttk.LabelFrame(main_frame, text="Instrucciones", padding=10)
        self.instruction_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.instruction_label = ttk.Label(self.instruction_frame, text="", wraplength=800, justify=tk.LEFT)
        self.instruction_label.pack()
        
        # Frame superior con controles
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Botones principales
        self.load_btn = ttk.Button(control_frame, text="1. Abrir Imagen", command=self.load_image)
        self.load_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.calibrate_btn = ttk.Button(control_frame, text="2. Trazar Línea Horizontal", command=self.start_calibration, state=tk.DISABLED)
        self.calibrate_btn.pack(side=tk.LEFT, padx=5)
        
        self.crop_btn = ttk.Button(control_frame, text="3. Seleccionar Región", command=self.start_region_selection, state=tk.DISABLED)
        self.crop_btn.pack(side=tk.LEFT, padx=5)
        
        self.config_btn = ttk.Button(control_frame, text="4. Configurar Elementos", command=self.open_config, state=tk.DISABLED)
        self.config_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(control_frame, text="5. Guardar Imagen", command=self.save_image, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # Separador
        ttk.Separator(control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Botones de deshacer/rehacer
        self.undo_btn = ttk.Button(control_frame, text="↶ Deshacer", command=self.undo, state=tk.DISABLED)
        self.undo_btn.pack(side=tk.LEFT, padx=5)
        
        self.redo_btn = ttk.Button(control_frame, text="↷ Rehacer", command=self.redo, state=tk.DISABLED)
        self.redo_btn.pack(side=tk.LEFT, padx=5)
        
        # Frame para la imagen
        image_frame = ttk.Frame(main_frame)
        image_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas con scrollbars
        canvas_frame = ttk.Frame(image_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white')
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Eventos del canvas
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        # Atajos de teclado para deshacer/rehacer
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-y>", lambda e: self.redo())
        self.root.bind("<Control-Z>", lambda e: self.redo())  # Ctrl+Shift+Z también para rehacer
        
        # Hacer que la ventana pueda recibir focus para los atajos de teclado
        self.root.focus_set()
    
    def update_workflow_instructions(self):
        instructions = [
            "Paso 1: Haga clic en 'Abrir Imagen' para cargar una imagen SEM (.tiff, .png, .jpeg)",
            "Paso 2: Haga clic en 'Trazar Línea Horizontal' y dibuje una línea horizontal sobre un elemento de referencia conocido",
            "Paso 3: Haga clic en 'Seleccionar Región' y arrastre para seleccionar el área de la imagen que desea conservar",
            "Paso 4: Configure los elementos adicionales (color del borde, escala, cell size opcional)",
            "Paso 5: Guarde la imagen procesada en formato TIFF"
        ]
        
        if self.workflow_step < len(instructions):
            self.instruction_label.config(text=instructions[self.workflow_step])
    
    def save_state(self):
        """Guarda el estado actual en el historial para poder deshacerlo"""
        if self.current_image is None:
            return
        
        # Crear estado actual
        state = {
            'current_image': self.current_image.copy() if self.current_image else None,
            'unprocessed_image': self.unprocessed_image.copy() if self.unprocessed_image else None,
            'processed_image': self.processed_image.copy() if self.processed_image else None,
            'workflow_step': self.workflow_step,
            'pixels_per_micron': self.pixels_per_micron,
            'scale_length_um': self.scale_length_um,
            'border_color': self.border_color,
            'border_width': self.border_width,
            'cell_size_enabled': self.cell_size_enabled,
            'cell_size_value': self.cell_size_value,
            'density_overlay_enabled': self.density_overlay_enabled,
            'density_mode': self.density_mode,
            'density_value': self.density_value
        }
        
        # Eliminar estados futuros si estamos en medio del historial
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        
        # Añadir nuevo estado
        self.history.append(state)
        self.history_index += 1
        
        # Limitar tamaño del historial
        if len(self.history) > self.max_history:
            self.history.pop(0)
            self.history_index -= 1
        
        self.update_undo_redo_buttons()
    
    def undo(self):
        """Deshace la última acción"""
        if self.history_index > 0:
            self.history_index -= 1
            self.restore_state(self.history[self.history_index])
            self.update_undo_redo_buttons()
    
    def redo(self):
        """Rehace la siguiente acción"""
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.restore_state(self.history[self.history_index])
            self.update_undo_redo_buttons()
    
    def restore_state(self, state):
        """Restaura un estado del historial"""
        self.current_image = state['current_image'].copy() if state['current_image'] else None
        self.unprocessed_image = state['unprocessed_image'].copy() if state['unprocessed_image'] else None
        self.processed_image = state['processed_image'].copy() if state['processed_image'] else None
        self.workflow_step = state['workflow_step']
        self.pixels_per_micron = state['pixels_per_micron']
        self.scale_length_um = state['scale_length_um']
        self.border_color = state['border_color']
        self.border_width = state['border_width']
        self.cell_size_enabled = state['cell_size_enabled']
        self.cell_size_value = state['cell_size_value']
        self.density_overlay_enabled = state.get('density_overlay_enabled', False)
        self.density_mode = state.get('density_mode', 'rho_f')
        self.density_value = state.get('density_value', '')
        
        # Actualizar interfaz
        self.display_image_on_canvas()
        self.update_workflow_instructions()
        self.update_workflow_buttons()
    
    def update_undo_redo_buttons(self):
        """Actualiza el estado de los botones deshacer/rehacer"""
        can_undo = self.history_index > 0
        can_redo = self.history_index < len(self.history) - 1
        
        self.undo_btn.config(state=tk.NORMAL if can_undo else tk.DISABLED)
        self.redo_btn.config(state=tk.NORMAL if can_redo else tk.DISABLED)
    
    def update_workflow_buttons(self):
        """Actualiza el estado de los botones según el paso del workflow"""
        # Calibración
        self.calibrate_btn.config(state=tk.NORMAL if self.workflow_step >= 1 else tk.DISABLED)
        
        # Recorte
        self.crop_btn.config(state=tk.NORMAL if self.workflow_step >= 2 else tk.DISABLED)
        
        # Configuración
        self.config_btn.config(state=tk.NORMAL if self.workflow_step >= 3 else tk.DISABLED)
        
        # Guardar
        self.save_btn.config(state=tk.NORMAL if self.workflow_step >= 4 else tk.DISABLED)
    
    def load_image(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar imagen SEM",
            initialdir=self.last_open_dir,
            filetypes=[
                ("Archivos de imagen", "*.tiff *.tif *.png *.jpg *.jpeg"),
                ("TIFF files", "*.tiff *.tif"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.original_image = Image.open(file_path)
                # Guardar la ruta del archivo original para usarla al guardar
                self.original_filepath = file_path
                # Actualizar la última carpeta de apertura
                try:
                    self.last_open_dir = os.path.dirname(file_path)
                except Exception:
                    pass
                if self.original_image.mode != 'RGB':
                    self.original_image = self.original_image.convert('RGB')
                
                self.current_image = self.original_image.copy()
                self.unprocessed_image = self.original_image.copy()  # Inicializar imagen sin procesar
                self.display_image_on_canvas()
                
                self.workflow_step = 1
                self.update_workflow_instructions()
                self.update_workflow_buttons()
                
                # Guardar estado inicial
                self.save_state()
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar la imagen: {str(e)}")
    
    def start_region_selection(self):
        if self.pixels_per_micron is None:
            messagebox.showerror("Error", "Flujo de trabajo incorrecto. Primero calibre la escala.")
            return
        
        self.selecting_region = True
        self.selection_rect = None
        self.selection_start = None
        messagebox.showinfo("Selección de Región", "Arrastre para seleccionar la región de la imagen que desea conservar.")
    
    def start_calibration(self):
        if self.current_image is None:
            messagebox.showerror("Error", "Flujo de trabajo incorrecto. Primero cargue una imagen.")
            return
        
        self.drawing_line = True
        self.calibration_line = None
        self.line_start = None
        self.line_end = None
        messagebox.showinfo("Calibración", "Dibuje una línea horizontal sobre un elemento de referencia conocido en la imagen.")
    
    def on_canvas_click(self, event):
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        # Convertir coordenadas del canvas a coordenadas de la imagen
        # Usando el offset dinámico calculado al mostrar la imagen
        img_x = (canvas_x - self.image_offset_x) / self.scale_factor
        img_y = (canvas_y - self.image_offset_y) / self.scale_factor
        
        if self.selecting_region:
            self.selection_start = (img_x, img_y)
            # Limpiar selección anterior si existe
            if self.selection_rect:
                self.canvas.delete(self.selection_rect)
                
        elif self.drawing_line:
            self.line_start = (img_x, img_y)
            # Limpiar línea anterior si existe
            if self.calibration_line:
                self.canvas.delete(self.calibration_line)
    
    def on_canvas_drag(self, event):
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        if self.selecting_region and self.selection_start:
            # Limpiar selección anterior
            if self.selection_rect:
                self.canvas.delete(self.selection_rect)
            
            # Dibujar nuevo rectángulo de selección
            start_x = self.selection_start[0] * self.scale_factor + self.image_offset_x
            start_y = self.selection_start[1] * self.scale_factor + self.image_offset_y
            
            self.selection_rect = self.canvas.create_rectangle(
                start_x, start_y, canvas_x, canvas_y,
                outline="red", width=2, tags="selection"
            )
            
        elif self.drawing_line and self.line_start:
            # Limpiar línea anterior
            if self.calibration_line:
                self.canvas.delete(self.calibration_line)
            
            # Forzar línea horizontal - solo cambiar X, mantener Y del punto inicial
            horizontal_y = self.line_start[1] * self.scale_factor + self.image_offset_y
            start_x = self.line_start[0] * self.scale_factor + self.image_offset_x
            
            self.calibration_line = self.canvas.create_line(
                start_x, horizontal_y, canvas_x, horizontal_y,
                fill="red", width=2, tags="calibration"
            )
    
    def on_canvas_release(self, event):
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        # Convertir coordenadas del canvas a coordenadas de la imagen
        # Usando el offset dinámico calculado al mostrar la imagen
        img_x = (canvas_x - self.image_offset_x) / self.scale_factor
        img_y = (canvas_y - self.image_offset_y) / self.scale_factor
        
        if self.selecting_region and self.selection_start:
            # Completar selección de región
            x1, y1 = self.selection_start
            x2, y2 = img_x, img_y
            
            # Asegurar que x1,y1 sea la esquina superior izquierda
            left = int(min(x1, x2))
            top = int(min(y1, y2))
            right = int(max(x1, x2))
            bottom = int(max(y1, y2))
            
            if right - left > 10 and bottom - top > 10:  # Región mínima válida
                # Recortar la imagen
                self.current_image = self.current_image.crop((left, top, right, bottom))
                self.unprocessed_image = self.current_image.copy()  # Guardar imagen sin procesar
                self.display_image_on_canvas()
                
                # Limpiar selección
                self.canvas.delete("selection")
                self.selecting_region = False
                
                self.workflow_step = 3
                self.update_workflow_instructions()
                self.update_workflow_buttons()
                
                # Guardar estado después del recorte
                self.save_state()
            else:
                messagebox.showwarning("Región muy pequeña", "Por favor, seleccione una región más grande.")
                if self.selection_rect:
                    self.canvas.delete(self.selection_rect)
                self.selection_start = None
                
        elif self.drawing_line and self.line_start:
            # Forzar línea horizontal - solo usar X del punto final, Y del inicial
            self.line_end = (img_x, self.line_start[1])
            
            # Calcular longitud en píxeles (solo horizontal)
            pixel_length = abs(self.line_end[0] - self.line_start[0])
            
            if pixel_length > 5:  # Línea mínima válida
                self.ask_real_length(pixel_length)
            else:
                messagebox.showwarning("Línea muy corta", "Por favor, dibuje una línea más larga.")
                if self.calibration_line:
                    self.canvas.delete(self.calibration_line)
                self.line_start = None
                self.line_end = None
    
    def ask_real_length(self, pixel_length):
        # Ventana para pedir la longitud real
        dialog = tk.Toplevel(self.root)
        dialog.title("Calibración de Escala")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Ingrese la longitud real de la línea trazada:").pack(pady=10)
        
        entry_frame = ttk.Frame(dialog)
        entry_frame.pack(pady=5)
        
        length_var = tk.StringVar()
        entry = ttk.Entry(entry_frame, textvariable=length_var, width=10)
        entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(entry_frame, text="μm").pack(side=tk.LEFT)
        
        def confirm_calibration():
            try:
                real_length = float(length_var.get())
                if real_length <= 0:
                    raise ValueError("La longitud debe ser positiva")
                
                self.pixels_per_micron = pixel_length / real_length
                self.scale_length_um = real_length
                
                dialog.destroy()
                self.drawing_line = False
                
                # Limpiar línea de calibración
                self.canvas.delete("calibration")
                
                # Si aún no se ha recortado, guardar la imagen sin procesar
                if self.unprocessed_image is None:
                    self.unprocessed_image = self.current_image.copy()
                
                self.workflow_step = 2
                self.update_workflow_instructions()
                self.update_workflow_buttons()
                
                # Guardar estado después de la calibración
                self.save_state()
                
                messagebox.showinfo("Calibración", f"Calibración completada: {real_length} μm = {pixel_length:.1f} píxeles")
                
            except ValueError as e:
                messagebox.showerror("Error", "Por favor, ingrese un número válido mayor que 0")
        
        def cancel_calibration():
            dialog.destroy()
            self.drawing_line = False
            if self.calibration_line:
                self.canvas.delete(self.calibration_line)
            self.line_start = None
            self.line_end = None
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Confirmar", command=confirm_calibration).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=cancel_calibration).pack(side=tk.LEFT, padx=5)
        
        entry.focus()
        dialog.bind('<Return>', lambda e: confirm_calibration())
    
    def open_config(self):
        if self.pixels_per_micron is None:
            messagebox.showerror("Error", "Flujo de trabajo incorrecto. Primero complete la calibración y el recorte.")
            return
        
        # Ventana de configuración
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuración de Elementos")
        config_window.geometry("500x480")
        config_window.transient(self.root)
        config_window.grab_set()
        
        # Color del borde
        color_frame = ttk.LabelFrame(config_window, text="Color del Borde y Escala", padding=10)
        color_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.color_display = tk.Label(color_frame, bg=self.border_color, width=10, height=2)
        self.color_display.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(color_frame, text="Elegir Color", command=self.choose_color).pack(side=tk.LEFT)
        
        # Grosor del borde
        border_frame = ttk.Frame(color_frame)
        border_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(border_frame, text="Grosor del borde: ").pack(side=tk.LEFT)
        self.border_width_var = tk.StringVar(value=str(self.border_width))
        border_entry = ttk.Entry(border_frame, textvariable=self.border_width_var, width=5)
        border_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(border_frame, text="píxeles").pack(side=tk.LEFT)
        
        # Cell size opcional
        cellsize_frame = ttk.LabelFrame(config_window, text="Cell Size (Opcional)", padding=10)
        cellsize_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.cellsize_var = tk.BooleanVar(value=self.cell_size_enabled)
        cellsize_check = ttk.Checkbutton(cellsize_frame, text="Añadir cell size (ø)", 
                                       variable=self.cellsize_var, command=self.toggle_cellsize)
        cellsize_check.pack(anchor=tk.W)
        
        cellsize_entry_frame = ttk.Frame(cellsize_frame)
        cellsize_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(cellsize_entry_frame, text="ø = ").pack(side=tk.LEFT)
        self.cellsize_entry_var = tk.StringVar(value=self.cell_size_value)
        self.cellsize_entry = ttk.Entry(cellsize_entry_frame, textvariable=self.cellsize_entry_var, width=10)
        self.cellsize_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(cellsize_entry_frame, text="μm").pack(side=tk.LEFT)
        
        self.toggle_cellsize()

        # Overlay de densidad
        density_frame = ttk.LabelFrame(config_window, text="Densidad (Opcional)", padding=10)
        density_frame.pack(fill=tk.X, padx=10, pady=5)

        self.density_enabled_var = tk.BooleanVar(value=self.density_overlay_enabled)
        density_check = ttk.Checkbutton(
            density_frame,
            text="Añadir etiqueta de densidad",
            variable=self.density_enabled_var,
            command=self.toggle_density_controls
        )
        density_check.pack(anchor=tk.W)

        density_options_frame = ttk.Frame(density_frame)
        density_options_frame.pack(fill=tk.X, pady=(5, 0))

        self.density_mode_var = tk.StringVar(value=self.density_mode)
        self.density_mode_var.trace_add("write", lambda *_: self.update_density_units_label())

        options = [
            ("ρf (kg/m³)", "rho_f"),
            ("ρᵣ", "rho_r"),
            ("X", "expansion"),
        ]
        self.density_radio_buttons = []
        for idx, (label, value) in enumerate(options):
            rb = ttk.Radiobutton(
                density_options_frame,
                text=label,
                value=value,
                variable=self.density_mode_var
            )
            rb.grid(row=0, column=idx, padx=5, sticky=tk.W)
            self.density_radio_buttons.append(rb)

        density_value_frame = ttk.Frame(density_frame)
        density_value_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(density_value_frame, text="Valor:").pack(side=tk.LEFT)
        self.density_value_var = tk.StringVar(value=self.density_value)
        self.density_entry = ttk.Entry(density_value_frame, textvariable=self.density_value_var, width=12)
        self.density_entry.pack(side=tk.LEFT, padx=(5, 5))
        self.density_units_label = ttk.Label(density_value_frame, text="")
        self.density_units_label.pack(side=tk.LEFT)

        self.update_density_units_label()
        self.toggle_density_controls()
        
        # Longitud de escala
        scale_frame = ttk.LabelFrame(config_window, text="Configuración de Escala", padding=10)
        scale_frame.pack(fill=tk.X, padx=10, pady=5)
        
        scale_entry_frame = ttk.Frame(scale_frame)
        scale_entry_frame.pack(fill=tk.X)
        
        ttk.Label(scale_entry_frame, text="Longitud de escala: ").pack(side=tk.LEFT)
        self.scale_entry_var = tk.StringVar(value=str(int(self.scale_length_um)))
        scale_entry = ttk.Entry(scale_entry_frame, textvariable=self.scale_entry_var, width=10)
        scale_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(scale_entry_frame, text="μm").pack(side=tk.LEFT)
        
        # Botones
        button_frame = ttk.Frame(config_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def apply_config():
            try:
                self.scale_length_um = float(self.scale_entry_var.get())
                self.border_width = int(self.border_width_var.get())
                self.cell_size_enabled = self.cellsize_var.get()
                self.cell_size_value = self.cellsize_entry_var.get() if self.cell_size_enabled else ""
                self.density_overlay_enabled = self.density_enabled_var.get()
                self.density_mode = self.density_mode_var.get()
                raw_density_value = self.density_value_var.get().strip()
                
                if self.scale_length_um <= 0:
                    raise ValueError("La longitud de escala debe ser positiva")
                
                if self.border_width < 1:
                    raise ValueError("El grosor del borde debe ser al menos 1 píxel")
                
                if self.cell_size_enabled and not self.cell_size_value:
                    raise ValueError("Ingrese un valor para el cell size")
                
                if self.density_overlay_enabled and not raw_density_value:
                    raise ValueError("Ingrese un valor para la densidad seleccionada")
                
                self.density_value = raw_density_value
                
                config_window.destroy()
                self.apply_final_processing()
                
            except ValueError as e:
                messagebox.showerror("Error", str(e))
        
        ttk.Button(button_frame, text="Aplicar", command=apply_config).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancelar", command=config_window.destroy).pack(side=tk.RIGHT)
    
    def choose_color(self):
        color = colorchooser.askcolor(initialcolor=self.border_color)
        if color[1]:
            self.border_color = color[1]
            self.color_display.config(bg=self.border_color)
    
    def toggle_cellsize(self):
        if self.cellsize_var.get():
            self.cellsize_entry.config(state=tk.NORMAL)
        else:
            self.cellsize_entry.config(state=tk.DISABLED)
    
    def toggle_density_controls(self):
        if not hasattr(self, 'density_enabled_var'):
            return
        enabled = self.density_enabled_var.get()
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, 'density_radio_buttons'):
            for rb in self.density_radio_buttons:
                rb.config(state=state)
        if hasattr(self, 'density_entry'):
            self.density_entry.config(state=state)
        self.update_density_units_label()
    
    def update_density_units_label(self, *_args):
        if hasattr(self, 'density_units_label'):
            mode = self.density_mode_var.get() if hasattr(self, 'density_mode_var') else self.density_mode
            unit = self._density_unit_for_mode(mode)
            if hasattr(self, 'density_enabled_var') and not self.density_enabled_var.get():
                self.density_units_label.config(text="")
            else:
                self.density_units_label.config(text=unit)
    
    def _density_unit_for_mode(self, mode):
        if mode == "rho_f":
            return "kg/m³"
        return ""
    
    def _density_components_for_mode(self, mode):
        if mode == "rho_f":
            return "ρ", "f", "kg/m³"
        if mode == "rho_r":
            return "ρ", "r", ""
        if mode == "expansion":
            return "X", "", ""
        return "", "", ""
    
    def _draw_density_overlay(self, draw, image, font, font_size):
        main_char, sub_char, unit_text = self._density_components_for_mode(self.density_mode)
        value_text = self.density_value.strip()
        if not main_char or not value_text:
            return

        display_label = main_char + (sub_char if sub_char else "")
        measurement_text = f"{display_label} = {value_text}"
        if unit_text:
            measurement_text += f" {unit_text}"

        padding_x = 10
        padding_y = 8
        sub_extra = int(font_size * 0.35) if sub_char else 0
        base_extra = int(font_size * 0.25)
        extra_bottom = base_extra + sub_extra

        text_bbox = draw.textbbox((0, 0), measurement_text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        bg_width = text_width + padding_x * 2
        bg_height = text_height + padding_y * 2 + extra_bottom
        bg_x = self.border_width
        bg_y = self.border_width

        draw.rectangle(
            [bg_x, bg_y, bg_x + bg_width, bg_y + bg_height],
            fill='lightgray',
            outline=self.border_color,
            width=2
        )

        text_x = bg_x + padding_x
        text_y = bg_y + padding_y

        main_bbox = draw.textbbox((0, 0), main_char, font=font)
        main_width = main_bbox[2] - main_bbox[0]
        main_height = main_bbox[3] - main_bbox[1]
        draw.text((text_x, text_y), main_char, fill='black', font=font)

        current_x = text_x + main_width

        if sub_char:
            sub_font_size = max(8, int(font_size * 0.65))
            try:
                sub_font = ImageFont.truetype("arial.ttf", sub_font_size)
            except:
                sub_font = font
            sub_bbox = draw.textbbox((0, 0), sub_char, font=sub_font)
            sub_width = sub_bbox[2] - sub_bbox[0]
            sub_height = sub_bbox[3] - sub_bbox[1]
            sub_y = text_y + max(2, int(main_height * 0.55))
            draw.text((current_x, sub_y), sub_char, fill='black', font=sub_font)
            current_x += sub_width + 4
        else:
            current_x += 4

        rest_text = f"= {value_text}"
        rest_bbox = draw.textbbox((0, 0), rest_text, font=font)
        rest_width = rest_bbox[2] - rest_bbox[0]
        draw.text((current_x, text_y), rest_text, fill='black', font=font)
        current_x += rest_width + 4

        if unit_text:
            draw.text((current_x, text_y), unit_text, fill='black', font=font)
    
    def apply_final_processing(self):
        if self.unprocessed_image is None or self.pixels_per_micron is None:
            return
        
        # Usar la imagen sin procesar como base (sin elementos anteriores)
        img = self.unprocessed_image.copy()
        
        # Añadir borde
        img_with_border = Image.new('RGB', 
                                  (img.width + 2*self.border_width, img.height + 2*self.border_width),
                                  self.border_color)
        img_with_border.paste(img, (self.border_width, self.border_width))
        
        draw = ImageDraw.Draw(img_with_border)
        
        # Configurar fuente
        font_size = max(12, int(min(img.width, img.height) / 40))
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
        
        # Añadir escala en la esquina inferior derecha
        scale_pixels = self.scale_length_um * self.pixels_per_micron
        
        # Calcular dimensiones del texto para ajustar el recuadro
        scale_text = f"{int(self.scale_length_um)} μm"
        text_bbox = draw.textbbox((0, 0), scale_text, font=font)
        text_height = text_bbox[3] - text_bbox[1]
        
        # El recuadro de la escala debe coincidir con los bordes internos de la imagen
        scale_bg_width = scale_pixels + 20  # Padding interno
        scale_bg_height = max(40, text_height + 25)  # Altura aumentada para el texto
        
        # Posición del recuadro de escala (alineado con bordes internos)
        scale_bg_x = img_with_border.width - scale_bg_width - self.border_width
        scale_bg_y = img_with_border.height - scale_bg_height - self.border_width
        
        # Dibujar fondo gris claro para la escala
        draw.rectangle([scale_bg_x, scale_bg_y, 
                       img_with_border.width - self.border_width, img_with_border.height - self.border_width],
                      fill='lightgray', outline=self.border_color, width=2)
        
        # Posición de la línea de escala (centrada en el recuadro)
        line_x = scale_bg_x + (scale_bg_width - scale_pixels) / 2
        line_y = scale_bg_y + scale_bg_height - 12
        
        # Línea de escala horizontal con líneas verticales en los extremos
        draw.line([line_x, line_y, line_x + scale_pixels, line_y],
                 fill='black', width=2)
        
        # Líneas verticales en los extremos
        tick_height = 4
        draw.line([line_x, line_y - tick_height/2, line_x, line_y + tick_height/2],
                 fill='black', width=2)
        draw.line([line_x + scale_pixels, line_y - tick_height/2, 
                  line_x + scale_pixels, line_y + tick_height/2],
                 fill='black', width=2)
        
        # Texto de escala (más separado de la línea)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = line_x + (scale_pixels - text_width) / 2
        text_y = line_y - text_height - 8  # Más separación
        draw.text((text_x, text_y), scale_text, fill='black', font=font)
        
        # Añadir cell size si está habilitado
        if self.cell_size_enabled and self.cell_size_value:
            cellsize_text = f"ø = {self.cell_size_value} μm"
            
            # Posición superior derecha
            cs_bbox = draw.textbbox((0, 0), cellsize_text, font=font)
            cs_width = cs_bbox[2] - cs_bbox[0]
            cs_height = cs_bbox[3] - cs_bbox[1]
            
            # Recuadro que coincide con los bordes internos superior y derecho
            cs_bg_width = cs_width + 20
            cs_bg_height = cs_height + 15  # Altura aumentada para evitar solapamiento
            cs_bg_x = img_with_border.width - cs_bg_width - self.border_width
            cs_bg_y = self.border_width
            
            # Fondo gris claro para cell size
            draw.rectangle([cs_bg_x, cs_bg_y, 
                           img_with_border.width - self.border_width, cs_bg_y + cs_bg_height],
                          fill='lightgray', outline=self.border_color, width=2)
            
            # Texto centrado en el recuadro
            text_x = cs_bg_x + (cs_bg_width - cs_width) / 2
            text_y = cs_bg_y + (cs_bg_height - cs_height) / 2
            draw.text((text_x, text_y), cellsize_text, fill='black', font=font)
        
        if self.density_overlay_enabled and self.density_value:
            self._draw_density_overlay(draw, img_with_border, font, font_size)
        
        # Reemplazar la imagen procesada anterior
        self.processed_image = img_with_border
        self.current_image = img_with_border
        self.display_image_on_canvas()
        
        self.workflow_step = 4
        self.update_workflow_instructions()
        self.update_workflow_buttons()
        
        # Guardar estado después del procesamiento final
        self.save_state()
    
    def save_image(self):
        if self.processed_image is None:
            messagebox.showerror("Error", "Flujo de trabajo incorrecto. Complete todos los pasos anteriores.")
            return
        # Preparar nombre por defecto: nombre_original_edited + extensión original (si existe)
        initialfile = "image_edited.tiff"
        # Preferir la última carpeta usada para guardar; si no existe, usar la del archivo original
        initialdir = self.last_save_dir or (os.path.dirname(self.original_filepath) if self.original_filepath else None)
        default_ext = ".tiff"

        if self.original_filepath:
            try:
                original_base = os.path.splitext(os.path.basename(self.original_filepath))[0]
                original_ext = os.path.splitext(self.original_filepath)[1]
                if original_ext:
                    default_ext = original_ext
                    initialfile = f"{original_base}_edited{original_ext}"
                else:
                    initialfile = f"{original_base}_edited{default_ext}"

                # Si no hay last_save_dir, usar la carpeta del archivo original
                if not self.last_save_dir:
                    initialdir = os.path.dirname(self.original_filepath)
            except Exception:
                # Fallback si ocurre algún error al parsear la ruta
                initialfile = "image_edited.tiff"
                default_ext = ".tiff"

        file_path = filedialog.asksaveasfilename(
            title="Guardar imagen procesada",
            initialfile=initialfile,
            initialdir=initialdir,
            defaultextension=default_ext,
            filetypes=[
                ("TIFF files", "*.tiff *.tif"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("Todos los archivos", "*.*")
            ]
        )

        if file_path:
            try:
                # Guardar con parámetros por defecto (calidad y dpi)
                self.processed_image.save(file_path, quality=95, dpi=(300, 300))

                # Actualizar la última carpeta de guardado
                try:
                    self.last_save_dir = os.path.dirname(file_path)
                except Exception:
                    pass

                messagebox.showinfo("Éxito", f"Imagen guardada en: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar la imagen: {str(e)}")
    
    def display_image_on_canvas(self):
        if self.current_image is None:
            return
        
        # Calcular escala para mostrar la imagen de forma óptima
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            self.root.after(100, self.display_image_on_canvas)
            return
        
        # Para imágenes SEM de 1280x960, optimizar la visualización
        img_width = self.current_image.width
        img_height = self.current_image.height
        
        # Para imágenes de exactamente 1280x960, mostrar a escala 1:1 con mínimo margen
        if img_width == 1280 and img_height == 960:
            # Mostrar a escala 1:1 exacta
            self.scale_factor = 1.0
            display_width = 1280
            display_height = 960
        else:
            # Para otras dimensiones, calcular escala para que quepa con margen mínimo
            max_display_width = canvas_width - 20  # Solo 10px de margen a cada lado
            max_display_height = canvas_height - 20  # Solo 10px de margen arriba y abajo
            
            scale_x = max_display_width / img_width
            scale_y = max_display_height / img_height
            
            self.scale_factor = min(scale_x, scale_y, 1.0)  # Máximo 1:1
            display_width = int(img_width * self.scale_factor)
            display_height = int(img_height * self.scale_factor)
        
        # Redimensionar imagen para mostrar
        display_img = self.current_image.resize((display_width, display_height), Image.Resampling.LANCZOS)
        self.display_image = ImageTk.PhotoImage(display_img)
        
        # Centrar la imagen en el canvas con mínimo offset
        offset_x = max(5, (canvas_width - display_width) // 2)
        offset_y = max(5, (canvas_height - display_height) // 2)
        
        # Guardar el offset para las conversiones de coordenadas
        self.image_offset_x = offset_x
        self.image_offset_y = offset_y
        
        # Limpiar canvas y mostrar imagen centrada
        self.canvas.delete("image")
        self.canvas.create_image(offset_x, offset_y, anchor=tk.NW, image=self.display_image, tags="image")
        
        # Configurar región de scroll para incluir toda la imagen con mínimo margen
        scroll_region = (0, 0, offset_x + display_width + 10, offset_y + display_height + 10)
        self.canvas.configure(scrollregion=scroll_region)


class SEMModule:
    """Wrapper class for the SEM Image Editor to match the expected interface"""
    def __init__(self, root):
        # Create the SEM Image Editor directly
        self.editor = SEMImageEditor(root)


def main():
    root = tk.Tk()
    app = SEMImageEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
