import os
import re
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# --- Diccionario de Unidades y Factores de Conversión a Base ---
# Base: Amperios (A) y Voltios (V)
UNIT_FACTORS = {
    "I": {
        "A": 1.0,
        "mA": 1e-3,
        "uA": 1e-6,
        "μA": 1e-6,
        "nA": 1e-9,
        "pA": 1e-12
    },
    "V": {
        "V": 1.0,
        "mV": 1e-3,
        "uV": 1e-6,
        "μV": 1e-6
    }
}

class IVCurveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Curvas IV")
        self.root.geometry("1400x900")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.current_folder = os.path.abspath("datos")
        self.excel_data = [] # Lista de dicts con info de cada excel
        self.summary_data = None
        self.valid_files = set()
        self.selected_files = set()
        
        self.default_i_unit = "μA"
        self.default_v_unit = "V"
        
        self.fig, self.ax = plt.subplots(figsize=(7, 5))
        self.canvas = None
        self.toolbar = None
        self.annot = None
        self.hover_cid = None
        # Flags para control de generación bajo demanda
        self.needs_regen = False
        self.graph_generated_iv = False
        self.graph_generated_isc = False
        
        self.setup_ui()
        self.scan_folder(self.current_folder)

    def setup_ui(self):
        # Frame Principal
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- PANEL IZQUIERDO (Controles y Lista) ---
        left_panel = ttk.Frame(main_frame, width=300)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # Selección de Carpeta
        folder_frame = ttk.LabelFrame(left_panel, text="Carpeta de Datos")
        folder_frame.pack(fill=tk.X, pady=5)
        
        self.lbl_folder = ttk.Label(folder_frame, text=self.current_folder, wraplength=250)
        self.lbl_folder.pack(pady=5, padx=5)
        
        btn_browse = ttk.Button(folder_frame, text="Seleccionar Carpeta", command=self.browse_folder)
        btn_browse.pack(pady=5)
        
        btn_refresh = ttk.Button(folder_frame, text="Actualizar / Escanear", command=lambda: self.scan_folder(self.current_folder))
        btn_refresh.pack(pady=5)
        
        # Resumen
        self.lbl_summary = ttk.Label(left_panel, text="Buscando...")
        self.lbl_summary.pack(pady=5)
        
        # Lista de Archivos
        list_frame = ttk.LabelFrame(left_panel, text="Archivos Encontrados")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Treeview para la lista con selección manual
        columns = ("Sel", "Estado", "Archivo", "Unidad I", "Unidad V")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="none")
        self.tree.heading("Sel", text="Sel")
        self.tree.heading("Estado", text="")
        self.tree.heading("Archivo", text="Archivo")
        self.tree.heading("Unidad I", text="I")
        self.tree.heading("Unidad V", text="V")
        self.tree.column("Sel", width=40, anchor=tk.CENTER)
        self.tree.column("Estado", width=30, anchor=tk.CENTER)
        self.tree.column("Archivo", width=120)
        self.tree.column("Unidad I", width=50, anchor=tk.CENTER)
        self.tree.column("Unidad V", width=50, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree.bind("<Button-1>", self.on_tree_click)

        select_buttons = ttk.Frame(list_frame)
        select_buttons.pack(fill=tk.X, padx=5, pady=(0,5))
        self.btn_select_all = ttk.Button(select_buttons, text="Seleccionar todos", command=self.toggle_select_all, state=tk.DISABLED)
        self.btn_select_all.pack(fill=tk.X)
        
        # Unidades a mostrar
        units_frame = ttk.LabelFrame(left_panel, text="Unidades del Gráfico")
        units_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(units_frame, text="Intensidad:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_i = ttk.Combobox(units_frame, values=list(UNIT_FACTORS["I"].keys()), width=5, state="readonly")
        self.cb_unit_i.set(self.default_i_unit)
        self.cb_unit_i.grid(row=0, column=1, padx=5, pady=5)
        self.cb_unit_i.bind("<<ComboboxSelected>>", lambda e: setattr(self, 'needs_regen', True))
        
        ttk.Label(units_frame, text="Voltaje:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_v = ttk.Combobox(units_frame, values=list(UNIT_FACTORS["V"].keys()), width=5, state="readonly")
        self.cb_unit_v.set(self.default_v_unit)
        self.cb_unit_v.grid(row=1, column=1, padx=5, pady=5)
        self.cb_unit_v.bind("<<ComboboxSelected>>", lambda e: setattr(self, 'needs_regen', True))
        
        # Botón Generar
        self.btn_generate = ttk.Button(left_panel, text="Generar Gráfico IV", command=self.generate_plot, state=tk.DISABLED)
        self.btn_generate.pack(fill=tk.X, pady=10)

        # --- PANEL DERECHO (Gráfico y Controles de Gráfico) ---
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Canvas de matplotlib
        self.notebook = ttk.Notebook(right_panel)
        self.iv_tab = ttk.Frame(self.notebook)
        self.isc_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.iv_tab, text="IV")
        self.notebook.add(self.isc_tab, text="Isc")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.iv_tab)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.iv_tab)
        self.toolbar.update()
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.on_tab_change())
        
        # Controles de Gráfico (Limites y ejes)
        plot_ctrl_frame = ttk.LabelFrame(right_panel, text="Controles del Gráfico (Tras generar)")
        plot_ctrl_frame.pack(fill=tk.X, pady=5)
        
        # Limites X
        ttk.Label(plot_ctrl_frame, text="X min:").grid(row=0, column=0, padx=2)
        self.entry_xmin = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_xmin.grid(row=0, column=1, padx=2)
        
        ttk.Label(plot_ctrl_frame, text="X max:").grid(row=0, column=2, padx=2)
        self.entry_xmax = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_xmax.grid(row=0, column=3, padx=2)
        
        # Limites Y
        ttk.Label(plot_ctrl_frame, text="Y min:").grid(row=1, column=0, padx=2, pady=5)
        self.entry_ymin = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_ymin.grid(row=1, column=1, padx=2)
        
        ttk.Label(plot_ctrl_frame, text="Y max:").grid(row=1, column=2, padx=2)
        self.entry_ymax = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_ymax.grid(row=1, column=3, padx=2)
        
        btn_apply_lims = ttk.Button(plot_ctrl_frame, text="Aplicar", command=self.apply_limits)
        btn_apply_lims.grid(row=0, column=4, rowspan=2, padx=10)
        
        btn_reset_lims = ttk.Button(plot_ctrl_frame, text="Reiniciar Límites", command=self.reset_limits)
        btn_reset_lims.grid(row=0, column=5, rowspan=2, padx=5)
        
        # Checkboxes Ejes
        self.var_swap_axes = tk.BooleanVar(value=False)
        cb_swap = ttk.Checkbutton(plot_ctrl_frame, text="Intercambiar X/Y", variable=self.var_swap_axes, command=lambda: setattr(self, 'needs_regen', True))
        cb_swap.grid(row=0, column=6, padx=10)
        
        self.var_invert_x = tk.BooleanVar(value=False)
        cb_inv_x = ttk.Checkbutton(plot_ctrl_frame, text="Invertir Eje X", variable=self.var_invert_x, command=self.update_axes_direction)
        cb_inv_x.grid(row=1, column=6, padx=10, sticky=tk.W)
        
        self.var_invert_y = tk.BooleanVar(value=False)
        cb_inv_y = ttk.Checkbutton(plot_ctrl_frame, text="Invertir Eje Y", variable=self.var_invert_y, command=self.update_axes_direction)
        cb_inv_y.grid(row=2, column=6, padx=10, sticky=tk.W)
        
        # Estilo de línea
        ttk.Label(plot_ctrl_frame, text="Estilo de curva:").grid(row=0, column=7, padx=10, sticky=tk.E)
        self.cb_line_style = ttk.Combobox(plot_ctrl_frame, values=["Línea continua (-)", "Línea rayada (--)", "Puntos (.)", "Puntos y línea (.-)"], width=15, state="readonly")
        self.cb_line_style.set("Línea continua (-)")
        self.cb_line_style.grid(row=0, column=8, padx=5, sticky=tk.W)
        self.cb_line_style.bind("<<ComboboxSelected>>", lambda e: setattr(self, 'needs_regen', True))
        
        # Exportar
        export_frame = ttk.Frame(plot_ctrl_frame)
        export_frame.grid(row=1, column=7, columnspan=2, rowspan=2, padx=20)
        btn_export = ttk.Button(export_frame, text="Exportar Gráfico", command=self.export_plot)
        btn_export.pack(side=tk.LEFT)
        btn_export_excel = ttk.Button(export_frame, text="Exportar Excel combinado", command=self.export_combined_excel)
        btn_export_excel.pack(side=tk.LEFT, padx=(10, 0))
        self.btn_export_excel = btn_export_excel

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.current_folder)
        if folder:
            self.current_folder = folder
            self.lbl_folder.config(text=self.current_folder)
            self.scan_folder(self.current_folder)
            
    def refresh_selection_columns(self):
        for item_id in self.tree.get_children():
            values = list(self.tree.item(item_id, "values"))
            filename = values[2]
            if filename in self.valid_files:
                values[0] = "✓" if filename in self.selected_files else ""
            else:
                values[0] = ""
            self.tree.item(item_id, values=values)

    def update_select_all_button(self):
        if not self.valid_files:
            self.btn_select_all.config(text="Seleccionar todos", state=tk.DISABLED)
            return
        if self.selected_files >= self.valid_files:
            self.btn_select_all.config(text="Deseleccionar todos", state=tk.NORMAL)
        else:
            self.btn_select_all.config(text="Seleccionar todos", state=tk.NORMAL)

    def update_generate_button_state(self):
        if self.selected_files:
            self.btn_generate.config(state=tk.NORMAL)
        else:
            self.btn_generate.config(state=tk.DISABLED)

    def toggle_select_all(self):
        if self.selected_files >= self.valid_files:
            self.selected_files.clear()
        else:
            self.selected_files = set(self.valid_files)
        self.refresh_selection_columns()
        self.update_select_all_button()
        self.update_generate_button_state()

    def on_tree_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id:
            return
        if col != "#1":
            return
        values = list(self.tree.item(row_id, "values"))
        if not values:
            return
        filename = values[2]
        if filename not in self.valid_files:
            return
        if filename in self.selected_files:
            self.selected_files.remove(filename)
        else:
            self.selected_files.add(filename)
        self.refresh_selection_columns()
        self.update_select_all_button()
        self.update_generate_button_state()

    def get_unit_from_str(self, s, var_type):
        """Busca unidades en paréntesis ej: (mA), (V)."""
        match = re.search(r'\((.*?)\)', s)
        if match:
            u = match.group(1).strip()
            if u in UNIT_FACTORS[var_type]:
                return u
        # Por defecto si no encuentra, asume la base
        return "A" if var_type == "I" else "V"

    def scan_folder(self, folder):
        self.excel_data = []
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not os.path.exists(folder):
            self.lbl_summary.config(text="Carpeta no encontrada.")
            self.btn_generate.config(state=tk.DISABLED)
            return
            
        files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        valid_count = 0
        invalid_count = 0
        
        self.summary_data = []
        for file in files:
            path = os.path.join(folder, file)
            try:
                # Intentar buscar la cabecera dinámicamente en las primeras 50 filas de la hoja datos_IV
                df_preview = pd.read_excel(path, sheet_name="datos_IV", header=None, nrows=50)
                
                header_row = None
                i_cols_idx = []
                v_cols_idx = []
                summary_df = None
                
                for idx, row in df_preview.iterrows():
                    row_str = [str(x).strip().upper() if pd.notna(x) else "" for x in row]
                    
                    curr_i = []
                    curr_v = []
                    for col_idx, val in enumerate(row_str):
                        # Criterio de búsqueda para I:
                        is_i = False
                        if val in ["I", "INTENSIDAD", "CURRENT"]:
                            is_i = True
                        elif "I (" in val or "INTENSIDAD (" in val or "CURRENT (" in val:
                            is_i = True
                        elif val.startswith("I_") or val.startswith("CURRENT_"):
                            is_i = True
                        
                        # Criterio de búsqueda para V:
                        is_v = False
                        if val in ["V", "VOLTAJE", "VOLTAGE"]:
                            is_v = True
                        elif "V (" in val or "VOLTAJE (" in val or "VOLTAGE (" in val:
                            is_v = True
                        elif val.startswith("V_") or val.startswith("VOLTAGE_"):
                            is_v = True
                            
                        if is_i:
                            curr_i.append(col_idx)
                        if is_v:
                            curr_v.append(col_idx)
                            
                    if curr_i and curr_v:
                        header_row = idx
                        i_cols_idx = curr_i
                        v_cols_idx = curr_v
                        break
                
                status = "✅"
                i_unit = "-"
                v_unit = "-"
                i_col_name = None
                v_col_name = None
                
                if header_row is None:
                    status = "❌"
                    invalid_count += 1
                else:
                    # Recuperar nombres reales de las columnas en la fila detectada
                    row_raw = list(df_preview.iloc[header_row])
                    i_col_name = row_raw[i_cols_idx[0]]
                    v_col_name = row_raw[v_cols_idx[0]]
                    
                    if len(i_cols_idx) > 1 or len(v_cols_idx) > 1:
                        status = "⚠️"
                    
                    i_unit = self.get_unit_from_str(str(i_col_name), "I")
                    v_unit = self.get_unit_from_str(str(v_col_name), "V")
                    valid_count += 1
                    summary_df = self.read_summary_sheet(path)
                
                self.excel_data.append({
                    "filename": file,
                    "path": path,
                    "status": status,
                    "header_row": header_row,
                    "i_col": i_col_name,
                    "v_col": v_col_name,
                    "i_unit": i_unit,
                    "v_unit": v_unit,
                    "summary_df": summary_df
                })
                if status in ["✅", "⚠️"] and summary_df is not None:
                    self.summary_data.append({"filename": file, "data": summary_df})
                
                self.tree.insert("", "end", values=("", status, file, i_unit, v_unit))
                
            except Exception as e:
                print(f"Error reading {file}: {e}")
                self.tree.insert("", "end", values=("", "❌", file, "-", "-"))
                invalid_count += 1

        self.lbl_summary.config(text=f"Total: {len(files)} | Válidos: {valid_count} | Inválidos: {invalid_count}")

        self.valid_files = {item['filename'] for item in self.excel_data if item['status'] in ['✅', '⚠️']}
        self.selected_files.intersection_update(self.valid_files)
        self.refresh_selection_columns()
        self.update_select_all_button()
        self.update_generate_button_state()

        if valid_count > 0:
            self.btn_export_excel.config(state=tk.NORMAL)
            self.btn_select_all.config(state=tk.NORMAL)
        else:
            self.btn_export_excel.config(state=tk.DISABLED)
            self.btn_select_all.config(state=tk.DISABLED)

    def generate_plot(self):
        # Generar IV (y preparar Isc) sólo cuando se pulsa el botón
        if not self.selected_files:
            messagebox.showwarning('Advertencia', 'Seleccione al menos un archivo valido para generar el gráfico.')
            return

        self.needs_regen = False
        self.graph_generated_iv = True

        # Construir combinados para uso posterior (Isc)
        try:
            self.last_combined_iv = self.build_combined_iv()
        except Exception:
            self.last_combined_iv = pd.DataFrame()
        try:
            self.last_combined_summary = self.build_combined_summary()
        except Exception:
            self.last_combined_summary = pd.DataFrame()

        # Marcar que Isc está disponible si hay resumen
        self.graph_generated_isc = not self.last_combined_summary.empty

        self.ax.clear()
        if self.hover_cid:
            self.canvas.mpl_disconnect(self.hover_cid)
            
        self.annot = self.ax.annotate("", xy=(0,0), xytext=(20,20), textcoords="offset points",
                                      bbox=dict(boxstyle="round", fc="w", alpha=0.9),
                                      arrowprops=dict(arrowstyle="->"))
        self.annot.set_visible(False)
        
        target_i_unit = self.cb_unit_i.get()
        target_v_unit = self.cb_unit_v.get()
        
        swap = self.var_swap_axes.get()
        
        style_map = {
            "Línea continua (-)": "-",
            "Línea rayada (--)": "--",
            "Puntos (.)": ".",
            "Puntos y línea (.-)": ".-"
        }
        fmt = style_map.get(self.cb_line_style.get(), "-")
        
        lines = []
        labels = []
        
        selected_items = [item for item in self.excel_data if item['filename'] in self.selected_files and item['status'] in ['✅', '⚠️']]
        for data in selected_items:
            try:
                df = pd.read_excel(data["path"], sheet_name="datos_IV", header=data["header_row"])
                # Convertimos al factor base (A o V) y luego a la unidad destino
                base_i_factor = UNIT_FACTORS["I"].get(data["i_unit"], 1.0)
                base_v_factor = UNIT_FACTORS["V"].get(data["v_unit"], 1.0)
                
                target_i_factor = UNIT_FACTORS["I"][target_i_unit]
                target_v_factor = UNIT_FACTORS["V"][target_v_unit]
                
                # Multiplicador = (Unidad origen a base) / (Base a unidad destino)
                mult_i = base_i_factor / target_i_factor
                mult_v = base_v_factor / target_v_factor
                
                i_vals = pd.to_numeric(df[data["i_col"]], errors='coerce').dropna() * mult_i
                v_vals = pd.to_numeric(df[data["v_col"]], errors='coerce').dropna() * mult_v
                
                # Asegurarnos de que tengan la misma longitud
                min_len = min(len(i_vals), len(v_vals))
                i_vals = i_vals.iloc[:min_len]
                v_vals = v_vals.iloc[:min_len]
                
                x_vals = v_vals if not swap else i_vals
                y_vals = i_vals if not swap else v_vals
                
                line, = self.ax.plot(x_vals, y_vals, fmt, label=data["filename"])
                lines.append(line)
                    
                # Guardar info para el tooltip
                labels.append(f"Archivo: {data['filename']}\nCol V: {data['v_col']}\nCol I: {data['i_col']}")
                    
            except Exception as e:
                print(f"Error ploting {data['filename']}: {e}")
                    
        x_label = f"Voltaje ({target_v_unit})" if not swap else f"Intensidad ({target_i_unit})"
        y_label = f"Intensidad ({target_i_unit})" if not swap else f"Voltaje ({target_v_unit})"
        
        self.ax.set_xlabel(x_label)
        self.ax.set_ylabel(y_label)
        self.ax.set_title("Curvas IV")
        self.ax.grid(True)
        
        # Configurar Tooltips manuales
        def hover(event):
            vis = self.annot.get_visible()
            if event.inaxes == self.ax:
                for i, line in enumerate(lines):
                    cont, ind = line.contains(event)
                    if cont:
                        x, y = line.get_data()
                        # Si es un clic muy cercano
                        idx = ind["ind"][0]
                        self.annot.xy = (x[idx], y[idx])
                        self.annot.set_text(f"{labels[i]}\nX: {x[idx]:.4g}\nY: {y[idx]:.4g}")
                        self.annot.set_visible(True)
                        self.canvas.draw_idle()
                        return
            if vis:
                self.annot.set_visible(False)
                self.canvas.draw_idle()

        self.hover_cid = self.canvas.mpl_connect("motion_notify_event", hover)
        
        self.update_axes_direction()
        
        # Limpiar Entry de limites para reflejar ajuste automático
        self.entry_xmin.delete(0, tk.END)
        self.entry_xmax.delete(0, tk.END)
        self.entry_ymin.delete(0, tk.END)
        self.entry_ymax.delete(0, tk.END)
        
        self.canvas.draw()

        # flags ya actualizados al inicio
        
    def update_axes_direction(self):
        if not self.ax.lines:
            return
            
        xlim = self.ax.get_xlim()
        if self.var_invert_x.get():
            if xlim[0] < xlim[1]:
                self.ax.set_xlim(xlim[1], xlim[0])
        else:
            if xlim[0] > xlim[1]:
                self.ax.set_xlim(xlim[1], xlim[0])
                
        ylim = self.ax.get_ylim()
        if self.var_invert_y.get():
            if ylim[0] < ylim[1]:
                self.ax.set_ylim(ylim[1], ylim[0])
        else:
            if ylim[0] > ylim[1]:
                self.ax.set_ylim(ylim[1], ylim[0])
                
        self.canvas.draw()
        
    def apply_limits(self):
        try:
            xmin = float(self.entry_xmin.get()) if self.entry_xmin.get() else None
            xmax = float(self.entry_xmax.get()) if self.entry_xmax.get() else None
            ymin = float(self.entry_ymin.get()) if self.entry_ymin.get() else None
            ymax = float(self.entry_ymax.get()) if self.entry_ymax.get() else None
            
            if xmin is not None and xmax is not None:
                self.ax.set_xlim(xmin, xmax)
            elif xmin is not None:
                self.ax.set_xlim(left=xmin)
            elif xmax is not None:
                self.ax.set_xlim(right=xmax)
                
            if ymin is not None and ymax is not None:
                self.ax.set_ylim(ymin, ymax)
            elif ymin is not None:
                self.ax.set_ylim(bottom=ymin)
            elif ymax is not None:
                self.ax.set_ylim(top=ymax)
                
            self.update_axes_direction() # Mantener inversión
            self.canvas.draw()
        except ValueError:
            messagebox.showerror("Error", "Los límites deben ser valores numéricos.")

    def reset_limits(self):
        # Reinicia límites y regenera sólo si el gráfico correspondiente ya fue generado
        current = self.notebook.tab(self.notebook.select(), "text") if hasattr(self, 'notebook') else 'IV'
        if current == 'IV' and self.graph_generated_iv:
            self.generate_plot()
        elif current == 'Isc' and self.graph_generated_isc:
            self.generate_isc_plot()
        else:
            self.ax.clear()
            self.ax.set_title('Pulse "Generar Gráfico IV" para generar los gráficos')
            self.canvas.draw()

        self.entry_xmin.delete(0, tk.END)
        self.entry_xmax.delete(0, tk.END)
        self.entry_ymin.delete(0, tk.END)
        self.entry_ymax.delete(0, tk.END)

    def get_clean_column_name(self, col_name):
        if not isinstance(col_name, str):
            return str(col_name)
        return ''.join(c for c in col_name.lower() if c.isalnum())

    def read_summary_sheet(self, path):
        try:
            df = pd.read_excel(path, sheet_name="resumen")
        except Exception:
            return None

        def clean(s):
            if not isinstance(s, str):
                return ''
            return ''.join(ch for ch in s.lower() if ch.isalnum())

        fecha_col = None
        isc_col = None
        for col in df.columns:
            col_clean = clean(col)
            if 'fecha' in col_clean and 'inicio' in col_clean:
                fecha_col = col
            if col_clean == 'isc' or 'isc' in col_clean:
                isc_col = col

        if fecha_col is None or isc_col is None:
            return None

        result = pd.DataFrame({
            'Fecha y hora inicio': pd.to_datetime(df[fecha_col], errors='coerce'),
            'Isc': pd.to_numeric(df[isc_col], errors='coerce')
        })
        return result.dropna(subset=['Fecha y hora inicio', 'Isc'])

    def same_voltage_series(self, s1, s2):
        if len(s1) != len(s2):
            return False
        s1_arr = np.asarray(s1, dtype=float)
        s2_arr = np.asarray(s2, dtype=float)
        return np.allclose(s1_arr, s2_arr, atol=1e-9, rtol=1e-6)

    def build_combined_iv(self):
        combined_columns = []
        top_headers = []
        bottom_headers = []
        data_dict = {}

        unique_voltages = []
        voltage_keys = []

        selected_items = [item for item in self.excel_data if item['filename'] in self.selected_files and item['status'] in ['✅', '⚠️']]
        for item in selected_items:
            try:
                df = pd.read_excel(item['path'], sheet_name='datos_IV', header=item['header_row'])
                v_values = pd.to_numeric(df[item['v_col']], errors='coerce').dropna().reset_index(drop=True)
                i_values = pd.to_numeric(df[item['i_col']], errors='coerce').dropna().reset_index(drop=True)
                length = min(len(v_values), len(i_values))
                v_values = v_values.iloc[:length]
                i_values = i_values.iloc[:length]
            except Exception:
                continue

            matched_index = None
            for idx, existing_v in enumerate(unique_voltages):
                if self.same_voltage_series(existing_v, v_values):
                    matched_index = idx
                    break

            if matched_index is None:
                unique_voltages.append(v_values)
                key = f'V_{len(unique_voltages)}'
                voltage_keys.append(key)
                combined_columns.append(key)
                top_headers.append(item['filename'])
                bottom_headers.append('V')
                data_dict[key] = v_values

                key_i = f'I_{len(unique_voltages)}_{item["filename"]}'
                combined_columns.append(key_i)
                top_headers.append(item['filename'])
                bottom_headers.append('I')
                data_dict[key_i] = i_values
            else:
                key_i = f'I_{matched_index + 1}_{item["filename"]}'
                combined_columns.append(key_i)
                top_headers.append(item['filename'])
                bottom_headers.append('I')
                data_dict[key_i] = i_values

        max_len = max((len(col) for col in data_dict.values()), default=0)
        for key, series in data_dict.items():
            if len(series) < max_len:
                data_dict[key] = series.reindex(range(max_len))

        if not combined_columns:
            return pd.DataFrame()

        df_combined = pd.DataFrame({key: data_dict[key] for key in combined_columns})
        df_combined.columns = pd.MultiIndex.from_arrays([top_headers, bottom_headers])
        return df_combined

    def build_combined_summary(self):
        rows = []
        for item in self.excel_data:
            if item['filename'] not in self.selected_files or item['status'] not in ['✅', '⚠️']:
                continue
            summary_df = item.get('summary_df')
            if summary_df is None or summary_df.empty:
                continue
            temp = summary_df.copy()
            temp['Archivo'] = item['filename']
            rows.append(temp)

        if not rows:
            return pd.DataFrame(columns=['Archivo', 'Fecha y hora inicio', 'Isc'])

        return pd.concat(rows, ignore_index=True)

    def export_combined_excel(self):
        combined_iv = self.build_combined_iv()
        combined_summary = self.build_combined_summary()

        if combined_iv.empty:
            messagebox.showwarning('Advertencia', 'No hay datos IV combinados para exportar.')
            return

        folder = self.current_folder
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            initialfile='iv_combinado.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel Workbook', '*.xlsx'), ('All Files', '*.*')]
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # Si las columnas son MultiIndex (top_headers/bottom_headers), aplanarlas para Excel
                df_to_write = combined_iv.copy()
                if isinstance(df_to_write.columns, pd.MultiIndex):
                    new_cols = []
                    for a, b in df_to_write.columns:
                        a_str = str(a)
                        b_str = str(b)
                        # Nombre legible: Archivo_Tipo (ej. datos.xlsx_V)
                        new_cols.append(f"{a_str}_{b_str}")
                    df_to_write.columns = new_cols
                df_to_write.to_excel(writer, sheet_name='datos_IV', index=False)
                if not combined_summary.empty:
                    combined_summary[['Fecha y hora inicio', 'Isc']].to_excel(writer, sheet_name='resumen', index=False)
            messagebox.showinfo('Éxito', f'Archivo exportado a:\n{save_path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar el archivo:\n{e}')

    def export_plot(self):
        # Exporta el gráfico correspondiente a la pestaña activa con nombre por defecto
        current = self.notebook.tab(self.notebook.select(), "text") if hasattr(self, 'notebook') else 'IV'
        default_name = 'grafico_iv.png' if current == 'IV' else 'grafico_isc.png'

        if current == 'IV' and not self.graph_generated_iv:
            messagebox.showwarning('Advertencia', 'No hay gráfico IV generado para exportar.')
            return
        if current == 'Isc' and not self.graph_generated_isc:
            messagebox.showwarning('Advertencia', 'No hay gráfico Isc generado para exportar.')
            return

        folder = self.current_folder
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            initialfile=default_name,
            defaultextension='.png',
            filetypes=[('PNG Image', '*.png'), ('PDF', '*.pdf'), ('All Files', '*.*')]
        )
        if not save_path:
            return
        try:
            # Asegurar que el contenido de la figura corresponde a la pestaña
            if current == 'Isc':
                # regenerar la gráfica Isc en los ejes actuales
                self.generate_isc_plot()
            else:
                self.generate_plot()
            self.fig.savefig(save_path, dpi=300, bbox_inches='tight')
            messagebox.showinfo('Éxito', f'Gráfico exportado a:\n{save_path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar el gráfico:\n{e}')

    def on_tab_change(self):
        # Recrea el canvas en la pestaña activa y dibuja el gráfico correspondiente
        selected = self.notebook.tab(self.notebook.select(), "text")
        parent = self.iv_tab if selected == 'IV' else self.isc_tab

        # destruir toolbar y canvas actuales y recrearlos bajo el nuevo padre
        try:
            self.toolbar.destroy()
        except Exception:
            pass
        try:
            self.canvas.get_tk_widget().pack_forget()
        except Exception:
            pass

        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, parent)
        self.toolbar.update()

        # Mostrar gráfico sólo si fue generado con el botón
        if selected == 'IV':
            if self.graph_generated_iv:
                self.generate_plot()
            else:
                self.ax.clear()
                self.ax.set_title('Pulse "Generar Gráfico IV" para generar los gráficos')
                self.canvas.draw()
        else:
            if self.graph_generated_isc:
                self.generate_isc_plot()
            else:
                self.ax.clear()
                self.ax.set_title('Pulse "Generar Gráfico IV" para generar los gráficos')
                self.canvas.draw()

    def generate_isc_plot(self):
        # Dibujar Isc usando los datos preparados por generate_plot (last_combined_summary)
        combined_summary = getattr(self, 'last_combined_summary', None)
        if combined_summary is None or combined_summary.empty:
            messagebox.showwarning('Advertencia', 'No hay datos de resumen para generar la gráfica de Isc.')
            return

        # Marcar que Isc fue generado
        self.graph_generated_isc = True

        self.ax.clear()
        # desconectar handler previo
        if getattr(self, 'hover_cid', None):
            try:
                self.canvas.mpl_disconnect(self.hover_cid)
            except Exception:
                pass
            self.hover_cid = None

        combined_summary['Fecha y hora inicio'] = pd.to_datetime(combined_summary['Fecha y hora inicio'], errors='coerce')
        combined_summary = combined_summary.dropna(subset=['Fecha y hora inicio', 'Isc'])

        lines = []
        labels = []
        xdata_lists = []
        xnum_lists = []
        ydata_lists = []
        for filename, group in combined_summary.groupby('Archivo'):
            # Parsear fechas con tolerancia a distintos formatos (dayfirst para dd.mm.yyyy)
            series_dt = pd.to_datetime(group['Fecha y hora inicio'], dayfirst=True, errors='coerce')
            series_y = pd.to_numeric(group['Isc'], errors='coerce')
            mask = series_dt.notna() & series_y.notna()
            if not mask.any():
                continue
            series_dt = series_dt[mask]
            series_y = series_y[mask]
            # convertir a objetos datetime de Python
            xd_dt = series_dt.dt.to_pydatetime()
            yd = series_y.to_numpy(dtype=float)
            # Plot usando datetime objects (matplotlib maneja bien)
            line, = self.ax.plot(xd_dt, yd, marker='o', linestyle='-', label=filename)
            lines.append(line)
            labels.append(filename)
            xdata_lists.append(xd_dt)
            # también almacenar los valores numéricos de matplotlib para el tooltip
            xnum_lists.append(mdates.date2num(xd_dt))
            ydata_lists.append(yd)

        self.ax.set_xlabel('Fecha y hora inicio')
        self.ax.set_ylabel('Isc')
        self.ax.set_title('Isc en función del tiempo')
        self.ax.grid(True)
        # sin leyenda
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M:%S'))
        self.fig.autofmt_xdate(rotation=30)

        # Desconectar cualquier handler de tooltip previo y mostrar leyenda simple
        if getattr(self, 'hover_cid', None):
            try:
                self.canvas.mpl_disconnect(self.hover_cid)
            except Exception:
                pass
            self.hover_cid = None

        # Mostrar leyenda simple en la gráfica Isc
        self.ax.legend()
        self.canvas.draw()

    def on_closing(self):
        plt.close('all')
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = IVCurveApp(root)
    root.mainloop()
