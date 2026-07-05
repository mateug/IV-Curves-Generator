import os
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
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
        self.root.geometry("1200x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.current_folder = os.path.abspath("datos")
        self.excel_data = [] # Lista de dicts con info de cada excel
        
        self.default_i_unit = "μA"
        self.default_v_unit = "V"
        
        self.fig, self.ax = plt.subplots(figsize=(7, 5))
        self.canvas = None
        self.toolbar = None
        self.annot = None
        self.hover_cid = None
        
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
        
        # Treeview para la lista
        columns = ("Estado", "Archivo", "Unidad I", "Unidad V")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.tree.heading("Estado", text="")
        self.tree.heading("Archivo", text="Archivo")
        self.tree.heading("Unidad I", text="I")
        self.tree.heading("Unidad V", text="V")
        self.tree.column("Estado", width=30, anchor=tk.CENTER)
        self.tree.column("Archivo", width=120)
        self.tree.column("Unidad I", width=50, anchor=tk.CENTER)
        self.tree.column("Unidad V", width=50, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Unidades a mostrar
        units_frame = ttk.LabelFrame(left_panel, text="Unidades del Gráfico")
        units_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(units_frame, text="Intensidad:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_i = ttk.Combobox(units_frame, values=list(UNIT_FACTORS["I"].keys()), width=5, state="readonly")
        self.cb_unit_i.set(self.default_i_unit)
        self.cb_unit_i.grid(row=0, column=1, padx=5, pady=5)
        self.cb_unit_i.bind("<<ComboboxSelected>>", lambda e: self.generate_plot())
        
        ttk.Label(units_frame, text="Voltaje:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_v = ttk.Combobox(units_frame, values=list(UNIT_FACTORS["V"].keys()), width=5, state="readonly")
        self.cb_unit_v.set(self.default_v_unit)
        self.cb_unit_v.grid(row=1, column=1, padx=5, pady=5)
        self.cb_unit_v.bind("<<ComboboxSelected>>", lambda e: self.generate_plot())
        
        # Botón Generar
        self.btn_generate = ttk.Button(left_panel, text="Generar Gráfico IV", command=self.generate_plot, state=tk.DISABLED)
        self.btn_generate.pack(fill=tk.X, pady=10)

        # --- PANEL DERECHO (Gráfico y Controles de Gráfico) ---
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Canvas de matplotlib
        self.canvas_frame = ttk.Frame(right_panel)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.canvas_frame)
        self.toolbar.update()
        
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
        cb_swap = ttk.Checkbutton(plot_ctrl_frame, text="Intercambiar X/Y", variable=self.var_swap_axes, command=self.generate_plot)
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
        self.cb_line_style.bind("<<ComboboxSelected>>", lambda e: self.generate_plot())
        
        # Exportar
        export_frame = ttk.Frame(plot_ctrl_frame)
        export_frame.grid(row=1, column=7, columnspan=2, rowspan=2, padx=20)
        btn_export = ttk.Button(export_frame, text="Exportar Gráfico", command=self.export_plot)
        btn_export.pack()

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.current_folder)
        if folder:
            self.current_folder = folder
            self.lbl_folder.config(text=self.current_folder)
            self.scan_folder(self.current_folder)
            
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
        
        for file in files:
            path = os.path.join(folder, file)
            try:
                # Solo leemos la cabecera para ver las columnas
                df = pd.read_excel(path, nrows=0)
                cols = list(df.columns)
                
                # Buscar columnas I y V
                i_cols = [c for c in cols if 'I' in c.upper() or 'INTENSIDAD' in c.upper() or 'CURRENT' in c.upper()]
                v_cols = [c for c in cols if 'V' in c.upper() or 'VOLTAJE' in c.upper() or 'VOLTAGE' in c.upper()]
                
                status = "✅"
                i_unit = "-"
                v_unit = "-"
                i_col_name = None
                v_col_name = None
                
                if not i_cols or not v_cols:
                    status = "❌"
                    invalid_count += 1
                else:
                    if len(i_cols) > 1 or len(v_cols) > 1:
                        status = "⚠️"
                    i_col_name = i_cols[0]
                    v_col_name = v_cols[0]
                    i_unit = self.get_unit_from_str(i_col_name, "I")
                    v_unit = self.get_unit_from_str(v_col_name, "V")
                    valid_count += 1
                
                self.excel_data.append({
                    "filename": file,
                    "path": path,
                    "status": status,
                    "i_col": i_col_name,
                    "v_col": v_col_name,
                    "i_unit": i_unit,
                    "v_unit": v_unit
                })
                
                self.tree.insert("", "end", values=(status, file, i_unit, v_unit))
                
            except Exception as e:
                print(f"Error reading {file}: {e}")
                self.tree.insert("", "end", values=("❌", file, "-", "-"))
                invalid_count += 1

        self.lbl_summary.config(text=f"Total: {len(files)} | Válidos: {valid_count} | Inválidos: {invalid_count}")
        
        if valid_count > 0:
            self.btn_generate.config(state=tk.NORMAL)
        else:
            self.btn_generate.config(state=tk.DISABLED)

    def generate_plot(self):
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
        
        for data in self.excel_data:
            if data["status"] in ["✅", "⚠️"]:
                try:
                    df = pd.read_excel(data["path"])
                    # Convertimos al factor base (A o V) y luego a la unidad destino
                    base_i_factor = UNIT_FACTORS["I"].get(data["i_unit"], 1.0)
                    base_v_factor = UNIT_FACTORS["V"].get(data["v_unit"], 1.0)
                    
                    target_i_factor = UNIT_FACTORS["I"][target_i_unit]
                    target_v_factor = UNIT_FACTORS["V"][target_v_unit]
                    
                    # Multiplicador = (Unidad origen a base) / (Base a unidad destino)
                    mult_i = base_i_factor / target_i_factor
                    mult_v = base_v_factor / target_v_factor
                    
                    i_vals = df[data["i_col"]].dropna() * mult_i
                    v_vals = df[data["v_col"]].dropna() * mult_v
                    
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
        if not self.ax.lines:
            return
        self.ax.relim()
        self.ax.autoscale_view()
        self.update_axes_direction()
        
        self.entry_xmin.delete(0, tk.END)
        self.entry_xmax.delete(0, tk.END)
        self.entry_ymin.delete(0, tk.END)
        self.entry_ymax.delete(0, tk.END)
        
        self.canvas.draw()

    def export_plot(self):
        if not self.ax.lines:
            messagebox.showwarning("Advertencia", "No hay ningún gráfico para exportar.")
            return
            
        base_name = "curvas_iv"
        ext = ".png"
        folder = self.current_folder
        
        # Buscar nombre disponible
        counter = 1
        filename = f"{base_name}{ext}"
        filepath = os.path.join(folder, filename)
        
        while os.path.exists(filepath):
            counter += 1
            filename = f"{base_name}_{counter}{ext}"
            filepath = os.path.join(folder, filename)
            
        # Permitir al usuario cambiarlo
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            initialfile=filename,
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("PDF Document", "*.pdf"), ("SVG Image", "*.svg"), ("All Files", "*.*")]
        )
        
        if save_path:
            try:
                self.fig.savefig(save_path, bbox_inches='tight')
                messagebox.showinfo("Éxito", f"Gráfico guardado en:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el gráfico:\n{e}")

    def on_closing(self):
        plt.close('all')
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = IVCurveApp(root)
    root.mainloop()
