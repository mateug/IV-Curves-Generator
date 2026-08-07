import os
import sys
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# Asegurar que la raíz del proyecto esté en sys.path cuando se ejecuta como script o exe.
# Esto permite importar los módulos como paquete `src` desde `src/main.py`.
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.excel_loader import UNIT_FACTORS, get_unit_from_str, prepare_iv_series, prepare_summary_dataframe, load_excel_file_info
from src.graphing import build_combined_iv, build_combined_summary, build_style_map, prepare_isc_plot_data
from src.ui import IVCurveUI
from src.exporting import export_combined_excel


class IVCurveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Curvas IV")
        self.root.geometry("1400x900")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.current_folder = os.path.abspath("datos")
        self.excel_data = []
        self.summary_data = None
        self.valid_files = set()
        self.selected_files = set()
        
        self.default_i_unit = "μA"
        self.default_v_unit = "V"
        self.unit_factors = UNIT_FACTORS
        
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
        self.ui = IVCurveUI(self.root, self)
        self.lbl_folder = self.ui.lbl_folder
        self.lbl_summary = self.ui.lbl_summary
        self.progress_bar = self.ui.progress_bar
        self.tree = self.ui.tree
        self.btn_select_all = self.ui.btn_select_all
        self.cb_unit_i = self.ui.cb_unit_i
        self.cb_unit_v = self.ui.cb_unit_v
        self.btn_generate = self.ui.btn_generate
        self.btn_export_excel = self.ui.btn_export_excel
        self.notebook = self.ui.notebook
        self.iv_tab = self.ui.iv_tab
        self.iv_log_tab = self.ui.iv_log_tab
        self.isc_tab = self.ui.isc_tab
        self.entry_xmin = self.ui.entry_xmin
        self.entry_xmax = self.ui.entry_xmax
        self.entry_ymin = self.ui.entry_ymin
        self.entry_ymax = self.ui.entry_ymax
        self.var_swap_axes = self.ui.var_swap_axes
        self.var_invert_x = self.ui.var_invert_x
        self.var_invert_y = self.ui.var_invert_y
        self.cb_line_style = self.ui.cb_line_style
        self.canvas = self.ui.canvas
        self.toolbar = self.ui.toolbar

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.iv_tab)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.iv_tab)
        self.toolbar.update()

    def browse_folder(self):
        """Abre el selector de carpeta y reescanea su contenido."""
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

    def update_selection_state(self):
        """Sincroniza la tabla, los botones y la habilitación de generación."""
        self.refresh_selection_columns()
        self.update_select_all_button()
        self.update_generate_button_state()

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
        self.update_selection_state()

    def on_tree_click(self, event):
        # Maneja la selección de archivos en la tabla que se van a graficar. Ctrl+Click permite seleccionar múltiples archivos.
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

        if event.state & 0x1:  # Shift + Click --> 
            # Selección de rango
            all_valid_files = sorted(self.valid_files)
            if not self.selected_files:
                self.selected_files.add(filename)
            else:
                last_selected = sorted(self.selected_files)[-1]
                try:
                    start_index = all_valid_files.index(last_selected)
                    end_index = all_valid_files.index(filename)
                    if start_index > end_index:
                        start_index, end_index = end_index, start_index
                    for f in all_valid_files[start_index:end_index + 1]:
                        self.selected_files.add(f)
                except ValueError:
                    self.selected_files.add(filename)

        # Opción A: Si solo se hace click solo se selecciona un único archivo
            """
        elif event.state & 0x4:  # Ctrl + Click --> Se mantienen los archivos anteriores seleccionados y se alterna el estado del archivo clickeado
                    if filename in self.selected_files:
                        self.selected_files.remove(filename)
                    else:
                        self.selected_files.add(filename)
        else: # Click --> Se deseleccionan todos los archivos seleccionados y se selecciona al que se le ha hecho click
            if filename in self.selected_files and len(self.selected_files) == 1:
                self.selected_files.clear()
            else:
                self.selected_files = {filename}
            """

        # Opción B: No hace falta hacer Ctrl+Click para seleccionar varios archivos
        else: # Click --> Se selecciona o deselecciona un nuevo archivo
            if filename in self.selected_files:
                self.selected_files.remove(filename)
            else:
                self.selected_files.add(filename)

        self.update_selection_state()

    def get_unit_from_str(self, s, var_type):
        return get_unit_from_str(s, var_type)

    def get_selected_items(self):
        """Devuelve los archivos Excel válidos que están actualmente seleccionados."""
        return [
            item for item in self.excel_data
            if item['filename'] in self.selected_files and item['status'] in ['✅', '⚠️']
        ]

    def scan_folder(self, folder):
        """Escanea una carpeta y carga la información de los archivos Excel detectados."""
        self.excel_data = []
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not os.path.exists(folder):
            self.lbl_summary.config(text="Carpeta no encontrada.")
            self.btn_generate.config(state=tk.DISABLED)
            return

        files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        total_files = len(files)

        valid_count = 0
        invalid_count = 0
        self.summary_data = []

        self.progress_bar.config(maximum=max(total_files, 1), value=0)
        self.lbl_summary.config(text=f"Escaneando {total_files} archivos...")
        self.root.update_idletasks()

        for index, file_name in enumerate(files, start=1):
            path = os.path.join(folder, file_name)
            try:
                info = load_excel_file_info(path, file_name)
                status = info["status"]
                summary_df = info["summary_df"]

                if status == "❌":
                    invalid_count += 1
                else:
                    valid_count += 1

                self.excel_data.append(info)
                if status in ["✅", "⚠️"] and summary_df is not None:
                    self.summary_data.append({"filename": file_name, "data": summary_df})

                self.tree.insert("", "end", values=("", status, file_name, info["i_unit"], info["v_unit"]))
                self.progress_bar.config(value=index)
                self.lbl_summary.config(text=f"Escaneando {total_files} archivos... ({index}/{total_files})")
                self.root.update_idletasks()

            except Exception as e:
                print(f"Error reading {file_name}: {e}")
                self.tree.insert("", "end", values=("", "❌", file_name, "-", "-"))
                invalid_count += 1

        self.progress_bar.config(value=total_files)
        self.lbl_summary.config(text=f"Total: {total_files} | Válidos: {valid_count} | Inválidos: {invalid_count}")

        self.valid_files = {item['filename'] for item in self.excel_data if item['status'] in ['✅', '⚠️']}
        self.selected_files.intersection_update(self.valid_files)
        self.update_selection_state()

        if valid_count > 0:
            self.btn_export_excel.config(state=tk.NORMAL)
            self.btn_select_all.config(state=tk.NORMAL)
        else:
            self.btn_export_excel.config(state=tk.DISABLED)
            self.btn_select_all.config(state=tk.DISABLED)

    def generate_plot(self, log_scale=False):
        """Genera la gráfica IV a partir de los archivos seleccionados."""
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
        
        fmt = build_style_map().get(self.cb_line_style.get(), "-")
        
        lines = []
        labels = []
        
        selected_items = self.get_selected_items()
        for data in selected_items:
            try:
                df = pd.read_excel(data["path"], sheet_name="datos_IV", header=data["header_row"])
                x_vals, y_vals = prepare_iv_series(
                    df=df,
                    v_col=data["v_col"],
                    i_col=data["i_col"],
                    source_i_unit=data["i_unit"],
                    source_v_unit=data["v_unit"],
                    target_i_unit=target_i_unit,
                    target_v_unit=target_v_unit,
                    swap=swap,
                )

                if log_scale:
                    # Solo el eje Y será logarítmico. Se omiten los puntos con valor de Y no positivo.
                    mask = y_vals > 0
                    x_plot = x_vals[mask]
                    y_plot = y_vals[mask]
                else:
                    x_plot = x_vals
                    y_plot = y_vals

                line, = self.ax.plot(x_plot, y_plot, fmt, label=data["filename"][:-5]) # Quito del label el ".xlsx"
                lines.append(line)
                    
                # Guardar info para el tooltip
                labels.append(f"Archivo: {data['filename']}\nCol V: {data['v_col']}\nCol I: {data['i_col']}")
                    
            except Exception as e:
                print(f"Error ploting {data['filename']}: {e}")
                    
        x_label = f"Voltaje ({target_v_unit})" if not swap else f"Intensidad ({target_i_unit})"
        y_label = f"Intensidad ({target_i_unit})" if not swap else f"Voltaje ({target_v_unit})"
        
        self.ax.set_xlabel(x_label)
        self.ax.set_ylabel(y_label)
        self.ax.set_title("Curvas IV" + (" (log)" if log_scale else ""))
        self.ax.grid(True)
        self.ax.set_xscale('linear')
        self.ax.set_yscale('log' if log_scale else 'linear')
        self.ax.legend(loc="upper right", fontsize='small', framealpha=0.9)
        
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
            self.ax.set_xlim(max(xlim), min(xlim))
        else:
            self.ax.set_xlim(min(xlim), max(xlim))
                
        ylim = self.ax.get_ylim()
        if self.var_invert_y.get():
            self.ax.set_ylim(max(ylim), min(ylim))
        else:
            self.ax.set_ylim(min(ylim), max(ylim))
                
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
            self.generate_plot(log_scale=False)
        elif current == 'IV log' and self.graph_generated_iv:
            self.generate_plot(log_scale=True)
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

    def read_summary_sheet(self, path):
        return prepare_summary_dataframe(path)

    def same_voltage_series(self, s1, s2):
        if len(s1) != len(s2):
            return False
        s1_arr = np.asarray(s1, dtype=float)
        s2_arr = np.asarray(s2, dtype=float)
        return np.allclose(s1_arr, s2_arr, atol=1e-9, rtol=1e-6)

    def build_combined_iv(self):
        return build_combined_iv(self.get_selected_items(), self.same_voltage_series)

    def build_combined_summary(self):
        return build_combined_summary(self.excel_data, self.selected_files)

    def export_combined_excel(self):
        """Exporta los datos combinados a un libro Excel usando el módulo dedicado."""
        combined_iv = self.build_combined_iv()
        combined_summary = self.build_combined_summary()
        export_combined_excel(combined_iv, combined_summary, self.current_folder)

    def export_plot(self):
        # Exporta el gráfico correspondiente a la pestaña activa con nombre por defecto
        current = self.notebook.tab(self.notebook.select(), "text") if hasattr(self, 'notebook') else 'IV'
        default_name = 'grafico_iv.png' if current == 'IV' else 'grafico_iv_log.png' if current == 'IV log' else 'grafico_isc.png'

        if current in ['IV', 'IV log'] and not self.graph_generated_iv:
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
            elif current == 'IV log':
                self.generate_plot(log_scale=True)
            else:
                self.generate_plot()
            self.fig.savefig(save_path, dpi=300, bbox_inches='tight')
            messagebox.showinfo('Éxito', f'Gráfico exportado a:\n{save_path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar el gráfico:\n{e}')

    def on_tab_change(self):
        # Recrea el canvas en la pestaña activa y dibuja el gráfico correspondiente
        selected = self.notebook.tab(self.notebook.select(), "text")
        if selected == 'IV':
            parent = self.iv_tab
        elif selected == 'IV log':
            parent = self.iv_log_tab
        else:
            parent = self.isc_tab

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
                self.generate_plot(log_scale=False)
            else:
                self.ax.clear()
                self.ax.set_title('Pulse "Generar Gráfico IV" para generar los gráficos')
                self.canvas.draw()
        elif selected == 'IV log':
            if self.graph_generated_iv:
                self.generate_plot(log_scale=True)
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
        for filename, series_dt, series_y in prepare_isc_plot_data(combined_summary):
            xd_dt = series_dt.dt.to_pydatetime()
            yd = series_y.to_numpy(dtype=float)
            line, = self.ax.plot(xd_dt, yd, marker='o', linestyle='-', label=filename)
            lines.append(line)
            labels.append(filename)

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
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    root = tk.Tk()
    
    # Aumentar la fuente por defecto para que se vea mejor en todas partes
    from tkinter import font
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(size=9)
    
    app = IVCurveApp(root)
    root.mainloop()
