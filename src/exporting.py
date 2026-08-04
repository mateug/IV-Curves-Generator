import os
import pandas as pd
from tkinter import filedialog, messagebox


def export_combined_excel(combined_iv, combined_summary, current_folder):
    """Exporta los datos IV y resumen combinados a un libro Excel."""
    if combined_iv.empty:
        messagebox.showwarning('Advertencia', 'No hay datos IV combinados para exportar.')
        return None

    save_path = filedialog.asksaveasfilename(
        initialdir=current_folder,
        initialfile='iv_combinado.xlsx',
        defaultextension='.xlsx',
        filetypes=[('Excel Workbook', '*.xlsx'), ('All Files', '*.*')]
    )
    if not save_path:
        return None

    try:
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df_to_write = combined_iv.copy()
            if isinstance(df_to_write.columns, pd.MultiIndex):
                df_to_write.columns = [
                    f"{a}_{b}" if not (isinstance(a, str) and isinstance(b, str)) else f"{a}_{b}"
                    for a, b in df_to_write.columns
                ]
            df_to_write.to_excel(writer, sheet_name='datos_IV', index=False)
            if not combined_summary.empty:
                cols = combined_summary.columns.tolist()
                if 'Archivo' in cols:
                    cols.insert(0, cols.pop(cols.index('Archivo')))
                    combined_summary = combined_summary[cols]
                combined_summary.to_excel(writer, sheet_name='resumen', index=False)
        messagebox.showinfo('Éxito', f'Archivo exportado a:\n{save_path}')
        return save_path
    except Exception as e:
        messagebox.showerror('Error', f'No se pudo exportar el archivo:\n{e}')
        return None
