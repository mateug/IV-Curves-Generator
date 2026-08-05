import tkinter as tk
from tkinter import ttk


class IVCurveUI:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_panel = ttk.Frame(main_frame, width=300)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        folder_frame = ttk.LabelFrame(left_panel, text="Carpeta de Datos")
        folder_frame.pack(fill=tk.X, pady=5)

        self.lbl_folder = ttk.Label(folder_frame, text=self.app.current_folder, wraplength=250)
        self.lbl_folder.pack(pady=5, padx=5)

        ttk.Button(folder_frame, text="Seleccionar Carpeta", command=self.app.browse_folder).pack(pady=5)
        ttk.Button(folder_frame, text="Actualizar / Escanear", command=lambda: self.app.scan_folder(self.app.current_folder)).pack(pady=5)

        self.lbl_summary = ttk.Label(left_panel, text="Buscando...")
        self.lbl_summary.pack(pady=(5, 2))

        self.progress_bar = ttk.Progressbar(left_panel, orient="horizontal", length=250, mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=5, pady=(0, 5))

        ttk.Label(left_panel, text="Ctrl + clic para seleccionar varios archivos", foreground="#555").pack(padx=5, pady=(0, 5))

        list_frame = ttk.LabelFrame(left_panel, text="Archivos Encontrados")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

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
        self.tree.bind("<Button-1>", self.app.on_tree_click)

        select_buttons = ttk.Frame(list_frame)
        select_buttons.pack(fill=tk.X, padx=5, pady=(0, 5))
        self.btn_select_all = ttk.Button(select_buttons, text="Seleccionar todos", command=self.app.toggle_select_all, state=tk.DISABLED)
        self.btn_select_all.pack(fill=tk.X)

        units_frame = ttk.LabelFrame(left_panel, text="Unidades del Gráfico")
        units_frame.pack(fill=tk.X, pady=5)

        ttk.Label(units_frame, text="Intensidad:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_i = ttk.Combobox(units_frame, values=list(self.app.unit_factors["I"].keys()), width=5, state="readonly")
        self.cb_unit_i.set(self.app.default_i_unit)
        self.cb_unit_i.grid(row=0, column=1, padx=5, pady=5)
        self.cb_unit_i.bind("<<ComboboxSelected>>", lambda e: setattr(self.app, 'needs_regen', True))

        ttk.Label(units_frame, text="Voltaje:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.cb_unit_v = ttk.Combobox(units_frame, values=list(self.app.unit_factors["V"].keys()), width=5, state="readonly")
        self.cb_unit_v.set(self.app.default_v_unit)
        self.cb_unit_v.grid(row=1, column=1, padx=5, pady=5)
        self.cb_unit_v.bind("<<ComboboxSelected>>", lambda e: setattr(self.app, 'needs_regen', True))

        self.btn_generate = ttk.Button(left_panel, text="Generar Gráfico IV", command=self.app.generate_plot, state=tk.DISABLED)
        self.btn_generate.pack(fill=tk.X, pady=(10, 5))

        self.btn_export_excel = ttk.Button(left_panel, text="Exportar Excel combinado", command=self.app.export_combined_excel, state=tk.DISABLED)
        self.btn_export_excel.pack(fill=tk.X, pady=(0, 10))

        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(right_panel)
        self.iv_tab = ttk.Frame(self.notebook)
        self.iv_log_tab = ttk.Frame(self.notebook)
        self.isc_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.iv_tab, text="IV")
        self.notebook.add(self.iv_log_tab, text="IV log")
        self.notebook.add(self.isc_tab, text="Isc")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.canvas = None
        self.toolbar = None
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.app.on_tab_change())

        plot_ctrl_frame = ttk.LabelFrame(right_panel, text="Controles del Gráfico (Tras generar)")
        plot_ctrl_frame.pack(fill=tk.X, pady=5)

        ttk.Label(plot_ctrl_frame, text="X min:").grid(row=0, column=0, padx=2)
        self.entry_xmin = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_xmin.grid(row=0, column=1, padx=2)

        ttk.Label(plot_ctrl_frame, text="X max:").grid(row=0, column=2, padx=2)
        self.entry_xmax = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_xmax.grid(row=0, column=3, padx=2)

        ttk.Label(plot_ctrl_frame, text="Y min:").grid(row=1, column=0, padx=2, pady=5)
        self.entry_ymin = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_ymin.grid(row=1, column=1, padx=2)

        ttk.Label(plot_ctrl_frame, text="Y max:").grid(row=1, column=2, padx=2)
        self.entry_ymax = ttk.Entry(plot_ctrl_frame, width=8)
        self.entry_ymax.grid(row=1, column=3, padx=2)

        ttk.Button(plot_ctrl_frame, text="Aplicar", command=self.app.apply_limits).grid(row=0, column=4, rowspan=2, padx=10)
        ttk.Button(plot_ctrl_frame, text="Reiniciar Límites", command=self.app.reset_limits).grid(row=0, column=5, rowspan=2, padx=5)

        self.var_swap_axes = tk.BooleanVar(value=False)
        ttk.Checkbutton(plot_ctrl_frame, text="Intercambiar X/Y", variable=self.var_swap_axes, command=lambda: self.app.generate_plot() if self.app.graph_generated_iv else None).grid(row=0, column=6, padx=10)

        self.var_invert_x = tk.BooleanVar(value=False)
        ttk.Checkbutton(plot_ctrl_frame, text="Invertir Eje X", variable=self.var_invert_x, command=self.app.update_axes_direction).grid(row=1, column=6, padx=10, sticky=tk.W)

        self.var_invert_y = tk.BooleanVar(value=False)
        ttk.Checkbutton(plot_ctrl_frame, text="Invertir Eje Y", variable=self.var_invert_y, command=self.app.update_axes_direction).grid(row=2, column=6, padx=10, sticky=tk.W)

        ttk.Label(plot_ctrl_frame, text="Estilo de curva:").grid(row=0, column=7, padx=10, sticky=tk.E)
        self.cb_line_style = ttk.Combobox(plot_ctrl_frame, values=["Línea continua (-)", "Línea rayada (--)", "Puntos (.)", "Puntos y línea (.-)"], width=15, state="readonly")
        self.cb_line_style.set("Línea continua (-)")
        self.cb_line_style.grid(row=0, column=8, padx=5, sticky=tk.W)
        self.cb_line_style.bind("<<ComboboxSelected>>", lambda e: self.app.generate_plot() if self.app.graph_generated_iv else None)

        export_frame = ttk.Frame(plot_ctrl_frame)
        export_frame.grid(row=1, column=7, columnspan=2, rowspan=2, padx=20)
        ttk.Button(export_frame, text="Exportar Gráfico", command=self.app.export_plot).pack(side=tk.LEFT)
