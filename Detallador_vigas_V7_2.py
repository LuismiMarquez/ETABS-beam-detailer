import sys
import json
import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import comtypes.client
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class DetalladorVigasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor CAD de Despiece y Modulación - ETABS (Estable Final)")
        self.root.geometry("1500x950")
        
        self.SapModel = None
        self.viga_label = "" 
        self.demanda_top = []
        self.demanda_bot = []
        self.estaciones = []
        self.columnas_graficas = [] 
        self.vigas_graficas = [] 
        self.longitud_eje_real = 0 
        
        self.barras = []
        self.bar_id_counter = 1
        self.rebar_db = {"#3": 71, "#4": 129, "#5": 200, "#6": 284, "#7": 387, "#8": 510}
        self.diametros_db = {"#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1, "#7": 22.2, "#8": 25.4}
        self.flejes_db = ["#3", "#4", "#5", "#6"]
        
        self.datos_cortante = []
        self.cortes = []
        self.ejes_extraidos = {}
        self.ruta_json_acumulado = None

        # --- Fase 1: Estado Drag & Drop (ax_barras interactivo) ---
        self._dragging_bar = None   # dict ref directa a self.barras[i]
        self._drag_edge = None      # 'x1' | 'x2'
        self._drag_tol = 0.15       # m tolerancia hit-test en X
        self._snap = 0.05           # m snap
        self._bar_artists = {}      # id -> {'line': Line2D, 'hook_ini': Line2D|None, 'hook_fin': Line2D|None, 'text': Text, 'y_pos': float, 'c': str}
        self._cort_artists = []     # [] de PolyCollections (axvspan) para pick Fase 2
        # --- Directriz 1.1: Handles visuales ---
        self._selected_bar_id = None
        self._handle_artists = []   # marcadores cuadrados en x1/x2 de la barra seleccionada
        
        self.crear_interfaz()
        self.root.after(500, self.conectar_etabs)
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)
        
    def cerrar_aplicacion(self):
        self.SapModel = None
        if hasattr(self, 'fig'): plt.close(self.fig)
        if hasattr(self, 'fig_flex'): plt.close(self.fig_flex)
        if hasattr(self, 'fig_cortante'): plt.close(self.fig_cortante)
        if hasattr(self, 'fig_cortes'): plt.close(self.fig_cortes)
        self.root.quit()
        self.root.destroy()
        
    def conectar_etabs(self):
        try:
            self.lbl_estado.config(text="Buscando ETABS...", fg="orange")
            self.root.update()
            helper = comtypes.client.CreateObject('ETABSv1.Helper')
            helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
            try:
                myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject")
                self.SapModel = myETABSObject.SapModel
            except Exception:
                try:
                    myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
                    self.SapModel = myETABSObject.SapModel
                except Exception:
                    self.lbl_estado.config(text="Iniciando ETABS...", fg="blue")
                    self.root.update()
                    myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                    myETABSObject.ApplicationStart()
                    self.SapModel = myETABSObject.SapModel
            
            self.SapModel.SetPresentUnits(9) 
            self.lbl_estado.config(text="¡Conectado a ETABS!", fg="green")
            self.btn_extraer.config(state="normal")
            self.btn_todos.config(state="normal")
        except Exception as e:
            self.lbl_estado.config(text="Fallo de conexión.", fg="red")

    def crear_interfaz(self):
        frame_izq = tk.Frame(self.root, width=450)
        frame_izq.pack(side="left", fill="y", padx=10, pady=5)
        
        # 1. Extraccion
        f_ext = tk.LabelFrame(frame_izq, text="1. Extracción de ETABS", padx=10, pady=5)
        f_ext.pack(fill="x", pady=5)
        self.lbl_estado = tk.Label(f_ext, text="Iniciando...", font=("Arial", 9, "bold"))
        self.lbl_estado.pack(anchor="w")
        self.btn_extraer = tk.Button(f_ext, text="⬇ Extraer Eje Seleccionado", state="disabled", bg="#fcf8e3", font=("Arial", 9, "bold"), command=self.extraer_diseno)
        self.btn_extraer.pack(fill="x", pady=2)
        self.btn_todos = tk.Button(f_ext, text="🌐 Extraer TODAS las Vigas", state="disabled", bg="#d9edf7", font=("Arial", 9, "bold"), command=self.extraer_todo_modelo)
        self.btn_todos.pack(fill="x", pady=2)
        
        f_sel_eje = tk.Frame(f_ext)
        f_sel_eje.pack(fill="x", pady=2)
        self.cmb_ejes = ttk.Combobox(f_sel_eje, state="readonly", width=28)
        self.cmb_ejes.pack(side="left", expand=True, fill="x")
        self.cmb_ejes.bind("<<ComboboxSelected>>", lambda e: self.cargar_eje_combo())
        
        self.aplicar_nsr_envolvente = tk.BooleanVar(value=False)
        tk.Checkbutton(f_ext, text="Aplicar mínimos NSR-10 (0.25Mu / 0.5Mu)", variable=self.aplicar_nsr_envolvente, command=self.refrescar_envolvente_activa, font=("Arial", 8), fg="#cc6600").pack(anchor="w")
        
        # 2. Modulacion Automatica
        f_auto = tk.LabelFrame(frame_izq, text="2. Modulación Automática", padx=10, pady=5)
        f_auto.pack(fill="x", pady=5)
        f_a1 = tk.Frame(f_auto)
        f_a1.pack(fill="x", pady=2)
        tk.Label(f_a1, text="L. Comercial(m):").pack(side="left")
        self.ent_lcom = tk.Entry(f_a1, width=5)
        self.ent_lcom.insert(0, "12.0")
        self.ent_lcom.pack(side="left", padx=2)
        tk.Label(f_a1, text="Traslapo(m):").pack(side="left")
        self.ent_ltras = tk.Entry(f_a1, width=5)
        self.ent_ltras.insert(0, "0.7")
        self.ent_ltras.pack(side="left", padx=2)
        tk.Label(f_a1, text="Gancho(m):").pack(side="left")
        self.ent_gancho_auto = tk.Entry(f_a1, width=5)
        self.ent_gancho_auto.insert(0, "0.3")
        self.ent_gancho_auto.pack(side="left", padx=2)
        tk.Label(f_a1, text="Rec(cm):").pack(side="left")
        self.ent_rec = tk.Entry(f_a1, width=4)
        self.ent_rec.insert(0, "4.0")
        self.ent_rec.pack(side="left", padx=2)
        
        f_a2 = tk.Frame(f_auto)
        f_a2.pack(fill="x", pady=2)
        tk.Label(f_a2, text="Cant Sup:").pack(side="left")
        self.spn_a_csup = tk.Spinbox(f_a2, from_=2, to=10, width=3)
        self.spn_a_csup.pack(side="left")
        self.cmb_a_tsup = ttk.Combobox(f_a2, values=list(self.rebar_db.keys()), width=4, state="readonly")
        self.cmb_a_tsup.current(2)
        self.cmb_a_tsup.pack(side="left", padx=2)
        tk.Label(f_a2, text="Cant Inf:").pack(side="left", padx=(10,0))
        self.spn_a_cinf = tk.Spinbox(f_a2, from_=2, to=10, width=3)
        self.spn_a_cinf.pack(side="left")
        self.cmb_a_tinf = ttk.Combobox(f_a2, values=list(self.rebar_db.keys()), width=4, state="readonly")
        self.cmb_a_tinf.current(2)
        self.cmb_a_tinf.pack(side="left", padx=2)
        tk.Button(f_auto, text="⚡ Generar Cortes Automáticos", bg="#d9edf7", command=self.generar_continuo).pack(fill="x", pady=5)

        # 3. Editor Manual
        f_edit = tk.LabelFrame(frame_izq, text="3. Editor Manual", padx=10, pady=5)
        f_edit.pack(fill="x", pady=5)
        f_e1 = tk.Frame(f_edit)
        f_e1.pack(fill="x")
        self.cmb_cara = ttk.Combobox(f_e1, values=["Superior", "Inferior"], width=8, state="readonly")
        self.cmb_cara.current(0)
        self.cmb_cara.pack(side="left", padx=2)
        self.cmb_rol = ttk.Combobox(f_e1, values=["Continuo", "Bastón"], width=8, state="readonly")
        self.cmb_rol.current(1)
        self.cmb_rol.pack(side="left", padx=2)
        self.spn_cant = tk.Spinbox(f_e1, from_=1, to=10, width=3)
        self.spn_cant.pack(side="left", padx=2)
        self.cmb_tipo = ttk.Combobox(f_e1, values=list(self.rebar_db.keys()), width=4, state="readonly")
        self.cmb_tipo.current(2)
        self.cmb_tipo.pack(side="left", padx=2)
        
        f_e2 = tk.Frame(f_edit)
        f_e2.pack(fill="x", pady=5)
        tk.Label(f_e2, text="X1:").pack(side="left")
        self.ent_x1 = tk.Entry(f_e2, width=5)
        self.ent_x1.pack(side="left", padx=1)
        tk.Label(f_e2, text="X2:").pack(side="left")
        self.ent_x2 = tk.Entry(f_e2, width=5)
        self.ent_x2.pack(side="left", padx=1)
        tk.Label(f_e2, text="G.Ini(m):").pack(side="left", padx=(5,0))
        self.ent_g1 = tk.Entry(f_e2, width=4)
        self.ent_g1.insert(0, "0.0")
        self.ent_g1.pack(side="left", padx=1)
        tk.Label(f_e2, text="G.Fin(m):").pack(side="left")
        self.ent_g2 = tk.Entry(f_e2, width=4)
        self.ent_g2.insert(0, "0.0")
        self.ent_g2.pack(side="left", padx=1)
        
        f_ebtns = tk.Frame(f_edit)
        f_ebtns.pack(fill="x", pady=2)
        tk.Button(f_ebtns, text="➕ Añadir", bg="#dff0d8", width=12, command=self.add_bar).pack(side="left", expand=True, padx=2)
        tk.Button(f_ebtns, text="🔄 Actualizar", bg="#fcf8e3", width=12, command=self.update_bar).pack(side="left", expand=True, padx=2)
        tk.Button(f_ebtns, text="❌ Borrar", bg="#f2dede", width=12, command=self.delete_bar).pack(side="left", expand=True, padx=2)

        # 4. Tabla Maestra
        f_tbl = tk.Frame(frame_izq)
        f_tbl.pack(fill="both", expand=True, pady=5)
        cols = ("ID", "Cara", "Rol", "Barra", "X1", "X2", "G1", "G2", "L_Cort")
        self.tree = ttk.Treeview(f_tbl, columns=cols, show="headings", height=6)
        for c in cols: 
            self.tree.heading(c, text=c)
            self.tree.column(c, width=35, anchor="center")
        self.tree.column("Cara", width=55)
        self.tree.column("Rol", width=55)
        self.tree.column("Barra", width=45)
        self.tree.column("L_Cort", width=45)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        f_vis = tk.Frame(frame_izq)
        f_vis.pack(fill="x", pady=5)
        self.mostrar_confinamiento = tk.BooleanVar(value=False)
        tk.Checkbutton(f_vis, text="👁 Mostrar Zonas 2H", variable=self.mostrar_confinamiento, command=self.actualizar_grafica, font=("Arial", 9, "bold"), fg="#cc6600").pack(side="left")

        f_json = tk.Frame(frame_izq)
        f_json.pack(fill="x", pady=2)
        tk.Button(f_json, text="💾 Guardar JSON Acumulado", bg="#4cae4c", fg="white", command=self.exportar_json).pack(side="left", expand=True, padx=1)
        tk.Button(f_json, text="📂 Archivo", bg="#f0ad4e", fg="white", command=self.cambiar_archivo_json).pack(side="left", padx=1)
        self.lbl_ruta_json = tk.Label(frame_izq, text="Sin archivo acumulado", font=("Arial", 7), fg="#777777")
        self.lbl_ruta_json.pack(anchor="w")

        # --- PANEL DERECHO ---
        self.frame_der = tk.Frame(self.root)
        self.frame_der.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.notebook = ttk.Notebook(self.frame_der)
        self.notebook.pack(fill="both", expand=True)
        
        # Pestañas
        self.tab_despiece = tk.Frame(self.notebook)
        self.notebook.add(self.tab_despiece, text="📐 Despiece Físico")
        self.fig, (self.ax_top, self.ax_barras, self.ax_bot) = plt.subplots(3, 1, figsize=(8, 10), sharex=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_despiece)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        # Fase 1: wiring Drag & Drop en ax_barras
        self.canvas.mpl_connect('button_press_event', self.on_barras_press)
        self.canvas.mpl_connect('motion_notify_event', self.on_barras_motion)
        self.canvas.mpl_connect('button_release_event', self.on_barras_release)

        self.tab_flexion = tk.Frame(self.notebook)
        self.notebook.add(self.tab_flexion, text="📊 Verificación Flexión")
        self.fig_flex, self.ax_flex = plt.subplots(figsize=(8, 6))
        self.canvas_flex = FigureCanvasTkAgg(self.fig_flex, master=self.tab_flexion)
        self.canvas_flex.get_tk_widget().pack(fill="both", expand=True)
        self.txt_alertas = tk.Text(self.tab_flexion, height=5, bg="#fff3cd", fg="#856404", font=("Consolas", 9, "bold"))
        self.txt_alertas.pack(fill="x", padx=5, pady=5)
        
        f_combos = tk.LabelFrame(self.tab_flexion, text="Combos de Diseño Extraídos de ETABS", padx=5, pady=3)
        f_combos.pack(fill="x", padx=5, pady=3)
        self.lbl_combos = tk.Label(f_combos, text="Extrayendo...", fg="#777777")
        self.lbl_combos.pack(anchor="w")
        f_combo_body = tk.Frame(f_combos)
        f_combo_body.pack(fill="both", expand=True)
        scroll_combos = tk.Scrollbar(f_combo_body, orient="vertical")
        self.lst_combos = tk.Listbox(f_combo_body, selectmode="multiple", exportselection=False, height=4, yscrollcommand=scroll_combos.set, font=("Consolas", 8))
        scroll_combos.config(command=self.lst_combos.yview)
        self.lst_combos.pack(side="left", fill="both", expand=True)
        scroll_combos.pack(side="right", fill="y")
        f_combo_btns = tk.Frame(f_combos)
        f_combo_btns.pack(fill="x")
        tk.Button(f_combo_btns, text="☑ Todos", command=lambda: self.lst_combos.select_set(0, 'end')).pack(side="left", padx=2)
        tk.Button(f_combo_btns, text="☐ Ninguno", command=lambda: self.lst_combos.select_clear(0, 'end')).pack(side="left", padx=2)
        tk.Button(f_combo_btns, text="🔄 Recalcular Envolvente", bg="#337ab7", fg="white", font=("Arial", 8, "bold"), command=self.recalcular_con_combos).pack(side="left", padx=8)
        
        self.tab_cortante = tk.Frame(self.notebook)
        self.notebook.add(self.tab_cortante, text="✂️ Cortante Capacidad")
        f_param_c = tk.Frame(self.tab_cortante, bg="#e8f4f8", pady=5, padx=5)
        f_param_c.pack(fill="x")
        
        tk.Label(f_param_c, text="Los valores de Carga Axial y Carga Distribuida se extraen automáticamente.", bg="#e8f4f8", font=("Arial", 9, "italic"), fg="blue").pack(side="left", padx=10)
        tk.Button(f_param_c, text="🔄 Analizar Tramos por Capacidad", command=self.actualizar_verificacion_cortante, bg="#5bc0de", fg="white", font=("Arial", 9, "bold")).pack(side="left", padx=20)
        
        cols_c = ("Tramo", "Zona", "X1", "X2", "Ve (kN)", "Vu (kN)", "Phi", "Req_Vs (kN)", "Ramas", "Calibre", "S_req (cm)", "S_col (cm)", "Estado")
        self.tree_cort = ttk.Treeview(self.tab_cortante, columns=cols_c, show="headings", height=8)
        for c in cols_c: 
            self.tree_cort.heading(c, text=c)
            self.tree_cort.column(c, width=50, anchor="center")
        self.tree_cort.column("Tramo", width=70)
        self.tree_cort.column("Zona", width=110)
        self.tree_cort.pack(fill="x", pady=5)
        self.tree_cort.bind('<<TreeviewSelect>>', self.on_tree_cortante_select)
        
        f_edit_c = tk.LabelFrame(self.tab_cortante, text="Editar Zona", padx=10, pady=5, bg="#fcf8e3")
        f_edit_c.pack(fill="x", pady=5)
        tk.Label(f_edit_c, text="Ramas:", bg="#fcf8e3").pack(side="left")
        self.spn_c_ramas = tk.Spinbox(f_edit_c, from_=2, to=8, increment=2, width=4)
        self.spn_c_ramas.pack(side="left", padx=5)
        tk.Label(f_edit_c, text="Calibre:", bg="#fcf8e3").pack(side="left")
        self.cmb_c_tipo = ttk.Combobox(f_edit_c, values=self.flejes_db, width=5, state="readonly")
        self.cmb_c_tipo.pack(side="left", padx=5)
        tk.Label(f_edit_c, text="S. Colocada (cm):", bg="#fcf8e3").pack(side="left")
        self.ent_c_s = tk.Entry(f_edit_c, width=6)
        self.ent_c_s.pack(side="left", padx=5)
        tk.Button(f_edit_c, text="💾 Actualizar", command=self.actualizar_zona_cortante, bg="#5cb85c", fg="white").pack(side="left", padx=20)
        
        self.fig_cortante, self.ax_v2 = plt.subplots(figsize=(8, 4))
        self.canvas_cortante = FigureCanvasTkAgg(self.fig_cortante, master=self.tab_cortante)
        self.canvas_cortante.get_tk_widget().pack(fill="both", expand=True)
        # Fase 2: pick de zonas de cortante (axvspan picker) -> Treeview
        self.canvas_cortante.mpl_connect('pick_event', self.on_cortante_pick)
        # también para zonas 2H en ax_barras si se activa picker allí
        try:
            self.canvas.mpl_connect('pick_event', self.on_cortante_pick)
        except Exception:
            pass
        
        self.tab_cortes = tk.Frame(self.notebook)
        self.notebook.add(self.tab_cortes, text="🏗️ Cortes de Sección")
        f_btn_cortes = tk.Frame(self.tab_cortes)
        f_btn_cortes.pack(fill="x")
        tk.Button(f_btn_cortes, text="✂️ Generar Cortes", command=self.generar_cortes, bg="#337ab7", fg="white", font=("Arial", 9, "bold")).pack(side="left", padx=5, pady=3)
        self.lbl_info_cortes = tk.Label(f_btn_cortes, text="Sin cortes", font=("Arial", 8), fg="#555555")
        self.lbl_info_cortes.pack(side="left", padx=10)
        self.fig_cortes = plt.figure(figsize=(9, 6))
        self.canvas_cortes = FigureCanvasTkAgg(self.fig_cortes, master=self.tab_cortes)
        self.canvas_cortes.get_tk_widget().pack(fill="both", expand=True)

    # ================= EXTRACCION DESDE ETABS =================

    def _obtener_vigas_modelo(self):
        ret = self.SapModel.FrameObj.GetNameList()
        if not ret or ret[-1] != 0: return []
        vigas = []
        for name in ret[1]:
            name = str(name)
            orient = None
            try:
                ro = self.SapModel.FrameObj.GetDesignOrientation(name)
                if ro and ro[-1] == 0: orient = int(ro[0])
            except: pass
            if orient is None:
                try:
                    rp = self.SapModel.FrameObj.GetPoints(name)
                    c1 = self.SapModel.PointObj.GetCoordCartesian(str(rp[0]))
                    c2 = self.SapModel.PointObj.GetCoordCartesian(str(rp[1]))
                    if abs(float(c1[2]) - float(c2[2])) < 0.1 and abs(float(c2[0])-float(c1[0])) + abs(float(c2[1])-float(c1[1])) > 0.5:
                        orient = 1
                except: pass
            if orient == 1: vigas.append(name)
        return vigas

    def _agrupar_en_ejes(self, vigas):
        info = {}
        for v in vigas:
            try:
                rp = self.SapModel.FrameObj.GetPoints(v)
                p1, p2 = str(rp[0]), str(rp[1])
                c1 = self.SapModel.PointObj.GetCoordCartesian(p1)
                c2 = self.SapModel.PointObj.GetCoordCartesian(p2)
                info[v] = {'p1': p1, 'p2': p2,
                           'x1': float(c1[0]), 'y1': float(c1[1]), 'z': float(c1[2]),
                           'x2': float(c2[0]), 'y2': float(c2[1])}
            except: pass
        
        nodos = {}
        for v, d in info.items():
            nodos.setdefault(d['p1'], []).append(v)
            nodos.setdefault(d['p2'], []).append(v)
        
        visitados = set()
        cadenas = []
        for v in info:
            if v in visitados: continue
            cola = [v]; visitados.add(v); cadena = []
            while cola:
                actual = cola.pop()
                cadena.append(actual)
                d = info[actual]
                dx, dy = d['x2']-d['x1'], d['y2']-d['y1']
                L = math.hypot(dx, dy) or 1.0
                ux, uy = dx/L, dy/L
                for nodo in (d['p1'], d['p2']):
                    for vecino in nodos.get(nodo, []):
                        if vecino in visitados or vecino == actual: continue
                        dv = info[vecino]
                        dxv, dyv = dv['x2']-dv['x1'], dv['y2']-dv['y1']
                        Lv = math.hypot(dxv, dyv) or 1.0
                        uvx, uvy = dxv/Lv, dyv/Lv
                        if abs(abs(ux*uvx + uy*uvy) - 1.0) < 1e-3:
                            visitados.add(vecino)
                            cola.append(vecino)
            cadenas.append(cadena)
        
        ejes = []
        for cadena in cadenas:
            frame_data = []
            for f in cadena:
                d = info[f]
                L_eje = math.hypot(d['x2']-d['x1'], d['y2']-d['y1']) / 1000.0
                frame_data.append({'name': f, 'xc': (d['x1']+d['x2'])/2, 'yc': (d['y1']+d['y2'])/2,
                                   'L': L_eje, 'x1': d['x1'], 'x2': d['x2'], 'y1': d['y1'], 'y2': d['y2'],
                                   'p1': d['p1'], 'p2': d['p2']})
            if not frame_data: continue
            dx = max(f['xc'] for f in frame_data) - min(f['xc'] for f in frame_data)
            dy = max(f['yc'] for f in frame_data) - min(f['yc'] for f in frame_data)
            
            if dx >= dy: 
                frame_data.sort(key=lambda item: item['xc'])
                for f in frame_data: f['rev'] = f['x1'] > f['x2']
            else: 
                frame_data.sort(key=lambda item: item['yc'])
                for f in frame_data: f['rev'] = f['y1'] > f['y2']
            
            labels = []
            story = ""
            for f in frame_data:
                dl = self.SapModel.FrameObj.GetLabelFromName(f['name'])
                if dl and dl[-1] == 0:
                    labels.append(str(dl[0]))
                    if not story: story = str(dl[1]) if len(dl) > 1 else ""
                else:
                    labels.append(f['name'])
            
            nombre = f"[{story}] " + "-".join(dict.fromkeys(labels)) if story else "-".join(labels)
            ejes.append((nombre, frame_data))
        
        ejes.sort(key=lambda t: (round(min(f['yc'] for f in t[1])/1000.0, 1), min(f['xc'] for f in t[1])))
        return ejes

    def _obtener_combos_diseno_reales(self, vigas):
        """Extrae los combos de diseño reales leyendo el reporte de las vigas directamente"""
        design_combos = set()
        
        # 1. Intentar obtener de GetComboDef (Combos explícitos de diseño de concreto)
        try:
            ret_dcon = self.SapModel.DesignConcrete.GetComboDef()
            if ret_dcon and ret_dcon[-1] == 0 and ret_dcon[1]:
                for c in ret_dcon[1]:
                    if str(c).strip(): design_combos.add(str(c).strip())
        except: pass
        
        # 2. Extraer de los resultados de diseño de las vigas seleccionadas (infalible)
        for v in vigas:
            try:
                res = self.SapModel.DesignConcrete.GetSummaryResultsBeam(v)
                if res and res[-1] == 0:
                    for array in res:
                        if isinstance(array, (tuple, list)):
                            for item in array:
                                if isinstance(item, str):
                                    val = item.strip()
                                    if val and val != v and not val.replace('.', '').replace('-', '').isdigit():
                                        design_combos.add(val)
            except: pass
            
        # 3. Fallback a RespCombo solo si no se encontró nada (y filtrando 'envelope')
        if not design_combos:
            try:
                ret_nc = self.SapModel.RespCombo.GetNameList()
                if ret_nc and ret_nc[-1] == 0 and ret_nc[1]:
                    for c in ret_nc[1]:
                        val = str(c).strip()
                        if val and "envelope" not in val.lower():
                            design_combos.add(val)
            except: pass
            
        return sorted(list(design_combos))

    def _activar_combos_output(self, combos):
        """Activa los combos para salida de resultados de FrameForce"""
        activados = []
        try:
            self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            for combo in combos:
                try:
                    if self.SapModel.Results.Setup.SetComboSelectedForOutput(combo, True) == 0:
                        activados.append(combo)
                except:
                    try:
                        if self.SapModel.Results.Setup.SetComboSelectedForOutput(combo) == 0:
                            activados.append(combo)
                    except: pass
            self.SapModel.Results.Setup.SetOptionMultiValuedCombo(2) 
        except: pass
        return activados

    def extraer_diseno(self):
        if not self.SapModel: return
        try:
            res_sel = self.SapModel.SelectObj.GetSelected()
            if len(res_sel) < 4 or res_sel[-1] != 0 or res_sel[0] == 0:
                return messagebox.showwarning("Aviso", "Selecciona los tramos en ETABS.")
            
            num_sel, obj_types, obj_names = res_sel[0], res_sel[1], res_sel[2]
            vigas_sel = [str(obj_names[i]) for i in range(num_sel) if obj_types[i] == 2]
            if not vigas_sel:
                return messagebox.showwarning("Aviso", "Selecciona tramos de VIGA en ETABS.")
            
            combos = self._obtener_combos_diseno_reales(vigas_sel)
            frame_data = []
            for f in vigas_sel:
                rp = self.SapModel.FrameObj.GetPoints(f)
                p1, p2 = str(rp[0]), str(rp[1])
                x1, y1, z1 = [float(c) for c in self.SapModel.PointObj.GetCoordCartesian(p1)[:3]]
                x2, y2, z2 = [float(c) for c in self.SapModel.PointObj.GetCoordCartesian(p2)[:3]]
                L_eje = math.sqrt((x2-x1)**2 + (y2-y1)**2 + (z2-z1)**2) / 1000.0
                frame_data.append({'name': f, 'xc': (x1+x2)/2, 'yc': (y1+y2)/2, 'L': L_eje,
                                   'x1': x1, 'x2': x2, 'y1': y1, 'y2': y2, 'p1': p1, 'p2': p2})
            
            dx = max(f['xc'] for f in frame_data) - min(f['xc'] for f in frame_data)
            dy = max(f['yc'] for f in frame_data) - min(f['yc'] for f in frame_data)
            
            if dx >= dy: 
                frame_data.sort(key=lambda item: item['xc'])
                for f in frame_data: f['rev'] = f['x1'] > f['x2']
            else: 
                frame_data.sort(key=lambda item: item['yc'])
                for f in frame_data: f['rev'] = f['y1'] > f['y2']
            
            labels = []
            for f in frame_data:
                dl = self.SapModel.FrameObj.GetLabelFromName(f['name'])
                labels.append(str(dl[0]) if dl and dl[-1] == 0 else f['name'])
            nombre = "Eje: " + "-".join(labels)
            
            datos = self._ensamblar_eje(nombre, frame_data, combos)
            if datos:
                self.ejes_extraidos[nombre] = datos
                self.cmb_ejes['values'] = list(self.ejes_extraidos.keys())
                self.cmb_ejes.set(nombre)
                self.aplicar_eje(datos)
                messagebox.showinfo("Éxito", f"Eje ensamblado correctamente. \nSe detectaron {len(combos)} combos de diseño.")
        except Exception as e: messagebox.showerror("Error", str(e))

    def extraer_todo_modelo(self):
        if not self.SapModel: return
        try:
            vigas = self._obtener_vigas_modelo()
            if not vigas:
                return messagebox.showwarning("Aviso", "No se encontraron vigas en el modelo.")
            
            ejes = self._agrupar_en_ejes(vigas)
            combos = self._obtener_combos_diseno_reales(vigas[:20])
            
            self.ejes_extraidos.clear()
            errores = 0
            for nombre, frame_data in ejes:
                try:
                    datos = self._ensamblar_eje(nombre, frame_data, combos)
                    if datos: self.ejes_extraidos[nombre] = datos
                except Exception:
                    errores += 1
            
            self.cmb_ejes['values'] = list(self.ejes_extraidos.keys())
            if self.ejes_extraidos:
                self.cmb_ejes.current(0)
                self.cargar_eje_combo()
                msg = f"Se extrajeron {len(self.ejes_extraidos)} ejes constructivos.\n({errores} ejes con errores se omitieron)"
                messagebox.showinfo("Éxito", msg)
        except Exception as e: messagebox.showerror("Error", str(e))

    def _ensamblar_eje(self, nombre, frame_data, combos):
        """
        NÚCLEO ESTABLE: Usa FrameForce para las envolventes puras y GetSummaryResultsBeam para el acero (como en V6).
        """
        combos_usados = self._activar_combos_output(combos)
        
        estaciones, demanda_top, demanda_bot, demanda_v = [], [], [], []
        columnas_graf, vigas_graf = [], []
        env_raw, env_nsr = [], []
        offset_longitud = 0.0
        col_x_registrados = set()
        
        def obtener_seccion_columna(nodo, viga_actual):
            try:
                res_c = self.SapModel.PointObj.GetConnectivity(nodo)
                if res_c and res_c[-1] == 0:
                    for i in range(res_c[0]):
                        f_name = str(res_c[2][i])
                        if res_c[1][i] == 2 and f_name != viga_actual:
                            rp = self.SapModel.FrameObj.GetPoints(f_name)
                            rc1 = self.SapModel.PointObj.GetCoordCartesian(str(rp[0]))
                            rc2 = self.SapModel.PointObj.GetCoordCartesian(str(rp[1]))
                            if abs(float(rc1[2]) - float(rc2[2])) > 0.1: 
                                rc = self.SapModel.FrameObj.GetSection(f_name)
                                if rc and rc[-1] == 0: return str(rc[0])
            except: pass
            return ""
        
        def interp_m(x, xp, yp):
            if not xp or not yp: return 0.0
            for i in range(len(xp)-1):
                if xp[i] <= x <= xp[i+1] or abs(x - xp[i]) < 1e-4:
                    if abs(xp[i+1] - xp[i]) < 1e-4: return yp[i]
                    return yp[i] + (yp[i+1] - yp[i]) * (x - xp[i]) / (xp[i+1] - xp[i])
            if x < xp[0]: return yp[0]
            if x > xp[-1]: return yp[-1]
            return 0.0

        for tramo in frame_data:
            vu = tramo['name']
            L_tramo = tramo['L']
            
            prop_viga = "Viga"
            H_m, B_m = 0.45, 0.30
            fc_val, fy_val = 21.0, 420.0 
            try:
                ret_prof = self.SapModel.FrameObj.GetSection(vu)
                if ret_prof and ret_prof[-1] == 0:
                    prop_viga = str(ret_prof[0])
                    ret_mat = self.SapModel.PropFrame.GetMaterial(prop_viga)
                    if ret_mat and ret_mat[-1] == 0:
                        mat_conc = str(ret_mat[0])
                        ret_fc = self.SapModel.PropMaterial.GetOConcrete_1(mat_conc)
                        if ret_fc and ret_fc[-1] == 0: fc_val = float(ret_fc[0]) 
                    ret_rebar = self.SapModel.PropFrame.GetRebarBeam(prop_viga)
                    if ret_rebar and ret_rebar[-1] == 0:
                        mat_rebar = str(ret_rebar[0])
                        ret_fy = self.SapModel.PropMaterial.GetORebar_1(mat_rebar)
                        if ret_fy and ret_fy[-1] == 0: fy_val = float(ret_fy[0]) 
                    ret_rect = self.SapModel.PropFrame.GetRectangle(prop_viga)
                    if ret_rect and ret_rect[-1] == 0:
                        H_m = float(ret_rect[1]) / 1000.0
                        B_m = float(ret_rect[2]) / 1000.0
            except: pass

            prop_col_izq = obtener_seccion_columna(tramo['p1'], vu)
            prop_col_der = obtener_seccion_columna(tramo['p2'], vu)
            native_face_izq, native_face_der = 0.0, L_tramo
            ancho_col_izq, ancho_col_der = 0.0, 0.0
            
            try:
                res = self.SapModel.DesignConcrete.GetSummaryResultsBeam(vu)
                if res and len(res) >= 8 and res[-1] == 0 and res[0] != 0: 
                    o_sta = res[3] if isinstance(res[2][0], str) else res[2]
                    o_top = res[5] if isinstance(res[2][0], str) else res[4]
                    o_bot = res[7] if isinstance(res[2][0], str) else res[6]
                    o_v   = res[9] if isinstance(res[2][0], str) else res[8] 
                    
                    native_s_m = sorted(list(set([float(x)/1000.0 for x in o_sta])))
                    if native_s_m:
                        native_face_izq = native_s_m[0]
                        native_face_der = native_s_m[-1]
                        ancho_col_izq = native_face_izq * 2.0
                        ancho_col_der = (L_tramo - native_face_der) * 2.0

                    s_m = [float(x)/1000.0 for x in o_sta]
                    t_a, b_a, v_a = [float(x) for x in o_top], [float(x) for x in o_bot], [float(x) for x in o_v]
                    
                    if tramo['rev']: s_m = [L_tramo - x for x in s_m]
                    pts = sorted(list(zip(s_m, t_a, b_a, v_a)), key=lambda p: p[0])
                    s_m, t_a, b_a, v_a = [p[0] for p in pts], [p[1] for p in pts], [p[2] for p in pts], [p[3] for p in pts]
                    
                    if len(s_m) > 0:
                        x_izq_global = round(offset_longitud, 3) 
                        if x_izq_global not in col_x_registrados:
                            columnas_graf.append({'x': offset_longitud, 'w': ancho_col_izq, 'prop': prop_col_izq})
                            col_x_registrados.add(x_izq_global)
                        x_der_global = round(offset_longitud + L_tramo, 3)
                        if x_der_global not in col_x_registrados:
                            columnas_graf.append({'x': offset_longitud + L_tramo, 'w': ancho_col_der, 'prop': prop_col_der})
                            col_x_registrados.add(x_der_global)
                            
                        vigas_graf.append({
                            'x_ini': offset_longitud, 'L': L_tramo,
                            'w_izq': ancho_col_izq, 'w_der': ancho_col_der,
                            'prop': prop_viga, 'H': H_m, 'B': B_m,
                            'fc': fc_val, 'fy': fy_val
                        })
                        for i in range(len(s_m)):
                            estaciones.append(s_m[i] + offset_longitud)
                            demanda_top.append(t_a[i])
                            demanda_bot.append(b_a[i])
                            demanda_v.append(v_a[i] * 1000.0) 
            except: pass

            frame_sta, m_max_raw, m_min_raw = [], [], []
            v_max_raw, v_min_raw = [], []
            p_max_raw, p_min_raw = [], [] 
            try:
                ret_forces = self.SapModel.Results.FrameForce(vu, 0)
                if ret_forces and len(ret_forces) >= 14 and ret_forces[-1] == 0:
                    obj_sta = ret_forces[2]
                    p_vals, v2_vals, m3_vals = ret_forces[8], ret_forces[9], ret_forces[13]
                    
                    env_dict = {}
                    for s, p, v, m in zip(obj_sta, p_vals, v2_vals, m3_vals):
                        s_r = round(float(s) / 1000.0, 4)
                        p_kn, v_kn, m_knm = float(p)/1000.0, float(v)/1000.0, float(m)/1e6   
                        if s_r not in env_dict: 
                            env_dict[s_r] = {'m_max': m_knm, 'm_min': m_knm, 'v_max': v_kn, 'v_min': v_kn, 'p_max': p_kn, 'p_min': p_kn}
                        else:
                            e = env_dict[s_r]
                            e['m_max'] = max(e['m_max'], m_knm); e['m_min'] = min(e['m_min'], m_knm)
                            e['v_max'] = max(e['v_max'], v_kn); e['v_min'] = min(e['v_min'], v_kn)
                            e['p_max'] = max(e['p_max'], p_kn); e['p_min'] = min(e['p_min'], p_kn)
                    
                    for k in sorted(env_dict.keys()):
                        frame_sta.append(k)
                        m_max_raw.append(env_dict[k]['m_max']); m_min_raw.append(env_dict[k]['m_min'])
                        v_max_raw.append(env_dict[k]['v_max']); v_min_raw.append(env_dict[k]['v_min'])
                        p_max_raw.append(env_dict[k]['p_max']); p_min_raw.append(env_dict[k]['p_min'])
            except: pass

            if frame_sta:
                top_i_base = interp_m(native_face_izq, frame_sta, m_min_raw)
                top_j_base = interp_m(native_face_der, frame_sta, m_min_raw)
                bot_i_base = interp_m(native_face_izq, frame_sta, m_max_raw)
                bot_j_base = interp_m(native_face_der, frame_sta, m_max_raw)
                top_m_base = interp_m((native_face_izq + native_face_der)/2.0, frame_sta, m_min_raw)
                bot_m_base = interp_m((native_face_izq + native_face_der)/2.0, frame_sta, m_max_raw)
                
                L_libre = native_face_der - native_face_izq
                
                # Mínimos normativos DES NSR-10
                max_m_ends = max(abs(top_i_base), abs(top_j_base), abs(bot_i_base), abs(bot_j_base))
                req_global_cuarto = 0.25 * max_m_ends
                req_bot_i_mitad = 0.5 * abs(top_i_base)
                req_bot_j_mitad = 0.5 * abs(top_j_base)
                
                tramo_pts_raw = {'x': [], 'm_pos': [], 'm_neg': [], 'v_pos': [], 'v_neg': [], 'v_grav': [], 'pu': 0.0}
                tramo_pts_nsr = {'x': [], 'm_pos': [], 'm_neg': [], 'v_pos': [], 'v_neg': [], 'v_grav': [], 'pu': 0.0}
                pu_max = 0.0
                
                for st, mx, mn, vx, vn, pv in zip(frame_sta, m_max_raw, m_min_raw, v_max_raw, v_min_raw, [max(a,b) for a,b in zip(p_max_raw, p_min_raw)]):
                    if st < native_face_izq - 1e-4 or st > native_face_der + 1e-4: continue
                    s_actual = (L_tramo - st) if tramo['rev'] else st
                    
                    v_pos = (-vn) if tramo['rev'] else vx
                    v_neg = (-vx) if tramo['rev'] else vn
                    
                    # Extraer V_gravitacional puro (el promedio de la envolvente anula el sismo)
                    v_grav = (vx + vn) / 2.0
                    if tramo['rev']: v_grav = -v_grav
                    
                    tramo_pts_raw['x'].append(s_actual + offset_longitud)
                    tramo_pts_raw['m_pos'].append(max(0.0, mx)); tramo_pts_raw['m_neg'].append(min(0.0, mn))
                    tramo_pts_raw['v_pos'].append(v_pos); tramo_pts_raw['v_neg'].append(v_neg)
                    tramo_pts_raw['v_grav'].append(v_grav)
                    tramo_pts_raw['pu'] = max(tramo_pts_raw['pu'], pv)
                    
                    mx_n = max(mx, req_global_cuarto); mn_n = min(mn, -req_global_cuarto)
                    if L_libre > 0:
                        native_mid = (native_face_izq + native_face_der) / 2.0
                        if st <= native_mid:
                            taper_n = req_bot_i_mitad * (1.0 - ((st - native_face_izq) / (L_libre / 2.0))) if (L_libre/2.0) > 1e-9 else 0
                        else:
                            taper_n = req_bot_j_mitad * (1.0 - ((native_face_der - st) / (L_libre / 2.0))) if (L_libre/2.0) > 1e-9 else 0
                        if taper_n > 0: mx_n = max(mx_n, taper_n)
                        
                    tramo_pts_nsr['x'].append(s_actual + offset_longitud)
                    tramo_pts_nsr['m_pos'].append(max(0.0, mx_n)); tramo_pts_nsr['m_neg'].append(min(0.0, mn_n))
                    tramo_pts_nsr['v_pos'].append(v_pos); tramo_pts_nsr['v_neg'].append(v_neg)
                    tramo_pts_nsr['v_grav'].append(v_grav)
                    tramo_pts_nsr['pu'] = max(tramo_pts_nsr['pu'], pv)
                
                if tramo['rev']:
                    text_top_raw, text_bot_raw = [top_j_base, top_m_base, top_i_base], [bot_j_base, bot_m_base, bot_i_base]
                    text_top_nsr = [min(top_j_base, -req_global_cuarto), min(top_m_base, -req_global_cuarto), min(top_i_base, -req_global_cuarto)]
                    bot_j_nsr = max(bot_j_base, req_global_cuarto, req_bot_j_mitad)
                    bot_i_nsr = max(bot_i_base, req_global_cuarto, req_bot_i_mitad)
                    bot_m_nsr = max(bot_m_base, req_global_cuarto)
                    text_bot_nsr = [bot_j_nsr, bot_m_nsr, bot_i_nsr]
                else:
                    text_top_raw, text_bot_raw = [top_i_base, top_m_base, top_j_base], [bot_i_base, bot_m_base, bot_j_base]
                    text_top_nsr = [min(top_i_base, -req_global_cuarto), min(top_m_base, -req_global_cuarto), min(top_j_base, -req_global_cuarto)]
                    bot_i_nsr = max(bot_i_base, req_global_cuarto, req_bot_i_mitad)
                    bot_j_nsr = max(bot_j_base, req_global_cuarto, req_bot_j_mitad)
                    bot_m_nsr = max(bot_m_base, req_global_cuarto)
                    text_bot_nsr = [bot_i_nsr, bot_m_nsr, bot_j_nsr]
                
                for destino, pts_d, tt, tb in [(env_raw, tramo_pts_raw, text_top_raw, text_bot_raw),
                                               (env_nsr, tramo_pts_nsr, text_top_nsr, text_bot_nsr)]:
                    destino.append({
                        'x': pts_d['x'], 'm_pos': pts_d['m_pos'], 'm_neg': pts_d['m_neg'],
                        'v_pos': pts_d['v_pos'], 'v_neg': pts_d['v_neg'],
                        'v_grav': pts_d['v_grav'],
                        'text_top': tt, 'text_bot': tb,
                        'pu': pts_d['pu']       
                    })

            offset_longitud += L_tramo 
        
        if not estaciones: return None
        
        return {
            'eje': nombre, 'L_total': offset_longitud,
            'estaciones': estaciones, 'demanda_top': demanda_top, 'demanda_bot': demanda_bot,
            'demanda_v_req': demanda_v,
            'columnas': columnas_graf, 'vigas': vigas_graf,
            'env_raw': env_raw, 'env_nsr': env_nsr,
            'envolventes_tramos': env_raw,
            'combos_usados': combos_usados,
            'frame_data': frame_data,
        }

    def recalcular_con_combos(self):
        label = self.viga_label
        datos_previos = self.ejes_extraidos.get(label)
        if not datos_previos or not datos_previos.get('frame_data'):
            return messagebox.showwarning("Aviso", "Extrae primero un eje desde ETABS.")
        if not self.SapModel:
            return messagebox.showwarning("Aviso", "No hay conexión con ETABS.")
        
        seleccion = [self.lst_combos.get(i) for i in self.lst_combos.curselection()]
        if not seleccion:
            return messagebox.showwarning("Aviso", "Marca al menos un combo en la lista.")
        
        try:
            datos = self._ensamblar_eje(label, datos_previos['frame_data'], seleccion)
            if not datos:
                return messagebox.showerror("Error", "No se pudieron extraer resultados con esos combos.")
            
            barras_previas = datos_previos.get('barras')
            self.ejes_extraidos[label] = datos
            if barras_previas is not None:
                datos['barras'] = barras_previas
            
            n_combos = len(datos.get('combos_usados', []))
            self._actualizar_lista_combos(seleccion)
            self.aplicar_eje(datos)
            messagebox.showinfo("Éxito", f"Envolvente recalculada con {n_combos} combo(s).")
        except Exception as e: messagebox.showerror("Error", str(e))

    def _actualizar_lista_combos(self, combos):
        self.lst_combos.delete(0, tk.END)
        for c in combos:
            self.lst_combos.insert(tk.END, c)
        self.lst_combos.select_set(0, 'end')
        self.lbl_combos.config(text=f"{len(combos)} combos del modelo (marcados = incluidos en la envolvente)")

    def cargar_eje_combo(self):
        self._guardar_estado_en_registro()
        label = self.cmb_ejes.get()
        if label in self.ejes_extraidos:
            self.aplicar_eje(self.ejes_extraidos[label])

    def _guardar_estado_en_registro(self):
        if self.viga_label and self.viga_label in self.ejes_extraidos:
            self.ejes_extraidos[self.viga_label]['barras'] = [dict(b) for b in self.barras]

    def aplicar_eje(self, datos):
        self.viga_label = datos['eje']
        self.longitud_eje_real = datos['L_total']
        self.estaciones = list(datos['estaciones'])
        self.demanda_top = list(datos['demanda_top'])
        self.demanda_bot = list(datos['demanda_bot'])
        self.demanda_v_req = list(datos.get('demanda_v_req', []))
        self.columnas_graficas = [dict(c) for c in datos['columnas']]
        self.vigas_graficas = [dict(v) for v in datos['vigas']]
        self.envolventes_tramos = datos['env_raw'] if not self.aplicar_nsr_envolvente.get() else datos['env_nsr']
        
        previas = datos.get('barras')
        if previas is not None:
            self.barras = [dict(b) for b in previas]
            self.bar_id_counter = max([b['id'] for b in self.barras], default=0) + 1
        else:
            self.barras.clear()
            self.bar_id_counter = 1
        
        self.refresh_tree()
        self.actualizar_grafica()
        self.actualizar_verificacion_cortante()
        if datos.get('combos_usados') is not None:
            self._actualizar_lista_combos(datos['combos_usados'])

    def refrescar_envolvente_activa(self):
        label = self.viga_label
        if label in self.ejes_extraidos:
            datos = self.ejes_extraidos[label]
            self.envolventes_tramos = datos['env_raw'] if not self.aplicar_nsr_envolvente.get() else datos['env_nsr']
            self.actualizar_verificacion_flexion()
            self.plotear_cortante()

    # ================= MODULACION Y DESPIECE =================

    def generar_continuo(self):
        if self.longitud_eje_real == 0: return
        try:
            L_com = float(self.ent_lcom.get())
            L_tras = float(self.ent_ltras.get())
            L_g_val = float(self.ent_gancho_auto.get())
            L_tot = self.longitud_eje_real
            
            self.barras = [b for b in self.barras if b['rol'] != 'Continuo']
            
            for cara, cant, tipo in [("Superior", int(self.spn_a_csup.get()), self.cmb_a_tsup.get()),
                                     ("Inferior", int(self.spn_a_cinf.get()), self.cmb_a_tinf.get())]:
                pos_x = 0.0
                is_first = True
                while pos_x < L_tot:
                    g_i = L_g_val if is_first else 0.0
                    
                    if pos_x + L_com - g_i >= L_tot:
                        g_f = L_g_val
                        self.create_bar_dict(cara, "Continuo", cant, tipo, pos_x, L_tot, g_i, g_f)
                        break
                    else:
                        end_x = pos_x + L_com - g_i
                        self.create_bar_dict(cara, "Continuo", cant, tipo, pos_x, end_x, g_i, 0.0)
                        pos_x = end_x - L_tras 
                    is_first = False
                        
            self.refresh_tree()
            self.actualizar_grafica()
        except ValueError: messagebox.showerror("Error", "Parámetros inválidos")

    def create_bar_dict(self, cara, rol, cant, tipo, x1, x2, g_ini, g_fin):
        b = {'id': self.bar_id_counter, 'cara': cara, 'rol': rol, 'cant': cant, 'tipo': tipo, 
             'area': cant * self.rebar_db[tipo], 'x1': x1, 'x2': x2, 'g_ini': g_ini, 'g_fin': g_fin}
        self.barras.append(b)
        self.bar_id_counter += 1

    def add_bar(self):
        try:
            x1, x2 = float(self.ent_x1.get()), float(self.ent_x2.get())
            g1, g2 = float(self.ent_g1.get()), float(self.ent_g2.get())
            if x1 >= x2: return messagebox.showwarning("Error", "X1 debe ser menor a X2")
            
            self.create_bar_dict(self.cmb_cara.get(), self.cmb_rol.get(), int(self.spn_cant.get()), self.cmb_tipo.get(), x1, x2, g1, g2)
            self.refresh_tree()
            self.actualizar_grafica()
        except ValueError: messagebox.showwarning("Error", "Datos numéricos inválidos.")

    def update_bar(self):
        sel = self.tree.selection()
        if not sel: return
        try:
            b_id = int(self.tree.item(sel[0])['values'][0])
            x1, x2 = float(self.ent_x1.get()), float(self.ent_x2.get())
            g1, g2 = float(self.ent_g1.get()), float(self.ent_g2.get())
            if x1 >= x2: return messagebox.showwarning("Error", "X1 debe ser menor a X2")
            
            for b in self.barras:
                if b['id'] == b_id:
                    b['cara'] = self.cmb_cara.get()
                    b['rol'] = self.cmb_rol.get()
                    b['cant'] = int(self.spn_cant.get())
                    b['tipo'] = self.cmb_tipo.get()
                    b['area'] = b['cant'] * self.rebar_db[b['tipo']]
                    b['x1'] = x1; b['x2'] = x2; b['g_ini'] = g1; b['g_fin'] = g2
                    break
            self.refresh_tree()
            self.actualizar_grafica()
        except ValueError: messagebox.showwarning("Error", "Valores inválidos.")

    def delete_bar(self):
        sel = self.tree.selection()
        if not sel: return
        b_id = int(self.tree.item(sel[0])['values'][0])
        self.barras = [b for b in self.barras if b['id'] != b_id]
        self.refresh_tree()
        self.actualizar_grafica()

    def on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])['values']
        self.cmb_cara.set(vals[1])
        self.cmb_rol.set(vals[2])
        self.spn_cant.delete(0, tk.END); self.spn_cant.insert(0, vals[3].split()[0])
        self.cmb_tipo.set(vals[3].split()[1])
        self.ent_x1.delete(0, tk.END); self.ent_x1.insert(0, vals[4])
        self.ent_x2.delete(0, tk.END); self.ent_x2.insert(0, vals[5])
        self.ent_g1.delete(0, tk.END); self.ent_g1.insert(0, vals[6])
        self.ent_g2.delete(0, tk.END); self.ent_g2.insert(0, vals[7])

    def refresh_tree(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for b in sorted(self.barras, key=lambda x: (x['cara'], x['x1'])):
            L_real = b['x2'] - b['x1'] + b.get('g_ini', 0) + b.get('g_fin', 0)
            self.tree.insert("", tk.END, values=(
                b['id'], b['cara'], b['rol'], f"{b['cant']} {b['tipo']}", 
                f"{b['x1']:.2f}", f"{b['x2']:.2f}", f"{b.get('g_ini', 0):.2f}", f"{b.get('g_fin', 0):.2f}", f"{L_real:.2f}"
            ))

    # ================= GRAFICO PRINCIPAL =================

    def actualizar_grafica(self):
        if not self.estaciones: return
        self.ax_top.clear()
        self.ax_barras.clear()
        self.ax_bot.clear()
        
        self.ax_top.set_title(f"Detallado Estructural - {self.viga_label} (L= {self.longitud_eje_real:.2f}m)", fontweight='bold')
        self.ax_barras.set_title("Esquema Físico de Barras (Elevación)", fontsize=10, fontweight='bold', color='gray')
        
        y_max_top = max(self.demanda_top) * 1.3 if self.demanda_top else 500
        
        for col in self.columnas_graficas:
            xs, w = col['x'] - col['w']/2, col['w']
            self.ax_top.axvspan(xs, xs + w, color='gray', alpha=0.25)
            self.ax_barras.axvspan(xs, xs + w, color='gray', alpha=0.15)
            self.ax_bot.axvspan(xs, xs + w, color='gray', alpha=0.25)
            
            if col.get('prop'):
                self.ax_top.text(col['x'], y_max_top * 1.05, col['prop'], fontsize=8, fontweight='bold', color='#444444', ha='center', bbox=dict(facecolor='#f5f5f5', alpha=0.8, edgecolor='none', pad=2))

        for v in self.vigas_graficas:
            x_face_izq = v['x_ini'] + v['w_izq']/2
            x_face_der = v['x_ini'] + v['L'] - v['w_der']/2
            
            if self.mostrar_confinamiento.get():
                if v['w_izq'] > 0.01:
                    z_conf_izq = min(x_face_izq + (2.0 * v['H']), (x_face_izq + x_face_der)/2)
                    self.ax_barras.axvspan(x_face_izq, z_conf_izq, color='#ff9933', alpha=0.08, hatch='//', edgecolor='none')
                    self.ax_barras.vlines([x_face_izq, z_conf_izq], -3.5, 3.5, colors='#cc6600', linestyles=':', alpha=0.4)
                    self.ax_barras.text(x_face_izq, -3.3, f"{x_face_izq:.2f}", color='#cc6600', fontsize=7.5, ha='center', va='bottom', fontweight='bold')
                    self.ax_barras.text(z_conf_izq, -3.3, f"{z_conf_izq:.2f}", color='#cc6600', fontsize=7.5, ha='center', va='bottom', fontweight='bold')

                if v['w_der'] > 0.01:
                    z_conf_der = max(x_face_der - (2.0 * v['H']), (x_face_izq + x_face_der)/2)
                    self.ax_barras.axvspan(z_conf_der, x_face_der, color='#ff9933', alpha=0.08, hatch='//', edgecolor='none')
                    self.ax_barras.vlines([z_conf_der, x_face_der], -3.5, 3.5, colors='#cc6600', linestyles=':', alpha=0.4)
                    self.ax_barras.text(z_conf_der, -3.3, f"{z_conf_der:.2f}", color='#cc6600', fontsize=7.5, ha='center', va='bottom', fontweight='bold')
                    self.ax_barras.text(x_face_der, -3.3, f"{x_face_der:.2f}", color='#cc6600', fontsize=7.5, ha='center', va='bottom', fontweight='bold')

            x_center_vano = (x_face_izq + x_face_der) / 2
            self.ax_top.text(x_center_vano, y_max_top * 0.88, f"Viga: {v['prop']} (H={v['H']:.2f}m)", fontsize=8, fontweight='bold', color='blue', ha='center', bbox=dict(facecolor='white', alpha=0.8, edgecolor='blue', linewidth=0.5, pad=2))

        self.ax_top.plot(self.estaciones, self.demanda_top, 'r-', linewidth=1.5, zorder=5)
        self.ax_top.fill_between(self.estaciones, self.demanda_top, color='red', alpha=0.1)
        self.ax_bot.plot(self.estaciones, self.demanda_bot, 'r-', linewidth=1.5, zorder=5)
        self.ax_bot.fill_between(self.estaciones, self.demanda_bot, color='red', alpha=0.1)
        
        pts_x = [i * (self.longitud_eje_real / 500.0) for i in range(501)]
        cap_t_cont, cap_t_bast = [0]*501, [0]*501
        cap_b_cont, cap_b_bast = [0]*501, [0]*501
        
        for b in self.barras:
            for i, x in enumerate(pts_x):
                if b['x1'] <= x <= b['x2']:
                    if b['cara'] == "Superior":
                        if b['rol'] == "Continuo": cap_t_cont[i] = max(cap_t_cont[i], b['area'])
                        else: cap_t_bast[i] += b['area']
                    else:
                        if b['rol'] == "Continuo": cap_b_cont[i] = max(cap_b_cont[i], b['area'])
                        else: cap_b_bast[i] += b['area']

        cap_t = [c + b for c, b in zip(cap_t_cont, cap_t_bast)]
        cap_b = [c + b for c, b in zip(cap_b_cont, cap_b_bast)]

        self.ax_top.fill_between(pts_x, cap_t, step="mid", color='blue', alpha=0.15, zorder=2)
        self.ax_top.step(pts_x, cap_t, where="mid", color='blue', linewidth=1.5, label="Cap. Total (Sin picos)", zorder=3)
        self.ax_bot.fill_between(pts_x, cap_b, step="mid", color='green', alpha=0.15, zorder=2)
        self.ax_bot.step(pts_x, cap_b, where="mid", color='green', linewidth=1.5, label="Cap. Total (Sin picos)", zorder=3)
        
        # Flejes dibujados como lineas verticales por zona
        for z in self.datos_cortante:
            s_m = max(z['s_col_cm'] / 100.0, 0.02)
            n_lines = min(int((z['x2'] - z['x1']) / s_m) + 1, 400)
            for k in range(n_lines):
                xf = z['x1'] + k * s_m
                if xf > z['x2']: break
                self.ax_barras.vlines(xf, -1.6, 1.6, colors='#8e44ad', linewidth=0.6, alpha=0.7)
        
        self.ax_barras.axhline(0, color='black', linestyle='-.', alpha=0.3) 
        
        # --- Fase 1: guardar artistas Line2D para drag visual (Opción A) ---
        self._bar_artists = {}
        for b in self.barras:
            c_fact = 1 if b['cara'] == "Superior" else -1
            r_fact = 2 if b['rol'] == "Continuo" else 1
            
            y_base = c_fact * r_fact
            y_off = c_fact * 0.4 * (b['id'] % 2) 
            y_pos = y_base + y_off
            
            c = 'blue' if b['cara'] == 'Superior' else 'green'
            ls = '-' if b['rol'] == "Continuo" else '--'
            
            line, = self.ax_barras.plot([b['x1'], b['x2']], [y_pos, y_pos], color=c, linestyle=ls, linewidth=2.5, zorder=4)
            
            hook_ini = None
            hook_fin = None
            if b.get('g_ini', 0) > 0:
                hook_ini, = self.ax_barras.plot([b['x1'], b['x1']], [y_pos, y_pos - c_fact*0.5], color=c, linewidth=2.5, zorder=4)
            if b.get('g_fin', 0) > 0:
                hook_fin, = self.ax_barras.plot([b['x2'], b['x2']], [y_pos, y_pos - c_fact*0.5], color=c, linewidth=2.5, zorder=4)
                
            L_real = b['x2'] - b['x1'] + b.get('g_ini', 0) + b.get('g_fin', 0)
            txt = self.ax_barras.text((b['x1']+b['x2'])/2, y_pos + c_fact*0.2, f"{b['cant']}{b['tipo']} (Lc={L_real:.2f}m)", fontsize=7, color=c, ha='center', va='center', fontweight='bold')
            self._bar_artists[b['id']] = {'line': line, 'hook_ini': hook_ini, 'hook_fin': hook_fin, 'text': txt, 'y_pos': y_pos, 'c_fact': c_fact, 'c': c}

        # --- Directriz 1.1: handles de barra seleccionada ---
        self._handle_artists = getattr(self, '_handle_artists', [])
        for h in list(self._handle_artists):
            try: h.remove()
            except Exception: pass
        self._handle_artists.clear()
        if getattr(self, '_selected_bar_id', None) is not None:
            b_sel = next((x for x in self.barras if x['id'] == self._selected_bar_id), None)
            if b_sel and b_sel['id'] in self._bar_artists:
                y_sel = self._bar_artists[b_sel['id']]['y_pos']
                for xe in [float(b_sel['x1']), float(b_sel['x2'])]:
                    hh, = self.ax_barras.plot([xe], [y_sel], marker='s', markersize=8, markerfacecolor='red', markeredgecolor='black', markeredgewidth=1.2, linestyle='None', zorder=7, picker=False)
                    self._handle_artists.append(hh)

        # --- Directriz 1.2: cotas traslapo continuo (visual) ---
        try:
            for cara in ["Superior", "Inferior"]:
                cont = [b for b in self.barras if b['cara'] == cara and b['rol'] == 'Continuo']
                cont.sort(key=lambda x: float(x['x1']))
                for i in range(len(cont)):
                    for j in range(i+1, len(cont)):
                        b1, b2 = cont[i], cont[j]
                        ox1 = max(float(b1['x1']), float(b2['x1']))
                        ox2 = min(float(b1['x2']), float(b2['x2']))
                        Ltras = ox2 - ox1
                        if Ltras > 0.05:
                            mid = (ox1 + ox2) / 2.0
                            y_cota = 2.9 if cara == "Superior" else -2.9
                            va = 'bottom' if cara == "Superior" else 'top'
                            # linea de cota con flechas
                            self.ax_barras.annotate("", xy=(ox1, y_cota), xytext=(ox2, y_cota),
                                arrowprops=dict(arrowstyle='<->', color='#b34700', lw=1.2, shrinkA=2, shrinkB=2))
                            self.ax_barras.text(mid, y_cota + (0.15 if cara=="Superior" else -0.15), f"L_traslapo={Ltras:.2f}m",
                                ha='center', va=va, fontsize=7, color='#b34700', fontweight='bold',
                                bbox=dict(boxstyle="round,pad=0.2", fc="#fff2e6", ec="#b34700", alpha=0.9))
        except Exception:
            pass

        # --- Directriz 2.2: líneas de corte A-A en ax_barras ---
        try:
            if getattr(self, 'cortes', None):
                for corte in self.cortes:
                    xm = float(corte.get('x_mid', 0))
                    nombre = str(corte.get('nombre', ''))
                    if 0 <= xm <= self.longitud_eje_real + 1e-6:
                        self.ax_barras.axvline(xm, color='#e67e22', linestyle=':', linewidth=1.2, alpha=0.9, zorder=5)
                        self.ax_barras.text(xm, 3.3, nombre, fontsize=7, color='#e67e22', ha='center', va='bottom', fontweight='bold',
                                            bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="#e67e22", alpha=0.8))
        except Exception:
            pass

        self.ax_top.set_ylabel("Área Sup. (mm²)")
        self.ax_top.set_ylim(0, y_max_top * 1.15)
        self.ax_top.grid(True, linestyle=':', alpha=0.6)
        
        self.ax_barras.set_ylim(-3.5, 3.5)
        self.ax_barras.set_yticks([])
        self.ax_barras.set_ylabel("Despiece", fontweight='bold')
        
        self.ax_bot.set_ylabel("Área Inf. (mm²)")
        self.ax_bot.set_xlabel("Coordenada del Eje Constructivo (m)")
        self.ax_bot.grid(True, linestyle=':', alpha=0.6)
        self.ax_bot.invert_yaxis() 
        
        ticks = []
        val = 0.0
        while val <= self.longitud_eje_real + 0.01:
            ticks.append(round(val, 2))
            val += 0.5
            
        self.ax_bot.set_xticks(ticks)
        self.ax_bot.set_xticklabels([f"{t:.2f}" for t in ticks], rotation=45, fontsize=8)
        
        self.fig.tight_layout()
        self.canvas.draw()
        self.actualizar_verificacion_flexion()

    # ================= FASE 1: DRAG & DROP AX_BARRAS (corrección visual) =================
    def on_barras_press(self, event):
        if event.inaxes is not self.ax_barras or event.button != 1 or event.xdata is None:
            return
        if not self.barras:
            return
        x = float(event.xdata)
        y = float(event.ydata) if getattr(event, 'ydata', None) is not None else None
        tol = float(getattr(self, '_drag_tol', 0.15))
        # --- Directriz 1.1: prioridad a handles de barra seleccionada ---
        if getattr(self, '_selected_bar_id', None) is not None and self._selected_bar_id in getattr(self, '_bar_artists', {}):
            b_sel = next((b for b in self.barras if b['id'] == self._selected_bar_id), None)
            if b_sel is not None:
                y_sel = self._bar_artists[b_sel['id']].get('y_pos', 0)
                # hit test handles: distancia en X y Y
                for edge in ('x1', 'x2'):
                    xe = float(b_sel[edge])
                    dx = abs(x - xe)
                    dy = abs(y - y_sel) if y is not None else 0
                    # tolerancia en Y: 0.5 (espacio entre rieles ~0.4)
                    if dx <= tol and dy <= 0.6:
                        self._dragging_bar = b_sel
                        self._drag_edge = edge
                        try:
                            self.canvas.get_tk_widget().config(cursor="sb_h_double_arrow")
                        except Exception:
                            pass
                        art = self._bar_artists.get(b_sel['id'])
                        if art and art.get('line') is not None:
                            try:
                                art['line'].set_linewidth(4.0)
                                art['line'].set_alpha(0.9)
                                self.canvas.draw_idle()
                            except Exception:
                                pass
                        return
        # --- No fue handle: buscar barra cercana para seleccionar (hit en línea o en borde) ---
        best = None
        best_edge = None
        best_dist = tol + 1e-9
        # primero buscar proximidad a borde (para selección por borde)
        for b in self.barras:
            art = self._bar_artists.get(b['id'])
            if art is None:
                continue
            y_pos = art.get('y_pos', 0)
            dy = abs(y - y_pos) if y is not None else 0
            if dy > 0.6:
                continue
            for edge in ('x1', 'x2'):
                d = abs(x - float(b[edge]))
                if d < best_dist and d <= tol:
                    best_dist = d
                    best = b
                    best_edge = edge
        # si no hay borde cercano, buscar selección por cuerpo de barra (click en medio)
        if best is None:
            best_body = None
            best_body_dy = 1e9
            for b in self.barras:
                art = self._bar_artists.get(b['id'])
                if art is None:
                    continue
                y_pos = art.get('y_pos', 0)
                if not (min(float(b['x1']), float(b['x2'])) - tol <= x <= max(float(b['x1']), float(b['x2'])) + tol):
                    continue
                dy = abs(y - y_pos) if y is not None else 999
                if dy < 0.45 and dy < best_body_dy:
                    best_body_dy = dy
                    best_body = b
            if best_body is not None:
                best = best_body
                best_edge = None
        if best is not None:
            # Seleccionar barra (Directriz 1.1)
            prev_sel = getattr(self, '_selected_bar_id', None)
            self._selected_bar_id = best['id']
            # si hizo clic en borde y la barra ya estaba seleccionada, iniciar drag directamente (segundo clic)
            if best_edge is not None and prev_sel == best['id']:
                self._dragging_bar = best
                self._drag_edge = best_edge
                try:
                    self.canvas.get_tk_widget().config(cursor="sb_h_double_arrow")
                except Exception:
                    pass
                art = self._bar_artists.get(best['id'])
                if art and art.get('line') is not None:
                    try:
                        art['line'].set_linewidth(4.0)
                        art['line'].set_alpha(0.9)
                    except Exception:
                        pass
            # refrescar para dibujar handles
            try:
                self.actualizar_grafica()
            except Exception:
                try:
                    self.canvas.draw_idle()
                except Exception:
                    pass
            # si es selección por cuerpo (sin borde), no iniciar drag
            # si es borde pero primera selección, esperar siguiente clic en handle
            if best_edge is None:
                return
            if prev_sel != best['id']:
                # primera selección por borde: no arrastrar aún, solo mostrar handles
                return
            else:
                # ya estaba seleccionado y clic en handle/borde -> el return de handle arriba ya manejó, aquí es segundo caso
                self.canvas.draw_idle()
                return
        else:
            # Clic en fondo: deseleccionar
            if getattr(self, '_selected_bar_id', None) is not None:
                self._selected_bar_id = None
                # limpiar handles visuales
                for h in list(getattr(self, '_handle_artists', [])):
                    try: h.remove()
                    except Exception: pass
                self._handle_artists.clear()
                try:
                    self.actualizar_grafica()
                except Exception:
                    try: self.canvas.draw_idle()
                    except Exception: pass
            return

    def on_barras_motion(self, event):
        if self._dragging_bar is None or self._drag_edge is None:
            return
        if event.inaxes is not self.ax_barras or event.xdata is None:
            return
        b = self._dragging_bar
        snap = getattr(self, '_snap', 0.05)
        new_x = round(event.xdata / snap) * snap if snap and snap > 0 else event.xdata
        min_gap = 0.05
        if self._drag_edge == 'x1':
            new_x = min(new_x, b['x2'] - min_gap)
            b['x1'] = float(f"{new_x:.2f}")
        else:
            new_x = max(new_x, b['x1'] + min_gap)
            b['x2'] = float(f"{new_x:.2f}")
        # --- Solución visual Opción A: actualizar Line2D directamente ---
        art = self._bar_artists.get(b['id'])
        if art is not None:
            try:
                art['line'].set_xdata([b['x1'], b['x2']])
                if art.get('hook_ini') is not None:
                    art['hook_ini'].set_xdata([b['x1'], b['x1']])
                if art.get('hook_fin') is not None:
                    art['hook_fin'].set_xdata([b['x2'], b['x2']])
                # texto centrado y Lc
                if art.get('text') is not None:
                    mid = (b['x1'] + b['x2']) / 2.0
                    y = art.get('y_pos', art['text'].get_position()[1])
                    # mantener y_pos original
                    art['text'].set_position((mid, art['text'].get_position()[1]))
                    L_real = b['x2'] - b['x1'] + b.get('g_ini', 0) + b.get('g_fin', 0)
                    art['text'].set_text(f"{b['cant']}{b['tipo']} (Lc={L_real:.2f}m)")
                # sincronizar entries
                try:
                    self.ent_x1.delete(0, tk.END); self.ent_x1.insert(0, f"{b['x1']:.2f}")
                    self.ent_x2.delete(0, tk.END); self.ent_x2.insert(0, f"{b['x2']:.2f}")
                except Exception:
                    pass
                # mover handles con la barra
                try:
                    if getattr(self, '_handle_artists', None) and len(self._handle_artists)==2:
                        y_h = art.get('y_pos', 0)
                        self._handle_artists[0].set_data([b['x1']], [y_h])
                        self._handle_artists[1].set_data([b['x2']], [y_h])
                except Exception:
                    pass
                self.canvas.draw_idle()
                return
            except Exception:
                pass
        # Fallback Opción B: limpieza rápida solo de ax_barras si no hay artistas
        try:
            self._redraw_barras_quick()
        except Exception:
            self.canvas.draw_idle()

    def _redraw_barras_quick(self):
        """Opción B fallback: clear solo ax_barras y re-dibujar esquema sin recalcular flexión/cortante."""
        if not hasattr(self, 'ax_barras'):
            return
        self.ax_barras.clear()
        self.ax_barras.set_title("Esquema Físico de Barras (Elevación)", fontsize=10, fontweight='bold', color='gray')
        self.ax_barras.set_ylim(-3.5, 3.5)
        self.ax_barras.set_yticks([])
        self.ax_barras.set_ylabel("Despiece", fontweight='bold')
        self.ax_barras.axhline(0, color='black', linestyle='-.', alpha=0.3)
        # re-dibujar columnas y zonas 2H
        for col in getattr(self, 'columnas_graficas', []):
            xs, w = col['x'] - col['w']/2, col['w']
            self.ax_barras.axvspan(xs, xs + w, color='gray', alpha=0.15)
        for v in getattr(self, 'vigas_graficas', []):
            x_face_izq = v['x_ini'] + v['w_izq']/2
            x_face_der = v['x_ini'] + v['L'] - v['w_der']/2
            if self.mostrar_confinamiento.get():
                if v['w_izq'] > 0.01:
                    z_conf_izq = min(x_face_izq + (2.0 * v['H']), (x_face_izq + x_face_der)/2)
                    self.ax_barras.axvspan(x_face_izq, z_conf_izq, color='#ff9933', alpha=0.08, hatch='//', edgecolor='none')
                    self.ax_barras.vlines([x_face_izq, z_conf_izq], -3.5, 3.5, colors='#cc6600', linestyles=':', alpha=0.4)
                if v['w_der'] > 0.01:
                    z_conf_der = max(x_face_der - (2.0 * v['H']), (x_face_izq + x_face_der)/2)
                    self.ax_barras.axvspan(z_conf_der, x_face_der, color='#ff9933', alpha=0.08, hatch='//', edgecolor='none')
                    self.ax_barras.vlines([z_conf_der, x_face_der], -3.5, 3.5, colors='#cc6600', linestyles=':', alpha=0.4)
        # flejes
        for z in getattr(self, 'datos_cortante', []):
            s_m = max(z['s_col_cm'] / 100.0, 0.02)
            n_lines = min(int((z['x2'] - z['x1']) / s_m) + 1, 400)
            for k in range(n_lines):
                xf = z['x1'] + k * s_m
                if xf > z['x2']: break
                self.ax_barras.vlines(xf, -1.6, 1.6, colors='#8e44ad', linewidth=0.6, alpha=0.7)
        # barras (con registro de artistas para drag & handles)
        self._bar_artists = {}
        for b in self.barras:
            c_fact = 1 if b['cara'] == "Superior" else -1
            r_fact = 2 if b['rol'] == "Continuo" else 1
            y_base = c_fact * r_fact
            y_off = c_fact * 0.4 * (b['id'] % 2)
            y_pos = y_base + y_off
            c = 'blue' if b['cara'] == 'Superior' else 'green'
            ls = '-' if b['rol'] == "Continuo" else '--'
            line, = self.ax_barras.plot([b['x1'], b['x2']], [y_pos, y_pos], color=c, linestyle=ls, linewidth=2.5, zorder=4)
            hook_ini = None
            hook_fin = None
            if b.get('g_ini', 0) > 0:
                hook_ini, = self.ax_barras.plot([b['x1'], b['x1']], [y_pos, y_pos - c_fact*0.5], color=c, linewidth=2.5, zorder=4)
            if b.get('g_fin', 0) > 0:
                hook_fin, = self.ax_barras.plot([b['x2'], b['x2']], [y_pos, y_pos - c_fact*0.5], color=c, linewidth=2.5, zorder=4)
            L_real = b['x2'] - b['x1'] + b.get('g_ini', 0) + b.get('g_fin', 0)
            txt = self.ax_barras.text((b['x1']+b['x2'])/2, y_pos + c_fact*0.2, f"{b['cant']}{b['tipo']} (Lc={L_real:.2f}m)", fontsize=7, color=c, ha='center', va='center', fontweight='bold')
            self._bar_artists[b['id']] = {'line': line, 'hook_ini': hook_ini, 'hook_fin': hook_fin, 'text': txt, 'y_pos': y_pos, 'c_fact': c_fact, 'c': c}
        # --- Directriz 1.1: handles ---
        self._handle_artists = getattr(self, '_handle_artists', [])
        for h in list(self._handle_artists):
            try: h.remove()
            except Exception: pass
        self._handle_artists.clear()
        if getattr(self, '_selected_bar_id', None) is not None:
            b_sel = next((x for x in self.barras if x['id'] == self._selected_bar_id), None)
            if b_sel and b_sel['id'] in self._bar_artists:
                y_sel = self._bar_artists[b_sel['id']]['y_pos']
                for xe in [float(b_sel['x1']), float(b_sel['x2'])]:
                    hh, = self.ax_barras.plot([xe], [y_sel], marker='s', markersize=8, markerfacecolor='red', markeredgecolor='black', markeredgewidth=1.2, linestyle='None', zorder=7)
                    self._handle_artists.append(hh)
        # --- Directriz 1.2: cotas traslapo ---
        try:
            for cara in ["Superior", "Inferior"]:
                cont = [b for b in self.barras if b['cara'] == cara and b['rol'] == 'Continuo']
                cont.sort(key=lambda x: float(x['x1']))
                for i in range(len(cont)):
                    for j in range(i+1, len(cont)):
                        b1, b2 = cont[i], cont[j]
                        ox1 = max(float(b1['x1']), float(b2['x1']))
                        ox2 = min(float(b1['x2']), float(b2['x2']))
                        Ltras = ox2 - ox1
                        if Ltras > 0.05:
                            mid = (ox1 + ox2) / 2.0
                            y_cota = 2.9 if cara == "Superior" else -2.9
                            va = 'bottom' if cara == "Superior" else 'top'
                            self.ax_barras.annotate("", xy=(ox1, y_cota), xytext=(ox2, y_cota),
                                arrowprops=dict(arrowstyle='<->', color='#b34700', lw=1.2, shrinkA=2, shrinkB=2))
                            self.ax_barras.text(mid, y_cota + (0.15 if cara=="Superior" else -0.15), f"L_traslapo={Ltras:.2f}m",
                                ha='center', va=va, fontsize=7, color='#b34700', fontweight='bold',
                                bbox=dict(boxstyle="round,pad=0.2", fc="#fff2e6", ec="#b34700", alpha=0.9))
        except Exception:
            pass
        # --- Directriz 2.2: líneas corte A-A ---
        try:
            if getattr(self, 'cortes', None):
                for corte in self.cortes:
                    xm = float(corte.get('x_mid', 0))
                    nombre = str(corte.get('nombre', ''))
                    if 0 <= xm <= self.longitud_eje_real + 1e-6:
                        self.ax_barras.axvline(xm, color='#e67e22', linestyle=':', linewidth=1.2, alpha=0.9, zorder=5)
                        self.ax_barras.text(xm, 3.3, nombre, fontsize=7, color='#e67e22', ha='center', va='bottom', fontweight='bold',
                                            bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="#e67e22", alpha=0.8))
        except Exception:
            pass
        self.canvas.draw_idle()

    def on_barras_release(self, event):
        if self._dragging_bar is None:
            return
        # restaurar grosor
        try:
            # el dict puede haber sido reasignado; usar id almacenado
            for art in self._bar_artists.values():
                try:
                    art['line'].set_linewidth(2.5)
                    art['line'].set_alpha(1.0)
                except Exception:
                    pass
            self.canvas.get_tk_widget().config(cursor="")
        except Exception:
            pass
        self._dragging_bar = None
        self._drag_edge = None
        self.refresh_tree()
        self.actualizar_grafica()
        # actualizar_grafica ya llama a actualizar_verificacion_flexion; se vuelve a llamar por spec + cortante
        self.actualizar_verificacion_cortante()

    # ================= VERIFICACION A FLEXION =================

    def calcular_capacidad_flexion(self, As_traccion, As_compresion, b, h, diam_traccion, diam_compresion, fc, fy, rec):
        """Momento nominal de seccion rectangular doblemente reforzada"""
        if As_traccion == 0: return 0, "Sin acero a tracción", 0
        
        Es = 200000 
        epsilon_cu = 0.003
        epsilon_y = fy / Es 
        
        d_prime = rec + 9.5 + (diam_compresion / 2.0)
        d = h - (rec + 9.5 + (diam_traccion / 2.0))
        
        if 17 <= fc <= 28: beta1 = 0.85
        elif 28 < fc < 55: beta1 = 0.85 - 0.05 * (fc - 28) / 7
        else: beta1 = 0.65
        
        coef_A = 0.85 * fc * beta1 * b
        coef_B = As_compresion * Es * epsilon_cu - As_traccion * fy
        coef_C = -As_compresion * Es * epsilon_cu * d_prime
        discriminante = coef_B**2 - 4 * coef_A * coef_C
        
        if discriminante < 0: return 0, "Error Matemático", 0
        
        c_inicial = (-coef_B + math.sqrt(discriminante)) / (2 * coef_A)
        c, a, Mn = 0, 0, 0
        
        if c_inicial < d_prime:
            epsilon_s_prime = epsilon_cu * (d_prime - c_inicial) / c_inicial
            if epsilon_s_prime >= epsilon_y:
                a = ((As_traccion + As_compresion) * fy) / (0.85 * fc * b)
                c = a / beta1
                Mn = (As_traccion * fy * (d - a/2)) + (As_compresion * fy * (d_prime - a/2))
            else:
                c = c_inicial
                a = beta1 * c
                fs_prime = Es * epsilon_cu * (c - d_prime) / c
                Mn = (As_traccion * fy * (d - a/2)) + (As_compresion * fs_prime * (d_prime - a/2))
        else:
            epsilon_s_prime = epsilon_cu * (c_inicial - d_prime) / c_inicial
            if epsilon_s_prime >= epsilon_y:
                c = (As_traccion * fy - As_compresion * (fy - 0.85 * fc)) / (0.85 * fc * beta1 * b)
                a = beta1 * c
                Mn = (As_traccion * fy * (d - a/2)) + (As_compresion * fy * (a/2 - d_prime))
            else:
                c = c_inicial
                a = beta1 * c
                fs_prime = Es * epsilon_cu * (c - d_prime) / c
                Mn = (As_traccion * fy * (d - a/2)) + (As_compresion * fs_prime * (a/2 - d_prime))
                
        epsilon_t = epsilon_cu * (d - c) / c
        if epsilon_t >= (0.003 + epsilon_y):
            phi = 0.90
            control = "Controlado por Tracción"
        elif epsilon_t > epsilon_y and epsilon_t < (0.003 + epsilon_y):
            phi = 0.65 + 0.25 * (epsilon_t - epsilon_y) / 0.003
            control = "En Transición"
        else:
            phi = 0.65
            control = "Controlado por Compresión"
            
        phi_Mn = (phi * Mn) / 1e6 
        Mn_kNm = Mn / 1e6 
        return phi_Mn, control, Mn_kNm

    def actualizar_verificacion_flexion(self):
            if not hasattr(self, 'envolventes_tramos') or not self.vigas_graficas: 
                return
    
            self.ax_flex.clear()
            self.txt_alertas.delete('1.0', tk.END)
            
            modo = "PURA de ETABS" if not self.aplicar_nsr_envolvente.get() else "con mínimos NSR-10"
            self.ax_flex.set_title(f"φMn Suministrado vs Mu Extraído ({modo})", fontweight='bold')
            
            pts_x = [i * (self.longitud_eje_real / 200.0) for i in range(201)]
            cap_positiva, cap_negativa = [], []
            
            viga_base = self.vigas_graficas[0]
            b_mm = viga_base.get('B', 0.30) * 1000
            h_mm = viga_base['H'] * 1000
            fc_ext = viga_base.get('fc', 21.0)
            fy_ext = viga_base.get('fy', 420.0)
            
            try: rec_mm = float(self.ent_rec.get()) * 10 
            except Exception: rec_mm = 40.0
            
            diams = {"#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1, "#7": 22.2, "#8": 25.4}
            alertas = set()
    
            for x in pts_x:
                As_sup_cont, As_sup_bast = 0, 0
                As_inf_cont, As_inf_bast = 0, 0
                d_sup_max, d_inf_max = 0, 0
                
                for b in self.barras:
                    if b['x1'] <= x <= b['x2']:
                        diam_actual = diams.get(b['tipo'], 15.9)
                        if b['cara'] == "Superior": 
                            if b['rol'] == "Continuo":
                                As_sup_cont = max(As_sup_cont, b['area'])
                            else:
                                As_sup_bast += b['area']
                            d_sup_max = max(d_sup_max, diam_actual)
                        else: 
                            if b['rol'] == "Continuo":
                                As_inf_cont = max(As_inf_cont, b['area'])
                            else:
                                As_inf_bast += b['area']
                            d_inf_max = max(d_inf_max, diam_actual)
                
                As_sup_x = As_sup_cont + As_sup_bast
                As_inf_x = As_inf_cont + As_inf_bast
                
                if d_sup_max == 0: d_sup_max = 15.9
                if d_inf_max == 0: d_inf_max = 15.9

                # --- Fase 3: verificación cuantías DES per PDF p7-9 ---
                # ρmin = max(0.25*sqrt(fc)/fy, 1.4/fy) en MPa (ej 21/420 => 0.0033)
                try:
                    rho_min = max(0.25 * math.sqrt(fc_ext) / fy_ext, 1.4 / fy_ext)
                except Exception:
                    rho_min = 0.0033
                # altura efectiva para cada cara
                d_sup_eff = h_mm - (rec_mm + 9.5 + d_sup_max/2.0)
                d_inf_eff = h_mm - (rec_mm + 9.5 + d_inf_max/2.0)
                if d_sup_eff > 0 and As_sup_x > 0:
                    rho_sup = As_sup_x / (b_mm * d_sup_eff)
                    if rho_sup < rho_min - 1e-9:
                        alertas.add(f"⚠️ X={x:.2f}m (Top): Acero inferior al mínimo normativo ρ={rho_sup:.4f} < ρmin={rho_min:.4f} (As={As_sup_x:.0f}mm²)")
                    if rho_sup > 0.025 + 1e-9:
                        alertas.add(f"⚠️ X={x:.2f}m (Top): Cuantía excede límite DES (0.025) ρ={rho_sup:.4f}")
                if d_inf_eff > 0 and As_inf_x > 0:
                    rho_inf = As_inf_x / (b_mm * d_inf_eff)
                    if rho_inf < rho_min - 1e-9:
                        alertas.add(f"⚠️ X={x:.2f}m (Bot): Acero inferior al mínimo normativo ρ={rho_inf:.4f} < ρmin={rho_min:.4f} (As={As_inf_x:.0f}mm²)")
                    if rho_inf > 0.025 + 1e-9:
                        alertas.add(f"⚠️ X={x:.2f}m (Bot): Cuantía excede límite DES (0.025) ρ={rho_inf:.4f}")
                
                phi_Mn_neg, ctrl_neg, _ = self.calcular_capacidad_flexion(As_sup_x, As_inf_x, b_mm, h_mm, d_sup_max, d_inf_max, fc_ext, fy_ext, rec_mm)
                cap_negativa.append(phi_Mn_neg)
                if ctrl_neg != "Controlado por Tracción" and As_sup_x > 0:
                    alertas.add(f"⚠️ X={x:.2f}m (Top): Sección '{ctrl_neg}' - no dúctil.")
                    
                phi_Mn_pos, ctrl_pos, _ = self.calcular_capacidad_flexion(As_inf_x, As_sup_x, b_mm, h_mm, d_inf_max, d_sup_max, fc_ext, fy_ext, rec_mm)
                cap_positiva.append(phi_Mn_pos)
                if ctrl_pos != "Controlado por Tracción" and As_inf_x > 0:
                    alertas.add(f"⚠️ X={x:.2f}m (Bot): Sección '{ctrl_pos}' - no dúctil.")
    
            # Fix Image1: evitar truncamiento de etiquetas Bot/Top bajo el eje con invert_yaxis
            self.fig_flex.subplots_adjust(bottom=0.22, top=0.92)
            
            if self.envolventes_tramos:
                self.ax_flex.text(-0.02, -0.06, "(-) Moment", transform=self.ax_flex.transAxes, ha='right', va='top', fontsize=8.5, color='darkred', fontweight='bold', clip_on=False)
                self.ax_flex.text(-0.02, -0.11, "(+) Moment", transform=self.ax_flex.transAxes, ha='right', va='top', fontsize=8.5, color='darkblue', fontweight='bold', clip_on=False)
                
                for idx, tramo in enumerate(self.envolventes_tramos):
                    lbl_neg = 'Mu Demandado (-) ETABS' if idx == 0 else ""
                    lbl_pos = 'Mu Demandado (+) ETABS' if idx == 0 else ""
                    
                    self.ax_flex.plot(tramo['x'], tramo['m_neg'], color='red', linewidth=1.0, label=lbl_neg)
                    self.ax_flex.fill_between(tramo['x'], tramo['m_neg'], 0, facecolor='none', edgecolor='red', hatch='///', alpha=0.5)
                    
                    self.ax_flex.plot(tramo['x'], tramo['m_pos'], color='blue', linewidth=1.0, label=lbl_pos)
                    self.ax_flex.fill_between(tramo['x'], tramo['m_pos'], 0, facecolor='none', edgecolor='blue', hatch='///', alpha=0.5)
                    
                    x_start, x_end = tramo['x'][0], tramo['x'][-1]
                    L_span = x_end - x_start
                    if L_span <= 0: continue
                    x_L = x_start + (L_span * 0.15)
                    x_M = x_start + (L_span * 0.50)
                    x_R = x_start + (L_span * 0.85)
    
                    self.ax_flex.text(x_L, -0.06, f"{tramo['text_top'][0]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkred', clip_on=False)
                    self.ax_flex.text(x_L, -0.11, f"{tramo['text_bot'][0]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkblue', clip_on=False)
                    self.ax_flex.text(x_M, -0.06, f"{tramo['text_top'][1]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkred', clip_on=False)
                    self.ax_flex.text(x_M, -0.11, f"{tramo['text_bot'][1]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkblue', clip_on=False)
                    self.ax_flex.text(x_R, -0.06, f"{tramo['text_top'][2]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkred', clip_on=False)
                    self.ax_flex.text(x_R, -0.11, f"{tramo['text_bot'][2]:.2f}", transform=self.ax_flex.get_xaxis_transform(), ha='center', va='top', fontsize=7.5, color='darkblue', clip_on=False)
            else:
                self.txt_alertas.insert(tk.END, "⚠️ No se detectaron combos de diseño.\n")
    
            self.ax_flex.step(pts_x, [-m for m in cap_negativa], where="mid", color='darkred', linewidth=2.5, label='φMn Suministrado (-)')
            self.ax_flex.step(pts_x, cap_positiva, where="mid", color='darkblue', linewidth=2.5, label='φMn Suministrado (+)')
            
            self.ax_flex.axhline(0, color='black', linewidth=1)
            self.ax_flex.set_ylabel("Momento Flector (kN.m)")
            self.ax_flex.set_xlabel("Longitud del Eje Constructivo (m)")
            
            self.ax_flex.invert_yaxis() 
            self.ax_flex.legend(loc='upper right', fontsize=8)
            self.ax_flex.grid(True, linestyle=':', alpha=0.6)
            
            self.canvas_flex.draw()
            
            if not alertas and self.envolventes_tramos: 
                self.txt_alertas.insert(tk.END, "✅ DISEÑO OK: El acero suministrado soporta la envolvente requerida y cumple ductilidad.")
            else:
                for alerta in sorted(list(alertas))[:8]: 
                    self.txt_alertas.insert(tk.END, alerta + "\n")

    # ================= CORTANTE ESTILO NSR-10 / DC-CAD =================

    def exportar_json(self):
        if not self.estaciones: return
        self._guardar_estado_en_registro()
        entry = {
            "eje": self.viga_label, 
            "L_total": self.longitud_eje_real, 
            "columnas": self.columnas_graficas,
            "vigas": self.vigas_graficas,
            "barras": self.barras,
            "zonas_cortante": [{k: z[k] for k in ('nombre','x1','x2','ramas','tipo_est','s_col_cm','estado')} for z in self.datos_cortante],
            "cortes": self.cortes
        }
        
        if not self.ruta_json_acumulado:
            ruta = filedialog.asksaveasfilename(defaultextension=".json", initialfile="Despiece_Acumulado.json", filetypes=[("JSON", "*.json")])
            if not ruta: return
            self.ruta_json_acumulado = ruta
        
        try:
            data = {"ejes": []}
            try:
                with open(self.ruta_json_acumulado, 'r', encoding='utf-8') as f:
                    cargado = json.load(f)
                data = cargado if isinstance(cargado, dict) and 'ejes' in cargado else {"ejes": [cargado]}
            except FileNotFoundError:
                pass
            
            data["ejes"] = [e for e in data["ejes"] if e.get("eje") != self.viga_label]
            data["ejes"].append(entry)
            
            with open(self.ruta_json_acumulado, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            
            self.lbl_ruta_json.config(text=f"Archivo: {self.ruta_json_acumulado} ({len(data['ejes'])} ejes)")
            messagebox.showinfo("Éxito", f"'{self.viga_label}' guardado en el JSON acumulado.\nTotal de ejes en el archivo: {len(data['ejes'])}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def cambiar_archivo_json(self):
        ruta = filedialog.askopenfilename(filetypes=[("JSON", "*.json")], title="Selecciona el JSON acumulado existente (o Cancela para elegir uno nuevo al guardar)")
        if ruta:
            self.ruta_json_acumulado = ruta
            try:
                with open(ruta, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                n = len(data.get('ejes', [])) if isinstance(data, dict) else 1
                self.lbl_ruta_json.config(text=f"Archivo: {ruta} ({n} ejes)")
            except Exception:
                self.lbl_ruta_json.config(text=f"Archivo: {ruta}")

    def obtener_as_en_x(self, x):
        """Obtiene el área de acero real colocada (tracción y compresión) en una coordenada x"""
        As_sup_cont, As_sup_bast, As_inf_cont, As_inf_bast = 0, 0, 0, 0
        d_sup_max, d_inf_max = 0, 0
        
        for b in self.barras:
            if b['x1'] <= x <= b['x2']:
                diam = self.diametros_db.get(b['tipo'], 15.9)
                if b['cara'] == "Superior": 
                    if b['rol'] == "Continuo": As_sup_cont = max(As_sup_cont, b['area'])
                    else: As_sup_bast += b['area']
                    d_sup_max = max(d_sup_max, diam)
                else: 
                    if b['rol'] == "Continuo": As_inf_cont = max(As_inf_cont, b['area'])
                    else: As_inf_bast += b['area']
                    d_inf_max = max(d_inf_max, diam)
                    
        return (As_sup_cont + As_sup_bast), (As_inf_cont + As_inf_bast), (d_sup_max or 15.9), (d_inf_max or 15.9)

    def calcular_shear_zone(self, zona_info):
        b_w_mm = zona_info['b_w']
        h_mm = zona_info['h']
        d_mm = h_mm - 60
        fc = zona_info['fc']
        fy = zona_info['fy']
        
        Ve_diseno = zona_info['Vu_analisis'] 
        Vp = zona_info.get('Vp', 0.0)
        Pu = zona_info['Pu']
        Ag = b_w_mm * h_mm
        
        is_conf = "Conf" in zona_info['nombre']
        
        if is_conf:
            phi = 0.75  
            s_norma_mm = min(d_mm / 4.0, 150.0)  
        else:
            phi = 0.75
            s_norma_mm = min(d_mm / 2.0, 300.0)  
            
        Vc_calc = 0.17 * math.sqrt(fc) * b_w_mm * d_mm / 1000.0  
        limite_axial = (Ag * fc / 20.0) / 1000.0  
        
        # Anulación Vc: Vc=0 si Vp≥0.5·Ve_diseno y Pu≤Ag·fc'/20 en zona conf.
        cond_vc_cero = is_conf and Vp >= (0.5 * Ve_diseno) and Pu <= limite_axial and zona_info.get('sin_vc_conf', True)
        if cond_vc_cero:
            Vc = 0.0
            phi = 0.60  
        else:
            Vc = Vc_calc
            
        phi_Vc = phi * Vc
        Vs_req = max(0.0, (Ve_diseno - phi_Vc) / phi)
        
        ramas = zona_info['ramas']
        diam_fleje = self.diametros_db.get(zona_info['tipo_est'], 9.5)
        Av = ramas * (math.pi * diam_fleje**2 / 4.0)
        
        if Vs_req > 0:
            s_calc_mm = (Av * fy * d_mm) / (Vs_req * 1000.0)
            s_req_mm = min(s_calc_mm, s_norma_mm)
        else:
            s_req_mm = s_norma_mm
            
        s_req_cm = s_req_mm / 10.0
        s_col = zona_info['s_col_cm'] * 10.0
        Vs_prov = (Av * fy * d_mm) / (s_col * 1000.0) if s_col > 0 else 0.0
            
        phi_Vn = phi * (Vc + Vs_prov)
        estado = "OK" if phi_Vn >= Ve_diseno else "FALLA"
        
        zona_info.update({
            'V_diseno': Ve_diseno, 'phi': phi, 'Vc': Vc, 
            'Vs_req': Vs_req, 's_req_cm': s_req_cm,
            'phi_Vn': phi_Vn, 'estado': estado
        })
        return zona_info

    def autodisenar_flejes(self):
        """Escala automaticamente: primero sube ramas, luego calibre, hasta lograr s >= s_min."""
        if not self.datos_cortante:
            return messagebox.showwarning("Aviso", "Primero extrae un eje y analiza el cortante.")
        try:
            s_min_cm = float(self.ent_smin.get())
        except ValueError:
            return messagebox.showwarning("Error", "S.mín inválido.")
        
        sin_solucion = []
        for z in self.datos_cortante:
            mejor = None
            for tipo in self.flejes_db:
                for ramas in [2, 4, 6, 8]:
                    prueba = dict(z)
                    prueba['ramas'] = ramas
                    prueba['tipo_est'] = tipo
                    resultado = self.calcular_shear_zone(prueba)
                    if resultado['s_req_cm'] >= s_min_cm - 1e-9:
                        mejor = (tipo, ramas, resultado['s_req_cm'])
                        break
                if mejor: break
            
            if mejor:
                z['tipo_est'], z['ramas'] = mejor[0], mejor[1]
                s_col = math.floor(mejor[2] * 2) / 2.0
                z['s_col_cm'] = max(s_col, s_min_cm)
            else:
                z['tipo_est'], z['ramas'] = "#6", 8
                z['s_col_cm'] = s_min_cm
                sin_solucion.append(z['nombre'])
            self.calcular_shear_zone(z)
        
        self.refresh_tree_cortante()
        self.plotear_cortante()
        
        if sin_solucion:
            messagebox.showwarning("Auto-diseño flejes", f"Zonas donde ni #6 con 8 ramas alcanza con s≥{self.ent_smin.get()}cm:\n" + "\n".join(sin_solucion))
        else:
            messagebox.showinfo("Auto-diseño flejes", "Todas las zonas diseñadas automáticamente.")

    def actualizar_verificacion_cortante(self):
        if not self.vigas_graficas: return
        
        try:
            rec_mm = float(self.ent_rec.get()) * 10
        except ValueError:
            return messagebox.showwarning("Error", "Recubrimiento inválido.")
            
        if not hasattr(self, 'datos_cortante'): self.datos_cortante = []
        self.datos_cortante.clear()
        
        # 1. Detectar TRASLAPOS en todo el modelo (Solapamiento de barras de la misma cara)
        overlaps = []
        for i, b1 in enumerate(self.barras):
            for j, b2 in enumerate(self.barras):
                if i < j and b1['cara'] == b2['cara']:
                    ox1 = max(b1['x1'], b2['x1'])
                    ox2 = min(b1['x2'], b2['x2'])
                    if ox1 < ox2:
                        overlaps.append((ox1, ox2))
        
        fuente = self.fuente_cortante.get() if hasattr(self, 'fuente_cortante') else 'Vu'
        sin_vc_conf = self.chk_sin_vc_conf.get() if hasattr(self, 'chk_sin_vc_conf') else True
        
        for idx, v in enumerate(self.vigas_graficas):
            b_mm = v.get('B', 0.3) * 1000
            h_mm = v.get('H', 0.45) * 1000
            fc = v.get('fc', 21.0)
            fy = v.get('fy', 420.0)
            Ln = v['L'] - (v['w_izq']/2 + v['w_der']/2)
            if Ln <= 0: continue
            
            x_face_izq = v['x_ini'] + v['w_izq']/2
            x_face_der = v['x_ini'] + v['L'] - v['w_der']/2
            
            # Obtener áreas de acero real en las caras de los apoyos para Mpr
            As_sup_izq, As_inf_izq, dt_sup_izq, dt_inf_izq = self.obtener_as_en_x(x_face_izq + 0.01)
            As_sup_der, As_inf_der, dt_sup_der, dt_inf_der = self.obtener_as_en_x(x_face_der - 0.01)
            
            # Momentos Probables (Sismo) usando fy * 1.25
            fy_pr = fy * 1.25
            _, _, Mpr1_neg = self.calcular_capacidad_flexion(As_sup_izq, As_inf_izq, b_mm, h_mm, dt_sup_izq, dt_inf_izq, fc, fy_pr, rec_mm)
            _, _, Mpr1_pos = self.calcular_capacidad_flexion(As_inf_izq, As_sup_izq, b_mm, h_mm, dt_inf_izq, dt_sup_izq, fc, fy_pr, rec_mm)
            _, _, Mpr2_neg = self.calcular_capacidad_flexion(As_sup_der, As_inf_der, b_mm, h_mm, dt_sup_der, dt_inf_der, fc, fy_pr, rec_mm)
            _, _, Mpr2_pos = self.calcular_capacidad_flexion(As_inf_der, As_sup_der, b_mm, h_mm, dt_inf_der, dt_sup_der, fc, fy_pr, rec_mm)
            
            # Perfil lineal de V_diseño(x) = max(Ve(x), Vu(x))
            tramo_x = []
            v_diseno_x = []
            Pu_tramo = 0.0
            Vp_max = 0.0
            
            if hasattr(self, 'envolventes_tramos') and len(self.envolventes_tramos) > idx:
                tramo = self.envolventes_tramos[idx]
                Pu_tramo = tramo.get('pu', 0.0)
                tramo_x = tramo['x']
                
                # Cortante Sísmico Puro (Vp)
                Vp1 = (Mpr1_neg + Mpr2_pos) / Ln
                Vp2 = (Mpr2_neg + Mpr1_pos) / Ln
                Vp_max = max(abs(Vp1), abs(Vp2))
                
                # Perfil lineal de V_diseño(x) = max(Ve(x), Vu(x))
                for i, px in enumerate(tramo_x):
                    vg = tramo['v_grav'][i]
                    vu_max_env = max(abs(tramo['v_pos'][i]), abs(tramo['v_neg'][i]))
                    
                    # Cortante por capacidad (Sismo hiperestático + Gravedad)
                    ve_1 = abs(vg + Vp1)
                    ve_2 = abs(vg - Vp2)
                    ve_max_x = max(ve_1, ve_2)
                    
                    v_diseno_x.append(max(ve_max_x, vu_max_env))
                
                self.envolventes_tramos[idx]['v_diseno_x'] = v_diseno_x
            
            # --- ZONIFICACION DINAMICA (2H y Traslapos) ---
            L_conf = 2.0 * (h_mm / 1000.0)
            pts = [x_face_izq, x_face_der, x_face_izq + L_conf, x_face_der - L_conf]
            for ox1, ox2 in overlaps:
                pts.extend([ox1, ox2])
            
            # Filtrar puntos estrictamente dentro de la luz libre y ordenar
            pts = sorted(list(set([round(p, 4) for p in pts if x_face_izq - 1e-3 <= p <= x_face_der + 1e-3])))
            
            zonas_tramo = []
            for p1, p2 in zip(pts[:-1], pts[1:]):
                if p2 - p1 < 0.01: continue 
                pm = (p1 + p2) / 2.0 
                
                if pm <= x_face_izq + L_conf: name = "Conf_Izq"
                elif pm >= x_face_der - L_conf: name = "Conf_Der"
                else:
                    name = "Centro"
                    for ox1, ox2 in overlaps:
                        if ox1 <= pm <= ox2:
                            name = "Conf_Traslapo"
                            break
                            
                if zonas_tramo and zonas_tramo[-1]['name'] == name:
                    zonas_tramo[-1]['x2'] = p2
                else:
                    zonas_tramo.append({'name': name, 'x1': p1, 'x2': p2})
            
# 2. EXTRACCIÓN EXACTA DE Vu Y CÁLCULO POR ZONAS
            for z_info in zonas_tramo:
                z_name = z_info['name']
                z_x1, z_x2 = z_info['x1'], z_info['x2']
                
                V_diseno_zona = 0
                if tramo_x and v_diseno_x:
                    for px, vdx in zip(tramo_x, v_diseno_x):
                        if z_x1 - 1e-3 <= px <= z_x2 + 1e-3:
                            V_diseno_zona = max(V_diseno_zona, vdx)
                
                z_dict = {
                    'viga_idx': idx, 
                    'nombre': f"V{idx+1} - {z_name}",
                    'x1': z_x1, 'x2': z_x2,
                    'b_w': b_mm, 'h': h_mm, 'fc': fc, 'fy': fy,
                    'Vu_analisis': V_diseno_zona, 
                    'Vp': Vp_max, 
                    'Pu': Pu_tramo, 'ramas': 2, 'tipo_est': "#3", 
                    's_col_cm': 10.0 if "Conf" in z_name else 20.0,
                    'fuente': fuente, 'sin_vc_conf': sin_vc_conf
                }
                self.datos_cortante.append(self.calcular_shear_zone(z_dict))
                
        self.refresh_tree_cortante()
        self.plotear_cortante()

    def refresh_tree_cortante(self):
        for item in self.tree_cort.get_children(): self.tree_cort.delete(item)
        for i, z in enumerate(self.datos_cortante):
            tag = 'falla' if z['estado'] == "FALLA" else 'ok'
            self.tree_cort.insert("", tk.END, iid=i, values=(
                f"Viga {z['viga_idx']+1}", z['nombre'].split(" - ")[1],
                f"{z['x1']:.2f}", f"{z['x2']:.2f}",
                f"{z['V_diseno']:.1f}", f"{z['phi']:.2f}", f"{z['Vc']:.1f}",
                f"{z['Vs_req']:.1f}", z['ramas'], z['tipo_est'], 
                f"{z['s_req_cm']:.1f}", f"{z['s_col_cm']:.1f}", z['estado']
            ), tags=(tag,))
        self.tree_cort.tag_configure('falla', background='#ffcccc')
        self.tree_cort.tag_configure('ok', background='#e6ffe6')

    def on_tree_cortante_select(self, event):
        sel = self.tree_cort.selection()
        if not sel: return
        idx = int(sel[0])
        z = self.datos_cortante[idx]
        
        self.spn_c_ramas.delete(0, tk.END); self.spn_c_ramas.insert(0, str(z['ramas']))
        self.cmb_c_tipo.set(z['tipo_est'])
        self.ent_c_s.delete(0, tk.END); self.ent_c_s.insert(0, str(z['s_col_cm']))

    # Fase 2: handler pick_event para axvspan de cortante (reutiliza Treeview)
    def on_cortante_pick(self, event):
        # event.artist es el PolyCollection del axvspan; gid almacena índice en self.datos_cortante
        artist = getattr(event, 'artist', None)
        if artist is None:
            return
        idx = None
        try:
            gid = artist.get_gid()
            if gid is not None and str(gid).isdigit():
                idx = int(gid)
        except Exception:
            pass
        if idx is None:
            # fallback: buscar por identidad en _cort_artists
            try:
                if hasattr(self, '_cort_artists') and artist in self._cort_artists:
                    idx = self._cort_artists.index(artist)
            except Exception:
                return
        if idx is None or not (0 <= idx < len(self.datos_cortante)):
            return
        # Seleccionar en Treeview y cargar en panel f_edit_c (usa lógica existente)
        iid = str(idx)
        try:
            if iid in self.tree_cort.get_children():
                self.tree_cort.selection_set(iid)
                self.tree_cort.focus(iid)
                self.tree_cort.see(iid)
        except Exception:
            pass
        try:
            z = self.datos_cortante[idx]
            self.spn_c_ramas.delete(0, tk.END); self.spn_c_ramas.insert(0, str(z['ramas']))
            self.cmb_c_tipo.set(z['tipo_est'])
            self.ent_c_s.delete(0, tk.END); self.ent_c_s.insert(0, str(z['s_col_cm']))
            # highlight sutil en gráfica sin recalcular
            self.canvas_cortante.draw_idle()
        except Exception:
            pass

    def actualizar_zona_cortante(self):
        sel = self.tree_cort.selection()
        if not sel: return
        idx = int(sel[0])
        
        try:
            ramas = int(self.spn_c_ramas.get())
            tipo = self.cmb_c_tipo.get()
            s_col = float(self.ent_c_s.get())
            
            self.datos_cortante[idx]['ramas'] = ramas
            self.datos_cortante[idx]['tipo_est'] = tipo
            self.datos_cortante[idx]['s_col_cm'] = s_col
            
            self.datos_cortante[idx] = self.calcular_shear_zone(self.datos_cortante[idx])
            self.refresh_tree_cortante()
            self.plotear_cortante()
        except ValueError:
            messagebox.showwarning("Error", "Datos de edición inválidos")

    def plotear_cortante(self):
        self.ax_v2.clear()
        # Fase 2: reset artistas pickeables
        if hasattr(self, '_cort_artists'):
            self._cort_artists.clear()
        else:
            self._cort_artists = []
        self.ax_v2.set_title("Envolvente V2 de ETABS vs φVn Provisto", fontweight='bold')
        
        for idx, z in enumerate(self.datos_cortante):
            if "Conf" in z['nombre']:
                color = '#ffe6cc' if "Traslapo" in z['nombre'] else '#ffd9b3' 
                poly = self.ax_v2.axvspan(z['x1'], z['x2'], facecolor=color, alpha=0.5, edgecolor='none', picker=True)
                try:
                    poly.set_gid(str(idx))
                    poly.set_picker(True)
                except Exception:
                    pass
                self._cort_artists.append(poly)

        for col in self.columnas_graficas:
            w = col['w'] if col['w'] > 0.05 else 0.30
            xs = col['x'] - w/2
            self.ax_v2.axvspan(xs, xs + w, facecolor='#666666', alpha=0.5, hatch='//')
                
        pts_x = [i * (self.longitud_eje_real / 500.0) for i in range(501)]
        cap_vn = []
        
        for x in pts_x:
            v_cap = 0
            for z in self.datos_cortante:
                if z['x1'] <= x <= z['x2']:
                    v_cap = z['phi_Vn']
                    break
            cap_vn.append(v_cap)
            
        if hasattr(self, 'envolventes_tramos'):
            for tramo in self.envolventes_tramos:
                if 'v_diseno_x' in tramo:
                    v_env = tramo['v_diseno_x']
                    self.ax_v2.plot(tramo['x'], v_env, color='purple', linewidth=1.5, alpha=0.9, label='V_diseño (Ve DES)' if tramo is self.envolventes_tramos[0] else "")
                    self.ax_v2.fill_between(tramo['x'], v_env, 0, color='purple', alpha=0.15)
                
        self.ax_v2.step(pts_x, cap_vn, where="mid", color='green', linewidth=2.5, label='φVn Provisto (+)')
        self.ax_v2.step(pts_x, [-c for c in cap_vn], where="mid", color='green', linewidth=2.5, label='φVn Provisto (-)')
        
        self.ax_v2.axhline(0, color='black', linewidth=1.5)
        self.ax_v2.set_ylabel("Fuerza Cortante (kN)")
        self.ax_v2.set_xlabel("Coordenada del Eje Constructivo (m)")
        self.ax_v2.legend(loc="upper right", fontsize=8)
        self.ax_v2.grid(True, linestyle=':', alpha=0.7)
        
        if cap_vn:
            self.ax_v2.set_ylim(-max(cap_vn)*1.2, max(cap_vn)*1.2)
            
        self.fig_cortante.tight_layout()
        self.canvas_cortante.draw()

    # ================= CORTES DE SECCION =================

    def generar_cortes(self):
        """Genera un corte en cada punto donde cambia el refuerzo efectivo (ignorando solapes de continuos traslapados) o el fleje. Asigna nomenclatura A-A, B-B..."""
        if not self.barras or not self.vigas_graficas:
            return messagebox.showwarning("Aviso", "Genera primero el despiece (modulación automática o manual).")
        try:
            rec_mm = float(self.ent_rec.get()) * 10
        except ValueError:
            rec_mm = 40.0
        
        bordes = set([0.0, round(self.longitud_eje_real, 4)])
        for b in self.barras:
            bordes.add(round(float(b['x1']), 4)); bordes.add(round(float(b['x2']), 4))
        for z in getattr(self, 'datos_cortante', []) or []:
            try: bordes.add(round(float(z['x1']), 4)); bordes.add(round(float(z['x2']), 4))
            except Exception: pass
        bordes = sorted(bordes)
        
        def firma_en(x):
            # --- Directriz 2.1: capacidad efectiva (ignorar duplicado por traslapo continuo) ---
            # Para continuos traslapados mismo tipo/cara, solo contar max (no suma)
            def firma_cara(cara):
                cont = [b for b in self.barras if b['cara'] == cara and b['rol'] == "Continuo" and float(b['x1']) <= x <= float(b['x2'])]
                bast = [b for b in self.barras if b['cara'] == cara and b['rol'] == "Bastón" and float(b['x1']) <= x <= float(b['x2'])]
                # Continuos: agrupar por tipo y tomar max cant (traslapo = no duplica)
                cont_eff = []
                if cont:
                    grp = {}
                    for b in cont:
                        try: c = int(b['cant']); t = str(b['tipo'])
                        except Exception: continue
                        grp[t] = max(grp.get(t, 0), c)
                    cont_eff = [("Continuo", c, t) for t, c in grp.items()]
                bast_tup = [(b['rol'], int(b['cant']), str(b['tipo'])) for b in bast]
                return tuple(sorted(cont_eff + bast_tup))
            sup = firma_cara("Superior")
            inf = firma_cara("Inferior")
            fleje = None
            for z in getattr(self, 'datos_cortante', []) or []:
                try:
                    if float(z['x1']) <= x <= float(z['x2']):
                        fleje = (str(z['tipo_est']), int(z['ramas']), float(z['s_col_cm']))
                        break
                except Exception: continue
            return sup, inf, fleje
        
        segmentos = []
        for x1, x2 in zip(bordes[:-1], bordes[1:]):
            if x2 - x1 < 0.05: continue
            f = firma_en((x1 + x2) / 2.0)
            if segmentos and segmentos[-1][2] == f:
                segmentos[-1][1] = x2
            else:
                segmentos.append([x1, x2, f])
        
        # --- Directriz 2.2: nomenclatura A-A, B-B ---
        def idx_a_letras(idx):
            letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            if idx < 26:
                return letras[idx]
            # para >26, genera AA, AB, etc. (Excel-style)
            base = ""
            n = idx
            while True:
                base = letras[n % 26] + base
                n = n // 26 - 1
                if n < 0: break
            return base
        self.cortes = []
        v_ref = self.vigas_graficas[len(self.vigas_graficas)//2] if self.vigas_graficas else {'B':0.30,'H':0.45}
        for idx, (x1, x2, (sup, inf, fleje)) in enumerate(segmentos):
            xm = round((x1 + x2) / 2.0, 3)
            letra = idx_a_letras(idx)
            nombre = f"{letra}-{letra}"
            
            def agrupar(firma):
                grupos = {}
                for rol, cant, tipo in firma:
                    grupos[tipo] = grupos.get(tipo, 0) + int(cant)
                return [{"cant": c, "tipo": t} for t, c in sorted(grupos.items(), key=lambda kv: -kv[1])]
            
            estribo = {"tipo": fleje[0], "ramas": fleje[1], "s_cm": fleje[2]} if fleje else {"tipo": "#3", "ramas": 2, "s_cm": 20.0}
            
            self.cortes.append({
                'x1': x1, 'x2': x2, 'x_mid': xm,
                'b_mm': float(v_ref.get('B', 0.30) * 1000) if isinstance(v_ref.get('B',0.30), (int,float)) else 300,
                'h_mm': float(v_ref.get('H', 0.45) * 1000) if isinstance(v_ref.get('H',0.45), (int,float)) else 450,
                'rec_mm': rec_mm,
                'sup': agrupar(sup),
                'inf': agrupar(inf),
                'estribo': estribo,
                'nombre': nombre
            })
        
        self.dibujar_cortes_preview()
        try:
            self.lbl_info_cortes.config(text=f"{len(self.cortes)} cortes generados (incluidos al guardar JSON)")
        except Exception: pass
        # refrescar elevación para mostrar líneas A-A (Directriz 2.2)
        try: self.actualizar_grafica()
        except Exception: pass

    def dibujar_cortes_preview(self):
        self.fig_cortes.clear()
        if not self.cortes:
            self.fig_cortes.text(0.5, 0.5, "Sin cortes. Genera despiece y pulsa 'Generar Cortes'.", ha='center')
            self.canvas_cortes.draw()
            return
        
        n = len(self.cortes)
        ncols = min(n, 4)
        nrows = math.ceil(n / ncols)
        axs = self.fig_cortes.subplots(nrows, ncols, squeeze=False, gridspec_kw={'wspace': 0.35, 'hspace': 0.55})
        
        for i, corte in enumerate(self.cortes):
            ax = axs[i // ncols][i % ncols]
            self._dibujar_un_corte(ax, corte)
        for j in range(n, nrows * ncols):
            axs[j // ncols][j % ncols].axis('off')
        
        self.fig_cortes.tight_layout()
        self.canvas_cortes.draw()

    def _dibujar_un_corte(self, ax, corte):
        b = float(corte.get('b_mm', 300)); h = float(corte.get('h_mm', 400)); rec = float(corte.get('rec_mm', 40))
        df = self.diametros_db.get(corte.get('estribo', {}).get('tipo', '#3'), 9.5)
        nombre = str(corte.get('nombre', ''))
        
        ax.add_patch(plt.Rectangle((-b/2, -h), b, h, fill=False, edgecolor='black', linewidth=1.5))
        inset = rec + df/2
        ax.add_patch(plt.Rectangle((-b/2 + inset, -h + inset), b - 2*inset, h - 2*inset, fill=False, edgecolor='#8e44ad', linewidth=1.2))
        if int(corte.get('estribo', {}).get('ramas', 2)) >= 4:
            for xf in [-b/6, b/6]:
                ax.plot([xf, xf], [-h + inset, -inset], color='#8e44ac', linewidth=1.0)
        
        def colocar(grupos, es_sup):
            total = sum(int(g['cant']) for g in grupos)
            if total == 0: return
            db_max = max(self.diametros_db.get(g['tipo'], 15.9) for g in grupos)
            margen_x = inset + db_max/2
            xs = [-b/2 + margen_x + i * (b - 2*margen_x) / max(total-1, 1) for i in range(total)]
            y = -(rec + df + db_max/2) if es_sup else -(h - rec - df - db_max/2)
            color = '#1560bd' if es_sup else '#1e8449'
            k = 0
            for g in grupos:
                db = self.diametros_db.get(g['tipo'], 15.9)
                for _ in range(int(g['cant'])):
                    ax.add_patch(plt.Circle((xs[k], y), db/2, color=color, clip_on=False))
                    # Directriz 4.3: calibre junto a cada barra si hay mezcla; si no, cada 1 para claridad
                    try:
                        # offset vertical opuesto a concreto
                        off = 6 if es_sup else -8
                        ax.text(xs[k], y + off, str(g['tipo']), fontsize=5, ha='center', va='center', color=color, fontweight='bold',
                                bbox=dict(boxstyle="circle,pad=0.1", fc="white", ec=color, alpha=0.7, lw=0.5))
                    except Exception: pass
                    k += 1
        
        colocar(corte.get('sup', []), True)
        colocar(corte.get('inf', []), False)
        
        etiqueta = f"{nombre}  x≈{float(corte.get('x_mid',0)):.2f}m\nSup: " + "+".join(f"{g['cant']}{g['tipo']}" for g in corte.get('sup', [])) + "\n"
        etiqueta += "Inf: " + "+".join(f"{g['cant']}{g['tipo']}" for g in corte.get('inf', [])) + "\n"
        etiqueta += f"Fleje: {corte.get('estribo',{}).get('tipo','#3')}@{float(corte.get('estribo',{}).get('s_cm',20)):.0f} ({corte.get('estribo',{}).get('ramas',2)}r)"
        ax.set_title(etiqueta, fontsize=7, fontweight='bold' if nombre else 'normal')
        ax.set_aspect('equal')
        ax.autoscale()
        ax.axis('off')
        # gancho 135° preview (esquina superior derecha)
        try:
            inset2 = rec + df/2
            hook_len = df * 1.8
            # esquina sup derecha interior
            x0 = b/2 - inset2; y0 = -inset2
            ang = math.radians(135)
            x1 = x0 + hook_len * math.cos(ang)
            y1 = y0 + hook_len * math.sin(ang)
            ax.plot([x0, x1], [y0, y1], color='#8e44ad', linewidth=1.2)
        except Exception: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = DetalladorVigasApp(root)
    root.mainloop()