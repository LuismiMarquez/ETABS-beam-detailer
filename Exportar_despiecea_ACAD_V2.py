import json
import math
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from pyautocad import Autocad, APoint

ESCALA_Y = 4.5      # exageracion visual del peralte en la elevacion
GAP_EJES = 10.0     # separacion vertical (m de dibujo) entre ejes apilados
GAP_CORTES = 4.0    # espacio entre la elevacion y la fila de cortes

# --- Helpers seguros Directriz 3 ---
def _f(v, default=0.0):
    try: return float(v)
    except Exception: return float(default)
def _i(v, default=0):
    try: return int(float(v))
    except Exception: return int(default)
def _s(v, default=""):
    try: return str(v).strip() if v is not None else default
    except Exception: return default
def P(x, y, z=0.0):
    """APoint 3D seguro para pyautocad (Directriz 3.1)"""
    return APoint(_f(x), _f(y), _f(z))

def conectar_autocad():
    """Se conecta al dibujo ACTIVO. Nunca crea ni abre archivos nuevos:
    todos los despieces se acumulan en un solo dibujo."""
    try:
        comtypes_check = None
        try:
            import comtypes.client
            comtypes.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            return None
        acad = Autocad(create_if_not_exists=False)
        acad.prompt("Iniciando dibujado del despiece acumulado...\n")
        return acad
    except Exception:
        return None

def set_color(obj, color_cad):
    try: obj.Color = color_cad
    except Exception: pass

def crear_cota(acad, p1, p2, p_texto):
    cota = acad.model.AddDimAligned(p1, p2, p_texto)
    def safe_set(propiedad, valor):
        try: setattr(cota, propiedad, valor)
        except Exception: pass
    safe_set('TextHeight', 0.12)
    safe_set('ArrowheadSize', 0.10)
    safe_set('ExtensionLineOffset', 0.05)
    safe_set('ExtensionLineExtend', 0.05)
    safe_set('TextGap', 0.03)
    safe_set('TextColor', 1)
    safe_set('DimensionLineColor', 1)
    safe_set('ExtensionLineColor', 1)
    # 4.2: ticks oblicuos estilo arquitectónico
    safe_set('Arrowhead1Type', 5)
    safe_set('Arrowhead2Type', 5)
    safe_set('Arrowhead1Block', "_OBLIQUE")
    safe_set('Arrowhead2Block', "_OBLIQUE")
    safe_set('VerticalTextPosition', 1)
    safe_set('DimLine1Suppress', False)
    safe_set('DimLine2Suppress', False)
    return cota

def dibujar_linea(acad, p1, p2, color_cad=256):
    linea = acad.model.AddLine(p1, p2)
    set_color(linea, color_cad)
    return linea

def agregar_texto(acad, texto, punto, altura=0.12, color_cad=7):
    t = acad.model.AddText(str(texto), punto, altura)
    set_color(t, color_cad)
    return t

# ================= ELEVACION =================

def dibujar_elevacion(acad, eje, y0):
    L_total = _f(eje.get("L_total", 0.0))
    columnas = eje.get("columnas", []) or []
    vigas = eje.get("vigas", []) or []
    barras = eje.get("barras", []) or []
    cortes = eje.get("cortes", []) or []

    # H viga con parseo estricto
    Hs = []
    for v in vigas:
        try: Hs.append(_f(v.get('H', 0.40)))
        except Exception: Hs.append(0.40)
    H_viga_real = max(Hs) if Hs else 0.40
    prop_viga_nombre = _s(vigas[0].get('prop', 'VIGA')) if vigas else 'VIGA'
    H_visual = H_viga_real * ESCALA_Y

    y0 = _f(y0)

    # Tabla de la viga (APoint 3D)
    dibujar_linea(acad, P(0, y0), P(L_total, y0), 8)
    dibujar_linea(acad, P(0, y0 - H_visual), P(L_total, y0 - H_visual), 8)
    dibujar_linea(acad, P(0, y0), P(0, y0 - H_visual), 8)
    dibujar_linea(acad, P(L_total, y0), P(L_total, y0 - H_visual), 8)

    titulo = f"ELEVACIÓN ESTRUCTURAL: {_s(eje.get('eje', 'Eje'))}"
    agregar_texto(acad, titulo, P(-2.5, y0 + 1.2), 0.18)

    letras_ejes = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for i, col in enumerate(columnas):
        try:
            x = _f(col.get('x', 0)); w = _f(col.get('w', 0.30))
        except Exception:
            continue
        if w <= 0.05: w = 0.30
        x_izq, x_der = x - w/2, x + w/2
        
        dibujar_linea(acad, P(x_izq, y0), P(x_izq, y0 - H_visual), 8)
        dibujar_linea(acad, P(x_der, y0), P(x_der, y0 - H_visual), 8)
        
        eje_l = dibujar_linea(acad, P(x, y0 + 0.8), P(x, y0 - H_visual - 0.5), 1)
        try:
            eje_l.Linetype = "DOT"
            eje_l.LinetypeScale = 0.5 
        except Exception: pass
        
        burbuja = acad.model.AddCircle(P(x, y0 + 1.05), 0.25)
        set_color(burbuja, 7)
        agregar_texto(acad, letras_ejes[i % 26], P(x - 0.08, y0 + 0.95), 0.25)

        crear_cota(acad, P(x_izq, y0), P(x_der, y0), P(x, y0 + 0.15)) 
        if i > 0:
            try:
                x_prev = _f(columnas[i-1].get('x', 0))
                crear_cota(acad, P(x_prev, y0), P(x, y0), P((x_prev + x)/2, y0 + 0.40)) 
            except Exception: pass

    # Rieles de barras
    def dibujar_capa_barras(lista_barras, y_base, dir_gancho, color_cad):
        tracks_cont, tracks_bast = [], []
        # ordenar seguro
        try:
            lista_sorted = sorted(lista_barras, key=lambda x: _f(x.get('x1',0)))
        except Exception:
            lista_sorted = lista_barras
        for b in lista_sorted:
            try:
                x1f = _f(b.get('x1',0)); x2f = _f(b.get('x2',0))
                rol = _s(b.get('rol',''))
            except Exception: continue
            if rol == 'Continuo':
                placed = False
                for t_idx, end_x in enumerate(tracks_cont):
                    if x1f > end_x + 0.15: 
                        tracks_cont[t_idx] = x2f
                        b['track'] = t_idx
                        placed = True
                        break
                if not placed:
                    b['track'] = len(tracks_cont)
                    tracks_cont.append(x2f)
            else: 
                placed = False
                for t_idx, end_x in enumerate(tracks_bast):
                    if x1f > end_x + 0.15:
                        tracks_bast[t_idx] = x2f
                        b['track'] = t_idx + 3 
                        placed = True
                        break
                if not placed:
                    b['track'] = len(tracks_bast) + 3
                    tracks_bast.append(x2f)

        for b in lista_sorted:
            try:
                x1, x2 = _f(b.get('x1',0)), _f(b.get('x2',0))
                g_ini, g_fin = _f(b.get('g_ini', 0)), _f(b.get('g_fin', 0))
                cant = _i(b.get('cant',0))
                tipo = _s(b.get('tipo','#5'))
                track = _i(b.get('track',0))
                rol = _s(b.get('rol',''))
            except Exception:
                continue
            
            y_actual = y_base + (dir_gancho * 0.12 * track)
            
            dibujar_linea(acad, P(x1, y_actual), P(x2, y_actual), color_cad)
            
            if g_ini > 0:
                dibujar_linea(acad, P(x1, y_actual), P(x1, y_actual + (dir_gancho * g_ini)), color_cad)
            if g_fin > 0:
                dibujar_linea(acad, P(x2, y_actual), P(x2, y_actual + (dir_gancho * g_fin)), color_cad)
                
            L_corte = (x2 - x1) + g_ini + g_fin
            texto = f"{cant}{tipo} L={L_corte:.2f}"
            
            pos_y_txt = y_actual - 0.12 if rol == "Bastón" else y_actual + 0.04
            
            t_varilla = acad.model.AddText(texto, P(0, 0), 0.08)
            try:
                t_varilla.Alignment = 1
                t_varilla.TextAlignmentPoint = P((x1 + x2)/2, pos_y_txt)
            except Exception: pass
            set_color(t_varilla, 7)

    barras_top = [dict(b) for b in barras if _s(b.get('cara','')) == "Superior"]
    barras_bot = [dict(b) for b in barras if _s(b.get('cara','')) == "Inferior"]

    dibujar_capa_barras(barras_top, y_base=y0 - 0.15, dir_gancho=-1, color_cad=5)
    dibujar_capa_barras(barras_bot, y_base=y0 - H_visual + 0.15, dir_gancho=1, color_cad=3)

    # Flejes por zona (lineas verticales a su separacion real, tal como DC-CAD)
    for z in eje.get("zonas_cortante", []) or []:
        try:
            x1z, x2z = _f(z.get('x1',0)), _f(z.get('x2',0))
            s_m = max(_f(z.get('s_col_cm',10.0)) / 100.0, 0.02)
            ramas = _i(z.get('ramas', 2))
            tipo_e = _s(z.get('tipo_est','#3'))
            s_cm_f = _f(z.get('s_col_cm',10.0))
            etiqueta = f"{tipo_e}@{s_cm_f:.0f} ({ramas}r)"
        except Exception:
            continue
        n_lines = min(int((x2z - x1z) / s_m) + 1, 400) if s_m>0 else 0
        for k in range(n_lines):
            xf = x1z + k * s_m
            if xf > x2z: break
            dibujar_linea(acad, P(xf, y0 - 0.15), P(xf, y0 - H_visual + 0.15), 4)
        agregar_texto(acad, etiqueta, P((x1z + x2z)/2 - 0.3, y0 - H_visual - 0.25), 0.09, 4)

    # Acotado de traslapos
    def acotar_traslapos(lista_barras, dir_gancho):
        continuas = [b for b in lista_barras if _s(b.get('rol','')) == 'Continuo']
        try: continuas.sort(key=lambda x: _f(x.get('x1',0)))
        except Exception: pass
        for i in range(len(continuas)):
            for j in range(i+1, len(continuas)):
                b1, b2 = continuas[i], continuas[j]
                try:
                    overlap_start = max(_f(b1.get('x1',0)), _f(b2.get('x1',0)))
                    overlap_end = min(_f(b1.get('x2',0)), _f(b2.get('x2',0)))
                except Exception: continue
                
                if overlap_end > overlap_start + 0.05: 
                    if dir_gancho == -1:
                        y_ref, y_txt = y0 - 0.15, y0 - 0.05
                    else:
                        y_ref = y0 - H_visual + 0.15
                        y_txt = y0 - H_visual + 0.65
                    crear_cota(acad, P(overlap_start, y_ref), P(overlap_end, y_ref), P((overlap_start+overlap_end)/2, y_txt))

    acotar_traslapos(barras_top, dir_gancho=-1)
    acotar_traslapos(barras_bot, dir_gancho=1)

    # --- Directriz 4.6: Líneas de corte en elevación (Dash-Dot) ---
    for corte in cortes:
        try:
            xm = _f(corte.get('x_mid', None), None)
            if xm is None: continue
            nombre = _s(corte.get('nombre','')).strip()
            # letra para indicador: primera letra antes de "-"
            letra = nombre.split('-')[0] if '-' in nombre else (nombre[:1] if nombre else "?")
            if not ( -0.5 <= xm <= L_total + 0.5): continue
            # línea vertical Dash-Dot a través de la viga extendida
            line_c = dibujar_linea(acad, P(xm, y0 + 0.6), P(xm, y0 - H_visual - 0.9), 1)
            try:
                line_c.Linetype = "DASHDOT"
                line_c.LinetypeScale = 0.4
            except Exception:
                try: line_c.Linetype = "DOT"
                except Exception: pass
            # textos arriba y abajo
            agregar_texto(acad, letra, P(xm - 0.07, y0 + 0.75), 0.13, 1)
            agregar_texto(acad, letra, P(xm - 0.07, y0 - H_visual - 0.90), 0.13, 1)
            # pequeño símbolo: círculo alrededor de letra para estilo plano
            try:
                c1 = acad.model.AddCircle(P(xm, y0+0.85), 0.12)
                set_color(c1, 1)
                c2 = acad.model.AddCircle(P(xm, y0 - H_visual -0.80), 0.12)
                set_color(c2, 1)
            except Exception: pass
        except Exception:
            continue
    
    return H_visual

# ================= CORTES DE SECCION =================

def dibujar_corte(acad, corte, cx, cy_top):
    """Dibuja un corte transversal centrado en cx, con la cara superior en cy_top. Escala S=2 sobre dimensiones reales. Incluye Directriz 4."""
    S = 2.0
    # --- Directriz 3.1: parseo estricto ---
    b = _f(corte.get('b_mm', 300)) / 1000.0 * S
    h = _f(corte.get('h_mm', 400)) / 1000.0 * S
    rec = _f(corte.get('rec_mm', 40.0)) / 1000.0 * S
    cx = _f(cx); cy_top = _f(cy_top)
    
    diam_db = {"#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1, "#7": 22.2, "#8": 25.4}
    estribo = corte.get('estribo', {}) if isinstance(corte.get('estribo'), dict) else {}
    tipo_e = _s(estribo.get('tipo', '#3'), "#3")
    ramas_e = _i(estribo.get('ramas', 2), 2)
    s_cm_e = _f(estribo.get('s_cm', 20), 20)
    df = diam_db.get(tipo_e, 9.5) / 1000.0 * S
    
    inset = rec + df/2
    
    # Concreto y estribo
    p_sup_i, p_sup_d = P(cx - b/2, cy_top), P(cx + b/2, cy_top)
    p_inf_i, p_inf_d = P(cx - b/2, cy_top - h), P(cx + b/2, cy_top - h)
    dibujar_linea(acad, p_sup_i, p_sup_d, 8)
    dibujar_linea(acad, p_inf_i, p_inf_d, 8)
    dibujar_linea(acad, p_sup_i, p_inf_i, 8)
    dibujar_linea(acad, p_sup_d, p_inf_d, 8)
    
    est_rect = [P(cx - b/2 + inset, cy_top - inset),
                P(cx + b/2 - inset, cy_top - inset),
                P(cx + b/2 - inset, cy_top - h + inset),
                P(cx - b/2 + inset, cy_top - h + inset)]
    for i in range(4):
        dibujar_linea(acad, est_rect[i], est_rect[(i+1) % 4], 4)
    if ramas_e >= 4:
        for xf in [-b/6, b/6]:
            if abs(xf) < b/2 - inset - 1e-6:
                dibujar_linea(acad, P(cx + xf, cy_top - inset), P(cx + xf, cy_top - h + inset), 4)
    
    # --- Directriz 4.4: ganchos sísmicos 135° en esquina sup. derecha interior ---
    try:
        # longitud gancho ~ 6*df o 75mm escala real, aquí en dibujo escala S
        hook_len = max(0.075 * S, 6 * df)
        # esquina interior sup derecha
        x0 = cx + b/2 - inset
        y0_hook = cy_top - inset
        # dirección 135° desde horizontal (noroeste: -cos45, -sin45? 135° desde +X ccw va NW)
        ang = math.radians(135)
        # pero en CAD Y hacia arriba, el gancho debe ir hacia interior y abajo (SW) -> 225° es SW, 135° es NW arriba; para estribo superior derecho, gancho interior va hacia abajo-izq (225°)
        # según norma 135° de doblez, el extremo va a 45° interior. Usaremos -135° (315°?).
        # Implementamos SW (225°) para entrar
        ang_sw = math.radians(225)
        x1 = x0 + hook_len * math.cos(ang_sw)
        y1 = y0_hook + hook_len * math.sin(ang_sw)
        dibujar_linea(acad, P(x0, y0_hook), P(x1, y1), 4)
        # segundo tramo corto horizontal para rematar gancho a 90° adicional (opcional detalle)
        # pequeño quiebre de 10mm hacia interior
        x2 = x1 + (0.02 * S) * math.cos(math.radians(180))
        y2 = y1
        dibujar_linea(acad, P(x1, y1), P(x2, y2), 4)
    except Exception:
        pass
    
    # Barras por cara
    def colocar(grupos, es_sup):
        try: grupos = grupos or []
        except Exception: grupos = []
        total = sum(_i(g.get('cant',0)) for g in grupos) if grupos else 0
        if total <= 0: return ""
        # db max para margen
        db_vals = [diam_db.get(_s(g.get('tipo','#5')), 15.9) for g in grupos]
        db_max = max(db_vals) / 1000.0 * S if db_vals else 15.9/1000*S
        margen_x = inset + db_max/2
        if total == 1:
            xs = [0.0]
        else:
            xs = [-b/2 + margen_x + i * (b - 2*margen_x) / max(total-1, 1) for i in range(total)]
        r_vis = max(db_max/2, 0.018)
        y = cy_top - (rec + df + db_max/2) if es_sup else cy_top - h + (rec + df + db_max/2)
        color = 5 if es_sup else 3
        k = 0
        calibre_textos = []
        for g in grupos:
            cant_g = _i(g.get('cant',0))
            tipo_g = _s(g.get('tipo','#5'))
            db = diam_db.get(tipo_g, 15.9) / 1000.0 * S
            for _ in range(cant_g):
                try:
                    circulo = acad.model.AddCircle(P(cx + xs[k], y), max(r_vis, db/2))
                    set_color(circulo, color)
                except Exception: pass
                # --- Directriz 4.3: texto calibre cerca de cada barra ---
                try:
                    # offset pequeño arriba/abajo según cara
                    off_y = (db_max + 0.04) if es_sup else -(db_max + 0.08)
                    # si hay mezcla de calibres, etiquetar cada una; si solo un tipo, igual etiquetar todas suave
                    txt_c = acad.model.AddText(tipo_g, P(cx + xs[k] - 0.04, y + off_y), 0.07)
                    set_color(txt_c, color)
                    try:
                        txt_c.Alignment = 4  # Center?
                    except Exception: pass
                except Exception: pass
                calibre_textos.append(tipo_g)
                k += 1
        return "+".join(f"{_i(g.get('cant',0))}{_s(g.get('tipo','#5'))}" for g in grupos)
    
    # parseo sup/inf con fallback
    sup_grupos = corte.get('sup', []) if isinstance(corte.get('sup'), list) else []
    inf_grupos = corte.get('inf', []) if isinstance(corte.get('inf'), list) else []
    txt_sup = colocar(sup_grupos, True)
    txt_inf = colocar(inf_grupos, False)
    txt_fleje = f"{tipo_e}@{s_cm_e:.0f} ({ramas_e}r)"
    
    # --- Directriz 4.1: Títulos superiores izquierda en dos líneas ---
    nombre = _s(corte.get('nombre',''), "X-X")
    # Si no hay nombre, usar x_mid como fallback pero preferir nombre
    try:
        titulo1 = "Detalle Sección"
        titulo2 = f"  {nombre}"
        # posición top-left del corte (ligeramente fuera del concreto)
        x_tit = cx - b/2 - 0.05
        y_tit1 = cy_top + 0.48
        y_tit2 = cy_top + 0.30
        agregar_texto(acad, titulo1, P(x_tit, y_tit1), 0.10, 7)
        t2 = agregar_texto(acad, titulo2, P(x_tit, y_tit2), 0.12, 7)
        try: t2.TextString = titulo2  # ensure
        except Exception: pass
    except Exception: pass

    # --- Directriz 4.2: Cotas físicas ancho y alto ---
    try:
        # ancho en base (horizontal)
        p1w = P(cx - b/2, cy_top - h)
        p2w = P(cx + b/2, cy_top - h)
        ptxt_w = P(cx, cy_top - h - 0.35)
        cota_w = crear_cota(acad, p1w, p2w, ptxt_w)
        # Override texto para mostrar dimensión real (b/S) en metros: ej 0.30
        try:
            real_b = _f(corte.get('b_mm',300))/1000.0
            cota_w.TextOverride = f"{real_b:.2f}"
        except Exception: pass

        # alto en lateral derecho (vertical)
        p1h = P(cx + b/2, cy_top - h)
        p2h = P(cx + b/2, cy_top)
        ptxt_h = P(cx + b/2 + 0.35, cy_top - h/2)
        cota_h = crear_cota(acad, p1h, p2h, ptxt_h)
        try:
            real_h = _f(corte.get('h_mm',400))/1000.0
            cota_h.TextOverride = f"{real_h:.2f}"
        except Exception: pass
    except Exception: pass

    # --- Directriz 4.5: Resumen material bajo cota inferior ---
    try:
        # armar líneas Sup/Inf/Estribos
        y_res_base = cy_top - h - 0.62
        # Sup
        if sup_grupos:
            # juntar por si hay varios tipos, ya agrupados; mostrar cada grupo
            for idx, g in enumerate(sup_grupos):
                cant_s = _i(g.get('cant',0)); tipo_s = _s(g.get('tipo','#5'))
                agregar_texto(acad, f"{cant_s} Barras {tipo_s} (Sup)", P(cx - b/2, y_res_base - idx*0.13), 0.09, 5)
            off_sup = len(sup_grupos)*0.13
        else:
            agregar_texto(acad, "0 Barras (Sup)", P(cx - b/2, y_res_base), 0.09, 5)
            off_sup = 0.13
        # Inf
        if inf_grupos:
            for idx, g in enumerate(inf_grupos):
                cant_i = _i(g.get('cant',0)); tipo_i = _s(g.get('tipo','#5'))
                agregar_texto(acad, f"{cant_i} Barras {tipo_i} (Inf)", P(cx - b/2, y_res_base - off_sup - idx*0.13), 0.09, 3)
            off_inf = len(inf_grupos)*0.13
        else:
            agregar_texto(acad, "0 Barras (Inf)", P(cx - b/2, y_res_base - off_sup), 0.09, 3)
            off_inf = 0.13
        # Estribos
        agregar_texto(acad, f"Estribos {tipo_e} @ {s_cm_e:.0f} ({ramas_e}r)", P(cx - b/2, y_res_base - off_sup - off_inf), 0.09, 4)

        # x_mid referencia pequeña arriba derecha (coordenada)
        agregar_texto(acad, f"x≈{_f(corte.get('x_mid',0)):.2f}m", P(cx + b/2 + 0.12, cy_top - h/2 + 0.35), 0.07, 7)
    except Exception: pass

def dibujar_fila_cortes(acad, eje, y_base_cortes, L_total):
    cortes = eje.get("cortes", []) or []
    if not cortes: return 0.0
    
    ancho_usado = 0.0
    x_cursor = 0.0
    for corte in cortes:
        try:
            b = _f(corte.get('b_mm',300)) / 1000.0 * 2.0
            cx = x_cursor + b/2 + 0.4  # margen para títulos/cotas
            dibujar_corte(acad, corte, cx, y_base_cortes)
            x_cursor += b + 1.6 + 0.5  # separación + margen cotas
            ancho_usado = x_cursor
        except Exception:
            continue
    return ancho_usado

# ================= PRINCIPAL =================

def main():
    root = tk.Tk()
    root.withdraw()
    
    ruta_json = filedialog.askopenfilename(
        title="Selecciona el JSON acumulado de Despieces",
        filetypes=[("Archivos JSON", "*.json")]
    )
    if not ruta_json: return
        
    with open(ruta_json, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if isinstance(data, dict) and 'ejes' in data:
        ejes = data['ejes']
    elif isinstance(data, dict):
        ejes = [data]
    else:
        return messagebox.showerror("Error", "Formato de JSON no reconocido.")
    
    if not ejes:
        return messagebox.showwarning("Aviso", "El JSON no contiene ejes.")

    acad = conectar_autocad()
    if acad is None:
        return messagebox.showerror(
            "Error de Conexión",
            "No hay ningún dibujo de AutoCAD abierto.\n"
            "Abre (o crea) tu plano en AutoCAD y vuelve a ejecutar.\n"
            "Los despieces se acumularán todos en ese MISMO archivo."
        )

    # Cargar variables del documento para saber hasta donde se dibujo la ultima vez.
    # Asi cada corrida agrega los ejes debajo de lo anterior, sin abrir archivos nuevos.
    try:
        y_offset = float(acad.doc.GetVariable("USERR1"))
    except Exception:
        y_offset = 0.0
    
    if y_offset == 0.0:
        try: acad.doc.SetVariable("DIMSCALE", 1.0)
        except Exception: pass
        try:
            acad.doc.SetVariable("DIMBLK", "_OBLIQUE")
            acad.doc.SetVariable("DIMTAD", 1)
            acad.doc.SetVariable("DIMTXT", 0.12)
            acad.doc.SetVariable("DIMASZ", 0.10)
        except Exception: pass
        try: acad.doc.Linetypes.Load("DOT", "acad.lin")
        except Exception: pass
        try: acad.doc.Linetypes.Load("DASHDOT", "acad.lin")
        except Exception: pass
    
    # USERR1 guarda el |y| ya consumido; dibujamos hacia abajo (y negativo)
    y_cursor = -(y_offset)
    
    resumen = []
    for eje in ejes:
        nombre = str(eje.get('eje', 'Eje'))
        L_total = _f(eje.get('L_total', 0.0))
        
        H_visual = dibujar_elevacion(acad, eje, y0=y_cursor)
        y_cortes = y_cursor - H_visual - GAP_CORTES
        ancho_cortes = dibujar_fila_cortes(acad, eje, y_base_cortes=y_cortes, L_total=L_total)
        
        alto_bloque = H_visual + GAP_CORTES + 3.6  # aumentado por resumen + cotas
        y_cursor -= alto_bloque + GAP_EJES
        resumen.append(f"• {nombre}")
    
    try:
        acad.doc.SetVariable("USERR1", abs(y_cursor))
    except Exception: pass
    
    messagebox.showinfo(
        "¡Exportación Exitosa!",
        f"Se agregaron {len(resumen)} eje(s) AL DIBUJO ACTIVO (nada se creó ni abrió nuevo):\n"
        + "\n".join(resumen[:15])
        + ("\n..." if len(resumen) > 15 else "")
        + "\n\nGuarda el dibujo cuando quieras conservar todo en un solo archivo."
    )

if __name__ == "__main__":
    main()
