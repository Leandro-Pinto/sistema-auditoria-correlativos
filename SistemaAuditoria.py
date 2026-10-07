# -*- coding: utf-8 -*-
"""
Sistema de Auditoría de Correlativos – Documentos de Gestión
Aplicación de escritorio para Windows (Python + Tkinter).
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from auditoria_core import Auditoria, COLORES, DESCRIPCION, ESTADOS

APP = "Sistema de Auditoría de Correlativos"
AZUL = "#1F4E78"
FUENTE = ("Segoe UI", 10)
# En pantalla: CONFIRMADO sin color, para que resalten los problemas
COLORES_PANTALLA = {"CONFIRMADO": "#FFFFFF", "REPETIDO": "#FFF4CC", "FALTA": "#F4B6A0",
                    "SIN NÚMERO": "#F8CBAD", "SIN AÑO": "#F8CBAD", "AÑO INCOMPLETO": "#DCE9F7",
                    "AÑO INFERIDO": "#DCE9F7", "FORMATO": "#EDF5E6", "CORREGIDO": "#E8E2F2",
                    "CON_FALT": "#FFFFFF"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP} – Documentos de Gestión")
        self.geometry("1280x760")
        self.minsize(1000, 600)
        self.aud = None
        self.config_pdf = self._leer_config()
        self._estilos()
        self._barra_superior()
        self._pestanas()
        self._barra_estado()
        self.estado("Abra un archivo Excel para comenzar (botón «Abrir Excel»).")
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self.after(200, lambda: self.cargar(sys.argv[1]))

    # ------------------------------------------------------------ configuración
    @staticmethod
    def _ruta_config():
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "SistemaAuditoriaCorrelativos.json")

    def _leer_config(self):
        try:
            import json
            with open(self._ruta_config(), encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError):
            return {}

    def guardar_config_pdf(self, entidad, auditor):
        import json
        self.config_pdf.update(entidad=entidad, auditor=auditor)
        try:
            with open(self._ruta_config(), "w", encoding="utf-8") as fh:
                json.dump(self.config_pdf, fh, ensure_ascii=False)
        except OSError:
            pass

    # ------------------------------------------------------------ estilo
    def _estilos(self):
        st = ttk.Style(self)
        try:
            st.theme_use("vista" if sys.platform == "win32" else "clam")
        except tk.TclError:
            st.theme_use("clam")
        import tkinter.font as tkfont
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(nombre).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass
        st.configure("Treeview", rowheight=24, font=FUENTE)
        st.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        st.configure("Titulo.TLabel", font=("Segoe UI", 15, "bold"), foreground=AZUL)
        st.configure("KPI.TLabel", font=("Segoe UI", 22, "bold"), foreground=AZUL)
        st.configure("KPIt.TLabel", font=("Segoe UI", 9), foreground="#555555")
        st.configure("Accion.TButton", font=("Segoe UI", 10, "bold"))

    def _barra_superior(self):
        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        ttk.Label(top, text=APP, style="Titulo.TLabel").pack(side="left")
        ttk.Button(top, text="Exportar informe PDF", style="Accion.TButton",
                   command=self.exportar_pdf).pack(side="right", padx=(6, 0))
        ttk.Button(top, text="Exportar reporte Excel", style="Accion.TButton",
                   command=self.exportar).pack(side="right", padx=(6, 0))
        ttk.Button(top, text="Recargar", command=self.recargar).pack(side="right", padx=(6, 0))
        ttk.Button(top, text="Abrir Excel…", style="Accion.TButton",
                   command=self.abrir).pack(side="right")

        sub = ttk.Frame(self, padding=(12, 0, 12, 6))
        sub.pack(fill="x")
        ttk.Label(sub, text="Tipo de documento:").pack(side="left")
        self.cb_tipo = ttk.Combobox(sub, state="readonly", width=42)
        self.cb_tipo.pack(side="left", padx=6)
        self.cb_tipo.bind("<<ComboboxSelected>>", lambda e: self.refrescar_todo())
        self.lbl_archivo = ttk.Label(sub, text="Sin archivo", foreground="#777777")
        self.lbl_archivo.pack(side="left", padx=12)

    def _barra_estado(self):
        bar = ttk.Frame(self, padding=(12, 3))
        bar.pack(fill="x", side="bottom")
        self.lbl_estado = ttk.Label(bar, text="", foreground="#444444")
        self.lbl_estado.pack(side="left")
        self.pb = ttk.Progressbar(bar, length=220, mode="determinate")

    def estado(self, txt):
        self.lbl_estado.config(text=txt)
        self.update_idletasks()

    # ------------------------------------------------------------ pestañas
    def _pestanas(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=4)
        self._tab_panel()
        self._tab_auditoria()
        self._tab_revisar()
        self._tab_datos()
        self._tab_ayuda()

    def _tree(self, parent, cols, anchos, alto=20):
        frame = ttk.Frame(parent)
        tv = ttk.Treeview(frame, columns=cols, show="headings", height=alto)
        for c, w in zip(cols, anchos):
            tv.heading(c, text=c, command=lambda c=c: self._ordenar(tv, c, False))
            tv.column(c, width=w, anchor="center" if w < 140 else "w", stretch=w >= 140)
        ys = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=ys.set)
        tv.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        for est, col in COLORES_PANTALLA.items():
            tv.tag_configure(est, background=col)
        tv.tag_configure("FALTA", background="#F4B6A0", foreground="#9C0006")
        return frame, tv

    def _ordenar(self, tv, col, desc):
        datos = [(tv.set(k, col), k) for k in tv.get_children("")]
        def clave(v):
            try:
                return (0, float(str(v[0]).replace("%", "").replace(",", ".")))
            except ValueError:
                return (1, str(v[0]))
        datos.sort(key=clave, reverse=desc)
        for i, (_, k) in enumerate(datos):
            tv.move(k, "", i)
        tv.heading(col, command=lambda: self._ordenar(tv, col, not desc))

    # ---- Panel
    def _tab_panel(self):
        f = ttk.Frame(self.nb, padding=10)
        self.nb.add(f, text="  Panel  ")
        kp = ttk.Frame(f)
        kp.pack(fill="x", pady=(0, 10))
        self.kpi = {}
        for nombre in ("Documentos registrados", "Números faltantes", "Registros repetidos",
                       "Por revisar", "Años con faltantes"):
            box = ttk.Frame(kp, padding=(14, 6), relief="groove")
            box.pack(side="left", padx=(0, 10), fill="x", expand=True)
            v = ttk.Label(box, text="—", style="KPI.TLabel")
            v.pack(anchor="w")
            ttk.Label(box, text=nombre, style="KPIt.TLabel").pack(anchor="w")
            self.kpi[nombre] = v
        cols = ("Año", "Registros", "N° distintos", "Último N°", "Faltantes",
                "Registros repetidos", "Sin número", "% Completitud", "Situación")
        fr, self.tv_panel = self._tree(f, cols, (70, 95, 110, 95, 90, 135, 105, 125, 160))
        fr.pack(fill="both", expand=True)
        self.tv_panel.bind("<Double-1>", self._panel_a_auditoria)
        ttk.Label(f, text="Doble clic en un año para auditar su secuencia completa.",
                  foreground="#666666").pack(anchor="w", pady=(6, 0))

    def _panel_a_auditoria(self, _e):
        sel = self.tv_panel.selection()
        if sel:
            self.cb_anio.set(self.tv_panel.set(sel[0], "Año"))
            self.refrescar_auditoria()
            self.nb.select(1)

    # ---- Auditoría por año
    def _tab_auditoria(self):
        f = ttk.Frame(self.nb, padding=10)
        self.nb.add(f, text="  Auditoría por año  ")
        top = ttk.Frame(f)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Año:").pack(side="left")
        self.cb_anio = ttk.Combobox(top, state="readonly", width=8)
        self.cb_anio.pack(side="left", padx=(4, 16))
        self.cb_anio.bind("<<ComboboxSelected>>", lambda e: self.refrescar_auditoria())
        ttk.Label(top, text="Mostrar:").pack(side="left")
        self.cb_filtro = ttk.Combobox(top, state="readonly", width=14,
                                      values=["Todos", "FALTA", "REPETIDO", "CONFIRMADO", "SIN NÚMERO"])
        self.cb_filtro.set("Todos")
        self.cb_filtro.pack(side="left", padx=4)
        self.cb_filtro.bind("<<ComboboxSelected>>", lambda e: self.refrescar_auditoria())
        self.lbl_aud = ttk.Label(top, text="", font=("Segoe UI", 10, "bold"), foreground=AZUL)
        self.lbl_aud.pack(side="left", padx=20)

        cuerpo = ttk.Frame(f)
        cuerpo.pack(fill="both", expand=True)
        cols = ("N°", "Veces", "Estado", "Título registrado", "id")
        fr, self.tv_aud = self._tree(cuerpo, cols, (70, 60, 110, 420, 200))
        fr.pack(side="left", fill="both", expand=True)

        der = ttk.LabelFrame(cuerpo, text=" Números faltantes ", padding=8)
        der.pack(side="right", fill="y", padx=(10, 0))
        self.txt_falt = tk.Text(der, width=26, wrap="word", font=("Consolas", 10),
                                background="#FFF5F0", relief="flat")
        self.txt_falt.pack(fill="both", expand=True)
        ttk.Button(der, text="Copiar lista", command=self._copiar_faltantes).pack(fill="x", pady=(6, 0))

    def _copiar_faltantes(self):
        self.clipboard_clear()
        self.clipboard_append(self.txt_falt.get("1.0", "end").strip())
        self.estado("Lista de faltantes copiada al portapapeles.")

    # ---- Por revisar
    def _tab_revisar(self):
        f = ttk.Frame(self.nb, padding=10)
        self.nb.add(f, text="  Por revisar  ")
        ttk.Label(f, text="Títulos con problemas. Doble clic en una fila para corregir el N° o el año.",
                  foreground="#666666").pack(anchor="w", pady=(0, 6))
        cols = ("Fila", "id", "Título", "N°", "Año", "Estado", "Detalle", "Nota")
        fr, self.tv_rev = self._tree(f, cols, (60, 90, 230, 60, 60, 120, 330, 220))
        fr.pack(fill="both", expand=True)
        self.tv_rev.bind("<Double-1>", lambda e: self._editar_desde(self.tv_rev))

    # ---- Datos
    def _tab_datos(self):
        f = ttk.Frame(self.nb, padding=10)
        self.nb.add(f, text="  Datos  ")
        top = ttk.Frame(f)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Buscar:").pack(side="left")
        self.ent_buscar = ttk.Entry(top, width=30)
        self.ent_buscar.pack(side="left", padx=4)
        self.ent_buscar.bind("<Return>", lambda e: self.refrescar_datos())
        ttk.Label(top, text="Año:").pack(side="left", padx=(12, 0))
        self.cb_anio_d = ttk.Combobox(top, state="readonly", width=8)
        self.cb_anio_d.pack(side="left", padx=4)
        ttk.Label(top, text="Estado:").pack(side="left", padx=(12, 0))
        self.cb_est_d = ttk.Combobox(top, state="readonly", width=16,
                                     values=["Todos"] + [e for e in ESTADOS if e != "FALTA"])
        self.cb_est_d.set("Todos")
        self.cb_est_d.pack(side="left", padx=4)
        ttk.Button(top, text="Filtrar", command=self.refrescar_datos).pack(side="left", padx=8)
        self.lbl_datos = ttk.Label(top, text="", foreground="#666666")
        self.lbl_datos.pack(side="left", padx=8)
        for cb in (self.cb_anio_d, self.cb_est_d):
            cb.bind("<<ComboboxSelected>>", lambda e: self.refrescar_datos())
        cols = ("Fila", "id", "Título", "N°", "Año", "Estado", "Nota")
        fr, self.tv_datos = self._tree(f, cols, (60, 90, 300, 70, 70, 130, 300))
        fr.pack(fill="both", expand=True)
        self.tv_datos.bind("<Double-1>", lambda e: self._editar_desde(self.tv_datos))

    # ---- Ayuda
    def _tab_ayuda(self):
        f = ttk.Frame(self.nb, padding=16)
        self.nb.add(f, text="  Ayuda  ")
        pasos = [
            "1. «Abrir Excel…» y elija el archivo con columnas id, Tipo de norma y Título (también acepta .csv).",
            "2. Elija el tipo de documento arriba. Cada tipo tiene su propia numeración por año.",
            "3. «Panel» resume cada año; doble clic en un año abre su auditoría.",
            "4. «Auditoría por año» muestra del N° 1 al último; los faltantes salen en rojo y en la lista de la derecha.",
            "5. «Por revisar» y «Datos»: doble clic en un registro para corregir N° o año. Las correcciones se guardan "
            "en un archivo .correcciones.json junto al Excel y se aplican cada vez que lo abra.",
            "6. «Exportar reporte Excel» genera un libro con Resumen, Faltantes, Por revisar, Todos y una hoja por año.",
            "7. «Exportar informe PDF» genera el informe de auditoría en hoja A4 horizontal: portada, resumen ejecutivo "
            "con gráficos, cuadro consolidado con nivel de riesgo, una sección por año con sus gráficos y hallazgos, "
            "conclusiones, recomendaciones y firmas.",
        ]
        for p in pasos:
            ttk.Label(f, text=p, wraplength=1100, justify="left").pack(anchor="w", pady=2)
        ttk.Label(f, text="Estados", font=("Segoe UI", 11, "bold"), foreground=AZUL).pack(anchor="w", pady=(14, 4))
        for e in ESTADOS:
            row = ttk.Frame(f)
            row.pack(anchor="w", pady=1)
            tk.Label(row, text=f" {e} ", background="#" + COLORES[e], width=16, anchor="w").pack(side="left")
            ttk.Label(row, text=DESCRIPCION[e]).pack(side="left", padx=8)

    # ------------------------------------------------------------ acciones
    def abrir(self):
        ruta = filedialog.askopenfilename(
            title="Abrir registro de documentos",
            filetypes=[("Excel o CSV", "*.xlsx *.xlsm *.csv"), ("Todos", "*.*")])
        if ruta:
            self.cargar(ruta)

    def recargar(self):
        if self.aud:
            self.cargar(self.aud.ruta)

    def cargar(self, ruta):
        self.estado("Leyendo archivo…")
        self.config(cursor="watch")
        try:
            self.aud = Auditoria(ruta)
        except Exception as ex:  # noqa: BLE001
            messagebox.showerror(APP, f"No se pudo abrir el archivo:\n\n{ex}")
            self.estado("Error al abrir el archivo.")
            return
        finally:
            self.config(cursor="")
        tipos = self.aud.tipos()
        self.cb_tipo["values"] = tipos
        if self.cb_tipo.get() not in tipos:
            self.cb_tipo.set(tipos[0] if tipos else "")
        self.lbl_archivo.config(text=os.path.basename(ruta))
        self.refrescar_todo()
        self.estado(f"{len(self.aud.registros):,} registros cargados.".replace(",", " "))

    def tipo(self):
        return self.cb_tipo.get()

    def refrescar_todo(self):
        if not self.aud:
            return
        anios = [str(a) for a in self.aud.anios(self.tipo())]
        self.cb_anio["values"] = anios
        if self.cb_anio.get() not in anios and anios:
            self.cb_anio.set(anios[-1])
        self.cb_anio_d["values"] = ["Todos"] + anios
        if not self.cb_anio_d.get():
            self.cb_anio_d.set("Todos")
        self.refrescar_panel()
        self.refrescar_auditoria()
        self.refrescar_revisar()
        self.refrescar_datos()

    def refrescar_panel(self):
        tv = self.tv_panel
        tv.delete(*tv.get_children())
        res = self.aud.resumen(self.tipo())
        for s in res:
            comp = f"{s['completitud']*100:.1f}%" if s["completitud"] is not None else ""
            tag = "CONFIRMADO"
            tv.insert("", "end", values=(s["anio"], s["registros"], s["distintos"], s["ultimo"],
                                         s["faltantes"], s["repetidos"], s["sin_numero"], comp,
                                         ("⚠ " if s["faltantes"] else "✓ ") + s["situacion"]), tags=(tag,))
        n = lambda x: f"{x:,}".replace(",", " ")
        self.kpi["Documentos registrados"].config(text=n(sum(s["registros"] for s in res)))
        self.kpi["Números faltantes"].config(text=n(sum(s["faltantes"] for s in res)))
        self.kpi["Registros repetidos"].config(text=n(sum(s["repetidos"] for s in res)))
        self.kpi["Por revisar"].config(text=n(len(self.aud.por_revisar(self.tipo()))))
        self.kpi["Años con faltantes"].config(text=f"{sum(1 for s in res if s['faltantes'])} / {len(res)}")

    def refrescar_auditoria(self):
        tv = self.tv_aud
        tv.delete(*tv.get_children())
        self.txt_falt.delete("1.0", "end")
        if not self.cb_anio.get():
            return
        anio = int(self.cb_anio.get())
        filtro = self.cb_filtro.get()
        seq = self.aud.secuencia(self.tipo(), anio)
        for fila in seq:
            if filtro != "Todos" and fila[2] != filtro:
                continue
            tv.insert("", "end", values=fila, tags=(fila[2],))
        falt = [str(x[0]) for x in seq if x[2] == "FALTA"]
        ultimo = max([x[0] for x in seq if isinstance(x[0], int)], default=0)
        comp = (ultimo - len(falt)) / ultimo * 100 if ultimo else 0
        self.lbl_aud.config(text=f"Último N°: {ultimo}    Faltantes: {len(falt)}    "
                                 f"Completitud: {comp:.1f}%")
        self.txt_falt.insert("1.0", ", ".join(falt) if falt else "Sin faltantes. Secuencia completa.")

    def refrescar_revisar(self):
        tv = self.tv_rev
        tv.delete(*tv.get_children())
        self._map_rev = {}
        for r in self.aud.por_revisar(self.tipo()):
            det = r.problema if r.problema and r.problema != r.estado else r.estado
            iid = tv.insert("", "end", values=(r.fila, r.id, r.titulo, r.n if r.n is not None else "",
                                               r.y if r.y is not None else "", r.estado,
                                               DESCRIPCION.get(det, ""), r.nota),
                            tags=(r.problema or r.estado,))
            self._map_rev[iid] = r
        self.nb.tab(2, text=f"  Por revisar ({len(self._map_rev)})  ")

    def refrescar_datos(self):
        if not self.aud:
            return
        tv = self.tv_datos
        tv.delete(*tv.get_children())
        self._map_datos = {}
        q = self.ent_buscar.get().strip().lower()
        an = self.cb_anio_d.get()
        es = self.cb_est_d.get()
        sel = [r for r in self.aud.registros if r.tipo == self.tipo()
               and (not q or q in r.titulo.lower() or q == str(r.id))
               and (an in ("", "Todos") or str(r.y) == an)
               and (es in ("", "Todos") or r.estado == es or r.problema == es)]
        sel.sort(key=lambda r: (r.y or 0, r.n if r.n is not None else 10**9, r.titulo))
        LIM = 3000
        for r in sel[:LIM]:
            iid = tv.insert("", "end", values=(r.fila, r.id, r.titulo, r.n if r.n is not None else "",
                                               r.y if r.y is not None else "", r.estado, r.nota),
                            tags=(r.estado if r.estado != "CONFIRMADO" else (r.problema or "CONFIRMADO"),))
            self._map_datos[iid] = r
        extra = f" (se muestran los primeros {LIM}; use los filtros)" if len(sel) > LIM else ""
        self.lbl_datos.config(text=f"{len(sel)} registros{extra}")

    # ------------------------------------------------------------ corrección
    def _editar_desde(self, tv):
        sel = tv.selection()
        if not sel:
            return
        mapa = self._map_rev if tv is self.tv_rev else self._map_datos
        reg = mapa.get(sel[0])
        if reg:
            DialogoCorreccion(self, reg)

    def aplicar_correccion(self, reg, n, y, nota):
        self.aud.guardar_correccion(reg, n, y, nota)
        self.refrescar_todo()
        self.estado(f"Corrección guardada para «{reg.titulo}».")

    # ------------------------------------------------------------ exportar
    def exportar(self):
        if not self.aud:
            messagebox.showinfo(APP, "Primero abra un archivo Excel.")
            return
        base = os.path.splitext(os.path.basename(self.aud.ruta))[0]
        destino = filedialog.asksaveasfilename(
            title="Guardar reporte de auditoría", defaultextension=".xlsx",
            initialfile=f"Reporte_auditoria_{base}.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not destino:
            return
        solo = messagebox.askyesno(APP, f"¿Exportar solo «{self.tipo()}»?\n\n"
                                        "Sí: solo este tipo de documento.\nNo: todos los tipos.")
        tipos = [self.tipo()] if solo else None
        self.pb.pack(side="right")
        self.pb["value"] = 0
        self.estado("Generando reporte Excel…")

        def prog(p):
            self.after(0, lambda: self.pb.configure(value=p * 100))

        def trabajo():
            try:
                self.aud.exportar_excel(destino, tipos, prog)
                self.after(0, lambda: self._export_ok(destino))
            except PermissionError:
                self.after(0, lambda: self._export_err("El archivo está abierto en Excel. Ciérrelo e intente de nuevo."))
            except Exception as ex:  # noqa: BLE001
                self.after(0, lambda: self._export_err(str(ex)))

        threading.Thread(target=trabajo, daemon=True).start()

    def exportar_pdf(self):
        if not self.aud:
            messagebox.showinfo(APP, "Primero abra un archivo Excel.")
            return
        DialogoPDF(self)

    def generar_pdf(self, opciones):
        base = os.path.splitext(os.path.basename(self.aud.ruta))[0]
        destino = filedialog.asksaveasfilename(
            title="Guardar informe PDF", defaultextension=".pdf",
            initialfile=f"Informe_auditoria_{base}.pdf", filetypes=[("PDF", "*.pdf")])
        if not destino:
            return
        self.pb.pack(side="right")
        self.pb["value"] = 0
        self.estado("Generando informe PDF…")

        def prog(p):
            self.after(0, lambda: self.pb.configure(value=p * 100))

        def trabajo():
            try:
                from reporte_pdf import exportar_pdf
                exportar_pdf(self.aud, self.tipo(), destino, progreso=prog, **opciones)
                self.after(0, lambda: self._export_ok(destino))
            except PermissionError:
                self.after(0, lambda: self._export_err("El PDF está abierto en otro programa. Ciérrelo e intente de nuevo."))
            except Exception as ex:  # noqa: BLE001
                self.after(0, lambda: self._export_err(str(ex)))

        threading.Thread(target=trabajo, daemon=True).start()

    def _export_ok(self, destino):
        self.pb.pack_forget()
        self.estado(f"Reporte guardado: {destino}")
        if messagebox.askyesno(APP, "Archivo generado.\n\n¿Desea abrirlo ahora?"):
            try:
                os.startfile(destino)  # Windows
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open" if sys.platform.startswith("linux") else "open", destino])

    def _export_err(self, msg):
        self.pb.pack_forget()
        self.estado("No se pudo generar el reporte.")
        messagebox.showerror(APP, f"No se pudo generar el reporte:\n\n{msg}")


class DialogoPDF(tk.Toplevel):
    """Datos del informe: entidad, responsable, años y nivel de detalle."""
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Informe de auditoría en PDF")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()
        cfg = app.config_pdf
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        ttk.Label(f, text=f"Tipo de documento: {app.tipo()}", font=("Segoe UI", 10, "bold"),
                  foreground=AZUL).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(f, text="Entidad:").grid(row=1, column=0, sticky="w", pady=3)
        self.e_ent = ttk.Entry(f, width=48)
        self.e_ent.insert(0, cfg.get("entidad", ""))
        self.e_ent.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(f, text="Responsable:").grid(row=2, column=0, sticky="w", pady=3)
        self.e_aud = ttk.Entry(f, width=48)
        self.e_aud.insert(0, cfg.get("auditor", ""))
        self.e_aud.grid(row=2, column=1, sticky="w", pady=3)

        ttk.Label(f, text="Años a incluir:").grid(row=3, column=0, sticky="nw", pady=(8, 3))
        caja = ttk.Frame(f)
        caja.grid(row=3, column=1, sticky="w", pady=(8, 3))
        self.lb = tk.Listbox(caja, selectmode="extended", height=8, width=14, exportselection=False)
        self.anios = app.aud.anios(app.tipo())
        for a in self.anios:
            self.lb.insert("end", a)
        self.lb.select_set(0, "end")
        sb = ttk.Scrollbar(caja, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left")
        sb.pack(side="left", fill="y")
        bt = ttk.Frame(caja)
        bt.pack(side="left", padx=8, anchor="n")
        ttk.Button(bt, text="Todos", command=lambda: self.lb.select_set(0, "end")).pack(fill="x")
        ttk.Button(bt, text="Ninguno", command=lambda: self.lb.select_clear(0, "end")).pack(fill="x", pady=4)
        ttk.Label(bt, text="Ctrl+clic para elegir\nvarios años.", foreground="#666666").pack()

        self.v_det = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Incluir detalle completo de la secuencia (todas las numeraciones; más páginas)",
                        variable=self.v_det).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(f, text="El informe sale en hoja A4 horizontal: portada, resumen ejecutivo con gráficos, cuadro "
                          "consolidado,\nuna sección por año (gráficos, hallazgo, faltantes, repetidos, "
                          "observaciones), conclusiones y firmas.",
                  foreground="#666666", justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        ttk.Button(b, text="Generar PDF", style="Accion.TButton", command=self.ok).pack(side="left", padx=4)
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def ok(self):
        sel = [self.anios[i] for i in self.lb.curselection()]
        if not sel:
            messagebox.showwarning(APP, "Elija al menos un año.", parent=self)
            return
        opciones = dict(entidad=self.e_ent.get().strip(), auditor=self.e_aud.get().strip(),
                        anios=sel, detalle=self.v_det.get())
        self.app.guardar_config_pdf(opciones["entidad"], opciones["auditor"])
        self.destroy()
        self.app.generar_pdf(opciones)


class DialogoCorreccion(tk.Toplevel):
    def __init__(self, app, reg):
        super().__init__(app)
        self.app, self.reg = app, reg
        self.title("Corregir registro")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        ttk.Label(f, text="Título original:").grid(row=0, column=0, sticky="w")
        ttk.Label(f, text=reg.titulo, font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(f, text=f"id {reg.id}  ·  fila {reg.fila}  ·  estado {reg.estado}",
                  foreground="#666666").grid(row=1, column=1, sticky="w", pady=(0, 10))
        self.e_n = self._campo(f, 2, "N° de documento:", reg.n)
        self.e_y = self._campo(f, 3, "Año:", reg.y)
        ttk.Label(f, text="Nota del auditor:").grid(row=4, column=0, sticky="w", pady=4)
        self.e_nota = ttk.Entry(f, width=46)
        self.e_nota.insert(0, reg.nota or "")
        self.e_nota.grid(row=4, column=1, sticky="w", pady=4)
        b = ttk.Frame(f)
        b.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        if str(reg.id) in app.aud.correcciones:
            ttk.Button(b, text="Quitar corrección", command=self.quitar).pack(side="left", padx=4)
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        ttk.Button(b, text="Guardar corrección", style="Accion.TButton",
                   command=self.guardar).pack(side="left", padx=4)
        self.bind("<Return>", lambda e: self.guardar())
        self.bind("<Escape>", lambda e: self.destroy())
        self.e_n.focus_set()
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _campo(self, f, fila, texto, valor):
        ttk.Label(f, text=texto).grid(row=fila, column=0, sticky="w", pady=4)
        e = ttk.Entry(f, width=12)
        e.insert(0, "" if valor is None else str(valor))
        e.grid(row=fila, column=1, sticky="w", pady=4)
        return e

    def guardar(self):
        try:
            n = int(self.e_n.get()) if self.e_n.get().strip() else None
            y = int(self.e_y.get()) if self.e_y.get().strip() else None
        except ValueError:
            messagebox.showwarning(APP, "N° y año deben ser números enteros.", parent=self)
            return
        if y is not None and not (1900 <= y <= 2100):
            messagebox.showwarning(APP, "El año debe tener 4 dígitos (ej. 2024).", parent=self)
            return
        if n is None or y is None:
            messagebox.showwarning(APP, "Complete el N° y el año.", parent=self)
            return
        self.app.aplicar_correccion(self.reg, n, y, self.e_nota.get().strip())
        self.destroy()

    def quitar(self):
        self.app.aplicar_correccion(self.reg, None, None, "")
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
