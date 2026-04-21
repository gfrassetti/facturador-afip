"""
Interfaz gráfica — Facturador AFIP
  python facturador_ui.py
  Sin consola: doble clic en Iniciar_facturador.vbs o: pyw -3 facturador_ui.py
  .exe sin ventana negra: pyinstaller --onefile --windowed --name FacturadorAFIP facturador_ui.py
"""

import math
import queue
import sys
import threading
from typing import Optional
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from bot import ejecutar_facturador, limpiar_progreso, resumen_progreso_excel

# Captura de ejemplo para la ayuda (tooltip)
RUTA_IMG_GUIA_EXCEL = _ROOT / "static" / "screenshoot.png"


TEXTO_GUIA_EXCEL = """Cómo armar tu planilla

Usá un archivo Excel (.xlsx). La primera fila puede tener títulos para vos
(como «Código de venta», «Fecha»…); no hace falta un nombre exacto.

Empezá a cargar ventas desde la segunda fila, de izquierda a derecha:

1) Código de venta — número que identifica esa factura (solo en la primera
   fila de cada venta).
2) Fecha — la misma fila que el código (podés usar la fecha de Excel o
   escribirla como 11/02/2026).
3) Código del servicio o producto (por ejemplo 01 o 04).
4) Nombre o detalle del servicio.
5) Importe de esa línea (solo números).

¿Una venta con varios ítems? En las filas de abajo cargá cada ítem con su
código, nombre e importe, y dejá vacíos el código de venta y la fecha hasta
empezar otra factura.

Guardá en formato Excel (.xlsx) y elegí el archivo acá."""


class SimpleTooltip:
    """
    Un solo contenido: solo texto O solo imagen (dos tooltips distintos en la UI).
    Posición según el puntero al entrar; cierre diferido para poder pasar al cuadro.
    """

    _SHOW_MS = 280
    _HIDE_MS = 450

    def __init__(
        self,
        widget: tk.Widget,
        *,
        texto: Optional[str] = None,
        imagen: Optional[Path] = None,
        wraplength: int = 440,
        max_ancho_imagen: int = 680,
    ):
        if (texto is None) == (imagen is None):
            raise ValueError("SimpleTooltip: pasá solo texto=... o solo imagen=...")
        self.widget = widget
        self.texto = texto
        self.imagen = imagen
        self.wraplength = wraplength
        self.max_ancho_imagen = max_ancho_imagen
        self.tip: Optional[tk.Toplevel] = None
        self._photo: Optional[tk.PhotoImage] = None
        self._id_show: Optional[str] = None
        self._id_hide: Optional[str] = None
        self._x_root = 0
        self._y_root = 0

        widget.bind("<Enter>", self._enter)
        widget.bind("<Motion>", self._motion)
        widget.bind("<Leave>", self._leave)
        widget.bind("<ButtonPress>", self._cerrar_ya)

    def _motion(self, e):
        self._x_root, self._y_root = e.x_root, e.y_root

    def _enter(self, e):
        self._x_root, self._y_root = e.x_root, e.y_root
        self._cancel_hide()
        self._cancel_show()
        self._id_show = self.widget.after(self._SHOW_MS, self._abrir)

    def _leave(self, _e=None):
        self._cancel_show()
        self._programar_cierre()

    def _cancel_show(self):
        if self._id_show:
            self.widget.after_cancel(self._id_show)
            self._id_show = None

    def _cancel_hide(self):
        if self._id_hide:
            self.widget.after_cancel(self._id_hide)
            self._id_hide = None

    def _programar_cierre(self):
        self._cancel_hide()
        self._id_hide = self.widget.after(self._HIDE_MS, self._cerrar)

    def _cerrar_ya(self, _e=None):
        self._cancel_show()
        self._cancel_hide()
        self._cerrar()

    def _cerrar(self):
        self._id_hide = None
        if self.tip:
            try:
                self.tip.destroy()
            except tk.TclError:
                pass
            self.tip = None
        self._photo = None

    def _cargar_foto(self) -> Optional[tk.PhotoImage]:
        if not self.imagen or not self.imagen.is_file():
            return None
        try:
            photo = tk.PhotoImage(file=str(self.imagen))
            w = photo.width()
            if w > self.max_ancho_imagen:
                step = max(2, math.ceil(w / self.max_ancho_imagen))
                photo = photo.subsample(step, step)
            return photo
        except tk.TclError:
            return None

    def _abrir(self):
        self._id_show = None
        self._cancel_hide()
        self._cerrar()

        root = self.widget.winfo_toplevel()
        self.tip = tk.Toplevel(root)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_attributes("-topmost", True)
        self.tip.transient(root)
        bg = "#fffde7"
        borde = "#bdae5e"
        marco = tk.Frame(self.tip, bg=bg, highlightthickness=1, highlightbackground=borde)
        marco.pack()

        if self.texto is not None:
            tk.Label(
                marco,
                text=self.texto,
                justify=tk.LEFT,
                bg=bg,
                fg="#263238",
                font=("Segoe UI", 9),
                wraplength=self.wraplength,
                padx=12,
                pady=10,
            ).pack()
        else:
            self._photo = self._cargar_foto()
            if self._photo:
                tk.Label(marco, image=self._photo, bg=bg).pack(padx=8, pady=8)
            else:
                tk.Label(
                    marco,
                    text="(No se encontró la imagen de ejemplo)",
                    bg=bg,
                    fg="#666",
                    padx=12,
                    pady=8,
                ).pack()

        def en_tooltip(_e=None):
            self._cancel_hide()

        def salir_tooltip(_e=None):
            self._programar_cierre()

        for w in (self.tip, marco, *marco.winfo_children()):
            w.bind("<Enter>", en_tooltip)
            w.bind("<Leave>", salir_tooltip)

        root.update_idletasks()
        self.tip.update_idletasks()
        tw = max(self.tip.winfo_reqwidth(), 80)
        th = max(self.tip.winfo_reqheight(), 40)

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        m = 8
        x = self._x_root + 14
        y = self._y_root + 14
        if x + tw > sw - m:
            x = max(m, sw - tw - m)
        if y + th > sh - m:
            y = max(m, sh - th - m)
        if x < m:
            x = m
        if y < m:
            y = m

        self.tip.geometry(f"+{x}+{y}")
        self.tip.lift()


class FacturadorApp:
    BG = "#eef1f6"
    CARD = "#ffffff"
    ACCENT = "#1e3a5f"
    OK = "#0d7d3c"
    ERR = "#b71c1c"
    INFO = "#1565c0"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Facturador AFIP")
        self.root.minsize(720, 560)
        self.root.configure(bg=self.BG)

        self.var_excel = tk.StringVar()
        self.var_cuit = tk.StringVar()
        self.var_password = tk.StringVar()
        self.var_hoja = tk.StringVar(value="Facturador")

        self._log_queue: queue.Queue = queue.Queue()
        self._running = False

        self._setup_styles()

        outer = tk.Frame(self.root, bg=self.BG, padx=18, pady=16)
        outer.pack(fill=tk.BOTH, expand=True)

        header = tk.Frame(outer, bg=self.BG)
        header.pack(fill=tk.X, pady=(0, 14))
        tk.Label(
            header,
            text="Facturador AFIP",
            font=("Segoe UI", 20, "bold"),
            fg=self.ACCENT,
            bg=self.BG,
        ).pack(anchor=tk.W)
        tk.Label(
            header,
            text="Completá los datos y pulsá «Empezar» para facturar según el Excel.",
            font=("Segoe UI", 10),
            fg="#455a64",
            bg=self.BG,
        ).pack(anchor=tk.W, pady=(4, 0))

        card = tk.Frame(outer, bg=self.CARD, padx=16, pady=14, highlightthickness=1, highlightbackground="#cfd8dc")
        card.pack(fill=tk.X, pady=(0, 10))

        # Excel (+ ayuda formato: tooltip)
        r0 = tk.Frame(card, bg=self.CARD)
        r0.pack(fill=tk.X, pady=(0, 10))
        r0_lbl = tk.Frame(r0, bg=self.CARD)
        r0_lbl.pack(fill=tk.X)
        lbl_excel = tk.Label(
            r0_lbl,
            text="Archivo Excel (.xlsx)",
            font=("Segoe UI", 9, "bold"),
            bg=self.CARD,
            fg="#37474f",
        )
        lbl_excel.pack(side=tk.LEFT, anchor=tk.W)
        lbl_ayuda = tk.Label(
            r0_lbl,
            text="  ⓘ Cómo armar la planilla",
            font=("Segoe UI", 9),
            bg=self.CARD,
            fg=self.INFO,
            cursor="hand2",
        )
        lbl_ayuda.pack(side=tk.LEFT, anchor=tk.W)
        SimpleTooltip(lbl_ayuda, texto=TEXTO_GUIA_EXCEL)

        lbl_captura = tk.Label(
            r0_lbl,
            text="  🖼 Ejemplo visual",
            font=("Segoe UI", 9),
            bg=self.CARD,
            fg=self.INFO,
            cursor="hand2",
        )
        lbl_captura.pack(side=tk.LEFT, anchor=tk.W, padx=(10, 0))
        SimpleTooltip(lbl_captura, imagen=RUTA_IMG_GUIA_EXCEL)

        SimpleTooltip(
            lbl_excel,
            texto="Pasá el mouse sobre «ⓘ Cómo armar la planilla» o «🖼 Ejemplo visual».",
        )
        row = tk.Frame(r0, bg=self.CARD)
        row.pack(fill=tk.X, pady=(4, 0))
        ent_excel = ttk.Entry(row, textvariable=self.var_excel, width=70)
        ent_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(row, text="Examinar…", command=self._examinar_excel).pack(side=tk.RIGHT)

        # Credenciales
        cred = ttk.LabelFrame(card, text=" Credenciales AFIP ", padding=(12, 8))
        cred.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(cred, text="CUIT (sin guiones)").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(cred, textvariable=self.var_cuit, width=48).grid(
            row=0, column=1, sticky=tk.EW, padx=(12, 0), pady=4
        )
        ttk.Label(cred, text="Contraseña").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(cred, textvariable=self.var_password, show="•", width=48).grid(
            row=1, column=1, sticky=tk.EW, padx=(12, 0), pady=4
        )
        cred.columnconfigure(1, weight=1)

        r2 = tk.Frame(card, bg=self.CARD)
        r2.pack(fill=tk.X)
        lbl_hoja = tk.Label(
            r2,
            text="Nombre de la hoja",
            font=("Segoe UI", 9, "bold"),
            bg=self.CARD,
            fg="#37474f",
        )
        lbl_hoja.pack(side=tk.LEFT)
        SimpleTooltip(
            lbl_hoja,
            texto=(
                "Es el nombre de la pestaña del Excel (abajo, como las solapas). "
                "Tiene que coincidir con la que usás en el archivo. Si no la cambiaste, dejá «Facturador»."
            ),
            wraplength=320,
        )
        ttk.Entry(r2, textvariable=self.var_hoja, width=28).pack(side=tk.LEFT, padx=(12, 0))

        # Botones
        btns = tk.Frame(outer, bg=self.BG)
        btns.pack(fill=tk.X, pady=(4, 8))
        self.btn_run = ttk.Button(btns, text="Empezar / Correr programa", command=self._ejecutar, style="Accent.TButton")
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10), ipady=4, ipadx=12)
        ttk.Button(btns, text="Estado y reinicio de progreso…", command=self._dialogo_progreso).pack(
            side=tk.LEFT, padx=(0, 8), ipady=4
        )
        ttk.Button(btns, text="Actualizar resumen", command=self._refrescar_estado).pack(side=tk.LEFT, ipady=4)

        self.lbl_estado = tk.Label(
            outer,
            text="Resumen: —",
            font=("Segoe UI", 9),
            fg="#546e7a",
            bg=self.BG,
            wraplength=680,
            justify=tk.LEFT,
        )
        self.lbl_estado.pack(anchor=tk.W, pady=(0, 6))

        tk.Label(outer, text="Registro (OK / error por factura)", font=("Segoe UI", 9, "bold"), bg=self.BG, fg="#37474f").pack(
            anchor=tk.W
        )
        self.txt = scrolledtext.ScrolledText(
            outer,
            height=18,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#fafafa",
            fg="#263238",
            insertbackground=self.ACCENT,
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cfd8dc",
        )
        self.txt.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.txt.tag_configure("ok", foreground=self.OK)
        self.txt.tag_configure("err", foreground=self.ERR)
        self.txt.tag_configure("info", foreground=self.INFO)

        self._poll_log()
        self._log_plain(
            "Indicá Excel, CUIT, contraseña y hoja. [OK] = factura generada · [ERR] = fallo · [INFO] = avisos.\n"
        )
        self._refrescar_estado()

    def _setup_styles(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def _log_plain(self, text: str):
        self.txt.insert(tk.END, text)
        self.txt.see(tk.END)

    def _examinar_excel(self):
        path = filedialog.askopenfilename(
            title="Elegir Excel",
            filetypes=[("Libro Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if path:
            self.var_excel.set(path)
            self._refrescar_estado()

    def _refrescar_estado(self):
        path = self.var_excel.get().strip()
        hoja = self.var_hoja.get().strip()
        if not path or not hoja:
            self.lbl_estado.config(text="Resumen: completá archivo y hoja para ver progreso.")
            return
        r = resumen_progreso_excel(path, hoja)
        if not r:
            self.lbl_estado.config(text="Resumen: no se pudo leer el archivo o la hoja indicada.")
            return
        self.lbl_estado.config(
            text=(
                f"Resumen — Última fila OK (checkpoint): {r['ultima_fila_checkpoint']} · "
                f"Filas con venta en Excel: {r['facturas_en_excel']} · "
                f"Estimadas hechas: {r['estimadas_ya_cargadas']} · "
                f"Pendientes: {r['pendientes']}"
            )
        )

    def _dialogo_progreso(self):
        path = self.var_excel.get().strip()
        hoja = self.var_hoja.get().strip()
        if not path or not hoja:
            messagebox.showwarning(
                "Faltan datos",
                "Indicá la ruta del Excel y el nombre de la hoja para calcular el progreso.",
            )
            return
        r = resumen_progreso_excel(path, hoja)
        if not r:
            messagebox.showerror("Error", "No se pudo leer el archivo o la hoja.")
            return
        msg = (
            f"Última fila guardada (checkpoint): {r['ultima_fila_checkpoint']}\n\n"
            f"Filas con código de venta en el Excel: {r['facturas_en_excel']}\n"
            f"Estimadas ya facturadas: {r['estimadas_ya_cargadas']}\n"
            f"Pendientes estimadas: {r['pendientes']}\n\n"
            "Si borrás el checkpoint, la próxima ejecución volverá a intentar todas las filas "
            "(el bot no duplica en AFIP, pero conviene coordinar).\n\n"
            "¿Borrar archivo de progreso y empezar desde cero la próxima vez?"
        )
        if messagebox.askyesno("Progreso y reinicio", msg, icon="question"):
            limpiar_progreso()
            messagebox.showinfo("Listo", "Checkpoint borrado. La próxima corrida no omitirá filas por progreso.")
            self._refrescar_estado()
            self._log_plain("[INFO] Checkpoint borrado manualmente.\n")

    def _log_print(self, *args, **kwargs):
        kwargs.pop("flush", None)
        msg = " ".join(str(a) for a in args)
        s = msg.strip()
        tag = None
        if s.startswith("[OK]"):
            tag = "ok"
        elif s.startswith("[ERR]"):
            tag = "err"
        elif s.startswith("[INFO]"):
            tag = "info"
        self._log_queue.put((msg + "\n", tag))

    def _poll_log(self):
        try:
            while True:
                line, tag = self._log_queue.get_nowait()
                if tag:
                    self.txt.insert(tk.END, line, (tag,))
                else:
                    self.txt.insert(tk.END, line)
                self.txt.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_log)

    def _al_fin_lote_ui(self):
        messagebox.showinfo(
            "Carga terminada",
            "Se procesaron todas las filas con datos de este Excel.\n\n"
            "El progreso local se reinició; podés cargar otro archivo o volver a ejecutar.",
        )
        self._refrescar_estado()

    def _ejecutar(self):
        if self._running:
            return
        path = self.var_excel.get().strip()
        cuit = self.var_cuit.get().strip()
        password = self.var_password.get()
        hoja = self.var_hoja.get().strip()

        if not path:
            messagebox.showerror("Error", "Elegí un archivo Excel (.xlsx).")
            return
        if not cuit:
            messagebox.showerror("Error", "Ingresá el CUIT.")
            return
        if not password:
            messagebox.showerror("Error", "Ingresá la contraseña.")
            return
        if not hoja:
            messagebox.showerror("Error", "Ingresá el nombre de la hoja.")
            return

        self._running = True
        self.btn_run.configure(state=tk.DISABLED)
        self._log_plain("\n──────── Inicio de ejecución ────────\n")

        def al_fin():
            self.root.after(0, self._al_fin_lote_ui)

        def worker():
            # Todos los datos vienen de los campos de arriba; se pasan explícito a bot.ejecutar_facturador
            try:
                ejecutar_facturador(
                    path,
                    cuit,
                    password,
                    hoja,
                    log_print=self._log_print,
                    al_terminar_lote=al_fin,
                )
            except Exception as e:
                self._log_queue.put((f"\n[ERR] Error general: {e}\n", "err"))
            finally:
                self._log_queue.put(("\n──────── Fin de ejecución ────────\n", None))
                self.root.after(0, self._fin_ejecucion)

        threading.Thread(target=worker, daemon=True).start()

    def _fin_ejecucion(self):
        self._running = False
        self.btn_run.configure(state=tk.NORMAL)
        self._refrescar_estado()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    FacturadorApp().run()
