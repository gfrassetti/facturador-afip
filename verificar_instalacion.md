# Verificación de instalación (esta PC)

En la otra PC funcionaba; en esta puede faltar algo. Revisá lo siguiente.

## 1. Python

- Tener **Python 3** instalado (recomendado 3.8+).
- En la terminal: `python --version` o `py --version`.

## 2. Dependencias Python

Desde la carpeta del proyecto, con el entorno activado si usás venv:

```bash
pip install -r requirements.txt
```

O a mano:

```bash
pip install openpyxl selenium webdriver-manager
```

## 2b. Módulos de `facturador_ui.py` (stdlib, sin pip)

| Módulo | Nota |
|--------|------|
| `queue`, `threading` | Siempre con Python |
| `pathlib` | Siempre con Python |
| `tkinter` (+ `filedialog`, `messagebox`, `scrolledtext`, `ttk`) | Tcl/Tk con el instalador de python.org |
| `sys` | Stdlib (solo uso interno de rutas) |

Nada de esto va en `pip install`; si `import tkinter` falla, reinstalá Python con Tcl/Tk marcado.

## 3. Google Chrome

- Tener **Google Chrome** instalado (no solo Edge).
- `webdriver-manager` descarga el ChromeDriver que coincide con tu versión de Chrome; si Chrome no está o está muy viejo, puede fallar.

## 4. Entorno virtual (recomendado)

Si en la otra PC usabas un `venv`:

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python bot.py
```

## 5. Si algo falla al ejecutar

- Si dice **"Chrome not found"** o error de ChromeDriver: instalá/actualizá Chrome.
- Si dice **"No module named 'openpyxl'"** (o selenium, etc.): `pip install -r requirements.txt`.
- Si **no se abre la ventana** o se cierra sola: revisá que no haya otro Chrome en segundo plano y que el antivirus no bloquee el script.
