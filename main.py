from proceso_excel import * 
import tkinter
from tkinter import messagebox

try:
    app = ventana_principal()
    app.ventana.mainloop()
except Exception as error:
    import traceback
    traceback.print_exc()
    messagebox.showerror(
        "Error al iniciar",
        f"No se pudo iniciar correctamente el programa.\n\n{error}"
    )
