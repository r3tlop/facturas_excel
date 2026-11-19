import tkinter
from tkinter import ttk
import win32com.client
import openpyxl
import complementos_proceso as complemento
import os
import sys
from datetime import datetime 
from openpyxl.styles import Font, PatternFill, Alignment
import shutil
import re
import subprocess
import traceback
import json

def ruta_recurso(nombre_archivo):
    if getattr(sys, "frozen", False):
        carpeta_base = sys._MEIPASS
    else:
        carpeta_base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(carpeta_base, nombre_archivo)

VERSION_PROGRAMA = "1.4.4"
NOMBRE_PROGRAMA = "Sistema de Facturas"

class ventana_principal:
    def __init__(self):
        self.excel = excel_datos()
        self.archivo_seleccionado = None
        self.hoja_seleccionada = None
        self.fila_seleccionada = None
        self.ventana_resultados = None
        # Colores de la interfaz
        self.color_fondo = "#0F172A" 
        self.color_primario = "#1E3A5F" 
        self.color_formulario = "#1E293B" 
        self.color_texto_claro = "#FFFFFF" 
        self.color_texto = "#E2E8F0" 
        self.color_etiquetas = "#E2E8F0"
        self.color_secundario = "#94A3B8"
        self.color_entrada = "#334155"
        # Aqui se crea la ventana propia para el programa
        self.ventana = tkinter.Tk()
        self.ventana.title(f"{NOMBRE_PROGRAMA} | v{VERSION_PROGRAMA}")
        self.ventana.geometry("750x640")
        self.ventana.minsize(700, 600)
        self.ventana.iconbitmap(ruta_recurso("iconos/icono_principal.ico"))
        self.ventana.configure(bg = self.color_fondo)
        self.crear_menu_superior()
        # Encabezado 
        self.encabezado = tkinter.Frame(
            self.ventana,
            bg = self.color_primario,
            height = 90
        )
        self.encabezado.pack(
            fill = "x",
            padx = 20,
            pady = (20, 10)
        )
        self.encabezado.pack_propagate(False)
        # Aqui se le asigna un titulo a la ventana
        self.titulo = tkinter.Label(
            self.encabezado,
            text = "REGISTRO DE INFORMACION DE FACTURAS",
            font = ("Segoe UI", 16, "bold"),
            bg = self.color_primario,
            fg = self.color_texto_claro
        )
        self.titulo.pack(fill = "x", pady = (14,0))
        # Subtitulo
        self.subtitulo = tkinter.Label(
            self.encabezado,
            text = "Administracion y seguimiento de facturas",
            font = ("Segoe UI", 10),
            bg = self.color_primario,
            fg = "#DCE6F1"
        )
        self.subtitulo.pack(
            fill = "x",
            pady = (4, 0)
        )
        self.campos()
        self.botones()
        self.ventana.after(500, self.cargar_datos_iniciales)
        self.ventana.after(1500, self.revisar_carpeta_facturas)
    
    def campos(self):
        # Se agregan los campos a continuacion 
        self.macro_campos = tkinter.Frame(
            self.ventana,
            bg = self.color_formulario,
            padx = 25,
            pady = 15
            )
        self.macro_campos.pack(fill = "x", padx = 20, pady = 10)
        # Se crean las 2 columnas
        self.macro_campos.columnconfigure(0, weight = 1)
        self.macro_campos.columnconfigure(1, weight = 1)
        # Titulo del formulario
        tkinter.Label(
            self.macro_campos,
            text = "INFORMACION DE LA FACTURA",
            font = ("Segoe UI", 11, "bold"),
            bg = self.color_formulario,
            fg = self.color_etiquetas
        ).grid(
            row = 0,
            column = 0,
            columnspan = 2,
            sticky = "w",
            pady = (0, 15)
        )
        # "Fecha / Fila 1"
        tkinter.Label(
            self.macro_campos,
            text = "Fecha",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row = 1, column = 0, sticky = "w", padx = 10)
        self.entrada_fecha = tkinter.Entry(
            self.macro_campos, 
            state = "readonly",
            font = ("Segoe UI", 11),
            readonlybackground = self.color_entrada,
            fg = self.color_texto_claro,
            relief = "flat",
            bd = 0
            )
        self.entrada_fecha.grid(row = 2, column = 0, sticky = "ew", padx =10, pady = (5, 12), ipady = 5 )
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        self.mostrar_fecha(fecha_actual)
        # "Numero de Cotizacion / Fila 1"
        tkinter.Label(
            self.macro_campos, 
            text = "Numero de Cotizacion",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=1, column=1,sticky = "w", padx=10)
        self.entrada_cotizacion = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_cotizacion.grid(row = 2, column = 1, sticky = "ew", padx=10, pady=(5, 12), ipady =5)
        # "Numero Orde de Compra" / Fila 2"
        tkinter.Label(
            self.macro_campos, 
            text = "Numero Orden de Compra",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=3, column=0,sticky = "w", padx=10)
        self.entrada_ord_cot = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_ord_cot.grid(row=4, column=0,sticky = "ew", padx=10, pady=(5, 12), ipady= 5)
        # "Cliente / Fila 2"
        tkinter.Label(
            self.macro_campos, 
            text = "Cliente",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=3, column=1, sticky = "w", padx=10)
        self.entrada_cliente = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_cliente.grid(row=4, column=1,sticky= "ew", padx=10, pady=(5, 12), ipady= 5)
        # "Folio Factura / Fila 3"
        tkinter.Label(
            self.macro_campos, 
            text = "Folio Factura",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=5, column=0,sticky= "w", padx=10)
        self.entrada_fol_fact = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_fol_fact.grid(row=6, column=0, sticky="ew", padx=10, pady=(5, 12), ipady=5)
        # "Estatus" solo permite las opciones brindadas para evitar los errores de ortografia
        tkinter.Label(
            self.macro_campos, 
            text = "Estatus",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=5, column=1, sticky="w", padx=10)
        self.estilo_campos = ttk.Style(self.ventana)
        self.estilo_campos.theme_use("clam")
        self.estilo_campos.configure(
            "Facturas.TCombobox",
            fieldbackground = self.color_entrada,
            background = self.color_entrada,
            foreground = self.color_texto_claro,
            arrowcolor = self.color_texto_claro,
            bordercolor = self.color_entrada,
            lightcolor =  self.color_entrada,
            darkcolor = self.color_entrada,
            borderwidth = 0,
            padding = 2      
        )
        self.estilo_campos.map(
            "Facturas.TCombobox",
            fieldbackground = [("readonly", self.color_entrada)],
            background = [("readonly" , self.color_entrada)],
            foreground = [("readonly", self.color_texto_claro)],
            arrowcolor = [("readonly", self.color_texto_claro)]
        )
        self.entrada_estatus = ttk.Combobox(
            self.macro_campos,
            values = complemento.estado_factura(),
            state = "readonly",
            font = ("Segoe UI", 11),
            style = "Facturas.TCombobox"
        )
        # Colores de la lista desplegable de los estatus
        self.entrada_estatus.option_add(
            "*TCombobox*Listbox.background",
            self.color_entrada
        )
        self.entrada_estatus.option_add(
            "*TCombobox*Listbox.foreground",
            self.color_texto_claro
        )
        self.entrada_estatus.option_add(
            "*TCombobox*Listbox.selectBackground",
            "#2563EB"
        )
        self.entrada_estatus.option_add(
            "*TCombobox*Listbox.selectForeground",
            "#FFFFFF"
        )
        self.entrada_estatus.grid(row=6, column=1, sticky="ew", padx=10, pady=(5,12))
        # "Concepto / Fila 4"
        tkinter.Label(
            self.macro_campos, 
            text = "Concepto",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=7, column=0, sticky = "w", padx=10)
        self.entrada_concepto = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_concepto.grid(row=8, column=0, sticky = "ew", padx=10, pady=(5,12), ipady=5)
        # "Monto Facturado / Fila 4"
        tkinter.Label(
            self.macro_campos, 
            text = "Monto Facturado",
            font = ("Segoe UI", 10),
            bg = self.color_formulario,
            fg = self.color_etiquetas
            ).grid(row=7, column=1, sticky = "w", padx=10)
        self.entrada_monto_fact = tkinter.Entry(
            self.macro_campos, 
            font = ("Segoe UI", 11),
            bg = self.color_entrada,
            fg = self.color_texto_claro,
            insertbackground = self.color_texto_claro,
            relief = "flat",
            bd = 0
        )
        self.entrada_monto_fact.grid(row=8, column=1, sticky = "ew", padx=10, pady=(5, 12), ipady=5)
        
    def botones(self):
        # Aqui se guardan los botones en un Frame para usarlos en **init**
        # Frame — Crea un contenedor para organizar otros elementos
        self.macro_botones = tkinter.Frame(self.ventana, bg = self.color_fondo)
        self.macro_botones.columnconfigure(0, weight = 1)
        self.macro_botones.columnconfigure(1, weight = 1)
        self.macro_botones.columnconfigure(2, weight = 1)
        self.macro_botones.pack(fill = "x", padx = 20, pady = (15, 20))
        # BOTON AGREGAR
        self.boton_agregar = tkinter.Button(
            self.macro_botones,
            text = "AGREGAR",
            command = self.agregar_registro,
            font = ("Segoe UI", 11, "bold"),
            bg = "#2563EB",
            fg = "#FFFFFF",
            activebackground = "#1D4ED8",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"
        )
        self.boton_agregar.grid(row = 0, column = 0, sticky = "ew", padx = 5)
        # BOTON ACTUALIZAR
        self.boton_actualizar = tkinter.Button(
            self.macro_botones,
            text = "ACTUALIZAR",
            command = self.actualizar_registro,
            font = ("Segoe UI", 11, "bold"),
            bg = "#475569",
            fg = "#FFFFFF",
            activebackground = "#334155",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"  
            )
        self.boton_actualizar.grid(row = 0, column = 1, sticky = "ew", padx = 5)
        # BOTON BUSCAR
        self.boton_buscar = tkinter.Button(
            self.macro_botones,
            text = "BUSCAR",
            command = self.buscar_campos,
            font = ("Segoe UI", 11, "bold"),
            bg = "#0F766E",
            fg = "#FFFFFF",
            activebackground = "#115E59",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"
        )
        self.boton_buscar.grid(row = 0, column = 2, sticky = "ew", padx = 5)

    def crear_menu_superior(self):
        self.barra_superior = tkinter.Frame(
            self.ventana,
            bg = "#0F172A",
            height = 42
        )
        self.barra_superior.pack(fill = "x", side = "top")
        self.barra_superior.pack_propagate(False)
        self.boton_manual = tkinter.Button(
            self.barra_superior,
            text = "MANUAL DE USUARIO",
            command = lambda: ventana_maual_usuario(self.ventana), 
            font = ("Segoe UI", 9, "bold"),
            bg = "#1E3A5F",
            fg = "#FFFFFF",
            activebackground = "#2563EB",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            padx = 14,
            cursor = "hand2" 
        )
        self.boton_manual.pack(
            side = "left",
            padx = 12,
            pady = 5
        )
        self.boton_manual.bind(
            "<Enter>",
            lambda evento: self.boton_manual.config(bg = "#2563EB")
        )
        self.boton_manual.bind(
            "<Leave>",
            lambda evento: self.boton_manual.config(bg = "#1E3A5F")
        )

    def agregar_registro(self):
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        cotizacion = self.entrada_cotizacion.get().strip()
        orden_cotizacion = self.entrada_ord_cot.get().strip()
        cliente = self.entrada_cliente.get().strip()
        folio_factura = self.entrada_fol_fact.get().strip()
        concepto = self.entrada_concepto.get().strip()
        monto_facturado = self.entrada_monto_fact.get().strip()
        estatus = self.entrada_estatus.get().strip()
        campos_obligatorios = {
            "Numero de Cotizacion" : cotizacion,
            "Numero Orden de Compra" : orden_cotizacion,
            "Cliente" : cliente,
            "Folio Factura" : folio_factura,
            "Monto Facturado" : monto_facturado,
            "Estatus": estatus,
        }
        campos_vacios = []
        for nombre, valor in campos_obligatorios.items():
            if not valor:
                campos_vacios.append(nombre)
        if campos_vacios:
            complemento.mensajes.campos_obligatorios_vacios(campos_vacios)
            return
        datos = [
            fecha_actual,
            cotizacion,
            orden_cotizacion,
            cliente,
            folio_factura,
            concepto,
            monto_facturado,
            estatus,
            "",
            ""
        ]
        folio_repetido = self.excel.folio_ya_registrado(folio_factura)
        if folio_repetido:
            guardado = self.excel.registrar_folio_repetido(datos)
        else:
            guardado = self.excel.agregar_registro(datos)
        if guardado:
            datos_factura = {
                    "fecha": fecha_actual,
                    "cotizacion": cotizacion,
                    "orden_cotizacion" : orden_cotizacion,
                    "cliente": cliente,
                    "folio_factura": folio_factura,
                    "concepto": concepto,
                    "monto_facturado": monto_facturado,
                    "estatus": estatus
                    }
            self.limpiar_campos()
            respuesta = complemento.mensajes.agregar_complemento()
            if respuesta:
                try:
                    self.ventana_complemento = ventana_complementos(
                        self.ventana,
                        datos_factura
                    )
                except Exception as error:
                    import traceback
                    traceback.print_exc()
                    complemento.mensajes.mensaje_error(error)
                    
    def buscar_campos(self):
        self.datos_busqueda = {
            "cotizacion" : self.entrada_cotizacion.get(),
            "orden_cotizacion" : self.entrada_ord_cot.get(),
            "cliente" : self.entrada_cliente.get(),
            "folio_factura" : self.entrada_fol_fact.get(),
            "concepto" : self.entrada_concepto.get(),
            "monto_facturado" : self.entrada_monto_fact.get(),
            "estatus" : self.entrada_estatus.get(),        }       
        datos_busqueda = complemento.busqueda_valores(self.datos_busqueda)
        if not any(datos_busqueda.values()):
            complemento.mensajes.campos_busqueda_vacios()
            return
        resultados = self.excel.buscar_registros(datos_busqueda)
        if self.ventana_resultados is None:
            self.ventana_resultados = ventana_busqueda(
                self.ventana,
                self,
                resultados
                )
        else:
            self.ventana_resultados.mostrar_resultados(resultados)            

    def actualizar_registro(self):
        if (
            self.archivo_seleccionado is None
            or self.hoja_seleccionada is None
            or self.fila_seleccionada is None
            ):
            complemento.mensajes.registro_no_seleccionado()
            return
        confirmacion = complemento.mensajes.confirmar_actualizacion(self.fila_seleccionada)
        if not confirmacion:
            return
        datos = {
            "fecha" : self.entrada_fecha.get(),
            "cotizacion" : self.entrada_cotizacion.get(),
            "orden_cotizacion" : self.entrada_ord_cot.get(),
            "cliente" : self.entrada_cliente.get(),
            "folio_factura" : self.entrada_fol_fact.get(),
            "concepto" : self.entrada_concepto.get(),
            "monto_facturado" : self.entrada_monto_fact.get(),
            "estatus" : self.entrada_estatus.get(),
        }
        actualizado = self.excel.actualizar_registro(
            self.archivo_seleccionado,
            self.hoja_seleccionada,
            self.fila_seleccionada,
            datos
            )
        if actualizado:
            self.limpiar_campos()
            if (self.ventana_resultados is not None and self.ventana_resultados.ventana_busqueda.winfo_exists()):
                resultados = self.excel.buscar_registros(
                    complemento.busqueda_valores(self.datos_busqueda)
                )
                self.ventana_resultados.mostrar_resultados(resultados)

    def mostrar_fecha(self, fecha):
        self.entrada_fecha.config(state = "normal")
        self.entrada_fecha.delete(0, tkinter.END)
        self.entrada_fecha.insert(0, fecha)
        self.entrada_fecha.config(state = "readonly")

    def limpiar_campos(self):
        self.entrada_cotizacion.delete(0, tkinter.END)
        self.entrada_ord_cot.delete(0, tkinter.END)
        self.entrada_cliente.delete(0, tkinter.END)
        self.entrada_fol_fact.delete(0, tkinter.END)
        self.entrada_concepto.delete(0, tkinter.END)
        self.entrada_monto_fact.delete(0, tkinter.END)
        self.entrada_estatus.set("")
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        self.mostrar_fecha(fecha_actual)
        self.archivo_seleccionado = None
        self.hoja_seleccionada = None
        self.fila_seleccionada = None

    def revisar_carpeta_facturas(self):
        self.excel.organizar_archivos_facturas()
        self.ventana.after(3000, self.revisar_carpeta_facturas)

    def cargar_datos_iniciales(self):
        self.excel.cargar_folios_memoria()

class ventana_busqueda:
    def __init__(self, ventana_padre, ventana_principal, resultados):
        # Toma la ventana principal
        self.ventana_principal = ventana_principal
        self.ventana_info_factura = None
        # Colores de interfaz
        self.color_fondo = "#0F172A"
        self.color_primario = "#1E3A5F"
        self.color_tabla = "#1E293B"
        self.color_entrada ="#334155"
        self.color_texto = "#FFFFFF"
        self.color_secundario = "#94A3B8"
        # Se crea la ventana secundaria para buscar algun registro
        self.ventana_busqueda = tkinter.Toplevel(ventana_padre)
        self.ventana_busqueda.title("Ventana de Busqueda")
        self.ventana_busqueda.geometry("1200x650")
        self.ventana_busqueda.minsize(950,500)
        self.ventana_busqueda.iconbitmap(ruta_recurso("iconos/icono_busqueda.ico"))
        self.ventana_busqueda.configure(bg = self.color_fondo)
        self.ventana_busqueda.protocol("WM_DELETE_WINDOW", self.cerrar_ventana)
        # Ecabezado
        self.encabezado = tkinter.Frame(
            self.ventana_busqueda,
            bg = self.color_primario,
            height = 85
        )
        self.encabezado.pack(
            fill = "x", padx = 20, pady = (20,10)
        )
        self.encabezado.pack_propagate(False)
        tkinter.Label(
            self.encabezado,
            text = "RESULTADOS DE BUSQUEDA",
            font = ("Segoe UI", 17, "bold"),
            bg = self.color_primario,
            fg = self.color_texto
        ).pack(pady = (12, 2))
        tkinter.Label(
            self.encabezado,
            text = "Consulta y administracion de registros de facturas",
            font = ("Segoe UI", 10),
            bg = self.color_primario,
            fg = "#CBD5E1"
        ).pack()
        self.crear_tabla()
        self.configurar_columnas()
        self.colores_fila()
        self.botones_busqueda()
        self.mostrar_resultados(resultados)
        
    def crear_tabla(self):
        # Oscuro de la tabla
        self.estilo_tabla = ttk.Style(self.ventana_busqueda)
        self.estilo_tabla.theme_use("clam")
        self.estilo_tabla.configure(
            "Facturas.Treeview",
            background = "#1E293B",
            foreground = "#FFFFFF",
            fieldbackground = "#1E293B",
            borderwidth = 0,
            relief = "flat",
            rowheight = 30,
            font = ("Segoe UI", 10)
        )
        # Color de encabezados de columnas
        self.estilo_tabla.configure(
            "Facturas.Treeview.Heading",
            background = "#1E3A5F",
            foreground = "#FFFFFF",
            relief = "flat",
            font = ("Segoe UI", 10, "bold"),
            padding = (8,10)
        )
        # Color de seleccion de fila
        self.estilo_tabla.map(
            "Facturas.Treeview",
            background = [("selected", "#2563EB")],
            foreground = [("selected", "#FFFFFF")]
        )
        self.estilo_tabla.map(
            "Facturas.Treeview.Heading",
            background = [("active", "#274C77"), ("pressed" ,"#1E3A5F")],
            foreground = [("active", "#FFFFFF")]
        )
        # Se nombran las columnas de la tabla de busqueda
        columnas = complemento.columnas_busqueda()
        # Frame para contener la tabla y el scroll
        frame_tabla = tkinter.Frame(self.ventana_busqueda, bg = self.color_tabla)
        frame_tabla.pack(
            fill = "both",
            expand = True,
            padx = 10,
            pady = 10
        )
        # Se crea el scroll 'X' e 'Y'
        scroll_x = ttk.Scrollbar(
            frame_tabla,
            orient = "horizontal",
            style = "Facturas.Horizontal.TScrollbar"
        )
        scroll_x.pack(
            side = "bottom",
            fill = "x"
        )
        scroll_y = ttk.Scrollbar(
            frame_tabla,
            orient = "vertical",
            style = "Facturas.Vertical.TScrollbar"
        )
        scroll_y.pack(
            side = "right",
            fill = "y"
        )
        # Se crea la tabla de resultados
        self.tabla_resultados = ttk.Treeview(
            frame_tabla,
            columns = columnas,
            show = "headings",
            style = "Facturas.Treeview",
            xscrollcommand = scroll_x.set,
            yscrollcommand = scroll_y.set 
        )
        self.tabla_resultados.pack(
            fill = "both",
            expand = True
        )
        self.estilo_tabla.configure(
            "Facturas.Vertical.TScrollbar",
            background="#64748B",
            troughcolor="#0F172A",
            bordercolor="#0F172A",
            arrowcolor="#FFFFFF",
            lightcolor="#64748B",
            darkcolor="#64748B",
            relief="flat",
            borderwidth=0,
            width=14
        )
        self.estilo_tabla.configure(
            "Facturas.Horizontal.TScrollbar",
            background="#64748B",
            troughcolor="#0F172A",
            bordercolor="#0F172A",
            arrowcolor="#FFFFFF",
            lightcolor="#64748B",
            darkcolor="#64748B",
            relief="flat",
            borderwidth=0,
            width=14
        )
        for estilo in ("Facturas.Vertical.TScrollbar", "Facturas.Horizontal.TScrollbar"):
            self.estilo_tabla.map(
                estilo,
                background = [
                    ("pressed", "#94A3BB"),
                    ("active", "#788CA5")
                ]
            )
        # Se conectan los scrolls
        scroll_x.config(
            command = self.tabla_resultados.xview
        )
        scroll_y.config(
            command = self.tabla_resultados.yview    
        )

    def configurar_columnas(self):
        columnas = complemento.columnas_busqueda()
        # Crea los encabezados de cada columna
        for columna in columnas:
            self.tabla_resultados.heading(
                columna,
                text = columna
            )
        # Ajusta el ancho de cada columna visualizar mejor el
        # contenido
        self.tabla_resultados.column(
            "Archivo",
            width = 150,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Hoja",
            width = 100,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Fila",
            width = 50,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Fecha",
            width = 120,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Numero de Cotizacion",
            width = 150,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Numero Orden de Compra",
            width = 190,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Cliente",
            width = 180,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Folio Factura",
            width = 130,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Concepto",
            width = 280,
        )
        self.tabla_resultados.column(
            "Monto Facturado",
            width = 130,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Estatus",
            width = 110,
            anchor = "center"
        )
        self.tabla_resultados.column(
            "Complemento 1",
            width = 130,
        )
        self.tabla_resultados.column(
            "Complemento 2",
            width = 130,
        )

    def botones_busqueda(self):
        self.macro_botones_busqueda = tkinter.Frame(
            self.ventana_busqueda,
            bg = self.color_fondo
            )
        self.macro_botones_busqueda.pack(
            fill = "x",
            padx = 20,
            pady = (10, 20)
            )
        for columna in range(3):
            self.macro_botones_busqueda.columnconfigure(columna, weight = 1)
        # BOTON SELECCIONAR
        boton_seleccionar = tkinter.Button(
            self.macro_botones_busqueda,
            text="SELECCIONAR",
            command=self.seleccionar_registro,
            font = ("Segoe UI", 11, "bold"),
            bg = "#2563EB",
            fg = "#FFFFFF",
            activebackground = "#1D4ED8",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"
        )
        boton_seleccionar.grid(
            row = 0,
            column = 0,
            sticky = "ew",
            padx = 5
        )
        # BOTON ABRIR EN EXCEL
        boton_abrir_fila = tkinter.Button(
            self.macro_botones_busqueda,
            text="ABRIR EN EXCEL",
            command=self.abrir_fila_excel,
            font = ("Segoe UI", 11, "bold"),
            bg = "#475569",
            fg = "#FFFFFF",
            activebackground = "#334155",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"
        )
        boton_abrir_fila.grid(
            row = 0,
            column = 1,
            sticky = "ew",
            padx = 5
        )
        # BOTON ABRIR REGISTRO
        boton_abrir_factura = tkinter.Button(
            self.macro_botones_busqueda,
            text="ABRIR REGISTRO",
            command=self.abrir_registro_factura,
            font = ("Segoe UI", 11, "bold"),
            bg = "#0F766E",
            fg = "#FFFFFF",
            activebackground = "#115E59",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            pady = 12,
            cursor = "hand2"

        )
        boton_abrir_factura.grid(
            row = 0, 
            column = 2,
            sticky = "ew",
            padx = 5
        )

    def colores_fila(self):
        colores = complemento.colores_estatus()
        for estatus, (fondo, texto) in colores.items():
            self.tabla_resultados.tag_configure(
                estatus, 
                background = fondo,
                foreground = texto
                )

    def mostrar_resultados(self, resultados):
        for item in self.tabla_resultados.get_children():
            self.tabla_resultados.delete(item)
        for registro in resultados:
            estatus = str(registro[10]).strip().upper()
            self.tabla_resultados.insert(
                "",
                "end",
                values = registro,
                tags = (estatus,)
            )

    def seleccionar_registro(self):
        seleccion = self.tabla_resultados.selection()
        if not seleccion:
            complemento.mensajes.registro_no_seleccionado()
            return
        registro = self.tabla_resultados.item(seleccion[0])
        valores = registro["values"]
        # Guardar la fila seleccionada en la ventana principal
        self.ventana_principal.archivo_seleccionado = valores[0]
        self.ventana_principal.hoja_seleccionada = valores[1]
        self.ventana_principal.fila_seleccionada = valores[2]
        # Fecha
        self.ventana_principal.mostrar_fecha(valores[3])
        # Cotizacion
        self.ventana_principal.entrada_cotizacion.delete(0, tkinter.END)
        self.ventana_principal.entrada_cotizacion.insert(0, valores[4])
        # Orden de Cotizacion
        self.ventana_principal.entrada_ord_cot.delete(0, tkinter.END)
        self.ventana_principal.entrada_ord_cot.insert(0, valores[5])
        # Cliente
        self.ventana_principal.entrada_cliente.delete(0, tkinter.END)
        self.ventana_principal.entrada_cliente.insert(0, valores[6])
        # Folio Factura
        self.ventana_principal.entrada_fol_fact.delete(0, tkinter.END)
        self.ventana_principal.entrada_fol_fact.insert(0, valores[7])
        # Concepto
        self.ventana_principal.entrada_concepto.delete(0, tkinter.END)
        self.ventana_principal.entrada_concepto.insert(0, valores[8])
        # Monto Factura
        self.ventana_principal.entrada_monto_fact.delete(0, tkinter.END)
        self.ventana_principal.entrada_monto_fact.insert(0, valores[9])
        # Estatus
        self.ventana_principal.entrada_estatus.set(valores[10])
        # Complemento 1
        self.ventana_principal.entrada_complem1.delete(0, tkinter.END)
        self.ventana_principal.entrada_complem1.insert(0, valores[11])
        # Complementas 2
        self.ventana_principal.entrada_complem2.delete(0, tkinter.END)
        self.ventana_principal.entrada_complem2.insert(0, valores[12])

    def abrir_fila_excel(self): 
        try:
            seleccion = self.tabla_resultados.selection()
            if not seleccion:
                complemento.mensajes.registro_no_seleccionado()
                return
            registro = self.tabla_resultados.item(seleccion[0])
            valores = registro["values"]
            nombre_archivo = valores[0]
            nombre_hoja = valores[1]
            fila_excel = int(valores[2])
            ruta_excel = os.path.join(self.ventana_principal.excel.carpeta_excel, nombre_archivo) 
            libro = self.ventana_principal.excel.obtener_libro_abierto(ruta_excel)
            if libro is not None:
                excel = libro.Application
                excel.Visible = True
                libro.Activate()
            else:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = True
                libro = excel.Workbooks.Open(ruta_excel)
            hoja = libro.Worksheets(nombre_hoja)
            hoja.Activate()
            hoja.Range(f"A{fila_excel}:J{fila_excel}").Select()
        except Exception as error:
            self.ventana_principal.excel.logs.escribir(
                "ERROR", {
                    "Proceso": "Abrir fila en excel",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)

    def abrir_registro_factura(self):
        try:
            seleccion = self.tabla_resultados.selection()
            if not seleccion:
                complemento.mensajes.registro_no_seleccionado()
                return
            registro = self.tabla_resultados.item(seleccion[0])            
            valores = registro["values"]
            folio_factura = str(valores[7]).strip()
            archivos_encontrados = (self.ventana_principal.excel.organizar_archivos_facturas(folio_factura))
            if not archivos_encontrados:
                complemento.mensajes.archivos_factura_no_encontrados(folio_factura)
                return
            if (self.ventana_info_factura is None or not self.ventana_info_factura.ventana.winfo_exists()):
                self.ventana_info_factura = ventana_info_facturas(
                    self.ventana_busqueda, 
                    valores, 
                    archivos_encontrados, 
                    self.ventana_principal.excel.carpeta_facturas,
                    self.ventana_principal.excel.logs
                    )
            else:
                self.ventana_info_factura.actualizar_datos(
                    valores,
                    archivos_encontrados
                )
                self.ventana_info_factura.ventana.lift()
                self.ventana_info_factura.ventana.focus_force()
        except Exception as error:
            self.ventana_principal.excel.logs.escribir(
                "ERROR", {
                    "Proceso": "Abrir registro de factura",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_al_abrir_factura(error)

    def cerrar_ventana(self):
        self.ventana_principal.ventana_resultados = None
        self.ventana_busqueda.destroy()

class ventana_info_facturas:
    def __init__(self, ventana_padre, datos_registro, archivos, carpetas_facturas, logs):
        # Datos del registro
        self.datos_registro = datos_registro
        self.archivos = archivos
        self.carpetas_facturas = carpetas_facturas
        self.logs = logs
        self.folio_factura = str(datos_registro[7]).strip()
        self.ruta_incorporaciones = os.path.join(
            os.path.dirname(self.carpetas_facturas),
            "incorporaciones_facturas.json"
        ) 
        # Colores interfaz
        self.color_fondo = "#0F172A"
        self.color_primario = "#1E3A5F"
        self.color_panel = "#1E293B"
        self.color_entrada = "#334155"
        self.color_texto = "#FFFFFF"
        self.color_secundario = "#94A3B8"
        # Ventana de info facts
        self.ventana = tkinter.Toplevel(ventana_padre)
        self.ventana.title("Ventana del Registro")
        self.ventana.geometry("850x750")
        self.ventana.minsize(800, 700)
        self.ventana.iconbitmap(ruta_recurso("iconos/icono_facturas.ico"))
        self.ventana.configure(bg = self.color_fondo)
        # Encabezado
        self.encabezado = tkinter.Frame(
            self.ventana,
            bg = self.color_primario,
            height = 85
        )
        self.encabezado.pack(
            fill = "x",
            padx = 20,
            pady = (20,10)
        )
        self.encabezado.pack_propagate(False)
        tkinter.Label(
            self.encabezado,
            text = "INFORMACION DE FACTURAS",
            font = ("Segoe UI", 17, "bold"),
            bg = self.color_primario,
            fg = self.color_texto
        ).pack(pady = (12, 2))
        tkinter.Label(
            self.encabezado,
            text = "Consulta de documentos y datos del registro",
            font = ("Segoe UI", 10),
            bg = self.color_primario,
            fg = "#CBD5E1"
        ).pack()
        self.crear_botones()
        self.mostrar_datos_registro()
        self.crear_tabla_archivos()
        self.mostrar_archivos()

    def crear_tabla_archivos(self):
        self.frame_archivos = tkinter.Frame(
            self.ventana,
            bg = self.color_panel,
            padx = 25,
            pady = 12
        )
        self.frame_archivos.pack(
            fill = "x",
            padx = 20,
            pady = (5,10)
        )
        self.label_archivos = tkinter.Label(
            self.frame_archivos,
            text = f"Archivos encontrados del Folio: {self.folio_factura}",
            font = ("Segoe UI", 11,"bold"),
            bg = self.color_panel,
            fg = self.color_texto,
            anchor = "w"
            )
        self.label_archivos.pack(fill = "x", pady = (0,10))
        estilo = ttk.Style(self.ventana)
        estilo.theme_use("clam")
        estilo.configure(
            "Archivos.Vertical.TScrollbar",
            background = "#64748B",
            troughcolor = "#0F172A",
            bordercolor = "#0F172A",
            arrowcolor = "#FFFFFF",
            lightcolor = "#64748B",
            darkcolor = "#64748B",
            relief = "flat",
            borderwidth = 0,
            width = 14
        )
        estilo.map(
            "Archivos.Vertical.TScrollbar",
            background = [
                ("pressed", "#94A3B8"),
                ("active", "#788CA5")
            ]
            )
        estilo.configure(
            "Archivos.Treeview",
            background = self.color_panel,
            foreground = self.color_texto,
            fieldbackground = self.color_panel,
            rowheight = 30,
            font = ("Segoe UI", 10),
            borderwidth = 0
        )
        estilo.configure(
            "Archivos.Treeview.Heading",
            background = self.color_primario,
            foreground = self.color_texto,
            font = ("Segoe UI", 10, "bold"),
            padding = (8, 10)
        )
        estilo.map(
            "Archivos.Treeview",
            background = [("selected", "#2563EB")],
            foreground = [("selected", "#FFFFFF")]
        )
        frame_contenido = tkinter.Frame(
            self.frame_archivos,
            bg = self.color_panel
        )
        frame_contenido.pack(
            fill = "both",
            expand = True
        )
        scroll_y = ttk.Scrollbar(
            frame_contenido,
            orient = "vertical",
            style = "Archivos.Vertical.TScrollbar"
        )
        scroll_y.pack(
            side = "right",
            fill = "y"
        )
        self.tabla_archivos = ttk.Treeview(
            frame_contenido,
            columns = (
                "Folio",
                "Creacion",
                "Incorporacion",
                "Archivo"
            ),
            show = "headings",
            selectmode = "browse",
            style = "Archivos.Treeview",
            height = 7,
            yscrollcommand = scroll_y.set
            )
        columnas = complemento.columnas_info_facturas()
        for columna, (titulo, ancho) in columnas.items():
            self.tabla_archivos.heading(
                columna,
                text = titulo
            )
            self.tabla_archivos.column(
                columna,
                width = ancho,
                minwidth = 80,
                anchor = "center",
                stretch = True
            )
        self.tabla_archivos.pack(
            side = "left",
            fill = "both",
            expand = True
        )
        scroll_y.config(
            command = self.tabla_archivos.yview
        )
    
    def mostrar_archivos(self):
        incorporaciones = {}
        if os.path.isfile(self.ruta_incorporaciones):
            try:
                with open(
                    self.ruta_incorporaciones,
                    "r",
                    encoding = "utf-8"
                ) as archivo_json:
                    incorporaciones = json.load(archivo_json)
            except (OSError, ValueError, json. JSONDecodeError) as error:
                self.logs.escribir(
                    "ERROR", {
                        "Proceso": "Leer incorporaciones",
                        "Tipo": type(error).__name__,
                        "Error": str(error)
                    }
                )
        for archivo in self.archivos:
            ruta_archivo  = os.path.join(
                self.carpetas_facturas,
                self.folio_factura,
                archivo
            )
            fecha_creacion = "No disponible"
            if os.path.isfile(ruta_archivo):
                fecha_creacion = datetime.fromtimestamp(
                    os.path.getctime(ruta_archivo)
                ).strftime("%d/%m/%Y %H:%M")
            clave = f"{self.folio_factura}/{archivo}"
            fecha_incorporacion = incorporaciones.get(clave, "Sin registro")
            self.tabla_archivos.insert(
                "",
                "end",
                values = (
                    self.folio_factura,
                    fecha_creacion,
                    fecha_incorporacion,
                    archivo
                )
            )

    def crear_botones(self):
        frame_botones = tkinter.Frame(
            self.ventana,
            bg = self.color_fondo
        )
        frame_botones.pack(
            side = "bottom",
            fill =  "x",
            padx = 20,
            pady = (10,20)
            )
        frame_botones.columnconfigure(0, weight = 1)
        frame_botones.columnconfigure(1, weight = 1)
        # BOTON ABRIR FACTURA
        boton_abrir_factura = tkinter.Button(
            frame_botones,
            text = "ABRIR FACTURA",
            command = self.abrir_factura,
            font = ("Segoe UI", 11, "bold"),
            bg = "#2563EB",
            fg = "#FFFFFF",
            activebackground = "#1D4ED8",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            cursor = "hand2",
            pady = 12
        )
        boton_abrir_factura.grid(
            row = 0,
            column = 0,
            sticky = "ew", 
            padx = (0,5)
        )
        # BOTON ABRIR CARPETA
        boton_abrir_carpeta = tkinter.Button(
            frame_botones,
            text = "ABRIR CARPETA",
            command = self.abrir_carpeta,
            font = ("Segoe UI", 11, "bold"),
            bg = "#0F766E",
            fg = "#FFFFFF",
            activebackground = "#115E59",
            activeforeground = "#FFFFFF",
            relief = "flat",
            bd = 0,
            cursor = "hand2",
            pady = 12
        )
        boton_abrir_carpeta.grid(
            row = 0,
            column = 1,
            sticky = "ew", 
            padx = (5,0)
        )

    def abrir_factura(self):
        try:
            seleccion = self.tabla_archivos.selection()
            if not seleccion:
                complemento.mensajes.factura_no_seleccionada()
                return
            datos = self.tabla_archivos.item(seleccion[0])
            archivo = datos["values"][3]
            carpeta_folio = os.path.join(self.carpetas_facturas, self.folio_factura)
            ruta_archivo = os.path.join(carpeta_folio, archivo)
            if not os.path.isfile(ruta_archivo):
                complemento.mensajes.archivos_factura_no_encontrados(self.folio_factura)
                return
            confirmacion = complemento.mensajes.confirmar_abrir_factura(archivo)
            if not confirmacion:
                return
            subprocess.Popen(
                ["cmd", "/c", "start", "", ruta_archivo],
                shell = False
            )
        except Exception as error:
            traceback.print_exc()
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Abrir factura",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_al_abrir_factura(error)

    def abrir_carpeta(self):
        try:    
            carpeta_folio = os.path.join(self.carpetas_facturas, self.folio_factura)
            if not os.path.isdir(carpeta_folio):
                complemento.mensajes.archivos_factura_no_encontrados(self.folio_factura)
                return
            subprocess.Popen(["explorer.exe", carpeta_folio])
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Abrir carpeta de factura",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_al_abrir_carpeta(error)

    def mostrar_datos_registro(self):
        self.frame_datos = tkinter.Frame(
            self.ventana, 
            bg = self.color_panel,
            padx = 20,
            pady = 10
        )
        self.frame_datos.pack(
            fill = "x",
            padx = 20,
            pady = 10
        )
        self.frame_datos.columnconfigure(0, weight = 1)
        self.frame_datos.columnconfigure(1, weight = 1)
        tkinter.Label(
            self.frame_datos,
            text = "INFORMACION DEL REGISTRO",
            font = ("Segoe UI", 11, "bold"),
            bg = self.color_panel,
            fg = self.color_texto
        ).grid(
            row = 0,
            column = 0,
            columnspan = 2,
            sticky = "w",
            pady = (0,8)
        ) 
        nombres = complemento.nombres_columnas()
        valores = self.datos_registro[3:13]
        self.labels_datos = []
        for indice, (nombre, valor) in enumerate(zip(nombres, valores)):
            fila = (indice // 2) + 1
            columna = indice % 2
            contenedor = tkinter.Frame(
                self.frame_datos,
                bg = self.color_panel
            )
            contenedor.grid(
                row = fila,
                column = columna,
                sticky = "ew",
                padx = "10",
                pady = 3
            )
            tkinter.Label(
                contenedor,
                text = nombre,
                font = ("Segoe UI", 9),
                bg = self.color_panel,
                fg = self.color_secundario
            ).pack(anchor = "w")
            label_valor = tkinter.Label(
                contenedor,
                text = str(valor),
                font = ("Segoe UI", 10, "bold"),
                bg = self.color_panel,
                fg = self.color_texto,
                anchor = "w",
                justify = "left",
                wraplength = 280
            )
            label_valor.pack(
                anchor = "w",
                pady = (1,0)
            )
            self.labels_datos.append(label_valor)

    def actualizar_datos(self, datos_registro, archivo):
        self.datos_registro = datos_registro
        self.archivos = archivo
        self.folio_factura = str(datos_registro[7]).strip()
        self.ventana.title(
            f"Datos del Folio: {self.folio_factura}"
        )
        self.label_archivos.config(
            text = f"Archivos encontrados del Folio: {self.folio_factura}"
        )
        valores = datos_registro[3:13]
        for label, valor in zip(self.labels_datos, valores):
            label.config(text = str(valor))
        for item in self.tabla_archivos.get_children():
            self.tabla_archivos.delete(item)
        self.mostrar_archivos()

class ventana_maual_usuario:
    def __init__(self, ventana_padre):
        self.ventana = tkinter.Toplevel(ventana_padre)
        self.ventana.title("Ventana Manual de Usuario")
        self.ventana.geometry("900x650")
        self.ventana.minsize(750,500)
        self.ventana.configure(bg = "#10243E")
        self.ventana.iconbitmap(ruta_recurso("iconos/icono_manual.ico"))
        self.ventana.transient(ventana_padre)
        self.encabezado = tkinter.Frame(
            self.ventana,
            bg = "#173B62"
        )
        self.encabezado.pack(fill = "x")
        self.titulo = tkinter.Label(
            self.encabezado,
            text = "MANUAL DE USUARIO",
            font = ("Arial", 20, "bold"),
            bg = "#173B62",
            fg = "#FFFFFF"
        )
        self.titulo.pack(pady = (20, 5))
        self.subtitulo = tkinter.Label(
            self.encabezado,
            text = F"Guia de Uso | {NOMBRE_PROGRAMA} v{VERSION_PROGRAMA}",
            font = ("Arial", 11),
            bg = "#173B62",
            fg = "#94A3B8"
        )
        self.subtitulo.pack(pady = (0,20))
        self.contenedor_manual = tkinter.Frame(
            self.ventana,
            bg = "#10243E"
        )
        self.contenedor_manual.pack(
            fill="both",
            expand=True,
            padx=20, 
            pady=10
        )
        self.area_contenido = tkinter.Frame(
            self.contenedor_manual,
            bg = "#142B46",
            width = 600
        )
        self.area_contenido.pack(
            side="left",
            fill="both",
            expand=True
        )
        self.area_contenido.pack_propagate(False)
        self.crear_menu_secciones()

    def crear_menu_secciones(self):
        secciones = complemento.secciones_manual()
        self.botones_secciones = {}
        self.contenedor_menu = tkinter.Frame(
            self.contenedor_manual,
            bg="#10243E",
            width=270
        )
        self.contenedor_menu.pack(
            side="right",
            fill="y",
            padx=(10,0)
        )
        self.contenedor_menu.pack_propagate(False)
        self.lienzo_menu = tkinter.Canvas(
            self.contenedor_menu,
            bg = "#10243E",
            highlightthickness = 0,
            width = 250,
            height = 500
        )
        self.lienzo_menu.pack(
            side = "left",
            fill = "both",
            expand = True
        )
        self.scroll_menu = tkinter.Scrollbar(
            self.contenedor_menu,
            orient = "vertical",
            command = self.lienzo_menu.yview
        )
        self.scroll_menu.pack(
            side = "right",
            fill = "y"
        )
        self.lienzo_menu.configure(
            yscrollcommand = self.scroll_menu.set
        )
        self.menu_secciones = tkinter.Frame(
            self.lienzo_menu, 
            bg = "#10243E",
            width = 230
        )
        self.ventana_menu = self.lienzo_menu.create_window(
            (0,0),
            window = self.menu_secciones,
            anchor = "nw"
        )
        self.lienzo_menu.bind(
            "<MouseWheel>",
            lambda evento: self.lienzo_menu.yview_scroll(
                int(-1 * (evento.delta/120)),
                "units"
            )
        )
        self.menu_secciones.bind(
            "<MouseWheel>",
            lambda evento: self.lienzo_menu.yview_scroll(
                int(-1 * (evento.delta/120)),
                "units"
            )
        )
        self.menu_secciones.bind(
            "<Configure>",
            lambda evento: self.actualizar_scroll_menu
        )
        self.lienzo_menu.bind(
            "<Configure>",
            lambda evento: self.lienzo_menu.itemconfigure(
                self.ventana_menu,
                width=evento.width
            )
        )
        self.lienzo_menu.bind(
            "<MouseWheel>",
            lambda evento: self.lienzo_menu.yview_scroll(
                -1 * int(evento.delta / 120)),
                "units"
            )  
        for fila, seccion in enumerate(secciones):
            boton = tkinter.Button(
                self.menu_secciones,
                text = f" {seccion}",
                font = ("Segoe UI", 9),
                bg = "#1E3A5F",
                fg = "#FFFFFF",
                activebackground = "#2563EB",
                activeforeground = "#FFFFFF",
                relief = "flat",
                anchor = "w",
                padx = 12,
                pady = 10,
                cursor = "hand2",
                command = lambda nombre = seccion: self.mostrar_contenido(nombre)
            )
            boton.grid(
                row = fila, 
                column = 0, 
                sticky = "ew", 
                padx = 4, 
                pady = 4
            )
            self.botones_secciones[seccion] = boton
            boton.bind(
                "<MouseWheel>",
                lambda evento: self.lienzo_menu.yview_scroll(
                    int(-1 * (evento.delta/120)),
                    "units"
                )
            )
        self.menu_secciones.columnconfigure(0, weight=1)
        self.lienzo_menu.update_idletasks()
        self.lienzo_menu.configure(scrollregion = self.lienzo_menu.bbox("all"))
        self.lienzo_menu.yview_moveto(0)
        if secciones:
            self.mostrar_contenido(secciones[0])

    def mostrar_contenido(self, seccion = "Introduccion"):
        for nombre, boton in self.botones_secciones.items():
            if nombre == seccion:
                boton.config(
                    text = f"▸  {nombre}",
                    bg = "#245C91",
                    fg = "#FFFFFF",
                    font = ("Segoe UI", 10, "bold"),
                    relief = "solid",
                    bd = 1
                )
            else:
                boton.config(
                    text = f"   {nombre}",
                    bg = "#1E3A5F",
                    fg = "#FFFFFF",
                    font = ("Segoe UI", 10),
                    relief = "flat",
                    bd = 0
                )
        for widget in self.area_contenido.winfo_children():
            widget.destroy()
        tkinter.Label(
            self.area_contenido,
            text = seccion.upper(),
            font = ("Arial", 17, "bold"),
            bg = "#142B46",
            fg ="#FFFFFF",
        ).pack(anchor = "w", padx = 20, pady = 20)
        contenidos = self.obtener_contenido_manual()
        texto = contenidos.get(
            seccion,
            "No hay informacion disponible para esta seeccion."
        )
        frame_texto = tkinter.Frame(
            self.area_contenido,
            bg = "#142B46"
        )
        frame_texto.pack(
            fill = "both",
            expand = True,
            padx = 20,
            pady = (0,20)
        )
        scroll = tkinter.Scrollbar(frame_texto)
        scroll.pack(side = "right", fill = "y")
        self.texto_manual = tkinter.Text(
            frame_texto,
            font = ("Segoe UI", 11),
            bg = "#142B46",
            fg = "#E2E8F0",
            insertbackground = "#FFFFFF",
            relief = "flat",
            wrap = "word",
            padx = 12,
            pady = 12,
            yscrollcommand = scroll.set
        )
        self.texto_manual.pack(
            side = "left",
            fill = "both", 
            expand = True
        )
        scroll.config(command = self.texto_manual.yview)
        self.texto_manual.insert("1.0", texto)
        self.texto_manual.config(state = "disabled")

    def obtener_contenido_manual(self):
        return complemento.obtener_contenido_manual()

    def actualizar_scroll_menu(self):
        self.lienzo_menu.configure(
            scrollregion = self.lienzo_menu.bbox("all")
        )

class ventana_complementos:
    def __init__(self, ventana_padre, datos_factura):
        self.ventana_padre = ventana_padre
        self.datos_factura = datos_factura
        self.ventana = tkinter.Toplevel(ventana_padre)
        self.ventana.title("Sistema de Facturas | Complementos")
        self.ventana.geometry("760x650")
        self.ventana.resizable(False, False)
        self.ventana.configure(bg = "#111B27")
        self.crear_campos()
        self.crear_botones()

    def crear_campos(self):
        encabezado = tkinter.Frame(
            self.ventana,
            bg = "#203F63",
            padx = 15,
            pady = 15
        )
        encabezado.pack(fill = "x", padx = 20, pady = (20,15))
        tkinter.Label(
            encabezado,
            text = "REGISTRO DE COMPLEMENTO DE PAGO",
            font = ("Arial", 16, "bold"),
            fg = "#FFFFFF",
            bg = "#203F63"
        ).pack(pady = (0,5))
        tkinter.Label(
            encabezado,
            text = "Administracion y seguimiento de pagos",
            font = ("Arial", 10),
            fg = "#D6E3F3",
            bg = "#203F63"
        ).pack()
        formulario = tkinter.Frame(
            self.ventana,
            bg = "#1E293B",
            padx = 20,
            pady= 15
        )
        formulario.pack(fill="both", expand=True, padx=20, pady=5)

        def crear_entrada(texto, fila, columna, atributo):
            contenedor = tkinter.Frame(
                formulario,
                bg = "#1E2938"
            )
            contenedor.grid(
                row = fila,
                column = columna,
                padx = 10,
                pady = 9,
                sticky = "ew"
            )
            tkinter.Label(
                contenedor,
                text=texto,
                font=("Arial", 10, "bold"),
                fg = "#E5E7EB",
                bg="#1E293B",
                anchor="w"
            ).pack(fill="x",pady=(0,5))
            entrada = tkinter.Entry(
                contenedor,
                font=("Arial", 11),
                bg = "#344256",
                fg="#FFFFFF",
                insertbackground="#FFFFFF",
                relief="flat",
                width=28
            )
            entrada.pack(fill="x",ipady=7)
            setattr(self, atributo, entrada)
        formulario.columnconfigure(0,weight=1)
        formulario.columnconfigure(1,weight=1)
        crear_entrada(
            "Fecha emitida:",
            0,0,
            "entrada_fecha_emitida"
        )
        crear_entrada(
            "Folio de Factura",
            1,0,
            "entrada_folio_factura"
        )
        crear_entrada(
            "Monto Facturado",
            2,0,
            "entrada_monto_facturado"
        )
        crear_entrada(
            "Monto Pagado",
            3,0,
            "entrada_monto_pagado"
        )
        crear_entrada(
            "Estatus",
            4,0,
            "entrada_estatus"
        )
        crear_entrada(
            "Cliente",
            0,1,
            "entrada_cliente"
        )
        crear_entrada(
            "Complemento (referencia)",
            1,1,
            "entrada_complemento"
        )
        crear_entrada(
            "Saldo anterior",
            2,1,
            "entrada_saldo_anterior"
        )
        crear_entrada(
            "fecha de pago",
            3,1,
            "entrada_fecha_pago"
        )
        placeholders = [
            (self.entrada_fecha_emitida, "Ej. 22/10/2026"),
            (self.entrada_cliente, "Ej. Empresa ABC"),
            (self.entrada_folio_factura, "Ej. 0001-ABC"),
            (self.entrada_complemento, "Ej. CP-0001"),
            (self.entrada_monto_facturado, "Ej. 15000.55"),
            (self.entrada_saldo_anterior, "Ej. 15000.55"),
            (self.entrada_monto_pagado, "Ej. 5000.99"),
            (self.entrada_fecha_pago, "Ej. 22/10/2026"),
            (self.entrada_estatus, "Ej. Pagado"),
        ]
        for entrada, texto in placeholders:
            self.agregar_placeholder(entrada, texto)

    def crear_botones(self):
        macro_botones = tkinter.Frame(
            self.ventana,
            bg ="#111827"
        )
        macro_botones.pack(fill = "x", padx= 20, pady = 15)
        configuracion_botones = [
            ("GUARDAR", self.guardar_complemento, "#2563EB"),
            ("LIMPIAR", self. limpiar_campos, "#475569"),
            ("CERRAR", self.ventana.destroy, "#0F766E")
        ]
        for columna, (texto, comando, color) in enumerate(configuracion_botones):
            boton = tkinter.Button(
                macro_botones,
                text = texto,
                command = comando,
                font=("Arial", 11, "bold"),
                bg=color,
                fg="#FFFFFF",
                activebackground=color,
                activeforeground="#FFFFFF",
                relief="flat",
                cursor="hand2",
                height=2
            )
            boton.grid(
                row = 0, 
                column = columna,
                padx = 5,
                sticky="ew"
                )
            macro_botones.columnconfigure(
                columna,
                weight=1
            )
            if texto == "GUARDAR":
                self.boton_guardar = boton
            elif texto == "LIMPIAR":
                self.boton_limpiar = boton
            else:
                self.boton_cerrar = boton

    def guardar_complemento(self):
        mensaje_guardado = complemento.mensajes.guardar_complemento()

    def limpiar_campos(self):
        entradas = [
            self.entrada_fecha_emitida,
            self.entrada_cliente,
            self.entrada_folio_factura,
            self.entrada_complemento,
            self.entrada_monto_facturado,
            self.entrada_saldo_anterior,
            self.entrada_monto_pagado,
            self.entrada_fecha_pago,
            self.entrada_estatus
        ]
        for entrada in entradas:
            entrada.delete(0, tkinter.END)

    def agregar_placeholder(self, entrada, texto):
        color_normal = "#FFFFFF"
        color_ejemplo = "#AAB7C7"
        entrada.insert(0, texto)
        entrada.config(fg = color_ejemplo)
        def enfocado(event):
            if entrada.get() == texto:
                entrada.delete(0, tkinter.END)
                entrada.config(fg = color_normal)
        def desenfocado(event):
            if not entrada.get().strip():
                entrada.insert(0, texto)
                entrada.config(fg = color_ejemplo)
        entrada.bind("<FocusIn>", enfocado)
        entrada.bind("<FocusOut>", desenfocado)

class gestor_de_complementos:
    def __init__(self, configuracion, logs):
        self.excel = excel_datos()
        self.configuracion = self.excel.configuracion
        self.logs = self.excel.logs

    def obtener_o_crear_hoja_complementos(self, libro):
        nombre_hoja = "COMPLEMENTOS"   
        for hoja in libro.Worksheets:
            if hoja.Name == nombre_hoja:
                return hoja, False
        hoja = libro.Worksheets.Add(
            After = libro.Worksheets(libro.Worksheets.Count)
        )
        hoja.Name = nombre_hoja
        return hoja, True

    def preparar_hoja_complementos(self, hoja):
        encabezados = complemento.encabezados_complementos()
        for numero_columna, encabezado in enumerate(encabezados, start= 1):
            hoja.Cells(1, numero_columna).Value = encabezado
        rango_encabezado = hoja.Range("A1:I1")
        rango_encabezado.Interior.Color = self.excel.convertir_color_com("1F4E78")
        rango_encabezado.Font.Bold = True
        rango_encabezado.Font.Color = 16777215
        rango_encabezado.HorizontalAlignment = -4108
        rango_encabezado.VerticalAlignment = -4108
        rango_encabezado.Borders.LineStyle = 1
        rango_encabezado.Borders.Weight = 2
        hoja.Rows(1).RowHeight = 25
        rango_encabezado.AutoFilter()

    def preparar_hoja_complementos_openpyxl(self, hoja):
        encabezados = complemento.encabezados_complementos()
        for numero_columna, encabezado in enumerate(encabezados, start= 1):
            celda = hoja.cell(row = 1, column = numero_columna, value = encabezado)
            celda.font = Font(bold = True, color = "FFFFFF")
            celda.fill = PatternFill(fill_type = "solid", fgColor = "1F4E78")
            celda.alignment = Alignment(
                horizontal = "center",
                vertical = "center"
            )
            celda.border = complemento.borde_celdas()
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = "A1:I1"
        hoja.row_dimensions[1].height = 25

    def preparar_hoja_complementos_archivo(self):
        ruta_excel = self.configuracion.obtener_ruta_excel_actual()
        if not os.path.exists(ruta_excel):
            ruta_excel, _ = self.configuracion.preparar_excel_actual()
        libro_abierto = self.excel.obtener_libro_abierto(ruta_excel)
        if libro_abierto is not None:
            hoja, creada = self.obtener_o_crear_hoja_complementos(libro_abierto)
            if creada:
                self.preparar_hoja_complementos(hoja)
            libro_abierto.Save()
        else:
            libro = openpyxl.load_workbook(ruta_excel)
            try:
                if "COMPLEMENTOS" not in libro.sheetnames:
                    hoja = libro.create_sheet("COMPLEMENTOS")
                    self.preparar_hoja_complementos_openpyxl(hoja)
                libro.save(ruta_excel)
            finally:
                
                libro.close()
        return True

    def guardar_complemento(self, datos):
        self.preparar_hoja_complementos_archivo()
        ruta_excel = self.configuracion.obtener_ruta_excel_actual()
        libro_abierto = self.excel.obtener_libro_abierto(ruta_excel)
        if libro_abierto is not None: 
            hoja = libro_abierto.Worksheets("COMPLEMENTOS")
            siguiente_fila = hoja.Cells(
                hoja.Rows.Count,
                1
            ).End(-4162).Row + 1 
            for columna, valor in enumerate(datos.values(), start = 1):
                hoja.Cells(siguiente_fila, columna).Value = valor
            libro_abierto.Save()
        else:
            libro = openpyxl.load_workbook(ruta_excel)
            try:
                hoja = libro["COMPLEMENTOS"]
                siguiente_fila = self.obtener_siguiente_fila_openpyxl(hoja)
                for columna, valor in enumerate(datos.values(), start = 1):
                    hoja.cell(
                        row = siguiente_fila,
                        column = columna,
                        value = valor
                    )
                libro.save(ruta_excel)
            finally:
                libro.close()
        return True

    def obtener_siguiente_fila_openpyxl(self, hoja):
        for fila in range(hoja.max_row, 1, -1):
            if hoja.cell(row = fila, column = 1).value is not None:
                 return fila + 1
        return 2

class excel_datos: 
    def __init__(self):
        if getattr(sys, "frozen",False):
            carpeta_programa = os.path.dirname(sys.executable)
        else:
            carpeta_programa = os.path.dirname(os.path.abspath(__file__))
        self.carpeta_programa = carpeta_programa
        self.carpeta_excel = os.path.join(carpeta_programa, "archivos_excel")
        self.carpeta_facturas = os.path.join(carpeta_programa, "folios_facturas")
        self.ruta_incorporaciones = os.path.join(
            self.carpeta_programa,
            "incorporaciones_facturas.json"
        )
        self.logs = registros_logs(self.carpeta_programa)
        try:
            os.makedirs(self.carpeta_excel, exist_ok = True)
            os.makedirs(self.carpeta_facturas, exist_ok = True)
        except OSError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Inicializar carpetas del programa",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            raise
        self.configuracion = configuracion_excel(self.carpeta_excel, self.logs)
        self.folios_registrados = set()

    def agregar_registro(self, datos):
        libro = None
        fila_libre = None
        guardado = False
        libro_abierto = None
        try:
            ruta_excel = self.configuracion.obtener_ruta_excel_actual()
            nombre_hoja = self.configuracion.obtener_mes_actual()
            libro_abierto = self.obtener_libro_abierto(ruta_excel)
            if libro_abierto is not None: 
                hoja, hoja_creada = self.obtener_o_crear_hoja_com(libro_abierto, nombre_hoja)
                if hoja_creada:
                    self.preparar_hoja_com(hoja)
                fila_libre = self.buscar_fila_libre_com(hoja)
                for numero_columna, valor in enumerate(datos, start = 1):
                    if numero_columna == 1:
                        valor = datetime.strptime(valor, "%d/%m/%Y")
                        hoja.Cells(fila_libre, numero_columna).Value = valor      
                        hoja.Cells(fila_libre, numero_columna).NumberFormat = "dd/mm/aaaa"
                    else:
                        hoja.Cells(fila_libre,numero_columna).Value = valor         
                estatus = datos[7]
                self.preparar_fila_com(hoja, fila_libre, estatus)
                libro_abierto.Save()
                if hoja_creada:
                    self.logs.escribir(
                        "CREAR HOJA",
                        {
                            "Archivo": os.path.basename(ruta_excel),
                            "Hoja": nombre_hoja
                        }
                    )
                guardado = True
            else:
                ruta_excel, nombre_hoja = (self.configuracion.preparar_excel_actual())
                libro = openpyxl.load_workbook(ruta_excel)
                hoja = libro[nombre_hoja]
                fila_libre = self.buscar_fila_libre(hoja)
                for numero_columna, valor in enumerate(datos, start = 1):
                    hoja.cell(
                        row = fila_libre,
                        column = numero_columna,
                        value = valor
                    )
                estatus = datos[7]
                self.configuracion.preparar_fila(hoja, fila_libre, estatus)
                libro.save(ruta_excel)
                guardado = True
        except FileNotFoundError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Agregar registro",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                    }
                )
            complemento.mensajes.archivo_no_encontrado()
        except PermissionError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Agregar registro",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                    }
                )
            complemento.mensajes.archivo_en_uso()
        except KeyError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Agregar registro",
                    "Tipo": type(error).__name__,
                    "Error": str(error),
                    "Hoja" : nombre_hoja
                    }
                )
            complemento.mensajes.hoja_no_encontrada(nombre_hoja)
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Agregar registro",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                    }
                )
            complemento.mensajes.error_inesperado(error)
        finally:
            if libro is not None:
                libro.close()
        if guardado:
            if libro_abierto is not None:
                comprobacion = self.comprobar_registro_com(hoja, fila_libre, datos)
            else:
                comprobacion = self.comprobar_registro_guardado(ruta_excel, nombre_hoja, fila_libre, datos)
            if comprobacion is True:
                self.logs.escribir(
                    "AGREGAR", {
                        "Archivo" : os.path.basename(ruta_excel),
                        "Hoja" : nombre_hoja,
                        "Fila" : fila_libre,
                        "Cliente" : datos[3],
                        "Folio Factura" : datos[4],
                        "Monto Facturado" : datos[6],
                        "Estatus": datos[7] 
                    }
                )
                folio_guardado = self.normalizar_valor(datos[4])
                if folio_guardado:
                    self.folios_registrados.add(folio_guardado)
                complemento.mensajes.registro_guardado(fila_libre)
                return True
            elif comprobacion is False:
                complemento.mensajes.error_comprobar_guardado()
                return False
            return False     

    def fila_vacia(self, fila):
        return all(celda is None or str(celda).strip() == "" for celda in fila)

    def comprobar_coincidencias(self, fila, criterios):
        if self.fila_vacia(fila):
            return False
        mapa_columnas = complemento.mapa_columnas()
        for campo, valor_buscado in criterios.items():
            if valor_buscado is None:
                continue
            valor_buscado = str(valor_buscado).strip()
            if valor_buscado == "":
                continue
            indice = mapa_columnas[campo]
            valor_excel = fila[indice]
            if valor_excel is None:
                valor_excel = ""
            texto_buscado = str(valor_buscado).strip().lower()
            texto_excel = str(valor_excel).strip().lower()
            if texto_buscado not in texto_excel:
                return False
        return True

    def buscar_registros(self, criterios):
        libro = None
        try:
            resultados = []
            archivos_excel = self.obtener_archivos_excel()
            for ruta_excel in archivos_excel:
                libro_abierto = self.obtener_libro_abierto(ruta_excel)
                if libro_abierto is not None:
                    for hoja in libro_abierto.Worksheets:
                        nombre_hoja = hoja.Name
                        if nombre_hoja.upper().startswith("FOLIOS_REPETIDOS_"):
                            continue
                        if not self.hoja_compatible_com(hoja):
                            continue
                        ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
                        coincidencias_hoja = 0
                        for numero_fila in range(2, ultima_fila + 1):
                            fila = []
                            for columna in range(1, 11):
                                valor = hoja.Cells(numero_fila, columna).Value
                                valor = self.formatear_valor_busqueda(valor, columna)
                                fila.append(valor)
                            if self.comprobar_coincidencias(fila, criterios):
                                coincidencias_hoja += 1
                                nombre_archivo = os.path.basename(ruta_excel)
                                resultados.append([nombre_archivo, nombre_hoja, numero_fila] + fila)
                else:
                    # OPENPYXL
                    libro = openpyxl.load_workbook(
                    ruta_excel,
                    read_only = True,
                    data_only = True
                    )
                    hojas_excel = libro.sheetnames            
                    for nombre_hoja in hojas_excel:
                        if nombre_hoja.upper().startswith("FOLIOS_REPETIDOS_"):
                            continue
                        hoja = libro[nombre_hoja]
                        if not self.hoja_compatible(hoja):
                            continue
                        for numero_fila, fila in enumerate(hoja.iter_rows(min_row = 2, values_only = True), start = 2):
                            fila_formateada = []
                            for columna, valor in enumerate(fila[:10], start = 1):
                                valor = self.formatear_valor_busqueda(valor, columna)
                                fila_formateada.append(valor)
                            if self.comprobar_coincidencias(fila_formateada, criterios):
                                nombre_archivo = os.path.basename(ruta_excel)
                                resultados.append([nombre_archivo, nombre_hoja, numero_fila] + fila_formateada)
                    libro.close()
                    libro = None
            return resultados
        except FileNotFoundError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Buscar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_no_encontrado()
            return []
        except PermissionError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Buscar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_en_uso()
            return []
        except KeyError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Buscar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return []
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Buscar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return []
        finally:
            if libro is not None:
                libro.close()

    def actualizar_registro(self, nombre_archivo, nombre_hoja, fila, datos):
        libro = None
        actualizado = False
        datos_anteriores = None
        libro_abierto = None
        try:
            ruta_excel = os.path.join(self.carpeta_excel, nombre_archivo)
            fila = int(fila)
            libro_abierto = self.obtener_libro_abierto(ruta_excel)
            if libro_abierto is not None:
                # COM
                hoja = libro_abierto.Worksheets(nombre_hoja)
                datos_anteriores = []
                for columna in range(1, 11):
                    valor = hoja.Cells(fila, columna).Value
                    datos_anteriores.append(valor)
                for numero_columna, valor in enumerate(datos.values(), start = 1):
                    if numero_columna == 1:
                        valor = datetime.strptime(valor, "%d/%m/%Y")
                        hoja.Cells(fila, numero_columna).Value = valor
                        hoja.Cells(fila, numero_columna).NumberFormat = "dd/mm/aaaa"
                    else:
                        hoja.Cells(fila, numero_columna).Value = valor
                estatus = datos["estatus"]
                self.preparar_fila_com(hoja, fila, estatus)
                libro_abierto.Save()
                actualizado = True
            else:
                # OPENPYXL
                libro = openpyxl.load_workbook(ruta_excel)
                hoja = libro[nombre_hoja]
                datos_anteriores = []
                for columna in range(1, 11):
                    valor = hoja.cell(row = fila, column = columna).value 
                    datos_anteriores.append(valor)
                for numero_columna, valor in enumerate(datos.values(), start = 1):
                    hoja.cell(
                        row = fila, 
                        column = numero_columna,
                        value = valor 
                    )
                estatus = datos["estatus"]
                self.configuracion.preparar_fila(hoja, fila, estatus)
                libro.save(ruta_excel)
                actualizado = True
        except FileNotFoundError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Actualizar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_no_encontrado()
            return False
        except PermissionError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Actualizar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_en_uso()
            return False
        except KeyError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Actualizar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.hoja_no_encontrada(nombre_hoja)
            return False
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Actualizar registros",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return False
        finally:
            if libro is not None:
                libro.close()
        if actualizado:
            if libro_abierto is not None:
                comprobacion = self.comprobar_registro_com(hoja, fila, datos.values())
            else:
                comprobacion = self.comprobar_registro_actualizado(nombre_archivo, nombre_hoja, fila, datos)
            if comprobacion is True:
                cambios = self.obtener_cambios(datos_anteriores, datos.values())
                detalles_log = {
                    "Archivo": nombre_archivo,
                    "Hoja": nombre_hoja,
                    "Fila": fila
                }
                detalles_log.update(cambios)
                self.logs.escribir(
                    "ACTUALIZAR",
                    detalles_log
                )
                folio_nuevo = self.normalizar_valor(datos["folio_factura"])
                self.cargar_folios_memoria()
                if folio_nuevo:
                    self.folios_registrados.add(folio_nuevo)
                complemento.mensajes.registro_actualizado(fila)
                return True
            elif comprobacion is False:
                complemento.mensajes.error_comprobar_actualizacion()
                return False
            return False

    def comprobar_registro_guardado(self, ruta_excel, nombre_hoja, fila, datos):
        libro_verificacion = None
        try:
            libro_verificacion = openpyxl.load_workbook(
                ruta_excel,
                read_only = True,
                data_only = True
            )
            hoja_verificacion = libro_verificacion[nombre_hoja]
            fila_guardada = next(
                hoja_verificacion.iter_rows(
                    min_row = fila,
                    max_row = fila,
                    min_col = 1,
                    max_col = 10,
                    values_only = True
                )
            )
            datos_excel_normalizados = [self.normalizar_valor(valor) for valor in fila_guardada]
            datos_programa_normalizados = [self.normalizar_valor(valor) for valor in datos]
            return (datos_excel_normalizados == datos_programa_normalizados)
        except Exception as error:
                self.logs.escribir(
                    "ERROR", {
                        "Proceso": "Comprobar registro guardado",
                        "Tipo": type(error).__name__,
                        "Error": str(error)
                    }
                )
                complemento.mensajes.error_inesperado(error)
                return None
        finally:
            if libro_verificacion is not None:
                libro_verificacion.close()
    
    def buscar_fila_libre(self, hoja):
        for numero_fila in range(2, hoja.max_row + 2):
            fila_vacia = True
            for columna in range(1, 11):
                valor = hoja.cell(
                    row = numero_fila,
                    column = columna
                ).value
                if valor is not None and str(valor).strip() != "":
                    fila_vacia = False
                    break
            if fila_vacia:
                return numero_fila

    def buscar_fila_libre_com(self, hoja):
        numero_fila = 2
        while True:
            fila_vacia = True
            for columna in range(1, 11):
                valor = hoja.Cells(numero_fila, columna).Value
                if valor is not None and str(valor).strip() != "":
                    fila_vacia = False
                    break
            if fila_vacia:
                return numero_fila
            numero_fila += 1

    def comprobar_registro_actualizado(self, nombre_archivo, nombre_hoja, fila, datos):
        libro = None
        try:
            ruta_excel = os.path.join(
            self.carpeta_excel,
            nombre_archivo
            )
            libro = openpyxl.load_workbook(
                ruta_excel,
                read_only = True,
                data_only = True
            )   
            hoja = libro[nombre_hoja]
            fila_excel = next(
                hoja.iter_rows(
                    min_row = fila,
                    max_row = fila,
                    min_col = 1,
                    max_col = 10,
                    values_only = True
                )
            )
            datos_excel = list(fila_excel)
            datos_programa = list(datos.values())
            datos_excel_normalizados = [self.normalizar_valor(valor) for valor in datos_excel]
            datos_programa_normalizados = [self.normalizar_valor(valor) for valor in datos_programa]
            return (datos_excel_normalizados == datos_programa_normalizados)
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Comprobar registro actualizado",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return None
        finally:
            if libro is not None:
                libro.close()

    def obtener_archivos_excel(self):
        archivos_excel = []
        os.makedirs(self.carpeta_excel, exist_ok = True)
        for archivo in os.listdir(self.carpeta_excel):
            if (archivo.lower().endswith(".xlsx") and not archivo.startswith("~$")): 
                ruta = os.path.join(self.carpeta_excel, archivo)
                archivos_excel.append(ruta)
        return archivos_excel
    
    def obtener_hojas_excel(self, ruta_excel):
        libro = None
        try:
            libro = openpyxl.load_workbook(
                ruta_excel,
                read_only = True,
                data_only = True
            )
            return libro.sheetnames
        except Exception as error:
            self.logs.escribir(
                "ERROR", { 
                 "Proceso": "Obtener hojas de Excel",
                 "Tipo": type(error).__name__,
                 "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return []
        finally:
            if libro is not None:
                libro.close()
        
    def hoja_compatible(self, hoja):
        encabezados_esperados = complemento.encabezados_excel()
        encabezados_excel = []
        for celda in hoja [1]:
            if celda.value is None:
                valor = ""
            else:
                valor = str(celda.value).strip().upper()
            encabezados_excel.append(valor)
        return encabezados_excel[:10] == encabezados_esperados

    def hoja_compatible_com(self, hoja):
        encabezados_esperados = complemento.encabezados_excel()
        encabezados_excel = []
        for columna in range (1, 11):
            valor = hoja.Cells(1, columna).Value
            if valor is None:
                valor = ""
            else:
                valor = str(valor).strip().upper()
            encabezados_excel.append(valor)
        return encabezados_excel == encabezados_esperados

    def obtener_libro_abierto(self, ruta_excel):
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return None
        ruta_buscada = os.path.normcase(os.path.abspath(ruta_excel))
        nombre_buscado = os.path.basename(ruta_excel).lower()
        for libro in excel.Workbooks:
            ruta_libro = str(libro.FullName)
            nombre_libro = str(libro.Name).lower()
            if ruta_libro.lower().startswith("http"):
                if nombre_libro == nombre_buscado:
                    return libro
            else:
                ruta_libro_normalizada = os.path.normcase(os.path.abspath(ruta_libro))
                if ruta_libro_normalizada == ruta_buscada:
                    return libro
        return None

    def obtener_o_crear_hoja_com(self, libro, nombre_hoja):
        for hoja in libro.Worksheets:
            if hoja.Name == nombre_hoja:
                return hoja, False
        hoja = libro.Worksheets.Add(
            After = libro.Worksheets(libro.Worksheets.Count)
            )
        hoja.Name = nombre_hoja
        return hoja, True

    def preparar_hoja_com(self, hoja):
        columnas = complemento.encabezados_excel()
        for numero_columna, encabezado in enumerate(columnas, start = 1):
            hoja.Cells(1, numero_columna).Value = encabezado
        rango_encabezado = hoja.Range("A1:J1")
        color_encabezado = self.convertir_color_com("1F4E78")
        rango_encabezado.Interior.Color = color_encabezado
        rango_encabezado.Font.Bold = True
        rango_encabezado.Font.Color = 16777215
        rango_encabezado.HorizontalAlignment = -4108
        rango_encabezado.VerticalAlignment = -4108 
        rango_encabezado.Borders.LineStyle = 1
        rango_encabezado.Borders.Weight = 2
        anchos = complemento.anchos_columnas()
        for columna, ancho in anchos.items():
            hoja.Columns(columna).ColumnWidth = ancho
        hoja.Rows(1).RowHeight = 25
        rango_encabezado.AutoFilter()        
        hoja.Activate()
        ventana_excel = hoja.Application.ActiveWindow
        ventana_excel.SplitRow = 1
        ventana_excel.SplitColumn = 0
        ventana_excel.FreezePanes = True
        
    def convertir_color_com(self, color):
        if isinstance(color, tuple):
            color = color[0]
        print("Color recibido:", repr(color))
        print("Tipo recibido:", type(color).__name__)
        color = color.lstrip("#")
        rojo = int(color[0:2], 16)
        verde = int(color[2:4], 16)
        azul = int(color[4:6], 16)
        color_com = (rojo + (verde << 8) + (azul << 16))
        return color_com

    def preparar_fila_com(self, hoja, numero_fila, estatus):
        colores = complemento.colores_estatus()
        estatus = str(estatus).strip().upper()
        color = colores.get(estatus, "FFFFFF")
        color_com = self.convertir_color_com(color)
        for numero_columna in range(1, 11):
            celda = hoja.Cells(numero_fila, numero_columna)
            celda.Interior.Color = color_com
            celda.Borders.LineStyle = 1
            celda.Borders.Weight = 2
            celda.HorizontalAlignment = -4108
            celda.VerticalAlignment = -4108

    def comprobar_registro_com(self, hoja, fila, datos):
        try:
            datos_excel = []
            for columna in range(1, 11):
                valor = hoja.Cells(fila, columna).Value
                datos_excel.append(valor)
            datos_excel_normalizados = [self.normalizar_valor(valor) for valor in datos_excel]
            datos_programa_normalizados = [self.normalizar_valor(valor) for valor in datos]
            return (datos_excel_normalizados == datos_programa_normalizados)
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Comprobar registro COM",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return None

    def normalizar_valor(self, valor):
        if valor is None:
            return ""
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y")
        texto = str(valor).strip()
        try:
            numero = float(texto)
            if numero.is_integer():
                return str(int(numero))
            return str(numero)
        except ValueError:
            return texto

    def formatear_valor_busqueda(self, valor, columna):
        if valor is None:
            return ""
        if columna == 1:
            if isinstance(valor, datetime):
                return valor.strftime("%d/%m/%Y")
            texto =  str(valor).strip()
            if "00:00:00" in texto:
                texto = texto.replace(" 00:00:00", "")
            return texto
        if columna in (2, 3, 5):
            try:
                numero = float(valor)
                if numero.is_integer():
                    return str(int(numero))
            except (ValueError, TypeError):
                pass
            return str(valor).strip()
        if columna == 8:
            return str(valor).strip().upper()
        return str(valor).strip()

    def obtener_cambios(self, datos_anteriores, datos_nuevos):
        nombres = complemento.nombres_columnas()
        cambios = {}
        for nombre, anterior, nuevo in zip(
            nombres,
            datos_anteriores,
            datos_nuevos
        ):
            anterior_normalizado = self.normalizar_valor(anterior)
            nuevo_normalizado = self.normalizar_valor(nuevo)
            if anterior_normalizado != nuevo_normalizado:
                cambios[nombre] = (f"{anterior_normalizado} --> {nuevo_normalizado}")
        return cambios

    def leer_incorporaciones(self):
        if not os.path.isfile(self.ruta_incorporaciones):
            return {}
        with open(self.ruta_incorporaciones, "r", encoding = "utf-8") as archivo:
            return json.load(archivo)

    def registrar_incorporacion(self, folio, archivo):
        datos = self.leer_incorporaciones()
        clave = f"{folio}/{archivo}"
        if clave in datos:
            return
        datos[clave] = datetime.now().strftime("%d/%m/%Y %H:%M")
        ruta_temporal = self.ruta_incorporaciones + ".tmp"
        with open(ruta_temporal, "w", encoding = "utf-8") as archivo_json:
            json.dump(datos, archivo_json, indent=4, ensure_ascii=False)
        os.replace(ruta_temporal, self.ruta_incorporaciones)

    def organizar_archivos_facturas(self, folio_factura = None):
        try:
            formatos = complemento.formatos_permitidos()
            if folio_factura is None:
                folios = self.folios_registrados
            else:
                folio_factura = self.normalizar_valor(folio_factura)
                if not folio_factura:
                    return []
                folios = {folio_factura}
            for archivo in os.listdir(self.carpeta_facturas):
                ruta_archivo = os.path.join(self.carpeta_facturas, archivo)
                if os.path.isdir(ruta_archivo):
                    continue
                nombre_archivo, extension = os.path.splitext(archivo)
                if extension.lower() not in formatos:
                    continue
                numeros_archivo = self.extraer_numeros_archivo(nombre_archivo)
                for folio in folios:
                    if folio not in numeros_archivo:
                        continue
                    carpeta_folio = os.path.join(self.carpeta_facturas, folio)
                    os.makedirs(carpeta_folio, exist_ok = True)
                    nueva_ruta = os.path.join(carpeta_folio, archivo)
                    if not os.path.exists(nueva_ruta):
                        shutil.move(ruta_archivo, nueva_ruta)
                        try:
                            self.registrar_incorporacion(folio, archivo)
                        except (OSError, ValueError, json.JSONDecodeError) as error:
                            self.logs.escribir(
                                "ERROR", {
                                    "Proceso": "Registrar incorporacion",
                                    "Archivo": archivo,
                                    "Tipo": type(error).__name__,
                                    "Error": str(error)
                                }
                            )
                    break
            if folio_factura is None:
                return []
            carpeta_folio = os.path.join(self.carpeta_facturas, folio_factura)
            if not os.path.isdir(carpeta_folio):
                return []
            archivos_encontrados = []
            for archivo in os.listdir(carpeta_folio):
                ruta_archivo = os.path.join(carpeta_folio, archivo)
                if not os.path.isfile(ruta_archivo):
                    continue
                nombre_archivo, extension = os.path.splitext(archivo)
                if extension.lower() in formatos:
                    archivos_encontrados.append(archivo)
            return archivos_encontrados
        except PermissionError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Organizar archivos de facturas",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            return []
        except OSError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Organizar archivos de facturas",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            return []
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Organizar archivos de facturas",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            return []
        
    def extraer_numeros_archivo(self, nombre_archivo):
        nombre_sin_extension = os.path.splitext(nombre_archivo)[0]
        numeros = re.findall(r"\d+", nombre_sin_extension)
        return numeros

    def obtener_folios_registrados(self):
        folios = set()
        libro = None
        try: 
            archivos_excel = self.obtener_archivos_excel()
            for ruta_excel in archivos_excel:
                libro_abierto = self.obtener_libro_abierto(ruta_excel)
                if libro_abierto is not None:
                    for hoja in libro_abierto.Worksheets:
                        if hoja.Name.upper().startswith("FOLIOS_REPETIDOS_"):
                            continue
                        if not self.hoja_compatible_com(hoja):
                            continue
                        ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
                        for fila in range (2, ultima_fila + 1):
                            folio = hoja.Cells(fila, 5).Value
                            folio = self.normalizar_valor(folio)
                            if folio:
                                folios.add(folio)
                else:
                    libro = openpyxl.load_workbook(
                        ruta_excel,
                        read_only = True,
                        data_only = True
                    )
                    for nombre_hoja in libro.sheetnames:
                        if nombre_hoja.upper().startswith("FOLIOS_REPETIDOS_"):
                            continue
                        hoja = libro[nombre_hoja]
                        if not self.hoja_compatible(hoja):
                            continue
                        for fila in hoja.iter_rows(
                            min_row = 2,
                            min_col = 5,
                            max_col = 5,
                            values_only = True
                            ):
                            folio = self.normalizar_valor(fila[0])
                            if folio:
                                folios.add(folio)    
                    libro.close()
                    libro = None
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Obtener folios registrados",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
        finally:
            if libro is not None:
                libro.close()
        return folios

    def cargar_folios_memoria(self):
        self.folios_registrados = (self.obtener_folios_registrados())

    def folio_ya_registrado(self, folio_factura):
        folio_factura = self.normalizar_valor(folio_factura)
        if not folio_factura:
            return False
        return folio_factura in self.folios_registrados

    def registrar_folio_repetido(self, datos):
        libro = None
        try:
            año = datetime.now().year
            ruta_excel = self.configuracion.obtener_ruta_excel_actual()
            nombre_hoja = f"FOLIOS_REPETIDOS_{año}"
            libro_abierto = self.obtener_libro_abierto(ruta_excel)
            if libro_abierto is not None:
                hoja, hoja_creada = self.obtener_o_crear_hoja_com(libro_abierto, nombre_hoja)
                if hoja_creada:
                    self.preparar_hoja_com(hoja)
                fila_libre = self.buscar_fila_libre_com(hoja)
                for numero_columna, valor in enumerate(
                    datos, start = 1
                ):
                    if numero_columna == 1:
                        valor = datetime.strptime(valor, "%d/%m/%Y")
                        hoja.Cells(fila_libre, numero_columna).Value = valor      
                        hoja.Cells(fila_libre, numero_columna).NumberFormat = "dd/mm/aaaa"
                    else:
                        hoja.Cells(fila_libre,numero_columna).Value = valor
                estatus = datos[7]
                self.preparar_fila_com(
                    hoja,
                    fila_libre,
                    estatus
                )
                libro_abierto.Save()
            else:
                self.configuracion.preparar_excel_actual()
                libro = openpyxl.load_workbook(ruta_excel)
                if nombre_hoja not in libro.sheetnames:
                    hoja = libro.create_sheet(nombre_hoja)
                    self.configuracion.preparar_hoja(hoja)
                else:
                    hoja = libro[nombre_hoja]
                fila_libre = self.buscar_fila_libre(hoja)
                for numero_columna, valor in enumerate(datos, start = 1):
                    hoja.cell(
                        row = fila_libre,
                        column = numero_columna,
                        value = valor 
                    )
                estatus = datos[7]
                self.configuracion.preparar_fila(hoja, fila_libre, estatus)
                libro.save(ruta_excel)
            self.logs.escribir(
                f"FOLIO_REPETIDO_{año}",
                {
                    "Archivo": os.path.basename(ruta_excel),
                    "Hoja": nombre_hoja,
                    "Fila": fila_libre,
                    "Fecha" : datos[0],
                    "Numero de Cotizacion" : datos[1],
                    "Numero Orden de Compra" : datos[2],
                    "Cliente" : datos[3],
                    "Folio Factura" : datos[4],
                    "Concepto" : datos[5],
                    "Monto Facturado" : datos[6],
                    "Estatus" : datos[7],
                    "Complemento 1" : datos[8],
                    "Complemento 2": datos[9]
                }
            )
            complemento.mensajes.folio_repetido_guardado(datos[4], nombre_hoja, fila_libre)
            return True
        except FileNotFoundError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Registrar folio repetido",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_no_encontrado()
            return False
        except PermissionError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Registrar folio repetido",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.archivo_en_uso()
            return False
        except KeyError as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Registrar folio repetido ",
                    "Tipo": type(error).__name__,
                    "Error": str(error),
                    "Hoja" : nombre_hoja
                    }
                )
            complemento.mensajes.hoja_no_encontrada(nombre_hoja)
            return False
        except Exception as error:
            self.logs.escribir(
                "ERROR", {
                    "Proceso": "Registrar folio repetido",
                    "Tipo": type(error).__name__,
                    "Error": str(error)
                }
            )
            complemento.mensajes.error_inesperado(error)
            return False
        finally:
            if libro is not None:
                libro.close()

class configuracion_excel:
    def __init__(self, carpeta_excel, logs):
        self.carpeta_excel = carpeta_excel
        self.logs = logs

    def obtener_ruta_excel_actual(self):
        año_actual = datetime.now().year
        nombre_archivo = f"FACTURAS_{año_actual}.xlsx"
        ruta_excel = os.path.join(self.carpeta_excel, nombre_archivo)
        return ruta_excel

    def obtener_mes_actual(self):
        meses = complemento.meses_excel()
        numero_mes = datetime.now().month
        return meses[numero_mes]
    
    def preparar_excel_actual(self):
        ruta_excel = self.obtener_ruta_excel_actual()
        nombre_hoja = self.obtener_mes_actual()
        libro = None
        try:
            os.makedirs(self.carpeta_excel, exist_ok = True)
            if not os.path.exists(ruta_excel):
                libro = openpyxl.Workbook()
                hoja = libro.active
                hoja.title = nombre_hoja
                self.preparar_hoja(hoja)
                libro.save(ruta_excel)
                self.logs.escribir(
                    "CREAR ARCHIVO", { 
                        "Archivo": os.path.basename(ruta_excel),
                        "Hoja Inicial": nombre_hoja
                    }
                )
            else:
                libro = openpyxl.load_workbook(ruta_excel)
                if nombre_hoja not in libro.sheetnames:
                    hoja = libro.create_sheet(nombre_hoja)
                    self.preparar_hoja(hoja)
                    libro.save(ruta_excel)
                    self.logs.escribir(
                        "CREAR HOJA", 
                        { 
                            "Archivo": os.path.basename(ruta_excel),
                            "Hoja Inicial": nombre_hoja
                        }
                    )
            return ruta_excel, nombre_hoja
        finally:
            if libro is not None:
                libro.close()

    def preparar_hoja(self, hoja):
        columnas = complemento.encabezados_excel()
        hoja.append(columnas)
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = "A1:J1"
        borde = complemento.borde_celdas()
        for celda in hoja[1]:
            # Fuente del texto / bolda, color, tamaño, etc.
            celda.font = Font(bold = True, color = "FFFFFF")
            # Fondo de la celda / fill_type y fgColor
            celda.fill = PatternFill(fill_type = "solid", fgColor = "1F4E78")
            # Posicion del contenido / horizontal y vertical
            celda.alignment = Alignment(horizontal = "center", vertical = "center")
            celda.border = borde
        anchos = complemento.anchos_columnas()
        for columna, ancho in anchos.items():
            hoja.column_dimensions[columna].width = ancho
        hoja.row_dimensions[1].height = 25

    def preparar_fila(self, hoja, numero_fila, estatus):
        borde = complemento.borde_celdas()
        colores = complemento.colores_estatus()
        estatus = str(estatus).strip().upper()
        color = colores.get(estatus, ("FFFFFF", "000000"))[0]
        if isinstance(color, str) and color.startswith("#"):
            color = color.lstrip("#")
        if len(color) == 6:
            color = "FF" + color
        for numero_columna in range (1, 11):
            celda = hoja.cell(row = numero_fila, column = numero_columna)
            celda.border = borde
            celda.fill = PatternFill(fill_type = "solid", fgColor = color)
            celda.alignment = Alignment(
                horizontal = "center",
                vertical = "center"
            )

class registros_logs:
    def __init__(self, capeta_programa):
       self.carpeta_registros = os.path.join(capeta_programa, "registros")
       
    def obtener_ruta_log(self):
       año_actual = datetime.now().year
       nombre_log = f"REGISTRO_{año_actual}.log"
       ruta_log = os.path.join(self.carpeta_registros, nombre_log)
       return ruta_log

    def escribir(self, acccion, detalles):
        try: 
            os.makedirs(self.carpeta_registros, exist_ok = True)
            ruta_log = self.obtener_ruta_log()
            fecha_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            with open(ruta_log, "a", encoding = "utf-8") as archivo:
                archivo.write(f"[{fecha_hora}] {acccion}\n")
                for clave, valor in detalles.items():
                    archivo.write(f"{clave}: {valor}\n")
                archivo.write("\n")
            return True
        except Exception:
           return False






