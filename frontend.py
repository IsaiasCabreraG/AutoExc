
import os, sys

if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
else:
    base = os.path.dirname(__file__)

os.environ["TCL_LIBRARY"] = os.path.join(base, "_tcl_data", "tcl8.6")
os.environ["TK_LIBRARY"]  = os.path.join(base, "_tk_data", "tk8.6")

import customtkinter as ctk
from tkinter import messagebox
from tkinter import filedialog
import AutoExcel as AE
import json
import logging


# Configuración de CustomTkinter
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

def solo_numeros(P):

    if P == "":
        return True
    return P.isdigit()



class App(ctk.CTk):
    def __init__(self, debug=True):
        super().__init__()
        
        self.debug = debug
        self.title('AutoExcel - Generador de Nóminas')
        ancho_ventana = 500
        alto_ventana = 300 if not self.debug else 550
        
        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        
        self.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
        self.resizable(False, False)
        ruta_icono = AE.obtener_ruta_recurso("static/logo.ico")
        self.iconbitmap(ruta_icono)
        # img = tk.PhotoImage(file='static/logo.png')
        # self.iconphoto(False, img)
        self.logger = logging.getLogger("AutoExcel")
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(levelname)s: %(message)s')

        # Siempre guardar en archivo (aunque el modo debug esté apagado)
        file_handler = logging.FileHandler("autoexcel_debug.log", mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        

        #Pagina Menu Principal
        self.main_fr = ctk.CTkFrame(master=self)
        aux_frame = ctk.CTkFrame(master=self.main_fr, fg_color="transparent")
        aux_frame.grid_columnconfigure(0, weight=1,uniform="a")
        aux_frame.grid_columnconfigure(1, weight=0)
        aux_frame.grid_columnconfigure(2, weight=1,uniform="a")
        title_lb= ctk.CTkLabel(master=aux_frame, text="AutoExcel",font =("",32,"bold"))
        title_lb.grid(row=0, column= 1)
        config_btn = ctk.CTkButton(master=aux_frame, text="⚙️", width=30 ,height=30, fg_color="transparent",hover_color=("gray75", "gray25"), text_color=("black", "white"), command=lambda: self.cambiarventana(self.config_fr, self.main_fr))
        config_btn.grid(row=0, column= 2, sticky="e", padx=10)
        aux_frame.pack(fill="x",pady=(15,40))
        Nnomina_lb = ctk.CTkLabel(master=self.main_fr, text="Ingrese Nómina y Año:")
        Nnomina_lb.pack(pady=(0, 5))

        self.container_inputs = ctk.CTkFrame(master=self.main_fr, fg_color="transparent")
        self.container_inputs.pack(pady=(0, 20))
        Nnomina_et = ctk.CTkEntry(master=self.container_inputs, placeholder_text="Ej: 12177", width=140)
        Nnomina_et.grid(row=0, column=0, padx=5)
        Nnomina_et.bind("<KeyRelease>", self.validar_en_tiempo_real)
        self.año_var = ctk.StringVar(value="----")
        self.año_dd = ctk.CTkOptionMenu(
            master=self.container_inputs,
            values=["----"],
            variable=self.año_var,
            width=100,
            corner_radius=5,
            fg_color=("#FFFFFF", "gray25"), 
            button_color=("#F0F0F0", "gray30"),
            button_hover_color=("#E5E5E5", "gray40"),
            text_color=("black", "white")
        )
        self.año_dd.grid(row=0, column=1, padx=5)
        self.año_dd.configure(state="disabled") 

        nomina_btn = ctk.CTkButton(master=self.main_fr, text="Generar Nómina", width=150 ,height=30, command=lambda: self.nomina_btn_fun(Nnomina_et.get()))
        nomina_btn.pack()
        self.main_fr.pack( fill="both", expand=True)


        #Pagina Configuración
        self.config_fr= ctk.CTkFrame(master=self)
        aux_frame = ctk.CTkFrame(master=self.config_fr, fg_color="transparent")
        aux_frame.grid_columnconfigure(0, weight=1,uniform="a")
        aux_frame.grid_columnconfigure(1, weight=0)
        aux_frame.grid_columnconfigure(2, weight=1,uniform="a")
        title_lb= ctk.CTkLabel(master=aux_frame, text="Configuracion",font =("",32,"bold"))
        title_lb.grid(row=0, column= 1)
        config_btn = ctk.CTkButton(master=aux_frame, text="❮" , width=30 ,height=30, fg_color="transparent",hover_color=("gray75", "gray25"), text_color=("black", "white"), command=lambda: self.confirmar_salida())
        config_btn.grid(row=0, column= 0, sticky="w", padx=10)
        aux_frame.pack(fill="x",pady=(15,25))


        path_fr = ctk.CTkFrame(master=self.config_fr, fg_color="transparent")
        path_fr.pack(pady=20, padx=20, fill="x")

        # Configuración de columnas: 3/10, 6/10 y 1/10
        path_fr.grid_columnconfigure(0, weight=2)
        path_fr.grid_columnconfigure(1, weight=7)
        path_fr.grid_columnconfigure(2, weight=1) 

        rut_lb = ctk.CTkLabel(path_fr, text="RUT Empresa :")
        rut_lb.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.rut_empresa_et = ctk.CTkEntry(path_fr, placeholder_text="12.345.678-9")
        self.rut_empresa_et.grid(row=0, column=1, columnspan=2, sticky="w", padx=5, pady=5)

            # Base de Datos
        db_path_lb = ctk.CTkLabel(path_fr, text="Base de Datos :")
        db_path_lb.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.db_path_et = ctk.CTkEntry(path_fr, placeholder_text="Seleccione un archivo...")
        self.db_path_et.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.db_path_et.configure(state="readonly")
        db_path_btn = ctk.CTkButton(path_fr, text="📁", width=40, 
                                    command=lambda: self.buscar_archivo(self.db_path_et, "db"))
        db_path_btn.grid(row=1, column=2, sticky="ew", padx=5, pady=5)

            # Carpeta de Salida
        output_path_lb = ctk.CTkLabel(path_fr, text="Carpeta de Salida :")
        output_path_lb.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.output_path_et = ctk.CTkEntry(path_fr, placeholder_text="Seleccione una carpeta...")
        self.output_path_et.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.output_path_et.configure(state="readonly")
        output_path_btn = ctk.CTkButton(path_fr, text="📁", width=40, 
                                        command=lambda: self.buscar_archivo(self.output_path_et, "folder"))
        output_path_btn.grid(row=2, column=2, sticky="ew", padx=5, pady=5)

        btn_fr = ctk.CTkFrame(master=self.config_fr, fg_color="transparent")
        btn_fr.pack(pady=20) # Este pady le da espacio respecto a lo de arriba

        guardar_btn = ctk.CTkButton(master=btn_fr, text="Guardar", width=150,command=lambda: self.guardar_configuracion())
        guardar_btn.pack(side="left", padx=10)

        cancelar_btn = ctk.CTkButton(master=btn_fr, text="Cancelar", width=150, height=30, command=lambda: self.confirmar_salida())
        cancelar_btn.pack(side="left", padx=10)

        if self.debug:
            self.debug_fr= ctk.CTkFrame(master=self)
            log_lb = ctk.CTkLabel(master=self.debug_fr, text="Consola de Depuración", font=("Arial", 12, "bold"))
            log_lb.pack(pady=(10, 0))

            self.log_widget = ctk.CTkTextbox(master=self.debug_fr, width=460, height=200, font=("Consolas", 11))
            self.log_widget.pack(pady=10, padx=20)
            self.log_widget.configure(state="disabled")

            # Conectar el logger al recuadro
            text_handler = TextboxHandler(self.log_widget)
            text_handler.setFormatter(formatter)
            self.logger.addHandler(text_handler)
            self.debug_fr.pack(pady=10, padx=20,side ="bottom")

        self.logger.info("Aplicación iniciada correctamente")

        self.refrescar_años_db()
        self.cargar_configuracion()

    def validar_en_tiempo_real(self, event):
        valor = event.widget.get()
        if valor and not valor.isdigit():
            event.widget.delete(len(valor)-1, "end")


    def buscar_archivo(self,entry_path,tipo):
        
        if tipo == "db":
           ruta = filedialog.askopenfilename(title="Seleccionar Base de Datos",filetypes=[("Archivos de base de datos", "*.accdb"), ("Todos los archivos", "*.*")],)
        elif tipo == "folder":
            ruta = filedialog.askdirectory(title="Seleccionar Carpeta de Salida",)
        if ruta: 
            self.actualizar_entry(entry_path, ruta)
            print(f"Ruta seleccionada: {ruta}")

    def cambiarventana(self, ventana_a_mostrar, ventana_a_ocultar):
        ventana_a_ocultar.pack_forget()
        ventana_a_mostrar.pack(fill="both", expand=True)
        if ventana_a_mostrar == self.config_fr:
            self.cargar_configuracion()

    def guardar_configuracion(self):
        config_data = {
            "ruta_db": self.db_path_et.get(),
            "ruta_output": self.output_path_et.get(),
            "rut_empresa": self.rut_empresa_et.get()
        }
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            messagebox.showinfo("Éxito", "Configuración guardada correctamente.")
            self.cambiarventana(self.main_fr,self.config_fr)
            self.refrescar_años_db()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {e}")

    def cargar_configuracion(self):
        self.logger.info("Cargando configuración desde JSON...")
        if not(os.path.exists("config.json")):
            self.logger.warning("No se encontró config.json. Creando uno nuevo con valores vacíos.")
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump({"ruta_db": "", "ruta_output": "", "rut_empresa": ""}, f, indent=4)
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                config_data = json.load(f)
            
            # Insertamos los datos en los entries
            # Recuerda: como son 'readonly', hay que habilitarlos primero
            self.logger.debug("JSON cargado correctamente.")
            self.actualizar_entry(self.db_path_et, config_data.get("ruta_db", ""))
            self.actualizar_entry(self.rut_empresa_et, config_data.get("rut_empresa", ""))
            self.actualizar_entry(self.output_path_et, config_data.get("ruta_output", ""))
            
        except Exception as e:
            self.logger.error(f"Error al cargar configuración desde JSON : {e}")

    def actualizar_entry(self, entry, texto):
        self.logger.debug(f"actualizando entry{texto}")
        entry.configure(state="normal") 
        entry.delete(0, "end")
        
        if texto and texto.strip() != "":
            entry.insert(0, texto)
        else:
            if entry == self.db_path_et:
                entry.configure(placeholder_text="Seleccione un archivo...")
            elif entry == self.output_path_et:
                entry.configure(placeholder_text="Seleccione una ruta...")
            elif entry == self.rut_empresa_et:
                entry.configure(placeholder_text="12345678-9")
                return
        entry.configure(state="readonly") 
    
    def confirmar_salida(self):
        self.logger.info("Confirmando salida de la configuración...")
        config_guardada = {"ruta_db": "", "ruta_output": "", "rut_empresa": ""}
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                config_guardada = json.load(f)

        db_actual = self.db_path_et.get()
        out_actual = self.output_path_et.get()
        rut_actual = self.rut_empresa_et.get()


        if db_actual != config_guardada.get("ruta_db", "") or \
        out_actual != config_guardada.get("ruta_output", "") or \
        rut_actual != config_guardada.get("rut_empresa", ""):
            self.logger.warning("Cambios detectados en la configuración sin guardar.")
            respuesta = messagebox.askyesno(
                "Cambios sin guardar", 
                "Has realizado cambios en la configuración. ¿Deseas salir sin guardar?"
            )
            if not respuesta:
                self.logger.info("El usuario decidió no salir sin guardar.")
                return 
        self.logger.info("Saliendo de la configuración sin guardar.")
        self.cambiarventana(self.main_fr, self.config_fr)
    
    def nomina_btn_fun(self, Nnomina):
        self.logger.info("Botón de generación de nómina presionado.")
        if not Nnomina.strip().isdigit():
            messagebox.showwarning("Entrada Inválida", "Por favor, ingrese un número de nómina válido.")
            self.logger.warning("Intento de generar nómina sin número ingresado.")
            return
        self.logger.info(f"Iniciando generación de nómina para N°{Nnomina}...")
        conn = AE.open_database()
        if conn is None:
            messagebox.showerror("Error de Conexión", "No se pudo conectar a la base de datos.\nVerifica la configuración de la ruta.")
            self.logger.error("Conexión a la base de datos establecida.")
            return
        self.logger.debug("Conexión a la base de datos establecida.")
        try:
            cur = conn.cursor()
            egreso_datos = AE.datos_egreso(cur, Nnomina, self.año_var.get())
            if egreso_datos:
                self.logger.info(f"Datos obtenidos para nómina N°{Nnomina}-{self.año_var.get()}: {len(egreso_datos)} registros.")
                total = AE.llenar_nomina_xlsx(egreso_datos, Nnomina, self.año_var.get(), RutEmpresa=self.rut_empresa_et.get())
                total_formateado = f"${total:,.0f}".replace(",", ".")
                mensaje = (
                    f"Proceso Completado con Éxito\n\n"
                    f"Nómina N°: {Nnomina}\n"
                    f"Año: {self.año_var.get()}\n"
                    f"Total a Pagar: {total_formateado}\n"
                    f"Registros: {len(egreso_datos)}\n\n"
                    f"Ubicación: {self.output_path_et.get()}"
                )
                messagebox.showinfo("Nómina Generada", mensaje)
            
            else:
                messagebox.showwarning("Sin Resultados", f"No se encontró ningún egreso con el número: {Nnomina}")
                self.logger.warning(f"No se encontraron datos para nómina N°{Nnomina}-{self.año_var.get()}.")

        except Exception as e:
            messagebox.showerror("Error Crítico", f"Ocurrió un error inesperado:\n{e}")
            self.logger.error(f"Error durante la generación de la nómina: {e}")
        
        finally:
            if conn:
                conn.close()
    
    def refrescar_años_db(self):
        """Llama a AE para obtener años y actualiza el estado del Dropdown"""
        self.logger.info("Intentando obtener años de la base de datos...")
        lista_años = AE.obtener_años_disponibles()
        
        if lista_años:
            self.año_dd.configure(state="normal", values=lista_años)
            self.año_var.set(lista_años[0]) # Año más reciente
            self.logger.info(f"Dropdown actualizado con {len(lista_años)} años.")
        else:
            self.año_dd.configure(state="disabled", values=["----"])
            self.año_var.set("----")
            self.logger.warning("No se pudo acceder a los años. Dropdown deshabilitado.")

class TextboxHandler(logging.Handler):
    def __init__(self, textbox):
        super().__init__()
        self.textbox = textbox

    def emit(self, record):
        msg = self.format(record)
        self.textbox.configure(state="normal") 
        self.textbox.insert("end", msg + "\n")
        self.textbox.configure(state="disabled") 
        self.textbox.see("end") 






if __name__ == "__main__":
    app = App(debug=False)
    app.mainloop()