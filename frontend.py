import customtkinter as ctk
from tkinter import messagebox
from tkinter import filedialog
import os
import AutoExcel as AE
import json

# Configuración de CustomTkinter
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")





class App(ctk.CTk):
    def __init__(self, debug=True):
        super().__init__()
        
        self.debug = debug
        self.title('AutoExcel - Generador de Nóminas')
        ancho_ventana = 500
        alto_ventana = 300
        
        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        
        self.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
        self.resizable(False, False)
        

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
        Nnomina_lb = ctk.CTkLabel(master = self.main_fr, text = "Ingrese el numero de nomina :")
        Nnomina_lb.pack(pady = (0,5))
        Nnomina_et = ctk.CTkEntry(master = self.main_fr, placeholder_text = "Ejemplo 12177")
        Nnomina_et.pack(pady = (0,20))
        nomina_btn = ctk.CTkButton(master=self.main_fr, text="Generar Nómina", width=150 ,height=30, command=lambda: self.nomina_btn_fun(123))
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
        config_btn = ctk.CTkButton(master=aux_frame, text="❮" , width=30 ,height=30, fg_color="transparent",hover_color=("gray75", "gray25"), text_color=("black", "white"), command=lambda: self.cambiarventana(self.main_fr, self.config_fr))
        config_btn.grid(row=0, column= 0, sticky="w", padx=10)
        aux_frame.pack(fill="x",pady=(15,40))


        path_fr = ctk.CTkFrame(master=self.config_fr, fg_color="transparent")
        path_fr.pack(pady=20, padx=20, fill="x")

        # Configuración de columnas: 3/10, 6/10 y 1/10
        path_fr.grid_columnconfigure(0, weight=2)
        path_fr.grid_columnconfigure(1, weight=7)
        path_fr.grid_columnconfigure(2, weight=1) 

            # Base de Datos
        db_path_lb = ctk.CTkLabel(path_fr, text="Base de Datos :")
        db_path_lb.grid(row=0, column=0, sticky="w", padx=5, pady=10)
        self.db_path_et = ctk.CTkEntry(path_fr, placeholder_text="Seleccione un archivo...")
        self.db_path_et.grid(row=0, column=1, sticky="ew", padx=5, pady=10)
        self.db_path_et.configure(state="readonly")
        db_path_btn = ctk.CTkButton(path_fr, text="📁", width=40, 
                                    command=lambda: self.buscar_archivo(self.db_path_et, "db"))
        db_path_btn.grid(row=0, column=2, sticky="ew", padx=5, pady=10)

            # Carpeta de Salida
        output_path_lb = ctk.CTkLabel(path_fr, text="Carpeta de Salida :")
        output_path_lb.grid(row=1, column=0, sticky="w", padx=5, pady=10)
        self.output_path_et = ctk.CTkEntry(path_fr, placeholder_text="Seleccione una carpeta...")
        self.output_path_et.grid(row=1, column=1, sticky="ew", padx=5, pady=10)
        self.output_path_et.configure(state="readonly")
        output_path_btn = ctk.CTkButton(path_fr, text="📁", width=40, 
                                        command=lambda: self.buscar_archivo(self.output_path_et, "folder"))
        output_path_btn.grid(row=1, column=2, sticky="ew", padx=5, pady=10)

        btn_fr = ctk.CTkFrame(master=self.config_fr, fg_color="transparent")
        btn_fr.pack(pady=20) # Este pady le da espacio respecto a lo de arriba

        guardar_btn = ctk.CTkButton(master=btn_fr, text="Guardar", width=150,command=lambda: self.guardar_configuracion())
        guardar_btn.pack(side="left", padx=10)

        cancelar_btn = ctk.CTkButton(master=btn_fr, text="Cancelar", width=150, height=30, command=lambda: self.confirmar_salida())
        cancelar_btn.pack(side="left", padx=10)


        self.cargar_configuracion()


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
        if ventana_a_mostrar == self.main_fr:
            self.cargar_configuracion()

    def guardar_configuracion(self):
        config_data = {
            "ruta_db": self.db_path_et.get(),
            "ruta_output": self.output_path_et.get(),
        }
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            messagebox.showinfo("Éxito", "Configuración guardada correctamente.")
            self.cambiarventana(self.main_fr,self.config_fr) # Regresar al menú principal tras guardar
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {e}")

    def cargar_configuracion(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    config_data = json.load(f)
                
                # Insertamos los datos en los entries
                # Recuerda: como son 'readonly', hay que habilitarlos primero
                self.actualizar_entry(self.db_path_et, config_data.get("ruta_db", ""))
                self.actualizar_entry(self.output_path_et, config_data.get("ruta_output", ""))
                
            except Exception as e:
                print(f"Error al cargar configuración: {e}")

    def actualizar_entry(self, entry, texto):
        entry.configure(state="normal") 
        entry.delete(0, "end")
        
        if texto and texto.strip() != "":
            entry.insert(0, texto)
            entry.configure(state="readonly") 
        else:
            entry.configure(placeholder_text="Seleccione una ruta...")
            entry.configure(state="readonly")
    
    def confirmar_salida(self):
        print("entro")
        config_guardada = {"ruta_db": "", "ruta_output": ""}
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                config_guardada = json.load(f)

        db_actual = self.db_path_et.get()
        out_actual = self.output_path_et.get()


        if db_actual != config_guardada.get("ruta_db", "") or \
        out_actual != config_guardada.get("ruta_output", ""):
            respuesta = messagebox.askyesno(
                "Cambios sin guardar", 
                "Has realizado cambios en la configuración. ¿Deseas salir sin guardar?"
            )
            if not respuesta:
                return 
        self.cambiarventana(self.main_fr, self.config_fr)
    
    def nomina_btn_fun(self, Nnomina):
        conn= AE.open_database()
        if conn is None:
            messagebox.showerror("Error", "No se pudo conectar a la base de datos. Verifica la configuración.")
            return
        cur = conn.cursor()
        print("si")






if __name__ == "__main__":
    app = App(debug=False) # Cambia a False para ocultar el log
    app.mainloop()