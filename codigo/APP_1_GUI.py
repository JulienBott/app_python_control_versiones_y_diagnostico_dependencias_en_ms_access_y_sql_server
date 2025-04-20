import pandas as pd
import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog as fd, ttk, scrolledtext
from threading import Thread

import APP_2_GENERAL as mod_gen
import APP_3_BACK_END_MS_ACCESS as mod_access
import APP_3_BACK_END_SQL_SERVER as mod_sql_server


#################################################################################################################################################################################
##                     VARIABLES GENERALES
#################################################################################################################################################################################

#colores para root y frames de la GUI
bg_GUI_inicio = "#309BBD"
bg_GUI_control_versiones = "#309BBD"
bg_GUI_control_versiones_merge = "#DBF18A"
bg_GUI_merge_bbdd_fisica = "#309BBD"
bg_GUI_merge_bbdd_fisica_errores = "#DDE1E2"

bg_GUI_scrolledtext_scripts = "#DDE1E2"


#colores para los scripts en el control de versiones
bg_GUI_lineas_control_versiones_cambios_localizados = "#05FB27"
bg_GUI_lineas_control_versiones_cambios_agregados = "#FBCB05"
bg_GUI_lineas_control_versiones_cambios_eliminados = "#05FBF0"


#resolucion pantalla --> se recomienda que sea de 1920x1080
#sino es posible que la GUI tenga los widgets descolocados debido al uso para colocarlos del metodo place en vez de usar pack o grid
resolucion_pantalla_recomendada = "1920x1080"


#################################################################################################################################################################################
##                    CLASE PARA GENERAR LOS WIDGETS
#################################################################################################################################################################################

class gui_widgets:
    #clase que permite crear los widgets recurrentes de la GUI

    def __init__(self, master, **kwargs):
        
        #parametros kwargs
        self.tipo_widget = kwargs.get("tipo_widget", None)
        self.text = kwargs.get("text", None)
        self.textvariable = kwargs.get("textvariable", None)
        self.bg = kwargs.get("bg", None)
        self.fg = kwargs.get("fg", None)
        self.width = kwargs.get("width", None)
        self.height = kwargs.get("height", None)
        self.justify = kwargs.get("justify", None)
        self.borderwidth = kwargs.get("borderwidth", None)
        self.relief = kwargs.get("relief", None)
        self.anchor = kwargs.get("anchor", None)
        self.font = kwargs.get("font", None)
        self.state = kwargs.get("state", None)
        self.place = kwargs.get("place", None)
        self.combobox_lista_valores = kwargs.get("combobox_lista_valores", None)
        self.combobox_values = kwargs.get("combobox_values", None)
        self.combobox_tipo_acceso_valores = kwargs.get("combobox_tipo_acceso_valores", None)
        self.combobox_bind_mousewheel = kwargs.get("combobox_bind_mousewheel", None)
        self.combobox_bind_comboboxselected_rutina = kwargs.get("combobox_bind_comboboxselected_rutina", None)
        self.command_proceso = kwargs.get("command_proceso", None)
        self.command = kwargs.get("command", None)
        self.command_parametro = kwargs.get("command_parametro", None)

        self.variable = kwargs.get("variable", None)
        self.wrap = kwargs.get("wrap", None)
        self.scrolledtext_item = kwargs.get("scrolledtext_item", None)
        self.scrolledtext_ini = kwargs.get("scrolledtext_ini", None)
        self.scrolledtext_fin = kwargs.get("scrolledtext_fin", None)
        self.scrolledtext_fin = kwargs.get("scrolledtext_fin", None)
        self.widget_frame = kwargs.get("widget_frame", None)

        self.combobox_bind_comboboxselected_rutina_con_param = kwargs.get("combobox_bind_comboboxselected_rutina_con_param", None)
        self.parametro_rutina_combobox = kwargs.get("parametro_rutina_combobox", None)


        #tipo widget soportados en la GUI del app
        if self.tipo_widget == "label":
            self.widget = tk.Label(master)

        elif self.tipo_widget == "combobox":
            self.widget = ttk.Combobox(master)

        elif self.tipo_widget == "button":
            self.widget = tk.Button(master)

        elif self.tipo_widget == "scrolledtext":
            self.widget = scrolledtext.ScrolledText(master)

        elif self.tipo_widget == "entry":
            self.widget = tk.Entry(master)

        elif self.tipo_widget == "entry":
            self.widget = tk.Entry(master)


        #configuraciones
        if self.text != None:
            self.widget.config(text = self.text)

        if self.textvariable != None:
            self.widget.config(textvariable = self.textvariable)

        if self.bg != None:
            self.widget.config(bg = self.bg)

        if self.fg != None:
            self.widget.config(fg = self.fg)

        if self.width != None:
            self.widget.config(width = self.width)

        if self.height != None:
            self.widget.config(height = self.height)

        if self.justify != None:
            self.widget.config(justify = self.justify)

        if self.borderwidth != None:
            self.widget.config(borderwidth = self.borderwidth)

        if self.relief != None:
            self.widget.config(relief = self.relief)

        if self.anchor != None:
            self.widget.config(anchor = self.anchor)
        
        if self.font != None:
            self.widget.config(font = self.font)

        if self.place != None:
            self.widget.place(x = self.place[0], y = self.place[1])

        if self.combobox_values != None:
            self.widget.configure(values = self.state)

        if self.combobox_lista_valores != None:
            self.widget['values'] = self.combobox_lista_valores

        if self.combobox_tipo_acceso_valores != None:
            self.widget['state'] = self.combobox_tipo_acceso_valores

        if self.combobox_bind_mousewheel != None:
            self.widget.bind("<MouseWheel>", lambda event: "break")

        if self.combobox_bind_comboboxselected_rutina != None:
            self.widget.bind('<<ComboboxSelected>>', self.combobox_bind_comboboxselected_rutina)



        if self.command_proceso != None:
            self.widget.config(command = lambda: self.command_proceso())


        if self.command != None:
            self.widget.config(command = lambda: self.command(self.command_parametro))


        if self.variable != None:
            self.widget.config(variable = self.variable)



    def config_lock_unlock(self, opcion_lock):

        if opcion_lock == "BLOQUEADO":
            self.widget.config(state = tk.DISABLED)

        elif opcion_lock == "DESBLOQUEADO":
            self.widget.config(state = tk.NORMAL)


    def combobox_update_lista_valores(self, **kwargs):
        self.widget['values'] = kwargs["combobox_lista_valores"]



    def def_combobox_bind_rutina_con_param(self, event):
        selected_item = self.widget.get()
        self.callback_method(selected_item)




    def combobox_bind_select(self, def_bind, opcion):
        self.widget.bind('<<ComboboxSelected>>', lambda event, combobox = self.widget: def_bind(event, combobox, opcion))


    def scrolledtext_delete(self, inicio, fin):
        self.widget.delete(inicio, fin)


    def scrolledtext_insert(self, index, texto):
        self.widget.insert(index, texto)


    def entry_bind(self, def_bind):
        self.widget.bind("<FocusOut>", lambda event, entry = self.widget: def_bind(event, entry))




#################################################################################################################################################################################
##                    CLASE PARA LA GUI DE INICIO
#################################################################################################################################################################################

class gui_ventana_inicio:
    #clase que permite generar la GUI de inicio con sus widgets y rutinas asociadas

    def __init__(self, master, **kwargs):

        
        resolucion_pantalla = str(master.winfo_screenwidth()) + "x" + str(master.winfo_screenheight())#se calcula es_resolucion_pantalla_recomendada para saber si informar en la GUI
        es_resolucion_pantalla_recomendada = "SI" if resolucion_pantalla == resolucion_pantalla_recomendada else "NO"#que la resolucion de la pantalla no es la recomendada


        self.master = master
        self.master.title(mod_gen.nombre_app)
        self.master.configure(bg = bg_GUI_inicio)

        if es_resolucion_pantalla_recomendada == "SI":
            self.master.geometry("840x460")
        else:
            self.master.geometry("840x500")



        self.master.iconbitmap(mod_gen.ico_app)


        self.strvar_combobox_proceso = tk.StringVar()
        self.strvar_access_textbox_bbddd_1 = tk.StringVar()
        self.strvar_access_textbox_bbddd_2 = tk.StringVar()
        self.strvar_sql_server_servidor_1 = tk.StringVar()
        self.strvar_sql_server_servidor_2 = tk.StringVar()
        self.strvar_sql_server_bbdd_1 = tk.StringVar()
        self.strvar_sql_server_bbdd_2 = tk.StringVar()

    



        ###############################################
        #       widgets PROCESO
        ###############################################

        lista_GUI_procesos = [mod_gen.dicc_procesos[i]["PROCESO"] for i in mod_gen.dicc_procesos.keys()]

        self.label_proceso = gui_widgets(master, tipo_widget = "label", text = "PROCESO", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 20))

        self.combobox_proceso = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_combobox_proceso, width = 25, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = lista_GUI_procesos, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                combobox_bind_comboboxselected_rutina = self.def_GUI_combobox_proceso, place = (100, 20)))

        self.boton_proceso = (gui_widgets(master, tipo_widget = "button", text = "CHECK", width = 8, height = 0, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                            command_proceso = self.def_GUI_threads, place = (340, 18)))



        ###############################################
        #       widgets MS Access
        ###############################################

        self.label_access = gui_widgets(master, tipo_widget = "label", text = "MS ACCESS", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 12, "bold", "underline"), place = (20, 60))

        self.label_access_bbdd_1 = gui_widgets(master, tipo_widget = "label", text = "BBDD_01", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 90))

        self.label_access_bbdd_path_1 = (gui_widgets(master, tipo_widget = "label", width = 85, height = 1, text = "", fg = "black", 
                                                        font = ("Calibri", 10, "bold", "italic"), justify = tk.LEFT, borderwidth = 1, relief = "groove", 
                                                        textvariable = self.strvar_access_textbox_bbddd_1, anchor="w", place = (100, 90)))

        self.label_access_bbdd_2 = (gui_widgets(master, tipo_widget = "label", text = "BBDD_02", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 120)))

        self.label_access_bbdd_path_2 = (gui_widgets(master, tipo_widget = "label", width = 85, height = 1, text = "", 
                                                        fg = "black", font = ("Calibri", 10, "bold", "italic"), justify = tk.LEFT, 
                                                        borderwidth = 1, relief = "groove", textvariable = self.strvar_access_textbox_bbddd_2, anchor="w", 
                                                        place = (100, 120)))

        self.boton_access_addfile_1 = (gui_widgets(master, tipo_widget = "button", text = "Add", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "ADD_ACCESS_BBDD_01", place = (720, 88)))

        self.boton_access_clearfile_1 = (gui_widgets(master, tipo_widget = "button", text = "Clear", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "CLEAR_ACCESS_BBDD_01", place = (770, 88)))

        boton_access_addfile_2 = (gui_widgets(master, tipo_widget = "button", text = "Add", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "ADD_ACCESS_BBDD_02", place = (720, 118)))

        boton_access_clearfile_2 = (gui_widgets(master, tipo_widget = "button", text = "Clear", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "CLEAR_ACCESS_BBDD_02", place = (770, 118)))




        ###############################################
        #       widgets SQL Server
        ###############################################

        self.label_sql_server = (gui_widgets(master, tipo_widget = "label", text = "SQL SERVER", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 12, "bold", "underline"), place = (20, 160)))

        self.label_sql_server_servidor_1 = (gui_widgets(master, tipo_widget = "label", text = "Servidor 1", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 190)))

        self.label_sql_server_bbdd_1 = (gui_widgets(master, tipo_widget = "label", text = "SQL_BBDD_01", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (360, 190)))

        self.label_sql_server_servidor_2 = (gui_widgets(master, tipo_widget = "label", text = "Servidor 2", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 220)))
        
        self.label_sql_server_servidor_2 = (gui_widgets(master, tipo_widget = "label", text = "SQL_BBDD_02", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (360, 220)))



    
        self.combobox_sql_server_servidor_1 = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_sql_server_servidor_1, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = mod_sql_server.lista_GUI_sql_server_servidor, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                combobox_bind_comboboxselected_rutina = self.def_GUI_combobox_sql_server_servidor_1, 
                                                place = (100, 190)))

        self.combobox_sql_server_bbdd_1 = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_sql_server_bbdd_1, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = [""], combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                place = (460, 190)))


        self.combobox_sql_server_servidor_2 = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_sql_server_servidor_2, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = mod_sql_server.lista_GUI_sql_server_servidor, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                combobox_bind_comboboxselected_rutina = self.def_GUI_combobox_sql_server_servidor_2, 
                                                place = (100, 220)))

        self.combobox_sql_server_bbdd_2 = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_sql_server_bbdd_2, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = [""], combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                place = (460, 220)))



        self.boton_sql_server_clear_1 = (gui_widgets(master, tipo_widget = "button", text = "Clear", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "CLEAR_SQL_SERVER_BBDD_01", place = (720, 190)))


        self.boton_sql_server_clear_2 = (gui_widgets(master, tipo_widget = "button", text = "Clear", width = 5, fg = "black", font = ("Calibri", 8, "bold"),
                                            command = self.def_GUI_access_add_clear, command_parametro = "CLEAR_SQL_SERVER_BBDD_02", place = (720, 220)))



        ###############################################
        #       widgets comentarios
        ###############################################

        self.scrolledtext_comentarios = (gui_widgets(master, tipo_widget = "scrolledtext", 
                                                    width = 97, height = 10, wrap = tk.WORD, fg = "black", place = (20, 270)))

        self.scrolledtext_comentarios.config_lock_unlock("BLOQUEADO")


        ##############################################################
        #       widgets WARNING RESOLUCION PANTALLA NO ES LA RECOMENDADA
        ##############################################################

        if es_resolucion_pantalla_recomendada == "NO":
            string_label_warning_resolucion_pantalla = "SE RECOMIENDA USAR UNA CONFIGURACIÓN DE PANTALLA DE " + resolucion_pantalla_recomendada + " (LA TUYA ES " + resolucion_pantalla + ")"

            self.label_warning_resolucion_pantalla = (gui_widgets(master, tipo_widget = "label", text = string_label_warning_resolucion_pantalla, bg = "red", fg = "white", 
                                                                  font = ("Calibri", 11, "bold"), place = (110, 460)))





    #################################################################################################################################################################################
    ##                     RUTINAS
    #################################################################################################################################################################################

    def def_GUI_threads(self):
        #rutina para ejecutar el proceso de control de versiones o el diagnostico de dependencias en MS Access (el de SQL Server es mediante otro toplevel intermedio)
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_gen.global_proceso_en_ejecucion == "SI":
            messagebox.showerror(title = mod_gen.nombre_app, message = "Espera a que acabe el proceso actualmente en ejecución.")

        else:
            Thread(target = self.def_GUI_boton_check).start()



    def def_GUI_combobox_proceso(self, event):
        #rutina de evento (asociada al metodo bind) que permite según el proceso seleccionado
        #actualizar el scrolledtext con la descripción del proceso

        proceso_selecc = self.strvar_combobox_proceso.get()
        proceso_selecc_id = next((key for key, value in mod_gen.dicc_procesos.items() if value["PROCESO"] == proceso_selecc), None)

        lista_comentario = mod_gen.dicc_procesos[proceso_selecc_id]["COMENTARIO"]

        self.scrolledtext_comentarios.config_lock_unlock("DESBLOQUEADO")
        self.scrolledtext_comentarios.scrolledtext_delete("1.0",  tk.END)

        for item in lista_comentario:
            self.scrolledtext_comentarios.scrolledtext_insert(tk.END, item)

        self.scrolledtext_comentarios.config_lock_unlock("BLOQUEADO")



    def def_GUI_access_add_clear(self, opcion):
        #rutina que permite añdir o borrar (según el parametro opcion) en la GUI los valores de las bbdd MS Access seleccionadas
        #o los servidores + bbdd SQL Server seleccionados 

        bbdd = opcion[len(opcion) - 7:]# - 7 es la longitud de los string BBDD_01 y BBDD_02

        if opcion in ["ADD_ACCESS_BBDD_01", "ADD_ACCESS_BBDD_02"]:

            msg = messagebox.askokcancel(mod_gen.nombre_app, message = "Selecciona la ubicación de " + bbdd + ".")

            if msg == True:
                path_bbdd = fd.askopenfilename(parent = self.master, title = "", filetypes = mod_access.lista_GUI_askopenfilename_ms_access)

                if bbdd == "BBDD_01":
                    self.strvar_access_textbox_bbddd_1.set(path_bbdd)

                elif bbdd == "BBDD_02":
                    self.strvar_access_textbox_bbddd_2.set(path_bbdd)

                mod_gen.dicc_codigos_bbdd[bbdd]["MS_ACCESS"]["PATH_BBDD"] = path_bbdd


        elif opcion == "CLEAR_ACCESS_BBDD_01":         
            self.strvar_access_textbox_bbddd_1.set("")
            mod_gen.dicc_codigos_bbdd[bbdd]["MS_ACCESS"]["PATH_BBDD"] = None

        elif opcion == "CLEAR_ACCESS_BBDD_02":         
            self.strvar_access_textbox_bbddd_2.set("")
            mod_gen.dicc_codigos_bbdd[bbdd]["MS_ACCESS"]["PATH_BBDD"] = None



        elif opcion == "CLEAR_SQL_SERVER_BBDD_01":         
            self.strvar_sql_server_servidor_1.set("")
            self.strvar_sql_server_bbdd_1.set("")
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["SERVIDOR"] = None
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["BBDD"] = None
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["CONNECTING_STRING"] = None

        elif opcion == "CLEAR_SQL_SERVER_BBDD_02":         
            self.strvar_sql_server_servidor_2.set("")
            self.strvar_sql_server_bbdd_2.set("")
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["SERVIDOR"] = None
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["BBDD"] = None
            mod_gen.dicc_codigos_bbdd[bbdd]["SQL_SERVER"]["CONNECTING_STRING"] = None




    def def_GUI_combobox_sql_server_servidor_1(self, *args):
        #rutina que permite cuando se informa el servidor 1 comprobar si el usuario tiene acceso a las bbdd del mismo
        #(en caso de que no se borra el servidor 1 seleccionado)
        

        servidor_selecc = self.strvar_sql_server_servidor_1.get()


        #se calcula la connecting string probando conectar 1ero por windows authentication
        #si funciona la conexion se almacena en mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONNECTING_STRING"]
        #si falla se pasa por SQL Server autentication abriendo un toplevel para informar el login y el password
        if mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "WINDOWS_AUTHENTICATION":

            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONNECTING_STRING"] = mod_sql_server.conn_str_sql_server_windows_authentication#aqui es BBDD_01

            #se localizan los permisos
            mod_sql_server.def_sql_server_servidor_permisos("BBDD_01", servidor_selecc)

            if mod_sql_server.global_acceso_servidor_selecc == "NO":
                messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes acceso al servidor seleccionado.")

            else:
                if not isinstance(mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo, list):
                    messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes permiso de acceso al código de los objetos de ninguna de las bbdd del servidor seleccionado.")
            
                else:
                    self.combobox_sql_server_bbdd_1.combobox_update_lista_valores(combobox_lista_valores = mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo)


        elif mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "SQL_SERVER_AUTHENTICATION":

            self.toplevel_sql_server_authentication = tk.Toplevel(self.master)
            self.toplevel_sql_server_authentication.transient(self.master)
            self.toplevel_sql_server_authentication.grab_set()

            call_gui_sql_server_authentication = (gui_sql_server_authentication(self.toplevel_sql_server_authentication, 
                                                                                combobox_sql_server_bbdd = self.combobox_sql_server_bbdd_1,
                                                                                opcion_bbdd = "BBDD_01", servidor_sql_server = servidor_selecc))#aqui es BBDD_01



    def def_GUI_combobox_sql_server_servidor_2(self, *args):
        #rutina que permite cuando se informa el servidor 2 comprobar si el usuario tiene acceso a las bbdd del mismo
        #(en caso de que no se borra el servidor 2 seleccionado)

        servidor_selecc = self.strvar_sql_server_servidor_2.get()


        #se calcula la connecting string probando conectar 1ero por windows authentication
        #si funciona la conexion se almacena en mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONNECTING_STRING"]
        #si falla se pasa por SQL Server autentication abriendo un toplevel para informar el login y el password
        if mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "WINDOWS_AUTHENTICATION":

            mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["CONNECTING_STRING"] = mod_sql_server.conn_str_sql_server_windows_authentication#aqui es BBDD_02

            #se localizan los permisos
            mod_sql_server.def_sql_server_servidor_permisos("BBDD_02", servidor_selecc)

            if mod_sql_server.global_acceso_servidor_selecc == "NO":
                messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes acceso al servidor seleccionado.")

            else:
                if not isinstance(mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo, list):
                    messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes permiso de acceso al código de los objetos de ninguna de las bbdd del servidor seleccionado.")

                else:
                    self.combobox_sql_server_bbdd_2.combobox_update_lista_valores(combobox_lista_valores = mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo)



        elif mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "SQL_SERVER_AUTHENTICATION":

            self.toplevel_sql_server_authentication = tk.Toplevel(self.master)
            self.toplevel_sql_server_authentication.transient(self.master)
            self.toplevel_sql_server_authentication.grab_set()

            call_gui_sql_server_authentication = (gui_sql_server_authentication(self.toplevel_sql_server_authentication, 
                                                                                combobox_sql_server_bbdd = self.combobox_sql_server_bbdd_2,
                                                                                opcion_bbdd = "BBDD_02", servidor_sql_server = servidor_selecc))#aqui es BBDD_02




    def def_GUI_boton_check(self):
        #rutina que permite ejecutar los procesos de control de versiones o diagnostico MS Access
        #En el caso del diagnostico en SQL Server genera un toplevel donde realizar la configuración


        proceso_selecc = self.strvar_combobox_proceso.get()
        path_bbdd_access_1 = self.strvar_access_textbox_bbddd_1.get()
        path_bbdd_access_2 = self.strvar_access_textbox_bbddd_2.get()

        servidor_sql_server_1 = self.strvar_sql_server_servidor_1.get()
        bbdd_sql_server_1 = self.strvar_sql_server_bbdd_1.get()
        servidor_sql_server_2 = self.strvar_sql_server_servidor_2.get()
        bbdd_sql_server_2 = self.strvar_sql_server_bbdd_2.get()



        mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["PATH_BBDD"] = path_bbdd_access_1 if len(path_bbdd_access_1) != 0 else None
        mod_gen.dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"] = path_bbdd_access_2 if len(path_bbdd_access_2) != 0 else None

        mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] = servidor_sql_server_1 if len(servidor_sql_server_1) != 0 else None
        mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"] = servidor_sql_server_2 if len(servidor_sql_server_2) != 0 else None
        mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"] = bbdd_sql_server_1 if len(bbdd_sql_server_1) != 0 else None
        mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"] = bbdd_sql_server_2 if len(bbdd_sql_server_2) != 0 else None


        proceso_selecc_id = next((key for key, value in mod_gen.dicc_procesos.items() if value["PROCESO"] == proceso_selecc), None)


        #se comprueba si se pueden realizar los procesos
        check_control_versiones_access = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS")
        check_control_versiones_sql_server = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "SQL_SERVER")
        check_diagnostico_access = mod_gen.func_se_puede_ejecutar_proceso("DIAGNOSTICO", "MS_ACCESS")
        check_diagnostico_sql_server = mod_gen.func_se_puede_ejecutar_proceso("DIAGNOSTICO", "SQL_SERVER")


        #empieza el proceso
        if len(proceso_selecc) == 0:
            messagebox.showerror(mod_gen.nombre_app, message = "No has seleccionado ningún proceso.")

        else:

            if proceso_selecc_id == "PROCESO_01":
                #control de versiones

                check_access_informado = "SI" if len(path_bbdd_access_1) + len(path_bbdd_access_2) != 0 else "NO"
                check_sql_server_informado = "SI" if len(servidor_sql_server_1) + len(bbdd_sql_server_1) + len(servidor_sql_server_2) + len(bbdd_sql_server_2) != 0 else "NO"


                if check_control_versiones_access == "NO" and check_control_versiones_sql_server == "NO":
                    mensaje = "No se ejecutara el proceso:\n\nMS ACCESS: las rutas configuradas han de ser distintas.\n\nSQL SERVER: [Servidor 1 + BBDD_01] ha de ser distinto a [Servidor 2 + BBDD_02]."            
                    messagebox.showerror(mod_gen.nombre_app, message = mensaje)

                else:
                    if check_control_versiones_access == "SI" and check_control_versiones_sql_server == "NO":
                        mensaje1 = "Se ejecutara el proceso sobre las 2 bbdd MS ACCESS seleccionadas.\n\nLa BBDD_02 es por defecto en la cual se hace el MERGE.\n\n"
                        mensaje2 = "SQL SERVER: no se ejecutara el proceso porque [Servidor 1 + BBDD_01] ha de ser distinto a [Servidor 2 + BBDD_02].\n\n" if check_sql_server_informado == "SI" else ""
                        mensaje3 = "Deseas continuar?"
                        mensaje = mensaje1 + mensaje2 + mensaje3

                    elif check_control_versiones_access == "NO" and check_control_versiones_sql_server == "SI":
                        mensaje1 = "Se ejecutara el proceso sobre las 2 bbdd SQL SERVER seleccionadas.\n\nLa BBDD_02 es por defecto en la cual se hace el MERGE.\n\n"
                        mensaje2 = "MS ACCESS: no se ejcutara el proceso porque las rutas configuradas han de ser distintas.\n\n" if check_access_informado == "SI" else ""
                        mensaje3 = "Deseas continuar?"
                        mensaje = mensaje1 + mensaje2 + mensaje3

                    if check_control_versiones_access == "SI" and check_control_versiones_sql_server == "SI":
                        mensaje1 = "Se ejecutara el proceso sobre las 2 bbdd MS Access seleccionadas y sobre las 2 bbdd SQL Server seleccionadas.\n\n"
                        mensaje2 = "La BBDD_02 en los 2 casos es por defecto en la cual se hace el MERGE.\n\nDeseas continuar?"
                        mensaje = mensaje1 + mensaje2


                    msg = messagebox.askokcancel(mod_gen.nombre_app, message = mensaje)

                    if msg == True:
                        mensaje = "SELECCIONA DONDE QUIERES GUARDAR LOS LOGS DE ERRORES (si los hubiese):"
                        ruta_destino_logs = fd.askdirectory(parent = root, title = mensaje)

                        self.master.config(cursor = "wait")
                        mod_gen.def_calc_global(proceso_selecc_id, ruta_destino_logs)
                        self.master.config(cursor = "")

                        if len(mod_gen.global_msg_errores_proceso_access) != 0 or len(mod_gen.global_msg_errores_proceso_sql_server) != 0:
                            messagebox.showerror(mod_gen.nombre_app, message = mod_gen.global_msg_errores_proceso_access + mod_gen.global_msg_errores_proceso_sql_server)

                        else:
                            self.toplevel_control_versiones = tk.Toplevel(self.master)
                            self.toplevel_control_versiones.transient(self.master)
                            self.toplevel_control_versiones.grab_set()

                            call_gui_ventana_control_versiones = gui_ventana_control_versiones(self.toplevel_control_versiones)



            elif proceso_selecc_id == "PROCESO_02":
                #diagnostico dependencias access

                if check_diagnostico_access == "NO":
                    messagebox.showerror(mod_gen.nombre_app, message = "Tienes que seleccionar la BBDD_01 de MS Access.")

                else:
                    mensaje1 = "Se realizara el diagnostico sobre la BBDD MS Access:\n\n" + path_bbdd_access_1 + "\n\n"
                    mensaje2 = "Tendras que seleccionar en que ruta quieres guardar el excel resultante.\n\nDeseas continuar?"
                    mensaje = mensaje1 + mensaje2
                    msg = messagebox.askokcancel(mod_gen.nombre_app, message = mensaje)

                    if msg == True:
                        mensaje = "SELECCIONA DONDE QUIERES GUARDAR EL EXCEL DE DIAGNOSTICO DE DEPENDENCIAS Y LOS LOGS DE ERRORES (si los hubiese):"
                        ruta_destino_output = fd.askdirectory(parent = root, title = mensaje)

                        self.master.config(cursor = "wait")
                        mod_gen.def_calc_global(proceso_selecc_id, ruta_destino_output, ruta_destino_excel_diagnostico_access = ruta_destino_output)
                        self.master.config(cursor = "")


                        if len(mod_gen.global_msg_errores_proceso_access) != 0:
                            messagebox.showerror(mod_gen.nombre_app, message = mod_gen.global_msg_errores_proceso_access)
                        else:
                            messagebox.showinfo(mod_gen.nombre_app, message = "Proceso finalizado.")



            elif proceso_selecc_id == "PROCESO_03":
                #diagnostico dependencias sql server

                if check_diagnostico_sql_server == "NO":
                    messagebox.showerror(mod_gen.nombre_app, message = "Tienes que seleccionar el servidor SQL SERVER de la BBDD_01.")

                else:
                    mensaje = "Se abrira una nueva ventana donde podras configurar el proceso.\n\nDeseas continuar?"
                    msg = messagebox.askokcancel(mod_gen.nombre_app, message = mensaje)

                    if msg == True:

                        self.toplevel_gui_diagnostico_bbdd_sql_server = tk.Toplevel(self.master)
                        self.toplevel_gui_diagnostico_bbdd_sql_server.transient(self.master)
                        self.toplevel_gui_diagnostico_bbdd_sql_server.grab_set()

                        call_gui_diagnostico_bbdd_sql_server = gui_diagnostico_bbdd_sql_server(self.toplevel_gui_diagnostico_bbdd_sql_server, servidor_sql_server = servidor_sql_server_1)



#################################################################################################################################################################################
##                    CLASE PARA ACCESO A SERVIDOR SQL SERVER (SQL SERVER AUTHENTICATION)
#################################################################################################################################################################################


class gui_sql_server_authentication:
    #clase que permite realizar la conexion al servidor SQL Server por SQL Server authentication
    #y almacenar la connecting string en dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"]

    def __init__(self, master, **kwargs):

        self.opcion_bbdd = kwargs.get("opcion_bbdd", None)
        servidor_sql_server = kwargs.get("servidor_sql_server", None)
        combobox_sql_server_bbdd = kwargs.get("combobox_sql_server_bbdd", None)


        self.master = master
        self.master.title("ACCESO SQL SERVER - " + self.opcion_bbdd)
        self.master.configure(bg = bg_GUI_inicio)
        self.master.geometry("400x120")
        self.master.resizable(0, 0)
        
        self.master.iconbitmap(mod_gen.ico_app)


        self.strvar_servidor = tk.StringVar()
        self.strvar_login = tk.StringVar()
        self.strvar_password = tk.StringVar()


        ###############################################
        #       widgets SQL SERVER AUTHENTICATION
        ###############################################
        
        self.label_servidor = gui_widgets(master, tipo_widget = "label", text = "SERVIDOR", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 20))

        self.textbox_servidor = (gui_widgets(master, tipo_widget = "entry", textvariable = self.strvar_servidor, justify = tk.CENTER, width = 30, bg = "white", fg = "black", 
                                             font = ("Calibri", 9, "bold"), place = (100, 20)))
        
        self.strvar_servidor.set(servidor_sql_server)
        self.textbox_servidor.config_lock_unlock("BLOQUEADO")


        self.label_login = gui_widgets(master, tipo_widget = "label", text = "LOGIN", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 50))

        self.textbox_login = (gui_widgets(master, tipo_widget = "entry", textvariable = self.strvar_login, justify = tk.CENTER, width = 30, bg = "white", fg = "black", 
                                             font = ("Calibri", 9, "bold"), place = (100, 50)))
        

        self.label_password = gui_widgets(master, tipo_widget = "label", text = "PASSWORD", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), show = "*", place = (20, 80))

        self.textbox_password = (gui_widgets(master, tipo_widget = "entry", textvariable = self.strvar_password, justify = tk.CENTER, width = 30, bg = "white", fg = "black", 
                                             font = ("Calibri", 9, "bold"), place = (100, 80)))
        


        self.boton_conexion = (gui_widgets(master, tipo_widget = "button", text = "CONEXIÓN", width = 10, height = 0, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                            command = self.def_GUI_conexion_servidor_sql_server, command_parametro = combobox_sql_server_bbdd, place = (300, 50)))
        

    ##########################################################################################################
    ##                     RUTINA
    ##########################################################################################################


    def def_GUI_conexion_servidor_sql_server(self, combobox_sql_server_bbdd):
        #rutina que permite (cuando la conexión a SQL Server es por SQL Server authentication es decir con login y password)
        #almacenar la connecting string SQL Server en dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"]
        
        login_selecc = self.strvar_login.get()
        password_selecc = self.strvar_password.get()
        servidor_selecc = self.strvar_servidor.get()

        if len(login_selecc) == 0 or len(password_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El login y el password son obligatorios.")

        elif len(login_selecc) != 0 and len(password_selecc) != 0:

            conn_string = mod_sql_server.conn_str_sql_server_login_password_authentication.replace("REEMPLAZA_LOGIN", login_selecc).replace("REEMPLAZA_PASSWORD", password_selecc)
            mod_gen.dicc_codigos_bbdd[self.opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"] = conn_string

            self.master.config(cursor = "wait")
            mod_sql_server.def_sql_server_servidor_permisos(self.opcion_bbdd, servidor_selecc)
            self.master.config(cursor = "")

            if mod_sql_server.global_acceso_servidor_selecc == "NO":
                messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes conexión al servidor seleccionado o el login / password son incorrectos.")
            else:

                combobox_sql_server_bbdd.combobox_update_lista_valores(combobox_lista_valores = mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo)

                self.master.destroy()


#################################################################################################################################################################################
##                    CLASE PARA LA GUI DE DIAGNOSTICO BBDD's SQL SERVER
#################################################################################################################################################################################

class gui_diagnostico_bbdd_sql_server:
    #clase que permite generar la GUI para la configuración y ejecución del proceso de diagnostico de dependencias en SQL Server con sus widgets y rutinas asociadas

    def __init__(self, master, **kwargs):

        self.master = master
        self.master.title(mod_gen.nombre_app)
        self.master.configure(bg = bg_GUI_inicio)
        self.master.geometry("290x300")
        self.master.resizable(0, 0)
        
        self.master.iconbitmap(mod_gen.ico_app)

        servidor_sql_server = kwargs.get("servidor_sql_server", None)


        self.strvar_sql_server_diagnostico_combobox_opciones = tk.StringVar()
        self.strvar_sql_server_diagnostico_listbox_bbdd = tk.StringVar()


        ###############################################
        #       widgets DIAGNOSTICO
        ###############################################


        self.label_sql_server_diagnostico_combobox_opciones = (gui_widgets(master, tipo_widget = "label", text = "SELECCIÓN", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 20)))

        self.combobox_sql_server_diagnostico = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_sql_server_diagnostico_combobox_opciones, width = 20, font = ("Calibri", 10, "bold"),
                                                        combobox_lista_valores = mod_sql_server.lista_GUI_diagnostico_combobox_sql_server, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                        place = (100, 20)))


        self.label_sql_server_diagnostico_listbox_bbdd = (gui_widgets(master, tipo_widget = "label", text = "BASES DE DATOS", bg = bg_GUI_inicio, fg = "black", font = ("Calibri", 11, "bold"), place = (20, 50)))

        self.boton_listbox_sql_server_all_none = (gui_widgets(master, tipo_widget = "button", text = "All / None", width = 8, height = 1, bg = "grey", fg = "white", font = ("Calibri", 7, "bold"),
                                                command_proceso = lambda: self.def_GUI_sql_server_diagnostico_listbox_all_none(), place = (210, 50)))


        self.listbox_sql_server_diagnostico = tk.Listbox(master, listvariable = self.strvar_sql_server_diagnostico_listbox_bbdd, selectmode = "multiple", width = 40, height = 10)
        self.listbox_sql_server_diagnostico.place(x = 20, y = 75)
        self.listbox_sql_server_diagnostico.configure(exportselection = False)


        self.boton_sql_server_diagnostico = (gui_widgets(master, tipo_widget = "button", text = "CHECK", width = 8, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                            command = self.def_GUI_sql_server_diagnostico_threads, command_parametro = servidor_sql_server, place = (110, 260)))




        #se crea la lista de bbdd asociadas al servidor
        mod_sql_server.def_sql_server_servidor_permisos("BBDD_01", servidor_sql_server)
        mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo



        #se crean las opciones del listbox con las bbdd (con permisos de acceso al codigo) del servidor 1 configurado en la GUI de inicio
        if isinstance(mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo, list):

            self.listbox_sql_server_diagnostico.delete(0, tk.END)
            for value in mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo:
                self.listbox_sql_server_diagnostico.insert(tk.END, value)



    ####################################################################
    ##                    RUTINAS
    ####################################################################

    def def_GUI_sql_server_diagnostico_threads(self, servidor_sql_server):
        #rutina para ejecutar el proceso de diagnostico de dependencias en SQL Server
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_gen.global_proceso_en_ejecucion == "SI":
            messagebox.showerror(title = mod_gen.nombre_app, message = "Espera a que acabe el proceso actualmente en ejecución.")

        else:
            Thread(target = self.def_GUI_sql_server_diagnostico_boton_check, args = (servidor_sql_server,)).start()



    def def_GUI_sql_server_diagnostico_listbox_all_none(self):
        #rutina para seleccionar o des-seleccionar las bbdd del servidor que entran en el calculo del proceso de diagnostico SQL Server

        lista_bbdd_selecc = [self.listbox_sql_server_diagnostico.get(i) for i in self.listbox_sql_server_diagnostico.curselection()]
        
        #seleccionar todo
        if len(lista_bbdd_selecc) == 0:
            self.listbox_sql_server_diagnostico.selection_set(0, tk.END)

        #des-seleccionar todo
        elif len(lista_bbdd_selecc) != 0:
            self.listbox_sql_server_diagnostico.selection_clear(0, tk.END)



    def def_GUI_sql_server_diagnostico_boton_check(self, servidor_sql_server):
        #rutina para ejecutar el proceso de diagnostico en SQL Server (PROCESO_03)

        opcion_diagnostico_sql_server = self.strvar_sql_server_diagnostico_combobox_opciones.get()
        lista_bbdd_selecc = [self.listbox_sql_server_diagnostico.get(i) for i in self.listbox_sql_server_diagnostico.curselection()]

   
        if len(opcion_diagnostico_sql_server) == 0 or len(lista_bbdd_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de selección y la(s) bbdd son obligatorios.")

        else:
            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] = servidor_sql_server
            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"] = lista_bbdd_selecc

            #opcion diagnostico
            if opcion_diagnostico_sql_server == mod_sql_server.lista_GUI_diagnostico_combobox_sql_server[0]:
                mensaje_1 = "SELECCIONA DONDE QUIERES GUARDAR EL EXCEL DE DIAGNOSTICO DE DEPENDENCIAS Y LOS LOGS DE ERRORES (si los hubiese):"
                mensaje_2 = "Diagnostico de dependencias SQL Server descargado en Excel en la ruta indicada."

            #opcion descarga codigo
            elif opcion_diagnostico_sql_server == mod_sql_server.lista_GUI_diagnostico_combobox_sql_server[1]:
                mensaje_1 = "SELECCIONA DONDE QUIERES GUARDAR LOS CODIGOS T-SQL DE LOS OBJETOS Y LOS LOGS DE ERRORES (si los hubiese):"
                mensaje_2 = "Códigos T-SQL de los objetos descargados en ficheros .sql en la ruta indicada."

            ruta_destino_diagnostico_sql_server = fd.askdirectory(parent = root, title = mensaje_1)

            self.master.config(cursor = "wait")
            mod_gen.def_calc_global("PROCESO_03", ruta_destino_diagnostico_sql_server
                                                                                    , opcion_diagnostico_sql_server = opcion_diagnostico_sql_server
                                                                                    , ruta_destino_diagnostico_sql_server = ruta_destino_diagnostico_sql_server)
                                                                                    
            self.master.config(cursor = "")


            if len(mod_gen.global_msg_errores_proceso_sql_server) != 0:
                messagebox.showerror(mod_gen.nombre_app, message = mod_gen.global_msg_errores_proceso_sql_server)

            else:
                messagebox.showinfo(mod_gen.nombre_app, message = mensaje_2)




#################################################################################################################################################################################
##                    CLASE PARA LA GUI DE CONTROL DE VERSIONES
#################################################################################################################################################################################

class gui_ventana_control_versiones:
    #clase que permite generar la GUI de control de versiones para MS_ACCESS y SQL_SERVER con sus widgets y rutinas asociadas

    def __init__(self, master, **kwargs):

        self.master = master
        self.master.title(mod_gen.nombre_control_versiones)
        self.master.configure(bg = bg_GUI_control_versiones)
        self.master.geometry("1740x780")
        self.master.resizable(0, 0)
        
        self.master.iconbitmap(mod_gen.ico_app)


        #string_var
        self.strvar_combobox_tipo_objeto = tk.StringVar()
        self.strvar_combobox_tipo_concepto = tk.StringVar()
        self.strvar_name_bbdd_1 = tk.StringVar()
        self.strvar_name_bbdd_2 = tk.StringVar()
        self.strvar_combobox_merge_accion = tk.StringVar()

        self.strvar_combobox_tipo_objeto.trace("w", self.def_control_versiones_cambio_tipo_objeto)



        self.lista_GUI_control_versiones_subform = [["TIPO BBDD", 140], ["TIPO OBJETO", 140], ["REPOSITORIO", 200], ["OBJETO", 200], ["BBDD_01", 60], ["BBDD_02", 60]]


        #combobox de tipo de seleccion + boton
        self.label_combobox_tipo_objeto = (gui_widgets(master, tipo_widget = "label", text = "Tipo Objeto", bg = bg_GUI_control_versiones, fg = "black", 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor="w", place = (20, 20)))


        #la lista de valores del combobox varia segun que se haya configurado o no control de versiones MS Access y/o SQL Server
        hacer_control_versiones_access = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS")
        hacer_control_versiones_sql_server = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "SQL_SERVER")

        lista_temp = []
        if hacer_control_versiones_access == "SI" and hacer_control_versiones_sql_server == "SI":
            lista_temp = mod_gen.lista_GUI_seleccion_tipo_objeto_access + mod_gen.lista_GUI_seleccion_tipo_objeto_sql_server

        elif hacer_control_versiones_access == "SI" and hacer_control_versiones_sql_server == "NO":
            lista_temp = mod_gen.lista_GUI_seleccion_tipo_objeto_access

        elif hacer_control_versiones_access == "NO" and hacer_control_versiones_sql_server == "SI":
            lista_temp = mod_gen.lista_GUI_seleccion_tipo_objeto_sql_server



        self.combobox_tipo_objeto = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_combobox_tipo_objeto, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = lista_temp, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                place = (140, 20)))



        self.label_combobox_tipo_concepto = (gui_widgets(master, tipo_widget = "label", text = "Tipo Concepto", bg = bg_GUI_control_versiones, fg = "black", 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor="w", place = (20, 50)))


    




        self.combobox_tipo_concepto = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_combobox_tipo_concepto, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = mod_gen.lista_GUI_seleccion_tipo_concepto, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                place = (140, 50)))




        self.boton_tipo_seleccion = (gui_widgets(master, tipo_widget = "button", text = "VER", width = 7, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_control_versiones_click_boton_seleccion(), 
                                                        place = (400, 35)))
        

        self.boton_excel = (gui_widgets(master, tipo_widget = "button", text = "EXCEL", width = 7, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_control_versiones_click_boton_excel(), 
                                                        place = (480, 35)))



        #subform con las rutinas con cambios
        tuple_columns = tuple([f"Column{i + 1}" for i in range(len(self.lista_GUI_control_versiones_subform))])
        self.subform_control_versiones_objetos = ttk.Treeview(master, columns = tuple_columns, show="headings")
        self.subform_control_versiones_objetos["height"] = 5

        for ind, item in enumerate(self.lista_GUI_control_versiones_subform):
            texto = item[0]
            width = item[1]

            self.subform_control_versiones_objetos.heading(f"Column{ind + 1}", text = texto)
            self.subform_control_versiones_objetos.column(f"Column{ind + 1}", width = width)

        width_control_versiones_subform = sum(item[1] for item in self.lista_GUI_control_versiones_subform)

        self.subform_control_versiones_objetos.place(x = 20, y = 100, width = width_control_versiones_subform)




        ###########################################################################################################
        #                SCRIPT BBDD_01
        ###########################################################################################################

        self.label_bbdd_01 = (gui_widgets(master, tipo_widget = "label", text = "BBDD_01", bg = "black", fg = "white",
                                            font = ("Calibri", 14, "bold"), justify = tk.LEFT, anchor="w", place = (20, 240)))

        self.label_nombre_bbdd_01 = (gui_widgets(master, tipo_widget = "label", textvariable = self.strvar_name_bbdd_1, bg = bg_GUI_control_versiones, fg = "white",
                                            font = ("Calibri", 14, "bold"), justify = tk.LEFT, anchor="w", place = (130, 240)))
        

        self.frame_bbdd_01 = tk.Frame(master)
        self.frame_bbdd_01.config(bg = "black", width = 800, height = 650)
        self.frame_bbdd_01.place(x = 20, y = 270)


        self.script_bbdd_01 = scrolledtext.ScrolledText(self.frame_bbdd_01, width = 98, height = 30, wrap = tk.NONE, fg = "black", bg = bg_GUI_scrolledtext_scripts)
        self.script_bbdd_01.pack(padx = 0, pady = 0)

        self.script_bbdd_01.tag_configure("CAMBIOS_LOCALIZADOS", background = bg_GUI_lineas_control_versiones_cambios_localizados)
        self.script_bbdd_01.tag_configure("AGREGADO", background = bg_GUI_lineas_control_versiones_cambios_agregados)
        self.script_bbdd_01.tag_configure("ELIMINADO", background = bg_GUI_lineas_control_versiones_cambios_eliminados)
        


        self.horinz_scrollbar_bbdd_01 = tk.Scrollbar(master, orient=tk.HORIZONTAL, command = self.script_bbdd_01.xview)
        self.script_bbdd_01.configure(xscrollcommand = self.horinz_scrollbar_bbdd_01.set)
        self.horinz_scrollbar_bbdd_01.place(x = 770, y = 250)



        ###########################################################################################################
        #                SCRIPT BBDD_02
        ###########################################################################################################

        self.label_bbdd_02 = (gui_widgets(master, tipo_widget = "label", text = "BBDD_02", bg = "red", fg = "white",
                                            font = ("Calibri", 14, "bold"), justify = tk.LEFT, anchor="w", place = (900, 240)))

        self.label_nombre_bbdd_02 = (gui_widgets(master, tipo_widget = "label", textvariable = self.strvar_name_bbdd_2, bg = bg_GUI_control_versiones, fg = "white",
                                            font = ("Calibri", 14, "bold"), justify = tk.LEFT, anchor="w", place = (1010, 240)))


        self.frame_bbdd_02 = tk.Frame(master)
        self.frame_bbdd_02.config(bg = "black", width = 800, height = 650)
        self.frame_bbdd_02.place(x = 900, y = 270)

        self.script_bbdd_02 = scrolledtext.ScrolledText(self.frame_bbdd_02, width = 98, height = 30, wrap=tk.NONE, fg = "black", bg = bg_GUI_scrolledtext_scripts)
        self.script_bbdd_02.pack(padx = 0, pady = 0)

        self.script_bbdd_02.tag_configure("CAMBIOS_LOCALIZADOS", background = bg_GUI_lineas_control_versiones_cambios_localizados)
        self.script_bbdd_02.tag_configure("AGREGADO", background = bg_GUI_lineas_control_versiones_cambios_agregados)
        self.script_bbdd_02.tag_configure("ELIMINADO", background = bg_GUI_lineas_control_versiones_cambios_eliminados)


        self.horinz_scrollbar_bbdd_02 = tk.Scrollbar(master, orient=tk.HORIZONTAL, command = self.script_bbdd_02.xview)
        self.script_bbdd_02.configure(xscrollcommand = self.horinz_scrollbar_bbdd_02.set)
        self.horinz_scrollbar_bbdd_02.place(x = 1650, y = 250)



        ###########################################################################################################
        #                PROCESO DE MERGE
        ###########################################################################################################

        self.frame_merge = tk.Frame(master = master, width = 800, height = 130, bg = bg_GUI_control_versiones_merge)
        self.frame_merge.pack(fill = "both", expand = True)
        self.frame_merge.place(x = 900, y = 90)

        #stringvar
        self.strvar_proceso_merge_bbdd_origen = tk.StringVar()
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1 = tk.StringVar()
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2 = tk.StringVar()
        self.strvar_proceso_merge_bbdd_lineas_destino_selecc = tk.StringVar()



        self.label_proceso_merge = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "MERGE", width = 8, bg = "black", fg = "white", 
                                            font = ("Calibri", 14, "bold"), justify = tk.CENTER, place = (0, 0)))



        self.label_proceso_merge_leyenda= (gui_widgets(self.frame_merge, tipo_widget = "label", text = "LEYENDA COLORES:", bg = bg_GUI_control_versiones_merge, fg = "black", 
                                                                        font = ("Calibri", 10, "bold"), justify = tk.CENTER, place = (150, 0)))

        self.label_proceso_merge_leyenda_cambios_localiz = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Cambios localizados", bg = bg_GUI_lineas_control_versiones_cambios_localizados, fg = "black", 
                                                                        width = 22, font = ("Calibri", 10, "bold"), justify = tk.CENTER, place = (280, 0)))
        
        self.label_proceso_merge_leyenda_agregado = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Agregado en bbdd MERGE", bg = bg_GUI_lineas_control_versiones_cambios_agregados, fg = "black", 
                                                                        width = 22, font = ("Calibri", 10, "bold"), justify = tk.CENTER, place = (460, 0)))
        
        self.label_proceso_merge_leyenda_eliminado = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Eliminado en bbdd MERGE", bg = bg_GUI_lineas_control_versiones_cambios_eliminados, fg = "black", 
                                                                        width = 22, font = ("Calibri", 10, "bold"), justify = tk.CENTER, place = (640, 0)))





        self.label_proceso_merge_bbdd_origen = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "BBDD", bg = bg_GUI_control_versiones_merge, fg = "black", width = 15, 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor = "w", place = (20, 70)))
        

        
        self.combobox_proceso_merge_bbdd_selecc = (gui_widgets(self.frame_merge, tipo_widget = "combobox", textvariable = self.strvar_proceso_merge_bbdd_origen, width = 10, justify = tk.CENTER, 
                                                               font = ("Calibri", 10, "bold"), combobox_lista_valores = mod_gen.lista_GUI_seleccion_bbdd, combobox_tipo_acceso_valores = "readonly", 
                                                               combobox_bind_mousewheel = True, combobox_bind_comboboxselected_rutina = self.def_proceso_merge_combobox_bbdd_selecc, 
                                                               place = (80, 70)))



        self.label_proceso_merge_accion = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Acción", bg = bg_GUI_control_versiones_merge, fg = "black", 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor = "w", place = (200, 70)))     

        self.combobox_proceso_merge_accion = (gui_widgets(self.frame_merge, tipo_widget = "combobox", textvariable = self.strvar_combobox_merge_accion, width = 15, justify = tk.CENTER, 
                                                font = ("Calibri", 10, "bold"), combobox_lista_valores = [""], combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True,
                                                place = (270, 70)))

    

        self.label_proceso_merge_lineas_origen = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Lineas Origen", bg = bg_GUI_control_versiones_merge, fg = "black", font = ("Calibri", 12, "bold"), 
                                                              justify = tk.LEFT, anchor="w", 
                                                              place = (420, 55)))
        

        self.entry_proceso_merge_lineas_origen_1 = (gui_widgets(self.frame_merge, tipo_widget = "entry", bg = "white", width = 5, fg = "black", 
                                                              textvariable = self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1, font = ("Calibri", 10), justify = tk.CENTER, 
                                                              place = (530, 55)))

        self.entry_proceso_merge_lineas_origen_2 = (gui_widgets(self.frame_merge, tipo_widget = "entry", bg = "white", width = 5, fg = "black", 
                                                              textvariable = self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2, font = ("Calibri", 10), justify = tk.CENTER, 
                                                              place = (570, 55)))
        
        self.entry_proceso_merge_lineas_origen_1.entry_bind(self.def_GUI_proceso_merge_lineas_exit)
        self.entry_proceso_merge_lineas_origen_2.entry_bind(self.def_GUI_proceso_merge_lineas_exit)
        



        self.label_proceso_merge_lineas_destino = (gui_widgets(self.frame_merge, tipo_widget = "label", text = "Lineas Destino", bg = bg_GUI_control_versiones_merge, fg = "black", font = ("Calibri", 12, "bold"), 
                                                               justify = tk.LEFT, anchor="w", 
                                                               place = (420, 85)))


        self.entry_proceso_merge_lineas_destino = (gui_widgets(self.frame_merge, tipo_widget = "entry", bg = "white", width = 11, fg = "black", 
                                                               textvariable = self.strvar_proceso_merge_bbdd_lineas_destino_selecc, font = ("Calibri", 10), justify = tk.CENTER, 
                                                               place = (530, 85)))
        
        self.entry_proceso_merge_lineas_destino.entry_bind(self.def_GUI_proceso_merge_lineas_exit)

        

        self.boton_proceso_merge_accion = (gui_widgets(self.frame_merge, tipo_widget = "button", text = "ACCIÓN", width = 7, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_click_proceso_merge_boton_merge(), 
                                                        place = (630, 70)))



        self.boton_proceso_merge_bbdd_fisicas = (gui_widgets(self.frame_merge, tipo_widget = "button", text = "MERGE", width = 7, bg = "red", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_click_proceso_merge_boton_cambios_en_bbdd(), 
                                                        place = (720, 70)))



    #####################################################################################################################################
    #             RUTINAS CONTROL VERSIONES
    #####################################################################################################################################

    def def_control_versiones_cambio_tipo_objeto(self, *args):
        #rutina que permite actualizar en la GUI el nombre de BBDD_01 y BBDD_02 según que se seleccionen objetos MS Access o SQL Server

        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()
        self.strvar_combobox_tipo_concepto.set("")

        tipo_bbdd = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD", valor = tipo_objeto_selecc)


        if tipo_bbdd == "MS_ACCESS":

            path_bbdd_access_1 = mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["PATH_BBDD"]
            path_bbdd_access_2 = mod_gen.dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"]

            name_bbdd_access_1 = os.path.basename(path_bbdd_access_1)
            name_bbdd_access_2 = os.path.basename(path_bbdd_access_2)

            self.strvar_name_bbdd_1.set(name_bbdd_access_1)
            self.strvar_name_bbdd_2.set(name_bbdd_access_2)



        elif tipo_bbdd == "SQL_SERVER":

            bbdd_sql_server_1 = "[" + mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] + "] " + mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"]
            bbdd_sql_server_2 = "[" + mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"] + "] " + mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"]

            self.strvar_name_bbdd_1.set(bbdd_sql_server_1)
            self.strvar_name_bbdd_2.set(bbdd_sql_server_2)




    def def_control_versiones_click_boton_seleccion(self):
        #rutina que permite al pulsar el boton VER y según el tipo de objeto y de concepto seleccionado actualiza
        #el sub-formulario con los objetos con cambios de una bbdd a otra
        #se combina con la rutina def_control_versiones_update_subform_objetos (más abajo)

        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()
        tipo_concepto_selecc = self.strvar_combobox_tipo_concepto.get()

        tipo_bbdd = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD", valor = tipo_objeto_selecc)
        tipo_objeto_selecc_key = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO", valor = tipo_objeto_selecc)


        #se recupera la lista lista_control_versiones_segun_tipo_objeto_selecc de la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) de dicc_control_versiones_tipo_objeto
        #que es sobre la cual se determinan que objetos se han de exportar a excel (se exporta solo segun el valor del combobox tipo objeto, el excel ya distingue el tipo de concepto)
        lista_control_versiones_segun_tipo_objeto_selecc = mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_selecc_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]


        if len(tipo_objeto_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de objeto es obligatorio.")
        else:
            if len(tipo_concepto_selecc) == 0:
                messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de concepto es obligatorio.")

            else:

                if not isinstance(lista_control_versiones_segun_tipo_objeto_selecc, list):
                    messagebox.showerror(title = mod_gen.nombre_app, message = "No se han localizado objetos con cambios.")

                else:
                    self.script_bbdd_01.config(state = tk.NORMAL)
                    self.script_bbdd_01.delete(1.0, tk.END)
                    self.script_bbdd_01.config(state = tk.DISABLED)

                    self.script_bbdd_02.config(state = tk.NORMAL)
                    self.script_bbdd_02.delete(1.0, tk.END)
                    self.script_bbdd_02.config(state = tk.DISABLED)

                    self.def_control_versiones_update_subform_objetos(tipo_objeto_selecc, tipo_concepto_selecc, lista_control_versiones_segun_tipo_objeto_selecc)

                    self.strvar_proceso_merge_bbdd_origen.set("")
                    self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
                    self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
                    self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")
                    self.strvar_combobox_merge_accion.set("")


                    num_objetos = sum(1 if dicc["CHECK_OBJETO"] == tipo_concepto_selecc else 0 for dicc in lista_control_versiones_segun_tipo_objeto_selecc)

                    messagebox.showinfo(title = mod_gen.nombre_app, message = str(num_objetos) + " objetos localizados con cambios.")




    def def_control_versiones_click_boton_excel(self):
        #rutina que permite descargar a Excel los objetos con cambios de una bbdd a otra según el tipo de objeto seleccionado

        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()

        tipo_objeto_selecc_key = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO", valor = tipo_objeto_selecc)
        tipo_bbdd = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD", valor = tipo_objeto_selecc)


        #se recupera la lista lista_control_versiones_selecc de la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) de dicc_control_versiones_tipo_objeto
        #que es sobre la cual se determinan que objetos se han de exportar a excel
        lista_control_versiones_selecc = mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_selecc_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]



        if len(tipo_objeto_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de objeto es obligatorio.")
        else:
            if not isinstance(lista_control_versiones_selecc, list):
                messagebox.showerror(title = mod_gen.nombre_app, message = "No se han localizado objetos con cambios.")

            else:
                mensaje = "Se generara un fichero excel en la ruta que indiques que detalle todos los cambios por objetos según el filtro por tipo de objeto seleccionado.\n\nDeseas continuar?"
                msg = messagebox.askyesno(title = mod_gen.nombre_app, message = mensaje)

                if msg == True:

                    path_xls = fd.askdirectory(parent = self.master, title = "INDICA DONDE QUIERES GUARDAR EL FICHERO EXCEL:")

                    self.master.config(cursor = "wait")
                    mod_gen.def_control_versiones_export_excel(tipo_bbdd, lista_control_versiones_selecc, path_xls)
                    self.master.config(cursor = "")
                    
                    del lista_control_versiones_selecc

                    messagebox.showinfo(mod_gen.nombre_app, message = "Excel generado.")



    def def_control_versiones_update_subform_objetos(self, tipo_objeto_selecc, tipo_concepto_selecc, lista_control_versiones_segun_tipo_objeto):
        #rutina que permite al pulsar el boton VER y según el tipo de objeto y de concepto seleccionado actualiza
        #el sub-formulario con los objetos con cambios de una bbdd a otra
        #se combina con la rutina def_control_versiones_click_boton_seleccion (más arriba)


        tipo_bbdd = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD", valor = tipo_objeto_selecc)
        tipo_objeto_selecc_key = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO", valor = tipo_objeto_selecc)

        self.subform_control_versiones_objetos.bind("<ButtonRelease-1>", lambda event: self.def_control_versiones_update_subform_objetos_click_item(event))

        if len(lista_control_versiones_segun_tipo_objeto) != 0:

            if tipo_objeto_selecc_key == "TODOS":

                lista_temp = [[dicc["TIPO_BBDD"], dicc["TIPO_OBJETO_SUBFORM"], dicc["REPOSITORIO"], dicc["NOMBRE_OBJETO"], dicc["NUM_CAMBIOS_SCRIPT_1"], dicc["NUM_CAMBIOS_SCRIPT_2"]] 
                            for dicc in lista_control_versiones_segun_tipo_objeto if dicc["TIPO_BBDD"] == tipo_bbdd and dicc["CHECK_OBJETO"] == tipo_concepto_selecc]

            else:

                lista_temp = [[dicc["TIPO_BBDD"], dicc["TIPO_OBJETO_SUBFORM"], dicc["REPOSITORIO"], dicc["NOMBRE_OBJETO"], dicc["NUM_CAMBIOS_SCRIPT_1"], dicc["NUM_CAMBIOS_SCRIPT_2"]] 
                            for dicc in lista_control_versiones_segun_tipo_objeto if dicc["TIPO_BBDD"] == tipo_bbdd and dicc["TIPO_OBJETO"] == tipo_objeto_selecc_key and dicc["CHECK_OBJETO"] == tipo_concepto_selecc]


            df_temp = pd.DataFrame(lista_temp, columns = [i[0] for i in self.lista_GUI_control_versiones_subform])
            df_temp = df_temp.replace({None:"---"})

            for item in self.subform_control_versiones_objetos.get_children():
                self.subform_control_versiones_objetos.delete(item)

            for index, row in df_temp.iterrows():
                self.subform_control_versiones_objetos.insert("", "end", values= tuple([row[i[0]] for i in self.lista_GUI_control_versiones_subform]))



    def def_control_versiones_update_subform_objetos_click_item(self, event):
        #rutina que permite tras hacer click en cada elemento del sub-formulario de objetos con cambios de una bbdd a otra
        #actualiza los cuadros de scripts de BBDD_01 y BBDD_02 marcando las lineas con cambios de color VERDE
        #se combina con la rutina def_control_versiones_rellenar_scripts_scrolledtext (ver más abajo)
        #
        #se usa las variables globales (global_tipo_objeto_subform, global_repositorio_subform y global_nombre_objeto_subform)
        #para almacenar el objeto seleccionado y poder ejecutar la rutina def_proceso_merge_realizar_cambios (modulo general)
        #integrada a la rutina asociada al boton ACCION (def_click_proceso_merge_boton_merge)

        global global_tipo_objeto_subform
        global global_repositorio_subform
        global global_nombre_objeto_subform


        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()

        item_selecc = self.subform_control_versiones_objetos.focus()
        
        if item_selecc:
            item_values = self.subform_control_versiones_objetos.item(item_selecc, 'values')

            global_tipo_objeto_subform = item_values[1]
            global_repositorio_subform = item_values[2]
            global_nombre_objeto_subform = item_values[3]


            #se recuperan los df codigos actuales de BBDD_01 y BBDD_02
            dicc_proceso_merge_anteriores = mod_gen.func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_selecc, global_tipo_objeto_subform, global_repositorio_subform, global_nombre_objeto_subform)
            
            df_codigo_actual_1 = dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_1"]
            df_codigo_actual_2 = dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_2"]


            #se rellenan los scrolledtext
            self.def_control_versiones_rellenar_scripts_scrolledtext(df_codigo_actual_1, df_codigo_actual_2)
 
            #se vacian los widgets de proceso merge
            self.strvar_proceso_merge_bbdd_origen.set("")
            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
            self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")
            self.strvar_combobox_merge_accion.set("")
            messagebox.showinfo(title = mod_gen.nombre_app, message = "Scripts actualizados en pantalla.")




    def def_control_versiones_rellenar_scripts_scrolledtext(self, df_codigo_bbdd_1, df_codigo_bbdd_2):
        #rutina que permite tras hacer click en cada elemento del sub-formulario de objetos con cambios de una bbdd a otra
        #actualiza los cuadros de scripts de BBDD_01 y BBDD_02 marcando las lineas con cambios de color VERDE
        #se combina con la rutina def_control_versiones_update_subform_objetos_click_item (ver más arriba)


        #se rellena el scrolltext script_bbdd_01 con el codigo de df_codigo_bbdd_1 (se añaden numeros de linea)
        #si CONTROL_CAMBIOS es ELIMINADO se quitan numeros de linea en el script modificado pero a su derecha se mantienen lineas del codigo anterior)
        self.script_bbdd_01.config(state = tk.NORMAL)
        self.script_bbdd_01.delete(1.0, tk.END)
        self.script_bbdd_01.config(state = tk.NORMAL)

        if isinstance(df_codigo_bbdd_1, pd.DataFrame):

            df_codigo_bbdd_1["CODIGO_CON_NUM_LINEA"] = None
            df_codigo_bbdd_1["NUM_LINEA"] = df_codigo_bbdd_1.apply(lambda x: None if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" else x["NUM_LINEA"], axis = 1)

            ind_col_num_linea = df_codigo_bbdd_1.columns.get_loc("NUM_LINEA")
            ind_col_codigo = df_codigo_bbdd_1.columns.get_loc("CODIGO")
            ind_col_codigo_con_num_linea = df_codigo_bbdd_1.columns.get_loc("CODIGO_CON_NUM_LINEA")
            ind_col_control_cambios = df_codigo_bbdd_1.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")

            lambda_func_num_lineas = None
            cont = 0
            for ind in df_codigo_bbdd_1.index:

                if df_codigo_bbdd_1.iloc[ind, ind_col_control_cambios] != "ELIMINADO":
                    cont += 1

                    lambda_func_num_lineas = lambda cont: f"{cont:04}\t"
                    df_codigo_bbdd_1.iloc[ind, ind_col_num_linea] = lambda_func_num_lineas(cont)
                    df_codigo_bbdd_1.iloc[ind, ind_col_codigo_con_num_linea] = lambda_func_num_lineas(cont) + df_codigo_bbdd_1.iloc[ind, ind_col_codigo]

                else:
                    df_codigo_bbdd_1.iloc[ind, ind_col_codigo_con_num_linea] = "    \t" + df_codigo_bbdd_1.iloc[ind, ind_col_num_linea] + df_codigo_bbdd_1.iloc[ind, ind_col_codigo]                   

                self.script_bbdd_01.insert(tk.END, df_codigo_bbdd_1.iloc[ind, ind_col_codigo_con_num_linea] + '\n', df_codigo_bbdd_1.iloc[ind, ind_col_control_cambios])

            self.script_bbdd_01.config(state = tk.DISABLED)




        #se rellena el scrolltext script_bbdd_02 con el codigo de df_codigo_bbdd_2 (se añaden numeros de linea)
        #si CONTROL_CAMBIOS es ELIMINADO se quitan numeros de linea en el script modificado pero a su derecha se mantienen lineas del codigo anterior)
        self.script_bbdd_02.config(state = tk.NORMAL)
        self.script_bbdd_02.delete(1.0, tk.END)
        self.script_bbdd_02.config(state = tk.NORMAL)

        if isinstance(df_codigo_bbdd_2, pd.DataFrame):

            df_codigo_bbdd_2["CODIGO_CON_NUM_LINEA"] = None
            df_codigo_bbdd_2["NUM_LINEA"] = df_codigo_bbdd_2.apply(lambda x: None if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" else x["NUM_LINEA"], axis = 1)

            ind_col_num_linea = df_codigo_bbdd_2.columns.get_loc("NUM_LINEA")
            ind_col_codigo = df_codigo_bbdd_2.columns.get_loc("CODIGO")
            ind_col_codigo_con_num_linea = df_codigo_bbdd_2.columns.get_loc("CODIGO_CON_NUM_LINEA")
            ind_col_control_cambios = df_codigo_bbdd_2.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")

            lambda_func_num_lineas = None
            cont = 0
            for ind in df_codigo_bbdd_2.index:

                if df_codigo_bbdd_2.iloc[ind, ind_col_control_cambios] != "ELIMINADO":
                    cont += 1

                    lambda_func_num_lineas = lambda cont: f"{cont:04}\t"
                    df_codigo_bbdd_2.iloc[ind, ind_col_num_linea] = lambda_func_num_lineas(cont)
                    df_codigo_bbdd_2.iloc[ind, ind_col_codigo_con_num_linea] = lambda_func_num_lineas(cont) + df_codigo_bbdd_2.iloc[ind, ind_col_codigo]

                else:
                    df_codigo_bbdd_2.iloc[ind, ind_col_codigo_con_num_linea] = "    \t" + df_codigo_bbdd_2.iloc[ind, ind_col_num_linea] + df_codigo_bbdd_2.iloc[ind, ind_col_codigo]   
                    
                self.script_bbdd_02.insert(tk.END, df_codigo_bbdd_2.iloc[ind, ind_col_codigo_con_num_linea] + '\n', df_codigo_bbdd_2.iloc[ind, ind_col_control_cambios])

            self.script_bbdd_02.config(state = tk.DISABLED)


    #####################################################################################################################################
    #             RUTINAS PROCESO MERGE
    #####################################################################################################################################

    def def_GUI_proceso_merge_lineas_exit(self, event, widget):
        #rutina de evento (asociada al metodo bind) del proceso merge entre una bbdd y otra genera un warning y borra el contenido
        #de los textbox de lineas origen (desde y hasta) y linea destino si no se informan solo caracteres numericos

        valor_lineas = widget.get()
        
        if not valor_lineas:
            return
            
        check_lineas = "OK" if valor_lineas == re.sub(r'\D', '', valor_lineas) else "KO"
            
        if check_lineas == "KO": 
            messagebox.showerror(title = mod_gen.nombre_app, message = "Solo se aceptan números enteros positivos.")
            widget.focus_set()



    def def_proceso_merge_combobox_bbdd_selecc(self, event):
        #rutina de evento (asociada al metodo bind) del proceso merge entre una bbdd y otra que permite
        #cuando se cambia el combobox "BBDD" (BBDD_01 o BBDD_02) actualizar las opciones del combobox "Acción" asociadas a BBDD_01 o BBDD_02


        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()
        bbdd_origen_selecc = self.strvar_proceso_merge_bbdd_origen.get()

        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
        self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")

        if len(tipo_objeto_selecc) == 0:
            self.strvar_proceso_merge_bbdd_origen.set("")
            messagebox.showerror(title = mod_gen.nombre_app, message = "No has seleccionado ningún objeto.")

        else:
            if bbdd_origen_selecc == "BBDD_01":
                self.combobox_proceso_merge_accion.combobox_update_lista_valores(combobox_lista_valores = mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_1)

            elif bbdd_origen_selecc == "BBDD_02":
                self.combobox_proceso_merge_accion.combobox_update_lista_valores(combobox_lista_valores = mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_2)



    def def_click_proceso_merge_boton_merge(self):
        #rutina que permite traspasar en la GUI los cambios realizados por el usuario de un script a otro al pulsar el botón "ACCIÓN"
        #y conserver los cambios, mediante la rutina def_proceso_merge_realizar_cambios (modulo general),

        tipo_objeto_selecc = self.strvar_combobox_tipo_objeto.get()
        bbdd_origen = self.strvar_proceso_merge_bbdd_origen.get()
        tipo_accion = self.strvar_combobox_merge_accion.get()

        lineas_origen_1 = self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.get()
        lineas_origen_2 = self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.get()
        lineas_destino = self.strvar_proceso_merge_bbdd_lineas_destino_selecc.get()



        #se crea la variable lineas_origen en funcion de lo informado en lineas origen para poder usarla como parametro en la rutina
        #def_proceso_merge_realizar_cambios (modulo general)
        lineas_origen = ""
        if len(lineas_origen_1) != 0 and len(lineas_origen_2) != 0:
            #si lineas origen (hasta) es menor que lineas origen (desde) se permutan los valores
            lineas_origen = str(int(lineas_origen_1)) + "-" + str(int(lineas_origen_2)) if int(lineas_origen_1) <= int(lineas_origen_2) else str(int(lineas_origen_2)) + "-" + str(int(lineas_origen_1))
            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set(lineas_origen.split("-")[0])
            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set(lineas_origen.split("-")[1])

        elif len(lineas_origen_1) != 0 and len(lineas_origen_2) == 0:
            #si solo se informa lineas origen (desde) se completa lineas origen (hasta) con el mismo valor
            lineas_origen = str(int(lineas_origen_1)) + "-" + str(int(lineas_origen_1))
            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set(lineas_origen_1)

        elif len(lineas_origen_1) == 0 and len(lineas_origen_2) != 0:
            #si solo se informa lineas origen (hasta) se completa lineas origen (desde) con el mismo valor
            lineas_origen = str(int(lineas_origen_2)) + "-" + str(int(lineas_origen_2))



        #se localiza si ya hay merge anteriores hechos (para saber si al optar por revertir cambios hay que generar un warning o no de que hay merge anteriores o no)
        dicc_proceso_merge_anteriores = mod_gen.func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_selecc, global_tipo_objeto_subform, global_repositorio_subform, global_nombre_objeto_subform)
        lista_merge_hechos = dicc_proceso_merge_anteriores["LISTA_DICC_MERGE_HECHOS"]

        

        if len(tipo_objeto_selecc) == 0:
            self.strvar_proceso_merge_bbdd_origen.set("")
            messagebox.showerror(title = mod_gen.nombre_app, message = "No has seleccionado ningún objeto.")

        else:
            if len(bbdd_origen) == 0 or len(tipo_accion) == 0:
                messagebox.showerror(title = mod_gen.nombre_app, message = "La selección de bbdd y el tipo de acción son obligatorios.")

            else:
                if tipo_accion == mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_1[1] and (len(lineas_origen) == 0 or len(lineas_destino) == 0):#Migrar por lineas
                    messagebox.showerror(title = mod_gen.nombre_app, message = "Las lineas de origen y destino son obligatorias.")
                
                elif tipo_accion == mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_2[1] and len(lineas_origen) == 0:#Quitar por lineas
                    messagebox.showerror(title = mod_gen.nombre_app, message = "Las lineas de origen son obligatorias.")


                else:
                    check_accion_revertir = "OK"
                    if tipo_accion in mod_gen.lista_acciones_revertir:#acciones de reversion

                        #se localiza si ya hay merge hechos anteriores para saber si se puede revertir cambios
                        if isinstance(lista_merge_hechos, list):
                            check_accion_revertir = "OK"
                        else:
                            check_accion_revertir = "KO"


                    if check_accion_revertir == "KO":
                        messagebox.showerror(title = mod_gen.nombre_app, message = "REVERTIR CAMBIOS:\n\nNo se realizaron merge anteriores.")

                    else:
                        msg = messagebox.askokcancel(title = mod_gen.nombre_app, message = "Estas segur@ de realizar los cambios?")

                        if msg == True:
                            #se realizan los cambios en los df y se guarda registro de los cambios en el diccionario del objeto (seleccionado en el sub-formulario)
                            #de la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) del diccionario dicc_GUI_control_versiones_tipo_objeto (modulo general)

                            mod_gen.def_proceso_merge_realizar_cambios(tipo_accion, tipo_objeto_selecc, global_tipo_objeto_subform, global_repositorio_subform, global_nombre_objeto_subform, lineas_origen, lineas_destino)


                            #se recuperan los df codigos actuales tras los cambios de BBDD_01 y BBDD_02
                            dicc_proceso_merge_anteriores = mod_gen.func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_selecc, global_tipo_objeto_subform, global_repositorio_subform, global_nombre_objeto_subform)
                            df_codigo_actual_1 = dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_1"]
                            df_codigo_actual_2 = dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_2"]



                            #se rellenan los scrolledtext
                            self.def_control_versiones_rellenar_scripts_scrolledtext(df_codigo_actual_1, df_codigo_actual_2)



                            #se reinician los widgets
                            self.strvar_combobox_merge_accion.set("")
                            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
                            self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
                            self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")

                            messagebox.showinfo(title = mod_gen.nombre_app, message = "Cambios realizados.")




    def def_click_proceso_merge_boton_cambios_en_bbdd(self):
        #rutina que permite acceder al toplevel de merge en bbdd fisica si se han registrado cambios realizados por el usuario
        #la funcion func_dicc_control_versiones_tipo_objeto_buscar_en_dicc (modulo general) con la opcion TIPO_BBDD_REALIZAR_MERGE_BBDD_FISICAS
        #crea lista de tipos de bbdd (MS_ACCESS y/o SQL_SERVER) donde se han localizado merge realizados por el usuario
        #si la lista resultante es vacia sale un warning en la GUI avisando de que no hay cambios y no se abre la GUI de merge en bbdd fisica

        lista_tipo_bbdd_merge_bbdd_fisica = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD_REALIZAR_MERGE_BBDD_FISICAS")


        if len(lista_tipo_bbdd_merge_bbdd_fisica) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "No se han localizado merge por realizar ni en MS Access ni en SQL Server.")

        else:
            mensaje = ""
            if "MS_ACCESS" in lista_tipo_bbdd_merge_bbdd_fisica and "SQL_SERVER" in lista_tipo_bbdd_merge_bbdd_fisica:
                mensaje = "Se han localizado merge realizados tanto en MS Access como en SQL Server.\n\nDeseas continuar?"

            elif "MS_ACCESS" in lista_tipo_bbdd_merge_bbdd_fisica and "SQL_SERVER" not in lista_tipo_bbdd_merge_bbdd_fisica:
                mensaje = "Se han localizado merge realizados solo en MS Access.\n\nDeseas continuar?"

            elif "MS_ACCESS" not in lista_tipo_bbdd_merge_bbdd_fisica and "SQL_SERVER" in lista_tipo_bbdd_merge_bbdd_fisica:
                mensaje = "Se han localizado merge realizados solo en SQL Server.\n\nDeseas continuar?"

            msg = messagebox.askokcancel(title = mod_gen.nombre_app, message = mensaje)

            if msg == True:

                self.toplevel_merge_bbdd_fisicas = tk.Toplevel(self.master)
                self.toplevel_merge_bbdd_fisicas.transient(self.master)
                self.toplevel_merge_bbdd_fisicas.grab_set()

                call_gui_ventana_merge_bbdd_fisicas = gui_ventana_merge_bbdd_fisicas(self.toplevel_merge_bbdd_fisicas)



#################################################################################################################################################################################
##                    CLASE PARA LA GUI DE MERGE EN BBDD FISICAS
#################################################################################################################################################################################

class gui_ventana_merge_bbdd_fisicas:
    #clase que permite generar la GUI para ejecutar los merge en bbdd fisica con sus widgets y rutinas asociadas

    def __init__(self, master, **kwargs):

        self.master = master
        self.master.title(mod_gen.nombre_merge_bbdd_fisicas)
        self.master.configure(bg = bg_GUI_merge_bbdd_fisica)
        self.master.geometry("1130x585")
        self.master.resizable(0, 0)
        
        self.master.iconbitmap(mod_gen.ico_app)


        #string_var
        self.strvar_combobox_tipo_seleccion = tk.StringVar()


        #se calculan las listas de objetos donde realizar merge en bbdd fisica asociadas a cada opcion del combobox
        mod_gen.def_merge_bbdd_fisica_lista_objetos()


        #se calculan los ajustes manuales a realizar en access
        if mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS") == "SI":
            mod_gen.def_merge_access_ajustes_manuales()
                                                                                                       

        #se calcula la lista para el combobox de seleccion
        lista_combobox_seleccion = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_COMBOBOX_MERGE_BBDD_FISICAS")



        self.label_combobox_opcion = (gui_widgets(master, tipo_widget = "label", text = "Tipo Selección", bg = bg_GUI_merge_bbdd_fisica, fg = "black", 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor="w", place = (20, 20)))


        self.combobox_objetos = (gui_widgets(master, tipo_widget = "combobox", textvariable = self.strvar_combobox_tipo_seleccion, width = 30, font = ("Calibri", 10, "bold"),
                                                combobox_lista_valores = lista_combobox_seleccion, combobox_tipo_acceso_valores = "readonly", combobox_bind_mousewheel = True, 
                                                place = (140, 20)))


        self.boton_ver = (gui_widgets(master, tipo_widget = "button", text = "VER", width = 7, bg = "black", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_merge_bbdd_fisicas_click_boton_seleccion(), 
                                                        place = (400, 18)))


        self.boton_merge_realizar = (gui_widgets(master, tipo_widget = "button", text = "MERGE", width = 7, bg = "red", fg = "white", font = ("Calibri", 10, "bold"),
                                                       command_proceso = lambda: self.def_merge_bbdd_fisicas_click_boton_realizar(), 
                                                        place = (500, 18)))




        #subform con los objetos donde se ha hecho algún merge en la ventana de control de versiones
        self.lista_GUI_merge_bbdd_fisica_subform = [["TIPO BBDD", 140], ["TIPO OBJETO", 140], ["TIPO_REPOSITORIO", 200], ["REPOSITORIO", 200], ["OBJETO", 200], ["ESTADO MIGRACIÓN", 200]]

        self.label_subform_merge_por_realizar = (gui_widgets(master, tipo_widget = "label", text = "MERGE POR REALIZAR", bg = "black", fg = "white", 
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor="w", place = (20, 75)))

        tuple_columns = tuple([f"Column{i + 1}" for i in range(len(self.lista_GUI_merge_bbdd_fisica_subform))])
        self.subform_merge_bbdd_fisica_objetos = ttk.Treeview(master, columns = tuple_columns, show="headings")
        self.subform_merge_bbdd_fisica_objetos["height"] = 5

        for ind, item in enumerate(self.lista_GUI_merge_bbdd_fisica_subform):
            texto = item[0]
            width = item[1]

            self.subform_merge_bbdd_fisica_objetos.heading(f"Column{ind + 1}", text = texto)
            self.subform_merge_bbdd_fisica_objetos.column(f"Column{ind + 1}", width = width)

        width_merge_bbdd_fisica_subform = sum(item[1] for item in self.lista_GUI_merge_bbdd_fisica_subform)

        self.subform_merge_bbdd_fisica_objetos.place(x = 20, y = 70, width = width_merge_bbdd_fisica_subform)



        #scrrolledtext con el script con los merge realizados pdte de migrar a bbdd fisica
        self.label_script_migrar = (gui_widgets(master, tipo_widget = "label", text = "SCRIPT A MIGRAR", bg = "black", fg = "white", width = 15,
                                                            font = ("Calibri", 12, "bold"), justify = tk.LEFT, anchor="w", place = (20, 215)))

        self.frame_script_migrar = tk.Frame(master)
        self.frame_script_migrar.config(width = 1400, height = 100)
        self.frame_script_migrar.place(x = 20, y = 240)


        self.script_migrar = scrolledtext.ScrolledText(self.frame_script_migrar, width = 132, height = 20, wrap = tk.NONE, fg = "black", bg = bg_GUI_scrolledtext_scripts)
        self.script_migrar.pack(padx = 0, pady = 0)

        self.script_migrar.tag_configure("CAMBIOS_LOCALIZADOS", background = bg_GUI_lineas_control_versiones_cambios_localizados)
        self.script_migrar.tag_configure("AGREGADO", background = bg_GUI_lineas_control_versiones_cambios_agregados)
        self.script_migrar.tag_configure("ELIMINADO", background = bg_GUI_lineas_control_versiones_cambios_eliminados)
        
        self.script_migrar.config(state = tk.DISABLED)


        self.horinz_scrollbar_bbdd_01 = tk.Scrollbar(master, orient=tk.HORIZONTAL, command = self.script_migrar.xview)
        self.script_migrar.configure(xscrollcommand = self.horinz_scrollbar_bbdd_01.set)
        self.horinz_scrollbar_bbdd_01.place(x = 1045, y = 215)



    #####################################################################################################################################
    #             RUTINAS
    #####################################################################################################################################


    def def_merge_bbdd_fisicas_click_boton_seleccion(self):
        #rutina que permite tras seleccionar el tipo de selección y pulsar el botón VER
        #actualizar el sub-formulario con los objetos donde el usuario ha realizado cambios
        #se asocia con la rutina def_merge_bbdd_fisicas_update_subform_objetos (ver más abajo)

        opcion_proceso_merge = self.strvar_combobox_tipo_seleccion.get()
        lista_temp = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)


        if len(opcion_proceso_merge) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de selección es obligatorio.")
        else:

            self.script_migrar.config(state = tk.NORMAL)
            self.script_migrar.delete(1.0, tk.END)
            self.script_migrar.config(state = tk.DISABLED)

            if isinstance(lista_temp, list):

                self.def_merge_bbdd_fisicas_update_subform_objetos(opcion_proceso_merge)

                num_objetos = len(lista_temp)
                messagebox.showinfo(title = mod_gen.nombre_app, message = str(num_objetos) + " objetos localizados.")

            else:
                messagebox.showinfo(title = mod_gen.nombre_app, message = "0 objetos localizados.")




    def def_merge_bbdd_fisicas_update_subform_objetos(self, opcion_proceso_merge):
        #rutina que permite tras seleccionar el tipo de selección y pulsar el botón VER
        #actualizar el sub-formulario con los objetos donde el usuario ha realizado cambios
        #se asocia con la rutina def_merge_bbdd_fisicas_click_boton_seleccion (ver más arriba)

        self.subform_merge_bbdd_fisica_objetos.bind("<ButtonRelease-1>", lambda event: self.def_merge_bbdd_fisicas_update_subform_objetos_click_item(event, opcion_proceso_merge))

        lista_dicc_temp = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)


        if isinstance(lista_dicc_temp, list):#los errores de migracion de inicio son None

            lista_subform = []
            for dicc in lista_dicc_temp:
                tipo_bbdd = dicc["TIPO_BBDD"]
                tipo_objeto_subform = dicc["TIPO_OBJETO_SUBFORM"]
                tipo_repositorio = dicc["TIPO_REPOSITORIO"]
                repositorio = dicc["REPOSITORIO"]
                nombre_objeto = dicc["NOMBRE_OBJETO"]
                estado_migracion = dicc["ESTADO_MIGRACION"]

                lista_subform.append([tipo_bbdd, tipo_objeto_subform, tipo_repositorio, repositorio, nombre_objeto, estado_migracion])

            df_temp = pd.DataFrame(lista_subform, columns = [i[0] for i in self.lista_GUI_merge_bbdd_fisica_subform])
            df_temp = df_temp.replace({None: "---"})


            for item in self.subform_merge_bbdd_fisica_objetos.get_children():
                self.subform_merge_bbdd_fisica_objetos.delete(item)

            for index, row in df_temp.iterrows():
                self.subform_merge_bbdd_fisica_objetos.insert("", "end", values = tuple([row[i[0]] for i in self.lista_GUI_merge_bbdd_fisica_subform]))

        else:
            for item in self.subform_merge_bbdd_fisica_objetos.get_children():
                self.subform_merge_bbdd_fisica_objetos.delete(item)



    def def_merge_bbdd_fisicas_update_subform_objetos_click_item(self, event, opcion_proceso_merge):
        #rutina que permite al hacer click en el sub-formulario de objetos con cambios realizados por el usuario
        #actualizar el cuadro de script con las lineas cambiadas marcadas en el color correspondiente según la acción realizada
        #se asocia con la rutina def_merge_bbdd_fisicas_rellenar_scripts_scrolledtext (ver más abajo)

        lista_dicc_merge_bbdd_fisicas = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)

        item_selecc = self.subform_merge_bbdd_fisica_objetos.focus()
        
        if item_selecc:
            item_values = self.subform_merge_bbdd_fisica_objetos.item(item_selecc, 'values')

            tipo_bbdd_seek = item_values[0]
            tipo_objeto_subform_seek = item_values[1]
            tipo_repositorio_seek = item_values[2]
            repositorio_seek = item_values[3]
            nombre_objeto_seek = item_values[4]
            estado_migracion = item_values[5]


            for dicc in lista_dicc_merge_bbdd_fisicas:
                tipo_bbdd = dicc["TIPO_BBDD"]
                tipo_objeto_subform = dicc["TIPO_OBJETO_SUBFORM"]
                tipo_repositorio = dicc["TIPO_REPOSITORIO"]
                repositorio = dicc["REPOSITORIO"]
                nombre_objeto = dicc["NOMBRE_OBJETO"]
                df_script = dicc["DF_CODIGO"]
  
                if tipo_bbdd == "MS_ACCESS":

                    #MS_ACCESS --> caso de que no son ajustes manuales por realizar
                    if not estado_migracion == mod_gen.label_merge_access_bbdd_fisica_en_manual:

                        #MS_ACCESS (TABLA_LOCAL, VINCULO_ODBC y VINCULO_OTRO) --> no hay tipo repositorio ni repositorio
                        if mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = tipo_objeto_subform_seek) in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:

                            if tipo_bbdd_seek == tipo_bbdd and tipo_objeto_subform_seek == tipo_objeto_subform and nombre_objeto_seek == nombre_objeto:
                                df_script_scrolledtext = df_script
                                break

                        #MS_ACCESS (VARIABLES_VBA) --> no hay nombre de objeto
                        elif mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = tipo_objeto_subform_seek) == "VARIABLES_VBA":

                            if tipo_bbdd_seek == tipo_bbdd and tipo_objeto_subform_seek == tipo_objeto_subform and tipo_repositorio_seek == tipo_repositorio and repositorio_seek == repositorio:
                                df_script_scrolledtext = df_script
                                break

                        #MS_ACCESS (RUTINAS_VBA)
                        elif mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = tipo_objeto_subform_seek) == "RUTINAS_VBA":
                            if tipo_bbdd_seek == tipo_bbdd and tipo_objeto_subform_seek == tipo_objeto_subform and tipo_repositorio_seek == tipo_repositorio and repositorio_seek == repositorio and nombre_objeto_seek == nombre_objeto:
                                df_script_scrolledtext = df_script
                                break

                    #MS_ACCESS --> caso de que SI son ajustes manuales por realizar
                    elif estado_migracion == mod_gen.label_merge_access_bbdd_fisica_en_manual:
                            df_script_scrolledtext = df_script
                            break



                #SQL SERVER --> no hay tipo repositorio
                elif tipo_bbdd == "SQL_SERVER":
                    if tipo_bbdd_seek == tipo_bbdd and tipo_objeto_subform_seek == tipo_objeto_subform and repositorio_seek == repositorio and nombre_objeto_seek == nombre_objeto:
                        df_script_scrolledtext = df_script
                        break


            #se rellena el scrolledtext
            self.def_merge_bbdd_fisicas_rellenar_scripts_scrolledtext(df_script_scrolledtext)



    def def_merge_bbdd_fisicas_rellenar_scripts_scrolledtext(self, df_script):
        #rutina que permite al hacer click en el sub-formulario de objetos con cambios realizados por el usuario
        #actualizar el cuadro de script con las lineas cambiadas marcadas en el color correspondiente según la acción realizada
        #se asocia con la rutina def_merge_bbdd_fisicas_update_subform_objetos_click_item (ver más arriba)


        #se rellena el scrolltext script_migrar con el codigo de df_merge (se añaden numeros de linea)
        #si CONTROL_CAMBIOS es ELIMINADO se quitan numeros de linea en el script modificado pero a su derecha se mantienen lineas del codigo anterior)
        self.script_migrar.config(state = tk.NORMAL)
        self.script_migrar.delete(1.0, tk.END)
        self.script_migrar.config(state = tk.NORMAL)

        if isinstance(df_script, pd.DataFrame):

            df_script["CODIGO_CON_NUM_LINEA"] = None
            df_script["NUM_LINEA"] = df_script.apply(lambda x: None if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" else x["NUM_LINEA"], axis = 1)

            ind_col_num_linea = df_script.columns.get_loc("NUM_LINEA")
            ind_col_codigo = df_script.columns.get_loc("CODIGO")
            ind_col_codigo_con_num_linea = df_script.columns.get_loc("CODIGO_CON_NUM_LINEA")
            ind_col_control_cambios = df_script.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")

            lambda_func_num_lineas = None
            cont = 0
            for ind in df_script.index:

                if df_script.iloc[ind, ind_col_control_cambios] != "ELIMINADO":
                    cont += 1

                    lambda_func_num_lineas = lambda cont: f"{cont:04}\t"
                    df_script.iloc[ind, ind_col_num_linea] = lambda_func_num_lineas(cont)
                    df_script.iloc[ind, ind_col_codigo_con_num_linea] = lambda_func_num_lineas(cont) + df_script.iloc[ind, ind_col_codigo]

                else:
                    df_script.iloc[ind, ind_col_codigo_con_num_linea] = "    \t" + df_script.iloc[ind, ind_col_num_linea] + df_script.iloc[ind, ind_col_codigo]   
                    
                self.script_migrar.insert(tk.END, df_script.iloc[ind, ind_col_codigo_con_num_linea] + '\n', df_script.iloc[ind, ind_col_control_cambios])

            self.script_migrar.config(state = tk.DISABLED)



    def def_merge_bbdd_fisicas_click_boton_realizar(self):
        #rutina que permite ejecutar el merge en bbdd fisica y generar los logs de OK y los de errores (si los hubiese)
        #genera tambien la documentacion del proceso en ficheros .txt

        opcion_proceso_merge = self.strvar_combobox_tipo_seleccion.get()
        lista_dicc_objetos_migrar_bbdd_fisica = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)


        if len(opcion_proceso_merge) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "No has seleccionado ninguna opción.")

        else:
            msg = None
            tipo_bbdd = None

            #MS ACCESS --> objetos para migrar en bbdd fisica
            if opcion_proceso_merge == mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["COMBOBOX_GUI"]:

                tipo_bbdd = "MS_ACCESS"

                mensaje1 = "Se realizaran los cambios en los distintos módulos VBA de la bbdd MS Access.\n\n"
                mensaje2 = "En caso de errores de migración se generara un fichero de logs en .txt en la ruta que indiques.\n\n"
                mensaje3 = "En la ruta que indicada para los posibles logs de errores tambien se creara una carpeta con la documentación del proceso.\n\n"
                mensaje4 = "Deseas continuar?"
                mensaje = mensaje1 + mensaje2 + mensaje3 + mensaje4
                msg = messagebox.askokcancel(title = mod_gen.nombre_app, message = mensaje)


            #SQL SERVER --> objetos para migrar en bbdd fisica
            elif opcion_proceso_merge == mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["COMBOBOX_GUI"]:

                tipo_bbdd = "SQL_SERVER"

                mensaje1 = "Los cambios en bbdd fisica se realizaran en este orden:\n\n"
                mensaje2 = "1. Se crearan esquemas nuevos (si los hubiese).\n"
                mensaje3 = "2. Se crearan las tablas.\n"
                mensaje4 = "3. Se crearan las funciones.\n"
                mensaje5 = "4. Se crearan las views.\n"
                mensaje6 = "5. Se crearan los stored procedures.\n\n"
                mensaje7 = "En caso de errores de migración se generara un fichero de logs en .txt en la ruta que indiques.\n\n"
                mensaje8 = "En la ruta que indicada para los posibles logs de errores tambien se creara una carpeta con la documentación del proceso.\n\n"
                mensaje9 = "Deseas continuar?"
                mensaje = mensaje1 + mensaje2 + mensaje3 + mensaje4 + mensaje5 + mensaje6 + mensaje7 + mensaje8 + mensaje9
                msg = messagebox.askokcancel(title = mod_gen.nombre_app, message = mensaje)


            if msg == True:

                ruta_export = fd.askdirectory(parent = root, title = "RUTA DONDE GUARDAR LOS FICHEROS DE LOGS (OK + ERRORES) + LA DOCUMENTACIÓN DEL PROCESO:")

                self.master.config(cursor = "wait")

                #se reestablece mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] a None
                mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = None

                #se ejecuta el proceso de merge
                lista_dicc_objetos_migrar_bbdd_fisica = [dicc for dicc in lista_dicc_objetos_migrar_bbdd_fisica if dicc["TIPO_BBDD"] == tipo_bbdd]
                mod_gen.def_merge_bbdd_fisicas(tipo_bbdd, lista_dicc_objetos_migrar_bbdd_fisica, ruta_export)


                #se ejecutan los logs (OK + errores) y crea los ficheros (se ejecuta sobre BBDD_02 que es por defecto donde se hace el merge)
                mod_gen.def_generacion_logs("MERGE_BBDD_FISICA_LOGS_OK", tipo_bbdd, ruta_export, opcion_bbdd = "BBDD_02")
                mod_gen.def_generacion_logs("MERGE_BBDD_FISICA_LOGS_ERRORES", tipo_bbdd, ruta_export, opcion_bbdd = "BBDD_02")


                self.master.config(cursor = "")

                #se genera el messagebox final
                lista_temp_ok = mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"]
                lista_temp_errores = mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

                num_ok = len(lista_temp_ok) if isinstance(lista_temp_ok, list) else 0
                num_errores = len(lista_temp_errores) if isinstance(lista_temp_errores, list) else 0

                mensaje1 = "Merge realizado en la bbdd.\n\n"
                mensaje2 = "--> " + str(num_ok) + " objetos migrados correctamente.\n\n" if num_ok != 0 else ""
                mensaje3 = "--> " + str(num_errores) + " objetos no migrados debido a errores.\n\n" if num_errores != 0 else ""
                mensaje4 = "\n\nConsulta los ficheros de logs."
                mensaje = mensaje1 + mensaje2 + mensaje3 + mensaje4

                del lista_temp_ok
                del lista_temp_errores

                #se vacia dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"] --> LISTA_DICC_OK_MIGRACION + LISTA_DICC_ERRORES_MIGRACION
                mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"] = None
                mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = None


                if num_ok != 0 and num_errores == 0:
                    messagebox.showinfo(title = mod_gen.nombre_app, message = mensaje)

                elif num_ok != 0 and num_errores != 0:
                    messagebox.showwarning(title = mod_gen.nombre_app, message = mensaje)

                elif num_ok == 0 and num_errores != 0:
                    messagebox.showerror(title = mod_gen.nombre_app, message = mensaje)




#################################################################################################################################################################################
##                    SE INICIA EL APP
#################################################################################################################################################################################

if __name__ == "__main__":

    root = tk.Tk()
    call_gui_ventana_inicio = gui_ventana_inicio(root)
    root.mainloop()

