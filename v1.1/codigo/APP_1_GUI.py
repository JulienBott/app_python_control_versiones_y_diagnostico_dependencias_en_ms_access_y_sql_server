
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog as fd
from threading import Thread
import datetime as dt
import time
import shutil

import APP_2_GUI_UTILS as mod_utils
import APP_3_GENERAL as mod_gen
import APP_4_BACK_END_MS_ACCESS as mod_access
import APP_5_BACK_END_SQL_SERVER as mod_sql_server

#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  CLASE - gui_ventana_inicio
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

class gui_ventana_inicio:

    def __init__(self, master, kwargs_gui_tags_scrolledtext_scripts = None, **dicc_kwargs_gui):

        self.master = master
        self.clase_gui_nombre = self.__class__.__name__
        self.kwargs_gui_ventana_inicio = dicc_kwargs_gui
        self.kwargs_gui_widgets_clase_actual = self.kwargs_gui_ventana_inicio[self.clase_gui_nombre].get("frames_root")
        self.kwargs_gui_tags_scrolledtext_scripts = kwargs_gui_tags_scrolledtext_scripts


        #se insertan los widgets y se almacenan en el diccionario dicc_widgets_gui_ventana_inicio
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_widgets_gui_ventana_inicio = {}
        for frame_contenedor in self.kwargs_gui_widgets_clase_actual.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_app_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_widgets_clase_actual.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_app_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_app_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_widgets_clase_actual[frame_contenedor].items() if widget != "frame"}

            dicc_widgets_frame_contenedor = {}
            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_app_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ","").strip()

                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]:
                    widget_objeto = mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "scrolledtext_propio":
                    widget_objeto = mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)

                elif tipo_widget_ajust == "treeview":
                    widget_objeto = mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "entry_propio":
                    widget_objeto = mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                dicc_widgets_frame_contenedor.update({frame_contenedor_widget:
                                                                                {"widget_objeto": widget_objeto
                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                }
                                                    })
                

            #se almacena el frame (objeto) en el diccionario dicc_widgets_gui_ventana_inicio junto con sus widgets (objetos)
            self.dicc_widgets_gui_ventana_inicio.update({frame_contenedor: 
                                                                        {"frame_contenedor_objeto": self.objeto_frame_contenedor
                                                                        , "dicc_widgets": dicc_widgets_frame_contenedor
                                                                        }
                                                        })
     
        #se recuperan los stringvar
        self.strvar_combobox_proceso = self.dicc_widgets_gui_ventana_inicio["frame_procesos"]["dicc_widgets"]["WIDGET_05"]["widget_objeto"].variable_enlace

        self.strvar_access_textbox_bbddd_1 = self.dicc_widgets_gui_ventana_inicio["frame_ms_access"]["dicc_widgets"]["WIDGET_10"]["widget_objeto"].variable_enlace
        self.strvar_access_textbox_bbddd_2 = self.dicc_widgets_gui_ventana_inicio["frame_ms_access"]["dicc_widgets"]["WIDGET_14"]["widget_objeto"].variable_enlace

        self.strvar_sql_server_servidor_1 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_19"]["widget_objeto"].variable_enlace
        self.strvar_sql_server_servidor_2 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_24"]["widget_objeto"].variable_enlace
        self.strvar_sql_server_bbdd_1 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_21"]["widget_objeto"].variable_enlace
        self.strvar_sql_server_bbdd_2 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_26"]["widget_objeto"].variable_enlace

        self.strvar_label_resolucion_pantalla = self.dicc_widgets_gui_ventana_inicio["frame_resolucion_pantalla"]["dicc_widgets"]["WIDGET_29"]["widget_objeto"].variable_enlace


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.widget_objeto_scrolledtext_comentario_proceso = self.dicc_widgets_gui_ventana_inicio["frame_procesos"]["dicc_widgets"]["WIDGET_07"]["widget_objeto"]
        self.widget_objeto_scrolledtext_comentario_proceso_height = self.kwargs_gui_widgets_clase_actual["frame_procesos"]["WIDGET_07"]["kwargs_config"]["height"]
        self.widget_objeto_sql_server_bbdd_asociada_servidor_1 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_21"]["widget_objeto"]
        self.widget_objeto_sql_server_bbdd_asociada_servidor_2 = self.dicc_widgets_gui_ventana_inicio["frame_sql_server"]["dicc_widgets"]["WIDGET_26"]["widget_objeto"]


        #se localiza si la resolucion de pantalla es la recomendada para poder informarlo en la GUI en caso de que no lo sea
        resolucion_pantalla_actual = str(self.master.widget_objeto.winfo_screenwidth()) + "x" + str(self.master.widget_objeto.winfo_screenheight())

        label_resolucion_pantalla = (f"{resolucion_pantalla_actual} (se recomienda usar {mod_gen.resolucion_pantalla_recomendada})"
                                    if resolucion_pantalla_actual != mod_gen.resolucion_pantalla_recomendada
                                    else f"{resolucion_pantalla_actual} (es la recomendada)")
        
        self.strvar_label_resolucion_pantalla.set(label_resolucion_pantalla)



    def def_GUI_guia_usuario(self):
        #rutina que permite descargar la guia de usuario en la ruta que indique el usuario

        msg = messagebox.askokcancel(title = mod_gen.nombre_app, message = "Se descargara el manual de usuario (pdf) y se abrira en la ruta que especifiques a continuación.\n\nDeseas continuar?")

        if msg:
            ruta_carpeta_guia_usuario = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres que se guarde la guia de usuario:")

            self.master.widget_objeto.config(cursor = "wait")

            now = dt.datetime.now()
            path_guia_usuario = os.path.join(ruta_carpeta_guia_usuario, mod_gen.nombre_guia_usuario + str(re.sub("[^0-9a-zA-Z]+", "_", str(now))) + ".pdf")
            shutil.copyfile(mod_gen.pdf_guia_usuario, path_guia_usuario)
            os.startfile(path_guia_usuario)

            self.master.widget_objeto.config(cursor = "")

            time.sleep(2)#para dar tiempo al pdf que se abra y no se ejecute antes del messagebox en la gui

            messagebox.showinfo(title = mod_gen.nombre_app, message = f"Guia de usuario descargada en: '{ruta_carpeta_guia_usuario}'.")



    def def_GUI_combobox_proceso(self):
        #rutina de evento (asociada al metodo bind) que permite según el proceso seleccionado
        #actualizar el scrolledtext con la descripción del proceso

        proceso_selecc = self.strvar_combobox_proceso.get()

        if proceso_selecc is not None:
            proceso_selecc_id = next((key for key, value in mod_gen.dicc_procesos.items() if value["PROCESO"] == proceso_selecc), None)

            comentario_proceso = "".join(mod_gen.dicc_procesos[proceso_selecc_id]["COMENTARIO"])

            self.widget_objeto_scrolledtext_comentario_proceso.modificaciones("borrar_contenido_y_tags")

            self.widget_objeto_scrolledtext_comentario_proceso.modificaciones("agregar_solo_contenido_desde_string"
                                                                                , string_texto_informar = comentario_proceso
                                                                                , height_scrolledtext = self.widget_objeto_scrolledtext_comentario_proceso_height)


    def def_GUI_sql_server(self, opcion, **kwargs):
        #rutina que permite cuando se informa el servidor comprobar si el usuario tiene acceso a las bbdd del mismo
        #(en caso de que no se borra el servidor seleccionado)

        opcion_servidor = kwargs.get("opcion_servidor", None)

        if opcion == "BBDD_ASOCIADAS_SERVIDOR":

            servidor_selecc = (self.strvar_sql_server_servidor_1.get() if opcion_servidor == "SERVIDOR_1"
                                else self.strvar_sql_server_servidor_2.get() if opcion_servidor == "SERVIDOR_2"
                                else None)

            tipo_bbdd = ("BBDD_01" if opcion_servidor == "SERVIDOR_1"
                           else "BBDD_02" if opcion_servidor == "SERVIDOR_2"
                           else None)


            widget_objeto_sql_server_bbdd_asociada_servidor_selecc = (self.widget_objeto_sql_server_bbdd_asociada_servidor_1 if opcion_servidor == "SERVIDOR_1"
                                                                        else self.widget_objeto_sql_server_bbdd_asociada_servidor_2 if opcion_servidor == "SERVIDOR_2"
                                                                        else None)
            

            if servidor_selecc is not None:
                #se calcula la connecting string probando conectar 1ero por windows authentication
                #si funciona la conexion se almacena en mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONNECTING_STRING"]
                #(o mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["CONNECTING_STRING"])
                #si falla se pasa por SQL Server autentication abriendo un toplevel para informar el login y el password
                if mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "WINDOWS_AUTHENTICATION":

                    #se localizan los permisos segun el combobox de servidor seleccionado
                    mod_gen.dicc_codigos_bbdd[tipo_bbdd]["SQL_SERVER"]["CONNECTING_STRING"] = mod_sql_server.conn_str_sql_server_windows_authentication
                    mod_sql_server.def_sql_server_servidor_permisos(tipo_bbdd, servidor_selecc)


                    if mod_sql_server.global_acceso_servidor_selecc == "NO":
                        messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes acceso al servidor seleccionado.")

                    else:
                        if not isinstance(mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo, list):
                            messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes permiso de acceso al código de los objetos de ninguna de las bbdd del servidor seleccionado.")
                    
                        else:
                            widget_objeto_sql_server_bbdd_asociada_servidor_selecc.config_atributos(**{"combobox_lista_opciones": mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo})


                elif mod_sql_server.func_sql_server_tipo_conexion_servidor(servidor_selecc) == "SQL_SERVER_AUTHENTICATION":

                    kwargs_gui_sql_server_authentication_dicc_config_root = self.kwargs_gui_ventana_inicio["gui_sql_server_authentication"]["dicc_config_root"]

                    self.toplevel_sql_server_authentication = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_sql_server_authentication_dicc_config_root)
                    self.toplevel_sql_server_authentication.config_atributos(**kwargs_gui_sql_server_authentication_dicc_config_root)

                    gui_sql_server_authentication(self.toplevel_sql_server_authentication
                                                , widget_objeto_combobox_bbdd_asociada_servidor = widget_objeto_sql_server_bbdd_asociada_servidor_selecc
                                                , tipo_bbdd = tipo_bbdd
                                                , servidor_sql_server = servidor_selecc
                                                , kwargs_gui_ventana_inicio = self.kwargs_gui_ventana_inicio
                                                )



        elif opcion == "CLEAR_SQL_SERVER_BBDD_01":     
            self.strvar_sql_server_servidor_1.set("")
            self.strvar_sql_server_bbdd_1.set("")
            self.widget_objeto_sql_server_bbdd_asociada_servidor_1.config_atributos(**{"combobox_lista_opciones": []})

            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] = None
            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"] = None
            mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONNECTING_STRING"] = None


        elif opcion == "CLEAR_SQL_SERVER_BBDD_02":
            self.strvar_sql_server_servidor_2.set("")
            self.strvar_sql_server_bbdd_2.set("")
            self.widget_objeto_sql_server_bbdd_asociada_servidor_2.config_atributos(**{"combobox_lista_opciones": []})

            mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"] = None
            mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"] = None
            mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["CONNECTING_STRING"] = None



    def def_GUI_ms_access(self, opcion, tipo_bbdd):
        #rutina que permite agregar o borrar (según el parametro opcion) en la GUI los valores de las bbdd MS Access seleccionadas

        if opcion == "ADD_MS_ACCESS":

            msg = messagebox.askokcancel(mod_gen.nombre_app, message = f"Selecciona la ubicación de {tipo_bbdd}.")

            if msg == True:
                path_bbdd = fd.askopenfilename(parent = self.master.widget_objeto, title = "", filetypes = mod_access.lista_GUI_askopenfilename_ms_access)

                if tipo_bbdd == "BBDD_01":
                    self.strvar_access_textbox_bbddd_1.set(path_bbdd)

                elif tipo_bbdd == "BBDD_02":
                    self.strvar_access_textbox_bbddd_2.set(path_bbdd)

                mod_gen.dicc_codigos_bbdd[tipo_bbdd]["MS_ACCESS"]["PATH_BBDD"] = path_bbdd


        elif opcion == "CLEAR_MS_ACCESS": 

            if tipo_bbdd == "BBDD_01":        
                self.strvar_access_textbox_bbddd_1.set("")
                mod_gen.dicc_codigos_bbdd[tipo_bbdd]["MS_ACCESS"]["PATH_BBDD"] = None

            elif tipo_bbdd == "BBDD_02":        
                self.strvar_access_textbox_bbddd_2.set("")
                mod_gen.dicc_codigos_bbdd[tipo_bbdd]["MS_ACCESS"]["PATH_BBDD"] = None


    def def_gui_threads(self):
        #rutina para ejecutar todos los procesos del app
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_gen.global_proceso_en_ejecucion == "SI":
            messagebox.showerror(title = mod_gen.nombre_app, message = "Espera a que acabe el proceso actualmente en ejecución.")

        else:
            Thread(target = self.def_gui_procesos).start()



    def def_gui_procesos(self):
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
                        mensaje2 = "MS ACCESS: no se ejecutara el proceso porque las rutas configuradas han de ser distintas.\n\n" if check_access_informado == "SI" else ""
                        mensaje3 = "Deseas continuar?"
                        mensaje = mensaje1 + mensaje2 + mensaje3

                    if check_control_versiones_access == "SI" and check_control_versiones_sql_server == "SI":
                        mensaje1 = "Se ejecutara el proceso sobre las 2 bbdd MS Access seleccionadas y sobre las 2 bbdd SQL Server seleccionadas.\n\n"
                        mensaje2 = "La BBDD_02 en los 2 casos es por defecto en la cual se hace el MERGE.\n\nDeseas continuar?"
                        mensaje = mensaje1 + mensaje2


                    msg = messagebox.askokcancel(mod_gen.nombre_app, message = mensaje)

                    if msg == True:
                        mensaje = "SELECCIONA DONDE QUIERES GUARDAR LOS LOGS DE ERRORES (si los hubiese):"
                        ruta_destino_logs = fd.askdirectory(parent = self.master.widget_objeto, title = mensaje)

                        self.master.widget_objeto.config(cursor = "wait")
                        mod_gen.def_calc_global(proceso_selecc_id, ruta_destino_logs)
                        self.master.widget_objeto.config(cursor = "")

                        if len(mod_gen.global_msg_errores_proceso_access) != 0 or len(mod_gen.global_msg_errores_proceso_sql_server) != 0:
                            messagebox.showerror(mod_gen.nombre_app, message = mod_gen.global_msg_errores_proceso_access + mod_gen.global_msg_errores_proceso_sql_server)

                        else:
                            kwargs_gui_control_versiones_dicc_config_root = self.kwargs_gui_ventana_inicio["gui_ventana_control_versiones"]["dicc_config_root"]

                            self.toplevel_control_versiones = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_control_versiones_dicc_config_root)
                            self.toplevel_control_versiones.config_atributos(**kwargs_gui_control_versiones_dicc_config_root)

                            gui_ventana_control_versiones(self.toplevel_control_versiones
                                                        , kwargs_gui_ventana_inicio = self.kwargs_gui_ventana_inicio
                                                        , kwargs_gui_tags_scrolledtext_scripts = self.kwargs_gui_tags_scrolledtext_scripts)



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
                        ruta_destino_output = fd.askdirectory(parent = self.master.widget_objeto, title = mensaje)

                        self.master.widget_objeto.config(cursor = "wait")
                        mod_gen.def_calc_global(proceso_selecc_id, ruta_destino_output, ruta_destino_excel_diagnostico_access = ruta_destino_output)
                        self.master.widget_objeto.config(cursor = "")


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

                        kwargs_gui_diagnostico_bbdd_sql_server_dicc_config_root = self.kwargs_gui_ventana_inicio["gui_diagnostico_bbdd_sql_server"]["dicc_config_root"]

                        self.toplevel_diagnostico_bbdd_sql_server = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_diagnostico_bbdd_sql_server_dicc_config_root)
                        self.toplevel_diagnostico_bbdd_sql_server.config_atributos(**kwargs_gui_diagnostico_bbdd_sql_server_dicc_config_root)

                        gui_diagnostico_bbdd_sql_server(self.toplevel_diagnostico_bbdd_sql_server, servidor_sql_server = servidor_sql_server_1, kwargs_gui_ventana_inicio = self.kwargs_gui_ventana_inicio)



#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  CLASE - gui_sql_server_authentication
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

class gui_sql_server_authentication():

    def __init__(self, master
                , widget_objeto_combobox_bbdd_asociada_servidor = None
                , servidor_sql_server = None
                , tipo_bbdd = None
                , kwargs_gui_ventana_inicio = None):

        self.master = master
        self.clase_gui_nombre = self.__class__.__name__

        self.kwargs_gui_widgets_clase_actual = kwargs_gui_ventana_inicio[self.clase_gui_nombre].get("frames_root")

        self.widget_objeto_combobox_bbdd_asociada_servidor = widget_objeto_combobox_bbdd_asociada_servidor
        self.servidor_sql_server = servidor_sql_server
        self.tipo_bbdd = tipo_bbdd


        #se insertan los widgets y se almacenan en el diccionario dicc_widgets_gui_sql_server_authentication
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_widgets_gui_sql_server_authentication = {}
        for frame_contenedor in self.kwargs_gui_widgets_clase_actual.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_app_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_widgets_clase_actual.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_app_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_app_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_widgets_clase_actual[frame_contenedor].items() if widget != "frame"}

            dicc_widgets_frame_contenedor = {}
            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_app_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ","").strip()

                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]:
                    widget_objeto = mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "scrolledtext_propio":
                    widget_objeto = mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)

                elif tipo_widget_ajust == "treeview":
                    widget_objeto = mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "entry_propio":
                    widget_objeto = mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                dicc_widgets_frame_contenedor.update({frame_contenedor_widget:
                                                                                {"widget_objeto": widget_objeto
                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                }
                                                    })
                

            #se almacena el frame (objeto) en el diccionario dicc_widgets_gui_sql_server_authentication junto con sus widgets (objetos)
            self.dicc_widgets_gui_sql_server_authentication.update({frame_contenedor: 
                                                                        {"frame_contenedor_objeto": self.objeto_frame_contenedor
                                                                        , "dicc_widgets": dicc_widgets_frame_contenedor
                                                                        }
                                                        })

        #se recuperan los stringvar
        self.strvar_servidor = self.dicc_widgets_gui_sql_server_authentication["frame_inicio"]["dicc_widgets"]["WIDGET_31"]["widget_objeto"].variable_enlace
        self.strvar_login = self.dicc_widgets_gui_sql_server_authentication["frame_inicio"]["dicc_widgets"]["WIDGET_33"]["widget_objeto"].variable_enlace
        self.strvar_password = self.dicc_widgets_gui_sql_server_authentication["frame_inicio"]["dicc_widgets"]["WIDGET_35"]["widget_objeto"].variable_enlace


        #se informa el entry del servidor
        self.strvar_servidor.set(self.servidor_sql_server)


    def def_GUI_conexion_servidor_sql_server(self):
        #rutina que permite (cuando la conexión a SQL Server es por SQL Server authentication es decir con login y password)
        #almacenar la connecting string SQL Server en dicc_codigos_bbdd[tipo_bbdd]["SQL_SERVER"]["CONNECTING_STRING"]
        
        servidor_selecc = self.strvar_servidor.get()
        login_selecc = self.strvar_login.get()
        password_selecc = self.strvar_password.get()

        if len(login_selecc) == 0 or len(password_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El login y el password son obligatorios.")

        elif len(login_selecc) != 0 and len(password_selecc) != 0:

            conn_string = mod_sql_server.conn_str_sql_server_login_password_authentication.replace("REEMPLAZA_LOGIN", login_selecc).replace("REEMPLAZA_PASSWORD", password_selecc)
            mod_gen.dicc_codigos_bbdd[self.tipo_bbdd]["SQL_SERVER"]["CONNECTING_STRING"] = conn_string

            self.master.widget_objeto.config(cursor = "wait")
            mod_sql_server.def_sql_server_servidor_permisos(self.tipo_bbdd, servidor_selecc)
            self.master.widget_objeto.config(cursor = "")

            if mod_sql_server.global_acceso_servidor_selecc == "NO":
                messagebox.showerror(title = mod_gen.nombre_app, message = "No tienes conexión al servidor seleccionado o el login / password son incorrectos.")
            else:
                #se actualiza la lista de opciones del widget (objeto) pasado como parametro y se cierra el toplevel de sql_server_authentication
                self.widget_objeto_combobox_bbdd_asociada_servidor.config_atributos(**{"combobox_lista_opciones": mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo})
                self.master.widget_objeto.destroy()



#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  CLASE - gui_sql_server_authentication
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

class gui_diagnostico_bbdd_sql_server():

    def __init__(self, master, servidor_sql_server = None, kwargs_gui_ventana_inicio = None):

        self.master = master
        self.clase_gui_nombre = self.__class__.__name__

        self.kwargs_gui_widgets_clase_actual = kwargs_gui_ventana_inicio[self.clase_gui_nombre].get("frames_root")
        self.servidor_sql_server = servidor_sql_server


        #se insertan los widgets y se almacenan en el diccionario dicc_widgets_gui_diagnostico_bbdd_sql_server
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_widgets_gui_diagnostico_bbdd_sql_server = {}
        for frame_contenedor in self.kwargs_gui_widgets_clase_actual.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_app_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_widgets_clase_actual.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_app_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_app_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_widgets_clase_actual[frame_contenedor].items() if widget != "frame"}

            dicc_widgets_frame_contenedor = {}
            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_app_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ","").strip()

                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]:
                    widget_objeto = mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "scrolledtext_propio":
                    widget_objeto = mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)

                elif tipo_widget_ajust == "treeview":
                    widget_objeto = mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "entry_propio":
                    widget_objeto = mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                dicc_widgets_frame_contenedor.update({frame_contenedor_widget:
                                                                                {"widget_objeto": widget_objeto
                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                }
                                                    })
                

            #se almacena el frame (objeto) en el diccionario dicc_widgets_gui_diagnostico_bbdd_sql_server junto con sus widgets (objetos)
            self.dicc_widgets_gui_diagnostico_bbdd_sql_server.update({frame_contenedor: 
                                                                        {"frame_contenedor_objeto": self.objeto_frame_contenedor
                                                                        , "dicc_widgets": dicc_widgets_frame_contenedor
                                                                        }
                                                                    })

        #se recuperan los stringvar
        self.strvar_sql_server_diagnostico_combobox_opciones = self.dicc_widgets_gui_diagnostico_bbdd_sql_server["frame_inicio"]["dicc_widgets"]["WIDGET_38"]["widget_objeto"].variable_enlace
        self.strvar_sql_server_diagnostico_listbox_bbdd = self.dicc_widgets_gui_diagnostico_bbdd_sql_server["frame_inicio"]["dicc_widgets"]["WIDGET_41"]["widget_objeto"].variable_enlace


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.widget_objeto_sql_server_listbox_bbdd = self.dicc_widgets_gui_diagnostico_bbdd_sql_server["frame_inicio"]["dicc_widgets"]["WIDGET_41"]["widget_objeto"]


        #se crea la lista de bbdd asociadas al servidor
        mod_sql_server.def_sql_server_servidor_permisos("BBDD_01", self.servidor_sql_server)
        mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo

        #se crean las opciones del listbox con las bbdd (con permisos de acceso al codigo) del servidor 1 configurado en la GUI de inicio
        if isinstance(mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo, list):
            self.widget_objeto_sql_server_listbox_bbdd.config_atributos(**{"listbox_lista_items": mod_sql_server.global_servidor_bbdd_permisos_acceso_codigo})



    def def_GUI_sql_server_diagnostico_threads(self):
        #rutina para ejecutar el proceso de diagnostico de dependencias en SQL Server
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_gen.global_proceso_en_ejecucion == "SI":
            messagebox.showerror(title = mod_gen.nombre_app, message = "Espera a que acabe el proceso actualmente en ejecución.")

        else:
            Thread(target = self.def_GUI_sql_server_diagnostico_boton_check).start()



    def def_GUI_sql_server_diagnostico_listbox_all_none(self):
        #rutina para seleccionar o des-seleccionar las bbdd del servidor que entran en el calculo del proceso de diagnostico SQL Server

        self.widget_objeto_sql_server_listbox_bbdd.config_atributos(**{"listbox_seleccionar_todo_o_nada": True})


    def def_GUI_sql_server_diagnostico_boton_check(self):
        #rutina para ejecutar el proceso de diagnostico en SQL Server (PROCESO_03)

        opcion_diagnostico_sql_server = self.strvar_sql_server_diagnostico_combobox_opciones.get()

        lista_bbdd_selecc = self.widget_objeto_sql_server_listbox_bbdd.config_atributos(**{"listbox_lista_items_seleccionados": True})

        if len(opcion_diagnostico_sql_server) == 0 or len(lista_bbdd_selecc) == 0:
            messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de selección y la(s) bbdd son obligatorios.")

        else:
            msg = messagebox.askokcancel(mod_gen.nombre_app, message = f"Se ejecutara el proceso '{opcion_diagnostico_sql_server}'.\n\nTendras que seleccionar una ruta donde generar el output.\n\nDeseas continuar?")

            if msg:
                mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] = self.servidor_sql_server
                mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"] = lista_bbdd_selecc

                #opcion diagnostico
                if opcion_diagnostico_sql_server == mod_sql_server.lista_GUI_diagnostico_combobox_sql_server[0]:
                    mensaje_1 = "SELECCIONA DONDE QUIERES GUARDAR EL EXCEL DE DIAGNOSTICO DE DEPENDENCIAS Y LOS LOGS DE ERRORES (si los hubiese):"
                    mensaje_2 = "Diagnostico de dependencias SQL Server descargado en Excel en la ruta indicada."

                #opcion descarga codigo
                elif opcion_diagnostico_sql_server == mod_sql_server.lista_GUI_diagnostico_combobox_sql_server[1]:
                    mensaje_1 = "SELECCIONA DONDE QUIERES GUARDAR LOS CODIGOS T-SQL DE LOS OBJETOS Y LOS LOGS DE ERRORES (si los hubiese):"
                    mensaje_2 = "Códigos T-SQL de los objetos descargados en ficheros .sql en la ruta indicada."

                ruta_destino_diagnostico_sql_server = fd.askdirectory(parent = self.master.widget_objeto, title = mensaje_1)

                self.master.widget_objeto.config(cursor = "wait")
                mod_gen.def_calc_global("PROCESO_03", ruta_destino_diagnostico_sql_server
                                                                                        , opcion_diagnostico_sql_server = opcion_diagnostico_sql_server
                                                                                        , ruta_destino_diagnostico_sql_server = ruta_destino_diagnostico_sql_server)
                                                                                        
                self.master.widget_objeto.config(cursor = "")


                if len(mod_gen.global_msg_errores_proceso_sql_server) != 0:
                    messagebox.showerror(mod_gen.nombre_app, message = mod_gen.global_msg_errores_proceso_sql_server)

                else:
                    messagebox.showinfo(mod_gen.nombre_app, message = mensaje_2)


#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  CLASE - gui_ventana_control_versiones
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

class gui_ventana_control_versiones():

    def __init__(self, master, kwargs_gui_ventana_inicio = None, kwargs_gui_tags_scrolledtext_scripts = None):

        self.master = master
        self.clase_gui_nombre = self.__class__.__name__

        self.kwargs_gui_ventana_inicio = kwargs_gui_ventana_inicio
        self.kwargs_gui_widgets_clase_actual = self.kwargs_gui_ventana_inicio[self.clase_gui_nombre].get("frames_root")
        self.lista_columnas_df_para_treeview = self.kwargs_gui_widgets_clase_actual["frame_inicio"]["WIDGET_49"]["kwargs_config"]["dicc_treeview"]["columnas_df"]
        self.kwargs_gui_tags_scrolledtext_scripts = kwargs_gui_tags_scrolledtext_scripts


        #se insertan los widgets y se almacenan en el diccionario dicc_widgets_gui_ventana_control_versiones
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_widgets_gui_ventana_control_versiones = {}
        for frame_contenedor in self.kwargs_gui_widgets_clase_actual.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_app_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_widgets_clase_actual.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_app_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_app_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_widgets_clase_actual[frame_contenedor].items() if widget != "frame"}

            dicc_widgets_frame_contenedor = {}
            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_app_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ","").strip()

                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]:
                    widget_objeto = mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "scrolledtext_propio":
                    widget_objeto = mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)

                elif tipo_widget_ajust == "treeview":
                    widget_objeto = mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "entry_propio":
                    widget_objeto = mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)



                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                dicc_widgets_frame_contenedor.update({frame_contenedor_widget:
                                                                                {"widget_objeto": widget_objeto
                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                }
                                                    })
                

            #se almacena el frame (objeto) en el diccionario dicc_widgets_gui_ventana_control_versiones junto con sus widgets (objetos)
            self.dicc_widgets_gui_ventana_control_versiones.update({frame_contenedor: 
                                                                        {"frame_contenedor_objeto": self.objeto_frame_contenedor
                                                                        , "dicc_widgets": dicc_widgets_frame_contenedor
                                                                        }
                                                                    })

        #se recuperan los stringvar
        self.strvar_combobox_tipo_objeto = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_44"]["widget_objeto"].variable_enlace
        self.strvar_combobox_tipo_concepto = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_46"]["widget_objeto"].variable_enlace
        self.strvar_name_bbdd_1 = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_51"]["widget_objeto"].variable_enlace
        self.strvar_name_bbdd_2 = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_54"]["widget_objeto"].variable_enlace

        self.strvar_proceso_merge_bbdd_origen = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_62"]["widget_objeto"].variable_enlace
        self.strvar_combobox_merge_accion = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_64"]["widget_objeto"].variable_enlace
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1 = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_66"]["widget_objeto"].variable_enlace
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2 = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_67"]["widget_objeto"].variable_enlace
        self.strvar_proceso_merge_bbdd_lineas_destino_selecc = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_69"]["widget_objeto"].variable_enlace


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        #(mediante el diccionario dicc_widgets_gui_ventana_control_versiones)
        self.widget_objeto_combobox_tipo_objeto = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_44"]["widget_objeto"]
        self.widget_objeto_treeview_objetos = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_49"]["widget_objeto"]
        self.widget_objeto_scrolledtext_script_1 = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_52"]["widget_objeto"]
        self.widget_objeto_scrolledtext_script_2 = self.dicc_widgets_gui_ventana_control_versiones["frame_inicio"]["dicc_widgets"]["WIDGET_55"]["widget_objeto"]
        self.widget_objeto_combobox_accion = self.dicc_widgets_gui_ventana_control_versiones["frame_merge"]["dicc_widgets"]["WIDGET_64"]["widget_objeto"]


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        #(mediante el diccionario kwargs_gui_widgets_clase_actual)
        self.widget_objeto_scrolledtext_script_1_height = self.kwargs_gui_widgets_clase_actual["frame_inicio"]["WIDGET_52"]["kwargs_config"]["height"]
        self.widget_objeto_scrolledtext_script_2_height = self.kwargs_gui_widgets_clase_actual["frame_inicio"]["WIDGET_55"]["kwargs_config"]["height"]


        #al inicio de la clase se establecen los nombres de las bbdd en los scripts a ""
        self.strvar_name_bbdd_1.set("")
        self.strvar_name_bbdd_2.set("")


        #la lista de valores del combobox tipo objeto varia segun que se haya configurado o no control de versiones MS Access y/o SQL Server
        hacer_control_versiones_access = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS")
        hacer_control_versiones_sql_server = mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "SQL_SERVER")

        lista_opciones_combobox_tipo_objeto = []
        if hacer_control_versiones_access == "SI" and hacer_control_versiones_sql_server == "SI":
            lista_opciones_combobox_tipo_objeto = mod_gen.lista_GUI_seleccion_tipo_objeto_access + mod_gen.lista_GUI_seleccion_tipo_objeto_sql_server

        elif hacer_control_versiones_access == "SI" and hacer_control_versiones_sql_server == "NO":
            lista_opciones_combobox_tipo_objeto = mod_gen.lista_GUI_seleccion_tipo_objeto_access

        elif hacer_control_versiones_access == "NO" and hacer_control_versiones_sql_server == "SI":
            lista_opciones_combobox_tipo_objeto = mod_gen.lista_GUI_seleccion_tipo_objeto_sql_server

        self.widget_objeto_combobox_tipo_objeto.config_atributos(**{"combobox_lista_opciones": lista_opciones_combobox_tipo_objeto})




    #####################################################################################################################################
    #             RUTINAS CONTROL VERSIONES
    #####################################################################################################################################

    def func_ajustes_df_scrolledtext(self, df_scrolledtext):
        #funcion interna que permite asignar el numero de linea seguido de la linea de codigo para informar en el scrolledtext
        #se realiza tanto en las columnas CODIGO_CON_NUM_LINEA como CODIGO_OTRA_BBDD_CON_NUM_LINEA

        if isinstance(df_scrolledtext, pd.DataFrame):

            df_scrolledtext["CODIGO_CON_NUM_LINEA"] = None
            df_scrolledtext["NUM_LINEA"] = df_scrolledtext.apply(lambda x: None if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" else x["NUM_LINEA"], axis = 1)

            ind_col_num_linea = df_scrolledtext.columns.get_loc("NUM_LINEA")
            ind_col_codigo = df_scrolledtext.columns.get_loc("CODIGO")
            ind_col_codigo_otra_bbdd = df_scrolledtext.columns.get_loc("CODIGO_OTRA_BBDD")
            ind_col_codigo_con_num_linea = df_scrolledtext.columns.get_loc("CODIGO_CON_NUM_LINEA")
            ind_col_codigo_con_num_linea_otra_bbdd = df_scrolledtext.columns.get_loc("CODIGO_CON_NUM_LINEA_OTRA_BBDD")
            ind_col_control_cambios = df_scrolledtext.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")

            lambda_func_num_lineas = None
            cont = 0
            for ind in df_scrolledtext.index:

                linea_codigo = df_scrolledtext.iloc[ind, ind_col_codigo]
                linea_codigo_otra_bbdd = df_scrolledtext.iloc[ind, ind_col_codigo_otra_bbdd]
                linea_num_linea = df_scrolledtext.iloc[ind, ind_col_num_linea]

                if df_scrolledtext.iloc[ind, ind_col_control_cambios] != "ELIMINADO":
                    cont += 1

                    lambda_func_num_lineas = lambda cont: f"{cont:04}\t"
                    df_scrolledtext.iloc[ind, ind_col_num_linea] = lambda_func_num_lineas(cont)

                    df_scrolledtext.iloc[ind, ind_col_codigo_con_num_linea] = (lambda_func_num_lineas(cont) + linea_codigo
                                                                               if linea_codigo is not None
                                                                               else None)
                    
                    df_scrolledtext.iloc[ind, ind_col_codigo_con_num_linea_otra_bbdd] = (lambda_func_num_lineas(cont) + linea_codigo_otra_bbdd
                                                                                        if linea_codigo_otra_bbdd is not None
                                                                                        else None)

                else:
                    df_scrolledtext.iloc[ind, ind_col_codigo_con_num_linea] = ("    \t" + linea_num_linea + linea_codigo
                                                                               if linea_codigo is not None
                                                                               else None)
                    
                    df_scrolledtext.iloc[ind, ind_col_codigo_con_num_linea_otra_bbdd] = ("    \t" + linea_num_linea + linea_codigo_otra_bbdd
                                                                                        if linea_codigo_otra_bbdd is not None
                                                                                        else None)
        else:
            df_scrolledtext = pd.DataFrame(mod_gen.lista_headers_df_codigo_control_versiones_2)


        return df_scrolledtext


    def def_control_versiones_cambio_tipo_objeto(self):
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



    def def_control_versiones_frame_inicio_click_boton(self, opcion_boton):

        ###########################################################################################################
        #permite al pulsar el boton VER y según el tipo de objeto y de concepto seleccionado actualiza
        #el sub-formulario con los objetos con cambios de una bbdd a otra
        #se combina con la rutina def_control_versiones_update_subform_objetos (más abajo)
        ###########################################################################################################
        if opcion_boton == "ACTUALIZAR_SUBFORM_OBJETOS_CONTROL_VERSIONES":

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
                        #se borra el contenido y tags de los scrolledtext de scripts
                        self.widget_objeto_scrolledtext_script_1.modificaciones("borrar_contenido_y_tags")
                        self.widget_objeto_scrolledtext_script_2.modificaciones("borrar_contenido_y_tags")


                        #se calcula el df para el treeview
                        df_treeview = pd.DataFrame(columns = self.lista_columnas_df_para_treeview)
                        if len(lista_control_versiones_segun_tipo_objeto_selecc) != 0:

                            if tipo_objeto_selecc_key == "TODOS":

                                lista_datos_treeview = [[dicc["TIPO_BBDD"], dicc["TIPO_OBJETO_SUBFORM"], dicc["REPOSITORIO"], dicc["NOMBRE_OBJETO"], dicc["NUM_CAMBIOS_SCRIPT_1"], dicc["NUM_CAMBIOS_SCRIPT_2"]] 
                                                        for dicc in lista_control_versiones_segun_tipo_objeto_selecc
                                                        if dicc["TIPO_BBDD"] == tipo_bbdd and dicc["CHECK_OBJETO"] == tipo_concepto_selecc]

                            else:
                                lista_datos_treeview = [[dicc["TIPO_BBDD"], dicc["TIPO_OBJETO_SUBFORM"], dicc["REPOSITORIO"], dicc["NOMBRE_OBJETO"], dicc["NUM_CAMBIOS_SCRIPT_1"], dicc["NUM_CAMBIOS_SCRIPT_2"]] 
                                                        for dicc in lista_control_versiones_segun_tipo_objeto_selecc
                                                        if dicc["TIPO_BBDD"] == tipo_bbdd and dicc["TIPO_OBJETO"] == tipo_objeto_selecc_key and dicc["CHECK_OBJETO"] == tipo_concepto_selecc]


                            df_treeview = pd.DataFrame(lista_datos_treeview, columns = self.lista_columnas_df_para_treeview)
                            df_treeview = df_treeview.replace({None:"---"})

                        #se rellena el treeview
                        self.widget_objeto_treeview_objetos.modificaciones("actualizar_desde_df", df_datos = df_treeview)


                        #se vacian los widgets del frame de merge
                        self.strvar_proceso_merge_bbdd_origen.set("")
                        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
                        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
                        self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")
                        self.strvar_combobox_merge_accion.set("")

                        #se genera el message box con el numero de objetos localizados
                        num_objetos = sum(1 if dicc["CHECK_OBJETO"] == tipo_concepto_selecc else 0 for dicc in lista_control_versiones_segun_tipo_objeto_selecc)

                        messagebox.showinfo(title = mod_gen.nombre_app, message = str(num_objetos) + " objetos localizados con cambios.")


        ###########################################################################################################
        #permite descargar a excel los objetos con cambios de una bbdd a otra según el tipo de objeto seleccionado
        ###########################################################################################################
        elif opcion_boton == "EXPORTAR_EXCEL_OBJETOS_CONTROL_VERSIONES":

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

                        path_xls = fd.askdirectory(parent = self.master.widget_objeto, title = "INDICA DONDE QUIERES GUARDAR EL FICHERO EXCEL:")

                        self.master.widget_objeto.config(cursor = "wait")
                        mod_gen.def_control_versiones_export_excel(tipo_bbdd, lista_control_versiones_selecc, path_xls)
                        self.master.widget_objeto.config(cursor = "")
                        
                        del lista_control_versiones_selecc

                        messagebox.showinfo(mod_gen.nombre_app, message = "Excel generado.")


    def def_control_versiones_update_subform_objetos_click_item(self):
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

        #el valor de la key lista_datos_item_seleccionado del atributo datos_item_seleccionado del widget widget_objeto_treeview_objetos
        #es lista de lista con a cada click en 1 item distinto una sola subliste (seleccion simple)
        global_tipo_objeto_subform = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][1]
        global_repositorio_subform = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][2]
        global_nombre_objeto_subform = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][3]


        #se recuperan los df codigos actuales de BBDD_01 y BBDD_02
        dicc_proceso_merge_anteriores = mod_gen.func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_selecc, global_tipo_objeto_subform, global_repositorio_subform, global_nombre_objeto_subform)
        
        df_codigo_actual_1 = self.func_ajustes_df_scrolledtext(dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_1"])
        df_codigo_actual_2 = self.func_ajustes_df_scrolledtext(dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_2"])

        self.widget_objeto_scrolledtext_script_1.modificaciones("borrar_contenido_y_tags")
        self.widget_objeto_scrolledtext_script_1.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                                                , df_datos = df_codigo_actual_1
                                                                , height_scrolledtext = self.widget_objeto_scrolledtext_script_1_height
                                                                , **self.kwargs_gui_tags_scrolledtext_scripts)
        
        self.widget_objeto_scrolledtext_script_1.config_atributos(**{"state": tk.DISABLED})

        self.widget_objeto_scrolledtext_script_2.modificaciones("borrar_contenido_y_tags")
        self.widget_objeto_scrolledtext_script_2.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                                                , df_datos = df_codigo_actual_2
                                                                , height_scrolledtext = self.widget_objeto_scrolledtext_script_1_height
                                                                , **self.kwargs_gui_tags_scrolledtext_scripts)
        
        self.widget_objeto_scrolledtext_script_2.config_atributos(**{"state": tk.DISABLED})
 


        #se vacian los widgets de proceso merge
        self.strvar_proceso_merge_bbdd_origen.set("")
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
        self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
        self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")
        self.strvar_combobox_merge_accion.set("")



    #####################################################################################################################################
    #             RUTINAS PROCESO MERGE
    #####################################################################################################################################

    def def_proceso_merge_combobox_bbdd_selecc(self):
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
            combobox_lista_opciones = (mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_1 if bbdd_origen_selecc == "BBDD_01"
                                       else mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_2 if bbdd_origen_selecc == "BBDD_02"
                                       else [])

            self.widget_objeto_combobox_accion.config_atributos(**{"combobox_lista_opciones": combobox_lista_opciones})


    def def_click_proceso_merge_boton_merge(self):
        #rutina que permite traspasar en la GUI los cambios realizados por el usuario de un script a otro al pulsar el botón "ACCIÓN"
        #y conserver los cambios, mediante la rutina def_proceso_merge_realizar_cambios (modulo general)

        global global_tipo_objeto_subform
        global global_repositorio_subform
        global global_nombre_objeto_subform

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
        try:
            dicc_proceso_merge_anteriores = mod_gen.func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_selecc
                                                                                                        , global_tipo_objeto_subform
                                                                                                        , global_repositorio_subform
                                                                                                        , global_nombre_objeto_subform)
            
            lista_merge_hechos = dicc_proceso_merge_anteriores["LISTA_DICC_MERGE_HECHOS"]

        except NameError as err:
            messagebox.showerror(title = mod_gen.nombre_app, message = "No has seleccionado ningún objeto.")

        else:
            if len(tipo_objeto_selecc) == 0:
                self.strvar_proceso_merge_bbdd_origen.set("")
                messagebox.showerror(title = mod_gen.nombre_app, message = "No has seleccionado ningún objeto.")

            else:
                if len(bbdd_origen) == 0 or len(tipo_accion) == 0:
                    messagebox.showerror(title = mod_gen.nombre_app, message = "La selección de bbdd y el tipo de acción son obligatorios.")

                else:
                    if tipo_accion == mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_1[1] and (len(lineas_origen) == 0 or len(lineas_destino) == 0):#migrar por lineas
                        messagebox.showerror(title = mod_gen.nombre_app, message = "Las lineas de origen y destino son obligatorias.")
                    
                    elif tipo_accion == mod_gen.lista_GUI_proceso_merge_tipo_accion_bbdd_2[1] and len(lineas_origen) == 0:#quitar por lineas
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

                                df_codigo_actual_1 = self.func_ajustes_df_scrolledtext(dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_1"])
                                df_codigo_actual_2 = self.func_ajustes_df_scrolledtext(dicc_proceso_merge_anteriores["DF_CODIGO_ACTUAL_2"])

                                #se rellenan los scrolledtext
                                self.widget_objeto_scrolledtext_script_1.modificaciones("borrar_contenido_y_tags")
                                self.widget_objeto_scrolledtext_script_1.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                                                                        , df_datos = df_codigo_actual_1
                                                                                        , height_scrolledtext = self.widget_objeto_scrolledtext_script_1_height
                                                                                        , **self.kwargs_gui_tags_scrolledtext_scripts)
                                
                                self.widget_objeto_scrolledtext_script_1.config_atributos(**{"state": tk.DISABLED})

                                self.widget_objeto_scrolledtext_script_2.modificaciones("borrar_contenido_y_tags")
                                self.widget_objeto_scrolledtext_script_2.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                                                                        , df_datos = df_codigo_actual_2
                                                                                        , height_scrolledtext = self.widget_objeto_scrolledtext_script_2_height
                                                                                        , **self.kwargs_gui_tags_scrolledtext_scripts)
                                
                                self.widget_objeto_scrolledtext_script_2.config_atributos(**{"state": tk.DISABLED})

                                #se reinician los algunos widgets del frame merge
                                self.strvar_combobox_merge_accion.set("")
                                self.strvar_proceso_merge_bbdd_lineas_origen_selecc_1.set("")
                                self.strvar_proceso_merge_bbdd_lineas_origen_selecc_2.set("")
                                self.strvar_proceso_merge_bbdd_lineas_destino_selecc.set("")


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

                kwargs_gui_control_versiones_merge_bbdd_fisicas_dicc_config_root = self.kwargs_gui_ventana_inicio["gui_ventana_merge_bbdd_fisicas"]["dicc_config_root"]

                self.toplevel_control_versiones_merge_bbdd_fisicas = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_control_versiones_merge_bbdd_fisicas_dicc_config_root)
                self.toplevel_control_versiones_merge_bbdd_fisicas.config_atributos(**kwargs_gui_control_versiones_merge_bbdd_fisicas_dicc_config_root)

                gui_ventana_merge_bbdd_fisicas(self.toplevel_control_versiones_merge_bbdd_fisicas
                                               , kwargs_gui_ventana_inicio = self.kwargs_gui_ventana_inicio
                                               , kwargs_gui_tags_scrolledtext_scripts = self.kwargs_gui_tags_scrolledtext_scripts)



#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  CLASE - gui_ventana_merge_bbdd_fisicas
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

class gui_ventana_merge_bbdd_fisicas():

    def __init__(self, master, kwargs_gui_ventana_inicio = None, kwargs_gui_tags_scrolledtext_scripts = None):

        self.master = master
        self.clase_gui_nombre = self.__class__.__name__

        self.kwargs_gui_widgets_clase_actual = kwargs_gui_ventana_inicio[self.clase_gui_nombre].get("frames_root")
        self.lista_columnas_df_para_treeview = self.kwargs_gui_widgets_clase_actual["frame_inicio"]["WIDGET_76"]["kwargs_config"]["dicc_treeview"]["columnas_df"]
        self.kwargs_gui_tags_scrolledtext_scripts = kwargs_gui_tags_scrolledtext_scripts


        #se insertan los widgets y se almacenan en el diccionario dicc_widgets_gui_ventana_merge_bbdd_fisicas
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_widgets_gui_ventana_merge_bbdd_fisicas = {}
        for frame_contenedor in self.kwargs_gui_widgets_clase_actual.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_app_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_widgets_clase_actual.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_app_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_app_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_widgets_clase_actual[frame_contenedor].items() if widget != "frame"}

            dicc_widgets_frame_contenedor = {}
            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_app_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ","").strip()

                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]:
                    widget_objeto = mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "scrolledtext_propio":
                    widget_objeto = mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)

                elif tipo_widget_ajust == "treeview":
                    widget_objeto = mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, self_clase_gui_donde_call_rutina = self, **kwargs_config)

                elif tipo_widget_ajust == "entry_propio":
                    widget_objeto = mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)



                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                dicc_widgets_frame_contenedor.update({frame_contenedor_widget:
                                                                                {"widget_objeto": widget_objeto
                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                }
                                                    })
                

            #se almacena el frame (objeto) en el diccionario dicc_widgets_gui_ventana_merge_bbdd_fisicas junto con sus widgets (objetos)
            self.dicc_widgets_gui_ventana_merge_bbdd_fisicas.update({frame_contenedor: 
                                                                        {"frame_contenedor_objeto": self.objeto_frame_contenedor
                                                                        , "dicc_widgets": dicc_widgets_frame_contenedor
                                                                        }
                                                                    })

        #se recuperan los stringvar
        self.strvar_combobox_tipo_seleccion = self.dicc_widgets_gui_ventana_merge_bbdd_fisicas["frame_inicio"]["dicc_widgets"]["WIDGET_73"]["widget_objeto"].variable_enlace



        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        #(mediante el diccionario dicc_widgets_gui_ventana_merge_bbdd_fisicas)
        self.widget_objeto_combobox_tipo_seleccion = self.dicc_widgets_gui_ventana_merge_bbdd_fisicas["frame_inicio"]["dicc_widgets"]["WIDGET_73"]["widget_objeto"]
        self.widget_objeto_treeview_objetos = self.dicc_widgets_gui_ventana_merge_bbdd_fisicas["frame_inicio"]["dicc_widgets"]["WIDGET_76"]["widget_objeto"]
        self.widget_objeto_scrolledtext_script_a_migrar = self.dicc_widgets_gui_ventana_merge_bbdd_fisicas["frame_inicio"]["dicc_widgets"]["WIDGET_78"]["widget_objeto"]


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        #(mediante el diccionario kwargs_gui_widgets_clase_actual)
        self.widget_objeto_scrolledtext_script_a_migrar_height = self.kwargs_gui_widgets_clase_actual["frame_inicio"]["WIDGET_78"]["kwargs_config"]["height"]


        #se calculan las listas de objetos donde realizar merge en bbdd fisica asociadas a cada opcion del combobox
        mod_gen.def_merge_bbdd_fisica_lista_objetos()


        #se calculan los ajustes manuales a realizar en access
        if mod_gen.func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS") == "SI":
            mod_gen.def_merge_access_ajustes_manuales()
                                                                                                       

        #se calcula la lista para el combobox de seleccion
        lista_combobox_seleccion = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_COMBOBOX_MERGE_BBDD_FISICAS")
        self.widget_objeto_combobox_tipo_seleccion.config_atributos(**{"combobox_lista_opciones": lista_combobox_seleccion})


    #####################################################################################################################################
    #             RUTINAS
    #####################################################################################################################################

    def def_merge_bbdd_fisicas_click_boton(self, opcion_selecc):

        if opcion_selecc == "ACTUALIZAR_SUBFORM_CONTROL_VERSIONES_MERGE_BBDD_FISICA":
            #permite tras seleccionar el tipo de selección y pulsar el botón VER
            #actualizar el sub-formulario con los objetos donde el usuario ha realizado cambios
            #se asocia con la rutina def_merge_bbdd_fisicas_update_subform_objetos (ver más abajo)

            opcion_proceso_merge = self.strvar_combobox_tipo_seleccion.get()

            if len(opcion_proceso_merge) == 0:
                messagebox.showerror(title = mod_gen.nombre_app, message = "El tipo de selección es obligatorio.")
            else:
                #se borra el contenido y tags del scrolledtext de script a migrar
                self.widget_objeto_scrolledtext_script_a_migrar.modificaciones("borrar_contenido_y_tags")


                #se calcula el df para el treeview
                lista_dicc_datos_para_treeview = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)

                df_treeview = pd.DataFrame(columns = self.lista_columnas_df_para_treeview)
                if isinstance(lista_dicc_datos_para_treeview, list) and len(lista_dicc_datos_para_treeview) != 0: #los errores de migracion de inicio son None

                    lista_datos_para_treeview = []
                    for dicc in lista_dicc_datos_para_treeview:
                        tipo_bbdd = dicc["TIPO_BBDD"]
                        tipo_objeto_subform = dicc["TIPO_OBJETO_SUBFORM"]
                        tipo_repositorio = dicc["TIPO_REPOSITORIO"]
                        repositorio = dicc["REPOSITORIO"]
                        nombre_objeto = dicc["NOMBRE_OBJETO"]
                        estado_migracion = dicc["ESTADO_MIGRACION"]

                        lista_datos_para_treeview.append([tipo_bbdd, tipo_objeto_subform, tipo_repositorio, repositorio, nombre_objeto, estado_migracion])

                    df_treeview = pd.DataFrame(lista_datos_para_treeview, columns = self.lista_columnas_df_para_treeview)
                    df_treeview = df_treeview.replace({None: "---"})


                    #se rellena el treeview
                    self.widget_objeto_treeview_objetos.modificaciones("actualizar_desde_df", df_datos = df_treeview)


                    #se genera el message box con el numero de objetos localizados
                    num_objetos = len(lista_datos_para_treeview)
                    messagebox.showinfo(title = mod_gen.nombre_app, message = str(num_objetos) + " objetos localizados.")

                else:
                    messagebox.showinfo(title = mod_gen.nombre_app, message = "0 objetos localizados.")



        if opcion_selecc == "REALIZAR_MERGE_BBDD_FISICA":
            #permite ejecutar el merge en bbdd fisica y generar los logs de OK y los de errores (si los hubiese)
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

                    ruta_export = fd.askdirectory(parent = self.master.widget_objeto, title = "RUTA DONDE GUARDAR LOS FICHEROS DE LOGS (OK + ERRORES) + LA DOCUMENTACIÓN DEL PROCESO:")

                    self.master.widget_objeto.config(cursor = "wait")

                    #se reestablece mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] a None
                    mod_gen.dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = None

                    #se ejecuta el proceso de merge
                    lista_dicc_objetos_migrar_bbdd_fisica = [dicc for dicc in lista_dicc_objetos_migrar_bbdd_fisica if dicc["TIPO_BBDD"] == tipo_bbdd]
                    mod_gen.def_merge_bbdd_fisicas(tipo_bbdd, lista_dicc_objetos_migrar_bbdd_fisica, ruta_export)

                    #se ejecutan los logs (OK + errores) y crea los ficheros (se ejecuta sobre BBDD_02 que es por defecto donde se hace el merge)
                    mod_gen.def_generacion_logs("MERGE_BBDD_FISICA_LOGS_OK", tipo_bbdd, ruta_export, opcion_bbdd = "BBDD_02")
                    mod_gen.def_generacion_logs("MERGE_BBDD_FISICA_LOGS_ERRORES", tipo_bbdd, ruta_export, opcion_bbdd = "BBDD_02")

                    self.master.widget_objeto.config(cursor = "")


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



    def def_merge_bbdd_fisicas_update_subform_objetos_click_item(self, opcion_proceso_merge):
        #rutina que permite al hacer click en el sub-formulario de objetos con cambios realizados por el usuario
        #actualizar el cuadro de script con las lineas cambiadas marcadas en el color correspondiente según la accion realizada

        lista_dicc_merge_bbdd_fisicas = mod_gen.func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS", opcion_gui_merge_bbdd_fisica = opcion_proceso_merge)


        #el valor de la key lista_datos_item_seleccionado del atributo datos_item_seleccionado del widget widget_objeto_treeview_objetos
        #es lista de lista con a cada click en 1 item distinto una sola subliste (seleccion simple)
        tipo_bbdd_seek = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][0]
        tipo_objeto_subform_seek = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][1]
        tipo_repositorio_seek = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][2]
        repositorio_seek = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][3]
        nombre_objeto_seek = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][4]
        estado_migracion = self.widget_objeto_treeview_objetos.datos_item_seleccionado["lista_datos_item_seleccionado"][0][5]

        df_script_scrolledtext = None
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
        self.widget_objeto_scrolledtext_script_a_migrar.modificaciones("borrar_contenido_y_tags")

        self.widget_objeto_scrolledtext_script_a_migrar.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                                              , df_datos = df_script_scrolledtext
                                                              , height_scrolledtext = self.widget_objeto_scrolledtext_script_a_migrar_height
                                                              , **self.kwargs_gui_tags_scrolledtext_scripts)
        
        self.widget_objeto_scrolledtext_script_a_migrar.config_atributos(**{"state": tk.DISABLED})



#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################
##  SE INICIA EL APP
#################################################################################################################################################################################
#################################################################################################################################################################################
#################################################################################################################################################################################

if __name__ == "__main__":

    #se crea el diccionario kwargs dicc_kwargs_gui que sirve para colocar todos los wigets de la gui
    dicc_kwargs_gui = {
                        #######################################################################
                        # gui_ventana_inicio
                        #######################################################################
                        "gui_ventana_inicio": #--> tiene que se el nombre exacto de la clase de la gui
                                            {"dicc_config_root":
                                                                {"title": mod_gen.nombre_app
                                                                , "iconbitmap": mod_gen.ico_app
                                                                , "tupla_geometry": (750, 570)
                                                                , "resizable": (0, 0)
                                                                }

                                            , "frames_root":
                                                    {"frame_guia_usuario":
                                                                {"frame":
                                                                        {"width": 720
                                                                        , "height": 40
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                        }

                                                                , "WIDGET_01":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label de la guia de usuario" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "guia usuario" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "width": 12
                                                                                                , "fg": "black"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 560, "coord_y": 10}
                                                                                                }
                                                                            }

                                                                , "WIDGET_02":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton asociado a la descarga de la guia de usuario" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_guia_usuario, "tupla_imagen_resize": (23, 23)}
                                                                                                , "controltiptext": "Descarga la guia de usuario"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 670, "coord_y": 10}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_GUI_guia_usuario"} #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                }
                                                                            }

                                                                }
 
                                                    , "frame_procesos":
                                                                {"frame":
                                                                        {"width": 700
                                                                        , "height": 170
                                                                        , "bg": "#ACADB1"
                                                                        , "bd": 2
                                                                        , "relief": "solid"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 50}
                                                                        }

                                                                , "WIDGET_03":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label titulo del frame 'frame_procesos'" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "PROCESOS" 
                                                                                                , "font": ("Calibri", 13, "bold")
                                                                                                , "width": 13
                                                                                                , "bd": 1
                                                                                                , "relief": "solid"
                                                                                                , "bg": "#1F40AD"
                                                                                                , "fg": "white"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_04":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label del combobox procesos" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "proceso" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "width": 12
                                                                                                , "bg": "#ACADB1"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 39}
                                                                                                }
                                                                            }

                                                                , "WIDGET_05": #el nombre de la key aqui es importante se hace referencia en este mismo diccionario (WIDGET_14)
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox de procesos" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 25
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 40}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": [mod_gen.dicc_procesos[item]["PROCESO"] for item in mod_gen.dicc_procesos.keys()]
                                                                                                , "lista_dicc_rutina_aplicar_eventos_widget":[{"tipo_bind": "<<ComboboxSelected>>"
                                                                                                                                                , "rutina": "def_GUI_combobox_proceso" #la rutina tiene que estar definida en la clase gui_ventana_inicio}
                                                                                                                                                }]
                                                                                                }
                                                                            }

                                                                , "WIDGET_06":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton asociado al combobox procesos" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_boton_procesos, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Ejecuta el proceso seleccionado"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 320, "coord_y": 38}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_gui_threads"} #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                }
                                                                            }

                                                                , "WIDGET_07":
                                                                            {"tipo_widget": "scrolledtext_propio"
                                                                            , "desc_tipo_widget": "scrolledtext que almacena la descripcion asociada al proceso (combobox) seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10, "bold")
                                                                                                , "width": 92
                                                                                                , "height": 5
                                                                                                , "state": tk.DISABLED
                                                                                                , "bg": "#B7C3F5"
                                                                                                , "fg": "black"
                                                                                                , "wrap": tk.WORD
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 70}
                                                                                                , "justify": tk.LEFT
                                                                                                , "anchor": "w"
                                                                                                }
                                                                            }

                                                                }

                                                    , "frame_ms_access":
                                                                    {"frame":
                                                                            {"width": 700
                                                                            , "height": 130
                                                                            , "bg": "#C06767"
                                                                            , "bd": 2
                                                                            , "relief": "solid"
                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 240}
                                                                            }

                                                                    , "WIDGET_08":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label titulo del frame 'frame_ms_access'" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "MS ACCESS" 
                                                                                                    , "font": ("Calibri", 13, "bold")
                                                                                                    , "width": 13
                                                                                                    , "bd": 1
                                                                                                    , "relief": "solid"
                                                                                                    , "bg": "#1F40AD"
                                                                                                    , "fg": "white"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_09":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "bbdd 01" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#C06767"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 39}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_10":
                                                                                {"tipo_widget": "entry"
                                                                                , "desc_tipo_widget": "entry BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 60
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 40}
                                                                                                    , "justify": tk.LEFT
                                                                                                    , "state": tk.DISABLED
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_11":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton add BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_add, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Agrega la bbdd 1 de MS Access"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 570, "coord_y": 38}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_ms_access" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("ADD_MS_ACCESS", "BBDD_01")   
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_12":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton clear BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_clear, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Limpia la bbdd 1 de MS Access"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 630, "coord_y": 38}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_ms_access" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("CLEAR_MS_ACCESS", "BBDD_01")    
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_13":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "bbdd 02" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#C06767"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 79}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_14":
                                                                                {"tipo_widget": "entry"
                                                                                , "desc_tipo_widget": "entry BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 60
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 80}
                                                                                                    , "justify": tk.LEFT
                                                                                                    , "state": tk.DISABLED
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_15":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton add BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_add, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Agrega la bbdd 2 de MS Access"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 570, "coord_y": 78}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_ms_access" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("ADD_MS_ACCESS", "BBDD_02")   
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_16":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton add BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_clear, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Limpia la bbdd 2 de MS Access"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 630, "coord_y": 78}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_ms_access" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("CLEAR_MS_ACCESS", "BBDD_02")    
                                                                                                                    }
                                                                                                    }
                                                                                }


                                                                    }

                                                    , "frame_sql_server":
                                                                    {"frame":
                                                                            {"width": 700
                                                                            , "height": 130
                                                                            , "bg": "#E2EE79"
                                                                            , "bd": 2
                                                                            , "relief": "solid"
                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 390}
                                                                            }

                                                                    , "WIDGET_17":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label titulo del frame 'frame_sql_server'" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "SQL SERVER" 
                                                                                                    , "font": ("Calibri", 13, "bold")
                                                                                                    , "width": 13
                                                                                                    , "bd": 1
                                                                                                    , "relief": "solid"
                                                                                                    , "bg": "#1F40AD"
                                                                                                    , "fg": "white"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_18":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label Servidor 1" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "servidor 1" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#E2EE79"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 39}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_19":
                                                                                {"tipo_widget": "combobox"
                                                                                , "desc_tipo_widget": "combobox de asociado a Servidor 1" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 25
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 40}
                                                                                                    , "justify": tk.LEFT
                                                                                                    , "combobox_lista_opciones": mod_sql_server.lista_GUI_sql_server_servidor
                                                                                                    , "lista_dicc_rutina_aplicar_eventos_widget":[{"tipo_bind": "<<ComboboxSelected>>"
                                                                                                                                                    , "rutina": "def_GUI_sql_server" #la rutina tiene que estar definida en la clase gui_ventana_inicio}
                                                                                                                                                    , "parametros_args": ("BBDD_ASOCIADAS_SERVIDOR",)
                                                                                                                                                    , "parametros_kwargs": {"opcion_servidor": "SERVIDOR_1"}
                                                                                                                                                    }]
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_20":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label bbbd 1" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "bbdd 01" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#E2EE79"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 320, "coord_y": 39}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_21":
                                                                                {"tipo_widget": "combobox"
                                                                                , "desc_tipo_widget": "combobox de asociado a bbdd del Servidor 1" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 25
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 390, "coord_y": 40}
                                                                                                    , "justify": tk.LEFT
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_22":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton clear Servidor 1 + BBDD" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_clear, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Limpia el Servidor 1 y bbdd 1 de SQL Server"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 620, "coord_y": 38}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_sql_server" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("CLEAR_SQL_SERVER_BBDD_01",)   
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_23":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label Servidor 2" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "servidor 2" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#E2EE79"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 79}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_24":
                                                                                {"tipo_widget": "combobox"
                                                                                , "desc_tipo_widget": "combobox de asociado a Servidor 2" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 25
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 80}
                                                                                                    , "justify": tk.LEFT
                                                                                                    , "combobox_lista_opciones": mod_sql_server.lista_GUI_sql_server_servidor
                                                                                                    , "lista_dicc_rutina_aplicar_eventos_widget":[{"tipo_bind": "<<ComboboxSelected>>"
                                                                                                                                                    , "rutina": "def_GUI_sql_server" #la rutina tiene que estar definida en la clase gui_ventana_inicio}
                                                                                                                                                    , "parametros_args": ("BBDD_ASOCIADAS_SERVIDOR",)
                                                                                                                                                    , "parametros_kwargs": {"opcion_servidor": "SERVIDOR_2"}
                                                                                                                                                    }]
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_25":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label bbbd 1" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "bbdd 02" 
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "width": 12
                                                                                                    , "bg": "#E2EE79"
                                                                                                    , "fg": "black"
                                                                                                    , "anchor": "w"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 320, "coord_y": 79}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_26":
                                                                                {"tipo_widget": "combobox"
                                                                                , "desc_tipo_widget": "combobox de asociado a bbdd del Servidor 2" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 25
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 390, "coord_y": 80}
                                                                                                    , "justify": tk.LEFT
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_27":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "boton clear Servidor 2 + BBDD" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 40
                                                                                                    , "dicc_imagen": {"png_imagen": mod_gen.img_boton_clear, "tupla_imagen_resize": (33, 23)}
                                                                                                    , "controltiptext": "Limpia el Servidor 2 y bbdd 2 de SQL Server"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 620, "coord_y": 78}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_GUI_sql_server" #la rutina tiene que estar definida en la clase gui_ventana_inicio
                                                                                                                    , "parametros_args": ("CLEAR_SQL_SERVER_BBDD_02",)   
                                                                                                                    }
                                                                                                    }
                                                                                }
                                                                    }

                                                    , "frame_resolucion_pantalla":
                                                                    {"frame":
                                                                            {"width": 700
                                                                            , "height": 30
                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 530}
                                                                            }

                                                                    , "WIDGET_28": #el nombre de la key aqui es importante se hace referencia mas abajo despues de declarar este diccionario
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label resolucion pantalla'" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "RESOLUCIÓN PANTALLA:"
                                                                                                    , "font": ("Calibri", 11, "bold")
                                                                                                    , "width": 20
                                                                                                    , "fg": "BLACK"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                    , "anchor": "w"
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_29": #el nombre de la key aqui es importante se hace referencia mas abajo despues de declarar este diccionario
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label resolucion pantalla'" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 11, "bold")
                                                                                                    , "width": 50
                                                                                                    , "fg": "red"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 160, "coord_y": 0}
                                                                                                    , "anchor": "w"
                                                                                                    }
                                                                                }
                                                                    }
                                                    }
                                            }

                        #######################################################################
                        # gui_sql_server_authentication
                        #######################################################################
                        , "gui_sql_server_authentication": #--> tiene que se el nombre exacto de la clase de la gui
                                            {"dicc_config_root":
                                                                {"title": "ACCESO SQL SERVER"
                                                                , "iconbitmap": mod_gen.ico_app
                                                                , "tupla_geometry": (400, 120)
                                                                , "mantener_nueva_ventana_encima_otras": True
                                                                , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                , "resizable": (0, 0)
                                                                }

                                            , "frames_root":
                                                    {"frame_inicio":
                                                                {"frame":
                                                                        {"width": 400
                                                                        , "height": 120
                                                                        , "bg": "#B9AA79"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                        }

                                                                , "WIDGET_30":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label servidor" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "servidor" 
                                                                                                , "font": ("Calibri", 11, "bold")
                                                                                                , "width": 13
                                                                                                , "bg": "#B9AA79"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 20}
                                                                                                }
                                                                            }

                                                                , "WIDGET_31":
                                                                            {"tipo_widget": "entry"
                                                                            , "desc_tipo_widget": "entry servidor" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 20}
                                                                                                , "justify": tk.LEFT
                                                                                                , "state": tk.DISABLED
                                                                                                }
                                                                            }

                                                                , "WIDGET_32":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label login" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "login" 
                                                                                                , "font": ("Calibri", 11, "bold")
                                                                                                , "width": 13
                                                                                                , "bg": "#B9AA79"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 50}
                                                                                                }
                                                                            }

                                                                , "WIDGET_33":
                                                                            {"tipo_widget": "entry"
                                                                            , "desc_tipo_widget": "entry login" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 50}
                                                                                                , "justify": tk.LEFT
                                                                                                }
                                                                            }

                                                                , "WIDGET_34":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label password" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "password" 
                                                                                                , "font": ("Calibri", 11, "bold")
                                                                                                , "width": 13
                                                                                                , "bg": "#B9AA79"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 80}
                                                                                                }
                                                                            }

                                                                , "WIDGET_35":
                                                                            {"tipo_widget": "entry"
                                                                            , "desc_tipo_widget": "entry password" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 80}
                                                                                                , "justify": tk.LEFT
                                                                                                }
                                                                            }

                                                                , "WIDGET_36":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton sql server authentication" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_boton_sql_server_authentication, "tupla_imagen_resize": (23, 23)}
                                                                                                , "controltiptext": "Crea la connecting string a SQL Server y la guarda en memoria"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 340, "coord_y": 48}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_GUI_conexion_servidor_sql_server"} #la rutina tiene que estar definida en la clase gui_sql_server_authentication
                                                                                                }
                                                                            }

                                                                }
                                                    }
                                            }

                        #######################################################################
                        # gui_diagnostico_bbdd_sql_server
                        #######################################################################
                        , "gui_diagnostico_bbdd_sql_server": #--> tiene que se el nombre exacto de la clase de la gui
                                            {"dicc_config_root":
                                                                {"title": "DEPENDENCIAS SQL SERVER"
                                                                , "iconbitmap": mod_gen.ico_app
                                                                , "tupla_geometry": (310, 290)
                                                                , "mantener_nueva_ventana_encima_otras": True
                                                                , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                , "resizable": (0, 0)
                                                                }

                                            , "frames_root":
                                                    {"frame_inicio":
                                                                {"frame":
                                                                        {"width": 310
                                                                        , "height": 290
                                                                        , "bg": "#64ACC2"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                        }

                                                                , "WIDGET_37":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label seleccion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "selección" 
                                                                                                , "font": ("Calibri", 11, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 20}
                                                                                                }
                                                                            }

                                                                , "WIDGET_38":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox de seleccion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 20
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 120, "coord_y": 20}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": mod_sql_server.lista_GUI_diagnostico_combobox_sql_server
                                                                                                }
                                                                            }

                                                                , "WIDGET_39":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label BBDD" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "bases de datos" 
                                                                                                , "font": ("Calibri", 11, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 56}
                                                                                                }
                                                                            }

                                                                , "WIDGET_40":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton (des)seleccionar todas la bbdd" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_seleccionar_all_none, "tupla_imagen_resize": (20, 20)}
                                                                                                , "controltiptext": "Selecciona o des-selecciona todas las bases de datos SQL Server"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 235, "coord_y": 50}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_GUI_sql_server_diagnostico_listbox_all_none"} #la rutina tiene que estar definida en la clase gui_diagnostico_bbdd_sql_server
                                                                                                }
                                                                            }

                                                                , "WIDGET_41":
                                                                            {"tipo_widget": "listbox"
                                                                            , "desc_tipo_widget": "listbox de seleccion de bbdd" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 43
                                                                                                , "height": 10
                                                                                                , "selectmode": "multiple"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 80}
                                                                                                }
                                                                            }

                                                                , "WIDGET_42":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton ejecucion diagnostico" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_boton_dependencias_sql_server, "tupla_imagen_resize": (20, 20)}
                                                                                                , "controltiptext": "Ejecuta el diagnostico de dependencias de objetos entre las bases de datos seleccionadas"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 130, "coord_y": 255}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_GUI_sql_server_diagnostico_boton_check"} #la rutina tiene que estar definida en la clase gui_diagnostico_bbdd_sql_server
                                                                                                                
                                                                                                }
                                                                            }

                                                                }

                                                    }
                                            }

                        #######################################################################
                        # gui_ventana_control_versiones
                        #######################################################################
                        , "gui_ventana_control_versiones": #--> tiene que se el nombre exacto de la clase de la gui
                                            {"dicc_config_root":
                                                                {"title": "CONTROL VERSIONES - MS ACCESS & SQL SERVER"
                                                                , "iconbitmap": mod_gen.ico_app
                                                                , "tupla_geometry": (1670, 780)
                                                                , "mantener_nueva_ventana_encima_otras": True
                                                                , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                , "resizable": (0, 0)
                                                                }

                                            , "frames_root":
                                                    {"frame_inicio":
                                                                {"frame":
                                                                        {"width": 1670
                                                                        , "height": 780
                                                                        , "bg": "#64ACC2"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                        }

                                                                , "WIDGET_43":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label tipo objeto" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "tipo objeto" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 20}
                                                                                                }
                                                                            }

                                                                , "WIDGET_44":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox tipo objeto" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 140, "coord_y": 20}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": []#se recalcula en el constructor de la clase gui_ventana_control_versiones
                                                                                                , "lista_dicc_rutina_trace_variable_enlace":
                                                                                                                [{"tipo_trace": "write"
                                                                                                                    , "rutina": "def_control_versiones_cambio_tipo_objeto"
                                                                                                                    } #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                                ]
                                                                                                }
                                                                            }

                                                                , "WIDGET_45":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label tipo concepto" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "tipo concepto" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 50}
                                                                                                }
                                                                            }

                                                                , "WIDGET_46":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox tipo concepto" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 140, "coord_y": 50}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": mod_gen.lista_GUI_seleccion_tipo_concepto,
                                                                                                }
                                                                            }

                                                                , "WIDGET_47":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton actualizar subform objetos control versiones" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_ver, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Actualiza el subformulario de objetos con variaciones entre las 2 bases de datos"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 400, "coord_y": 35}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_control_versiones_frame_inicio_click_boton" #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                                , "parametros_args": ("ACTUALIZAR_SUBFORM_OBJETOS_CONTROL_VERSIONES",)    
                                                                                                                }
                                                                                                }
                                                                            }

                                                                , "WIDGET_48":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton exportar a excel" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_excel, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Exporta a excel los objetos con variaciones entre las 2 bases de datos"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 480, "coord_y": 35}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_control_versiones_frame_inicio_click_boton" #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                                , "parametros_args": ("EXPORTAR_EXCEL_OBJETOS_CONTROL_VERSIONES",)    
                                                                                                                }
                                                                                                }
                                                                            }


                                                                , "WIDGET_49":
                                                                            {"tipo_widget": "treeview"
                                                                            , "desc_tipo_widget": "treeview con los objetos control versiones" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 100}
                                                                                                , "dicc_treeview": {"seleccion_item": "simple"
                                                                                                                    , "height": 5
                                                                                                                    ,"columnas_df": ["TIPO_BBDD", "TIPO_OBJETO_SUBFORM", "REPOSITORIO"
                                                                                                                                    , "NOMBRE_OBJETO", "NUM_CAMBIOS_SCRIPT_1", "NUM_CAMBIOS_SCRIPT_2"]

                                                                                                                    , "columnas_treeview": ["TIPO BBDD", "TIPO OBJETO", "REPOSITORIO", "OBJETO", "BBDD_01", "BBDD_02"]

                                                                                                                    , "width_columnas_treeview": [140, 140, 200, 200, 60, 60]
                                                                                                                    #las listas almacenadas en las keys columnas_df, columnas_treeview y width_columnas_treeview
                                                                                                                    #han de tener la misma longitud
                                                                                                                    }
                                                                                                , "dicc_rutina_click_item": {
                                                                                                                            #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                                            "rutina": "def_control_versiones_update_subform_objetos_click_item"}
                                                                                                }
                                                                            }

                                                                , "WIDGET_50":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label titulo script BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "BBDD 01" 
                                                                                                , "font": ("Calibri", 14, "bold")
                                                                                                , "width": 10
                                                                                                , "bd": 1
                                                                                                , "relief": "solid"
                                                                                                , "bg": "black"
                                                                                                , "fg": "white"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 243}
                                                                                                }
                                                                            }

                                                                , "WIDGET_51":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label nombre BBDD_01" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 14, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "white"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 130, "coord_y": 243}
                                                                                                }
                                                                            }


                                                                , "WIDGET_52":
                                                                            {"tipo_widget": "scrolledtext_propio"
                                                                            , "desc_tipo_widget": "scrolledtext del script 1" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 11)
                                                                                                , "width": 113
                                                                                                , "height": 27
                                                                                                , "state": tk.DISABLED
                                                                                                , "bg": "#DDE1E2"
                                                                                                , "wrap": tk.NONE
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 270}
                                                                                                , "colocacion_scrollbar_horizontal": {"metodo": "place", "coord_x": 765, "coord_y": 250}
                                                                                                , "justify": tk.LEFT
                                                                                                , "anchor": "w"
                                                                                                }
                                                                            }


                                                                , "WIDGET_53":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label titulo script BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "BBDD 02" 
                                                                                                , "font": ("Calibri", 14, "bold")
                                                                                                , "width": 10
                                                                                                , "bd": 1
                                                                                                , "relief": "solid"
                                                                                                , "bg": "red"
                                                                                                , "fg": "white"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 850, "coord_y": 243}
                                                                                                }
                                                                            }

                                                                , "WIDGET_54":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label nombre BBDD_02" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 14, "bold")
                                                                                                , "bg": "#64ACC2"
                                                                                                , "fg": "white"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 960, "coord_y": 243}
                                                                                                }
                                                                            }

                                                                , "WIDGET_55":
                                                                            {"tipo_widget": "scrolledtext_propio"
                                                                            , "desc_tipo_widget": "scrolledtext del script 2" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 11)
                                                                                                , "width": 113
                                                                                                , "height": 27
                                                                                                , "state": tk.DISABLED
                                                                                                , "bg": "#DDE1E2"
                                                                                                , "wrap": tk.NONE
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 850, "coord_y": 270}
                                                                                                , "colocacion_scrollbar_horizontal": {"metodo": "place", "coord_x": 1595, "coord_y": 250}
                                                                                                , "justify": tk.LEFT
                                                                                                , "anchor": "w"
                                                                                                }
                                                                            }
                                                                }

                                                , "frame_merge":
                                                                {"frame":
                                                                        {"width": 800
                                                                        , "height": 130
                                                                        , "bg": "#73C58C"
                                                                        , "bd": 2
                                                                        , "relief": "solid"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 850, "coord_y": 100}
                                                                        }

                                                                , "WIDGET_56":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label titulo frame" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "MERGE" 
                                                                                                , "font": ("Calibri", 14, "bold")
                                                                                                , "width": 8
                                                                                                , "bd": 1
                                                                                                , "relief": "solid"
                                                                                                , "bg": "#1F40AD"
                                                                                                , "fg": "white"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_57":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label leyenda colores" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "LEYENDA COLORES:" 
                                                                                                , "font": ("Calibri", 10, "bold")
                                                                                                , "bg": "#73C58C"
                                                                                                , "fg": "black"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 265, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_58":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label cambios localizados" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "cambios localizados" 
                                                                                                , "font": ("Calibri", 10, "bold")
                                                                                                , "bg": "#05FB27"
                                                                                                , "fg": "black"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 387, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_59":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label agregado en bbdd MERGE" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "agregado en bbdd merge" 
                                                                                                , "font": ("Calibri", 10, "bold")
                                                                                                , "bg": "#FBCB05"
                                                                                                , "fg": "black"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 500, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_60":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label eliminado en bbdd MERGE" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "eliminado en bbdd merge" 
                                                                                                , "font": ("Calibri", 10, "bold")
                                                                                                , "bg": "#05FBF0"
                                                                                                , "fg": "black"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 647, "coord_y": 0}
                                                                                                }
                                                                            }

                                                                , "WIDGET_61":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label combobox bbdd" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "bbdd" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#73C58C"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 70}
                                                                                                }
                                                                            }

                                                                , "WIDGET_62":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox bbdd" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 10
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 80, "coord_y": 70}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": mod_gen.lista_GUI_seleccion_bbdd
                                                                                                , "lista_dicc_rutina_aplicar_eventos_widget":[{"tipo_bind": "<<ComboboxSelected>>"
                                                                                                                                                , "rutina": "def_proceso_merge_combobox_bbdd_selecc"
                                                                                                                                                #la rutina tiene que estar definida en la clase gui_ventana_control_versiones}
                                                                                                                                                }]
                                                                                                }
                                                                            }

                                                                , "WIDGET_63":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label combobox accion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "acción" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#73C58C"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 200, "coord_y": 70}
                                                                                                }
                                                                            }

                                                                , "WIDGET_64":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox accion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 15
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 270, "coord_y": 70}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": [] #se asigna al ejecutar la rutina def_proceso_merge_combobox_bbdd_selecc
                                                                                                }
                                                                            }

                                                                , "WIDGET_65":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label lineas origen" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "lineas Origen" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#73C58C"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 420, "coord_y": 55}
                                                                                                }
                                                                            }

                                                                , "WIDGET_66":
                                                                            {"tipo_widget": "entry_propio"
                                                                            , "desc_tipo_widget": "entry lineas origen (desde)" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 5
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 530, "coord_y": 55}
                                                                                                , "dicc_entry":
                                                                                                            {"formato_validacion": "entero_positivo"
                                                                                                            , "titulo_messagebox_warning": "LINEAS ORIGEN (DESDE)"
                                                                                                            }
                                                                                                }
                                                                            }

                                                                , "WIDGET_67":
                                                                            {"tipo_widget": "entry_propio"
                                                                            , "desc_tipo_widget": "entry lineas origen (hasta)" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 5
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 570, "coord_y": 55}
                                                                                                , "dicc_entry":
                                                                                                            {"formato_validacion": "entero_positivo"
                                                                                                            , "titulo_messagebox_warning": "LINEAS ORIGEN (HASTA)"
                                                                                                            }
                                                                                                }
                                                                            }


                                                                , "WIDGET_68":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label lineas destino" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "lineas Destino" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#73C58C"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 420, "coord_y": 85}
                                                                                                }
                                                                            }

                                                                , "WIDGET_69":
                                                                            {"tipo_widget": "entry_propio"
                                                                            , "desc_tipo_widget": "entry lineas destino" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 12
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 530, "coord_y": 85}
                                                                                                , "dicc_entry":
                                                                                                            {"formato_validacion": "entero_positivo"
                                                                                                            , "titulo_messagebox_warning": "LINEAS DESTINO"
                                                                                                            }
                                                                                                }
                                                                            }

                                                                , "WIDGET_70":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton ACCION" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_migrar_lineas_codigo, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Ejecuta migraciones de lineas de código entre los scripts de objetos de las 2 bases de datos"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 640, "coord_y": 65}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_click_proceso_merge_boton_merge"} #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                }
                                                                            }

                                                                , "WIDGET_71":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton acceso a clase MERGE" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_merge_bbdd_fisica, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Permite acceder a la ventana para realizar los MERGE en base de datos fisica"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 720, "coord_y": 65}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_click_proceso_merge_boton_cambios_en_bbdd"} #la rutina tiene que estar definida en la clase gui_ventana_control_versiones
                                                                                                }
                                                                            }

                                                                }



                                                    }
                                            }

                        #######################################################################
                        # gui_ventana_merge_bbdd_fisicas
                        #######################################################################
                        , "gui_ventana_merge_bbdd_fisicas": #--> tiene que se el nombre exacto de la clase de la gui
                                            {"dicc_config_root":
                                                                {"title": "MERGE BBDD FISICA - MS ACCESS & SQL SERVER"
                                                                , "iconbitmap": mod_gen.ico_app
                                                                , "tupla_geometry": (1130, 585)
                                                                , "mantener_nueva_ventana_encima_otras": True
                                                                , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                , "resizable": (0, 0)
                                                                }

                                            , "frames_root":
                                                    {"frame_inicio":
                                                                {"frame":
                                                                        {"width": 1130
                                                                        , "height": 585
                                                                        , "bg": "#829ADB"
                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                        }

                                                                , "WIDGET_72":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label tipo seleccion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "tipo selección" 
                                                                                                , "font": ("Calibri", 12, "bold")
                                                                                                , "bg": "#829ADB"
                                                                                                , "fg": "black"
                                                                                                , "anchor": "w"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 20}
                                                                                                }
                                                                            }

                                                                , "WIDGET_73":
                                                                            {"tipo_widget": "combobox"
                                                                            , "desc_tipo_widget": "combobox tipo seleccion" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 10)
                                                                                                , "width": 30
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 140, "coord_y": 20}
                                                                                                , "justify": tk.LEFT
                                                                                                , "combobox_lista_opciones": []#se recalcula en el constructor de la clase gui_ventana_control_versiones
                                                                                                }
                                                                            }

                                                                , "WIDGET_74":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton actualizar subform objetos treeview" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_ver, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Actualiza el subformulario con los objetos a migrar en base de datos fisica"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 400, "coord_y": 18}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_merge_bbdd_fisicas_click_boton" #la rutina tiene que estar definida en la clase gui_ventana_merge_bbdd_fisicas
                                                                                                                , "parametros_args": ("ACTUALIZAR_SUBFORM_CONTROL_VERSIONES_MERGE_BBDD_FISICA",)    
                                                                                                                }
                                                                                                }
                                                                            }

                                                                , "WIDGET_75":
                                                                            {"tipo_widget": "button"
                                                                            , "desc_tipo_widget": "boton realizar merge en bbdd fisica" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"width": 40
                                                                                                , "dicc_imagen": {"png_imagen": mod_gen.img_control_versiones_boton_merge_bbdd_fisica, "tupla_imagen_resize": (33, 23)}
                                                                                                , "controltiptext": "Realiza el merge en base de datos fisica y genera un log de objetos migrados con los ok y ko"
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 500, "coord_y": 18}
                                                                                                , "dicc_rutina":
                                                                                                                {"rutina": "def_merge_bbdd_fisicas_click_boton" #la rutina tiene que estar definida en la clase gui_ventana_merge_bbdd_fisicas
                                                                                                                , "parametros_args": ("REALIZAR_MERGE_BBDD_FISICA",)  
                                                                                                                }
                                                                                                }
                                                                            }

                                                                , "WIDGET_76":
                                                                            {"tipo_widget": "treeview"
                                                                            , "desc_tipo_widget": "treeview con los objetos a migrar en bbdd fisica" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 70}
                                                                                                , "dicc_treeview": {"seleccion_item": "simple"
                                                                                                                    ,"height": 5
                                                                                                                    ,"columnas_df": ["TIPO_BBDD", "TIPO_OBJETO_SUBFORM", "TIPO_REPOSITORIO"
                                                                                                                                     , "REPOSITORIO", "NOMBRE_OBJETO", "ESTADO_MIGRACION"]

                                                                                                                    , "columnas_treeview": ["TIPO BBDD", "TIPO OBJETO", "TIPO_REPOSITORIO"
                                                                                                                                            , "REPOSITORIO", "OBJETO", "ESTADO MIGRACIÓN"]

                                                                                                                    , "width_columnas_treeview": [140, 140, 200, 200, 200, 200]
                                                                                                                    #las listas almacenadas en las keys columnas_df, columnas_treeview y width_columnas_treeview
                                                                                                                    #han de tener la misma longitud
                                                                                                                    }
                                                                                                , "dicc_rutina_click_item": {
                                                                                                                            #la rutina tiene que estar definida en la clase gui_ventana_merge_bbdd_fisicas
                                                                                                                            "rutina": "def_merge_bbdd_fisicas_update_subform_objetos_click_item"

                                                                                                                            #parametros_args se asocia al valor que toma el combobox WIDGET_73 (combobox)
                                                                                                                            #se pasa el atributo widget_objeto_combobox_tipo_seleccion de la clase que almacena el widget (objeto) 
                                                                                                                            #(es tupla de 1 solo elemento) y el boton coje el parametro cuando el usuario interactua con el despues de su creacion
                                                                                                                            #de ahi el "lambda widget"
                                                                                                                            , "parametros_args": (lambda widget: 
                                                                                                                                widget.widget_objeto_combobox_tipo_seleccion.widget_objeto.get(),)
                                                                                                                            }
                                                                                                }
                                                                            }

                                                                , "WIDGET_77":
                                                                            {"tipo_widget": "label"
                                                                            , "desc_tipo_widget": "label titulo script a migrar" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"text": "SCRIPT A MIGRAR" 
                                                                                                , "font": ("Calibri", 14, "bold")
                                                                                                , "width": 15
                                                                                                , "bd": 1
                                                                                                , "relief": "solid"
                                                                                                , "bg": "black"
                                                                                                , "fg": "white"
                                                                                                , "justify": tk.CENTER
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 218}
                                                                                                }
                                                                            }

                                                                , "WIDGET_78":
                                                                            {"tipo_widget": "scrolledtext_propio"
                                                                            , "desc_tipo_widget": "scrolledtext del script a migrar" #key informativa (no se usa en el resto del codigo del app)
                                                                            , "kwargs_config":
                                                                                                {"font": ("Calibri", 11)
                                                                                                , "width": 153
                                                                                                , "height": 17
                                                                                                , "state": tk.DISABLED
                                                                                                , "bg": "#DDE1E2"
                                                                                                , "wrap": tk.NONE
                                                                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 245}
                                                                                                , "colocacion_scrollbar_horizontal": {"metodo": "place", "coord_x": 1040, "coord_y": 225}
                                                                                                , "justify": tk.LEFT
                                                                                                , "anchor": "w"
                                                                                                }
                                                                            }



                                                                }
                                                    }
                                            }

                        }


    #se crea el diccionario kwargs dicc_kwargs_gui_tags_scrolledtext_scripts que sirve para los tags en los scrolledtext de scripts
    dicc_kwargs_gui_tags_scrolledtext_scripts = {"columna_df_para_informar": "CODIGO_CON_NUM_LINEA"
                                                 
                                                , "lista_dicc_tag_linea_completa":
                                                            [{"nombre_tag": "CAMBIOS_LOCALIZADOS"
                                                                , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                , "case_sensitive": False
                                                                , "dicc_config": {"background": "#05FB27"}
                                                                }

                                                            , {"nombre_tag": "AGREGADO"
                                                                , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                , "case_sensitive": True
                                                                , "dicc_config": {"background": "#FBCB05"}
                                                                }

                                                            , {"nombre_tag": "ELIMINADO"
                                                                , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                , "case_sensitive": True
                                                                , "dicc_config": {"background": "#05FBF0"}
                                                                }
                                                            ]

                                                , "lista_dicc_tag_caracteres_cambiantes_comparativa":
                                                            [{"nombre_tag": "CAMBIOS_LOCALIZADOS_POR_INDICES"
                                                                , "columna_df_filtro_registros_aplicar_tag": "CONTROL_CAMBIOS_ACTUAL"
                                                                , "columna_df_filtro_registros_aplicar_tag_valor": "CAMBIOS_LOCALIZADOS"
                                                                , "columna_df_comparar_1": "CODIGO_CON_NUM_LINEA"
                                                                , "columna_df_comparar_2": "CODIGO_CON_NUM_LINEA_OTRA_BBDD"
                                                                , "case_sensitive": True
                                                                , "marcar_toda_linea_si_todo_varia": False
                                                                , "dicc_config": {"foreground": "red"}
                                                                }
                                                            ]            
                                                }


    #se crea el root usando el kwargs dicc_config_root del diccionario creado anteriormente
    #y se inicia la clase gui_gui_app
    kwargs_gui_ventana_inicio = dicc_kwargs_gui["gui_ventana_inicio"]
    kwargs_gui_root = dicc_kwargs_gui["gui_ventana_inicio"]["dicc_config_root"]

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_gui_root)
    root.config_atributos(**kwargs_gui_ventana_inicio)

    #se inicializa la clase gui_ventana_inicio (se pasa como kwargs el diccionario dicc_kwargs_gui al completo)
    gui_ventana_inicio(root, kwargs_gui_tags_scrolledtext_scripts = dicc_kwargs_gui_tags_scrolledtext_scripts, **dicc_kwargs_gui)

    root.widget_objeto.mainloop()








