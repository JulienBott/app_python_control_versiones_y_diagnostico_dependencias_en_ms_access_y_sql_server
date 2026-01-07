import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import messagebox, filedialog as fd
from typing import Literal
from itertools import product
from PIL import Image, ImageTk
import difflib
import re
import pkgutil
import inspect
import os
import sys
import pandas as pd


###########################################################################################
###########################################################################################
# DISCLAIMER
###########################################################################################
###########################################################################################

#este modulo contiene una version muy muy reducida y no documentada del proyecto de refactorizacion 
#basado en herencias de clases de tkinter que comento en el README de mi 1er proyecto publico en github
#(app_python_control_versiones_y_diagnostico_dependencias_en_ms_access_y_sql_server)
#
#solo incluye lo estrictamente necesario para que funcione el app
###########################################################################################
###########################################################################################



###########################################################################################
#     CLASE MADRE - gui_tkinter_widgets
###########################################################################################
class gui_tkinter_widgets():

    def __init__(self, master, tipo_widget_param = None, self_clase_gui_donde_call_rutina = None, **kwargs_config_parametros):

        """
        Clase que permite crear widgets tkinter con metodos nativos y propios.

        args
        --
        --> tipo_widget_param\n
        Es el string del objeto tkinter que se va a crear (la clase tiene un mecanismo interno para localizar el modulo tkinter
        y asimismo como se ha de declarar el objeto para que se cree correctamente independientemente de que se haya declarado
        con/sin mayusculas/minusculas).

        --> self_clase_gui_donde_call_rutina\n
        Es la clase padre donde se usa en el proyecto de GUI donde se incorpora este modulo (el self)

        kwargs
        --
        Diccionario con los parametros de configuración.

        """


        #se asigna a master el widget_objeto cuando el root se crea desde la clase
        #si no se hace los frames que se declaren con la clase gui_tkinter_widgets
        #se tienen que configurar como "master = root.widget_objeto"
        if isinstance(master, self.__class__):
            master = master.widget_objeto


        #se inicializan los atributos siguientes para poder usarlos en metodos publicos o privados dentro de la clase que se crea:
        # master                           repositorio GUI donde se ubica el widget (root, frame etc etc)
        # widget_objeto                    objeto widget que se crea con la presente clase (None de inicio, se asigna mas adelante) 
        # tipo_widget_param                parametro de la clase
        # tipo_widget_lower_no_blank       parametro tipo_widget_param de la clase que se crea en minusculas y trimeado (correjido al nombre "tk" si es root)
        # clase_objeto                     objeto de la clase que se crea
        # clase_nombre                     nombre de la clase que se crea
        # nombre_alias_tkinter_import      alias (nombre) de la libreria tkinter importada en este modulo .py
        # objeto_alias_tkinter_import      alias (objeto) de la libreria tkinter importada en este modulo .py
        # nombre_modulo_python_py          nombre del presente modulo .py
        self.master = master
        self.tipo_widget_param = tipo_widget_param

        self.tipo_widget_lower_no_blank = (None if tipo_widget_param is None
                                           else tipo_widget_param.lower().replace(" ", "")
                                           if tipo_widget_param.lower().replace(" ", "") != "root" else "tk")
        

        self.widget_objeto  = None#se asigna mas adelante
        self.clase_objeto = self.__class__
        self.clase_nombre = self.__class__.__name__
        self.nombre_alias_tkinter_import = self.__varios_clase_madre("nombre_alias_libreria_python_import", libreria_python = "tkinter")
        self.objeto_alias_tkinter_import = self.__varios_clase_madre("objeto_alias_libreria_python_import", libreria_python = "tkinter")
        self.nombre_modulo_python_py = self.__varios_clase_madre("nombre_modulo_python_py")
        self.kwargs_config_parametros = kwargs_config_parametros
        self.self_clase_gui_donde_call_rutina = self_clase_gui_donde_call_rutina


        #se inicializa como atributos los kwargs nativos y propios (se calculan en el metodo config_atributos de la presente clase)
        #donde los atributos se netean a minusculas sin espacios en blanco y a posteriori se fusionan en kwargs_config_parametros
        self.kwargs_config_parametros_lower_no_blank = None
        self.kwargs_config_atributos_nativos = None
        self.kwargs_config_atributos_propios = None
        self.kwargs_config_parametros = None #son los kwargs donde los artibutos nativos salen antes que los propios



        #se asigna el nombre delos atributos propios
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs_dicc_config_root = "dicc_config_root"
        self.nombre_kwargs_dicc_config_root_tupla_geometry = "tupla_geometry"

        self.nombre_kwargs_colocacion_dicc = "colocacion_dicc"
        self.nombre_kwargs_colocacion_dicc_metodo = "metodo"
        self.nombre_kwargs_colocacion_dicc_coord_x = "coord_x"
        self.nombre_kwargs_colocacion_dicc_coord_y = "coord_y"

        self.nombre_kwargs_combobox_lista_opciones = "combobox_lista_opciones"

        self.nombre_kwargs_dicc_rutina = "dicc_rutina"
        self.nombre_kwargs_dicc_rutina_nombre_rutina = "rutina"
        self.nombre_kwargs_dicc_rutina_tipo_rutina = "tipo_rutina"
        self.nombre_kwargs_dicc_rutina_tipo_bind = "tipo_bind"
        self.nombre_kwargs_dicc_rutina_parametros_args = "parametros_args"
        self.nombre_kwargs_dicc_rutina_parametros_kwargs = "parametros_kwargs"
        self.nombre_tipo_rutina_asociada_widget_command = "command"
        self.nombre_tipo_rutina_asociada_widget_event = "event"

        self.nombre_imagen_tupla_resize = "imagen_tupla_resize"
        self.nombre_controltiptext = "controltiptext"

        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace = "lista_dicc_rutina_trace_variable_enlace"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina = "rutina"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace = "tipo_trace"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args = "parametros_args"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs = "parametros_kwargs"



        self.nombre_kwargs_alineacion = "alineacion"

        self.nombre_kwargs_bloquear = "bloquear"

        self.nombre_kwargs_listbox_lista_items = "listbox_lista_items"
        self.nombre_kwargs_listbox_lista_items_seleccionados = "listbox_lista_items_seleccionados"
        self.nombre_kwargs_listbox_listbox_seleccionar_todo_o_nada = "listbox_seleccionar_todo_o_nada"


        self.imagen_en_boton_tupla_resize_por_defecto = (20, 20)



        #se inicializa como atributo una lista de atributos que requieren una variable de enlace (StringVar) como valor
        self.lista_atributos_vinculados_variable_enlace = ["textvariable", "listvariable"]


        #se asigna un stringvar al objeto (aunque algunos widgets no lo admiten como atributo
        #mediante un try except pass se descarta mas abajo los tipos de widgets que no lo admiten)
        #(en caso de necesitar un IntVar la conversion ha de hacerse desde el proyecto en si donde este integrado este modulo .py)
        #
        #se asigna la variable_enlace tan solo si la clase madre no se usa para crear el root (sino da error)
        self.variable_enlace = self.objeto_alias_tkinter_import.StringVar() if self.tipo_widget_lower_no_blank not in ["root", "tk"] else None




        #se crea el widget mediante la rutina interna __varios_clase_madre (opcion = crear_widget_objeto)
        if self.tipo_widget_lower_no_blank is not None:
            self.__varios_clase_madre("crear_widget_objeto")



        # CONFIGURACION NATIVOS + ATRIBUTOS PROPIOS
        ###########################################################################################
        self.config_atributos(**kwargs_config_parametros)



    ###########################################################################################
    #     ATRIBUTOS  Y NATIVOS - RE-USABLES UNA VEZ CREADO EL OBJETO
    ###########################################################################################

    def config_atributos(self, **kwargs_config_parametros):

        """
        Metodo que permite configurar los atributos (nativos y/o propios) despues de haber creado el widget.

        kwargs
        --
        --> Diccionario con los parametros de configuración.
        """


        ###################################################################################################################################
        # ATRIBUTOS NATIVOS (realizar siempre la configuracion de los atributos nativos antes que los propios)
        ###################################################################################################################################

        #se configuran los atributos nativos (mediante config o configure) una vez creado el widget
        #se actualizan los atributos declarados en el constructor de la clase:
        # --> kwargs_config_atributos_nativos
        # --> kwargs_config_atributos_propios
        # --> kwargs_config_parametros

        self.kwargs_config_parametros_lower_no_blank = {key.lower().replace(" ", ""): valor for key, valor in kwargs_config_parametros.items()}

        if len(self.kwargs_config_parametros_lower_no_blank) != 0:

            self.kwargs_config_atributos_nativos = {}
            for atributo_config, valor_atributo_config in self.kwargs_config_parametros_lower_no_blank.items():

                #caso de imagenes a incorporar al widget
                if atributo_config == "image" and isinstance(valor_atributo_config, str):
                    
                    try:
                        imagen_tupla_resize = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_imagen_tupla_resize, self.imagen_en_boton_tupla_resize_por_defecto)
                        img = Image.open(valor_atributo_config).resize(imagen_tupla_resize, Image.LANCZOS)#se redimensiona la imagen sino los botones salen distorsionados en la GUI
                        img_tk = ImageTk.PhotoImage(img)

                        self.widget_objeto.image = img_tk
                        dicc_config = {"image": img_tk}
                    except:
                        pass

                #caso general
                else:
                    dicc_config = {atributo_config: valor_atributo_config}


                try:
                    self.widget_objeto.config(**dicc_config)

                except (AttributeError, TypeError):
                    #por si se hace una llamada a una clase hija herencia de esta clase gui_tkinter_widgets
                    #donde no se pasa por parametro el tipo de widget por lo que da el error en estos casos
                    #(AttributeError: 'NoneType' object has no attribute 'config')

                    pass


                except self.objeto_alias_tkinter_import.TclError as _:

                    try:
                        self.widget_objeto.configure(**dicc_config)

                    except self.objeto_alias_tkinter_import.TclError as Err: 
                        pass

                    else:
                        #el dicc_config que no da error pasa a formar parte de kwargs_config_atributos_nativos
                        self.kwargs_config_atributos_nativos.update(dicc_config)

                    pass

                else:
                    #el dicc_config que no da error pasa a formar parte de kwargs_config_atributos_nativos
                    self.kwargs_config_atributos_nativos.update(dicc_config)


                #cuando el atributo de la iteracion sobre lista_kwargs_config_atributos_nativos es 'text', 
                #puesto que todos los widgets creados con esta clase madre tienen variable_enlace (salvo el root)
                #se asigna el valor del atributo a esta variable_enlace
                if isinstance(dicc_config, dict) and dicc_config.get("text", None) is not None:
                    self.variable_enlace.set(dicc_config.get("text", None))


            #se actualiza kwargs_config_atributos_propios
            self.kwargs_config_atributos_propios = {atributo: valor_atributo for atributo, valor_atributo in self.kwargs_config_parametros_lower_no_blank.items()
                                                    if atributo not in list(self.kwargs_config_atributos_nativos.keys())}



            #se actualiza kwargs_config_parametros          
            self.kwargs_config_parametros = {}
            self.kwargs_config_parametros.update(self.kwargs_config_atributos_nativos)
            self.kwargs_config_parametros.update(self.kwargs_config_atributos_propios)


            ###################################################################################################################################
            # ATRIBUTOS PROPIOS
            ###################################################################################################################################

            #se configuran los atributos propios que esten definidos en la clase y que esten en las keys de kwargs_config_atributos_propios
            #(se crea la lista lista_config_atributos_propios)

            lista_config_atributos_propios = [key for key in self.kwargs_config_atributos_propios.keys()]

            # dicc_config_root
            if self.tipo_widget_lower_no_blank in ["root", "tk", "toplevel"]:

                try:
                    dicc_config_root = (self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_config_root]
                                        if self.tipo_widget_lower_no_blank in ["root", "tk"] else self.kwargs_config_parametros_lower_no_blank)

                    if isinstance(dicc_config_root, dict):

                        title = dicc_config_root.get("title", None)
                        bg = dicc_config_root.get("bg", None)
                        tupla_geometry = dicc_config_root.get(self.nombre_kwargs_dicc_config_root_tupla_geometry, None)
                        iconbitmap = dicc_config_root.get("iconbitmap", None)
                        resizable = dicc_config_root.get("resizable", None)
                        transient = dicc_config_root.get("transient", None)
                        grab_set = dicc_config_root.get("grab_set", None)

                        if isinstance(title, (str, int, float)):
                            self.widget_objeto.title(str(title))

                        if isinstance(bg, str):
                            self.widget_objeto.configure(bg)

                        if (isinstance(tupla_geometry, tuple)
                            and len(tupla_geometry) == 2
                            and sum(1 if isinstance(item, (int, float)) else 0 for item in tupla_geometry) == len(tupla_geometry)):

                            self.widget_objeto.geometry(f"{tupla_geometry[0]}x{tupla_geometry[1]}")


                        if iconbitmap is not None:
                            self.widget_objeto.iconbitmap(iconbitmap)


                        if (isinstance(resizable, tuple)
                            and len(resizable) == 2
                            and sum(1 if isinstance(item, (int, float)) else 0 for item in resizable) == len(resizable)):

                            self.widget_objeto.resizable(resizable[0], resizable[1])


                        if self.tipo_widget_lower_no_blank == "toplevel":
                            if transient:
                                self.widget_objeto.transient(self.master)

                            if grab_set:
                                self.widget_objeto.grab_set()

                except:
                    pass



            # colocacion_dicc
            if self.nombre_kwargs_colocacion_dicc in lista_config_atributos_propios:

                try:
                    if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_colocacion_dicc], dict):
                        
                        metodo = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_colocacion_dicc].get(self.nombre_kwargs_colocacion_dicc_metodo, None)
                        coord_x = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_colocacion_dicc].get(self.nombre_kwargs_colocacion_dicc_coord_x, None)
                        coord_y = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_colocacion_dicc].get(self.nombre_kwargs_colocacion_dicc_coord_y, None)


                        if (metodo is not None and coord_x is not None and coord_y is not None and
                            isinstance(metodo, str) and isinstance(coord_x, int) and isinstance(coord_y, int)):
                                
                                if metodo.lower().replace(" ", "") == "place":
                                    self.widget_objeto.place(x = coord_x, y = coord_y)

                                elif metodo.lower().replace(" ", "") == "pack":
                                    self.widget_objeto.pack(padx = coord_x, pady = coord_y)
                            
                            
                except:
                    pass




            # combobox_lista_opciones
            if self.nombre_kwargs_combobox_lista_opciones in lista_config_atributos_propios:

                try:
                    if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_combobox_lista_opciones], (list, tuple, set)):
                        self.widget_objeto["values"] = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_combobox_lista_opciones]

                        self.widget_objeto['state'] = "readonly"
                        self.widget_objeto.configure(exportselection = False)
                        self.widget_objeto.bind("<MouseWheel>", lambda event: "break")

                except:
                    pass


            ##############################################################
            # dicc_rutina
            ##############################################################
            #IMPORTANTE: para que la llamada de config_atributos desde una clase propia en otro modulo .py (que contiene la GUI personalizada del proyecto que sea)
            #funcione con la rutina pasada por string es necesario agregar a los kwargs de la presente clase (gui_tkinter_widgets) el entorno de la clase de la GUI
            # --> self_clase_gui_donde_call_rutina = self

            #el kwargs tiene que tener un diccionario 'dicc_rutina' (atributo self.nombre_kwargs_dicc_rutina) con las keys siguientes:
            # --> 'rutina' (atributo self.nombre_kwargs_dicc_rutina_nombre_rutina)
            # --> 'tipo_rutina' (atributo self.nombre_kwargs_dicc_rutina_tipo_rutina) --> boton o evento
            # --> 'tipo_bind' (atributo self.nombre_kwargs_dicc_rutina_tipo_bind) --> <<ComboboxSelected>>, <<TreeviewSelect>>, <ButtonRelease-1> etc etc (aqui aplica solo si tipo_rutina = evento)
            # --> 'parametros_args' (atributo self.nombre_kwargs_dicc_rutina_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_dicc_rutina_parametros_kwargs)
            if self.nombre_kwargs_dicc_rutina in lista_config_atributos_propios:

                try:
                    dicc_rutina = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_dicc_rutina, None)

                    if isinstance(dicc_rutina, dict):

                        nombre_rutina = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_nombre_rutina, None)
                        tipo_rutina = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_tipo_rutina, None)
                        tipo_bind = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_tipo_bind, None)
                        rutina_parametros_args = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_parametros_args, ())
                        rutina_parametros_kwargs = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_parametros_kwargs, {})

                        if isinstance(nombre_rutina, str) and self.self_clase_gui_donde_call_rutina is not None:
                            rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina, None)

                        else:
                            rutina_objeto = nombre_rutina


                        if callable(rutina_objeto):
                            
                            #para botones (se usa command = lambda para que se aplique la rutina al widget en el momento de interactuar con el y no en el momento de su creacion)
                            if tipo_rutina == self.nombre_tipo_rutina_asociada_widget_command:

                                self.widget_objeto.config(command = lambda: rutina_objeto(
                                                                                            #args dinamicos
                                                                                            *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg for arg in rutina_parametros_args)

                                                                                            #kwargs dinamicos
                                                                                            , **{key: (valor(self.self_clase_gui_donde_call_rutina) if callable(valor) else valor)
                                                                                                    for key, valor in rutina_parametros_kwargs.items()}
                                                                                        )
                                                        )
                                
                            #para rutinas de evento
                            elif tipo_rutina == self.nombre_tipo_rutina_asociada_widget_event:

                                self.widget_objeto.bind(tipo_bind, lambda event: rutina_objeto(
                                                                                                #args dinamicos
                                                                                                *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg for arg in rutina_parametros_args)

                                                                                                #kwargs dinamicos
                                                                                                , **{key: (valor(self.self_clase_gui_donde_call_rutina) if callable(valor) else valor)
                                                                                                        for key, valor in rutina_parametros_kwargs.items()}
                                                                                                )
                                                        )

                except:
                    pass




            ##############################################################
            # lista_dicc_rutina_trace_variable_enlace
            ##############################################################
            #IMPORTANTE: para que la llamada de config_atributos desde una clase propia en otro modulo .py (que contiene la GUI personalizada del proyecto que sea)
            #funcione con la rutina pasada por string es necesario agregar a los kwargs de la presente clase (gui_tkinter_widgets) el entorno de la clase de la GUI
            # --> self_clase_gui_donde_call_rutina = self

            #el kwargs tiene que tener un diccionario 'dicc_rutina' (atributo self.nombre_kwargs_dicc_rutina) con las keys siguientes:
            # --> 'rutina' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina)
            # --> 'tipo_trace' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace) --> boton o evento
            # --> 'parametros_args' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs)
            if self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace in lista_config_atributos_propios:

                try:
                    lista_dicc_rutina = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace, None)

                    if isinstance(lista_dicc_rutina, list) and sum(1 if isinstance(dicc, dict) else 0 for dicc in lista_dicc_rutina) == len(lista_dicc_rutina):

                        for dicc_rutina in lista_dicc_rutina:

                            nombre_rutina = dicc_rutina.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina, None)
                            tipo_trace = dicc_rutina.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace, None)
                            rutina_parametros_args = dicc_rutina.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args, ())
                            rutina_parametros_kwargs = dicc_rutina.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs, {})

                            if isinstance(nombre_rutina, str) and self.self_clase_gui_donde_call_rutina is not None:
                                rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina, None)

                            else:
                                rutina_objeto = nombre_rutina

                            if callable(rutina_objeto):
                            
                                #las rutinas asociadas a eventos trace de stringvar deben tener como parametro *args de ahi "lambda *trace_args"
                                #en la rutina de la GUI del proyecto donde se use esta clase ya no es necesario ponerle *args
                                # IMPORTANTE: poner la "captura_lambda" en "captura_lambda = rutina_objeto" y "rutina_parametros_kwargs = rutina_parametros_kwargs: captura_lambda("
                                #             pq en caso de que lista_dicc_rutina tenga mas de 1 diccionario el trace ejecuta tan solo el ultimo de la lista
                                #             de ahi la necesidad de almacenar rutina_objeto de la lambda en captura_lambda
                                self.variable_enlace.trace_add(tipo_trace
                                                                , lambda *trace_args
                                                                , captura_lambda = rutina_objeto
                                                                , rutina_parametros_args = rutina_parametros_args
                                                                , rutina_parametros_kwargs = rutina_parametros_kwargs: captura_lambda(
                                                                                                                                        #args dinamicos
                                                                                                                                        *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg
                                                                                                                                            for arg in rutina_parametros_args)

                                                                                                                                        #kwargs dinamicos
                                                                                                                                        , **{key: (valor(self.self_clase_gui_donde_call_rutina)
                                                                                                                                                   if callable(valor) else valor)
                                                                                                                                                    for key, valor in rutina_parametros_kwargs.items()}
                                                                                                                                    )
                                                                )

                except:
                    pass



            # alineacion
            if self.nombre_kwargs_alineacion in lista_config_atributos_propios:

                try:
                    opcion_alineacion = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_alineacion, None)

                    if opcion_alineacion.lower().replace(" ", "") == "center":
                        self.widget_objeto.config(anchor = "center")
                        self.widget_objeto.config(justify = self.objeto_alias_tkinter_import.CENTER)

                    elif opcion_alineacion.lower().replace(" ", "") == "left":
                        self.widget_objeto.config(anchor = "w")
                        self.widget_objeto.config(justify = self.objeto_alias_tkinter_import.LEFT)

                    elif opcion_alineacion.lower().replace(" ", "") == "right":
                        self.widget_objeto.config(anchor = "e")
                        self.widget_objeto.config(justify = self.objeto_alias_tkinter_import.RIGHT)

                    elif opcion_alineacion.lower().replace(" ", "") == "top_center":
                        self.widget_objeto.config(anchor = "n")

                    elif opcion_alineacion.lower().replace(" ", "") == "top_left":
                        self.widget_objeto.config(anchor = "nw")

                    elif opcion_alineacion.lower().replace(" ", "") == "top_right":
                        self.widget_objeto.config(anchor = "ne")

                    elif opcion_alineacion.lower().replace(" ", "") == "bottom_center":
                        self.widget_objeto.config(anchor = "s")

                    elif opcion_alineacion.lower().replace(" ", "") == "bottom_left":
                        self.widget_objeto.config(anchor = "sw")

                    elif opcion_alineacion.lower().replace(" ", "") == "bottom_right":
                        self.widget_objeto.config(anchor = "se")

                except:
                    pass



            # bloquear
            if self.nombre_kwargs_bloquear in lista_config_atributos_propios:

                try:
                    opcion_bloqueo = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_bloquear, None)

                    if opcion_bloqueo.lower().replace(" ", "") == "si":
                        self.widget_objeto.config(state = self.objeto_alias_tkinter_import.DISABLED)

                    elif opcion_bloqueo.lower().replace(" ", "") == "no":
                        self.widget_objeto.config(state = self.objeto_alias_tkinter_import.NORMAL)

                except:
                    pass


            # listbox_varios
            if self.nombre_kwargs_listbox_lista_items in lista_config_atributos_propios:

                try:
                    self.widget_objeto.delete(0, self.objeto_alias_tkinter_import.END)
                    for item in self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_listbox_lista_items]:
                        self.widget_objeto.insert(self.objeto_alias_tkinter_import.END, item)

                except:
                    pass


            if self.nombre_kwargs_listbox_lista_items_seleccionados in lista_config_atributos_propios:

                try:
                    item_selecc_indices = self.widget_objeto.curselection()
                    return [self.widget_objeto.get(i) for i in item_selecc_indices] if item_selecc_indices else []
                
                except:
                    pass


            if self.nombre_kwargs_listbox_listbox_seleccionar_todo_o_nada in lista_config_atributos_propios:

                try:
                    lista_items_selecc = [self.widget_objeto.get(i) for i in self.widget_objeto.curselection()]
            
                    #seleccionar todo
                    if len(lista_items_selecc) == 0:
                        self.widget_objeto.selection_set(0, self.objeto_alias_tkinter_import.END)

                    #des-seleccionar todo
                    elif len(lista_items_selecc) != 0:
                        self.widget_objeto.selection_clear(0, self.objeto_alias_tkinter_import.END)

                except:
                    pass


            # controltiptext
            if self.nombre_controltiptext in lista_config_atributos_propios:
                if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_controltiptext], str):
                    controltiptext(self.widget_objeto, (self.kwargs_config_parametros_lower_no_blank[self.nombre_controltiptext]))



    def destroy(self):

        """
        Metodo que permite eliminar el widget.
        """

        self.widget_objeto.destroy()



    ###########################################################################################
    #     METODOS PROPIOS - INTERNOS A LA CLASE (vienen precedidos de 2 guiones bajos __)
    ###########################################################################################

    def __varios_clase_madre(self, opcion_varios: Literal ["crear_widget_objeto", "nombre_modulo_python_py", "nombre_alias_libreria_python_import", "objeto_alias_libreria_python_import", 
                                                            "lista_atributos_constructor_clase", "dicc_metodos_propios_clase"]
                            , **kwargs):

        """
        Rutina interna (hibrido entre rutina y función) que permite realizar acciones varias dentro de la clase gui_tkinter_widgets.

        opcion_varios:
        --

        --> crear_widget_objeto\n
        \t\tCrea el widget.

        --> nombre_modulo_python_py\n
        \t\tDevuelve el nombre del presente módulo .py.

        --> nombre_alias_libreria_python_import\n
        \t\tDevuelve el string del alias de la libreria tkinter importada en el módulo .py (import tkinter as tk --> devuelve "tk").

        --> objeto_alias_libreria_python_import\n
        \t\tDevuelve el objeto del alisas de la libreria tkinter importada en el módulo .py (import tkinter as tk --> devuelve tk).
        
        --> lista_atributos_constructor_clase\n
        \t\tDevuelve todos los atributos incializados en el constructor de la clase (__init__) los que vienen precedidos por self.

        --> dicc_metodos_propios_clase\n
        \t\tDevuelve un diccionario que lista tantos los metodos propios públicos y privados de la clase.
 
        kwargs:
        --
        --> libreria_python (se usa solo en opcion_varios = 'nombre_alias_libreria_python_import' y 'objeto_alias_libreria_python_import').\n
        """

        resultado_funcion = None

        #parametros kwargs
        libreria_python = kwargs.get("libreria_python", None)


        if opcion_varios == "crear_widget_objeto":

            #se listan los modulos tkinter (mediante la libreria pkgutil), se agrega el nombre_alias_tkinter_import
            lista_modulos_tkinter = [self.nombre_alias_tkinter_import] + [modulo for _, modulo, _ in pkgutil.iter_modules(self.objeto_alias_tkinter_import.__path__, "")]


            #se listan todas las combinaciones posibles permutando por mayuscula y minuscula cada caracter de self.tipo_widget_lower_no_blank
            #(se usa para ello la funcion product de la libreria itertools)
            lista_tuplas_combinaciones_lower_upper = list(product(*[(caracter.lower(), caracter.upper()) for caracter in self.tipo_widget_lower_no_blank]))
            lista_combinaciones_lower_upper = ["".join(tupla) for tupla in lista_tuplas_combinaciones_lower_upper]


            # se crea el objeto widget mediante bucle por modulo y por combinacion mayusc/minusc se localzan en el modulo de tkinter y el objeto asociado
            #(el metodo globals() para localizar el modulo de la libreria tkinter solo funciona si se ha realizado en el modulo .py el 'from tkinter import ttk' etc etc)
            #se usa la variable check_localiz_modulo_y_objeto para informar los errores que pueden haber, no se puede hacer durante 
            #los bloques except pq la busqueda se realiza iterando tanto sobre lista_modulos_tkinter como sobre lista_combinaciones_lower_upper
            #hasta dar con la combinacion que funciona por lo tanto es normal que surjan errores y estos no se han de logear 
            #(tan solo se logea al final si no se ha encontraddo ningun macheo)
            check_localiz_modulo_y_objeto = ""

            str_modulo_tk = None
            combinacion_tipo_widget = None
            modulo_tk = None

            for str_modulo_tk in lista_modulos_tkinter:
                for combinacion_tipo_widget in lista_combinaciones_lower_upper:

                    try:
                        modulo_tk = globals()[str_modulo_tk]
                    except:
                        pass
                    else:

                        try:
                            widget_objeto_por_crear = getattr(modulo_tk, combinacion_tipo_widget)

                        except AttributeError:
                            pass

                        else:
                            check_localiz_modulo_y_objeto = "ok"
                            break

                if check_localiz_modulo_y_objeto == "ok":
                    break


            #se crea el widget_objeto en caso de que el macheo haya dado resultado
            #y se le asigna la variable_enlace creada en el constructor
            if check_localiz_modulo_y_objeto == "ok":
                self.widget_objeto = widget_objeto_por_crear(self.master)

                for atributo_variable_enlace in self.lista_atributos_vinculados_variable_enlace:
                    dicc_atributo_variable_enlace = {atributo_variable_enlace: self.variable_enlace}

                    try:
                        self.widget_objeto.config(**dicc_atributo_variable_enlace)
                    except:
                        try:
                            self.widget_objeto.configure(**dicc_atributo_variable_enlace)
                        except:
                            pass

                        pass

            resultado_funcion = None



        elif opcion_varios == "nombre_modulo_python_py":
            #devuelve el nombre del modulo python donde se ubica la clase

            frame_actual = inspect.currentframe()
            fichero_modulo_py = inspect.getfile(frame_actual)
            resultado_funcion = os.path.basename(fichero_modulo_py)



        elif "alias_libreria_python_import" in opcion_varios:
            #devuelve el el nombre o el objeto del alias de la librerias python importada en este modulo .py
            #el metodo globals() solo funciona si la libreria se importo en este modulo .py

            dicc_globals = dict(globals())
            for nombre, objeto in dicc_globals.items():
                if objeto is sys.modules.get(libreria_python):
                    resultado_funcion = nombre if "nombre" in opcion_varios else objeto if "objeto" in opcion_varios else None
                    break



        elif opcion_varios == "lista_atributos_constructor_clase":
            #devuelve una lista con todos los atributos inicializados en el constructor de la clase (__init__)
            #los que se declaran precedidos de self

            resultado_funcion = [key for key, _ in self.__dict__.items()]



        elif opcion_varios == "dicc_metodos_propios_clase":
            #devuelve un diccionario con 2 keys:
            # --> lista_metodos_propios_publicos       lista de metodos propios publicos (usables fuera de la clase)       
            # --> lista_atributos_propios_privados     lista de metodos propios privados (usables solo internamente en la clase)     

            #lista_metodos_propios_publicos y lista_atributos_propios_privados se obtienen directamente usando la libreria inspect
            str_inicio_rename_interno_clase_metodos_privados = f"_{self.clase_nombre}"

            lista_metodos_propios_publicos = [metodo_propio for metodo_propio, _ in inspect.getmembers(self.clase_objeto, inspect.isfunction)
                                                if metodo_propio[:len(str_inicio_rename_interno_clase_metodos_privados)] != str_inicio_rename_interno_clase_metodos_privados
                                                and metodo_propio[:2] != "__" and metodo_propio[-2:] != "__"]
            
            lista_metodos_propios_privados = [metodo_propio[len(str_inicio_rename_interno_clase_metodos_privados):] for metodo_propio, _ in inspect.getmembers(self.clase_objeto, inspect.isfunction)
                                                if metodo_propio[:len(str_inicio_rename_interno_clase_metodos_privados)] == str_inicio_rename_interno_clase_metodos_privados]


            resultado_funcion = {"lista_metodos_propios_publicos": lista_metodos_propios_publicos
                                 , "lista_metodos_propios_privados": lista_metodos_propios_privados
                                 }


        #resultado de la funcion
        return resultado_funcion



###########################################################################################
#     CLASE HIJA - treeview_propio
###########################################################################################
class treeview_propio(gui_tkinter_widgets):
    #clase hija de la clase gui_tkinter_widgets para crear widget de tipo treeview con metodos propios asociados

    def __init__(self, master, self_clase_gui_donde_call_rutina = None, **kwargs_config_widget):

        #se habilita la herencia atributos y metodos de la clase madre
        super().__init__(master, self_clase_gui_donde_call_rutina = self_clase_gui_donde_call_rutina, **kwargs_config_widget)


        #se asigna el nombre del diccionario kwargs donde recuperar los parametros de configuracion
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs = "treeview_dicc"
        self.nombre_kwargs_height = "height"
        self.nombre_kwargs_columnas_df = "columnas_df"
        self.nombre_kwargs_columnas_treeview = "columnas_treeview"
        self.nombre_kwargs_width_columnas_treeview = "width_columnas_treeview"

        self.nombre_kwargs_rutina_click_item = "dicc_rutina_click_item"
        self.nombre_kwargs_rutina_click_item_nombre_rutina = "rutina"
        self.nombre_kwargs_rutina_click_item_parametros_args = "parametros_args"
        self.nombre_kwargs_rutina_click_item_parametros_kwargs = "parametros_kwargs"
        

        #se inicializan atributos necesarios a la clase hija
        self.nombre_tipo_widget = "Treeview"
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget
        self.self_clase_gui_donde_call_rutina = self_clase_gui_donde_call_rutina


        self.widget_objeto = None
        self.datos_item_seleccionado = None


        #se pocede a crear el widget
        self.__crear_widget_objeto()


        #se ejecutan los atributos nativos y propios
        self.config_atributos(**kwargs_config_widget)



    def actualizar_desde_df(self, df_datos: pd.DataFrame):

        """
        Metodo que rellena el treeview con el dataframe pasado por parametro.

        args
        --
        --> df_datos\n
        Es el dataframe que se ha de usar para completar el treeview.
        """

        if self.widget_objeto is not None:

            if isinstance(df_datos, pd.DataFrame):

                lista_config_columnas_df = self.datos_item_seleccionado["lista_columnas_df"]
                lista_columnas_df_datos = [columna for columna in df_datos.columns]

                lista_columnas_df_datos_ok = [columna for columna in df_datos.columns if columna in lista_config_columnas_df]
                lista_columnas_faltantes_df_datos = [columna_config for columna_config in lista_config_columnas_df if columna_config not in lista_columnas_df_datos]
                lista_columnas_faltantes_df_datos_valores = ["" for columna_config in lista_config_columnas_df if columna_config not in lista_columnas_df_datos]


                #se rellena el treeview solo si df_datos no es vacio y hay al menos 1 columna en df_datos incluida en la lista de columnas configuradas para el df
                #en datos_item_seleccionado (lista_columnas_df)
                if len(df_datos) != 0 and len(lista_columnas_df_datos_ok) != 0:

                    #se agregan las columnas configuradas en el atributo dicc_treeview_columnas_y_width (columnas_df) en el df_datos si no aprecen
                    #(se les pone el valor "")
                    if len(lista_columnas_faltantes_df_datos) != 0:
                        df_datos[lista_columnas_faltantes_df_datos] = lista_columnas_faltantes_df_datos_valores

                    df_datos = df_datos[lista_config_columnas_df]


                    #se actualiza el tipo de dato de las columnas del df (se hace antes de rellenar el treeview
                    #pq segun el tamaño de df puede tardar unos segundos)
                    lista_columnas_df_tipo_datos = [df_datos[columna].dtype for columna in df_datos.columns]
                    self.datos_item_seleccionado["lista_tipo_dato_columna_df"] = lista_columnas_df_tipo_datos


                    #se rellena el treeview
                    for item in self.widget_objeto.get_children():
                        self.widget_objeto.delete(item)

                    for _, linea in df_datos.iterrows():
                        self.widget_objeto.insert("", "end", values = tuple([linea[columna] for columna in lista_config_columnas_df]))


                elif len(df_datos) == 0:
                    #se vacia el treeview
                    for item in self.widget_objeto.get_children():
                        self.widget_objeto.delete(item)


    def __crear_widget_objeto(self):
        #Rutina interna que permite crear el widget

        try:
            #se crea el objeto treeview
            #selectmode = browse significa seleccion de un solo elemento (extended es para seleccion multiple)
            self.widget_objeto = ttk.Treeview(self.master, selectmode = "browse")


            #se realizan ajustes especificos al objeto treeview
            height = self.kwargs_config_widget[self.nombre_kwargs][self.nombre_kwargs_height]
            columnas_df = self.kwargs_config_widget[self.nombre_kwargs][self.nombre_kwargs_columnas_df]
            columnas_treeview = self.kwargs_config_widget[self.nombre_kwargs][self.nombre_kwargs_columnas_treeview]
            width_columnas_treeview = self.kwargs_config_widget[self.nombre_kwargs][self.nombre_kwargs_width_columnas_treeview]

            tupla_columnas_treeview = tuple([columna for columna in columnas_treeview])
            self.widget_objeto.config(columns = tupla_columnas_treeview, show = "headings")

            self.widget_objeto["height"] = height

            width_acum = 0
            for ind, columna in enumerate(columnas_treeview):
                width = width_columnas_treeview[ind]
                self.widget_objeto.heading(columna, text = columna)

                width_corr = width if ind + 1 == len(columnas_treeview) else width - 1 #el - 1 es para que no machaque el borde vertical-derecha del treeview
                self.widget_objeto.column(columna, width = width_corr)

                width_acum += width_corr
                
            # width_widget_objeto = sum(width for width in columnas_treeview)
            self.widget_objeto.place(width = width_acum)


            #se enlaza la accion al cliquar sobre un item
            #(#add = "+" es para no machacar el bind _asociado a la rutina _rutina_click_item si esta definida en los kwargs)
            self.widget_objeto.bind("<ButtonRelease-1>", lambda event: self.__click_on_item(event), add = "+")


            #se crea la variable interna de la clase datos_item_seleccionado con los datos de configuracion del treeview
            #es un diccionario que contiene las keys siguientes:
            # --> lista_columnas_df                 es la lista de las columnas del df configuradas en el atributo propio dicc_treeview_columnas_y_width
            # --> lista_columnas_treeview           es la lista de las columnas del treeview configuradas en el atributo propio dicc_treeview_columnas_y_width
            # --> lista_width_columnas_treeview     es la lista de los width de las columnas del treeview configuradas en el atributo propio dicc_treeview_columnas_y_width
            # --> lista_tipo_dato_columna_df        es la lista de los tipos de datos de las columnas del df del parametro df_datos
            # --> lista_datos_item_seleccionado     es la lista de los datos del item seleccionado
            self.datos_item_seleccionado = {"lista_columnas_df": list(columnas_df)
                                            , "lista_columnas_treeview": list(columnas_treeview)
                                            , "lista_width_columnas_treeview": list(width_columnas_treeview)
                                            , "lista_tipo_dato_columna_df": None     #se asigna cuando se rellena el treeview con el df
                                            , "lista_datos_item_seleccionado": None  #se asigna cuando se realiza la accion de click sobre un item del treeview
                                            }
        except:
            pass

        finally:
            ##############################################################
            # dicc_rutina_click_item
            #IMPORTANTE: para que la llamada de config_atributos desde una clase propia en otro modulo .py (que contiene la GUI personalizada del proyecto que sea)
            #funcione con la rutina pasada por string es necesario agregar a los kwargs de la presente clase (treeview_propio) el entorno de la clase de la GUI
            # --> self_clase_gui_donde_call_rutina = self

            #el kwargs tiene que tener un diccionario 'dicc_rutina_click_item' (atributo self.nombre_kwargs_rutina_click_item) con las keys siguientes:
            # --> 'rutina' (atributo self.nombre_kwargs_rutina_click_item_nombre_rutina)
            # --> 'parametros_args' (atributo self.nombre_kwargs_rutina_click_item_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_rutina_click_item_parametros_kwargs)
            dicc_rutina_click_item = self.kwargs_config_widget.get(self.nombre_kwargs_rutina_click_item, None)

            try:
                if isinstance(dicc_rutina_click_item, dict):

                    nombre_rutina_click_item = dicc_rutina_click_item.get(self.nombre_kwargs_rutina_click_item_nombre_rutina, None)
                    rutina_click_item_parametros_args = dicc_rutina_click_item.get(self.nombre_kwargs_rutina_click_item_parametros_args, ())
                    rutina_click_item_parametros_kwargs = dicc_rutina_click_item.get(self.nombre_kwargs_rutina_click_item_parametros_kwargs, {})

                    if isinstance(nombre_rutina_click_item, str) and self.self_clase_gui_donde_call_rutina is not None:
                        rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina_click_item, None)

                    else:
                        rutina_objeto = nombre_rutina_click_item

                    if callable(rutina_objeto):

                        self.widget_objeto.bind("<ButtonRelease-1>", lambda event: rutina_objeto(
                                                                                                event

                                                                                                #args dinamicos
                                                                                                , *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg for arg in rutina_click_item_parametros_args)
                                                                                                
                                                                                                #kwargs dinamicos
                                                                                                , **{key: (valor(self.self_clase_gui_donde_call_rutina) if callable(valor) else valor)
                                                                                                        for key, valor in rutina_click_item_parametros_kwargs.items()}
                                                                                                )
                                                #add = "+" es para no machacar el bind __click_on_item
                                                , add = "+"
                                                )
            except:
                pass



    def __click_on_item(self, event = None):
        #Rutina interna que permite recuperar los datos del item seleccionado.

        item_selecc = self.widget_objeto.selection()

        if item_selecc:
            item_id = item_selecc[0]
            self.datos_item_seleccionado["lista_datos_item_seleccionado"] = list(self.widget_objeto.item(item_id, "values"))



###########################################################################################
#     CLASE HIJA - scrolledtext_propio
###########################################################################################
class scrolledtext_propio(gui_tkinter_widgets):
    #clase hija de la clase gui_tkinter_widgets para crear widget de tipo scrolledtext con o sin tags
    # con metodos propios asociados

    def __init__(self, master, **kwargs_config_widget):
        

        #se habilita la herencia atributos y metodos de la clase madre
        super().__init__(master, **kwargs_config_widget)

  
        #se asigna el nombre del diccionario kwargs donde recuperar los parametros de configuracion
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs_colocacion_scrollbar_horizontal = "colocacion_scrollbar_horizontal"
        self.nombre_kwargs_colocacion_scrollbar_horizontal_metodo = "metodo"
        self.nombre_kwargs_colocacion_scrollbar_horizontal_coord_x = "coord_x"
        self.nombre_kwargs_colocacion_scrollbar_horizontal_coord_y = "coord_y"

        self.nombre_kwargs_df_datos = "df_datos"
        self.nombre_kwargs_columna_df_para_informar = "columna_df_para_informar"

        self.nombre_kwargs_lista_dicc_tag_linea_completa = "lista_dicc_tag_linea_completa"
        self.nombre_kwargs_lista_dicc_tag_linea_completa_nombre_tag = "nombre_tag"
        self.nombre_kwargs_lista_dicc_tag_linea_completa_columna_df_tag_aplicar = "columna_df_tag_aplicar"
        self.nombre_kwargs_lista_dicc_tag_linea_completa_case_sensitive = "case_sensitive"
        self.nombre_kwargs_lista_dicc_tag_linea_completa_dicc_config = "dicc_config"

        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa = "lista_dicc_tag_caracteres_cambiantes_comparativa"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag = "nombre_tag"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag = "columna_df_filtro_registros_aplicar_tag"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor = "columna_df_filtro_registros_aplicar_tag_valor"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1 = "columna_df_comparar_1"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2 = "columna_df_comparar_2"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_case_sensitive = "case_sensitive"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_marcar_toda_linea_si_todo_varia = "marcar_toda_linea_si_todo_varia"
        self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_dicc_config = "dicc_config"


        #se inicializan atributos necesarios a la clase hija
        self.nombre_tipo_widget = "ScrolledText"
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget

        self.widget_objeto = None

        #se pocede a crear el widget
        self.__varios_clase("crear_widget_objeto")
        self.srollbar_vertical = self.widget_objeto.vbar


        #se ejecutan los atributos nativos y propios
        self.config_atributos(**kwargs_config_widget)


    def modificaciones(self, opcion_modif, **kwargs):

        """
        Metodo que permite realizar modificaciones sobre el contenido y/o tags en el scrolledtext

        args
        --
        --> borrar_contenido_y_tags\n
        Borra todo el contenido y todos los tags.

        --> agregar_solo_contenido_desde_string\n
        Agrega el contenido desde un string. Requiere el uso de los kwargs siguientes: string_texto_informar y height_scrolledtext.

        --> agregar_solo_contenido_desde_dataframe\n
        Agrega el contenido desde un dataframe. Requiere el uso de los kwargs siguientes: df_datos, columna_df_para_informar y height_scrolledtext.

        --> agregar_contenido_y_tags_desde_dataframe\n
        Agrega el contenido y los tags desde un dataframe. Requiere el uso de los kwargs siguientes: df_datos, columna_df_para_informar, height_scrolledtext.
        Asimismo, requiere otros kwargs: lista_dicc_tag_linea_completa o lista_dicc_tag_caracteres_cambiantes_comparativa (al menos uno de los 2).
        """

        #parametros kwargs
        string_texto_informar = kwargs.get("string_texto_informar", None)
        df_datos = kwargs.get(self.nombre_kwargs_df_datos, None)
        columna_df_para_informar = kwargs.get(self.nombre_kwargs_columna_df_para_informar, None)
        height_scrolledtext = kwargs.get("height_scrolledtext", None)
        lista_dicc_tag_linea_completa = kwargs.get(self.nombre_kwargs_lista_dicc_tag_linea_completa, None)
        lista_dicc_tag_caracteres_cambiantes_comparativa = kwargs.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa, None)


        if opcion_modif == "borrar_contenido_y_tags":

            #se borra el contenido
            self.widget_objeto.config(state = self.objeto_alias_tkinter_import.NORMAL)
            self.widget_objeto.delete(1.0, self.objeto_alias_tkinter_import.END)

            #se borran todos los tags existentes
            lista_tags = list(self.widget_objeto.tag_names())
            for tag in lista_tags:
                self.widget_objeto.tag_delete(tag)


        elif opcion_modif == "agregar_solo_contenido_desde_string":

            #se inserta el texto en el scrolledtext y se calcula el numero de lineas informado
            self.widget_objeto.insert(tk.END, string_texto_informar)
            numero_lineas_informadas = len(self.widget_objeto.get("1.0", tk.END).splitlines())

            self.widget_objeto.config(state = self.objeto_alias_tkinter_import.DISABLED)

            #se agrega scrollbar vertical si el texto excede el height
            if numero_lineas_informadas > height_scrolledtext:
                self.srollbar_vertical.pack(side = "right", fill = "y")
            else:
                self.widget_objeto.vbar.pack_forget()


        elif opcion_modif == "agregar_solo_contenido_desde_dataframe":

            if columna_df_para_informar is not None:
                existe_columna_en_df = "si" if columna_df_para_informar in [columna for columna in df_datos.columns] else "no"

                if existe_columna_en_df == "si":
                    df_datos.reset_index(drop = True, inplace = True)

                    for ind in df_datos.index:
                        self.widget_objeto.insert(self.objeto_alias_tkinter_import.END, df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_para_informar)] + "\n")

            #se agrega scrollbar vertical si el texto excede el height
            if len(df_datos) > height_scrolledtext:
                self.srollbar_vertical.pack(side = "right", fill = "y")
            else:
                self.widget_objeto.vbar.pack_forget()



        elif opcion_modif == "agregar_contenido_y_tags_desde_dataframe":

            #se crean los tags configurados si lista_dicc_tag_linea_completa es lista
            if isinstance(lista_dicc_tag_linea_completa, list):

                for ind, dicc_tag in enumerate(lista_dicc_tag_linea_completa):
                    nombre_tag = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_linea_completa_nombre_tag, None)
                    dicc_config = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_linea_completa_dicc_config, None)

                    if isinstance(nombre_tag, str):
                        
                        if isinstance(dicc_config, dict):
                            for atributo, valor_atributo in dicc_config.items():
                                try:
                                    self.widget_objeto.tag_configure(nombre_tag, **{atributo: valor_atributo})

                                except:
                                    pass

            #se crean los tags configurados si lista_dicc_tag_caracteres_cambiantes_comparativa es lista
            if isinstance(lista_dicc_tag_caracteres_cambiantes_comparativa, list):

                for ind, dicc_tag in enumerate(lista_dicc_tag_caracteres_cambiantes_comparativa):
                    nombre_tag = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag, None)
                    dicc_config = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_dicc_config, None)

                    if isinstance(nombre_tag, str):
                        
                        if isinstance(dicc_config, dict):
                            for atributo, valor_atributo in dicc_config.items():
                                try:
                                    self.widget_objeto.tag_configure(nombre_tag, **{atributo: valor_atributo})

                                except:
                                    pass




            #se informa el scrolledtext con los tags (si lista_dicc_tag no esta configurada o no esta bien configurada solo se agregan los datos sin tags) 
            if isinstance(df_datos, pd.DataFrame) and columna_df_para_informar is not None:

                lista_columnas_df = [columna for columna in df_datos.columns]
                existe_columna_en_df = "si" if columna_df_para_informar in lista_columnas_df else "no"

                if existe_columna_en_df == "si":

                    #se extraen de df_datos la columna que sirve para informar el contenido + las que sirven para los tags 
                    lista_columnas_df_tags_linea_completa = [dicc_tag[self.nombre_kwargs_lista_dicc_tag_linea_completa_columna_df_tag_aplicar] 
                                                            for dicc_tag in lista_dicc_tag_linea_completa
                                                            if dicc_tag[self.nombre_kwargs_lista_dicc_tag_linea_completa_columna_df_tag_aplicar] in lista_columnas_df]

                    lista_columnas_filtro_registros_aplicar_tag = [dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag] 
                                                                    for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                    if dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag] in lista_columnas_df]
                    
                    lista_columnas_df_filtro_registros_aplicar_tag_valor = [dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor] 
                                                                            for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                            if dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor] in lista_columnas_df]


                    lista_columnas_df_comparar_1 = [dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1] 
                                                    for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                    if dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1] in lista_columnas_df]
                    
                    lista_columnas_df_comparar_2 = [dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2] 
                                                    for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                    if dicc_tag[self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2] in lista_columnas_df]

                    lista_columnas_df = ([columna_df_para_informar] + lista_columnas_df_tags_linea_completa + lista_columnas_filtro_registros_aplicar_tag + lista_columnas_df_filtro_registros_aplicar_tag_valor 
                                        + lista_columnas_df_comparar_1 + lista_columnas_df_comparar_2)
                    lista_columnas_df = list(dict.fromkeys(lista_columnas_df))

                    df_datos = df_datos[lista_columnas_df]
                    df_datos.reset_index(drop = True, inplace = True)


                    #se borra el contenido del scrolledtext
                    self.widget_objeto.delete(1.0, self.objeto_alias_tkinter_import.END)
                
                    #se rellena el scrolledtext y se le aplican los tags interando por los registros del df df_datos
                    for ind in df_datos.index:

                        #se extrae la linea del df q    ue sirve para informar el scrolledtext
                        linea_texto_informar = str(df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_para_informar)])

                        #se inserta el registro del df df_datos en el scrolledtext
                        self.widget_objeto.insert(self.objeto_alias_tkinter_import.END, linea_texto_informar + "\n")

                        #se extraen los indices de posicion de los caracteres de la linea tanto el de inicio como el final
                        #se usa en los tags configurados LINEA_COMPLETA y CARACTERES_CAMBIANTES_COMPARATIVA
                        indice_ini_linea_texto_informar = 0
                        indice_fin_linea_texto_informar = len(linea_texto_informar)

                        #############################################################
                        #se aplican los tags configurados LINEA_COMPLETA
                        #############################################################
                        if isinstance(lista_dicc_tag_linea_completa, list):

                            #se itera por los distintos tags configurados
                            for dicc_tag in lista_dicc_tag_linea_completa:

                                #se extraen los parametros de cada tag
                                nombre_tag = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_linea_completa_nombre_tag, None)
                                columna_df_tag_aplicar = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_linea_completa_columna_df_tag_aplicar, None)
                                case_sensitive = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_linea_completa_case_sensitive, False)

                                if isinstance(columna_df_tag_aplicar, str) and columna_df_tag_aplicar in lista_columnas_df:

                                    #se extrae la linea de texto donde buscar el tag
                                    linea_texto_con_tag_aplicar = df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_tag_aplicar)]

                                    if nombre_tag is not None: 

                                        #se aplica el tag si existe aplicando case sensitive o no segun este configurado o no
                                        linea_texto_columna_busqueda_tag_ajust = str(linea_texto_con_tag_aplicar).lower() if not case_sensitive else str(linea_texto_con_tag_aplicar)
                                        nombre_tag_ajust = nombre_tag.lower() if not case_sensitive else nombre_tag

                                        #se crea lista de macheos usando el metodo finditer de la libreria re
                                        #y en caso de no ser vacia se aplica el tag sobre la linea completa
                                        lista_macheos_indices_string_buscado = list(re.finditer(nombre_tag_ajust, linea_texto_columna_busqueda_tag_ajust))

                                        if len(lista_macheos_indices_string_buscado) != 0:
                                            self.widget_objeto.tag_add(nombre_tag
                                                                    , f"{ind + 1}.{indice_ini_linea_texto_informar}"
                                                                    , f"{ind + 1}.{indice_fin_linea_texto_informar}")


                        #############################################################
                        #se aplican los tags configurados CARACTERES_CAMBIANTES_COMPARATIVA
                        #############################################################
                        if isinstance(lista_dicc_tag_caracteres_cambiantes_comparativa, list):

                            #se itera por los distintos tags configurados
                            for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa:

                                nombre_tag = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag, None)
                                columna_df_filtro_registros_aplicar_tag = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag, None)
                                columna_df_filtro_registros_aplicar_tag_valor = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor, None)
                                columna_df_comparar_1 = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1, None)
                                columna_df_comparar_2 = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2, None)
                                case_sensitive = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_case_sensitive, False)
                                marcar_toda_linea_si_todo_varia = dicc_tag.get(self.nombre_kwargs_lista_dicc_tag_caracteres_cambiantes_comparativa_marcar_toda_linea_si_todo_varia, False)


                                if (isinstance(columna_df_comparar_1, str) and isinstance(columna_df_comparar_2, str)
                                    and columna_df_comparar_1 in lista_columnas_df and columna_df_comparar_2 in lista_columnas_df):

                                    #se recuperan las lineas de texto a comparar y se las ajusta segun se haya configurado case sensitive o no
                                    linea_texto_comparar_1 = (df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_1)]
                                                                if df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_1)] is not None
                                                                else "")
                                    
                                    linea_texto_comparar_2 = (df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_2)]
                                                                if df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_2)] is not None
                                                                else "")


                                    if nombre_tag is not None:

                                        #se recuperan las lineas de texto a comparar y se las ajusta segun se haya configurado case sensitive o no
                                        linea_texto_comparar_1 = (df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_1)]
                                                                    if df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_1)] is not None
                                                                    else "")
                                        
                                        linea_texto_comparar_2 = (df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_2)]
                                                                    if df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_comparar_2)] is not None
                                                                    else "")
                                        
                                        linea_texto_comparar_1_ajust = linea_texto_comparar_1.lower() if not case_sensitive else linea_texto_comparar_1
                                        linea_texto_comparar_2_ajust = linea_texto_comparar_2.lower() if not case_sensitive else linea_texto_comparar_2


                                        #se localizan los caracteres cambiantes usando squencematcher de difflib
                                        #y se almacenan en unalista de tuplas (donde se quitan las tuplas donde indice_ini = indice_fin)
                                        macheo_sequencematcher = difflib.SequenceMatcher(None, linea_texto_comparar_1_ajust, linea_texto_comparar_2_ajust)

                                        lista_tuplas_indices_linea = []
                                        for tipo_macheo, ind_linea_ini, ind_linea_fin, *_ in macheo_sequencematcher.get_opcodes():

                                            if tipo_macheo != "equal":
                                                tupla_indices_linea = (ind_linea_ini, ind_linea_fin)
                                                lista_tuplas_indices_linea.append(tupla_indices_linea)

                                        lista_tuplas_indices_linea = [(indice_ini, indice_fin) for indice_ini, indice_fin in lista_tuplas_indices_linea if indice_ini != indice_fin]

                
                                        #se aplican los tags siempre y cuando marcar_solo_caracteres_cambiantes este configurado
                                        if len(lista_tuplas_indices_linea) != 0:

                                            #se ordena la lista lista_tuplas_indices_linea por el 1er tem de cada sublista
                                            #y se localiza si las variaciones de caracteres afectan toda la linea o no
                                            lista_tuplas_indices_linea.sort(key = lambda x: x[0], reverse = False)

                                            variaciones_afecta_linea_completa = (True
                                                                                if lista_tuplas_indices_linea[0][0] == indice_ini_linea_texto_informar 
                                                                                and lista_tuplas_indices_linea[-1][1] == indice_fin_linea_texto_informar
                                                                                else False)
                                            
                                            if columna_df_filtro_registros_aplicar_tag is not None and columna_df_filtro_registros_aplicar_tag_valor is not None:

                                                #se extrae la linea de texto para filtrar los registros donde aplicar o no el tag
                                                linea_texto_filtro_registros_aplicar_tag = df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_filtro_registros_aplicar_tag)]

                                                if linea_texto_filtro_registros_aplicar_tag == columna_df_filtro_registros_aplicar_tag_valor:

                                                    #se aplican los tags segun que marcar_toda_linea_si_todo_varia este configurado
                                                    #y si las variaciones afectan toda la linea 
                                                    if marcar_toda_linea_si_todo_varia:
                                                        self.widget_objeto.tag_add(nombre_tag
                                                                                    , f"{ind + 1}.{indice_ini_linea_texto_informar}"
                                                                                    , f"{ind + 1}.{indice_fin_linea_texto_informar}")
                                                        

                                                    else:                                                    
                                                        if not variaciones_afecta_linea_completa:

                                                            for indice_ini, indice_fin in lista_tuplas_indices_linea:
                                                                self.widget_objeto.tag_add(nombre_tag
                                                                                            , f"{ind + 1}.{indice_ini}"
                                                                                            , f"{ind + 1}.{indice_fin}")
                                                            
                                            else:
                                                #se aplican los tags segun que marcar_toda_linea_si_todo_varia este configurado
                                                #y si las variaciones afectan toda la linea 
                                                if marcar_toda_linea_si_todo_varia:
                                                    self.widget_objeto.tag_add(nombre_tag
                                                                                , f"{ind + 1}.{indice_ini_linea_texto_informar}"
                                                                                , f"{ind + 1}.{indice_fin_linea_texto_informar}")
                                                    

                                                else:                                                    
                                                    if not variaciones_afecta_linea_completa:

                                                        for indice_ini, indice_fin in lista_tuplas_indices_linea:
                                                            self.widget_objeto.tag_add(nombre_tag
                                                                                        , f"{ind + 1}.{indice_ini}"
                                                                                        , f"{ind + 1}.{indice_fin}")

            #se agrega scrollbar vertical si el texto excede el height
            if len(df_datos) > height_scrolledtext:
                self.srollbar_vertical.pack(side = "right", fill = "y")
            else:
                self.widget_objeto.vbar.pack_forget()



    def __varios_clase(self, opcion_varios: Literal["crear_widget_objeto", "agregar_scrollbar_horizontal"]):
        #rutina interna que permite crear el objeto widget y de agragarle (si esta configurado) un scrollbar horizontal

        if opcion_varios == "crear_widget_objeto":

            colocacion_scrollbar_horizontal = self.kwargs_config_widget.get(self.nombre_kwargs_colocacion_scrollbar_horizontal, None)

            #se crea el scrolledtext con o sin wrap segun los kwargs
            wrap_widget_objeto = self.kwargs_config_widget.get("wrap", None)

            if wrap_widget_objeto is None:
                self.widget_objeto = scrolledtext.ScrolledText(self.master) 
            else:
                self.widget_objeto = scrolledtext.ScrolledText(self.master, wrap = wrap_widget_objeto)

            #se quita el scrollbar vertical en la creacion (se agrega o no cuando se informa el texto)
            self.widget_objeto.vbar.pack_forget()

            #se agrega un scrollbar horizontal si esta configurado en los kwargs
            if colocacion_scrollbar_horizontal is not None:
                self.__varios_clase("agregar_scrollbar_horizontal")



        elif opcion_varios == "agregar_scrollbar_horizontal":
            #crea un scrollbar horizontal en el frame contenador del scrolledtext
            #se coloca por encima del scrolledtext a su derecha

            try:
                widget_objeto_coord_x = self.kwargs_config_widget[self.nombre_kwargs_colocacion_scrollbar_horizontal][self.nombre_kwargs_colocacion_scrollbar_horizontal_coord_x]
                widget_objeto_coord_y = self.kwargs_config_widget[self.nombre_kwargs_colocacion_scrollbar_horizontal][self.nombre_kwargs_colocacion_scrollbar_horizontal_coord_y]

                self.widget_objeto_scrollbar_horizontal = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = self.objeto_alias_tkinter_import.HORIZONTAL, command = self.widget_objeto.xview)
                self.widget_objeto.configure(xscrollcommand = self.widget_objeto_scrollbar_horizontal.set)

                self.widget_objeto_scrollbar_horizontal.place(x = widget_objeto_coord_x, y = widget_objeto_coord_y)

            except:
                pass
                    
    

###########################################################################################
#     CLASE HIJA - entry_propio
###########################################################################################

class entry_propio(gui_tkinter_widgets):
    #clase hija de la clase gui_tkinter_widgets para crear widget de tipo treeview con metodos propios asociados

    def __init__(self, master, **kwargs_config_widget):
        

        #se habilita la herencia atributos y metodos de la clase madre
        super().__init__(master, **kwargs_config_widget)


        #se asigna el nombre del diccionario kwargs donde recuperar los parametros de configuracion
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs = "entry_dicc"
        self.nombre_kwargs_formato_validacion = "formato_validacion"
        self.nombre_kwargs_texto_longitud_maxima = "texto_longitud_maxima"
        self.nombre_kwargs_titulo_messagebox_warning = "titulo_messagebox_warning"
        self.nombre_kwargs_boolean_incluir_calendario = "boolean_incluir_calendario"

        self.lista_atributos_entry_dicc = [self.nombre_kwargs_formato_validacion, self.nombre_kwargs_texto_longitud_maxima
                                           , self.nombre_kwargs_titulo_messagebox_warning, self.nombre_kwargs_boolean_incluir_calendario]

        
        #se inicializan atributos necesarios a la clase hija
        self.nombre_tipo_widget = "Entry"
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget


        self.widget_objeto = None
        self.widget_boton_calendario = None
        self.mostrar_calendario = None
        self.widget_objeto_calendario = None
        self.toplevel_calendario = None


        #se inicializa kwargs_config_widget como atributo al cual se excluyen atributos que se inicializan por separado
        # --> formato_validacion
        # --> texto_longitud_maxima
        # --> titulo_messagebox_warning
        # --> boolean_incluir_calendario
        self.formato_validacion = self.kwargs_config_widget[self.nombre_kwargs].get(self.nombre_kwargs_formato_validacion, None)
        self.texto_longitud_maxima = self.kwargs_config_widget[self.nombre_kwargs].get(self.nombre_kwargs_texto_longitud_maxima, None)
        self.titulo_messagebox_warning = self.kwargs_config_widget[self.nombre_kwargs].get(self.nombre_kwargs_titulo_messagebox_warning, None)
        self.boolean_incluir_calendario = self.kwargs_config_widget[self.nombre_kwargs].get(self.nombre_kwargs_boolean_incluir_calendario, None)



        #se inicializa como atributo el diccionario con los patrones de validacion
        #cada patron (key_1) contiene un diccionario con:
        # --> validacion_re              regla de validacion con la libreria re
        # --> mensaje_warning            mensaje que se genera al intentar salir del entry si la validacion del formato es incorrecta
        self.dicc_patrones_validacion = {
                                "entero_positivo": 
                                                    {"validacion_re": r"^\d+$"
                                                    , "mensaje_warning" : "Solo se admiten enteros positivos."
                                                    }
                                , "entero_negativo": 
                                                    {"validacion_re": r"^-\d+$"
                                                    , "mensaje_warning" : "Solo se admiten enteros negativos."
                                                    }
                                , "float_positivo": 
                                                    {"validacion_re": r"^\d*\.?\d+$"
                                                    , "mensaje_warning" : "Solo se admiten enteros o decimales positivos."
                                                    }
                                , "float_negativo": 
                                                    {"validacion_re": r"^-?\d*\.?\d+$"
                                                    , "mensaje_warning" : "Solo se admiten enteros o decimales negativos."
                                                    }
                                , "texto": 
                                                    {"validacion_re": None
                                                    , "mensaje_warning" : "Solo se admite texto REPLACE_ME."
                                                    }
                                , "fecha_ddmmaaaa": 
                                                    {"validacion_re": r"^\d{2}[-/.]\d{2}[-/.]\d{4}$"
                                                    , "mensaje_warning" : "Solo se admiten fechas en formato EUR (dd/mm/aaaa, dd-mm-aaaa o dd.mm.aaaa)."
                                                    }
                                , "fecha_yyyymmdd": 
                                                    {"validacion_re": r"^\d{4}[-/.]\d{2}[-/.]\d{2}$"
                                                    , "mensaje_warning" : "Solo se admiten fechas en formato USA (aaaa/mm/dd, aaaa-mm-dd o aaaa.mm.dd)."
                                                    }
                                , "alfanumerico": 
                                                    {"validacion_re": r"^[\w]+$"
                                                    , "mensaje_warning" :"Solo se admiten caracteres alfanumaricos."
                                                    }
                                }


        #se pocede a crear el widget
        self.__crear_widget_objeto()

        #se ejecutan los atributos nativos y propios
        self.config_atributos(**kwargs_config_widget)



    def __crear_widget_objeto(self):
        #Rutina interna que permite crear el widget

        try:
            formato_validacion_lower_no_blank = self.formato_validacion.lower().replace(" ", "")

            #se crea el widget y se le asigna la variable_enlace (creada en la clase madre)
            self.widget_objeto = self.objeto_alias_tkinter_import.Entry(self.master, textvariable = self.variable_enlace)

            #si no se configura formato_validacion o este no se encuentra en dicc_patrones_validacion (minusculas y sin espacios blancos)
            #se crea un entry normal
            if formato_validacion_lower_no_blank in list(self.dicc_patrones_validacion.keys()):
                self.widget_objeto.bind("<FocusOut>", lambda event: self.__exit_entry(event))

        except:
            pass


    def __exit_entry(self, event = None):
        #rutina interna que permite bloquear la salida del entry si los formatos de validación no corresponden a lo configurado.

        try:
            formato_validacion_lower_no_blank = self.formato_validacion.lower().replace(" ", "")

            #se recuperan los datos del diccionario dicc_patrones_validacion
            validacion_re = self.dicc_patrones_validacion[formato_validacion_lower_no_blank]["validacion_re"]
            mensaje_warning = self.dicc_patrones_validacion[formato_validacion_lower_no_blank]["mensaje_warning"]


            #se recupera el valor informado en el widget entry
            valor_entry = self.widget_objeto.get()

            if not valor_entry:
                return
            
            valor_entry = str(valor_entry)


            #se realiza el chequeo y en caso de que no coincida el valor informado en el entry se impide salir de el generando el warning
            if formato_validacion_lower_no_blank == "texto":

                if isinstance(self.texto_longitud_maxima, (int, float)) and self.texto_longitud_maxima >= 0:
                    if len(valor_entry) > self.texto_longitud_maxima:

                        mensaje_warning = mensaje_warning.replace("REPLACE_ME", f"(longitud máxima: {self.texto_longitud_maxima} caracteres)")
                        messagebox.showwarning(title = self.titulo_messagebox_warning, message = mensaje_warning)
                        self.widget_objeto.focus_set()

            else:

                if not re.fullmatch(validacion_re, valor_entry):

                    messagebox.showwarning(title = self.titulo_messagebox_warning, message = mensaje_warning)
                    self.widget_objeto.focus_set()

                else:
                    if formato_validacion_lower_no_blank == "texto" and isinstance(self.texto_longitud_maxima, (int, float)) and self.texto_longitud_maxima > 0:
                            
                        if len(valor_entry) > self.texto_longitud_maxima:

                            messagebox.showwarning(title = self.titulo_messagebox_warning, message = mensaje_warning)          
                            self.widget_objeto.focus_set()

        except:
            pass


class controltiptext:
    """
    Clase independiente (sin herencias) que permite generar un texto en pantalla cuando el usuario pone el cursor del ratón sobre un botón.
    """

    def __init__(self, widget_objeto, text, delay = 500):

        self.widget_objeto = widget_objeto
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self.after_id = None

        widget_objeto.bind("<Enter>", self.__tiempo_mostrar)
        widget_objeto.bind("<Leave>", self.__ocultar)
        widget_objeto.bind("<ButtonPress>", self.__ocultar)


    def __tiempo_mostrar(self, event = None):
        #rutina interna que aplica un tiempo de espera para que aparezca el mensaje cuando el usuario coloca el cursor del raton sobre el widget

        self.after_id = self.widget_objeto.after(self.delay, self.__mostrar)


    def __mostrar(self):
        #rutina interna que permite mostar el mensaje cuando el usuario coloca el cursor del raton sobre el widget

        if self.tipwindow or not self.text:
            return

        x = self.widget_objeto.winfo_rootx() + 20
        y = self.widget_objeto.winfo_rooty() + self.widget_objeto.winfo_height() + 5

        self.tip_ventana = tk.Toplevel(self.widget_objeto)
        self.tip_ventana.wm_overrideredirect(True)  # sin bordes
        self.tip_ventana.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tip_ventana, text = self.text, justify = "left", background = "#ffffe0", relief = "solid", borderwidth = 1, font = ("Calibri", 9))
        label.pack(ipadx = 6, ipady = 3)


    def __ocultar(self, event = None):
        #rutina interna que permite borrar el mensaje cuando el usuario al colocar el cursor del raton sobre el widget
        #lo mueve en otro sito

        try:
            if self.after_id:
                self.widget_objeto.after_cancel(self.after_id)
                self.after_id = None

            if self.tip_ventana:
                self.tip_ventana.destroy()
                self.tip_ventana = None

        except:
            pass


