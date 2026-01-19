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
from tkcalendar import Calendar


###########################################################################################
###########################################################################################
###########################################################################################
#     CLASE MADRE - gui_tkinter_widgets
###########################################################################################
###########################################################################################
###########################################################################################

class gui_tkinter_widgets():

    def __init__(self, master, tipo_widget_param = None, self_clase_gui_donde_call_rutina = None, **kwargs_config_parametros):

        """
        Clase que permite crear widgets tkinter con métodos nativos y propios.

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



        #se asigna el nombre de los atributos propios que figuran en los kwargs de configuracion
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs_dicc_config_root = "dicc_config_root"
        self.nombre_kwargs_dicc_config_root_tupla_geometry = "tupla_geometry"

        self.nombre_kwargs_dicc_colocacion = "dicc_colocacion"
        self.nombre_kwargs_dicc_colocacion_metodo = "metodo"
        self.nombre_kwargs_dicc_colocacion_coord_x = "coord_x"
        self.nombre_kwargs_dicc_colocacion_coord_y = "coord_y"


        self.nombre_kwargs_combobox_lista_opciones = "combobox_lista_opciones"

        self.nombre_kwargs_dicc_rutina = "dicc_rutina"
        self.nombre_kwargs_dicc_rutina_nombre_rutina = "rutina"
        self.nombre_kwargs_dicc_rutina_parametros_args = "parametros_args"
        self.nombre_kwargs_dicc_rutina_parametros_kwargs = "parametros_kwargs"

        self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget = "lista_dicc_rutina_aplicar_eventos_widget"
        self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_tipo_bind = "tipo_bind"
        self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina = "rutina"
        self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_args = "parametros_args"
        self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_kwargs = "parametros_kwargs"

        self.nombre_kwargs_dicc_imagen = "dicc_imagen"
        self.nombre_kwargs_dicc_imagen_png_imagen = "png_imagen"
        self.nombre_kwargs_dicc_imagen_tupla_imagen_resize = "tupla_imagen_resize"

        self.nombre_kwargs_controltiptext = "controltiptext"

        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace = "lista_dicc_rutina_trace_variable_enlace"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace = "tipo_trace"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina = "rutina"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args = "parametros_args"
        self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs = "parametros_kwargs"

        self.nombre_kwargs_alineacion = "alineacion"
        self.nombre_kwargs_bloquear = "bloquear"

        self.nombre_kwargs_listbox_lista_items = "listbox_lista_items"
        self.nombre_kwargs_listbox_lista_items_seleccionados = "listbox_lista_items_seleccionados"
        self.nombre_kwargs_listbox_listbox_seleccionar_todo_o_nada = "listbox_seleccionar_todo_o_nada"


        self.nombre_kwargs_bloquear_interaccion_nueva_ventana_con_otras = "bloquear_interaccion_nueva_ventana_con_otras"
        self.nombre_kwargs_mantener_nueva_ventana_encima_otras = "mantener_nueva_ventana_encima_otras"


        self.nombre_kwargs_destroy = "destroy"

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
        Método que permite configurar los atributos (nativos y/o propios) despues de haber creado el widget.

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
                        bloquear_interaccion_nueva_ventana_con_otras = dicc_config_root.get(self.nombre_kwargs_bloquear_interaccion_nueva_ventana_con_otras, None)
                        mantener_nueva_ventana_encima_otras = dicc_config_root.get(self.nombre_kwargs_mantener_nueva_ventana_encima_otras, None)


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

                            if bloquear_interaccion_nueva_ventana_con_otras:
                                self.widget_objeto.grab_set()

                            if mantener_nueva_ventana_encima_otras:
                                self.widget_objeto.transient(self.master)


                except:
                    pass



            # dicc_colocacion
            if self.nombre_kwargs_dicc_colocacion in lista_config_atributos_propios:

                try:
                    if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_colocacion], dict):
                        
                        metodo = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_colocacion].get(self.nombre_kwargs_dicc_colocacion_metodo, None)
                        coord_x = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_colocacion].get(self.nombre_kwargs_dicc_colocacion_coord_x, None)
                        coord_y = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_colocacion].get(self.nombre_kwargs_dicc_colocacion_coord_y, None)


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
            # --> 'parametros_args' (atributo self.nombre_kwargs_dicc_rutina_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_dicc_rutina_parametros_kwargs)
            if self.nombre_kwargs_dicc_rutina in lista_config_atributos_propios:

                try:
                    dicc_rutina = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_dicc_rutina, None)

                    if isinstance(dicc_rutina, dict):

                        nombre_rutina = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_nombre_rutina, None)
                        rutina_parametros_args = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_parametros_args, ())
                        rutina_parametros_kwargs = dicc_rutina.get(self.nombre_kwargs_dicc_rutina_parametros_kwargs, {})

                        if isinstance(nombre_rutina, str) and self.self_clase_gui_donde_call_rutina is not None:
                            rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina, None)
                        else:
                            rutina_objeto = nombre_rutina


                        if callable(rutina_objeto):
                            
                            #se usa command = lambda para que se aplique la rutina al widget en el momento de interactuar con el y no en el momento de su creacion)
                            self.widget_objeto.config(command = lambda: rutina_objeto(
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
            # lista_dicc_rutina_aplicar_eventos_widget
            ##############################################################
            #IMPORTANTE: para que la llamada de config_atributos desde una clase propia en otro modulo .py (que contiene la GUI personalizada del proyecto que sea)
            #funcione con la rutina pasada por string es necesario agregar a los kwargs de la presente clase (gui_tkinter_widgets) el entorno de la clase de la GUI
            # --> self_clase_gui_donde_call_rutina = self

            #el kwargs tiene que ser lista de diccionarios donde cada diccionario contiene las keys siguientes:
            # --> 'tipo_bind' (atributo self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_tipo_bind) --> <<ComboboxSelected>>, <<TreeviewSelect>>, <ButtonRelease-1> etc etc (aqui aplica solo si tipo_rutina = evento)
            # --> 'rutina' (atributo self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina)
            # --> 'parametros_args' (atributo self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_kwargs) 
            if self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget in lista_config_atributos_propios:

                try:
                    lista_dicc_rutina_aplicar_eventos_widget = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget, None)

                    if (isinstance(lista_dicc_rutina_aplicar_eventos_widget, list)
                        and sum(1 if isinstance(dicc, dict) else 0 for dicc in lista_dicc_rutina_aplicar_eventos_widget) == len(lista_dicc_rutina_aplicar_eventos_widget)):

                        for dicc in lista_dicc_rutina_aplicar_eventos_widget:
                            

                            tipo_bind = dicc.get(self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_tipo_bind, None)
                            nombre_rutina = dicc.get(self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina, None)
                            rutina_parametros_args = dicc.get(self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_args, ())
                            rutina_parametros_kwargs = dicc.get(self.nombre_kwargs_lista_dicc_rutina_aplicar_eventos_widget_rutina_parametros_kwargs, {})

                            if isinstance(nombre_rutina, str) and self.self_clase_gui_donde_call_rutina is not None:
                                rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina, None)
                            else:
                                rutina_objeto = nombre_rutina


                            if callable(rutina_objeto):

                                # IMPORTANTE: poner la "captura_lambda" en "captura_lambda = rutina_objeto" y "rutina_parametros_kwargs = rutina_parametros_kwargs: captura_lambda("
                                #             pq en caso de que lista_dicc_rutina_aplicar_eventos_widget tenga mas de 1 diccionario el trace ejecuta tan solo el ultimo de la lista
                                #             de ahi la necesidad de almacenar rutina_objeto de la lambda en captura_lambda
                                #             (cuando se realizan bucles sobre eventos bind con lambdas si no se procede asi solo se aplica el bind de la ultima iteracion del bucle)
                                self.widget_objeto.bind(tipo_bind
                                                        , lambda event
                                                        , captura_lambda = rutina_objeto
                                                        , rutina_parametros_args = rutina_parametros_args
                                                        , rutina_parametros_kwargs = rutina_parametros_kwargs: captura_lambda(
                                                                                                                                # args dinamicos
                                                                                                                                *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg
                                                                                                                                for arg in rutina_parametros_args),

                                                                                                                                # kwargs dinamicos
                                                                                                                                **{key: (valor(self.self_clase_gui_donde_call_rutina)
                                                                                                                                        if callable(valor) else valor)
                                                                                                                                for key, valor in rutina_parametros_kwargs.items()}
                                                                                                                            )
                                                        #add = "+" es para no machacar el resto de bind creados anteriormente
                                                        #si lista_dicc_rutina_aplicar_eventos_widget tiene mas de 1 item
                                                        , add = "+"
                                                        )

                except:
                    pass





            ##############################################################
            # lista_dicc_rutina_trace_variable_enlace
            ##############################################################
            #IMPORTANTE: para que la llamada de config_atributos desde una clase propia en otro modulo .py (que contiene la GUI personalizada del proyecto que sea)
            #funcione con la rutina pasada por string es necesario agregar a los kwargs de la presente clase (gui_tkinter_widgets) el entorno de la clase de la GUI
            # --> self_clase_gui_donde_call_rutina = self

            #el kwargs tiene que ser lista de diccionarios donde cada diccionario contiene las keys siguientes:
            # --> 'tipo_trace' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace) --> boton o evento
            # --> 'rutina' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina)
            # --> 'parametros_args' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs)
            if self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace in lista_config_atributos_propios:

                try:
                    lista_dicc_rutina_trace_variable_enlace = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace, None)

                    if isinstance(lista_dicc_rutina_trace_variable_enlace, list) and sum(1 if isinstance(dicc, dict) else 0 for dicc in lista_dicc_rutina_trace_variable_enlace) == len(lista_dicc_rutina_trace_variable_enlace):

                        for dicc in lista_dicc_rutina_trace_variable_enlace:

                            nombre_rutina = dicc.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_nombre_rutina, None)
                            tipo_trace = dicc.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_tipo_trace, None)
                            rutina_parametros_args = dicc.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_args, ())
                            rutina_parametros_kwargs = dicc.get(self.nombre_kwargs_lista_dicc_rutina_trace_variable_enlace_parametros_kwargs, {})

                            if isinstance(nombre_rutina, str) and self.self_clase_gui_donde_call_rutina is not None:
                                rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina, None)

                            else:
                                rutina_objeto = nombre_rutina

                            if callable(rutina_objeto):
                            
                                #las rutinas asociadas a eventos trace de stringvar deben tener como parametro *args de ahi "lambda *trace_args"
                                #en la rutina de la GUI del proyecto donde se use esta clase ya no es necesario ponerle *args
                                #
                                # IMPORTANTE: poner la "captura_lambda" en "captura_lambda = rutina_objeto" y "rutina_parametros_kwargs = rutina_parametros_kwargs: captura_lambda("
                                #             pq en caso de que lista_dicc_rutina_trace_variable_enlace tenga mas de 1 diccionario el trace ejecuta tan solo el ultimo de la lista
                                #             de ahi la necesidad de almacenar rutina_objeto de la lambda en captura_lambda
                                #             (cuando se realizan bucles sobre eventos bind con lambdas si no se procede asi solo se aplica el bind de la ultima iteracion del bucle)
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
                    opcion_bloqueo = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_bloquear, False)

                    if opcion_bloqueo:
                        self.widget_objeto.config(state = self.objeto_alias_tkinter_import.DISABLED)

                    else:
                        self.widget_objeto.config(state = self.objeto_alias_tkinter_import.NORMAL)

                except:
                    pass


            # listbox_varios
            if self.nombre_kwargs_listbox_lista_items in lista_config_atributos_propios:

                try:
                    if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_listbox_lista_items], (list, tuple)):

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



            #dicc_imagen
            if self.nombre_kwargs_dicc_imagen in lista_config_atributos_propios:

                if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_imagen], dict):
                    
                    png_imagen = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_imagen].get(self.nombre_kwargs_dicc_imagen_png_imagen, None)
                    tupla_imagen_resize = self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_dicc_imagen].get(self.nombre_kwargs_dicc_imagen_tupla_imagen_resize, self.imagen_en_boton_tupla_resize_por_defecto)

                    try:
                        #se redimensiona la imagen sino los botones salen distorsionados en la gui
                        img = Image.open(png_imagen).resize(tupla_imagen_resize, Image.LANCZOS)
                        img_tk = ImageTk.PhotoImage(img)

                        self.widget_objeto.image = img_tk
                        self.widget_objeto.config(**{"image": img_tk})

                    except:
                        pass




            # controltiptext
            if self.nombre_kwargs_controltiptext in lista_config_atributos_propios:
                if isinstance(self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_controltiptext], str):
                    controltiptext(self.widget_objeto, (self.kwargs_config_parametros_lower_no_blank[self.nombre_kwargs_controltiptext]))


            #destroy
            if self.nombre_kwargs_destroy in lista_config_atributos_propios:
                try:
                    opcion_destroy = self.kwargs_config_parametros_lower_no_blank.get(self.nombre_kwargs_destroy, False)

                    if opcion_destroy:
                        self.widget_objeto.destroy()
                        
                except:
                   pass


    ###########################################################################################
    #     METODOS PROPIOS - INTERNOS A LA CLASE
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
        \t\tDevuelve el objeto del alias de la libreria tkinter importada en el módulo .py (import tkinter as tk --> devuelve tk).
        
        --> lista_atributos_constructor_clase\n
        \t\tDevuelve todos los atributos incializados en el constructor de la clase (__init__) los que vienen precedidos por self.

        --> dicc_metodos_propios_clase\n
        \t\tDevuelve un diccionario que lista tantos los métodos propios públicos y privados de la clase.
 
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
###########################################################################################
###########################################################################################
#     CLASES HIJAS
###########################################################################################
###########################################################################################
###########################################################################################


###########################################################################################
#     CLASE HIJA - frame_con_scrollbar
###########################################################################################

class frame_con_scrollbar(gui_tkinter_widgets):
    #clase hija de la clase gui_tkinter_widgets para crear frames con scrollbars (vertical y/o horizontal)

    def __init__(self, master, **kwargs_config_widget):

        #se habilita la herencia atributos y metodos de la clase madre
        super().__init__(master, **kwargs_config_widget)


        #se establece que ajuste se ha realizar a la baja dentro del height del canvas contenedor del frame
        #para colocar el scrollbar horizontal si se configura en los kwargs
        self.scrollbar_vertical_coord_y_ajust_canvas = 15


        #se inicializan atributos necesarios a la clase hija
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget


        #se asigna el nombre del diccionario kwargs donde recuperar los parametros de configuracion
        #(esto es por si se desea cambiar en el futuro el nombre el mismo sin tener que modificar el codigo de la clase)
        self.nombre_kwargs_dicc_frame_scrollbar = "dicc_frame_scrollbar"
        self.nombre_kwargs_dicc_frame_scrollbar_tipo_scrollbar = "tipo_scrollbar"
        self.nombre_kwargs_dicc_frame_scrollbar_width_visible = "width_visible"
        self.nombre_kwargs_dicc_frame_scrollbar_width_total = "width_total"
        self.nombre_kwargs_dicc_frame_scrollbar_height_visible = "height_visible"
        self.nombre_kwargs_dicc_frame_scrollbar_height_total = "height_total"
        self.nombre_kwargs_dicc_frame_scrollbar_tupla_coord_place = "tupla_coord_place"
        self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling = "velocidad_scrolling"

        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical_y_horizontal = "vertical_y_horizontal"
        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_horizontal_y_vertical = "horizontal_y_vertical"
        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_horizontal = "horizontal"
        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical = "vertical"


        #se inicializan los atributos que contienen los distintos objetos necesarios
        #para crear un frame con scrollbar vertical y/o horizontal
        self.widget_objeto = None
        self.widget_objeto_canvas_contenedor = None
        self.widget_objeto_scrollbar_vertical = None
        self.widget_objeto_scrollbar_horizontal = None


        #se inicializa el atributo propio kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion
        #que almacena el kwargs de configuracion del frame con scrollbar reconstruido que se usa en su creacion
        # --> permite cuando se ejecuta el metodo modificaciones de la presente clase
        #     compararlos con los kwargs de modificaciones y localizar se hay variaciones y en este caso
        #     eliminar el frame creado previamente para volver a crearlo de nuevo
        #     pq si solo se aplican las modificaciones sobre el frame ya creado anteriormente inhabilita
        #     despues los scrollbars y el scrolling
        self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion = None

        

        #se procede a crear el widget
        # --> se calcula el kwargs reconstruido que se almacena en el atributo self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion
        #     (rutina interna __varios_clase_hija, opcion = check_kwargs_creacion)
        # --> se crea el widget con el kwargs reconstruido del paso anterior (rutina interna __varios_clase_hija, opcion = crear_widget_objeto)
        self.__varios_clase_hija("check_kwargs_creacion", **self.kwargs_config_widget)
        self.__varios_clase_hija("crear_widget_objeto", **self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion)



        #se ejecutan los atributos nativos y propios
        #(se excluyen los atributos nativos (height, width y place) y el atributo propio de la clase madre (dicc_colocacion)
        #pq ya se usan al crear el widget con la rutina __crear_widget_objeto y si se habilitan en la ejecucion del metodo propio
        #config_atributos de la clase madre hace que los scrollbars dejen de funcionar tras crear el frame scrollable)
        kwargs_config_widget_ajust = {key: valor for key, valor in kwargs_config_widget.items()
                                        if key not in ["height", "width", "place", self.nombre_kwargs_dicc_colocacion]}
        
        self.config_atributos(**kwargs_config_widget_ajust)


    def modificaciones(self, **kwargs):

        """
        Método propio de la clase 'frame_con_scrollbar' que permite modificar el frame posteriormente a su creación.\n
        Realiza mediante comparacion de los kwargs configurados y el atributo propio kwargs_dicc_frame_scrollbar_reconstruido\n
        localizar si se han de modificar el tipo de scrollbar, el tamaño y/o las coordenadas y la velocidad del scrolling del frame
        (en estos casos se elimina el frame incialmente creado y se vuelve a crear porque estas modificaciones anulan
        el canvas contenedor del frame tras su creación y los scrollbars dejan de funcionar).\n

        Al finalizar el proceso, se ejecuta el método propio 'config_atributos' de la clase madre gui_tkinter_widget.
        """

        ###############################################################################
        #se realiza la comparativa entre el diccionario de kwargs reconstruidos en la creacion del frame
        #(se saca del atributo self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion) y el diccionario de kwargs reconstruidos
        #en la modificacion (se saca con la rutina interna __varios_clase_hija, opcion = check_kwargs_modificacion, desde el kwargs
        #del presente metodo propio)
        #
        #en caso de alguna variacion en el valor de las keys (los 2 diccionarios tienen las mismas keys) se vuelve a crear el frame
        #usando como kwargs el diccionario reconstruido en la modificacion
        ###############################################################################
        kwargs_dicc_frame_scrollbar_reconstruido_modificacion = self.__varios_clase_hija("check_kwargs_modificacion", **kwargs)

        check_si_volver_a_crear_frame = (True if self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion is not None
                                                and kwargs_dicc_frame_scrollbar_reconstruido_modificacion is not None
                                                and sum(1 if valor != kwargs_dicc_frame_scrollbar_reconstruido_modificacion[key] else 0
                                                        for key, valor in self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion.items()) != 0
                                        else False)
        

        #se crea de nuevo el frame scrollable en base al kwargs reconstruido de modificacion
        if check_si_volver_a_crear_frame:
            self.__varios_clase_hija("crear_widget_objeto", **kwargs_dicc_frame_scrollbar_reconstruido_modificacion)



        #se ejecutan los atributos nativos y propios (se tenga que crear el frame scrollable de nuevo o no)
        #(se excluyen los atributos nativos (height, width y place) y el atributo propio de la clase madre (dicc_colocacion)
        #pq ya se usan al crear el widget con la rutina __crear_widget_objeto y si se habilitan en la ejecucion del metodo propio
        #config_atributos de la clase madre hace que los scrollbars dejen de funcionar tras crear el frame scrollable)
        kwargs_config_widget_ajust = {key: valor for key, valor in kwargs.items()
                                        if key not in ["height", "width", "place", self.nombre_kwargs_dicc_colocacion]}
        
        self.config_atributos(**kwargs_config_widget_ajust)



    def __varios_clase_hija(self, opcion: Literal["check_kwargs_creacion", "check_kwargs_modificacion", "crear_widget_objeto"]
                            , **kwargs):

        """
        Rutina interna que crea el frame scrollable. Realiza un check previo en los kwargs configurados y decide internamente que tipo de frame se ha de crear
        (ningún frame, frame normal o frame con scrollbar vertical y/o horizontal).\n
        """

        resultado_funcion = None

        if "check" in opcion:

            dicc_kwargs_config_frame = kwargs.get(self.nombre_kwargs_dicc_frame_scrollbar, None)

            if not isinstance(dicc_kwargs_config_frame, dict):

                dicc_kwargs_config_frame_reconstruido = {self.nombre_kwargs_dicc_frame_scrollbar_tipo_scrollbar: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_width_visible: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_width_total: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_height_visible: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_height_total: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_tupla_coord_place: None
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling: None
                                                        }

            else:
                width_visible = dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_width_visible, 0)
                width_total = dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_width_total, 0)
                height_visible = dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_height_visible, 0)
                height_total = dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_height_total, 0)
                tupla_coord_place = dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_tupla_coord_place, None)


                #se reajusta la velocidad scrolling si esta mal configurada (si no lo esta se establece al minimo es decir 1
                #y si esta configura pero no es numerica tambien se establece al minimo y si es numerico se corrije con el valor absoluto
                #para quitar negativos mal configurados)
                velocidad_scrolling = (abs(dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling, None))
                                        if isinstance(dicc_kwargs_config_frame.get(self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling, None), (int, float))
                                        else 1)



                #se chequea si se puede crear un frame y si se puede agregar scrollbar vertical y/o horizontal
                check_frame_por_crear = (True if (isinstance(tupla_coord_place, tuple)
                                                    and sum(1 if isinstance(item, (int, float)) and item > 0 else 0 for item in tupla_coord_place) == len(tupla_coord_place)
                                                    and len(tupla_coord_place) == 2)
                                        else False)

                check_agregar_scrollbar_vertical = True if height_visible + height_total != 0 and height_visible < height_total else False
                
                check_agregar_scrollbar_horizontal = True if width_visible + width_total != 0 and width_visible < width_total else False


                #se ajusta el tipo de scrollbar por ajustar segun como se hayan configurado los kwargs
                tipo_scrollbar_ajust = (self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical_y_horizontal
                                        if check_frame_por_crear and check_agregar_scrollbar_vertical and check_agregar_scrollbar_horizontal
                                        else
                                        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical
                                        if check_frame_por_crear and check_agregar_scrollbar_vertical and not check_agregar_scrollbar_horizontal
                                        else
                                        self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_horizontal
                                        if check_frame_por_crear and not check_agregar_scrollbar_vertical and check_agregar_scrollbar_horizontal
                                        else None)
                

                #se reconstruye el kwargs tras el chequeo
                dicc_kwargs_config_frame_reconstruido = {self.nombre_kwargs_dicc_frame_scrollbar_tipo_scrollbar: tipo_scrollbar_ajust
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_width_visible: width_visible
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_width_total: width_total
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_height_visible: height_visible
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_height_total: height_total
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_tupla_coord_place: tupla_coord_place
                                                        , self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling: velocidad_scrolling
                                                        }
                
                
                #se almacena dicc_kwargs_config_frame_reconstruido en en el atributo self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion
                #(esto evita cuando se aplique el metodo propio modificaciones de la presente clase tener que volver a realizar todo el chequeo)
                if opcion == "check_kwargs_creacion":
                    self.kwargs_dicc_frame_scrollbar_reconstruido_usado_creacion = dicc_kwargs_config_frame_reconstruido


                #se pasa dicc_kwargs_config_frame_reconstruido como resultado de la funcion
                elif opcion == "check_kwargs_modificacion":
                    resultado_funcion = dicc_kwargs_config_frame_reconstruido.copy()


        elif opcion == "crear_widget_objeto":

            try:
                #se recuperan los datos de los kwargs no de configuracion hecha por los usuarios sino recalculados tras pasar
                #por las opciones check_kwargs_creacion o check_kwargs_modificacion donde se calcula el tipo de scrollbar a aplicar
                tipo_scrollbar = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_tipo_scrollbar]
                width_visible = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_width_visible]
                width_total = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_width_total]
                height_visible = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_height_visible]
                height_total = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_height_total]
                tupla_coord_place = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_tupla_coord_place]
                velocidad_scrolling = kwargs[self.nombre_kwargs_dicc_frame_scrollbar_velocidad_scrolling]

                coord_x = tupla_coord_place[0]
                coord_y = tupla_coord_place[1]



                #se eliminan los objetos creados previamente y se renicializan a None los atributos que los contienen
                #(si la creacion del frame con scrollbar se hace al ejecutar el metodo modificaciones de la presente clase)
                dicc_objetos_eliminar = {"widget_objeto": self.widget_objeto
                                        , "widget_objeto_canvas_contenedor": self.widget_objeto_canvas_contenedor
                                        , "widget_objeto_scrollbar_vertical": self.widget_objeto_scrollbar_vertical
                                        , "widget_objeto_scrollbar_horizontal": self.widget_objeto_scrollbar_horizontal
                                        }

                for key, objeto in dicc_objetos_eliminar.items():
                    try:
                        objeto.destroy()
                        setattr(self, key, None)
                    except:
                        pass


                ######################################################
                # CASO 1 --> frame sin scrollbars
                ######################################################
                if tipo_scrollbar is None:
                    self.widget_objeto = self.objeto_alias_tkinter_import.Frame(self.master, width = width_visible, height = height_visible)
                    self.widget_objeto.place(x = coord_x, y = coord_y)


                ######################################################
                # CASO 2 --> frame con scrollbar vertical y/o horizontal
                ######################################################
                elif tipo_scrollbar is not None:

                    #se crea con canvas el frame visible en pantalla (depende del width_visible y height_visible)
                    self.widget_objeto_canvas_contenedor = self.objeto_alias_tkinter_import.Canvas(self.master, width = width_visible, height = height_visible, highlightthickness = 0)
                    self.widget_objeto_canvas_contenedor.place(x = coord_x, y = coord_y)


                    #se crean los scrollbars (no acorde a lo que en los kwargs se haya informado para tipo_scrollbar
                    #pero acorde a la validacion anterior de los mismos con la variable tipo_frame_por_crear)
                    if tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical_y_horizontal:

                        self.widget_objeto_scrollbar_vertical = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = "vertical", command = self.widget_objeto_canvas_contenedor.yview)
                        self.widget_objeto_scrollbar_vertical.place(x = coord_x + width_visible, y = coord_y, height = height_visible)
                        self.widget_objeto_canvas_contenedor.configure(yscrollcommand = self.widget_objeto_scrollbar_vertical.set)

                        #se usa aqui el atributo self.scrollbar_vertical_coord_y_ajust_canvas para colocar el scrollbar horizontal
                        #dentro del canvas al final (si no se hace este ajuste el scrollbar horizontal se coloca despues del canvas y no es visble en pantalla)
                        self.widget_objeto_scrollbar_horizontal = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = "horizontal", command = self.widget_objeto_canvas_contenedor.xview)
                        self.widget_objeto_scrollbar_horizontal.place(x = coord_x, y = coord_y + height_visible - self.scrollbar_vertical_coord_y_ajust_canvas, width = width_visible)
                        self.widget_objeto_canvas_contenedor.configure(xscrollcommand = self.widget_objeto_scrollbar_horizontal.set)


                    elif tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical:

                        self.widget_objeto_scrollbar_vertical = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = "vertical", command = self.widget_objeto_canvas_contenedor.yview)
                        self.widget_objeto_scrollbar_vertical.place(x = coord_x + width_visible, y = coord_y, height = height_visible)
                        self.widget_objeto_canvas_contenedor.configure(yscrollcommand = self.widget_objeto_scrollbar_vertical.set)


                    elif tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_horizontal:

                        #se usa aqui el atributo self.scrollbar_vertical_coord_y_ajust_canvas para colocar el scrollbar horizontal
                        #dentro del canvas al final (si no se hace este ajuste el scrollbar horizontal se coloca despues del canvas y no es visble en pantalla)
                        self.widget_objeto_scrollbar_horizontal = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = "horizontal", command = self.widget_objeto_canvas_contenedor.xview)
                        self.widget_objeto_scrollbar_horizontal.place(x = coord_x, y = coord_y + height_visible - self.scrollbar_vertical_coord_y_ajust_canvas, width = width_visible)
                        self.widget_objeto_canvas_contenedor.configure(xscrollcommand = self.widget_objeto_scrollbar_horizontal.set)



                    #se crea el frame dentro del canvas creado y se recupera el id del canvas con el frame integrado
                    #para poder reajustar el width del canvas con la rutina interna __varios_clase_hija_event
                    #y se agrega el objeto frame a la lista lista_objetos_scrolling
                    self.widget_objeto = self.objeto_alias_tkinter_import.Frame(self.widget_objeto_canvas_contenedor, width = width_total, height = height_total)
                    self.id_canvas = self.widget_objeto_canvas_contenedor.create_window((0, 0), window = self.widget_objeto, anchor = "nw")



                    #se asegura que cuando cambia la altura del contenido agregando widgets dentro del frame
                    #el canvas recalcula el area scrollable
                    self.widget_objeto.bind("<Configure>", lambda event: self.widget_objeto_canvas_contenedor.configure(scrollregion = self.widget_objeto_canvas_contenedor.bbox("all")))


                    #se establece la zona scrollable de forma fija desde el inicio
                    self.widget_objeto_canvas_contenedor.configure(scrollregion = (0, 0, width_total, height_total))


                    #se ajusta el width si el canvas cambia (solo para scrollbar vertical)
                    self.widget_objeto_canvas_contenedor.bind("<Configure>", lambda event: self.__varios_clase_hija_event("ajustar_width_canvas_para_scrollbar_vertical"
                                                                                                                           , event
                                                                                                                           , tipo_scrollbar = tipo_scrollbar
                                                                                                                           ), add = "+")



                    #se crea lista de objetos para poder habilitar el scroll con raton a los objetos incluidos en la lista lista_objetos_scrolling
                    #(se realiza para cada objeto de la lista un bind que se activa al colacar el taon sobre el objeto y otro que lo desactiva cuando se pierde el foco)
                    lista_objetos_scrolling = [self.widget_objeto
                                              , self.widget_objeto_canvas_contenedor
                                              , self.widget_objeto_scrollbar_vertical
                                              , self.widget_objeto_scrollbar_horizontal]
                    
                    lista_objetos_scrolling = [objeto for objeto in lista_objetos_scrolling if objeto is not None]

                    for objeto in lista_objetos_scrolling:

                        # IMPORTANTE: poner las variables captura_lambda_tipo_scrollbar y captura_lambda_velocidad_scrolling pq
                        #             cuando se realizan bucles sobre eventos bind con lambdas si no se procede asi solo se aplica el bind de la ultima iteracion del bucle
                        objeto.bind("<MouseWheel>"
                                    , lambda event
                                    , captura_lambda_tipo_scrollbar = tipo_scrollbar
                                    , captura_lambda_velocidad_scrolling = velocidad_scrolling: self.__varios_clase_hija_event("agregar_scrolling"
                                                                                                                                , event
                                                                                                                                , tipo_scrollbar = captura_lambda_tipo_scrollbar
                                                                                                                                , velocidad_scrolling = captura_lambda_velocidad_scrolling)
                                    , add = "+")



            except:
                pass


        #resultado funcion
        return resultado_funcion
        

    def __varios_clase_hija_event(self, opcion: Literal["ajustar_width_canvas_para_scrollbar_vertical", "agregar_scrolling"]
                                  , event = None
                                  , **kwargs):

        """
        Rutina interna de evento para usar en los binds al crear el frame con scrollbars.
        """

        #parametros kwargs (tipo_scrollbar se usa en las opciones de la rutina pero puede variar
        #entre los kwargs de creacion y los de modificacion del frame)
        tipo_scrollbar = kwargs.get("tipo_scrollbar", None)
        velocidad_scrolling = kwargs.get("velocidad_scrolling", None)


        if opcion == "ajustar_width_canvas_para_scrollbar_vertical":

            if tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical:
                self.widget_objeto_canvas_contenedor.itemconfig(self.id_canvas, width = event.width)


        elif opcion == "agregar_scrolling":

            velocidad_scrolling_ajust = int(-1 * (event.delta / 120) * velocidad_scrolling)

            if tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical_y_horizontal:
                #no se declara las 2 lineas siguientes pq al hacer el scroll con el raton mueve los 2 scrollbars al mismo tiempo
                # --> self.widget_objeto_canvas_contenedor.yview_scroll(velocidad_scrolling_ajust, "units")
                # --> self.widget_objeto_canvas_contenedor.xview_scroll(velocidad_scrolling_ajust, "units")
                #
                #para scrolling horizontal se establece el scrolling solo cuando el usuario
                #tiene la tecla Shift pulsada y hace scrolling con el raton (event.state & 0x0001)

                if event.state & 0x0001:
                    self.widget_objeto_canvas_contenedor.xview_scroll(velocidad_scrolling_ajust, "units")
                else:
                    self.widget_objeto_canvas_contenedor.yview_scroll(velocidad_scrolling_ajust, "units")


            elif tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_vertical:
                self.widget_objeto_canvas_contenedor.yview_scroll(velocidad_scrolling_ajust, "units")


            elif tipo_scrollbar == self.nombre_kwargs_dicc_frame_scrollbar_scrollbar_horizontal:
                self.widget_objeto_canvas_contenedor.xview_scroll(velocidad_scrolling_ajust, "units")




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
        self.nombre_kwargs_dicc_entry = "dicc_entry"
        self.nombre_kwargs_dicc_entry_formato_validacion = "formato_validacion"
        self.nombre_kwargs_dicc_entry_texto_longitud_maxima = "texto_longitud_maxima"
        self.nombre_kwargs_dicc_entry_titulo_messagebox_warning = "titulo_messagebox_warning"
        self.nombre_kwargs_dicc_entry_calendario_tupla_coord_place_y_width = "calendario_tupla_coord_place_y_width"
        self.nombre_kwargs_dicc_entry_calendario_iconbitmap = "calendario_iconbitmap"

        self.lista_atributos_entry_dicc = [self.nombre_kwargs_dicc_entry_formato_validacion, self.nombre_kwargs_dicc_entry_texto_longitud_maxima
                                           , self.nombre_kwargs_dicc_entry_titulo_messagebox_warning, self.nombre_kwargs_dicc_entry_calendario_tupla_coord_place_y_width
                                           , self.nombre_kwargs_dicc_entry_calendario_iconbitmap]

        
        #se inicializan atributos necesarios a la clase hija
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget


        self.widget_objeto = None
        self.widget_boton_calendario = None
        self.mostrar_calendario = None
        self.widget_objeto_calendario = None
        self.toplevel_calendario = None


        self.formato_validacion = None
        self.texto_longitud_maxima = None
        self.titulo_messagebox_warning = None
        self.calendario_tupla_coord_place_y_width = None
        self.calendario_iconbitmap = None


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
        self.__varios_clase_hija("crear_widget_objeto")

        #se ejecutan los atributos nativos y propios
        self.config_atributos(**kwargs_config_widget)



    def __varios_clase_hija(self, opcion_varios: Literal["crear_widget_objeto"
                                                        , "calendario_para_fecha_crear_boton"
                                                        , "calendario_para_fecha_abrir_toplevel_al_pulsar_boton"
                                                        , "calendario_para_fecha_informar_entry_valor_seleccionado"
                                                        , "calendario_eliminar_toplevel_cuando_focus_out"]):
        
        """
        Rutina interna (hibrido entre rutina y función) que permite realizar acciones varias dentro de la clase entry_propio.

        opcion_varios:
        --

        --> crear_widget_objeto\n
        \t\tCrea el widget.

        --> calendario_para_fecha_crear_boton\n
        \t\tCrea el botón calendario y lo colo según en las coordenadas configuradas en el atributo propio calendario_coord_place en los kwargs al llamar la clase.

        --> calendario_para_fecha_abrir_toplevel_al_pulsar_boton\n
        \t\tcrea un toplevel en el cual se incrusta un calendario (libreria Calendar) al cual se le agrega un boton para poder informar en el entry la fecha seleccionada.

        --> calendario_para_fecha_informar_entry_valor_seleccionado\n
        \t\tpermite, tras pulsar el botón 'seleccionar' dentro del toplevel del calendario, informar el entry se hace mediante la variable_enlace del entry (heredada de la clase madre gui_tkinter_widgets).

        --> calendario_eliminar_toplevel_cuando_focus_out\n
        \t\tpermite eliminar el calendario cuando se pierde el foco en el mismo.

        """

        if opcion_varios == "crear_widget_objeto":

            try:
                self.formato_validacion = self.kwargs_config_widget[self.nombre_kwargs_dicc_entry].get(self.nombre_kwargs_dicc_entry_formato_validacion, None)
                self.texto_longitud_maxima = self.kwargs_config_widget[self.nombre_kwargs_dicc_entry].get(self.nombre_kwargs_dicc_entry_texto_longitud_maxima, None)
                self.titulo_messagebox_warning = self.kwargs_config_widget[self.nombre_kwargs_dicc_entry].get(self.nombre_kwargs_dicc_entry_titulo_messagebox_warning, None)
                self.calendario_tupla_coord_place_y_width = self.kwargs_config_widget[self.nombre_kwargs_dicc_entry].get(self.nombre_kwargs_dicc_entry_calendario_tupla_coord_place_y_width, None)
                self.calendario_iconbitmap = self.kwargs_config_widget[self.nombre_kwargs_dicc_entry].get(self.nombre_kwargs_dicc_entry_calendario_iconbitmap, None)


                formato_validacion_lower_no_blank = self.formato_validacion.lower().replace(" ", "")

                #se crea el widget y se le asigna la variable_enlace (creada en la clase madre)
                self.widget_objeto = self.objeto_alias_tkinter_import.Entry(self.master, textvariable = self.variable_enlace)

                #se configura la regla de validacion
                if formato_validacion_lower_no_blank in list(self.dicc_patrones_validacion.keys()):
                    self.widget_objeto.bind("<FocusOut>", lambda event: self.__exit_entry(event))


                #se integra el boton de calendario en caso de que formato de validacion contenga "fecha"
                #y que self.calendario_tupla_coord_place_y_width sea tupla de 3 items numericos positivos
                if (isinstance(self.calendario_tupla_coord_place_y_width, tuple) 
                    and sum(1 if isinstance(item, (int, float)) and item >= 0 else 0 for item in self.calendario_tupla_coord_place_y_width) == len(self.calendario_tupla_coord_place_y_width)
                    and "fecha" in formato_validacion_lower_no_blank):

                    self.__varios_clase_hija("calendario_para_fecha_crear_boton")

            except:
                pass



        elif opcion_varios == "calendario_para_fecha_crear_boton":
            #crea el boton de acceso al calendario y lo coloca segun lo configurado en el atributo propio calendario_coord_place

            formato_validacion_lower_no_blank = self.formato_validacion.lower().replace(" ", "")

            widget_objeto_coord_x = self.calendario_tupla_coord_place_y_width[0]
            widget_objeto_coord_y = self.calendario_tupla_coord_place_y_width[1]
            widget_objeto_width = self.calendario_tupla_coord_place_y_width[2]

            self.widget_boton_calendario = (self.objeto_alias_tkinter_import.Button(self.master, text = "📅", width = widget_objeto_width
                                                                                    , command = lambda: self.__varios_clase_hija("calendario_para_fecha_abrir_toplevel_al_pulsar_boton")))
            
            self.widget_boton_calendario.place(x = widget_objeto_coord_x, y = widget_objeto_coord_y)


        elif opcion_varios == "calendario_para_fecha_abrir_toplevel_al_pulsar_boton":
            #crea un toplevel en el cual se incrusta un calendario (libreria Calendar)
            #al cual se le agrega un boton para poder informar en el entry la fecha seleccionada
        

            formato_validacion_lower_no_blank = self.formato_validacion.lower().replace(" ", "")

            date_pattern = "dd/mm/yyyy" if formato_validacion_lower_no_blank == "fecha_ddmmaaaa" else "yyyy/mm/dd"

            #se crea el toplevel
            self.toplevel_calendario = self.objeto_alias_tkinter_import.Toplevel(self.master)
            self.toplevel_calendario.resizable(0, 0)

            if self.calendario_iconbitmap is not None:
                try:
                    self.toplevel_calendario.iconbitmap(self.calendario_iconbitmap)
                except:
                    pass#por si self.calendario_iconbitmap no es un .ico

            self.toplevel_calendario.transient()#impide que teniendo el calendario abierto y al clicar en otra ventana el toplevel pase por debajo de la ventana

            self.toplevel_calendario.wm_attributes("-toolwindow", True)#se deja solo el boton cerrar toplevel

            self.toplevel_calendario.bind("<FocusOut>", lambda event: self.__varios_clase_hija("calendario_eliminar_toplevel_cuando_focus_out"))



            #se crea el calendario
            self.widget_objeto_calendario = globals()["Calendar"](self.toplevel_calendario, date_pattern = date_pattern)
            self.widget_objeto_calendario.pack()

            boton_seleccionar = (self.objeto_alias_tkinter_import.Button(self.toplevel_calendario, text = "SELECCIONAR", font = ("Calibri", 10, "bold")
                                                                        , command = lambda: self.__varios_clase_hija("calendario_para_fecha_informar_entry_valor_seleccionado")))
            boton_seleccionar.pack(pady = 5)



        elif opcion_varios == "calendario_para_fecha_informar_entry_valor_seleccionado":
            #permite tras pulsar el boton seleccionar dentro del toplevel del calendario informar el entry
            #se hace mediante la variable_enlace del entry

            fecha_seleccionada = self.widget_objeto_calendario.get_date()
            self.variable_enlace.set(fecha_seleccionada)
            self.toplevel_calendario.destroy()

        
        elif opcion_varios == "calendario_eliminar_toplevel_cuando_focus_out":
            #permite eliminar el calendario cuando se pierde el foco en el
            
            self.toplevel_calendario.destroy()
            



    def __exit_entry(self, event = None):
        
        """Rutina interna de evento que permite bloquear la salida del entry si los formatos de validación no corresponden a lo configurado.
        """

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
        self.nombre_kwargs_dicc_treeview = "dicc_treeview"
        self.nombre_kwargs_dicc_treeview_seleccion_item = "seleccion_item"
        self.nombre_kwargs_dicc_treeview_seleccion_item_ninguno = "ninguno"
        self.nombre_kwargs_dicc_treeview_seleccion_item_simple = "simple"
        self.nombre_kwargs_dicc_treeview_seleccion_item_multiple = "multiple"
        self.nombre_kwargs_dicc_treeview_height = "height"
        self.nombre_kwargs_dicc_treeview_columnas_df = "columnas_df"
        self.nombre_kwargs_dicc_treeview_columnas_treeview = "columnas_treeview"
        self.nombre_kwargs_dicc_treeview_width_columnas_treeview = "width_columnas_treeview"

        self.nombre_kwargs_dicc_treeview_rutina_click_item = "dicc_rutina_click_item"
        self.nombre_kwargs_dicc_treeview_rutina_click_item_nombre_rutina = "rutina"
        self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_args = "parametros_args"
        self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_kwargs = "parametros_kwargs"


        #se inicializa el atributo propio kwargs_usados_creacion_antes_modificaciones
        #que almacena el kwargs de configuracion del frame con scrollbar reconstruido que se usa en su creacion
        # --> permite cuando se ejecuta el metodo modificaciones de la presente clase
        #     compararlos con los kwargs de modificaciones (dicc_treeview) y localizar se hay variaciones y en este caso
        #     eliminar el treeview creado previamente para volver a crearlo de nuevo con las modificaciones
        self.kwargs_usados_creacion_antes_modificaciones = {self.nombre_kwargs_dicc_treeview: None
                                                            , self.nombre_kwargs_dicc_treeview_rutina_click_item: None}


        #se inicializa el atributo dicc_conversion_tipo_seleccion_item
        #los valores de las keys han de ser nativos de tkinter
        self.dicc_conversion_tipo_seleccion_item = {self.nombre_kwargs_dicc_treeview_seleccion_item_ninguno: "none"
                                                    , self.nombre_kwargs_dicc_treeview_seleccion_item_simple: "browse"
                                                    , self.nombre_kwargs_dicc_treeview_seleccion_item_multiple: "extended"}
        

        #se inicializan atributos necesarios a la clase hija
        self.master = master
        self.clase_madre_nombre = self.__class__.__bases__[0].__name__ if self.__class__.__bases__[0].__name__ != "object" else None
        self.clase_nombre = self.__class__.__name__
        self.kwargs_config_widget = kwargs_config_widget
        self.self_clase_gui_donde_call_rutina = self_clase_gui_donde_call_rutina


        self.widget_objeto = None
        self.datos_item_seleccionado = None


        #se pocede a crear el widget
        self.__varios_clase_hija("crear_widget_objeto")


        #se ejecutan los atributos nativos y propios
        self.config_atributos(**kwargs_config_widget)


    def modificaciones(self, opcion: Literal["actualizar_desde_df", "modificar_objeto"], **kwargs):

        """
        Método que permite rellenar el treeview o modificar el objeto.

        args
        --
        --> actualizar_desde_df\n
        Permite rellenar el treeview usando un dataframe.

        --> modificar_objeto\n
        Permite modificar el treeview (cambiar el tipo de selección, agregar / quitar columnas, cambiar el ancho y/o alto del treeview
        y asignarle otra rutina al clicar en items).

        kwargs
        --
        --> df_datos\n
        Se usa tan solo si se ha optado por el parámetro args 'actualizar_desde_df'.

        --> kwargs_modificacion_treeview\n
        Se usa tan solo si se ha optado por el parámetro args 'modificar_objeto'.
        """

        #parametros kwargs
        df_datos = kwargs.get("df_datos", None)


        if opcion == "actualizar_desde_df":

            if self.widget_objeto is not None:

                if isinstance(df_datos, pd.DataFrame):

                    try:
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

                    except:
                        pass


        elif opcion == "modificar_objeto":

            #se recuperan los kwargs usados en la creacion del treeview previamente a la ejecucion del presente metodo modificaciones
            dicc_treeview_usado_creacion = self.kwargs_usados_creacion_antes_modificaciones.get(self.nombre_kwargs_dicc_treeview, {})
            dicc_treeview_usado_creacion_seleccion_item = dicc_treeview_usado_creacion.get(self.nombre_kwargs_dicc_treeview_seleccion_item, None)
            dicc_treeview_usado_creacion_height = dicc_treeview_usado_creacion.get(self.nombre_kwargs_dicc_treeview_height, None)
            dicc_treeview_usado_creacion_columnas_df = dicc_treeview_usado_creacion.get(self.nombre_kwargs_dicc_treeview_columnas_df, None)
            dicc_treeview_usado_creacion_columnas_treeview = dicc_treeview_usado_creacion.get(self.nombre_kwargs_dicc_treeview_columnas_treeview, None)
            dicc_treeview_usado_creacion_width_columnas_treeview = dicc_treeview_usado_creacion.get(self.nombre_kwargs_dicc_treeview_width_columnas_treeview, None)

            dicc_treeview_rutina_click_item_usado_creacion = self.kwargs_usados_creacion_antes_modificaciones.get(self.nombre_kwargs_dicc_treeview_rutina_click_item, {})
            dicc_treeview_rutina_click_item_usado_creacion_nombre_rutina = dicc_treeview_rutina_click_item_usado_creacion.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_nombre_rutina, None)
            dicc_treeview_rutina_click_item_usado_creacion_parametros_args = dicc_treeview_rutina_click_item_usado_creacion.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_args, None)
            dicc_treeview_rutina_click_item_usado_creacion_parametros_kwargs = dicc_treeview_rutina_click_item_usado_creacion.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_kwargs, None)

            #se recuperan los kwargs de modificaciones del treeview pasados como parametros del presente metodo modificaciones
            dicc_treeview_modificaciones = kwargs.get(self.nombre_kwargs_dicc_treeview, {})
            dicc_treeview_modificaciones_seleccion_item = dicc_treeview_modificaciones.get(self.nombre_kwargs_dicc_treeview_seleccion_item, None)
            dicc_treeview_modificaciones_height = dicc_treeview_modificaciones.get(self.nombre_kwargs_dicc_treeview_height, None)
            dicc_treeview_modificaciones_columnas_df = dicc_treeview_modificaciones.get(self.nombre_kwargs_dicc_treeview_columnas_df, None)
            dicc_treeview_modificaciones_columnas_treeview = dicc_treeview_modificaciones.get(self.nombre_kwargs_dicc_treeview_columnas_treeview, None)
            dicc_treeview_modificaciones_width_columnas_treeview = dicc_treeview_modificaciones.get(self.nombre_kwargs_dicc_treeview_width_columnas_treeview, None)

            dicc_treeview_rutina_click_item_modificaciones =kwargs.get(self.nombre_kwargs_dicc_treeview_rutina_click_item, {})
            dicc_treeview_rutina_click_item_modificaciones_nombre_rutina = dicc_treeview_rutina_click_item_modificaciones.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_nombre_rutina, None)
            dicc_treeview_rutina_click_item_modificaciones_parametros_args = dicc_treeview_rutina_click_item_modificaciones.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_args, None)
            dicc_treeview_rutina_click_item_modificaciones_parametros_kwargs = dicc_treeview_rutina_click_item_modificaciones.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_kwargs, None)


            #se chequea si los kwargs configurados (dicc_treeview y dicc_treeview_rutina_click_item) para aplicar en modificaciones
            #requieren volver a crear el treeview pq han variado con respecto a los kwargs usados en la creacion previa del treeview
            #y almacenadas en el atributo kwargs_usados_creacion_antes_modificaciones
            check_si_volver_a_crear_treeview = (True if dicc_treeview_usado_creacion_seleccion_item != dicc_treeview_modificaciones_seleccion_item
                                                        or dicc_treeview_usado_creacion_height != dicc_treeview_modificaciones_height
                                                        or dicc_treeview_usado_creacion_columnas_df != dicc_treeview_modificaciones_columnas_df
                                                        or dicc_treeview_usado_creacion_columnas_treeview != dicc_treeview_modificaciones_columnas_treeview
                                                        or dicc_treeview_usado_creacion_width_columnas_treeview != dicc_treeview_modificaciones_width_columnas_treeview
                                                        or dicc_treeview_rutina_click_item_usado_creacion_nombre_rutina != dicc_treeview_rutina_click_item_modificaciones_nombre_rutina
                                                        or dicc_treeview_rutina_click_item_usado_creacion_parametros_args != dicc_treeview_rutina_click_item_modificaciones_parametros_args
                                                        or dicc_treeview_rutina_click_item_usado_creacion_parametros_kwargs != dicc_treeview_rutina_click_item_modificaciones_parametros_kwargs
                                                else False)
            

            #se crea de nuevo el frame scrollable en base al kwargs reconstruido de modificacion
            if check_si_volver_a_crear_treeview:
                self.__varios_clase_hija("crear_widget_objeto_desde_modificaciones", **kwargs)



    def __varios_clase_hija(self, opcion: Literal["crear_widget_objeto", "crear_widget_objeto_desde_modificaciones"], **kwargs_modificaciones):
        
        """Rutina interna que permite crear el widget (creación nueva o modificación).
        """


        kwargs_usados = self.kwargs_config_widget if opcion == "crear_widget_objeto" else kwargs_modificaciones

        try:
            #se elimina el treeview creado anteriormente (solo para opcion = crear_widget_objeto_desde_modificaciones)
            if opcion == "crear_widget_objeto_desde_modificaciones":
                self.widget_objeto.destroy()


            #se crea el objeto treeview
            seleccion_item = kwargs_usados[self.nombre_kwargs_dicc_treeview].get(self.nombre_kwargs_dicc_treeview_seleccion_item, self.nombre_kwargs_dicc_treeview_seleccion_item_simple) 
            
            seleccion_item_ajust = (self.dicc_conversion_tipo_seleccion_item[seleccion_item]
                                    if seleccion_item in list(self.dicc_conversion_tipo_seleccion_item.keys())
                                    else "browse")
                  
            self.widget_objeto = ttk.Treeview(self.master, selectmode = seleccion_item_ajust)


            #se realizan ajustes especificos al objeto treeview
            height = kwargs_usados[self.nombre_kwargs_dicc_treeview][self.nombre_kwargs_dicc_treeview_height]
            columnas_df = kwargs_usados[self.nombre_kwargs_dicc_treeview][self.nombre_kwargs_dicc_treeview_columnas_df]
            columnas_treeview = kwargs_usados[self.nombre_kwargs_dicc_treeview][self.nombre_kwargs_dicc_treeview_columnas_treeview]
            width_columnas_treeview = kwargs_usados[self.nombre_kwargs_dicc_treeview][self.nombre_kwargs_dicc_treeview_width_columnas_treeview]

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
            # --> lista_columnas_df                 es la lista de las columnas del df configuradas en el atributo propio treeview_dicc
            # --> lista_columnas_treeview           es la lista de las columnas del treeview configuradas en el atributo propio treeview_dicc
            # --> lista_width_columnas_treeview     es la lista de los width de las columnas del treeview configuradas en el atributo propio treeview_dicc
            # --> lista_tipo_dato_columna_df        es la lista de los tipos de datos de las columnas del df del parametro df_datos
            # --> lista_datos_item_seleccionado     es la lista de los datos de los items seleccionados (es lista de lista)
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

            #el kwargs tiene que tener un diccionario 'dicc_rutina_click_item' (atributo self.nombre_kwargs_dicc_treeview_rutina_click_item) con las keys siguientes:
            # --> 'rutina' (atributo self.nombre_kwargs_dicc_treeview_rutina_click_item_nombre_rutina)
            # --> 'parametros_args' (atributo self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_args)
            # --> 'parametros_kwargs' (atributo self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_kwargs)
            dicc_rutina_click_item = kwargs_usados.get(self.nombre_kwargs_dicc_treeview_rutina_click_item, None)

            try:
                if isinstance(dicc_rutina_click_item, dict):

                    nombre_rutina_click_item = dicc_rutina_click_item.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_nombre_rutina, None)
                    rutina_click_item_parametros_args = dicc_rutina_click_item.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_args, ())
                    rutina_click_item_parametros_kwargs = dicc_rutina_click_item.get(self.nombre_kwargs_dicc_treeview_rutina_click_item_parametros_kwargs, {})

                    if isinstance(nombre_rutina_click_item, str) and self.self_clase_gui_donde_call_rutina is not None:
                        rutina_objeto = getattr(self.self_clase_gui_donde_call_rutina, nombre_rutina_click_item, None)

                    else:
                        rutina_objeto = nombre_rutina_click_item

                    if callable(rutina_objeto):

                        self.widget_objeto.bind("<ButtonRelease-1>", lambda event: rutina_objeto(
                                                                                                #args dinamicos
                                                                                                *(arg(self.self_clase_gui_donde_call_rutina) if callable(arg) else arg for arg in rutina_click_item_parametros_args)
                                                                                                
                                                                                                #kwargs dinamicos
                                                                                                , **{key: (valor(self.self_clase_gui_donde_call_rutina) if callable(valor) else valor)
                                                                                                        for key, valor in rutina_click_item_parametros_kwargs.items()}
                                                                                                )
                                                #add = "+" es para no machacar el bind __click_on_item
                                                , add = "+"
                                                )
            except:
                pass

            finally:

                #se almacena el widget_objeto y el kwargs (dicc_treeview) + rutina al hacer click en item usado en la creacion en el atributo kwargs_usados_creacion_antes_modificaciones
                #para saber cuando se aplica el metodo propio modificaciones (opcion = modificar_objeto)
                self.kwargs_usados_creacion_antes_modificaciones[self.nombre_kwargs_dicc_treeview] = kwargs_usados[self.nombre_kwargs_dicc_treeview]
                self.kwargs_usados_creacion_antes_modificaciones[self.nombre_kwargs_dicc_treeview_rutina_click_item] = kwargs_usados[self.nombre_kwargs_dicc_treeview_rutina_click_item]




    def __click_on_item(self, event = None):
        
        """Rutina interna que permite recuperar los datos de los items seleccionados (es lista de lista).
        """

        item_selecc = self.widget_objeto.selection()
        self.datos_item_seleccionado["lista_datos_item_seleccionado"] = [list(self.widget_objeto.item(item_id, "values")) for item_id in item_selecc] if item_selecc else None



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
        self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal = "colocacion_scrollbar_horizontal"
        self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal_metodo = "metodo"
        self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal_coord_x = "coord_x"
        self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal_coord_y = "coord_y"

        self.nombre_kwargs_scrolledtext_propio_df_datos = "df_datos"
        self.nombre_kwargs_scrolledtext_propio_columna_df_para_informar = "columna_df_para_informar"

        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa = "lista_dicc_tag_linea_completa"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_nombre_tag = "nombre_tag"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_columna_df_tag_aplicar = "columna_df_tag_aplicar"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_case_sensitive = "case_sensitive"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_dicc_config = "dicc_config"

        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa = "lista_dicc_tag_caracteres_cambiantes_comparativa"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag = "nombre_tag"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag = "columna_df_filtro_registros_aplicar_tag"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor = "columna_df_filtro_registros_aplicar_tag_valor"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1 = "columna_df_comparar_1"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2 = "columna_df_comparar_2"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_case_sensitive = "case_sensitive"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_marcar_toda_linea_si_todo_varia = "marcar_toda_linea_si_todo_varia"
        self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_dicc_config = "dicc_config"


        #se inicializan atributos necesarios a la clase hija
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


    def modificaciones(self, opcion_modif: Literal["borrar_contenido_y_tags"
                                                   , "agregar_solo_contenido_desde_string"
                                                   , "agregar_solo_contenido_desde_dataframe"
                                                   , "agregar_contenido_y_tags_desde_dataframe"]
                                                   , **kwargs):

        """
        Método que permite realizar modificaciones sobre el contenido y/o tags en el scrolledtext

        
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

        
        kwargs
        --
        Los kwargs que se detallan a continuacín se explican en el manual MANUAL_tkinter_utils_v1.0 (y posteriores) en el repositorio Github.\n
        Se detalla aqui tan solo cuando usarlos según el parámetro args 'opcion_modif' seleccionado.\n

        --> string_texto_informar\n
        Aplica cuando opcion_modif = 'agregar_solo_contenido_desde_string' 

        --> height_scrolledtext\n
        Aplica cuando opcion_modif = 'agregar_solo_contenido_desde_string', 'agregar_solo_contenido_desde_dataframe' o 'agregar_contenido_y_tags_desde_dataframe'. 

        --> df_datos\n
        Aplica cuando opcion_modif = 'agregar_solo_contenido_desde_dataframe' o 'agregar_contenido_y_tags_desde_dataframe'. 

        --> columna_df_para_informar\n
        Aplica cuando opcion_modif = 'agregar_solo_contenido_desde_dataframe' o 'agregar_contenido_y_tags_desde_dataframe'. 

        --> lista_dicc_tag_linea_completa\n
        Aplica cuando opcion_modif = 'agregar_contenido_y_tags_desde_dataframe'. 

        --> lista_dicc_tag_caracteres_cambiantes_comparativa\n
        Aplica cuando opcion_modif = 'agregar_contenido_y_tags_desde_dataframe'. 
        """

        #parametros kwargs
        string_texto_informar = kwargs.get("string_texto_informar", None)
        height_scrolledtext = kwargs.get("height_scrolledtext", None)
        
        df_datos = kwargs.get(self.nombre_kwargs_scrolledtext_propio_df_datos, None)
        columna_df_para_informar = kwargs.get(self.nombre_kwargs_scrolledtext_propio_columna_df_para_informar, None)
        lista_dicc_tag_linea_completa = kwargs.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa, None)
        lista_dicc_tag_caracteres_cambiantes_comparativa = kwargs.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa, None)


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
            self.widget_objeto.insert(self.objeto_alias_tkinter_import.END, string_texto_informar)
            numero_lineas_informadas = len(self.widget_objeto.get("1.0", self.objeto_alias_tkinter_import.END).splitlines())

            self.widget_objeto.config(state = self.objeto_alias_tkinter_import.DISABLED)

            #se agrega scrollbar vertical si el texto excede el height
            try:
                if numero_lineas_informadas > height_scrolledtext:
                    self.srollbar_vertical.pack(side = "right", fill = "y")
                else:
                    self.widget_objeto.vbar.pack_forget()
            except:
                pass


        elif opcion_modif == "agregar_solo_contenido_desde_dataframe":

            if columna_df_para_informar is not None:
                existe_columna_en_df = "si" if columna_df_para_informar in [columna for columna in df_datos.columns] else "no"

                if existe_columna_en_df == "si":
                    df_datos.reset_index(drop = True, inplace = True)

                    for ind in df_datos.index:
                        self.widget_objeto.insert(self.objeto_alias_tkinter_import.END, df_datos.iloc[ind, df_datos.columns.get_loc(columna_df_para_informar)] + "\n")

            #se agrega scrollbar vertical si el texto excede el height
            try:
                if len(df_datos) > height_scrolledtext:
                    self.srollbar_vertical.pack(side = "right", fill = "y")
                else:
                    self.widget_objeto.vbar.pack_forget()
            except:
                pass



        elif opcion_modif == "agregar_contenido_y_tags_desde_dataframe":

            #se crean los tags configurados si lista_dicc_tag_linea_completa es lista
            if isinstance(lista_dicc_tag_linea_completa, list):

                for ind, dicc_tag in enumerate(lista_dicc_tag_linea_completa):
                    nombre_tag = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_nombre_tag, None)
                    dicc_config = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_dicc_config, None)

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
                    nombre_tag = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag, None)
                    dicc_config = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_dicc_config, None)

                    if isinstance(nombre_tag, str):
                        
                        if isinstance(dicc_config, dict):
                            for atributo, valor_atributo in dicc_config.items():
                                try:
                                    self.widget_objeto.tag_configure(nombre_tag, **{atributo: valor_atributo})

                                except:
                                    pass




            #se informa el scrolledtext con los tags (si lista_dicc_tag no esta configurada o no esta bien configurada solo se agregan los datos sin tags) 
            if isinstance(df_datos, pd.DataFrame) and columna_df_para_informar is not None:

                try:

                    lista_columnas_df = [columna for columna in df_datos.columns]
                    existe_columna_en_df = "si" if columna_df_para_informar in lista_columnas_df else "no"

                    if existe_columna_en_df == "si":
                        #se extraen de df_datos la columna que sirve para informar el contenido + las que sirven para los tags

                        lista_columnas_df_tags_linea_completa = []
                        lista_columnas_filtro_registros_aplicar_tag = []
                        lista_columnas_df_filtro_registros_aplicar_tag_valor = []
                        lista_columnas_df_comparar_1 = []
                        lista_columnas_df_comparar_2 = []

                        if isinstance(lista_dicc_tag_linea_completa, list):

                            if self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_columna_df_tag_aplicar is not None:

                                lista_columnas_df_tags_linea_completa = [dicc_tag[self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_columna_df_tag_aplicar] 
                                                                        for dicc_tag in lista_dicc_tag_linea_completa
                                                                        if dicc_tag[self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_columna_df_tag_aplicar] in lista_columnas_df]                                                                
                                                                    

                        if isinstance(lista_dicc_tag_caracteres_cambiantes_comparativa, list):

                            lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag = self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag
                            lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor = self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor
                            lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1 = self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1
                            lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2 = self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2

                            if lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag is not None:

                                lista_columnas_filtro_registros_aplicar_tag = [dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag] 
                                                                                for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                                if dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag]
                                                                                    in lista_columnas_df]
                                
                                lista_columnas_df_filtro_registros_aplicar_tag_valor = [dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor] 
                                                                                        for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                                        if dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor] 
                                                                                            in lista_columnas_df]



                            if lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1 is not None:

                                lista_columnas_df_comparar_1 = [dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1] 
                                                                for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                if dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1] in lista_columnas_df]


                            if lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2 is not None:

                                lista_columnas_df_comparar_2 = [dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2] 
                                                                for dicc_tag in lista_dicc_tag_caracteres_cambiantes_comparativa
                                                                if dicc_tag[lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2] in lista_columnas_df]


                        #se concatenan las listas anteriores, se quitan los duplicados y se convierte en df
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
                                    nombre_tag = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_nombre_tag, None)
                                    columna_df_tag_aplicar = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_columna_df_tag_aplicar, None)
                                    case_sensitive = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_linea_completa_case_sensitive, False)

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

                                    nombre_tag = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_nombre_tag, None)
                                    columna_df_filtro_registros_aplicar_tag = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag, None)
                                    columna_df_filtro_registros_aplicar_tag_valor = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_filtro_registros_aplicar_tag_valor, None)
                                    columna_df_comparar_1 = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_1, None)
                                    columna_df_comparar_2 = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_columna_df_comparar_2, None)
                                    case_sensitive = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_case_sensitive, False)
                                    marcar_toda_linea_si_todo_varia = dicc_tag.get(self.nombre_kwargs_scrolledtext_propio_lista_dicc_tag_caracteres_cambiantes_comparativa_marcar_toda_linea_si_todo_varia, False)


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
                                                                                                                           

                except:
                    pass


            #se agrega scrollbar vertical si el texto excede el height
            try:
                if len(df_datos) > height_scrolledtext:
                    self.srollbar_vertical.pack(side = "right", fill = "y")
                else:
                    self.widget_objeto.vbar.pack_forget()
            except:
                pass



    def __varios_clase(self, opcion_varios: Literal["crear_widget_objeto", "agregar_scrollbar_horizontal"]):

        """Rutina interna que permite crear el objeto widget y de agregarle (si esta configurado) un scrollbar horizontal.
        """

        if opcion_varios == "crear_widget_objeto":

            colocacion_scrollbar_horizontal = self.kwargs_config_widget.get(self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal, None)

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
                widget_objeto_coord_x = self.kwargs_config_widget[self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal][self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal_coord_x]
                widget_objeto_coord_y = self.kwargs_config_widget[self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal][self.nombre_kwargs_scrolledtext_propio_colocacion_scrollbar_horizontal_coord_y]

                self.widget_objeto_scrollbar_horizontal = self.objeto_alias_tkinter_import.Scrollbar(self.master, orient = self.objeto_alias_tkinter_import.HORIZONTAL, command = self.widget_objeto.xview)
                self.widget_objeto.configure(xscrollcommand = self.widget_objeto_scrollbar_horizontal.set)

                self.widget_objeto_scrollbar_horizontal.place(x = widget_objeto_coord_x, y = widget_objeto_coord_y)

            except:
                pass

   
###########################################################################################
###########################################################################################
###########################################################################################
#     CLASES INDEPENDIENTES
###########################################################################################
###########################################################################################
###########################################################################################

class controltiptext:

    """
    Clase independiente (sin herencias) que permite generar un texto en pantalla cuando el usuario pone el cursor del ratón sobre un widget.
    """

    def __init__(self, widget_objeto, mensaje_controltiptext, delay = 500):

        self.widget_objeto = widget_objeto
        self.mensaje_controltiptext = mensaje_controltiptext
        self.delay = delay
        self.tipwindow = None
        self.after_id = None
        self.objeto_alias_tkinter_import = self.__varios_clase("objeto_alias_libreria_python_import", libreria_python = "tkinter")

        widget_objeto.bind("<Enter>", lambda event: self.__varios_clase_event("tiempo_espera_para_mostrar_mensaje_controltiptext", event))
        widget_objeto.bind("<Leave>", lambda event: self.__varios_clase_event("eliminar_mensaje_controltiptext_cuando_focus_out", event))
        widget_objeto.bind("<ButtonPress>", lambda event: self.__varios_clase_event("eliminar_mensaje_controltiptext_cuando_focus_out", event))


    def __varios_clase(self, opcion: Literal["objeto_alias_libreria_python_import", "mostrar_mensaje_controltiptext_cuando_focus_widget"], **kwargs):

        """
        Rutina / función interna que realiza varias acciones

        --> objeto_alias_libreria_python_import\n
        \t\tDevuelve el objeto del alias de la libreria tkinter importada en el módulo .py (import tkinter as tk --> devuelve tk).

        --> mostrar_mensaje_controltiptext_cuando_focus_widget
        \t\tPermite mostar el mensaje cuando el usuario coloca el cursor del ratón sobre el widget.
        """

        resultado_funcion = None
        libreria_python = kwargs.get("libreria_python", None)


        if opcion == "objeto_alias_libreria_python_import":
            #devuelve el objeto del alias de la libreria tkinter importada en el módulo .py (import tkinter as tk --> devuelve tk)

            dicc_globals = dict(globals())
            for _, objeto in dicc_globals.items():
                if objeto is sys.modules.get(libreria_python):
                    resultado_funcion = objeto
                    break

        elif opcion == "mostrar_mensaje_controltiptext_cuando_focus_widget":
            #permite mostar el mensaje cuando el usuario coloca el cursor del ratón sobre el widget

            if self.tipwindow or not self.mensaje_controltiptext:
                return

            x = self.widget_objeto.winfo_rootx() + 20
            y = self.widget_objeto.winfo_rooty() + self.widget_objeto.winfo_height() + 5

            self.tip_ventana = self.objeto_alias_tkinter_import.Toplevel(self.widget_objeto)
            self.tip_ventana.wm_overrideredirect(True)  # sin bordes
            self.tip_ventana.wm_geometry(f"+{x}+{y}")

            label = (self.objeto_alias_tkinter_import.Label(self.tip_ventana, text = self.mensaje_controltiptext, justify = "left"
                                                            , background = "#ffffe0", relief = "solid", borderwidth = 1, font = ("Calibri", 9)))
            
            label.pack(ipadx = 6, ipady = 3)


        #resultado funcion
        return resultado_funcion


    def __varios_clase_event(self, opcion: Literal["tiempo_espera_para_mostrar_mensaje_controltiptext"
                                                   , "eliminar_mensaje_controltiptext_cuando_focus_out"]
                                , event = None):

        """
        Rutina interna de evento para usar en los binds al crear el frame con scrollbars.

        args:
        --> tiempo_espera_para_mostrar_mensaje_controltiptext
        \n\nAplica un tiempo de espera para que aparezca el mensaje cuando el usuario coloca el cursor del raton sobre el widget\n

        --> eliminar_mensaje_controltiptext_cuando_focus_out
        \n\nPermite borrar el mensaje cuando el usuario al colocar el cursor del raton sobre el widget lo mueve en otro sito
        """


        if opcion == "tiempo_espera_para_mostrar_mensaje_controltiptext":
            #aplica un tiempo de espera para que aparezca el mensaje cuando el usuario coloca el cursor del raton sobre el widget

            self.after_id = self.widget_objeto.after(self.delay, self.__varios_clase("mostrar_mensaje_controltiptext_cuando_focus_widget"))


        elif opcion == "eliminar_mensaje_controltiptext_cuando_focus_out":
            #permite borrar el mensaje cuando el usuario al colocar el cursor del raton sobre el widget lo mueve en otro sito

            try:
                if self.after_id:
                    self.widget_objeto.after_cancel(self.after_id)
                    self.after_id = None

                if self.tip_ventana:
                    self.tip_ventana.destroy()
                    self.tip_ventana = None

            except:
                pass




