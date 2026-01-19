import os
import sys
import pathlib
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog as fd
import pandas as pd
from tkcalendar import Calendar

#importasr el modulo (OBLIGATORIO)
import tkinter_UTILS_v1_0 as mod_utils


##########################################################################################################
#IMPORTANTE
##########################################################################################################
#para ejecutar los ejemplos uno a uno establecer los bloques de codigo a comentarios (ctrl + ç en Visual Studio Code)
#de los ejemplos que no se van a usar y dejar como codigo activo el que si se va a usar




##########################################################################################################
##########################################################################################################
# EJEMPLO 1 - root
##########################################################################################################
##########################################################################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root_1 = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }


kwargs_config_root_2 = {"dicc_config_root":
                                    {"title": "OTRA PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root_1)

root.config_atributos(**kwargs_config_root_2)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 2 - frames dentro de un root
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))


kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

root.widget_objeto.mainloop()



##########################################################################################################
##########################################################################################################
# EJEMPLO 3 - boton
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

png_para_boton = (os.path.join(sys._MEIPASS, "png_para_boton.png")
                    if getattr(sys, 'frozen', False)
                    else os.path.join(pathlib.Path(__file__).parent.absolute(), "png_para_boton.png"))


class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_boton_1 =  {"text": "botón"
                                                , "width": 5
                                                , "bg": "black"
                                                , "fg": "white"
                                                , "controltiptext": "este botón solo imprime"
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                , "dicc_rutina":
                                                                {"rutina": "def_rutina_boton_1"
                                                                , "parametros_args": ("args_1",)
                                                                , "parametros_kwargs": {"kwargs_1": "prueba_1", "kwargs_2": "prueba_1"}
                                                                }
                                                }

        self.kwargs_config_combobox = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "bd": 2
                                        , "relief": "solid"
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 50}
                                        , "combobox_lista_opciones": ["A", "B", "C"]
                                        }

        self.kwargs_config_boton_2 =  {"width": 40
                                                , "dicc_imagen": {"png_imagen": png_para_boton, "tupla_imagen_resize": (23, 23)}
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 48}
                                                , "dicc_rutina":
                                                                {"rutina": "def_rutina_boton_2"
                                                                , "parametros_args": (lambda widget: self.combobox.widget_objeto.get(),)
                                                                }
                                                }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)

        self.boton_1 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton_1)
        self.boton_2 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton_2)

        self.combobox = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", **self.kwargs_config_combobox)


    def def_rutina_boton_1(self, opcion_boton_1, **kwargs):

        kwargs_1 = kwargs.get("kwargs_1", None)

        if opcion_boton_1 == "args_1":
            print("prueba aaa")

        elif opcion_boton_1 == "args_2":
            print("prueba bbb")

        if kwargs_1 == "prueba_1":
            print("prueba ccc")

    def def_rutina_boton_2(self, opcion_boton_2):

        if opcion_boton_2 == "A":
            print("prueba ddd")

        elif opcion_boton_2 == "B":
            print("prueba eee")

        elif opcion_boton_2 == "C":
            print("prueba fff")



if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

png_para_boton = (os.path.join(sys._MEIPASS, "png_para_boton.png")
                    if getattr(sys, 'frozen', False)
                    else os.path.join(pathlib.Path(__file__).parent.absolute(), "png_para_boton.png"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_boton_1 =  {"text": "botón"
                        , "width": 5
                        , "bg": "black"
                        , "fg": "white"
                        , "controltiptext": "este botón solo imprime"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton_1"
                                        , "parametros_args": ("args_1",)
                                        , "parametros_kwargs": {"kwargs_1": "prueba_1", "kwargs_2": "prueba_1"}
                                        }
                        }

kwargs_config_combobox = {"font": ("Calibri", 10)
                        , "width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "justify": tk.LEFT
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 50}
                        , "combobox_lista_opciones": ["A", "B", "C"]
                        }

kwargs_config_boton_2 =  {"width": 40
                        , "dicc_imagen": {"png_imagen": png_para_boton, "tupla_imagen_resize": (23, 23)}
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 48}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton_2"
                                        , "parametros_args": (lambda widget: combobox.widget_objeto.get(),)
                                        }
                        }


modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

def def_rutina_boton_1(opcion_boton_1, **kwargs):

    kwargs_1 = kwargs.get("kwargs_1", None)

    if opcion_boton_1 == "args_1":
        print("prueba aaa")

    elif opcion_boton_1 == "args_2":
        print("prueba bbb")

    if kwargs_1 == "prueba_1":
        print("prueba ccc")

def def_rutina_boton_2(opcion_boton_2):

    if opcion_boton_2 == "A":
        print("prueba ddd")

    elif opcion_boton_2 == "B":
        print("prueba eee")

    elif opcion_boton_2 == "C":
        print("prueba fff")

combobox = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", **kwargs_config_combobox)
boton_1 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton_1)
boton_2 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton_2)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 4 - frames scrollables dentro del root
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        lista_kwargs_frame = [
                                {"width": 5000
                                , "height": 5000
                                , "bg": "#ACADB1"
                                , "bd": 2
                                , "relief": "solid"
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 100}
                                , "dicc_frame_scrollbar": {"width_visible": 450
                                                            , "width_total": 2000
                                                            , "height_visible": 150
                                                            , "height_total": 630
                                                            , "tupla_coord_place": (10, 10)
                                                            , "velocidad_scrolling": 3
                                                            }
                                }
        
                                , {"bg": "#88E271"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_frame_scrollbar": {"width_visible": 450
                                                                , "width_total": 450
                                                                , "height_visible": 150
                                                                , "height_total": 2000
                                                                , "tupla_coord_place": (500, 10)
                                                                , "velocidad_scrolling": 3
                                                                }
                                    }
        
                                , {"bg": "#DFC968"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_frame_scrollbar": {"width_visible": 450
                                                                , "width_total": 2000
                                                                , "height_visible": 150
                                                                , "height_total": 150
                                                                , "tupla_coord_place": (10, 210)
                                                                , "velocidad_scrolling": 3
                                                                }
                                    }
        
                                , {"bg": "#E3ABF0"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_frame_scrollbar": {"width_visible": 450
                                                                , "width_total": 2000
                                                                , "height_visible": 150
                                                                , "height_total": 630
                                                                , "tupla_coord_place": (500, 210)
                                                                , "velocidad_scrolling": 3
                                                                }
                                    }
                            ]


        self.lista_objetos_frame = []
        for ind_kwargs_label, kwargs_frame in enumerate(lista_kwargs_frame):

            kwargs_label = {"text": None
                            , "bg": kwargs_frame["bg"]
                            , "width": 20
                            , "font": ("Calibri", 12, "bold")
                            , "alineacion": "left"
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                            }
            
            kwargs_label["text"] = ("vertical_y_horizontal" if ind_kwargs_label in [0, 3]
                                    else "vertical" if ind_kwargs_label == 1
                                    else "horizontal" if ind_kwargs_label == 2
                                    else "")
            

            frame = mod_utils.frame_con_scrollbar(self.master.widget_objeto, **kwargs_frame)
            mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_label)

            self.lista_objetos_frame.append(frame)

            for coord in ["x", "y"]:

                kwargs_config_label_frame = {"text": None
                                            , "width": 15
                                            , "bd": 2
                                            , "bg": "black"
                                            , "fg": "white"
                                            , "relief": "solid"
                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                            }

                for ind in range(1, 20, 1):

                    kwargs_config_label_frame["text"] = f"label_{ind}" if ind == 1 else f"label_{ind}_{coord}"
                    kwargs_config_label_frame["dicc_colocacion"][f"coord_{coord}"] = (ind - 1) * 150 + 10 if coord == "x" else (ind - 1) * 30 + 40

                    mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_config_label_frame)

                    kwargs_config_label_frame["text"] = None


        kwargs_config_boton =  {"text": "modificar frame"
                                , "bg": "red"
                                , "fg": "white"
                                , "width": 15
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 450, "coord_y": 380}
                                , "dicc_rutina":
                                                {"rutina": "def_rutina_boton"}
                                }
        
        self.boton = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **kwargs_config_boton)


    def def_rutina_boton(self):

        kwargs_frame_modif = {"bg": "#CFE970"
                                , "bd": 2
                                , "relief": "solid"
                                , "dicc_frame_scrollbar": {"width_visible": 450
                                                            , "width_total": 450
                                                            , "height_visible": 150
                                                            , "height_total": 150
                                                            , "tupla_coord_place": (500, 210)
                                                            , "velocidad_scrolling": 3
                                                            }
                                }
        
        kwargs_frame_modif_label = {"text": "Frame cambiado de 'vertical y horizontal' a 'sin scrollbars'"
                                    , "bg": "#CFE970"
                                    , "fg": "red"
                                    , "width": 50
                                    , "font": ("Calibri", 12, "bold")
                                    , "alineacion": "left"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        #se modifica el frame 4 (el de abajo - derecha)
        self.frame_por_modificar = self.lista_objetos_frame[3]
        self.frame_por_modificar.modificaciones(**kwargs_frame_modif)

        mod_utils.gui_tkinter_widgets(self.frame_por_modificar.widget_objeto, tipo_widget_param = "label", **kwargs_frame_modif_label)

        for coord in ["x", "y"]:

            kwargs_config_label_frame = {"text": None
                                        , "width": 15
                                        , "bd": 2
                                        , "bg": "black"
                                        , "fg": "white"
                                        , "relief": "solid"
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                        }

            for ind in range(1, 20, 1):

                kwargs_config_label_frame["text"] = f"label_{ind}" if ind == 1 else f"label_{ind}_{coord}"
                kwargs_config_label_frame["dicc_colocacion"][f"coord_{coord}"] = (ind - 1) * 150 + 10 if coord == "x" else (ind - 1) * 30 + 40

                mod_utils.gui_tkinter_widgets(self.frame_por_modificar.widget_objeto, tipo_widget_param = "label", **kwargs_config_label_frame)

                kwargs_config_label_frame["text"] = None



if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (1000, 450)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()



##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))


kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (1000, 450)
                                    , "resizable": (0, 0)
                                    }
                    }

lista_kwargs_frame = [
                        {"width": 5000
                        , "height": 5000
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 100, "coord_y": 100}
                        , "dicc_frame_scrollbar": {"width_visible": 450
                                                    , "width_total": 2000
                                                    , "height_visible": 150
                                                    , "height_total": 630
                                                    , "tupla_coord_place": (10, 10)
                                                    , "velocidad_scrolling": 3
                                                    }
                        }

                        , {"bg": "#88E271"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_frame_scrollbar": {"width_visible": 450
                                                        , "width_total": 450
                                                        , "height_visible": 150
                                                        , "height_total": 2000
                                                        , "tupla_coord_place": (500, 10)
                                                        , "velocidad_scrolling": 3
                                                        }
                            }

                        , {"bg": "#DFC968"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_frame_scrollbar": {"width_visible": 450
                                                        , "width_total": 2000
                                                        , "height_visible": 150
                                                        , "height_total": 150
                                                        , "tupla_coord_place": (10, 210)
                                                        , "velocidad_scrolling": 3
                                                        }
                            }

                        , {"bg": "#E3ABF0"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_frame_scrollbar": {"width_visible": 450
                                                        , "width_total": 2000
                                                        , "height_visible": 150
                                                        , "height_total": 630
                                                        , "tupla_coord_place": (500, 210)
                                                        , "velocidad_scrolling": 3
                                                        }
                            }
                    ]


root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)

lista_objetos_frame = []
for ind_kwargs_label, kwargs_frame in enumerate(lista_kwargs_frame):

    kwargs_label = {"text": None
                    , "bg": kwargs_frame["bg"]
                    , "width": 20
                    , "font": ("Calibri", 12, "bold")
                    , "alineacion": "left"
                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                    }

    kwargs_label["text"] = ("vertical_y_horizontal" if ind_kwargs_label in [0, 3]
                            else "vertical" if ind_kwargs_label == 1
                            else "horizontal" if ind_kwargs_label == 2
                            else "")

    frame = mod_utils.frame_con_scrollbar(root.widget_objeto, **kwargs_frame)
    mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_label)

    lista_objetos_frame.append(frame)

    for coord in ["x", "y"]:

        kwargs_config_label_frame = {"text": None
                                    , "width": 15
                                    , "bd": 2
                                    , "bg": "black"
                                    , "fg": "white"
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                    }

        for ind in range(1, 20, 1):

            kwargs_config_label_frame["text"] = f"label_{ind}" if ind == 1 else f"label_{ind}_{coord}"
            kwargs_config_label_frame["dicc_colocacion"][f"coord_{coord}"] = (ind - 1) * 150 + 10 if coord == "x" else (ind - 1) * 30 + 40

            mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_config_label_frame)

            kwargs_config_label_frame["text"] = None



def def_rutina_boton():

    kwargs_frame_modif = {"bg": "#CFE970"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_frame_scrollbar": {"width_visible": 450
                                                        , "width_total": 450
                                                        , "height_visible": 150
                                                        , "height_total": 150
                                                        , "tupla_coord_place": (500, 210)
                                                        , "velocidad_scrolling": 3
                                                        }
                            }
    
    kwargs_frame_modif_label = {"text": "Frame cambiado de 'vertical y horizontal' a 'sin scrollbars'"
                                , "bg": "#CFE970"
                                , "fg": "red"
                                , "width": 50
                                , "font": ("Calibri", 12, "bold")
                                , "alineacion": "left"
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                }

    #se modifica el frame 4 (el de abajo - derecha)
    frame_por_modificar = lista_objetos_frame[3]
    frame_por_modificar.modificaciones(**kwargs_frame_modif)

    mod_utils.gui_tkinter_widgets(frame_por_modificar.widget_objeto, tipo_widget_param = "label", **kwargs_frame_modif_label)

    for coord in ["x", "y"]:

        kwargs_config_label_frame = {"text": None
                                    , "width": 15
                                    , "bd": 2
                                    , "bg": "black"
                                    , "fg": "white"
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                    }

        for ind in range(1, 20, 1):

            kwargs_config_label_frame["text"] = f"label_{ind}" if ind == 1 else f"label_{ind}_{coord}"
            kwargs_config_label_frame["dicc_colocacion"][f"coord_{coord}"] = (ind - 1) * 150 + 10 if coord == "x" else (ind - 1) * 30 + 40

            mod_utils.gui_tkinter_widgets(frame_por_modificar.widget_objeto, tipo_widget_param = "label", **kwargs_config_label_frame)

            kwargs_config_label_frame["text"] = None



kwargs_config_boton =  {"text": "modificar frame"
                        , "bg": "red"
                        , "fg": "white"
                        , "width": 15
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 450, "coord_y": 380}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton"}
                        }

modulo_python_actual = sys.modules[__name__]

boton = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton)

root.widget_objeto.mainloop()



##########################################################################################################
##########################################################################################################
# EJEMPLO 5 - label
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_label = {"text": "prueba_label"
                                    , "width": 15
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.label = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "label", **self.kwargs_config_label)

        self.strvar_label = self.label.variable_enlace
        self.strvar_label.set("prueba stringvar")

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_label = {"text": "prueba_label"
                        , "width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
label = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_config_label)

strvar_label = label.variable_enlace
strvar_label.set("prueba stringvar")

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 6 - entry (con gui_tkinter_widgets)
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_entry = {"width": 20
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.entry = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "entry", **self.kwargs_config_entry)

        self.strvar_entry = self.entry.variable_enlace
        self.strvar_entry.set("prueba stringvar entry")

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_entry = {"width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
entry = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "entry", **kwargs_config_entry)

strvar_entry = entry.variable_enlace
strvar_entry.set("prueba stringvar entry")

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 7 - entry (con entry_propio) con reglas validacion
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_entry_1 = {"width": 15
                                , "bd": 2
                                , "relief": "solid"
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                , "dicc_entry":
                                            {"formato_validacion": "entero_positivo"
                                            , "titulo_messagebox_warning": "titulo messagebox"
                                            }
                                }

        # self.kwargs_config_entry_1 = {"width": 15
        #                             , "bd": 2
        #                             , "relief": "solid"
        #                             , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
        #                             , "dicc_entry":
        #                                         {"formato_validacion": "texto"
        #                                         , "texto_longitud_maxima": 5
        #                                         , "titulo_messagebox_warning": "titulo messagebox"
        #                                         }
        #                             }

        self.kwargs_config_entry_2 = {"width": 15
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        
        self.entry_1 = mod_utils.entry_propio(self.frame.widget_objeto, **self.kwargs_config_entry_1)
        self.entry_2 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "entry", **self.kwargs_config_entry_2)

        self.strvar_entry = self.entry_1.variable_enlace

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()

##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_entry_1 = {"width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        , "dicc_entry":
                                    {"formato_validacion": "entero_positivo"
                                    , "titulo_messagebox_warning": "titulo messagebox"
                                    }
                        }

# kwargs_config_entry_1 = {"width": 15
#                             , "bd": 2
#                             , "relief": "solid"
#                             , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
#                             , "dicc_entry":
#                                         {"formato_validacion": "texto"
#                                         , "texto_longitud_maxima": 5
#                                         , "titulo_messagebox_warning": "titulo messagebox"
#                                         }
#                             }

kwargs_config_entry_2 = {"width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                        }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

entry_1 = mod_utils.entry_propio(frame.widget_objeto, **kwargs_config_entry_1)
entry_2 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "entry", **kwargs_config_entry_2)

strvar_entry = entry_1.variable_enlace

strvar_entry.set("prueba")

root.widget_objeto.mainloop()



##########################################################################################################
##########################################################################################################
# EJEMPLO 8 - entry fecha (con entry_propio) con calendario
##########################################################################################################
##########################################################################################################

##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_entry = {"width": 15
                                , "bd": 2
                                , "relief": "solid"
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                , "dicc_entry":
                                            {"formato_validacion": "fecha_ddmmaaaa"
                                            , "titulo_messagebox_warning": "titulo messagebox"
                                            , "calendario_tupla_coord_place_y_width": (120, 10, 5)
                                            , "calendario_iconbitmap": ico_tapar_pluma_tkinter
                                            }
                                }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.entry = mod_utils.entry_propio(self.frame.widget_objeto, **self.kwargs_config_entry)

        self.strvar_entry = self.entry.variable_enlace

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_entry = {"width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        , "dicc_entry":
                                    {"formato_validacion": "fecha_ddmmaaaa"
                                    , "titulo_messagebox_warning": "titulo messagebox"
                                    , "calendario_tupla_coord_place_y_width": (120, 10, 5)
                                    , "calendario_iconbitmap": ico_tapar_pluma_tkinter
                                    }
                        }


root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
entry = mod_utils.entry_propio(frame.widget_objeto, **kwargs_config_entry)

strvar_entry = entry.variable_enlace

strvar_entry.set("prueba")

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 9 - combobox
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_combobox = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "bd": 2
                                        , "relief": "solid"
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                        , "combobox_lista_opciones": ["A", "B"]
                                        }
        
        self.kwargs_config_boton =  {"text": "botón"
                                    , "width": 5
                                    , "bg": "black"
                                    , "fg": "white"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                                    , "dicc_rutina":
                                                    {"rutina": "def_rutina_boton"
                                                    , "parametros_args": (lambda widget: self.combobox.widget_objeto.get(),)
                                                    }
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.combobox = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", **self.kwargs_config_combobox)
        self.boton = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton)

        self.strvar_combobox = self.combobox.variable_enlace

    def def_rutina_boton(self, opcion_boton):
        
        if opcion_boton == "A":
            print("seleccion opcion A")

        elif opcion_boton == "B":
            self.combobox.config_atributos(**{"combobox_lista_opciones": ["A", "B", "C"]})


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_combobox = {"font": ("Calibri", 10)
                        , "width": 15
                        , "bd": 2
                        , "relief": "solid"
                        , "justify": tk.LEFT
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        , "combobox_lista_opciones": ["A", "B"]
                        }

kwargs_config_boton =  {"text": "botón"
                        , "width": 5
                        , "bg": "black"
                        , "fg": "white"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton"
                                        , "parametros_args": (lambda widget: combobox.widget_objeto.get(),)
                                        }
                        }

modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
combobox = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", **kwargs_config_combobox)

strvar_combobox = combobox.variable_enlace

def def_rutina_boton(opcion_boton):
    
    if opcion_boton == "A":
        print("seleccion opcion A")

    elif opcion_boton == "B":
        combobox.config_atributos(**{"combobox_lista_opciones": ["A", "B", "C"]})

boton = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton)
root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 10 - listbox
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_ventana_inicio:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_root = {"dicc_config_root":
                                                {"title": "PRUEBA ROOT"
                                                , "iconbitmap": ico_tapar_pluma_tkinter
                                                , "tupla_geometry": (400, 200)
                                                , "resizable": (0, 0)
                                                }
                                }

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.kwargs_config_listbox_seleccion_simple = {"font": ("Calibri", 10)
                                                    , "width": 15
                                                    , "height": 5
                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                    , "selectmode": "single"
                                                    , "exportselection": False
                                                    , "listbox_lista_items": ["A", "B", "C"]
                                                    }

        self.kwargs_config_listbox_seleccion_multiple = {"font": ("Calibri", 10)
                                                        , "width": 15
                                                        , "height": 5
                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 80}
                                                        , "selectmode": "multiple"
                                                        , "exportselection": False
                                                        , "listbox_lista_items": ["A", "B", "C"]
                                                        }

        self.kwargs_config_boton = {"text": "botón"
                                    , "font": ("Calibri", 10)
                                    , "bg": "black"
                                    , "fg": "white"
                                    , "width": 5
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                                    , "dicc_rutina":
                                                    {"rutina": "def_rutina_listbox"}
                                    }


        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)

        self.listbox_simple = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "listbox", **self.kwargs_config_listbox_seleccion_simple)
        self.listbox_multiple = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "listbox", **self.kwargs_config_listbox_seleccion_multiple)

        self.boton = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton)


    def def_rutina_listbox(self):
        lista_items_seleccionados_simple = self.listbox_simple.config_atributos(**{"listbox_lista_items_seleccionados": True})
        lista_items_seleccionados_multiplke = self.listbox_multiple.config_atributos(**{"listbox_lista_items_seleccionados": True})

        print(lista_items_seleccionados_simple)
        print(lista_items_seleccionados_multiplke)

        self.listbox_simple.config_atributos(**{"listbox_lista_items": ["A", "B", "C", "D"]})
        self.listbox_multiple.config_atributos(**{"listbox_lista_items": ["A", "B", "C", "D"]})

        self.listbox_simple.config_atributos(**{"listbox_seleccionar_todo_o_nada": True, "exportselection": False})
        self.listbox_multiple.config_atributos(**{"listbox_seleccionar_todo_o_nada": True, "exportselection": False})


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                        }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_ventana_inicio(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_listbox_seleccion_simple = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "height": 5
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                        , "selectmode": "single"
                                        , "exportselection": False
                                        , "listbox_lista_items": ["A", "B", "C"]
                                        }

kwargs_config_listbox_seleccion_multiple = {"font": ("Calibri", 10)
                                            , "width": 15
                                            , "height": 5
                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 80}
                                            , "selectmode": "multiple"
                                            , "exportselection": False
                                            , "listbox_lista_items": ["A", "B", "C"]
                                            }

kwargs_config_boton = {"text": "botón"
                        , "font": ("Calibri", 10)
                        , "bg": "black"
                        , "fg": "white"
                        , "width": 5
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 10}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_listbox"}
                        }

modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

listbox_simple = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "listbox", **kwargs_config_listbox_seleccion_simple)
listbox_multiple = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "listbox", **kwargs_config_listbox_seleccion_multiple)


def def_rutina_listbox():
    lista_items_seleccionados_simple = listbox_simple.config_atributos(**{"listbox_lista_items_seleccionados": True})
    lista_items_seleccionados_multiplke = listbox_multiple.config_atributos(**{"listbox_lista_items_seleccionados": True})

    print(lista_items_seleccionados_simple)
    print(lista_items_seleccionados_multiplke)

    listbox_simple.config_atributos(**{"listbox_lista_items": ["A", "B", "C", "D"]})
    listbox_multiple.config_atributos(**{"listbox_lista_items": ["A", "B", "C", "D"]})

    listbox_simple.config_atributos(**{"listbox_seleccionar_todo_o_nada": True, "exportselection": False})
    listbox_multiple.config_atributos(**{"listbox_seleccionar_todo_o_nada": True, "exportselection": False})


boton = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 11 - treeview
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.df_datos = pd.DataFrame({"COLUMNA_1": ["prueba 1", "prueba 2", "prueba 3"]
                                , "COLUMNA_2": ["aa", "bb", "cc"]
                                , "COLUMNA_3": [1.5, 1, 1]
                                , "COLUMNA_4": ["dd", "ee", "ff"]})

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.kwargs_config_treeview = {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 10}
                                        , "dicc_treeview": {"seleccion_item": "simple"
                                                            , "height": 4
                                                            ,"columnas_df": ["COLUMNA_1", "COLUMNA_2", "COLUMNA_3"]
                                                            , "columnas_treeview": ["col A", "col B", "col C"]
                                                            , "width_columnas_treeview": [50, 50, 50]
                                                            }
                                        , "dicc_rutina_click_item": {"rutina": "def_rutina_click_item"}
                                        }
        
        self.kwargs_config_boton_1 = {"text": "update df"
                                        , "width": 10
                                        , "bg": "black"
                                        , "fg": "white"
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 140}
                                        , "dicc_rutina":
                                                        {"rutina": "def_rutina_boton_1"}
                                    }
        
        self.kwargs_config_boton_2 = {"text": "modificar"
                                        , "width": 15
                                        , "bg": "black"
                                        , "fg": "white"
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 140}
                                        , "dicc_rutina":
                                                        {"rutina": "def_rutina_boton_2"}
                                    }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.treeview = mod_utils.treeview_propio(self.frame.widget_objeto, self_clase_gui_donde_call_rutina = self, **self.kwargs_config_treeview)
        self.boton_1 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton_1)
        self.boton_2 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton_2)


    def def_rutina_click_item(self):
        diccionario_datos_item_seleccionado = self.treeview.datos_item_seleccionado
        print(diccionario_datos_item_seleccionado)


    def def_rutina_boton_1(self):
        self.treeview.modificaciones("actualizar_desde_df", df_datos = self.df_datos)


    def def_rutina_boton_2(self):

        kwargs_config_treeview_update = {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 10}
                                        , "dicc_treeview": {"seleccion_item": "multiple"
                                                            , "height": 4
                                                            , "columnas_df": ["COLUMNA_1", "COLUMNA_2", "COLUMNA_3", "COLUMNA_4"]
                                                            , "columnas_treeview": ["col A", "col B", "col C", "col D"]
                                                            , "width_columnas_treeview": [50, 50, 50, 50]
                                                            }
                                        , "dicc_rutina_click_item": {"rutina": "def_rutina_click_item"}
                                        }

        self.treeview.modificaciones("modificar_objeto", **kwargs_config_treeview_update)
        self.treeview.config_atributos(**kwargs_config_treeview_update)

        self.treeview.modificaciones("actualizar_desde_df", df_datos = self.df_datos)


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))


df_datos = pd.DataFrame({"COLUMNA_1": ["prueba 1", "prueba 2", "prueba 3"]
                        , "COLUMNA_2": ["aa", "bb", "cc"]
                        , "COLUMNA_3": [1.5, 1, 1]
                        , "COLUMNA_4": ["dd", "ee", "ff"]})

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (1, 1)
                                    }
                        }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_treeview = {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 10}
                        , "dicc_treeview": {"seleccion_item": "simple"
                                            , "height": 4
                                            ,"columnas_df": ["COLUMNA_1", "COLUMNA_2", "COLUMNA_3"]
                                            , "columnas_treeview": ["col A", "col B", "col C"]
                                            , "width_columnas_treeview": [50, 50, 50]
                                            }
                        , "dicc_rutina_click_item": {"rutina": "def_rutina_click_item"}
                        }

kwargs_config_boton_1 = {"text": "update df"
                        , "width": 10
                        , "bg": "black"
                        , "fg": "white"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 140}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton_1"}
                        }

kwargs_config_boton_2 = {"text": "modificar"
                        , "width": 15
                        , "bg": "black"
                        , "fg": "white"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 140}
                        , "dicc_rutina":
                                        {"rutina": "def_rutina_boton_2"}
                        }


modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root, tipo_widget_param = "frame", **kwargs_config_frame)


def def_rutina_click_item():
    diccionario_datos_item_seleccionado = treeview.datos_item_seleccionado
    print(diccionario_datos_item_seleccionado)


def def_rutina_boton_1():
    treeview.modificaciones("actualizar_desde_df", df_datos = df_datos)


def def_rutina_boton_2():

        kwargs_config_treeview_update = {"dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 10}
                                        , "dicc_treeview": {"seleccion_item": "multiple"
                                                            , "height": 4
                                                            ,"columnas_df": ["COLUMNA_1", "COLUMNA_2", "COLUMNA_3", "COLUMNA_4"]
                                                            , "columnas_treeview": ["col A", "col B", "col C", "col D"]
                                                            , "width_columnas_treeview": [50, 50, 50, 50]
                                                            }
                                        , "dicc_rutina_click_item": {"rutina": "def_rutina_click_item_2"}
                                        }

        treeview.modificaciones("modificar_objeto", **kwargs_config_treeview_update)
        treeview.config_atributos(**kwargs_config_treeview_update)

        treeview.modificaciones("actualizar_desde_df", df_datos = df_datos)


def def_rutina_click_item_2():
    diccionario_datos_item_seleccionado = treeview.datos_item_seleccionado
    print(diccionario_datos_item_seleccionado)


treeview = mod_utils.treeview_propio(frame.widget_objeto, self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_treeview)
boton_1 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton_1)
boton_2 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton_2)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 12 - nuevo root tras pulsar un botón (toplevel)
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_boton = {"text": "nueva ventana"
                                        , "width": 15
                                        , "bg": "black"
                                        , "fg": "white"
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 180, "coord_y": 10}
                                        , "dicc_rutina":
                                                        {"rutina": "def_rutina_boton"}
                                    }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.boton = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton)


    def def_rutina_boton(self):

        kwargs_config_root_otra_ventana = {"dicc_config_root":
                                                        {"title": "PRUEBA OTRO ROOT"
                                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                                        , "tupla_geometry": (400, 200)
                                                        , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                        , "mantener_nueva_ventana_encima_otras": True
                                                        , "resizable": (0, 0)
                                                        }
                                            }
        
        self.otra_ventana = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_config_root_otra_ventana["dicc_config_root"])
        self.otra_ventana.config_atributos(**kwargs_config_root_otra_ventana["dicc_config_root"])

        clase_de_mi_proyecto_otra_ventana(self.otra_ventana)


class clase_de_mi_proyecto_otra_ventana:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#637EEA"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_label = {"text": "otra ventana"
                                    , "width": 10
                                    , "bg": "black"
                                    , "fg": "white"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.label = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "label", **self.kwargs_config_label)



if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (0, 0)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()




##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 200)
                                    , "resizable": (0, 0)
                                    }
                        }

kwargs_config_frame = {"width": 380
                            , "height": 180
                            , "bg": "#ACADB1"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                            }

kwargs_config_boton = {"text": "nueva ventana"
                                , "width": 15
                                , "bg": "black"
                                , "fg": "white"
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 180, "coord_y": 10}
                                , "dicc_rutina":
                                                {"rutina": "def_rutina_boton"}
                            }

modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)


def def_rutina_boton():

    kwargs_config_root_otra_ventana = {"dicc_config_root":
                                                    {"title": "PRUEBA ROOT"
                                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                                    , "tupla_geometry": (400, 200)
                                                    , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                    , "mantener_nueva_ventana_encima_otras": True
                                                    , "resizable": (0, 0)
                                                    }
                                        }
    
    otra_ventana = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "toplevel", **kwargs_config_root_otra_ventana["dicc_config_root"])
    otra_ventana.config_atributos(**kwargs_config_root_otra_ventana["dicc_config_root"])

    kwargs_config_frame = {"width": 380
                            , "height": 180
                            , "bg": "#637EEA"
                            , "bd": 2
                            , "relief": "solid"
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                            }
    
    kwargs_config_label = {"text": "otra ventana"
                            , "width": 10
                            , "bg": "black"
                            , "fg": "white"
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                            }

    frame = mod_utils.gui_tkinter_widgets(otra_ventana.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
    label = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "label", **kwargs_config_label)


boton = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_boton)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 13 - Aplicar a un widget varias rutinas de evento bind (ejemplo combobox)
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        

        self.kwargs_config_combobox = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                        , "combobox_lista_opciones": ["A", "B", "C"]
                                        , "lista_dicc_rutina_aplicar_eventos_widget": [{"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_1"}
                                                                                        , {"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_2"}
                                                                                        , {"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_3"}
                                                                                        ]
                                        }
    

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.combobox = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_combobox)

    def def_rutina_1(self):
        print("impresion desde def_rutina_1")

    def def_rutina_2(self):
        print("impresion desde def_rutina_2")

    def def_rutina_3(self):
        print("impresion desde def_rutina_3")


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (1, 1)
                                        }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }


kwargs_config_combobox = {"font": ("Calibri", 10)
                        , "width": 15
                        , "justify": tk.LEFT
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        , "combobox_lista_opciones": ["A", "B", "C"]
                        , "lista_dicc_rutina_aplicar_eventos_widget": [{"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_1"}
                                                                        , {"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_2"}
                                                                        , {"tipo_bind": "<<ComboboxSelected>>", "rutina": "def_rutina_3"}
                                                                        ]
                        }

modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

def def_rutina_1():
    print("impresion desde def_rutina_1")

def def_rutina_2():
    print("impresion desde def_rutina_2")

def def_rutina_3():
    print("impresion desde def_rutina_3")

combobox = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_combobox)

root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 14 - Aplicar a una variable enlace (stringvar) varias rutinas de evento en modo trace
#               (ejemplo lista opciones de combobox dependientes del valor escogido en otro combobox)
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.df_datos = pd.DataFrame({"COLUMNA_1": ["A", "A", "B", "B", "C", "C"]
                                        , "COLUMNA_2": ["AA", "AA", "BB", "BB", "CC", "CC"]
                                        , "COLUMNA_3": ["AAA", "AAA", "BBB", "BBB", "CCC", "CCC"]
                                        })
        
        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.kwargs_config_combobox_1 = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                        , "combobox_lista_opciones": ["A", "B", "C"]
                                        , "lista_dicc_rutina_trace_variable_enlace":[{"tipo_trace": "write", "rutina": "def_rutina_update_combobox_2"}
                                                                                        , {"tipo_trace": "write", "rutina": "def_rutina_update_combobox_3"}
                                                                                        ]
                                        }

        self.kwargs_config_combobox_2 = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 50}
                                        , "combobox_lista_opciones": []
                                        }
        
        self.kwargs_config_combobox_3 = {"font": ("Calibri", 10)
                                        , "width": 15
                                        , "justify": tk.LEFT
                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 90}
                                        , "combobox_lista_opciones": []
                                        }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.combobox_1 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_combobox_1)
        self.combobox_2 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", **self.kwargs_config_combobox_2)
        self.combobox_3 = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "combobox", **self.kwargs_config_combobox_3)


    def def_rutina_update_combobox_2(self):
        valor_combobox_1 = self.combobox_1.variable_enlace.get()
        
        df_datos_filtrado = self.df_datos.loc[self.df_datos["COLUMNA_1"] == valor_combobox_1, ["COLUMNA_2"]]
        lista_opciones = df_datos_filtrado.values.tolist()
        lista_opciones = [item[0] for item in lista_opciones]
        
        self.combobox_2.config_atributos(**{"combobox_lista_opciones": lista_opciones})


    def def_rutina_update_combobox_3(self):
        valor_combobox_1 = self.combobox_1.variable_enlace.get()
        
        df_datos_filtrado = self.df_datos.loc[self.df_datos["COLUMNA_1"] == valor_combobox_1, ["COLUMNA_3"]]
        lista_opciones = df_datos_filtrado.values.tolist()
        lista_opciones = [item[0] for item in lista_opciones]
        
        self.combobox_3.config_atributos(**{"combobox_lista_opciones": lista_opciones})


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 200)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

df_datos = pd.DataFrame({"COLUMNA_1": ["A", "A", "B", "B", "C", "C"]
                        , "COLUMNA_2": ["AA", "AA", "BB", "BB", "CC", "CC"]
                        , "COLUMNA_3": ["AAA", "AAA", "BBB", "BBB", "CCC", "CCC"]
                        })

kwargs_config_root = {"dicc_config_root":
                                {"title": "PRUEBA ROOT"
                                , "iconbitmap": ico_tapar_pluma_tkinter
                                , "tupla_geometry": (400, 200)
                                , "resizable": (1, 1)
                                }
                    }

kwargs_config_frame = {"width": 380
                        , "height": 180
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_combobox_1 = {"font": ("Calibri", 10)
                            , "width": 15
                            , "justify": tk.LEFT
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                            , "combobox_lista_opciones": ["A", "B", "C"]
                            , "lista_dicc_rutina_trace_variable_enlace":[{"tipo_trace": "write", "rutina": "def_rutina_update_combobox_2"}
                                                                            , {"tipo_trace": "write", "rutina": "def_rutina_update_combobox_3"}
                                                                            ]
                            }

kwargs_config_combobox_2 = {"font": ("Calibri", 10)
                            , "width": 15
                            , "justify": tk.LEFT
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 50}
                            , "combobox_lista_opciones": []
                            }

kwargs_config_combobox_3 = {"font": ("Calibri", 10)
                            , "width": 15
                            , "justify": tk.LEFT
                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 90}
                            , "combobox_lista_opciones": []
                            }

modulo_python_actual = sys.modules[__name__]

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)


def def_rutina_update_combobox_2():
    valor_combobox_1 = combobox_1.variable_enlace.get()

    df_datos_filtrado = df_datos.loc[df_datos["COLUMNA_1"] == valor_combobox_1, ["COLUMNA_2"]]
    lista_opciones = df_datos_filtrado.values.tolist()
    lista_opciones = [item[0] for item in lista_opciones]

    combobox_2.config_atributos(**{"combobox_lista_opciones": lista_opciones})


def def_rutina_update_combobox_3():
    valor_combobox_1 = combobox_1.variable_enlace.get()

    df_datos_filtrado = df_datos.loc[df_datos["COLUMNA_1"] == valor_combobox_1, ["COLUMNA_3"]]
    lista_opciones = df_datos_filtrado.values.tolist()
    lista_opciones = [item[0] for item in lista_opciones]

    combobox_3.config_atributos(**{"combobox_lista_opciones": lista_opciones})


combobox_1 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", self_clase_gui_donde_call_rutina = modulo_python_actual, **kwargs_config_combobox_1)
combobox_2 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", **kwargs_config_combobox_2)
combobox_3 = mod_utils.gui_tkinter_widgets(frame.widget_objeto, tipo_widget_param = "combobox", **kwargs_config_combobox_3)

root.widget_objeto.mainloop()



##########################################################################################################
##########################################################################################################
# EJEMPLO 15 - scrolledtext desde un string
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.lista_texto = [
                            "El diagnostico se realizara solo sobre la bbdd MS Access que tengas configurada para BBDD_01.\n\n"
                            , "IMPORTANTE: la BBDD MS Access debe tener el código VBA desbloqueado y no debe tener una macro AutoExec activada.\n\nAl finalizar el proceso, se generara un excel donde:\n\n"
                            , "LISTADO: se listan todos las objetos y se aportan diversas informaciones en función del tipo de objeto.\n\n"
                            , "DEPENDENCIAS: se listan para cada objeto en que modulos/rutinas VBA se usan.\n\nSIN DEPENDENCIAS: se listan que objetos no se usan modulos/rutinas VBA.\n\n"
                            , "TABLAS (CHECK MANUAL): se listan todas las tablas con las rutinas / funciones VBA en las que se usan como string encapsuladas entre comillas dobles"
                            , "pero que no parecen sentencias SQL o de maniplación de tablas via VBA. Sera el usuario quien ha de decidir si las tablas de este listado se usan o no en código VBA."
                            ]
        
        self.string_scrolledtext = "".join(self.lista_texto)
        self.df_scrolledtext = pd.DataFrame([[item] for item in self.lista_texto], columns = ["TEXTO"])

        self.kwargs_config_frame = {"width": 380
                                    , "height": 230
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.kwargs_config_scrolledtext_1 =  {"font": ("Calibri", 10, "bold")
                                                , "width": 40
                                                , "height": 5
                                                , "state": tk.NORMAL
                                                , "bg": "#B7C3F5"
                                                , "fg": "black"
                                                , "wrap": tk.WORD
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                , "justify": tk.LEFT
                                                }
        
        self.kwargs_config_scrolledtext_2 =  {"font": ("Calibri", 10, "bold")
                                                , "width": 40
                                                , "height": 5
                                                , "state": tk.NORMAL
                                                , "bg": "#B7C3F5"
                                                , "fg": "black"
                                                , "wrap": tk.NONE
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 130}
                                                , "justify": tk.LEFT
                                                , "colocacion_scrollbar_horizontal": {"metodo": "place", "coord_x": 255, "coord_y": 110}
                                                }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.scrolledtext_1 = mod_utils.scrolledtext_propio(self.frame.widget_objeto, **self.kwargs_config_scrolledtext_1)
        self.scrolledtext_2 = mod_utils.scrolledtext_propio(self.frame.widget_objeto, **self.kwargs_config_scrolledtext_2)

        self.scrolledtext_1.modificaciones("agregar_solo_contenido_desde_string"
                                        , string_texto_informar = self.string_scrolledtext
                                        , height_scrolledtext = self.kwargs_config_scrolledtext_1.get("height", 1))
        
        self.scrolledtext_2.modificaciones("agregar_solo_contenido_desde_dataframe"
                                        , df_datos = self.df_scrolledtext
                                        , columna_df_para_informar = "TEXTO"
                                        , height_scrolledtext = self.kwargs_config_scrolledtext_2.get("height", 1))


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 250)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

lista_texto = [
                "El diagnostico se realizara solo sobre la bbdd MS Access que tengas configurada para BBDD_01.\n\n"
                , "IMPORTANTE: la BBDD MS Access debe tener el código VBA desbloqueado y no debe tener una macro AutoExec activada.\n\nAl finalizar el proceso, se generara un excel donde:\n\n"
                , "LISTADO: se listan todos las objetos y se aportan diversas informaciones en función del tipo de objeto.\n\n"
                , "DEPENDENCIAS: se listan para cada objeto en que modulos/rutinas VBA se usan.\n\nSIN DEPENDENCIAS: se listan que objetos no se usan modulos/rutinas VBA.\n\n"
                , "TABLAS (CHECK MANUAL): se listan todas las tablas con las rutinas / funciones VBA en las que se usan como string encapsuladas entre comillas dobles"
                , "pero que no parecen sentencias SQL o de maniplación de tablas via VBA. Sera el usuario quien ha de decidir si las tablas de este listado se usan o no en código VBA."
                ]

string_scrolledtext = "".join(lista_texto)
df_scrolledtext = pd.DataFrame([[item] for item in lista_texto], columns = ["TEXTO"])

kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 250)
                                    , "resizable": (1, 1)
                                    }
                        }

kwargs_config_frame = {"width": 380
                        , "height": 230
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_scrolledtext_1 =  {"font": ("Calibri", 10, "bold")
                                , "width": 40
                                , "height": 5
                                , "state": tk.NORMAL
                                , "bg": "#B7C3F5"
                                , "fg": "black"
                                , "wrap": tk.WORD
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                , "justify": tk.LEFT
                                }

kwargs_config_scrolledtext_2 =  {"font": ("Calibri", 10, "bold")
                                , "width": 40
                                , "height": 5
                                , "state": tk.NORMAL
                                , "bg": "#B7C3F5"
                                , "fg": "black"
                                , "wrap": tk.NONE
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 130}
                                , "justify": tk.LEFT
                                , "colocacion_scrollbar_horizontal": {"metodo": "place", "coord_x": 255, "coord_y": 110}
                                }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)
scrolledtext_1 = mod_utils.scrolledtext_propio(frame.widget_objeto, **kwargs_config_scrolledtext_1)
scrolledtext_2 = mod_utils.scrolledtext_propio(frame.widget_objeto, **kwargs_config_scrolledtext_2)

scrolledtext_1.modificaciones("agregar_solo_contenido_desde_string"
                                , string_texto_informar = string_scrolledtext
                                , height_scrolledtext = kwargs_config_scrolledtext_1.get("height", 1))

scrolledtext_2.modificaciones("agregar_solo_contenido_desde_dataframe"
                                , df_datos = df_scrolledtext
                                , columna_df_para_informar = "TEXTO"
                                , height_scrolledtext = kwargs_config_scrolledtext_2.get("height", 1))

root.widget_objeto.mainloop()



##########################################################################################################
##########################################################################################################
# EJEMPLO 16 - scrolledtext desde un dataframe aplicando tags
##########################################################################################################
##########################################################################################################


##################################################
# clase propia
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.lista_texto_para_df = [['CREATE TABLE Production.Document', 'CREATE TABLE Production.Document', None]
                                    , ['\t(', '\t(', None], ['[DocumentNode] HIERARCHYID(892)', '[Owner] INT', 'CAMBIOS_LOCALIZADOS']
                                    , [', [DocumentLevel] SMALLINT', ', [DocumentLevel] smallint', 'CAMBIOS_LOCALIZADOS']
                                    , [', [Title] NVARCHAR(50)', ')', 'CAMBIOS_LOCALIZADOS']
                                    , [', [Owner] INT', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [FolderFlag] BIT', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [FileName] NVARCHAR(400)', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [FileExtension] NVARCHAR(8)', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [Revision] NCHAR(5)', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [ChangeNumber] INT', None, None]
                                    , [', [Status] TINYINT', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [DocumentSummary] NVARCHAR(-1)', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [Document] VARBINARY(-1)', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [rowguid] UNIQUEIDENTIFIER', None, 'CAMBIOS_LOCALIZADOS']
                                    , [', [ModifiedDate] DATETIME', None, 'CAMBIOS_LOCALIZADOS']
                                    , [')', None, None]]

        self.df_scrolledtext = pd.DataFrame(self.lista_texto_para_df, columns = ["CODIGO", "CODIGO_OTRA_BBDD", "CONTROL_CAMBIOS_ACTUAL"])

        self.kwargs_config_frame = {"width": 380
                                    , "height": 320
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }

        self.kwargs_config_scrolledtext_1 =  {"font": ("Calibri", 10, "bold")
                                                , "width": 40
                                                , "height": 5
                                                , "state": tk.NORMAL
                                                , "bg": "#B7C3F5"
                                                , "fg": "black"
                                                , "wrap": tk.NONE
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                , "justify": tk.LEFT
                                                , "lista_dicc_tag_linea_completa":
                                                                                [{"nombre_tag": "CAMBIOS_LOCALIZADOS"
                                                                                    , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                                    , "case_sensitive": False
                                                                                    , "dicc_config": {"background": "#05FB27"}
                                                                                    }
                                                                                ]
                                                }
        
        self.kwargs_config_scrolledtext_2 =  {"font": ("Calibri", 10, "bold")
                                                , "width": 40
                                                , "height": 5
                                                , "state": tk.NONE
                                                , "bg": "#B7C3F5"
                                                , "fg": "black"
                                                , "wrap": tk.NONE
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 110}
                                                , "justify": tk.LEFT
                                                , "lista_dicc_tag_caracteres_cambiantes_comparativa":
                                                                                                    [{"nombre_tag": "CAMBIOS_LOCALIZADOS_POR_INDICES"
                                                                                                        , "columna_df_filtro_registros_aplicar_tag": "CONTROL_CAMBIOS_ACTUAL"
                                                                                                        , "columna_df_filtro_registros_aplicar_tag_valor": "CAMBIOS_LOCALIZADOS"
                                                                                                        , "columna_df_comparar_1": "CODIGO"
                                                                                                        , "columna_df_comparar_2": "CODIGO_OTRA_BBDD"
                                                                                                        , "case_sensitive": True
                                                                                                        , "marcar_toda_linea_si_todo_varia": False
                                                                                                        , "dicc_config": {"foreground": "red"}
                                                                                                        }
                                                                                                    ] 
                                                }
        
        self.kwargs_config_scrolledtext_3 =  {"font": ("Calibri", 10, "bold")
                                                , "width": 40
                                                , "height": 5
                                                , "state": tk.NORMAL
                                                , "bg": "#B7C3F5"
                                                , "fg": "black"
                                                , "wrap": tk.NONE
                                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 210}
                                                , "justify": tk.LEFT
                                                , "lista_dicc_tag_linea_completa":
                                                                                [{"nombre_tag": "CAMBIOS_LOCALIZADOS"
                                                                                    , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                                    , "case_sensitive": True
                                                                                    , "dicc_config": {"background": "#05FB27"}
                                                                                    }
                                                                                ]
                                                , "lista_dicc_tag_caracteres_cambiantes_comparativa":
                                                                                                    [{"nombre_tag": "CAMBIOS_LOCALIZADOS_POR_INDICES"
                                                                                                        , "columna_df_filtro_registros_aplicar_tag": "CONTROL_CAMBIOS_ACTUAL"
                                                                                                        , "columna_df_filtro_registros_aplicar_tag_valor": "CAMBIOS_LOCALIZADOS"
                                                                                                        , "columna_df_comparar_1": "CODIGO"
                                                                                                        , "columna_df_comparar_2": "CODIGO_OTRA_BBDD"
                                                                                                        , "case_sensitive": True
                                                                                                        , "marcar_toda_linea_si_todo_varia": False
                                                                                                        , "dicc_config": {"foreground": "red"}
                                                                                                        }
                                                                                                    ] 
                                                }

        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)

        self.scrolledtext_1 = mod_utils.scrolledtext_propio(self.frame.widget_objeto, **self.kwargs_config_scrolledtext_1)
        self.scrolledtext_2 = mod_utils.scrolledtext_propio(self.frame.widget_objeto, **self.kwargs_config_scrolledtext_2)
        self.scrolledtext_3 = mod_utils.scrolledtext_propio(self.frame.widget_objeto, **self.kwargs_config_scrolledtext_3)

        self.scrolledtext_1.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                        , df_datos = self.df_scrolledtext
                                        , columna_df_para_informar = "CODIGO"
                                        , height_scrolledtext = self.kwargs_config_scrolledtext_1.get("height", 1)
                                        , lista_dicc_tag_linea_completa = self.kwargs_config_scrolledtext_1.get("lista_dicc_tag_linea_completa", None)
                                        )

        self.scrolledtext_2.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                        , df_datos = self.df_scrolledtext
                                        , columna_df_para_informar = "CODIGO"
                                        , height_scrolledtext = self.kwargs_config_scrolledtext_2.get("height", 1)
                                        , lista_dicc_tag_caracteres_cambiantes_comparativa = self.kwargs_config_scrolledtext_2.get("lista_dicc_tag_caracteres_cambiantes_comparativa", None)
                                        )
        
        self.scrolledtext_3.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                        , df_datos = self.df_scrolledtext
                                        , columna_df_para_informar = "CODIGO"
                                        , height_scrolledtext = self.kwargs_config_scrolledtext_3.get("height", 1)
                                        , lista_dicc_tag_linea_completa = self.kwargs_config_scrolledtext_3.get("lista_dicc_tag_linea_completa", None)
                                        , lista_dicc_tag_caracteres_cambiantes_comparativa = self.kwargs_config_scrolledtext_3.get("lista_dicc_tag_caracteres_cambiantes_comparativa", None)
                                        )


if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 340)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()


##################################################
# directamente en el modulo
##################################################

ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))



lista_texto_para_df = [['CREATE TABLE Production.Document', 'CREATE TABLE Production.Document', None]
                        , ['\t(', '\t(', None], ['[DocumentNode] HIERARCHYID(892)', '[Owner] INT', 'CAMBIOS_LOCALIZADOS']
                        , [', [DocumentLevel] SMALLINT', ', [DocumentLevel] smallint', 'CAMBIOS_LOCALIZADOS']
                        , [', [Title] NVARCHAR(50)', ')', 'CAMBIOS_LOCALIZADOS']
                        , [', [Owner] INT', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [FolderFlag] BIT', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [FileName] NVARCHAR(400)', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [FileExtension] NVARCHAR(8)', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [Revision] NCHAR(5)', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [ChangeNumber] INT', None, None]
                        , [', [Status] TINYINT', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [DocumentSummary] NVARCHAR(-1)', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [Document] VARBINARY(-1)', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [rowguid] UNIQUEIDENTIFIER', None, 'CAMBIOS_LOCALIZADOS']
                        , [', [ModifiedDate] DATETIME', None, 'CAMBIOS_LOCALIZADOS']
                        , [')', None, None]]

df_scrolledtext = pd.DataFrame(lista_texto_para_df, columns = ["CODIGO", "CODIGO_OTRA_BBDD", "CONTROL_CAMBIOS_ACTUAL"])


kwargs_config_root = {"dicc_config_root":
                                    {"title": "PRUEBA ROOT"
                                    , "iconbitmap": ico_tapar_pluma_tkinter
                                    , "tupla_geometry": (400, 340)
                                    , "resizable": (1, 1)
                                    }
                        }

kwargs_config_frame = {"width": 380
                        , "height": 320
                        , "bg": "#ACADB1"
                        , "bd": 2
                        , "relief": "solid"
                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                        }

kwargs_config_scrolledtext_1 =  {"font": ("Calibri", 10, "bold")
                                , "width": 40
                                , "height": 5
                                , "state": tk.NORMAL
                                , "bg": "#B7C3F5"
                                , "fg": "black"
                                , "wrap": tk.NONE
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                , "justify": tk.LEFT
                                , "lista_dicc_tag_linea_completa":
                                                                [{"nombre_tag": "CAMBIOS_LOCALIZADOS"
                                                                    , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                    , "case_sensitive": False
                                                                    , "dicc_config": {"background": "#05FB27"}
                                                                    }
                                                                ]
                                }
        
kwargs_config_scrolledtext_2 =  {"font": ("Calibri", 10, "bold")
                                , "width": 40
                                , "height": 5
                                , "state": tk.NONE
                                , "bg": "#B7C3F5"
                                , "fg": "black"
                                , "wrap": tk.NONE
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 110}
                                , "justify": tk.LEFT
                                , "lista_dicc_tag_caracteres_cambiantes_comparativa":
                                                                                    [{"nombre_tag": "CAMBIOS_LOCALIZADOS_POR_INDICES"
                                                                                        , "columna_df_filtro_registros_aplicar_tag": "CONTROL_CAMBIOS_ACTUAL"
                                                                                        , "columna_df_filtro_registros_aplicar_tag_valor": "CAMBIOS_LOCALIZADOS"
                                                                                        , "columna_df_comparar_1": "CODIGO"
                                                                                        , "columna_df_comparar_2": "CODIGO_OTRA_BBDD"
                                                                                        , "case_sensitive": True
                                                                                        , "marcar_toda_linea_si_todo_varia": False
                                                                                        , "dicc_config": {"foreground": "red"}
                                                                                        }
                                                                                    ] 
                                }
        
kwargs_config_scrolledtext_3 =  {"font": ("Calibri", 10, "bold")
                                , "width": 40
                                , "height": 5
                                , "state": tk.NORMAL
                                , "bg": "#B7C3F5"
                                , "fg": "black"
                                , "wrap": tk.NONE
                                , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 210}
                                , "justify": tk.LEFT
                                , "lista_dicc_tag_linea_completa":
                                                                [{"nombre_tag": "CAMBIOS_LOCALIZADOS"
                                                                    , "columna_df_tag_aplicar": "CONTROL_CAMBIOS_ACTUAL"
                                                                    , "case_sensitive": True
                                                                    , "dicc_config": {"background": "#05FB27"}
                                                                    }
                                                                ]
                                , "lista_dicc_tag_caracteres_cambiantes_comparativa":
                                                                                    [{"nombre_tag": "CAMBIOS_LOCALIZADOS_POR_INDICES"
                                                                                        , "columna_df_filtro_registros_aplicar_tag": "CONTROL_CAMBIOS_ACTUAL"
                                                                                        , "columna_df_filtro_registros_aplicar_tag_valor": "CAMBIOS_LOCALIZADOS"
                                                                                        , "columna_df_comparar_1": "CODIGO"
                                                                                        , "columna_df_comparar_2": "CODIGO_OTRA_BBDD"
                                                                                        , "case_sensitive": True
                                                                                        , "marcar_toda_linea_si_todo_varia": False
                                                                                        , "dicc_config": {"foreground": "red"}
                                                                                        }
                                                                                    ] 
                                }

root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
frame = mod_utils.gui_tkinter_widgets(root.widget_objeto, tipo_widget_param = "frame", **kwargs_config_frame)

scrolledtext_1 = mod_utils.scrolledtext_propio(frame.widget_objeto, **kwargs_config_scrolledtext_1)
scrolledtext_2 = mod_utils.scrolledtext_propio(frame.widget_objeto, **kwargs_config_scrolledtext_2)
scrolledtext_3 = mod_utils.scrolledtext_propio(frame.widget_objeto, **kwargs_config_scrolledtext_3)

scrolledtext_1.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                , df_datos = df_scrolledtext
                                , columna_df_para_informar = "CODIGO"
                                , height_scrolledtext = kwargs_config_scrolledtext_1.get("height", 1)
                                , lista_dicc_tag_linea_completa = kwargs_config_scrolledtext_1.get("lista_dicc_tag_linea_completa", None)
                                )

scrolledtext_2.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                , df_datos = df_scrolledtext
                                , columna_df_para_informar = "CODIGO"
                                , height_scrolledtext = kwargs_config_scrolledtext_2.get("height", 1)
                                , lista_dicc_tag_caracteres_cambiantes_comparativa = kwargs_config_scrolledtext_2.get("lista_dicc_tag_caracteres_cambiantes_comparativa", None)
                                )

scrolledtext_3.modificaciones("agregar_contenido_y_tags_desde_dataframe"
                                , df_datos = df_scrolledtext
                                , columna_df_para_informar = "CODIGO"
                                , height_scrolledtext = kwargs_config_scrolledtext_3.get("height", 1)
                                , lista_dicc_tag_linea_completa = kwargs_config_scrolledtext_3.get("lista_dicc_tag_linea_completa", None)
                                , lista_dicc_tag_caracteres_cambiantes_comparativa = kwargs_config_scrolledtext_3.get("lista_dicc_tag_caracteres_cambiantes_comparativa", None)
                                )


root.widget_objeto.mainloop()




##########################################################################################################
##########################################################################################################
# EJEMPLO 17 - Configuración de atributos nativos tkinter no integrados en mi_sistema_tkinter
#              (messagebox, filedialog etc) dentro de clases propias en mi_proyecto
##########################################################################################################
##########################################################################################################


ico_tapar_pluma_tkinter = (os.path.join(sys._MEIPASS, "ico_tapar_pluma_tkinter.ico")
                           if getattr(sys, 'frozen', False)
                           else os.path.join(pathlib.Path(__file__).parent.absolute(), "ico_tapar_pluma_tkinter.ico"))

class clase_de_mi_proyecto:

    def __init__(self, master):

        self.master = master

        self.kwargs_config_frame = {"width": 380
                                    , "height": 180
                                    , "bg": "#ACADB1"
                                    , "bd": 2
                                    , "relief": "solid"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    }
        
        self.kwargs_config_boton =  {"text": "botón"
                                    , "width": 5
                                    , "bg": "black"
                                    ,"fg": "white"
                                    , "controltiptext": "esto botón solo imprime"
                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                    , "dicc_rutina":
                                                    {"rutina": "def_rutina_boton"}
                                    }
        
        self.frame = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "frame", **self.kwargs_config_frame)
        self.boton = mod_utils.gui_tkinter_widgets(self.frame.widget_objeto, tipo_widget_param = "button", self_clase_gui_donde_call_rutina = self, **self.kwargs_config_boton)

    
    def def_rutina_boton(self):

        filedialog = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres que se guarde la guia de usuario:")

        self.master.widget_objeto.config(cursor = "wait")

        self.master.widget_objeto.config(cursor = "")

if __name__ == "__main__":

    kwargs_config_root = {"dicc_config_root":
                                        {"title": "PRUEBA ROOT"
                                        , "iconbitmap": ico_tapar_pluma_tkinter
                                        , "tupla_geometry": (400, 340)
                                        , "resizable": (1, 1)
                                        }
                            }

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **kwargs_config_root)
    clase_de_mi_proyecto(root)
    root.widget_objeto.mainloop()



