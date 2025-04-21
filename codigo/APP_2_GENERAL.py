import pandas as pd
import pyodbc
import os
import sys
import re
import pathlib
import traceback
import datetime as dt
import xlwings as xw
import difflib
from threading import Thread
import win32com.client
import subprocess
import time

import APP_3_BACK_END_MS_ACCESS as mod_access
import APP_3_BACK_END_SQL_SERVER as mod_sql_server



#los templates son los ficheros ademas del codigo Python que se usan dentro del app
#si se ejecuta desde consola o se ejecuta desde el ejecutable compilado (con estos templates incluidos dentro)
#el accesso a los path se codifica de forma distinta en Python
#aqui se prueba 1ero si la ejecucion es desde el ejecutable y sino se hace desde consola

try:
    #si es un ejecutable
    ico_app = sys._MEIPASS + r"\ico_app.ico" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() + r"\ico_app.ico"
    ruta_plantilla_control_versiones_xls = sys._MEIPASS + r"\PLANTILLA_CONTROL_VERSIONES.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() + r"\PLANTILLA_CONTROL_VERSIONES.xlsb"
    ruta_plantilla_diagnostico_access_xls = sys._MEIPASS + r"\PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() + r"\PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb"
    ruta_plantilla_diagnostico_sql_server_xls = sys._MEIPASS + r"\PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() + r"\PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb"

except:
    #si se ejecuta desde consola
    ico_app = sys._MEIPASS / "ico_app.ico" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() / "ico_app.ico"
    ruta_plantilla_control_versiones_xls = sys._MEIPASS / "PLANTILLA_CONTROL_VERSIONES.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() / "PLANTILLA_CONTROL_VERSIONES.xlsb"
    ruta_plantilla_diagnostico_access_xls = sys._MEIPASS / "PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() / "PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb"
    ruta_plantilla_diagnostico_sql_server_xls = sys._MEIPASS / "PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb" if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent.absolute() / "PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb"
    
    pass



#################################################################################################################################################################################
##                     VARIABLES GENERALES
#################################################################################################################################################################################

#############################################
# varios
#############################################

#es el nombre del app para que figura en el titulo de la GUI de inicio y como titulo en todos los messagebox del app
nombre_app = "APP CONTROL VERSIONES EN MS ACCESS & SQL SERVER"


#es el titulo que figura en la GUI de control de versiones
nombre_control_versiones = "CONTROL VERSIONES - MS ACCESS & SQL SERVER"


#es el titulo que figura en la GUI de merge en bbdd fisica
nombre_merge_bbdd_fisicas = "MERGE EN BBDD FISICAS - MS ACCESS & SQL SERVER"


#es una variable global que impide cuando se ejecute un proceso lanzar otro hasta que el proceso en curso acabe
global_proceso_en_ejecucion = "NO"


#listas que sirven para crear las opciones del combobox de "Acció" en la GUI de control versiones en el cuadro de MERGE
#--> lista_GUI_proceso_merge_tipo_accion_bbdd_1 es cuando se seleciona BBDD_01 en el combobox BBDD
#--> lista_GUI_proceso_merge_tipo_accion_bbdd_2 es cuando se seleciona BBDD_02 en el combobox BBDD
#IMPORTANTE: no cambiar el orden de las opciones sino habra que ir al codigo donde se usan estas 2 listas y cambiar los indices
lista_GUI_proceso_merge_tipo_accion_bbdd_1 = ["Migrar todo", "Migrar por lineas", "Revertir todo", "Revertir Último"]
lista_GUI_proceso_merge_tipo_accion_bbdd_2 = ["Quitar todo", "Quitar por lineas", "Revertir todo", "Revertir Último"]

lista_acciones_migrar_quitar = ["Migrar todo", "Migrar por lineas", "Quitar todo", "Quitar por lineas"]
lista_acciones_revertir = ["Revertir todo", "Revertir Último"]


#es el literal que sale en el sub-form de la GUI de merge en bbdd fisica (columna ESTADO MIGRACION) cuando se opta por MS_ACCESS y ajustes manuales
label_merge_access_bbdd_fisica_en_manual = "Migración manual"


#listas headers de los df de codigo que se calculan en el control de versiones
lista_headers_df_codigo_control_versiones_1 = ["NUM_LINEA", "CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]
lista_headers_df_codigo_control_versiones_2 = ["NUM_LINEA", "CODIGO", "CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]


#############################################
# diccionario de procesos
#############################################

#es el diccionario de los distintos procesos que se calculan en el app
#IMPORTANTE: las keys principales (PROCESO_01, PROCESO_02 y PROCESO_03) NO DEBEN CAMBIARSE pq se reutilizan en el codigo del app en multiples ocasiones
#IMPORTANTE: las keys de los diccionarios asociados a las keys principales (PROCESO y COMENTARIO) tampoco se han de cambiar por el mismo motivo (aunque solo se usan en la GUI)
#
#los valores asociados a las keys de estos diccionarios si son modificables (y no afectan el funcionamiento del app)

dicc_procesos = {"PROCESO_01":
                        {"PROCESO": "Control de versiones"
                         
                        , "COMENTARIO": [
                                    "El control de versiones puede hacerse según 3 opciones:\n\n"
                                    , "1. Solamente entre 2 bbdd MS Access,\n"
                                    , "2. Solamente entre 2 bbdd SQL Server.\n"
                                    , "3. Sobre 2 bbdd MS Access y sobre 2 bbdd SQL Server.\n\n"
                                    , "MS ACCESS:\n\n"
                                    , "   --> las 2 bbdd deben tener el código VBA desbloqueado.\n"
                                    , "   --> las 2 bbdd deben tener la macro AutoExec desactivada.\n"
                                    , "   --> no se ejecutara el proceso hasta tener 2 rutas de bbdd distintas configuradas.\n\n"
                                    , "SQL SERVER:\n\n"
                                    , "   --> las bbdd que se relacionan a un servidor son aquellas donde la(el) usuari@ tiene permiso\n"
                                    , "       de acceso al código (permiso de 'VIEW DEFINITION' en fn_my_permissions(NULL, 'DATABASE')).\n"
                                    , "   --> no se ejecutara el proceso si el mismo Servidor + bbdd se configura 2 veces.\n\n"
                                    ]
                        }
    
                    , "PROCESO_02":
                        {"PROCESO": "Diagnostico BBDD Access"
                         
                        , "COMENTARIO": [
                                    "El diagnostico se realizara solo sobre la bbdd MS Access que tengas configurada para BBDD_01.\n\n"
                                    , "IMPORTANTE: la BBDD MS Access debe tener el código VBA desbloqueado y no debe tener una macro AutoExec activada.\n\n"
                                    , "Al finalizar el proceso, se generara un excel donde:\n\n"
                                    , "LISTADO: se listan todos las objetos y se aportan diversas informaciones en función del tipo de objeto.\n\n"
                                    , "DEPENDENCIAS: se listan para cada objeto en que modulos/rutinas VBA se usan.\n\n"
                                    , "SIN DEPENDENCIAS: se listan que objetos no se usan modulos/rutinas VBA.\n\n"
                                    , "TABLAS (CHECK MANUAL): se listan todas las tablas con las rutinas / funciones VBA en las que se usan como string encapsuladas entre comillas dobles"
                                    , "pero que no parecen sentencias SQL o de maniplación de tablas via VBA. Sera el usuario quien ha de decidir si las tablas de este listado se usan o no en código VBA."
                                    ]
                        }

                    , "PROCESO_03":
                        {"PROCESO": "Diagnostico Servidor SQL Server"
                         
                        , "COMENTARIO": [
                                    "El diagnostico se realizara solo sobre el servidor que tengas configurado para BBDD_01.\n\n"
                                    , "IMPORTANTE:Se abrira una nueva ventana donde se listaran solo las bbdd donde tengas permiso de escriacceso a los códigos de objetos"
                                    , "y tienes la opción de poder escojer una sola bbdd o varias para ver las dependencias.\n\n"
                                    , "Al finalizar el proceso, se generara un excel con las hojas siguientes:\n\n"
                                    , "BBDD SELECCIONADAS: lista el servidor y bases de datos seleccionadas.\n\n"
                                    , "LISTADO: lista los objetos de las bases de datos selecionadas informaabdo del esquema, el tipo de objeto y de sus parametros si los tiene.\n\n"
                                    , "DEPENDENCIAS: para cada objeto de la hoja LISTADO se informa en que otros objetos se usan.\n\n"
                                    , "SIN DEPENDENCIAS: son los objetos de la hoja LISTADO que no dependen de ningún otro objeto.\n\n"
                                    ]
                        }
                    }


#############################################
# diccionario relacionado con la importacion de los codigos
#############################################

#es el diccionario donde se almacena todo lo relacionado con la importacion de codigo VBA y/o codigo T-SQL
#IMPORTANTE: ninguna key y subkey (1 y 2) deben cambiarse pq se usaqn en multiples sitios en el codigo del app
dicc_codigos_bbdd = {"BBDD_01":
                            {"MS_ACCESS":
                                    {
                                    #sirve para almacenar la ruta del access de BBDD_01
                                    "PATH_BBDD": None
                                     
                                    #sirve para almacenar una lista que contine las distintas librerias DLL que la BBDD_01 tiene activadas en VBA
                                    , "LISTA_LIBRERIAS_DLL": None

                                    #lista de diccionarios donde cada diccionario es comun a todos los tipos de objetos MS ACCESS (algunos items aplican o no 
                                    #segun el tipo de objeto, a cada objeto se le incorpora un df con su codigo en funcion del proceso seleccionado
                                    # --> CONTROL VERSIONES: se incorpora el codigo (en formato df)
                                    # --> DIAGNOSTICO: no se incorpora el codigo salvo para los vinculos ODBC u otros
                                    #
                                    #las variables publicas aqui se listan 1 a una
                                    , "LISTA_DICC_OBJETOS": None

                                    #df que contiene los datos necesarios para realizar el calculo del control de versiones o calculo de dependencias
                                    #tras el proceso de import en el cual se realizan calculos intermedios para dejar preparados los datos antes de ejecutar cada proceso
                                    , "DF_CODIGO_CALCULADO_TRAS_IMPORT": None
                                
                                    }
                                
                            , "SQL_SERVER":
                                    {
                                    #son el servidor + bbdd de BBDD_01
                                    "SERVIDOR": None
                                    , "BBDD": None

                                    #es la connecting string al SERVIDOR tras probar por Windows Authentication y en caso de fallo por SQL Server Authentication
                                    #(es decir por login y password)
                                    , "CONNECTING_STRING": None


                                    #lista de diccionarios donde cada diccionario contiene la info del objeto y su codigo T-SQL
                                    #es comun para el control de versiones y el diagnostico de dependencias
                                    #para mas detalle ir a la rutina def_proceso_sql_server_1_import (modulo APP_3_BACK_END_SQL_SERVER)
                                    , "LISTA_DICC_OBJETOS": None


                                    #lista de bbdd que conforman el servidor (aplica solo para el diagnostico)
                                    , "LISTA_BBDD_SERVIDOR": None

                                    }
                            }
                    
                    , "BBDD_02":
                            {"MS_ACCESS":
                                    {
                                    #sirve para almacenar la ruta del access de BBDD_02
                                    "PATH_BBDD": None
                                     
                                    #sirve para almacenar una lista que contine las distintas librerias DLL que la BBDD_01 tiene activadas en VBA
                                    , "LISTA_LIBRERIAS_DLL": None


                                    #lista de diccionarios donde cada diccionario es comun a todos los tipos de objetos MS ACCESS (algunos items aplican o no 
                                    #segun el tipo de objeto, a cada objetos se le incorpora un df con su codigo en funcion del proceso seleccionado
                                    # --> CONTROL VERSIONES: se incorpora el codigo (en formato df)
                                    # --> DIAGNOSTICO: no se incorpora el codigo salvo para los vinculos ODBC u otros
                                    #
                                    #las variables publicas aqui se listan 1 a una
                                    , "LISTA_DICC_OBJETOS": None


                                    #df que contiene los datos necesarios para realizar el calculo del control de versiones o calculo de dependencias
                                    #tras el proceso de import en el cual se realizan calculos intermedios para dejar preparados los datos antes de ejecutar cada proceso
                                    , "DF_CODIGO_CALCULADO_TRAS_IMPORT": None

                                    }
                                
                            , "SQL_SERVER":
                                    {
                                    #son el servidor + bbdd de BBDD_02
                                    "SERVIDOR": None
                                    , "BBDD": None


                                    #es la connecting string al SERVIDOR tras probar por Windows Authentication y en caso de fallo por SQL Server Authentication
                                    #(es decir por login y password)
                                    , "CONNECTING_STRING": None

                                    
                                    #lista de diccionarios donde cada diccionario contiene la info del objeto y su codigo T-SQL
                                    #es comun para el control de versiones y el diagnostico de dependencias
                                    #para mas detalle ir a la rutina def_proceso_sql_server_1_import (modulo APP_3_BACK_END_SQL_SERVER)
                                    , "LISTA_DICC_OBJETOS": None

                                    #lista de bbdd que conforman el servidor (aplica solo para el diagnostico por lo que aqui para BBDD_02 siempre sera None)
                                    , "LISTA_BBDD_SERVIDOR": None                                    
                                    }
                            }
                    }




#############################################
# diccionarios relacionados con el proceso de control de versiones
#############################################

#es el diccionario que sirve para almacenar todo lo relacionado con el control de versiones y merge de una bbdd a otra realizados por el usuario
#
#IMPORTANTE: no deben cambiarse los nommbres de las keys principales ni de las subkeys (1, 2 y 3) pq se usan en multiples sitios en el codigo del app
#(tan solo es modificable el valor de las subkeys 4 que tengan COMBOBOX en su nombre o que se llamen TIPO_OBJETO_SUBFORM)
#
#la subkey 4 llamada LISTA_DICC_OBJETOS_CONTROL_VERSIONES (asociada a la subkey 2 --> TIPO_OBJETO) es donde se almacena para el tipo de objeto asociado (subkey 3) 
#la lista de objetos con cambios de una bbdd a otra resultante del calculo de control de versiones tras el proceso de import de las 2 bbdd sean MS Access o SQL Server
#
#la subkey 4 llamada LISTA_DICC_OBJETOS (asociada a la subkey 2 --> MERGE_BBDD_FISICA) es donde se almacena la lista de diccionarios con todos los objetos donde el usuario
#ha realizado cambios de scripts de una bbdd a otra sean MS Access o SQL Server, y que sirve de base para poder realizar el MERGE en bbdd fisica
#
#las subkeys 4 llamadas LISTA_DICC_OK_MIGRACION y LISTA_DICC_ERRORES_MIGRACION (asociada a la subkey 2 --> MERGE_BBDD_FISICA) es donde se almacena la lista de diccionarios
#con OK y ERRORES para cada objeto cuando se ejecuta el MERGE en bbdd fisica (sirven para poder generar los logs en formato .txt)
#
#la subkey 4 llamada LISTA_OBJECT_ID (asociada a la subkey 2 --> MERGE_BBDD_FISICA), que solo afecta a la key principal SQL_SERVER, contiene una lista de listas que sirven
#cuando se realiza el MERGE en bbdd fisica para eliminar los objetos que entran en el MERGE si previamente ya existien en la bbdd fisica
#(se usa en la query query_drop_if_exists_objeto del modulo APP_3_BACK_END_SQL_SERVER)


dicc_control_versiones_tipo_objeto = {"MS_ACCESS":
                                                {"TIPO_OBJETO":
                                                            {"TODOS":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: <TODOS>"
                                                                            , "TIPO_OBJETO_SUBFORM": None
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            , "TABLA_LOCAL":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: tablas locales"
                                                                            , "TIPO_OBJETO_SUBFORM": "Tablas Locales"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            , "VINCULO_ODBC":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: vinculos ODBC"
                                                                            , "TIPO_OBJETO_SUBFORM": "Vinculos ODBC"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            , "VINCULO_OTRO":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: vinculos Otro"
                                                                            , "TIPO_OBJETO_SUBFORM": "Vinculos Otros"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            , "RUTINAS_VBA":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: rutinas VBA"
                                                                            , "TIPO_OBJETO_SUBFORM": "Rutinas VBA"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            , "VARIABLES_VBA":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "MS ACCESS: variables VBA"
                                                                            , "TIPO_OBJETO_SUBFORM": "Variables VBA"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }
                                                            }
                                                
                                                , "MERGE_BBDD_FISICA":
                                                                    {"AJUSTES_MANUALES":
                                                                                        {"COMBOBOX_GUI": "MS Access: ajustes manuales"
                                                                                        , "LISTA_DICC_OBJETOS": None
                                                                                        }

                                                                    , "OBJETOS_A_MIGRAR":
                                                                                        {"COMBOBOX_GUI": "MS Access: objetos a migrar"
                                                                                        , "LISTA_DICC_OBJETOS": None
                                                                                        }
                                                                    , "LISTA_DICC_OK_MIGRACION": None
                                                                    , "LISTA_DICC_ERRORES_MIGRACION": None
                                                                    }

                                                , "OBJETO_SELECCIONADO_TRAS_CLICK_SUBFORM":
                                                                            {"TIPO_BBDD": None
                                                                            , "TIPO_OBJETO": None
                                                                            , "REPOSITORIO": None
                                                                            , "NOMBRE_OBJETO": None
                                                                            , "DF_CODIGO_ACTUAL_1": None
                                                                            , "DF_CODIGO_ACTUAL_2": None
                                                                            }
                                                
                                                }

                                    , "SQL_SERVER":
                                                {"TIPO_OBJETO":
                                                            {"TODOS":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "SQL SERVER: <TODOS>"
                                                                            , "TIPO_OBJETO_SUBFORM": None
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            }  
                                                            , "TABLAS":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "SQL SERVER: tablas"
                                                                            , "TIPO_OBJETO_SUBFORM": "Tablas"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            , "LISTA_OBJECT_ID": [["TABLE", "U"]]
                                                                            }
                                                            , "VIEWS":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "SQL SERVER: views"
                                                                            , "TIPO_OBJETO_SUBFORM": "Views"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            , "LISTA_OBJECT_ID": [["VIEW", "V"]]
                                                                            }
                                                            , "STORED_PROCEDURES":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "SQL SERVER: stored procedures"
                                                                            , "TIPO_OBJETO_SUBFORM": "Stored Procedures"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            , "LISTA_OBJECT_ID": [["PROCEDURE", "P"]]
                                                                            }
                                                            , "FUNCIONES":
                                                                            {"COMBOBOX_CONTROL_VERSIONES": "SQL SERVER: funciones"
                                                                            , "TIPO_OBJETO_SUBFORM": "Funciones"
                                                                            , "LISTA_DICC_OBJETOS_CONTROL_VERSIONES": None
                                                                            , "LISTA_OBJECT_ID": [["FUNCTION", "FN"], ["FUNCTION", "TF"], ["FUNCTION", "IF"]]
                                                                            }
                                                            }


                                                , "MERGE_BBDD_FISICA":
                                                                            {"OBJETOS_A_MIGRAR":
                                                                                                {"COMBOBOX_GUI": "SQL Server: objetos a migrar"
                                                                                                , "LISTA_DICC_OBJETOS": None
                                                                                                }
                                                                            , "LISTA_DICC_OK_MIGRACION": None
                                                                            , "LISTA_DICC_ERRORES_MIGRACION": None
                                                                            }
                                                }
                                        }



#es el diccionario que sirve en la GUI de control de versiones para las opciones del combobox de tipo de concepto
#IMPORTANTE: no hay que modificar las keys (YA_EXISTE, SOLO_EN_BBDD_01 y SOLO_EN_BBDD_02) pq se usa en multiples sitios el codigo del app

dicc_control_versiones_tipo_concepto = {"YA_EXISTE": "Objetos ya existentes en las 2 BBDD's"
                                        ,"SOLO_EN_BBDD_01": "En BBDD_01 pero no en BBDD_02"
                                        ,"SOLO_EN_BBDD_02": "En BBDD_02 pero no en BBDD_01"}




#############################################
# listas para las opciones de los combobox de la GUI del control de versiones
# --> tipo de selección
# --> tipo concepto
# --> seleccion bbdd
#############################################

lista_GUI_seleccion_tipo_objeto_access = []
for key in dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"].keys():
    lista_GUI_seleccion_tipo_objeto_access.append(dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"][key]["COMBOBOX_CONTROL_VERSIONES"])

                                         
lista_GUI_seleccion_tipo_objeto_sql_server = []                                              
for key in dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"].keys():
    lista_GUI_seleccion_tipo_objeto_sql_server.append(dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"][key]["COMBOBOX_CONTROL_VERSIONES"])


lista_GUI_seleccion_tipo_concepto = [dicc_control_versiones_tipo_concepto[key] for key in dicc_control_versiones_tipo_concepto.keys()]

lista_GUI_seleccion_bbdd = [key for key in dicc_codigos_bbdd.keys()]


#############################################
# diccionario para almacenar los errores en los procesos de calculo de los procesos que se pueden ejecutar con el app
#############################################

#IMPORTANTE: no modificar el nombre de las keys 1, 2 y 3 pq se usan en multiples sitios en el codigo del app

dicc_errores_procesos = {"PROCESO_01":
                        #control versiones
                                    {"MS_ACCESS":
                                                {
                                                #lista de diccionarios con los posibles errores resultantes de la importacion del codigo de BBDD_01 y BBDD_02
                                                "LISTA_DICC_ERRORES_IMPORTACION_BBDD_01": None
                                                , "LISTA_DICC_ERRORES_IMPORTACION_BBDD_02": None


                                                #lista de diccionarios con los posibles errores resultantes del calculo del control de versiones tras realizar
                                                #la importacion de los codigos de BBDD_01 y BBDD_02
                                                , "LISTA_DICC_ERRORES_CALCULO": None
                                                }

                                    , "SQL_SERVER":
                                                {
                                                #lista de diccionarios con los posibles errores resultantes de la importacion del codigo de BBDD_01 y BBDD_02
                                                "LISTA_DICC_ERRORES_IMPORTACION_BBDD_01": None
                                                , "LISTA_DICC_ERRORES_IMPORTACION_BBDD_02": None


                                                #lista de diccionarios con los posibles errores resultantes del calculo del control de versiones tras realizar
                                                #la importacion de los codigos de BBDD_01 y BBDD_02
                                                , "LISTA_DICC_ERRORES_CALCULO": None
                                                }
                                    }

                        , "PROCESO_02":
                        #diagnostico access (solo afecta a BBDD_01)
                                    {"MS_ACCESS":
                                                {#lista de diccionarios con los posibles errores resultantes de la importacion del codigo de BBDD_01
                                                "LISTA_DICC_ERRORES_IMPORTACION_BBDD_01": None

                                                #lista de diccionarios con los posibles errores resultantes del calculo del control de versiones tras realizar
                                                #la importacion de los codigos de BBDD_01
                                                , "LISTA_DICC_ERRORES_CALCULO": None
                                                }
                                    }

                        , "PROCESO_03":
                        #diagnostico sql server (solo afecta a BBDD_01)
                                    {"SQL_SERVER":
                                                {#lista de diccionarios con los posibles errores resultantes de la importacion del codigo de BBDD_01
                                                "LISTA_DICC_ERRORES_IMPORTACION_BBDD_01": None

                                                #lista de diccionarios con los posibles errores resultantes del calculo del control de versiones tras realizar
                                                #la importacion de los codigos de BBDD_01
                                                , "LISTA_DICC_ERRORES_CALCULO": None
                                                }
                                    }
                        }


########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################
#                            RUTINAS + FUNCIONES - ERRORES y LOGS
########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################


def func_procesos_errores(proceso_id, tipo_bbdd, tipo_rutina):
    #funcion que localiza por procesos si se han generado errores duante la ejecucion

    if proceso_id == "PROCESO_01":
        #control versiones (access o sql server)

        if tipo_rutina == "IMPORT":
            error_bbdd_1 = "SI" if isinstance(dicc_errores_procesos[proceso_id][tipo_bbdd]["LISTA_DICC_ERRORES_IMPORTACION_BBDD_01"], list) else "NO"
            error_bbdd_2 = "SI" if isinstance(dicc_errores_procesos[proceso_id][tipo_bbdd]["LISTA_DICC_ERRORES_IMPORTACION_BBDD_02"], list) else "NO"

            temp = "ERROR" if error_bbdd_1 + error_bbdd_2 != "NONO" else "SIN_ERROR"

        elif tipo_rutina == "CALCULO":
            temp = "ERROR" if isinstance(dicc_errores_procesos[proceso_id][tipo_bbdd]["LISTA_DICC_ERRORES_CALCULO"], list) else "SIN_ERROR"


    elif proceso_id in ["PROCESO_02", "PROCESO_03"]:
        #diagnostico access o sql server

        if tipo_rutina == "IMPORT":
            temp = "ERROR" if isinstance(dicc_errores_procesos[proceso_id][tipo_bbdd]["LISTA_DICC_ERRORES_IMPORTACION_BBDD_01"], list)  else "SIN_ERROR"

        elif tipo_rutina == "CALCULO":
            temp = "ERROR" if isinstance(dicc_errores_procesos[proceso_id][tipo_bbdd]["LISTA_DICC_ERRORES_CALCULO"], list) else "SIN_ERROR"

    #resultado de la funcion
    return temp



def def_generacion_logs(opcion, tipo_bbdd_selecc, ruta_destino_fichero_logs, **kwargs):
    #rutina que permite generar el fichero de errores de merge en bbdd fisica
    #el proceso se ejecuta solo si dicc_errores_procesos[proceso_id][tipo_bbdd_selecc]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd] es una lista

    #parametros kwargs
    opcion_bbdd = kwargs.get("opcion_bbdd", None)
    proceso_id = kwargs.get("proceso_id", None)


    if opcion == "PROCESOS_LOGS_ERRORES_IMPORT_BBDD":

        if tipo_bbdd_selecc == "MS_ACCESS":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["PATH_BBDD"]

        elif tipo_bbdd_selecc == "SQL_SERVER":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["SERVIDOR"] + "(" + dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["BBDD"] + ")"


        lista_dicc_errores_procesos = dicc_errores_procesos[proceso_id][tipo_bbdd_selecc]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd]

        if isinstance(lista_dicc_errores_procesos, list):

            string_log = tipo_bbdd_selecc + ": ERRORES EN " + opcion_bbdd + ": " + str(nombre_bbdd) + "\n" + "-" * 200 + "\n" + "-" * 200
            for dicc in lista_dicc_errores_procesos:

                #TEMPORAL: nada mas encender el PC y ejecutar de inicio el proceso de control de versiones en MS ACCESS suele
                #aparecer el error "'NoneType' object has no attribute 'References'" (es como si el app mediante el framewwork win32com.client
                #no reconoce la bbdd cuando se ejecuta el hilo de threads por la 1era vez), la solucion cuando surje esto es ir a CMD 
                #y ejecutar el comando siguiente: taskkill /f /im msaccess.exe (no requiere cerrar el app y volver a abrirala
                # solo ejecutar el comando CMD y volver a jecutar el proceso desde la app)
                # --> es lo mismo que hacer Ctrl + Alt + Sup y cerrar los programas MS ACCESS activos)
                #se informa en los logs de errores si pasa este caso
                str_err_access_hilo_threads = ""
                if dicc["ERRORES"]  == "'NoneType' object has no attribute 'References'":

                    mensaje_err_1 = "\n\nSe ha producido un fallo en el hilo de threads de importación de las 2 bases de datos MS ACCESS. "
                    mensaje_err_2 = "Suele occurir cuando, nada más encender el PC, \nse ejecuta el proceso de control de versiones entre 2 bases de datos MS ACCESS. "
                    mensaje_err_3 = "Para resolver el problema basta con ir a CMD y ejecutar el comando siguiente:\ntaskkill /f /im msaccess.exe\n"
                    mensaje_err_4 = "(es lo mismo que pulsar la hotkey Ctrl + Alt + Sup y cerrar todas las bases de datos MS ACCESS activas)."
                    str_err_access_hilo_threads = mensaje_err_1 + mensaje_err_2 + mensaje_err_3 + mensaje_err_4


                #se genera el log acumulado
                modulo_python = dicc["MODULO_PYTHON"] if dicc["MODULO_PYTHON"] != None else "---"
                rutina_python = dicc["RUTINA_PYTHON"] if dicc["RUTINA_PYTHON"] != None else "---"
                linea_error = dicc["LINEA_ERROR"] if dicc["LINEA_ERROR"] != None else "---"
                errores = dicc["ERRORES"] + str_err_access_hilo_threads if dicc["ERRORES"] != None else "---"

                string_log = string_log + "\n\nMODULO PYTHON: " + modulo_python + "\n"
                string_log = string_log + "RUTINA PYTHON: " + rutina_python + "\n"
                string_log = string_log + "LINEA ERROR: " + str(linea_error) + "\n\n"
                string_log = string_log + "ERROR: " + errores + "\n\n"
                string_log = string_log + "-" * 200

                now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
                nombre_fich_logs = tipo_bbdd_selecc + "_LOGS_ERRORES_IMPORT_" + now + ".txt"
                saveas = str(ruta_destino_fichero_logs) + r"\%s" % nombre_fich_logs

                with open(saveas, 'w') as fich_log:
                    fich_log.write(string_log)


    elif opcion == "PROCESOS_LOGS_ERRORES_CALCULO":

        lista_dicc_errores_procesos = dicc_errores_procesos[proceso_id][tipo_bbdd_selecc]["LISTA_DICC_ERRORES_CALCULO"]

        if isinstance(lista_dicc_errores_procesos, list):

            proceso_concepto = dicc_procesos[proceso_id]["PROCESO"].upper()

            string_log = tipo_bbdd_selecc + ": ERRORES EN PROCESO " + proceso_concepto + "\n" + "-" * 200 + "\n" + "-" * 200
            for dicc in lista_dicc_errores_procesos:

                modulo_python = dicc["MODULO_PYTHON"] if dicc["MODULO_PYTHON"] != None else "---"
                rutina_python = dicc["RUTINA_PYTHON"] if dicc["RUTINA_PYTHON"] != None else "---"
                linea_error = dicc["LINEA_ERROR"] if dicc["LINEA_ERROR"] != None else "---"
                errores = dicc["ERRORES"] if dicc["ERRORES"] != None else "---"

                string_log = string_log + "\n\nMODULO PYTHON: " + modulo_python + "\n"
                string_log = string_log + "RUTINA PYTHON: " + rutina_python + "\n"
                string_log = string_log + "LINEA ERROR: " + str(linea_error) + "\n\n"
                string_log = string_log + "ERROR: " + errores + "\n\n"
                string_log = string_log + "-" * 200

                now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
                nombre_fich_logs = tipo_bbdd_selecc + "_LOGS_ERRORES_CALCULO_" + proceso_concepto.replace(" ", "_") + "_" + now + ".txt"
                saveas = str(ruta_destino_fichero_logs) + r"\%s" % nombre_fich_logs

                with open(saveas, 'w') as fich_log:
                    fich_log.write(string_log)



    elif opcion == "MERGE_BBDD_FISICA_LOGS_OK":

        if tipo_bbdd_selecc == "MS_ACCESS":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["PATH_BBDD"]

        elif tipo_bbdd_selecc == "SQL_SERVER":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["SERVIDOR"] + "(" + dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["BBDD"] + ")"


        lista_dicc_ok_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"]
        

        if isinstance(lista_dicc_ok_migracion, list):

            string_log = tipo_bbdd_selecc + ": ERRORES MERGE EN BBDD: " + str(nombre_bbdd) + "\n" + "-" * 200 + "\n" + "-" * 200
            for dicc in lista_dicc_ok_migracion:

                tipo_repositorio = dicc["TIPO_REPOSITORIO"] if dicc["TIPO_REPOSITORIO"] != None else "---"
                repositorio = dicc["REPOSITORIO"] if dicc["REPOSITORIO"] != None else "---"
                tipo_objeto = dicc["TIPO_OBJETO"] if dicc["TIPO_OBJETO"] != None else "---"
                nombre_objeto = dicc["NOMBRE_OBJETO"] if dicc["NOMBRE_OBJETO"] != None else "---"

                string_log = string_log + "\n\nTIPO REPOSITORIO: " + tipo_repositorio + "\n"
                string_log = string_log + "REPOSITORIO: " + repositorio + "\n"
                string_log = string_log + "TIPO OBJETO: " + tipo_objeto + "\n"
                string_log = string_log + "NOMBRE OBJETO: " + nombre_objeto + "\n\n"
                string_log = string_log + "-" * 200


                now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
                nombre_fich_logs = tipo_bbdd_selecc + "_LOGS_OK_MERGE_BBDD_" + now + ".txt"
                saveas = str(ruta_destino_fichero_logs) + r"\%s" % nombre_fich_logs

                with open(saveas, 'w') as fich_log:
                    fich_log.write(string_log)



    elif opcion == "MERGE_BBDD_FISICA_LOGS_ERRORES":

        if tipo_bbdd_selecc == "MS_ACCESS":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["PATH_BBDD"]

        elif tipo_bbdd_selecc == "SQL_SERVER":
            nombre_bbdd = dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["SERVIDOR"] + "(" + dicc_codigos_bbdd[opcion_bbdd][tipo_bbdd_selecc]["BBDD"] + ")"


        lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

        if isinstance(lista_dicc_errores_migracion, list):

            string_log = tipo_bbdd_selecc + ": ERRORES MERGE EN BBDD: " + str(nombre_bbdd) + "\n" + "-" * 200 + "\n" + "-" * 200
            for dicc in lista_dicc_errores_migracion:

                modulo_python = dicc["MODULO_PYTHON"] if dicc["MODULO_PYTHON"] != None else "---"
                rutina_python = dicc["RUTINA_PYTHON"] if dicc["RUTINA_PYTHON"] != None else "---"
                tipo_repositorio = dicc["TIPO_REPOSITORIO"] if dicc["TIPO_REPOSITORIO"] != None else "---"
                repositorio = dicc["REPOSITORIO"] if dicc["REPOSITORIO"] != None else "---"
                tipo_objeto = dicc["TIPO_OBJETO"] if dicc["TIPO_OBJETO"] != None else "---"
                nombre_objeto = dicc["NOMBRE_OBJETO"] if dicc["NOMBRE_OBJETO"] != None else "---"
                linea_error = dicc["LINEA_ERROR"] if dicc["LINEA_ERROR"] != None else "---"
                errores = dicc["ERRORES"] if dicc["ERRORES"] != None else "---"

                string_log = string_log + "\n\nMODULO PYTHON: " + modulo_python + "\n"
                string_log = string_log + "RUTINA PYTHON: " + rutina_python + "\n"
                string_log = string_log + "TIPO REPOSITORIO: " + tipo_repositorio + "\n"
                string_log = string_log + "REPOSITORIO: " + repositorio + "\n"
                string_log = string_log + "TIPO OBJETO: " + tipo_objeto + "\n"
                string_log = string_log + "NOMBRE OBJETO: " + nombre_objeto + "\n\n"
                string_log = string_log + "LINEA ERROR: " + str(linea_error) + "\n\n"
                string_log = string_log + "ERROR: " + errores + "\n\n"
                string_log = string_log + "-" * 200


                now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
                nombre_fich_logs = tipo_bbdd_selecc + "_LOGS_ERRORES_MERGE_BBDD_" + now + ".txt"
                saveas = str(ruta_destino_fichero_logs) + r"\%s" % nombre_fich_logs

                with open(saveas, 'w') as fich_log:
                    fich_log.write(string_log)


########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################
#                            RUTINAS + FUNCIONES - CALCULO DE PROCESOS
########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################


def func_se_puede_ejecutar_proceso(opcion_proceso, tipo_bbdd):
    #funcion que determina si se puede ejecutar los procesos de control de versiones y/o de diagnostico según el tipo de bbdd Access o SQL Server

    bbdd_access_1 = dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["PATH_BBDD"]
    bbdd_access_2 = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"]

    servidor_sql_server_1 = dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"]
    servidor_sql_server_2 = dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"]
    bbdd_sql_server_1 = dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"]
    bbdd_sql_server_2 = dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"]


    if opcion_proceso == "CONTROL_VERSIONES":

        if tipo_bbdd == "MS_ACCESS":

            #se comprueba que los 2 path estan configurados y son distintos
            return "SI" if bbdd_access_1 != None and bbdd_access_2 != None and bbdd_access_1 != bbdd_access_2 else "NO"


        elif tipo_bbdd == "SQL_SERVER":

            #se comprueba que servidor_1 + bbdd_1 y servidor_2 + bbdd_2 estan configurados y no son los mismos
            temp = "NO"
            if servidor_sql_server_1 != None and bbdd_sql_server_1 != None and servidor_sql_server_2 != None and bbdd_sql_server_2 != None:
                if servidor_sql_server_1 + bbdd_sql_server_1 != servidor_sql_server_2 + bbdd_sql_server_2:
                    temp = "SI"
                else:
                    temp = "NO"
            else:
                temp = "NO"


            return temp


    elif opcion_proceso == "DIAGNOSTICO":

        if tipo_bbdd == "MS_ACCESS":

            #se comprueba si el pathh de bbdd_1 esta configurado
            return "SI" if bbdd_access_1 != None else "NO"


        elif tipo_bbdd == "SQL_SERVER":

            #se comprueba si el Servidor_1 esta configurado
            return "SI" if servidor_sql_server_1 != None else "NO"



def func_dicc_control_versiones(**kwargs):
    #funcion que permite realizar el control de versiones sobre 2 df que contienen el mismo script en las 2 BBDD
    #en caso que el objeto este solo en una bbdd (BBDD_01 o BBDD_02) el control de versiones se limita a guardar el script de la bbdd donde esta
    #la función devuelve un diccionario con la misma estructura sea que el objeto ya existe en las 2 bbdd o que solo este en BBDD_01 o que solo este en BBDD_02


    #parametros kwargs
    tipo_bbdd = kwargs.get("tipo_bbdd", None)
    check_objeto = kwargs.get("check_objeto", None)
    tipo_objeto = kwargs.get("tipo_objeto", None)
    tipo_repositorio = kwargs.get("tipo_repositorio", None)
    repositorio = kwargs.get("repositorio", None)
    nombre_objeto = kwargs.get("nombre_objeto", None)
    df_codigo_bbdd_1 = kwargs.get("df_codigo_bbdd_1", None)
    df_codigo_bbdd_2 = kwargs.get("df_codigo_bbdd_2", None)

    tipo_objeto_subform = dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["TIPO_OBJETO_SUBFORM"]

    #se modifican los df_codigo_bbdd_1 + df_codigo_bbdd_2 (que ya vienen con la columna CODIGO) agregando las columnas adiciones (de momento se establecen a None)
    #y se reordenan dentro de los nuevos df:
    # --> NUM_LINEA
    # --> CODIGO_CON_NUM_LINEA
    # --> CONTROL_CAMBIOS_ORIGINAL
    # --> CONTROL_CAMBIOS_ACTUAL
    df_temp1 = pd.DataFrame()
    if isinstance(df_codigo_bbdd_1, pd.DataFrame):
        df_temp1 = df_codigo_bbdd_1.copy()

        df_temp1[lista_headers_df_codigo_control_versiones_1] = [None, None, None, None]
        df_temp1 = df_temp1[lista_headers_df_codigo_control_versiones_2]


    df_temp2 = pd.DataFrame()
    if isinstance(df_codigo_bbdd_2, pd.DataFrame):
        df_temp2 = df_codigo_bbdd_2.copy()

        df_temp2[lista_headers_df_codigo_control_versiones_1] = [None, None, None, None]
        df_temp2 = df_temp2[lista_headers_df_codigo_control_versiones_2]


    if check_objeto == dicc_control_versiones_tipo_concepto["YA_EXISTE"]:

        lista_script_1 = df_codigo_bbdd_1["CODIGO"].tolist()
        lista_script_2 = df_codigo_bbdd_2["CODIGO"].tolist()

        diff = difflib.ndiff(lista_script_1, lista_script_2)


        #difflib marca los scripts con los cambios de la forma siguiente al prinicipio de cada linea con cambios:
        # --> con el signo "-" las diferencias del script 1 con respecto al script 2
        # --> con el signo "+" las diferencias del script 2 con respecto al script 1
        #
        #se crean las listas lista_cambios_script_1 y lista_cambios_script_2 para almacenar las lineas de codigo con cambios 
        #(neteadas de este signo + o - al principio de la linea)
        #ademas se crea los contadores num_cambios_script_1 y num_cambios_script_2 de cambios en cada script
        lista_cambios_script_1 = []
        lista_cambios_script_2 = []

        num_cambios_script_1 = 0
        num_cambios_script_2 = 0

        for line in diff:

            if len(line.replace(" ", "").replace("\t", "")) != 0:

                if line.startswith('- ') and len(line[2:].replace(" ", "").replace("\t", "")) != 0:
                    lista_cambios_script_1.append(line[2:])
                    num_cambios_script_1 += 1

                elif line.startswith('+ ') and len(line[2:].replace(" ", "").replace("\t", "")) != 0:
                    lista_cambios_script_2.append(line[2:])
                    num_cambios_script_2 += 1


        #se crea la lista con los numeros de linea del script_1 donde hay cambios
        lista_cambios_lineas_script_1 = []
        if len(lista_cambios_script_1) != 0:

            lista_codigo_con_num_linea = [[ind + 1, item] for ind, item in enumerate(lista_script_1)]

            for linea_cambio in lista_cambios_script_1:
                for num_linea, linea_codigo in lista_codigo_con_num_linea:

                    if linea_cambio.strip() == linea_codigo.strip():
                        lista_cambios_lineas_script_1.append(num_linea)

                        break

            #se rellena CONTROL_CAMBIOS
            for ind in df_temp1.index:
                num_linea = ind + 1

                for item in lista_cambios_lineas_script_1:
                    if num_linea == item:
                        df_temp1.iloc[ind, df_temp1.columns.get_loc("CONTROL_CAMBIOS_ORIGINAL")] = "CAMBIOS_LOCALIZADOS"
                        df_temp1.iloc[ind, df_temp1.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")] = "CAMBIOS_LOCALIZADOS"
                        break


        #se crea la lista con los numeros de linea del script_2 donde hay cambios
        lista_cambios_lineas_script_2 = []
        if len(lista_cambios_script_2) != 0:

            lista_codigo_con_num_linea = [[ind + 1, item] for ind, item in enumerate(lista_script_2)]

            for linea_cambio in lista_cambios_script_2:
                for num_linea, linea_codigo in lista_codigo_con_num_linea:

                    if linea_cambio.strip() == linea_codigo.strip():
                        lista_cambios_lineas_script_2.append(num_linea)

                        break

            #se rellena CONTROL_CAMBIOS
            for ind in df_temp2.index:
                num_linea = ind + 1

                for item in lista_cambios_lineas_script_2:
                    if num_linea == item:
                        df_temp2.iloc[ind, df_temp2.columns.get_loc("CONTROL_CAMBIOS_ORIGINAL")] = "CAMBIOS_LOCALIZADOS"
                        df_temp2.iloc[ind, df_temp2.columns.get_loc("CONTROL_CAMBIOS_ACTUAL")] = "CAMBIOS_LOCALIZADOS"
                        break


        #se crea el diccionario que sirve de resultado de la funcion
        hay_diferencias = "SI" if len(lista_cambios_script_1) + len(lista_cambios_script_2) != 0 else "NO"

        dicc_temp = {"TIPO_BBDD": tipo_bbdd
                    , "CHECK_OBJETO": check_objeto if hay_diferencias == "SI" else None
                    , "TIPO_OBJETO": tipo_objeto if hay_diferencias == "SI" else None
                    , "TIPO_OBJETO_SUBFORM": tipo_objeto_subform if hay_diferencias == "SI" else None
                    , "TIPO_REPOSITORIO": tipo_repositorio if hay_diferencias == "SI" else None
                    , "REPOSITORIO": repositorio if hay_diferencias == "SI" else None
                    , "NOMBRE_OBJETO": nombre_objeto if hay_diferencias == "SI" else None
                    , "NUM_CAMBIOS_SCRIPT_1": num_cambios_script_1 if hay_diferencias == "SI" else None
                    , "NUM_CAMBIOS_SCRIPT_2": num_cambios_script_2 if hay_diferencias == "SI" else None
                    , "DF_CODIGO_ORIGINAL_1": df_temp1 if hay_diferencias == "SI" else None
                    , "DF_CODIGO_ORIGINAL_2": df_temp2 if hay_diferencias == "SI" else None
                    , "LISTA_DICC_MERGE_HECHOS": None              #se usa para saber que merge se han hecho (aqui es None al iniciar el app)
                    , "DF_CODIGO_ACTUAL_1": df_temp1 if hay_diferencias == "SI" else None
                    , "DF_CODIGO_ACTUAL_2": df_temp2 if hay_diferencias == "SI" else None                                 
                    }


        #resultado de la funcion (devuelve el diccionario dicc_temp si al menos uno de los valores asociados a las keys es distinto de None sino devuelve None)
        return dicc_temp if sum(1 if isinstance(dicc_temp[key], pd.DataFrame) else 1 if dicc_temp[key] != None else 0 for key in dicc_temp.keys()) != 0 else None
        



    elif check_objeto == dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_01"]:

        num_cambios_script_1 = len(df_temp1)
        df_temp1[["CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]] = ["CAMBIOS_LOCALIZADOS", "CAMBIOS_LOCALIZADOS"]


        dicc_temp =  {"TIPO_BBDD": tipo_bbdd
                    , "CHECK_OBJETO": check_objeto
                    , "TIPO_OBJETO": tipo_objeto
                    , "TIPO_OBJETO_SUBFORM": tipo_objeto_subform
                    , "TIPO_REPOSITORIO": tipo_repositorio
                    , "REPOSITORIO": repositorio
                    , "NOMBRE_OBJETO": nombre_objeto
                    , "NUM_CAMBIOS_SCRIPT_1": num_cambios_script_1
                    , "NUM_CAMBIOS_SCRIPT_2": 0
                    , "DF_CODIGO_ORIGINAL_1": df_temp1
                    , "DF_CODIGO_ORIGINAL_2": None
                    , "LISTA_DICC_MERGE_HECHOS": None       #se usa para saber que merge se han hecho (aqui es None al iniciar el app)
                    , "DF_CODIGO_ACTUAL_1": df_temp1
                    , "DF_CODIGO_ACTUAL_2": None
                    }

        #resultado de la funcion (devuelve el diccionario dicc_temp si al menos uno de los valores asociados a las keys es distinto de None sino devuelve None)
        return dicc_temp if sum(1 if isinstance(dicc_temp[key], pd.DataFrame) else 1 if dicc_temp[key] != None else 0 for key in dicc_temp.keys()) != 0 else None
    



    elif check_objeto == dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_02"]:

        num_cambios_script_2 = len(df_temp2)
        df_temp2[["CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]] = ["CAMBIOS_LOCALIZADOS", "CAMBIOS_LOCALIZADOS"]


        dicc_temp =  {"TIPO_BBDD": tipo_bbdd
                    , "CHECK_OBJETO": check_objeto
                    , "TIPO_OBJETO": tipo_objeto
                    , "TIPO_OBJETO_SUBFORM": tipo_objeto_subform
                    , "TIPO_REPOSITORIO": tipo_repositorio
                    , "REPOSITORIO": repositorio
                    , "NOMBRE_OBJETO": nombre_objeto
                    , "NUM_CAMBIOS_SCRIPT_1": 0
                    , "NUM_CAMBIOS_SCRIPT_2": num_cambios_script_2
                    , "DF_CODIGO_ORIGINAL_1": None
                    , "DF_CODIGO_ORIGINAL_2": df_temp2
                    , "LISTA_DICC_MERGE_HECHOS": None         #se usa para saber que merge se han hecho (aqui es None al iniciar el app)
                    , "DF_CODIGO_ACTUAL_1": None
                    , "DF_CODIGO_ACTUAL_2": df_temp2
                    }

        #resultado de la funcion (devuelve el diccionario dicc_temp si al menos uno de los valores asociados a las keys es distinto de None sino devuelve None)
        return dicc_temp if sum(1 if isinstance(dicc_temp[key], pd.DataFrame) else 1 if dicc_temp[key] != None else 0 for key in dicc_temp.keys()) != 0 else None




def def_calc_global(proceso_selecc_id, ruta_destino_fichero_logs, **kwargs):
    #rutina para ejecutar los threads segun el proceso escojido

    global global_proceso_en_ejecucion
    global global_msg_errores_proceso_access
    global global_msg_errores_proceso_sql_server


    #parametros kwargs
    ruta_destino_excel_diagnostico_access = kwargs.get("ruta_destino_excel_diagnostico_access", None)
    opcion_diagnostico_sql_server = kwargs.get("opcion_diagnostico_sql_server", None)
    ruta_destino_diagnostico_sql_server = kwargs.get("ruta_destino_diagnostico_sql_server", None)


    #se inicializan las variables globales de control de errores a ""
    global_msg_errores_proceso_access = ""
    global_msg_errores_proceso_sql_server = ""


    #se vacia dicc_codigos_bbdd al iniciar (salvo los conceptos estaticos)
    for num_bbdd in dicc_codigos_bbdd.keys():
        for tipo_bbdd in dicc_codigos_bbdd[num_bbdd].keys():

            if tipo_bbdd == "MS_ACCESS":
                for key in dicc_codigos_bbdd[num_bbdd][tipo_bbdd].keys():
                    if key != "PATH_BBDD":
                        dicc_codigos_bbdd[num_bbdd][tipo_bbdd][key] = None


            elif tipo_bbdd == "SQL_SERVER":
                for key in dicc_codigos_bbdd[num_bbdd][tipo_bbdd].keys():
                    if key not in ["SERVIDOR", "BBDD", "CONNECTING_STRING"]:
                        dicc_codigos_bbdd[num_bbdd][tipo_bbdd][key] = None



    #se vacia dicc_control_versiones_tipo_objeto al iniciar (salvo los conceptos estaticos)
    for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():

        dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"] = None
        dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = None
        dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"] = None

        if tipo_bbdd == "MS_ACCESS":
            dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["LISTA_DICC_OBJETOS"] = None

        for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():
            dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = None



    #se vacia dicc_errores_procesos al iniciar
    for proceso in dicc_errores_procesos.keys():
        for tipo_bbdd in dicc_errores_procesos[proceso].keys():
            for error in dicc_errores_procesos[proceso][tipo_bbdd].keys():
                dicc_errores_procesos[proceso][tipo_bbdd][error] = None




    #se recuperan el path de cada bbdd access y el servidor/bbdd de cada bbdd sql server
    access_path_bbdd_1 = dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["PATH_BBDD"]
    access_path_bbdd_2 = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"]

    sql_server_servidor_1 = dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"]
    sql_server_bbdd_lista_1 = dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"] if isinstance(dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"], list) else [dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"]]

    sql_server_servidor_2 = dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"]
    sql_server_bbdd_lista_2 = dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"] if isinstance(dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"], list) else [dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"]]



    if proceso_selecc_id == "PROCESO_01":
        #CONTROL VERSIONES
        #se ejecuta 1 hilo de threads sobre def_proceso_access_1_import y/o def_proceso_sql_server_1_import (segun se pueda realizar en access o sql server)
        #cuando finalice se ejecuta def_proceso_access_2_control_versiones y/o def_proceso_sql_server_2_control_versiones


        ##########################################################################
        #se activa la variable global global_proceso_en_ejecucion para impedir desde la GUI ejecutar otro proceso
        #antes de que acabe el proceso en curso
        ##########################################################################
        global_proceso_en_ejecucion = "SI"


        ##########################################################################
        #hilo de threads
        ##########################################################################

        lista_threads = []
        if func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS") == "SI":

            thread = Thread(target = mod_access.def_proceso_access_1_import, args = (proceso_selecc_id, "BBDD_01", access_path_bbdd_1))
            lista_threads.append(thread)

            thread = Thread(target = mod_access.def_proceso_access_1_import, args = (proceso_selecc_id, "BBDD_02", access_path_bbdd_2))
            lista_threads.append(thread)


        if func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "SQL_SERVER") == "SI":

            thread = Thread(target = mod_sql_server.def_proceso_sql_server_1_import, args = (proceso_selecc_id, "BBDD_01", sql_server_servidor_1, sql_server_bbdd_lista_1))
            lista_threads.append(thread)

            thread = Thread(target = mod_sql_server.def_proceso_sql_server_1_import, args = (proceso_selecc_id, "BBDD_02", sql_server_servidor_2, sql_server_bbdd_lista_2))
            lista_threads.append(thread)


        for thread in lista_threads:
            thread.start()

        for thread in lista_threads:
            thread.join()


        #se cierran todos los access abiertos al finalizar el hilo de threads si el proceso afecta access
        if func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS") == "SI":
            try:
                subprocess.Popen("taskkill /f /im msaccess.exe")
            except:
                pass


        ##########################################################################
        #se localiza si hay errores de ejecucion en el proceso de import del control de versiones
        #en caso de que no los hay se pasa a ejecutar el calculo del proceso de control de versiones (sin threads)
        ##########################################################################

        #se informa la variable global global_msg_errores_proceso_access en caso de errores de ejecucion
        #y se generan los logs
        #global_msg_errores_proceso_access se usa en la GUI para generar el mensaje que sale en el warning
        #al final de la ejecucion del proceso si hay errores

        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():

            tipo_bbdd_msg = tipo_bbdd.replace("_", " ")

            if func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", tipo_bbdd) == "SI":
                
                if func_procesos_errores(proceso_selecc_id, tipo_bbdd, "IMPORT") == "SIN_ERROR":

                    if tipo_bbdd == "MS_ACCESS":
                        mod_access.def_proceso_access_2_control_versiones()

                    elif tipo_bbdd == "SQL_SERVER":
                        mod_sql_server.def_proceso_sql_server_2_control_versiones()


                    if func_procesos_errores(proceso_selecc_id, tipo_bbdd, "CALCULO") == "ERROR":
                        mensaje1 = "Control versiones " + tipo_bbdd_msg + "\n\nEl proceso de importacion de los datos se ha realizado correctamente"
                        mensaje2 = "pero el proceso de calculo del control de versiones ha generado errores por lo que el proceso se ha cancelado (consulta los logs).\n\n"

                        if tipo_bbdd == "MS_ACCESS":
                            global_msg_errores_proceso_access = mensaje1 + mensaje2

                        elif tipo_bbdd == "SQL_SERVER":
                            global_msg_errores_proceso_sql_server = mensaje1 + mensaje2

                        #se generan los logs en ficheros .txt
                        def_generacion_logs("PROCESOS_LOGS_ERRORES_CALCULO", tipo_bbdd, ruta_destino_fichero_logs, proceso_id = proceso_selecc_id)


                else:
                    if tipo_bbdd == "MS_ACCESS":
                        global_msg_errores_proceso_access = "Control versiones " + tipo_bbdd_msg + "\n\nSe han localizado errores en el proceso de importacion de los datos por lo que el proceso se ha cancelado (consulta los logs).\n\n"
                    
                    elif tipo_bbdd == "SQL_SERVER":
                        global_msg_errores_proceso_sql_server = "Control versiones " + tipo_bbdd_msg + "\n\nSe han localizado errores en el proceso de importacion de los datos por lo que el proceso se ha cancelado (consulta los logs).\n\n"

                    #se generan los logs en ficheros .txt
                    def_generacion_logs("PROCESOS_LOGS_ERRORES_IMPORT_BBDD", tipo_bbdd, ruta_destino_fichero_logs, proceso_id = proceso_selecc_id, opcion_bbdd = "BBDD_01")
                    def_generacion_logs("PROCESOS_LOGS_ERRORES_IMPORT_BBDD", tipo_bbdd, ruta_destino_fichero_logs, proceso_id = proceso_selecc_id, opcion_bbdd = "BBDD_02")
                

        ##########################################################################
        #se desactiva la variable global global_proceso_en_ejecucion para poder desde la GUI ejecutar otro proceso
        ##########################################################################
        global_proceso_en_ejecucion = "NO"



    elif proceso_selecc_id == "PROCESO_02":
        #DIAGNOSTICO ACCESS (se realiza por defecto sobre BBDD_01)
        #aqui no hay hilos de threads puesto que no se puede ejecutar simultaneamente sobre una bbdd access y sql server


        ##########################################################################
        #se activa la variable global global_proceso_en_ejecucion para impedir desde la GUI ejecutar otro proceso
        #antes de que acabe el proceso en curso
        ##########################################################################
        global_proceso_en_ejecucion = "SI"


        ##########################################################################
        #se ejecuta el proceso de diagnostico si no se localizan errores en la fase de import
        ##########################################################################

        mod_access.def_proceso_access_1_import(proceso_selecc_id, "BBDD_01", access_path_bbdd_1)

        #se cierran todos los access abiertos al finalizar el hilo de threads si el proceso afecta access
        #(por si acaso no se hubiesen cerrado bien)
        try:
            subprocess.Popen("taskkill /f /im msaccess.exe")
        except:
            pass


        #se informa la variable global global_msg_errores_proceso_access en caso de errores de ejecucion
        #y se generan los logs
        if func_se_puede_ejecutar_proceso("DIAGNOSTICO", "MS_ACCESS") == "SI":
            
            if func_procesos_errores(proceso_selecc_id, "MS_ACCESS", "IMPORT") == "SIN_ERROR":
                mod_access.def_proceso_access_2_diagnostico(ruta_destino_excel_diagnostico_access)

                if func_procesos_errores(proceso_selecc_id, "MS_ACCESS", "CALCULO") == "ERROR":

                    mensaje1 = "Diagnostico MS ACCESS\n\nEl proceso de importacion de los datos se ha realizado correctamente"
                    mensaje2 = "pero el proceso de calculo del control de versiones ha generado errores por lo que el proceso se ha cancelado (consulta los logs).\n\n"
                    global_msg_errores_proceso_access = mensaje1 + mensaje2

                    def_generacion_logs("PROCESOS_LOGS_ERRORES_CALCULO", "MS_ACCESS", ruta_destino_fichero_logs, proceso_id = proceso_selecc_id)

            else:
                global_msg_errores_proceso_access = "Diagnostico MS ACCESS\n\nSe han localizado errores en el proceso de importacion de los datos por lo que el proceso se ha cancelado (consulta los logs).\n\n"

                def_generacion_logs("PROCESOS_LOGS_ERRORES_IMPORT_BBDD", "MS_ACCESS", ruta_destino_fichero_logs, proceso_id = proceso_selecc_id, opcion_bbdd = "BBDD_01")



        ##########################################################################
        #se desactiva la variable global global_proceso_en_ejecucion para poder desde la GUI ejecutar otro proceso
        ##########################################################################
        global_proceso_en_ejecucion = "NO"



    elif proceso_selecc_id == "PROCESO_03":
        #DIAGNOSTICO SQL SERVER (se realiza por defecto sobre BBDD_01)
        #aqui no hay hilos de threads puesto que no se puede ejecutar simultaneamente sobre una bbdd access y sql server


        ##########################################################################
        #se activa la variable global global_proceso_en_ejecucion para impedir desde la GUI ejecutar otro proceso
        #antes de que acabe el proceso en curso
        ##########################################################################
        global_proceso_en_ejecucion = "SI"


        ##########################################################################
        #se ejecuta el proceso de diagnostico si no se localizan errores en la fase de import
        ##########################################################################

        mod_sql_server.def_proceso_sql_server_1_import(proceso_selecc_id, "BBDD_01", sql_server_servidor_1, sql_server_bbdd_lista_1)

        #se informa la variable global global_msg_errores_proceso_sql_server en caso de errores de ejecucion y se generan los logs
        if func_se_puede_ejecutar_proceso("DIAGNOSTICO", "SQL_SERVER") == "SI":
            
            if func_procesos_errores(proceso_selecc_id, "SQL_SERVER", "IMPORT") == "SIN_ERROR":

                mod_sql_server.def_proceso_sql_server_2_diagnostico(opcion_diagnostico_sql_server, sql_server_servidor_1, sql_server_bbdd_lista_1, ruta_destino_diagnostico_sql_server)

                if func_procesos_errores(proceso_selecc_id, "SQL_SERVER", "CALCULO") == "ERROR":
                    mensaje1 = "Diagnostico SQL SERVER\n\nEl proceso de importacion de los datos se ha realizado correctamente "
                    mensaje2 = "pero el proceso de calculo del diagnostico de dependencias ha generado errores por lo que el proceso se ha cancelado (consulta los logs).\n\n"
                    global_msg_errores_proceso_sql_server = mensaje1 + mensaje2

                    def_generacion_logs("PROCESOS_LOGS_ERRORES_CALCULO", "SQL_SERVER", ruta_destino_fichero_logs, proceso_id = proceso_selecc_id)

            else:
                global_msg_errores_proceso_sql_server = "Diagnostico SQL SERVER\n\nSe han localizado errores en el proceso de importacion de los datos por lo que el proceso se ha cancelado (consulta los logs).\n\n"

                def_generacion_logs("PROCESOS_LOGS_ERRORES_IMPORT_BBDD", "SQL_SERVER", ruta_destino_fichero_logs, proceso_id = proceso_selecc_id, opcion_bbdd = "BBDD_01")



        ##########################################################################
        #se cierran todos los access abiertos al iniciar el proceso en caso de que este ultimo0 afecta a MS ACCESS
        ##########################################################################
        if func_se_puede_ejecutar_proceso("CONTROL_VERSIONES", "MS_ACCESS") == "SI":
            try:
                subprocess.Popen("taskkill /f /im msaccess.exe")
            except:
                pass


        ##########################################################################
        #se desactiva la variable global global_proceso_en_ejecucion para poder desde la GUI ejecutar otro proceso
        ##########################################################################
        global_proceso_en_ejecucion = "NO"




########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################
#                            FUNCIONES + RUTINAS - GUI CONTROL VERSIONES
########################################################################################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################

def func_dicc_control_versiones_tipo_objeto_buscar_en_dicc(opcion, **kwargs):
    #funcion que permite buscar keys asociadas a subkeys en dicc_control_versiones_tipo_objeto
    #permite localizar la lista de dicc de objetos por tipo de bbdd (MS_ACCESS o SQL_SERVER) donde hay que realizar merge3 en bbdd fisicas
    #
    #es más limpio realizar esta busqueda por el metodo next
    #ejemplo de 1 caso usado en el modulo APP_1_GUI: next((key for key, value in mod_gen.dicc_procesos.items() if value["PROCESO"] == proceso_selecc), None)
    #pero he preferido hacerlo por bucles anidados y con la instruccion break para darle mas legibilidad al codigo

    valor = kwargs.get("valor", None)
    opcion_gui_merge_bbdd_fisica = kwargs.get("opcion_gui_merge_bbdd_fisica", None)


    temp = ""
    if opcion == "TIPO_BBDD":
        #busca el tipo de bbdd asociado asociado a la opcion del combobox de seleccion de objetos en el control de versiones
        #que es configurable mientras que el tipo de bbdd es fijo (MS_ACCESS o SQL_SERVER), se reutiliza en el codigo en varios sitios

        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
            for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():
                if valor == dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["COMBOBOX_CONTROL_VERSIONES"]:
                    temp = tipo_bbdd
                    break



    elif opcion == "TIPO_OBJETO":
        #busca la key principal asociada a la opcion del combobox de seleccion de objetos en el control de versiones pq el combobox es configurable
        #pero no la key principal del objeto que esta ha de ser fija (se reutiliza en el codigo en varios sitios)
        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
            for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():
                if valor == dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["COMBOBOX_CONTROL_VERSIONES"]:
                    temp = tipo_objeto
                    break



    elif opcion == "TIPO_OBJETO_DESDE_SUBFORM":
        #busca la key tipo objeto asociada al literal que sale en el subformulario (este literal es configurable, la key no)

        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
            for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():
                if valor == dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["TIPO_OBJETO_SUBFORM"]:
                    temp = tipo_objeto
                    break



    elif opcion == "TIPO_BBDD_REALIZAR_MERGE_BBDD_FISICAS":
        #localiza en que tipo de bbdd se han realizado acciones de merge para replicarlas en bbdd fisicas

        temp = []
        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():

            for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():

                if tipo_objeto != "TODOS":
                    if isinstance(dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"], list):

                        for dicc in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]:

                            if isinstance(dicc["LISTA_DICC_MERGE_HECHOS"], list):

                                if tipo_bbdd not in temp:
                                    temp.append(tipo_bbdd)




    elif opcion == "LISTA_COMBOBOX_MERGE_BBDD_FISICAS":
        #crea la lista de opciones para el combobox de seleccion en la GUI de merge en bbdd fisica

        temp = []
        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():

            if tipo_bbdd == "MS_ACCESS":

                if isinstance(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["LISTA_DICC_OBJETOS"], list):
                    temp.append(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["COMBOBOX_GUI"])

                if isinstance(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"], list):
                    temp.append(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["COMBOBOX_GUI"])


            elif tipo_bbdd == "SQL_SERVER":

                if isinstance(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"], list):
                    temp.append(dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["COMBOBOX_GUI"])





    elif opcion == "LISTA_DICC_OBJETOS_MERGE_BBDD_FISICAS":
        #localiza la lista de objetos donde se ha hecho un merge en la gui de control de versiones para una opcion especifica del combobox merge bbdd fisica

        temp = []
        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
            for key_opciones_merge_bbdd_fisica in dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"].keys():

                if key_opciones_merge_bbdd_fisica not in ["LISTA_DICC_ERRORES_MIGRACION", "LISTA_DICC_OK_MIGRACION"]:
                    if dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"][key_opciones_merge_bbdd_fisica]["COMBOBOX_GUI"] == opcion_gui_merge_bbdd_fisica:
                        temp = dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"][key_opciones_merge_bbdd_fisica]["LISTA_DICC_OBJETOS"]
                        temp = temp if isinstance(temp, list) else None
                        break



    elif opcion == "KEY_MERGE_BBDD_FISICAS_COMBOBOX_GUI":
        #localiza la key segun la opcion del combobox de merge en bbdd fisica seleccionada

        temp = []
        for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
            for key_opciones_merge_bbdd_fisica in dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"].keys():

                if key_opciones_merge_bbdd_fisica not in ["LISTA_DICC_ERRORES_MIGRACION", "LISTA_DICC_OK_MIGRACION"]:
                    if dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"][key_opciones_merge_bbdd_fisica]["COMBOBOX_GUI"] == opcion_gui_merge_bbdd_fisica:
                        temp = key_opciones_merge_bbdd_fisica
                        break


    return temp



def func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_combobox_selecc, tipo_objeto_subform, repositorio_subform, nombre_objeto_subform):
    #funcion que permite localizar un objeto en la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) de dicc_control_versiones_tipo_objeto
    #devuelve un diccionario con estas keys:
    # --> INDICE_DICC_LISTA_DICC_OBJETOS_CONTROL_VERSIONES      es el indice del diccionario en LISTA_DICC_CONTROL_VERSIONES para el objeto buscado
    # --> LISTA_DICC_MERGE_HECHOS                               es la lista de diccionarios de merge hecgos
    # --> DICC_CONTROL_VERSIONES                                es el diccionario de LISTA_DICC_CONTROL_VERSIONES para el objeto buscado
    # --> DF_CODIGO_ORIGINAL_1                                  es el df codigo original del script de BBDD_01
    # --> DF_CODIGO_ORIGINAL_2                                  es el df codigo original del script de BBDD_02
    # --> DF_CODIGO_ACTUAL_1                                    es el df codigo actual del script de BBDD_01
    # --> DF_CODIGO_ACTUAL_2                                    es el df codigo actual del script de BBDD_02


    tipo_objeto_key = func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = tipo_objeto_subform)


    #se busca la lista lista_control_versiones asociada al tipo de objeto seleccionadao en el combobox
    lista_control_versiones = None
    for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():
        for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():

            tipo_objeto_config_combobox = dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["COMBOBOX_CONTROL_VERSIONES"]

            if tipo_objeto_config_combobox == tipo_objeto_combobox_selecc:
                lista_control_versiones = dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]
                break


    #dentro de la lista lista_control_versiones se busca el diccionario asociado al objeto buscado y el indice de este diccionario en la lista
    dicc_buscado_en_lista_control_versiones = None
    indice_dicc_buscado_en_lista_control_versiones = None

    for ind, dicc in enumerate(lista_control_versiones):

        tipo_bbdd_seek = dicc["TIPO_BBDD"]
        tipo_objeto_subform_seek = dicc["TIPO_OBJETO_SUBFORM"]
        repositorio_seek = dicc["REPOSITORIO"]
        nombre_objeto_seek = dicc["NOMBRE_OBJETO"]

        if tipo_bbdd_seek == "MS_ACCESS":

            #para TABLA_LOCAL, VINCULO_ODBC, VINCULO_OTRO la busqueda se hace por tipo de objeto y por nombre de objeto
            if tipo_objeto_key in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:

                if tipo_objeto_subform == tipo_objeto_subform_seek and nombre_objeto_subform == nombre_objeto_seek:

                    indice_dicc_buscado_en_lista_control_versiones = ind
                    dicc_buscado_en_lista_control_versiones = dicc
                    lista_merge_hechos = dicc["LISTA_DICC_MERGE_HECHOS"]

                    break


            #para RUTINAS_VBA la busqueda se hace por tipo de objeto, por repositorio y por nombre de objeto
            elif tipo_objeto_key == "RUTINAS_VBA":

                if tipo_objeto_subform == tipo_objeto_subform_seek and repositorio_subform == repositorio_seek and nombre_objeto_subform == nombre_objeto_seek:

                    indice_dicc_buscado_en_lista_control_versiones = ind
                    dicc_buscado_en_lista_control_versiones = dicc
                    lista_merge_hechos = dicc["LISTA_DICC_MERGE_HECHOS"]

                    break

            #para VARIABLES_VBA la busqueda se hace por tipo de objeto y por repositorio
            elif tipo_objeto_key == "VARIABLES_VBA":

                if tipo_objeto_subform == tipo_objeto_subform_seek and repositorio_subform == repositorio_seek:

                    indice_dicc_buscado_en_lista_control_versiones = ind
                    dicc_buscado_en_lista_control_versiones = dicc
                    lista_merge_hechos = dicc["LISTA_DICC_MERGE_HECHOS"]

                    break


        elif tipo_bbdd_seek == "SQL_SERVER":

            #para SQL_SERVER (todos los tipo de objeto) la busqueda se hace por tipo de objeto, por repositorio y por nombre de objeto
            if tipo_objeto_subform == tipo_objeto_subform_seek and repositorio_subform == repositorio_seek and nombre_objeto_subform == nombre_objeto_seek:

                indice_dicc_buscado_en_lista_control_versiones = ind
                dicc_buscado_en_lista_control_versiones = dicc
                lista_merge_hechos = dicc["LISTA_DICC_MERGE_HECHOS"]

                break

    #se calculan df codigo actuales de BBD_01 y BBDD_02
    #si no hay merge hechos se cojen las keys DF_CODIGO_ORIGINAL_1 y DF_CODIGO_ORIGINAL_2 del diccionario de LISTA_DICC_OBJETOS_CONTROL_VERSIONES
    #si hay merge hechos por el usuario se coje el ultimo diccionario de LISTA_DICC_MERGE_HECHOS para las keys DF_CODIGO_ACTUAL_1 y DF_CODIGO_ACTUAL_2
    df_codigo_original_1 = dicc_buscado_en_lista_control_versiones["DF_CODIGO_ORIGINAL_1"]
    df_codigo_original_2 = dicc_buscado_en_lista_control_versiones["DF_CODIGO_ORIGINAL_2"]

    df_codigo_actual_1 = lista_merge_hechos[-1]["DF_CODIGO_ACTUAL_1"] if isinstance(lista_merge_hechos, list) else df_codigo_original_1
    df_codigo_actual_2 = lista_merge_hechos[-1]["DF_CODIGO_ACTUAL_2"] if isinstance(lista_merge_hechos, list) else df_codigo_original_2


    #resultado de la funcion
    return {"INDICE_DICC_LISTA_DICC_OBJETOS_CONTROL_VERSIONES": indice_dicc_buscado_en_lista_control_versiones
            , "LISTA_DICC_MERGE_HECHOS": lista_merge_hechos
            , "DICC_CONTROL_VERSIONES": dicc_buscado_en_lista_control_versiones
            , "DF_CODIGO_ORIGINAL_1": df_codigo_original_1
            , "DF_CODIGO_ORIGINAL_2": df_codigo_original_2
            , "DF_CODIGO_ACTUAL_1": df_codigo_actual_1
            , "DF_CODIGO_ACTUAL_2": df_codigo_actual_2
            }




def def_control_versiones_export_excel(tipo_bbdd, lista_control_versiones_selecc, ruta_excel):
    #rutina que permite en la GUI de control de versiones listar los objetos localizados con cambios en un fichero excel
    #se exportan todos los objetos según el tipo de bbdd (MS_ACCESS o SQL_SERVER) asociado a la opcion seleccionada en el comobox de tipo de seleccion

    now = dt.datetime.now()
    saveas = str(ruta_excel) + r"\CONTROL_VERSIONES_" + tipo_bbdd + "_" + str(re.sub("[^0-9a-zA-Z]+", "_", str(now))) + ".xlsb"


    #se convierte en un df acorde a la plantilla excel ruta_plantilla_control_versiones_xls para su exportacion
    lista_df_xls = []
    for dicc in lista_control_versiones_selecc:
        if dicc["TIPO_OBJETO_SUBFORM"] != None:
            lista_df_xls_temp = [dicc["TIPO_BBDD"], dicc["TIPO_OBJETO_SUBFORM"], dicc["REPOSITORIO"], dicc["NOMBRE_OBJETO"], dicc["NUM_CAMBIOS_SCRIPT_1"], dicc["NUM_CAMBIOS_SCRIPT_2"], dicc["CHECK_OBJETO"]]
            lista_df_xls.append(lista_df_xls_temp)

    df_export_xls = pd.DataFrame(lista_df_xls, columns = ["TIPO_BBDD", "TIPO_OBJETO_SUBFORM", "REPOSITORIO", "NOMBRE_OBJETO", "NUM_CAMBIOS_SCRIPT_1", "NUM_CAMBIOS_SCRIPT_2", "CHECK_OBJETO"])


    lista_key_temp = []
    for key in dicc_control_versiones_tipo_concepto.keys():
        df_export_xls.loc[df_export_xls["CHECK_OBJETO"] == dicc_control_versiones_tipo_concepto[key], key] = 1
        df_export_xls.loc[df_export_xls["CHECK_OBJETO"] != dicc_control_versiones_tipo_concepto[key], key] = 0

        lista_key_temp.append(key)

    bbdd_1 = None
    bbdd_2 = None
    if tipo_bbdd == "MS_ACCESS":
        bbdd_1 = os.path.basename(dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["PATH_BBDD"])
        bbdd_2 = os.path.basename(dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"])

    elif tipo_bbdd == "SQL_SERVER":
        bbdd_1 = "[" + dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["SERVIDOR"] + "] " + dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["BBDD"]
        bbdd_2 = "[" + dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["SERVIDOR"] + "] " + dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["BBDD"]

    df_export_xls[["BBDD_01", "BBDD_02"]] = [bbdd_1, bbdd_2]


    lista_headers_temp = ["TIPO_BBDD", "TIPO_OBJETO_SUBFORM", "REPOSITORIO", "NOMBRE_OBJETO", lista_key_temp[0], lista_key_temp[1], lista_key_temp[2], 
                          "NUM_CAMBIOS_SCRIPT_1", "NUM_CAMBIOS_SCRIPT_2", "BBDD_01", "BBDD_02"]

    df_export_xls = df_export_xls[lista_headers_temp]
    df_export_xls = df_export_xls[lista_headers_temp].sort_values(["TIPO_BBDD", "TIPO_OBJETO_SUBFORM", "REPOSITORIO", "NOMBRE_OBJETO"], ascending = [True, True, True, True])



    #se exporta a excel
    wb = xw.Book(ruta_plantilla_control_versiones_xls, update_links = False)

    ws1 = wb.sheets["DETALLE_OBJETOS"]
    ws1.range("A3:FF65000").clear_contents()
    ws1["A3"].options(pd.DataFrame, header = 0, index = False, expand = "table").value = df_export_xls

    wb.save(saveas)


    del lista_control_versiones_selecc
    del df_export_xls
    del lista_headers_temp



def def_proceso_merge_realizar_cambios(tipo_accion, tipo_objeto_combobox_selecc, tipo_objeto_subform, repositorio_subform, nombre_objeto_subform, lineas_origen, lineas_destino):
    #permite calcular los nuevos scripts tras el merge con los tags correspondientes para las acciones de migrar y quitar
    #permite tambiene calcular el diccionario de merge realizados para poder incluirlo en LISTA_DICC_MERGE_HECHOS 
    #de la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) de dicc_control_versiones_tipo_objeto



    #se crean variables de tipo de accion para relacionarlas con el parametro tipo_accion de la rutina
    accion_migrar_todo = lista_GUI_proceso_merge_tipo_accion_bbdd_1[0]
    accion_migrar_lineas = lista_GUI_proceso_merge_tipo_accion_bbdd_1[1]
    accion_quitar_todo = lista_GUI_proceso_merge_tipo_accion_bbdd_2[0]
    accion_quitar_lineas = lista_GUI_proceso_merge_tipo_accion_bbdd_2[1]
    accion_revertir_todo = lista_GUI_proceso_merge_tipo_accion_bbdd_1[2]
    accion_revertir_ultimo_cambio = lista_GUI_proceso_merge_tipo_accion_bbdd_1[3]


    #se recuperan la key_1 y subkey_3 (asociada a la subkey_2 TIPO_OBJETO) asociadas al valor de la subkey_4 (COMBOBOX_CONTROL_VERSIONES) tomado por tipo_objeto_combobox_selecc
    #se recupera el indice del diccionario en la lista de diccionarios de la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) realcionada con la key_1 y subkey_3 mencionadas
    tipo_bbdd = func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_BBDD", valor = tipo_objeto_combobox_selecc)
    tipo_objeto_key = func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO", valor = tipo_objeto_combobox_selecc)


    #se localizan datos necesartios para el calculo del proceso de la rutina mediante la funcion func_control_versiones_dicc_proceso_merge_anteriores
    dicc_localiz_objeto = func_control_versiones_dicc_proceso_merge_anteriores(tipo_objeto_combobox_selecc, tipo_objeto_subform, repositorio_subform, nombre_objeto_subform)

    indice_dicc_control_versiones = dicc_localiz_objeto["INDICE_DICC_LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]
    lista_merge_hechos = dicc_localiz_objeto["LISTA_DICC_MERGE_HECHOS"]
    df_codigo_original_1 = dicc_localiz_objeto["DF_CODIGO_ORIGINAL_1"] if isinstance(dicc_localiz_objeto["DF_CODIGO_ORIGINAL_1"], pd.DataFrame) else pd.DataFrame(columns = lista_headers_df_codigo_control_versiones_2)
    df_codigo_original_2 = dicc_localiz_objeto["DF_CODIGO_ORIGINAL_2"] if isinstance(dicc_localiz_objeto["DF_CODIGO_ORIGINAL_2"], pd.DataFrame) else pd.DataFrame(columns = lista_headers_df_codigo_control_versiones_2)
    df_codigo_actual_1 = dicc_localiz_objeto["DF_CODIGO_ACTUAL_1"] if isinstance(dicc_localiz_objeto["DF_CODIGO_ACTUAL_1"], pd.DataFrame) else pd.DataFrame(columns = lista_headers_df_codigo_control_versiones_2)
    df_codigo_actual_2 = dicc_localiz_objeto["DF_CODIGO_ACTUAL_2"] if isinstance(dicc_localiz_objeto["DF_CODIGO_ACTUAL_2"], pd.DataFrame) else pd.DataFrame(columns = lista_headers_df_codigo_control_versiones_2)


    ##########################################################################################################################################
    ##########################################################################################################################################
    #            ACCIONES MIGRAR / QUITAR
    ##########################################################################################################################################
    ##########################################################################################################################################

    if tipo_accion in lista_acciones_migrar_quitar:

        ##############################################
        #        Migrar todo
        ##############################################

        if tipo_accion == accion_migrar_todo:

            #se conserva el script eliminado en pantalla al que se le suma el nuevo 
            # (aqui no hace falta recalcular los numeros de linea pq ya estan en los df)
            df_temp_agregado = df_codigo_actual_1.copy()
            df_temp_eliminado = df_codigo_actual_2.copy() if isinstance(df_codigo_actual_2, pd.DataFrame) else pd.DataFrame(columns = [i for i in df_codigo_actual_2])

            df_temp_agregado["CONTROL_CAMBIOS_ACTUAL"] = "AGREGADO"
            df_temp_eliminado["CONTROL_CAMBIOS_ACTUAL"] = "ELIMINADO"

            df_codigo_actual_despues_cambio_1 = df_temp_agregado

            df_codigo_actual_despues_cambio_2 = pd.concat([df_temp_agregado, df_temp_eliminado])
            df_codigo_actual_despues_cambio_2.reset_index(drop = True, inplace = True)


        ##############################################
        #        Migrar por lineas
        ##############################################

        elif tipo_accion == accion_migrar_lineas:

            df_codigo_actual_1_temp = df_codigo_actual_1.copy()
            df_codigo_actual_2_temp = df_codigo_actual_2.copy() if isinstance(df_codigo_actual_2, pd.DataFrame) else pd.DataFrame(columns = [i for i in df_codigo_actual_2])

            #se pone CONTROL_CAMBIOS a "AGREGADO" para resaltar el cambio en otro color en los scripts de no_merge y merge
            lista_temp = lineas_origen.split("-")

            if len(lista_temp) == 1:
                df_codigo_actual_1_temp["CONTROL_CAMBIOS_ACTUAL"] = df_codigo_actual_1_temp.apply(lambda x: "AGREGADO" if int(x.name) + 1 == int(lista_temp[0]) else x["CONTROL_CAMBIOS_ACTUAL"] , axis = 1)

            else:
                df_codigo_actual_1_temp["CONTROL_CAMBIOS_ACTUAL"] = (df_codigo_actual_1_temp.apply(lambda x: "AGREGADO" if int(x.name) + 1 >= int(lista_temp[0]) and int(x.name) + 1 <= int(lista_temp[1]) 
                                                                                        else x["CONTROL_CAMBIOS_ACTUAL"], axis = 1))

            #se reconstruye el df de la bbdd de merge --> se fragmenta 3 df que se concatenan en este orden
            # --> df_temp_1 = df_codigo_actual_2_temp <= lineas_destino
            # --> df_temp_2 = df_codigo_actual_1_temp entre lineas origen
            # --> df_temp_3 = df_codigo_actual_2_temp > lineas_destino

            df_temp_1 = df_codigo_actual_2_temp.loc[df_codigo_actual_2_temp.index + 1 <= int(lineas_destino), [i for i in df_codigo_actual_2_temp.columns]]

            if len(lista_temp) == 1:
                df_temp_2 = df_codigo_actual_1_temp.loc[df_codigo_actual_1_temp.index + 1 == int(lista_temp[0]), [i for i in df_codigo_actual_1_temp.columns]]
            else:
                df_temp_2 = df_codigo_actual_1_temp.loc[(df_codigo_actual_1_temp.index + 1 >= int(lista_temp[0])) & (df_codigo_actual_1_temp.index + 1 <= int(lista_temp[1])), [i for i in df_codigo_actual_1_temp.columns]]

            df_temp_2.reset_index(drop = True, inplace = True)

            df_temp_3 = df_codigo_actual_2_temp.loc[df_codigo_actual_2_temp.index + 1 > int(lineas_destino), [i for i in df_codigo_actual_2_temp.columns]]


            #se crean los df codigo despues de los cambios
            df_codigo_actual_despues_cambio_1 = df_codigo_actual_1_temp.copy()

            df_codigo_actual_despues_cambio_2 = pd.concat([df_temp_1, df_temp_2, df_temp_3])
            df_codigo_actual_despues_cambio_2.reset_index(drop = True, inplace = True)

            del df_temp_1
            del df_temp_2
            del df_temp_3
            


        ##############################################
        #        Quitar todo
        ##############################################

        elif tipo_accion == accion_quitar_todo:

            df_codigo_actual_despues_cambio_1 = df_codigo_actual_1.copy() if isinstance(df_codigo_actual_1, pd.DataFrame) else None

            df_codigo_actual_despues_cambio_2 = df_codigo_actual_2.copy()
            df_codigo_actual_despues_cambio_2["CONTROL_CAMBIOS_ACTUAL"] = "ELIMINADO"



        ##############################################
        #        Quitar por lineas
        ##############################################

        elif tipo_accion == accion_quitar_lineas:

            df_codigo_actual_despues_cambio_1 = df_codigo_actual_1.copy() if isinstance(df_codigo_actual_1, pd.DataFrame) else None
            df_codigo_actual_despues_cambio_2 = df_codigo_actual_2.copy()

            lista_temp = lineas_origen.split("-")
            if len(lista_temp) == 1:
                df_codigo_actual_despues_cambio_2["CONTROL_CAMBIOS_ACTUAL"] = df_codigo_actual_despues_cambio_2.apply(lambda x: "ELIMINADO" if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" and int(x["NUM_LINEA"][0:4]) == int(lista_temp[0]) 
                                                                                                                        else x["CONTROL_CAMBIOS_ACTUAL"]
                                                                                                                        , axis = 1)
            else:
                df_codigo_actual_despues_cambio_2["CONTROL_CAMBIOS_ACTUAL"] = (df_codigo_actual_despues_cambio_2.apply(lambda x: "ELIMINADO" if x["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO" and int(x["NUM_LINEA"][0:4]) >= int(lista_temp[0]) 
                                                                                                                        and int(x["NUM_LINEA"][0:4]) <= int(lista_temp[1])
                                                                                                                        else x["CONTROL_CAMBIOS_ACTUAL"], axis = 1))

            df_codigo_actual_despues_cambio_2.reset_index(drop = True, inplace = True)


        ##########################################################################################################################################
        #se actualiza LISTA_DICC_MERGE_HECHOS en la subkey_4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES) de dicc_control_versiones_tipo_objeto
        ##########################################################################################################################################

        #se crea el diccionario que almacena solo los df de codigos depues de los cambios 
        #y se agrega a la lista de merge hechos para actualizar despues dicc_control_versiones_tipo_objeto
        dicc_temp = {"DF_CODIGO_ACTUAL_1": df_codigo_actual_despues_cambio_1
                    , "DF_CODIGO_ACTUAL_2": df_codigo_actual_despues_cambio_2
                    }

        if isinstance(lista_merge_hechos, list):
            lista_merge_hechos.append(dicc_temp)
        else:
            lista_merge_hechos = [dicc_temp]


        dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["LISTA_DICC_MERGE_HECHOS"] = lista_merge_hechos 


    ##########################################################################################################################################
    ##########################################################################################################################################
    #            ACCIONES REVERTIR
    ##########################################################################################################################################
    ##########################################################################################################################################

    elif tipo_accion in lista_acciones_revertir:


        ##############################################
        #OPCION --> Revertir todo
        ##############################################

        if tipo_accion == accion_revertir_todo:

            dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["LISTA_DICC_MERGE_HECHOS"] = None 

            if isinstance(df_codigo_original_1, pd.DataFrame):
                df_codigo_original_1["CONTROL_CAMBIOS_ACTUAL"] = df_codigo_original_1["CONTROL_CAMBIOS_ORIGINAL"]
            
            dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_1"] = df_codigo_original_1


            if isinstance(df_codigo_original_2, pd.DataFrame):
                df_codigo_original_2["CONTROL_CAMBIOS_ACTUAL"] = df_codigo_original_2["CONTROL_CAMBIOS_ORIGINAL"]

            dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_2"] = df_codigo_original_2



        ##############################################
        #OPCION --> Revertir ultimo cambio
        ##############################################

        elif tipo_accion == accion_revertir_ultimo_cambio:

            if isinstance(lista_merge_hechos, list):

                #se elimina el ultimo item de lista_merge_hechos (los ultimos df con cambios)
                lista_merge_hechos.pop()
                lista_merge_hechos_tras_revertir_ultimo_cambio = lista_merge_hechos

                #si la lista despues de la reversion sigue sigue teniendo un len >= 1 (ha habido previamente al menos 1 merge)
                if len(lista_merge_hechos) != 0:

                    #se seleccionan los df codigos del ultimo item de lista_merge_hechos_tras_revertir_ultimo_cambio
                    df_temp_1 = lista_merge_hechos_tras_revertir_ultimo_cambio[-1]["DF_CODIGO_ACTUAL_1"]
                    df_temp_2 = lista_merge_hechos_tras_revertir_ultimo_cambio[-1]["DF_CODIGO_ACTUAL_2"]

                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["LISTA_DICC_MERGE_HECHOS"] = lista_merge_hechos_tras_revertir_ultimo_cambio 
                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_1"] = df_temp_1
                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_2"] = df_temp_2


                #si len = 0, solo hubo 1 merge tras la reversion ya no hay merge hechos por lo que la key LISTA_DICC_MERGE_HECHOS de subkey 4 (LISTA_DICC_OBJETOS_CONTROL_VERSIONES)
                #de dicc_control_versiones_tipo_objeto se re-establece a None
                else:
                    
                    df_temp_1 = df_codigo_original_1 if isinstance(df_codigo_original_1, pd.DataFrame) else None
                    df_temp_2 = df_codigo_original_2 if isinstance(df_codigo_original_2, pd.DataFrame) else None

                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["LISTA_DICC_MERGE_HECHOS"] = None 
                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_1"] = df_temp_1
                    dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto_key]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"][indice_dicc_control_versiones]["DF_CODIGO_ACTUAL_2"] = df_temp_2



###############################################################################################################################################################################################
###############################################################################################################################################################################################
###############################################################################################################################################################################################
##                               FUNCIONES + RUTINAS - GUI MERGE EN BBDD FISICAS
###############################################################################################################################################################################################
###############################################################################################################################################################################################
###############################################################################################################################################################################################


def def_merge_access_ajustes_manuales():
    #rutina que permite localizar todos los ajustes manuales a realizar (en caso de haberlos) en access para almacenarlos en
    #dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["LISTA_DICC_OBJETOS"]
    #se realizan localizando los ajustes manuales a realizar en BBDD_02 que estan en BBDD_01 pero no en BBDD_02

    path_bbdd_2 = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["PATH_BBDD"]

    #librerias DLL a activar manualmente en BBDD_02
    librerias_dll_bbdd_1 = dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["LISTA_LIBRERIAS_DLL"]
    librerias_dll_bbdd_2 = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["LISTA_LIBRERIAS_DLL"]

    lista_librerias_dll_ajuste_manual = []
    if isinstance(librerias_dll_bbdd_1, list) and isinstance(librerias_dll_bbdd_2, list):
        lista_librerias_dll_ajuste_manual = [dll for dll in librerias_dll_bbdd_1 if dll not in librerias_dll_bbdd_2]


    #se extraen del parametro mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"]
    #los modulos que no son estandares (formularios/reportes y/o userform) y se comprueba si existen en BBDD_02
    #en caso de que los modulos no existan se informa en los ajustes manuales

    df_temp = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"][["TIPO_MODULO", "NOMBRE_MODULO"]]
    df_temp.drop_duplicates(subset = [i for i in df_temp.columns], keep = "last", inplace = True)
    df_temp.reset_index(drop = True, inplace = True)
    lista_modulos_bbdd_merge = [[df_temp.iloc[ind, 0], df_temp.iloc[ind, 1]] for ind in df_temp.index]
    del df_temp


    tipo_objeto_subform_rutinas = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["RUTINAS_VBA"]["TIPO_OBJETO_SUBFORM"]
    tipo_objeto_subform_variables = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VARIABLES_VBA"]["TIPO_OBJETO_SUBFORM"]
 
    lista_modulos_ajuste_manual = []
    for dicc in dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"]:

        tipo_objeto_subform = dicc["TIPO_OBJETO_SUBFORM"]
        tipo_repositorio = dicc["TIPO_REPOSITORIO"]
        repositorio = dicc["REPOSITORIO"]


        #se realiza solo para objetos que pueden ser incluidos en modulos VBA (rutinas o variables publicas)
        if tipo_objeto_subform in [tipo_objeto_subform_rutinas, tipo_objeto_subform_variables]:

            if tipo_repositorio != "Estandar":

                cont = 0
                for tipo_repositorio_bbdd_merge, repositorio_bbdd_merge in lista_modulos_bbdd_merge:
                    if tipo_repositorio == tipo_repositorio_bbdd_merge and repositorio == repositorio_bbdd_merge:
                        cont = 0
                        break
                    else:
                        cont += 1

                if cont != 0:
                    lista_modulos_ajuste_manual.append([repositorio, tipo_repositorio])

    if len(lista_modulos_ajuste_manual) != 0:
        lista_modulos_ajuste_manual = [sublista for i, sublista in enumerate(lista_modulos_ajuste_manual) if sublista not in lista_modulos_ajuste_manual[:i]]
        lista_modulos_ajuste_manual = sorted(lista_modulos_ajuste_manual, key = lambda x: (x[0], x[1]))



    #se informa dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["LISTA_DICC_OBJETOS"]
    #con un df que fusiona losajustes a realizar en las librerias DLL y los ajustes en los formularios
    lista_para_df = []
    df_temp = None
    if len(lista_librerias_dll_ajuste_manual) + len(lista_modulos_ajuste_manual) != 0:

        if len(lista_librerias_dll_ajuste_manual) != 0:
            lista_para_df = lista_para_df + ["AJUSTES MANUALES", path_bbdd_2, "", "LIBRERIAS DLL POR ACTIVAR:"] + ["--> " + dll for dll in lista_librerias_dll_ajuste_manual] + ["", ""]

        if len(lista_modulos_ajuste_manual) != 0:
            lista_para_df = lista_para_df + ["MODULOS POR CREAR:"] + ["--> " + item[0] + "(" + item[1] + ")" for item in lista_modulos_ajuste_manual]

        df_temp = pd.DataFrame({"CODIGO": lista_para_df})
            
        #se crea el df con el mismo formato de columnas que los df de los objetos a migrar en bbdd fisica
        #y se almacena en una lista de diccionarios (la lista solo tiene 1 diccionario)
        #es para que funcione la rutina de actualizacion del subform en la GUI sin realizar igual con todos los tipos de objetos seleccionados
        #sin necesidad de realizar un ajuste espcial en codigo para el caso los ajustes manuales
        df_temp[["NUM_LINEA", "CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]] = [None, None, None, None]
        df_temp = df_temp[["NUM_LINEA", "CODIGO", "CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ORIGINAL", "CONTROL_CAMBIOS_ACTUAL"]]

    lista_dicc_ajustes_manuales = [
                                    {"TIPO_BBDD": "MS_ACCESS"
                                    , "TIPO_OBJETO_SUBFORM": "---"
                                    , "TIPO_REPOSITORIO": "---"
                                    , "REPOSITORIO": "---"
                                    , "NOMBRE_OBJETO": "---"
                                    , "ESTADO_MIGRACION": label_merge_access_bbdd_fisica_en_manual
                                    , "DF_CODIGO": df_temp
                                    }
                                ]

    dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["AJUSTES_MANUALES"]["LISTA_DICC_OBJETOS"] = lista_dicc_ajustes_manuales

    del lista_dicc_ajustes_manuales
    del df_temp
    del lista_librerias_dll_ajuste_manual
    del lista_modulos_ajuste_manual




def def_merge_bbdd_fisica_lista_objetos():
    #permite crear la lista de objetos (es lista de diccionarios) donde se han realizado cambios para poder replicarlos en bbdd fisica
    #la lista creada se informa en dicc_control_versiones_tipo_objeto[tipo_bbdd]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"]


    lista_objetos_access = []
    lista_objetos_sql_server= []
    for tipo_bbdd in dicc_control_versiones_tipo_objeto.keys():

        for tipo_objeto in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"].keys():

            if tipo_objeto != "TODOS":

                if isinstance(dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"], list):

                    for dicc in dicc_control_versiones_tipo_objeto[tipo_bbdd]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]:

                        if isinstance(dicc["LISTA_DICC_MERGE_HECHOS"], list):

                            #dentro de la lista dicc["LISTA_DICC_MERGE_HECHOS"] se recupera el df asociado a la key DF_CODIGO_ACTUAL_2 del ultimo diccionario
                            #es DF_CODIGO_ACTUAL_2 pq por defecto la BBD_02 es donde se realiza el merge
                            df_codigo = dicc["LISTA_DICC_MERGE_HECHOS"][-1]["DF_CODIGO_ACTUAL_2"]
                            
                            dicc_temp = {"TIPO_BBDD": dicc["TIPO_BBDD"]
                                        , "CHECK_OBJETO": dicc["CHECK_OBJETO"]
                                        , "TIPO_OBJETO_SUBFORM": dicc["TIPO_OBJETO_SUBFORM"]
                                        , "TIPO_REPOSITORIO": dicc["TIPO_REPOSITORIO"]
                                        , "REPOSITORIO": dicc["REPOSITORIO"]
                                        , "NOMBRE_OBJETO": dicc["NOMBRE_OBJETO"]
                                        , "ESTADO_MIGRACION": "Pendiente migrar"
                                        , "DF_CODIGO": df_codigo
                                        }


                            if tipo_bbdd == "MS_ACCESS":
                                if dicc_temp not in lista_objetos_access:
                                    lista_objetos_access.append(dicc_temp)



                            if tipo_bbdd == "SQL_SERVER":
                                tipo_objeto_subform = dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"][tipo_objeto]["TIPO_OBJETO_SUBFORM"]

                                if tipo_objeto_subform == dicc["TIPO_OBJETO_SUBFORM"]:
                                    if dicc_temp not in lista_objetos_sql_server:
                                        lista_objetos_sql_server.append(dicc_temp)


    #se informa dicc_control_versiones_tipo_objeto
    dicc_control_versiones_tipo_objeto["MS_ACCESS"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"] = lista_objetos_access
    dicc_control_versiones_tipo_objeto["SQL_SERVER"]["MERGE_BBDD_FISICA"]["OBJETOS_A_MIGRAR"]["LISTA_DICC_OBJETOS"] = lista_objetos_sql_server





def def_merge_bbdd_fisicas(tipo_bbdd_selecc, lista_dicc_objetos_migrar_bbdd_fisica, ruta_export):
    #rutina que permite realizar el merge en bbdd fisica y documentar el proceso mediante ficheros .txt en la 
    #ruta indicada en el parametro de la rutina


    #es para agregar un flag a la subcarpeta que se crea al iniciar el proceso de documentacion
    now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]


    ####################################################################################################################################
    ####################################################################################################################################
    #           MS ACCESS
    ####################################################################################################################################
    ####################################################################################################################################

    if tipo_bbdd_selecc == "MS_ACCESS":

        tipo_objeto_subform_tablas = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["TABLA_LOCAL"]["TIPO_OBJETO_SUBFORM"]
        tipo_objeto_subform_vinculos_odbc = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VINCULO_ODBC"]["TIPO_OBJETO_SUBFORM"]
        tipo_objeto_subform_vinculos_otros = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VINCULO_OTRO"]["TIPO_OBJETO_SUBFORM"]
        tipo_objeto_subform_variables_publicas = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VARIABLES_VBA"]["TIPO_OBJETO_SUBFORM"]
        tipo_objeto_subform_rutinas = dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["RUTINAS_VBA"]["TIPO_OBJETO_SUBFORM"]

        texto_check_objeto_solo_bbbd_01 = dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_01"]


        ####################################################################################################################################
        # CALCULOS --> TABLAS / VINCULOS
        ####################################################################################################################################

        #se crea la lista de los modulos donde se han realizado cambios en tablas / vinculos
        lista_dicc_tablas_y_vinculos_con_cambios = []
        for dicc in lista_dicc_objetos_migrar_bbdd_fisica:
            if dicc["TIPO_OBJETO_SUBFORM"] in [tipo_objeto_subform_tablas, tipo_objeto_subform_vinculos_odbc, tipo_objeto_subform_vinculos_otros]:

                tipo_objeto = func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = dicc["TIPO_OBJETO_SUBFORM"])
                nombre_objeto = dicc["NOMBRE_OBJETO"]
                df_codigo = dicc["DF_CODIGO"]


                #se prepara el codigo para documentarlo en fichero .txt (se le agrega al inicio de cada linea ADD o DEL segun se haya
                #agregado o eliminado la linea respectivamente)
                df_codigo_documentacion = df_codigo[["CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ACTUAL"]].copy()

                df_codigo_documentacion["ADD_DEL"] = (df_codigo_documentacion.apply(lambda x: "ADD\t" if x["CONTROL_CAMBIOS_ACTUAL"] == "AGREGADO" else "DEL\t" 
                                                                                    if x["CONTROL_CAMBIOS_ACTUAL"] == "ELIMINADO" else "---\t", 
                                                                                    axis = 1))
                

                df_codigo_documentacion["CODIGO"] = df_codigo_documentacion[["ADD_DEL", "CODIGO_CON_NUM_LINEA"]].astype(str).apply("".join, axis = 1)
                df_codigo_documentacion = df_codigo_documentacion[["CODIGO"]]


                #se prepara la sentencia create (quitando el flag ELIMINADO)
                df_codigo_sin_flag_eliminado = df_codigo.loc[df_codigo["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO", ["CODIGO"]]
                df_codigo_sin_flag_eliminado.reset_index(drop = True, inplace = True)


                #se crean las sentencias para poder realizar el merge en bbdd fisica (en los 3 casos TABLA_LOCAL, VINCULO_ODBC y VINCULO_OTRO
                #se crea la sentencia_create y el objeto de origen (solo aplica para vinculos),se hace asi para realizar el merge en bbdd fisica por bucle (ver mas abajo)
                #si df_codigo_sin_flag_eliminado devuelve un df vacio se estable a None                           
                if tipo_objeto == "TABLA_LOCAL":

                    sentencia_create = "\n".join([df_codigo_sin_flag_eliminado.iloc[ind, 0] for ind in df_codigo_sin_flag_eliminado.index]) if len(df_codigo_sin_flag_eliminado) != 0 else None
                    objeto_origen = None


                elif tipo_objeto == "VINCULO_ODBC":

                    sentencia_create = "\n".join([df_codigo_sin_flag_eliminado.iloc[ind, 0] for ind in df_codigo_sin_flag_eliminado.index
                                                  if len(df_codigo_sin_flag_eliminado.iloc[ind, 0]) != 0 and 
                                                  df_codigo_sin_flag_eliminado.iloc[ind, 0][0:6].upper() != "SOURCE"]).replace("\n", ";") if len(df_codigo_sin_flag_eliminado) != 0 else None
                    
                    objeto_origen = df_codigo_sin_flag_eliminado.iloc[len(df_codigo_sin_flag_eliminado) - 1, 0].replace("SOURCE: ", "").strip() if len(df_codigo_sin_flag_eliminado) != 0 else None


                elif tipo_objeto == "VINCULO_OTRO":
                    sentencia_create = "\n".join([df_codigo_sin_flag_eliminado.iloc[ind, 0]
                                                  for ind in df_codigo_sin_flag_eliminado.index if len(df_codigo_sin_flag_eliminado.iloc[ind, 0]) != 0 and 
                                                  df_codigo_sin_flag_eliminado.iloc[ind, 0][0:6].upper() != "SOURCE"]).replace("\n", ";") if len(df_codigo_sin_flag_eliminado) != 0 else None
                    
                    objeto_origen = df_codigo_sin_flag_eliminado.iloc[len(df_codigo_sin_flag_eliminado) - 1, 0].replace("SOURCE: ", "").strip() if len(df_codigo_sin_flag_eliminado) != 0 else None


                #se crea un diccionario temporal y se agrega a la lista lista_dicc_tablas_y_vinculos_con_cambios
                dicc_temp = {"TIPO_OBJETO": func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = dicc["TIPO_OBJETO_SUBFORM"])
                            , "NOMBRE_OBJETO": nombre_objeto
                            , "DF_CODIGO_DOCUMENTACION": df_codigo_documentacion

                            , "PARAMETROS_MERGE": {"SENTENCIA_CREATE": sentencia_create
                                                , "OBJETO_ORIGEN": objeto_origen
                                                }
                            }

                lista_dicc_tablas_y_vinculos_con_cambios.append(dicc_temp)

                del dicc_temp
                del df_codigo
                del df_codigo_documentacion


        ####################################################################################################################################
        # CALCULOS --> MODULOS VBA
        ####################################################################################################################################

        #se recupera el df de codigo de BBDD_02 (es por defecto la bbdd donde se hacer el merge en bbddd fisica)
        df_codigos_bbdd = dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]


        #se crea la lista de los modulos donde se han realizado cambios
        lista_modulos_con_cambios = [[dicc["TIPO_REPOSITORIO"], dicc["REPOSITORIO"]] 
                                     for dicc in lista_dicc_objetos_migrar_bbdd_fisica if dicc["TIPO_OBJETO_SUBFORM"] in [tipo_objeto_subform_variables_publicas, tipo_objeto_subform_rutinas]]
        
        lista_modulos_con_cambios = [sublista for i, sublista in enumerate(lista_modulos_con_cambios) if sublista not in lista_modulos_con_cambios[:i]]


        #se crea la lista lista_dicc_modulos_con_cambios que contiene diccionarios (cada uno por modulo con cambios)
        #donde se almacenan los datos necesarios para reconstruir el modulo con los cambios antes de exportarlo en la bbdd BBDD_02
        #(que es por defecto donde se hace el merge en bbdd fisica)
        #se usa tambien esta lista de diccionarios para almcenar los datos de cara a su documentación en ficheros .txt

        lista_dicc_modulos_con_cambios = []
        if len(lista_modulos_con_cambios) != 0:

            for tipo_modulo_con_cambios, nombre_modulo_con_cambios in lista_modulos_con_cambios:

                #se extrae de df_codigos_bbdd el codigo del modulo de la iteracion
                df_codigo_original_modulo = (df_codigos_bbdd.loc[(df_codigos_bbdd["TIPO_MODULO"] == tipo_modulo_con_cambios) & (df_codigos_bbdd["NOMBRE_MODULO"] == nombre_modulo_con_cambios),
                                                        [i for i in df_codigos_bbdd.columns]])

                df_codigo_original_modulo.reset_index(drop = True, inplace = True)


                ################################################################
                # ENCABEZADO DEL MODULO
                ################################################################

                #se extrae el encabezado de modulo al cual se le agrega una linea de codigo en blanco para tener separacion
                df_codigo_modulo_encabezado_reconstruido = df_codigo_original_modulo.loc[df_codigo_original_modulo["ES_ENCABEZADO_MODULO"] == "SI", ["CODIGO"]]

                df_codigo_modulo_encabezado_reconstruido = (pd.concat([df_codigo_modulo_encabezado_reconstruido, pd.DataFrame({"CODIGO": [""]})]) 
                                                                if len(df_codigo_modulo_encabezado_reconstruido) != 0 else pd.DataFrame({"CODIGO": [""]}))
                
                df_codigo_modulo_encabezado_reconstruido.reset_index(drop = True, inplace = True)


                ################################################################
                # VARIABLES PUBLICAS
                ################################################################

                #se extraen las variables publicas (antes de los cambios) a las cuales se les agrega una linea de codigo 
                #en blanco para tener separacion
                df_codigo_modulo_variables_publicas_antes_cambios = df_codigo_original_modulo.loc[df_codigo_original_modulo["ES_VARIABLE_PUBLICA"] != "NO", ["CODIGO"]]           
                df_codigo_modulo_variables_publicas_antes_cambios = pd.concat([df_codigo_modulo_variables_publicas_antes_cambios, pd.DataFrame({"CODIGO": [""]})])
                df_codigo_modulo_variables_publicas_antes_cambios.reset_index(drop = True, inplace = True)


                #se extraen las variables publicas (con los cambios, si los hay) a las cuales se les agrega una linea de codigo 
                #en blanco para tener separacion
                lista_df_variables_publicas_con_cambios = [dicc["DF_CODIGO"] for dicc in lista_dicc_objetos_migrar_bbdd_fisica 
                                                            if dicc["TIPO_REPOSITORIO"] == tipo_modulo_con_cambios and dicc["REPOSITORIO"] == nombre_modulo_con_cambios and
                                                            dicc["TIPO_OBJETO_SUBFORM"] == tipo_objeto_subform_variables_publicas
                                                            ]

                #se crea el df df_codigo_modulo_variables_publicas_para_reconstruccion donde se quitan las lineas eliminadas en el script de BBDD_02
                #se crea tambien el df df_codigo_modulo_variables_publicas_para_documentacion de codigo para la documentacion 
                #en fichero .txt donde se agrega al inicio de cada linea ADD o DEL segun se haya agregado o eliminado la linea respectivamente
                df_codigo_modulo_variables_publicas_para_reconstruccion = None
                df_codigo_modulo_variables_publicas_para_documentacion = None

                if len(lista_df_variables_publicas_con_cambios) != 0:

                    df_codigo_modulo_variables_publicas_con_cambios = lista_df_variables_publicas_con_cambios[0]

                    if isinstance(df_codigo_modulo_variables_publicas_con_cambios, pd.DataFrame):

                        #se crea el df df_codigo_modulo_variables_publicas_para_reconstruccion donde se quitan las lineas de codigo ELIMINADAS
                        df_codigo_modulo_variables_publicas_para_reconstruccion = (df_codigo_modulo_variables_publicas_con_cambios.loc[df_codigo_modulo_variables_publicas_con_cambios["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO", 
                                                                                                                               ["CODIGO"]])

                        df_codigo_modulo_variables_publicas_para_reconstruccion = pd.concat([df_codigo_modulo_variables_publicas_para_reconstruccion, pd.DataFrame({"CODIGO": [""]})])
                        df_codigo_modulo_variables_publicas_para_reconstruccion.reset_index(drop = True, inplace = True)


                        #se crea el df df_codigo_modulo_variables_publicas_para_documentacion para la documentacion
                        df_codigo_modulo_variables_publicas_para_documentacion = df_codigo_modulo_variables_publicas_con_cambios[["CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ACTUAL"]].copy()

                        df_codigo_modulo_variables_publicas_para_documentacion["ADD_DEL"] = (df_codigo_modulo_variables_publicas_para_documentacion.apply(lambda x: "ADD\t" 
                                                                                            if x["CONTROL_CAMBIOS_ACTUAL"] == "AGREGADO" else "DEL\t" 
                                                                                            if x["CONTROL_CAMBIOS_ACTUAL"] == "ELIMINADO" else "---\t", 
                                                                                            axis = 1))
                        
                        df_codigo_modulo_variables_publicas_para_documentacion["CODIGO"] = df_codigo_modulo_variables_publicas_para_documentacion[["ADD_DEL", "CODIGO_CON_NUM_LINEA"]].astype(str).apply("".join, axis = 1)
                        df_codigo_modulo_variables_publicas_para_documentacion = df_codigo_modulo_variables_publicas_para_documentacion[["CODIGO"]]

                    del df_codigo_modulo_variables_publicas_con_cambios


                #se crea el df reconstruido para las variables publicas (en funcion de si se han modificado o no variables publicas)
                #los cambios por el usuario pueden ser solo a nivel de rutinas / funciones por lo que en este caso es el codigo original
                df_codigo_modulo_variables_publicas_reconstruido = (df_codigo_modulo_variables_publicas_para_reconstruccion 
                                                                    if isinstance(df_codigo_modulo_variables_publicas_para_reconstruccion, pd.DataFrame) 
                                                                    else df_codigo_modulo_variables_publicas_antes_cambios)



                ################################################################
                # RUTINAS / FUNCIONES
                ################################################################

                #se crea la lista lista_rutinas_para_reconstruccion_modulo de las rutinas / funciones dentro del modulo de la iteracion
                #se recuperan todas las rutinas / funciones del modulo de BBDD_02 independientemente de que el usuario haya realizado cambios en algunas o no
                #para obtener en que orden aparecen en el modulo y mediante bucle sobre esta lista de rutinas / funciones en cada iteracion se localizan las que
                #tienen cambios realizados por el usuario (si los tienen se conserva el codigo de la rutina con cambios, quitando las lineas donde el flag pone ELIMINADO
                #en caso contrario se conserva el codigo original), se informa tambien de la accion a realizar (si el df de codigo de la rutina / funcion tras quitar las lineas
                #con el flag ELIMINADO devuelve un df vacio no se agrega la rutina / funcion en la reconstruccion del modulo para el merge en bbdd fisica)
                #
                #se agrega tambien el df con cambios preparado para su exportacion en fichero .txt donde en cada linea al principio se agrega ADD o DEL
                #segun que la linea se haya agregado o eliminado respectivamente
                #
                #al final del proceso se agregan a esta lista las rutinas que no estaban en BBDD_02 pero que el usuario ha migrado desde la BBDD_01
                #estas rutinas dentro del modulo reconstruido se agregan al final del modulo para conservar mas o menos el mismo orden del modulo original antes de los cambios
                #
                #es lista de lista donde cada sublista contiene:
                # --> el numero de orden de la rutina / funcion dentro del modulo
                # --> el nombre de la rutina / funcion del modulo
                # --> el codigo de la rutina con los cambios (si los hay y sino el codigo original)
                # --> el tipo de accion a realizar (MANTENER o ELIMINAR)
                #
                #la lista se re-ordena por el numero de orden de la rutina dentro del modulo


                #indices de las sublistas (es para mayor legibilidad del codigo)
                indice_lista_numero_orden_rutina = 0
                indice_lista_nombre_rutina = 1
                indice_lista_df_codigo_rutina_reconstruida = 2
                indice_lista_df_codigo_rutina_documentacion = 3
                indice_lista_accion_rutina_a_realizar = 4


                #se extrae el codigo original de BBDD_02 de las rutinas del modulo de la iteracion con elnumero de orden dentro del modulo
                df_codigo_original_modulo_rutinas = df_codigo_original_modulo.loc[df_codigo_original_modulo["NOMBRE_RUTINA"].isnull() == False, ["NOMBRE_RUTINA", "ORDEN_RUTINA_EN_MODULO", "CODIGO"]]
                df_codigo_original_modulo_rutinas.reset_index(drop = True, inplace = True)


                lista_rutinas_para_reconstruccion_modulo = []
                if len(df_codigo_original_modulo_rutinas) != 0:

                    lista_rutinas_para_reconstruccion_modulo = [[df_codigo_original_modulo_rutinas.iloc[ind, df_codigo_original_modulo_rutinas.columns.get_loc("ORDEN_RUTINA_EN_MODULO")]
                                            , df_codigo_original_modulo_rutinas.iloc[ind, df_codigo_original_modulo_rutinas.columns.get_loc("NOMBRE_RUTINA")]
                                            , None #es para el df de codigo con cambios (se calcula mas adelante)
                                            , None #es para el df de codigo para documentar (se calcula mas adelante)
                                            , None #es para el tipo de accion a realizar (se calcula mas adelante)
                                            ] for ind in df_codigo_original_modulo_rutinas.index]


                    #se quitan los duplicados y se ordena por el numero de orden de la rutina dentro del modulo
                    lista_rutinas_para_reconstruccion_modulo = [sublista for i, sublista in enumerate(lista_rutinas_para_reconstruccion_modulo) if sublista not in lista_rutinas_para_reconstruccion_modulo[:i]]
                    lista_rutinas_para_reconstruccion_modulo = sorted(lista_rutinas_para_reconstruccion_modulo, key = lambda x: x[0])


                    #se informa en lista_rutinas_para_reconstruccion_modulo el df de codigo original y el df con cambios (si lo hay)
                    for indice_lista_rutinas_modulo, item_lista_rutinas_modulo in enumerate(lista_rutinas_para_reconstruccion_modulo):

                        orden_rutina_en_modulo = item_lista_rutinas_modulo[indice_lista_numero_orden_rutina]
                        nombre_rutina = item_lista_rutinas_modulo[indice_lista_nombre_rutina]
               

                        #df codigo original al cual se le agrega una linea de codigo en blanco para tener separacion
                        df_codigo_original_rutina = df_codigo_original_modulo.loc[df_codigo_original_modulo["ORDEN_RUTINA_EN_MODULO"] == orden_rutina_en_modulo, ["CODIGO"]]
                        df_codigo_original_rutina = pd.concat([df_codigo_original_rutina, pd.DataFrame({"CODIGO": [""]})])
                        df_codigo_original_rutina.reset_index(drop = True, inplace = True)


                        #se crea el df df_codigo_rutina_para_reconstruccion de codigo con cambios (si los hay)
                        #se excluyen las lineas donde pone ELIMINADO en el control de cambios, si el df resultante 
                        #de quitar todas estas lineas devuelve un df vacio se informa la accion = ELIMINAR, sino es REEMPLAZAR
                        #se agrega a df_codigo_rutina_para_reconstruccion una linea de codigo en blanco para tener separacion
                        #
                        #se crea tambien el df df_codigo_rutina_para_documentacion de codigo para la documentacion 
                        #en fichero .txt donde se agrega al inicio de cada linea ADD o DEL segun se haya agregado o eliminado la linea respectivamente
                        lista_df_rutinas_con_cambios = [dicc["DF_CODIGO"] for dicc in lista_dicc_objetos_migrar_bbdd_fisica 
                                                        if dicc["TIPO_REPOSITORIO"] == tipo_modulo_con_cambios and dicc["REPOSITORIO"] == nombre_modulo_con_cambios and
                                                        dicc["TIPO_OBJETO_SUBFORM"] == tipo_objeto_subform_rutinas and dicc["NOMBRE_OBJETO"] == nombre_rutina
                                                        ]


                        accion_rutina_a_realizar = None
                        df_codigo_rutina_para_reconstruccion = None
                        df_codigo_rutina_para_documentacion = None

                        if len(lista_df_rutinas_con_cambios) != 0:

                            df_codigo_rutina_con_cambios = lista_df_rutinas_con_cambios[0]

                            if isinstance(df_codigo_rutina_con_cambios, pd.DataFrame):

                                #se crea el df df_codigo_rutina_para_reconstruccion
                                df_codigo_rutina_para_reconstruccion = df_codigo_rutina_con_cambios.loc[df_codigo_rutina_con_cambios["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO", ["CODIGO"]]
                                df_codigo_rutina_para_reconstruccion = pd.concat([df_codigo_rutina_para_reconstruccion, pd.DataFrame({"CODIGO": [""]})])
                                df_codigo_rutina_para_reconstruccion.reset_index(drop = True, inplace = True)


                                #se crea el df df_codigo_rutina_para_documentacion
                                df_codigo_rutina_para_documentacion = df_codigo_rutina_con_cambios[["CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ACTUAL"]].copy()

                                df_codigo_rutina_para_documentacion["ADD_DEL"] = (df_codigo_rutina_para_documentacion.apply(lambda x: "ADD\t" if x["CONTROL_CAMBIOS_ACTUAL"] == "AGREGADO" else "DEL\t" 
                                                                                                                    if x["CONTROL_CAMBIOS_ACTUAL"] == "ELIMINADO" else "---\t", 
                                                                                                                    axis = 1))
                                
                                df_codigo_rutina_para_documentacion["CODIGO"] = df_codigo_rutina_para_documentacion[["ADD_DEL", "CODIGO_CON_NUM_LINEA"]].astype(str).apply("".join, axis = 1)
                                df_codigo_rutina_para_documentacion = df_codigo_rutina_para_documentacion[["CODIGO"]]


                            #se informa de la accion a realizar
                            if isinstance(df_codigo_rutina_con_cambios, pd.DataFrame):
                                accion_rutina_a_realizar = "CON_CAMBIOS" if len(df_codigo_rutina_con_cambios) != 0 else "ELIMINAR"

                            del df_codigo_rutina_con_cambios


                        #se crea el df de codigo de rutina reconstruido
                        df_codigo_rutina_reconstruido = df_codigo_rutina_para_reconstruccion if isinstance(df_codigo_rutina_para_reconstruccion, pd.DataFrame) else df_codigo_original_rutina


                        #se informan los df y la accion a realizar en la sublista de lista_rutinas_para_reconstruccion_modulo
                        lista_rutinas_para_reconstruccion_modulo[indice_lista_rutinas_modulo][indice_lista_df_codigo_rutina_reconstruida] = df_codigo_rutina_reconstruido
                        lista_rutinas_para_reconstruccion_modulo[indice_lista_rutinas_modulo][indice_lista_df_codigo_rutina_documentacion] = df_codigo_rutina_para_documentacion
                        lista_rutinas_para_reconstruccion_modulo[indice_lista_rutinas_modulo][indice_lista_accion_rutina_a_realizar] = accion_rutina_a_realizar


                        del df_codigo_rutina_para_reconstruccion
                        del df_codigo_rutina_para_documentacion
                        del lista_df_rutinas_con_cambios



                #se localizan las rutinas / funciones nuevas que no estaban en BBDD_02 y que el usuario quiere importar de BBDD_01
                #y se agregan a la lista lista_rutinas_para_reconstruccion_modulo al final del todo (se agrega a cada nueva rutina una linea en blanco 
                #para tener sepreacion con la siguiente rutina)
                #se filtra en la lista lista_dicc_objetos_migrar_bbdd_fisica (parametro de la presente rutina)
                #por la key del de los diccionarios CHECK_OBJETO = SOLO_EN_BBDD_01
                #
                #se crea tambien el df df_codigo_nueva_rutina_documentacion de codigo para la documentacion 
                #en fichero .txt donde se agrega al inicio de cada linea ADD o DEL segun se haya agregado o eliminado la linea respectivamente
                lista_nuevas_rutinas = [[dicc["NOMBRE_OBJETO"], dicc["DF_CODIGO"]] for dicc in lista_dicc_objetos_migrar_bbdd_fisica 
                                        if dicc["TIPO_REPOSITORIO"] == tipo_modulo_con_cambios and dicc["REPOSITORIO"] == nombre_modulo_con_cambios and
                                        func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = dicc["TIPO_OBJETO_SUBFORM"]) == "RUTINAS_VBA" and
                                        dicc["CHECK_OBJETO"] == texto_check_objeto_solo_bbbd_01]
                
                if len(lista_nuevas_rutinas)!= 0:
                    for item_nuevas_rutinas in lista_nuevas_rutinas:
                        
                        numero_orden_en_modulo_nueva_rutina = len(lista_rutinas_para_reconstruccion_modulo) + 1
                        nombre_nueva_rutina = item_nuevas_rutinas[0]
                        df_codigo_nueva_rutina = item_nuevas_rutinas[1]


                        #se crea el df df_codigo_nueva_rutina_reconstruida para la reconstruccion del modulo
                        df_codigo_nueva_rutina_reconstruida = df_codigo_nueva_rutina.loc[df_codigo_nueva_rutina["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO", ["CODIGO"]]
                        df_codigo_nueva_rutina_reconstruida.reset_index(drop = True, inplace = True)

                        df_codigo_nueva_rutina_reconstruida = pd.concat([df_codigo_nueva_rutina_reconstruida, pd.DataFrame({"CODIGO": [""]})])


                        #se crea el df df_codigo_nueva_rutina_documentacion
                        df_codigo_nueva_rutina_documentacion = df_codigo_nueva_rutina[["CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ACTUAL"]].copy()

                        df_codigo_nueva_rutina_documentacion["ADD_DEL"] = (df_codigo_nueva_rutina_documentacion.apply(lambda x: "ADD\t" if x["CONTROL_CAMBIOS_ACTUAL"] == "AGREGADO" else "DEL\t" 
                                                                                                            if x["CONTROL_CAMBIOS_ACTUAL"] == "ELIMINADO" else "---\t", 
                                                                                                            axis = 1))
                        
                        df_codigo_nueva_rutina_documentacion["CODIGO"] = df_codigo_nueva_rutina_documentacion[["ADD_DEL", "CODIGO_CON_NUM_LINEA"]].astype(str).apply("".join, axis = 1)
                        df_codigo_nueva_rutina_documentacion = df_codigo_nueva_rutina_documentacion[["CODIGO"]]


                        #se agregan los datos a lista_rutinas_para_reconstruccion_modulo
                        lista_rutinas_para_reconstruccion_modulo.append([numero_orden_en_modulo_nueva_rutina
                                                                        , nombre_nueva_rutina
                                                                        , df_codigo_nueva_rutina_reconstruida
                                                                        , df_codigo_nueva_rutina_documentacion
                                                                        , "NEW"])


                #se reconstruye el codigo de las rutinas del modulo
                df_codigo_modulo_rutinas_reconstruido = pd.DataFrame(columns = ["CODIGO"])
                if len(lista_rutinas_para_reconstruccion_modulo):
                    lista_rutinas_para_reconstruccion_modulo = sorted(lista_rutinas_para_reconstruccion_modulo, key = lambda x: x[0])

                    for numero_orden_rutina, nombre_rutina, df_codigo_para_reconstrucciion_modulo, df_codigo_para_documentacion, accion_a_realizar in lista_rutinas_para_reconstruccion_modulo:

                        if accion_a_realizar != "ELIMINAR":
                            df_codigo_modulo_rutinas_reconstruido = pd.concat([df_codigo_modulo_rutinas_reconstruido, df_codigo_para_reconstrucciion_modulo])



                #se crea el df de codigo de modulo reconstruido concatenando:
                # --> el encabezado de modulo
                # --> las variables publicas
                # --> las rutinas / funciones
                df_codigo_modulo_reconstruido = pd.concat([df_codigo_modulo_encabezado_reconstruido, df_codigo_modulo_variables_publicas_reconstruido, 
                                                           df_codigo_modulo_rutinas_reconstruido])
                
                df_codigo_modulo_reconstruido.reset_index(drop = True, inplace = True)


                ################################################################
                # se agregan los datos a lista_dicc_modulos_con_cambios
                ################################################################

                #se crea previamente la lista de las rutinas para la documentacion como extraccion de la lista lista_rutinas_para_reconstruccion_modulo
                #donde se extrae el nombre dela rutina / funcion y el df para documentacion (se hace solo para rutinas / funciones con cambios)
                lista_rutinas_modulo_para_documentacion = [[item[1], item[3]] for item in lista_rutinas_para_reconstruccion_modulo if item[4] != None]


                #se crea diccionario temporal que se agrega a lista_dicc_modulos_con_cambios
                dicc_temp = {"TIPO_REPOSITORIO": tipo_modulo_con_cambios
                            , "NOMBRE_REPOSITORIO": nombre_modulo_con_cambios
                            , "DF_VARIABLES_PUBLICAS_DOCUMENTACION": df_codigo_modulo_variables_publicas_para_documentacion
                            , "LISTA_DF_RUTINAS_DOCUMENTACION": lista_rutinas_modulo_para_documentacion
                            , "DF_CODIGO_REPOSITORIO": df_codigo_modulo_reconstruido
                            }

                lista_dicc_modulos_con_cambios.append(dicc_temp)


                del dicc_temp
                del df_codigo_modulo_variables_publicas_para_reconstruccion
                del df_codigo_modulo_variables_publicas_para_documentacion
                del lista_df_variables_publicas_con_cambios
                del df_codigo_modulo_variables_publicas_antes_cambios


        ####################################################################################################################################
        # DOCUMENTACION
        ####################################################################################################################################

        #se crea el directorio dentro de ruta_export
        nombre_carpeta_documentacion = tipo_bbdd_selecc + "_MIGRACION_" + now  
        ruta_carpeta_documentacion = str(ruta_export) + r'\%s' % nombre_carpeta_documentacion
        os.makedirs(ruta_carpeta_documentacion, exist_ok = True)


        #documentacion para tablas / vinculos
        if len(lista_dicc_tablas_y_vinculos_con_cambios) != 0:

            nombre_carpeta_documentacion_tablas_y_vinculos = "TABLAS_Y_VINCULOS" 
            ruta_carpeta_documentacion_tablas_y_vinculos = str(ruta_carpeta_documentacion) + r'\%s' % nombre_carpeta_documentacion_tablas_y_vinculos
            os.makedirs(ruta_carpeta_documentacion_tablas_y_vinculos, exist_ok = True)

            for dicc_tablas_y_vinculos in lista_dicc_tablas_y_vinculos_con_cambios:

                tipo_objeto = dicc_tablas_y_vinculos["TIPO_OBJETO"]
                nombre_objeto = dicc_tablas_y_vinculos["NOMBRE_OBJETO"]
                df_codigo_documentacion = dicc_tablas_y_vinculos["DF_CODIGO_DOCUMENTACION"]

                nombre_fichero_tablas_y_vinculos = "[" + tipo_objeto + "]_" + nombre_objeto + ".txt"
                ruta_fichero_tablas_y_vinculos = str(ruta_carpeta_documentacion_tablas_y_vinculos) + r'\%s' % nombre_fichero_tablas_y_vinculos

                with open(ruta_fichero_tablas_y_vinculos, 'w') as fich:
                    for ind in df_codigo_documentacion.index:
                        fich.write(df_codigo_documentacion.iloc[ind, 0] + "\n")



        #documentacion para modulos VBA
        if len(lista_dicc_modulos_con_cambios) != 0:

            nombre_carpeta_documentacion_modulos = "MODULOS_VBA" 
            ruta_carpeta_documentacion_modulos = str(ruta_carpeta_documentacion) + r'\%s' % nombre_carpeta_documentacion_modulos
            os.makedirs(ruta_carpeta_documentacion_modulos, exist_ok = True)

            for dicc_modulos in lista_dicc_modulos_con_cambios:

                tipo_repositorio = dicc_modulos["TIPO_REPOSITORIO"]
                nombre_repositorio = dicc_modulos["NOMBRE_REPOSITORIO"]
                df_codigo_modulo_documentacion = dicc_modulos["DF_CODIGO_REPOSITORIO"]
                lista_rutinas_documentacion = dicc_modulos["LISTA_DF_RUTINAS_DOCUMENTACION"]


                #se crea la carpeta del modulo de la iteracion
                nombre_carpeta_documentacion_modulos_iteracion = tipo_repositorio + "_" + nombre_repositorio
                ruta_carpeta_documentacion_modulos_iteracion = str(ruta_carpeta_documentacion_modulos) + r'\%s' % nombre_carpeta_documentacion_modulos_iteracion
                os.makedirs(ruta_carpeta_documentacion_modulos_iteracion, exist_ok = True)
            

                #se documenta el codigo del modulo reconstruido
                nombre_fichero_modulo_reconstruido_iteracion = "[" + tipo_repositorio + "_" + nombre_repositorio + "]_CODIGO_MODULO_VBA.txt"
                ruta_fichero_modulo_reconstruido_iteracion  = str(ruta_carpeta_documentacion_modulos_iteracion) + r'\%s' % nombre_fichero_modulo_reconstruido_iteracion

                with open(ruta_fichero_modulo_reconstruido_iteracion, 'w') as fich:
                    for ind in df_codigo_modulo_documentacion.index:
                        fich.write(df_codigo_modulo_documentacion.iloc[ind, 0] + "\n")


                #se documenta el codigo de las rutinas / funciones con cambios
                for nombre_rutina_documentacion, df_codigo_rutina_documentacion in lista_rutinas_documentacion:

                    nombre_fichero_rutina_iteracion = nombre_rutina_documentacion + ".txt"
                    ruta_fichero_rutina_iteracion  = str(ruta_carpeta_documentacion_modulos_iteracion) + r'\%s' % nombre_fichero_rutina_iteracion

                    with open(ruta_fichero_rutina_iteracion, 'w') as fich:
                        for ind in df_codigo_rutina_documentacion.index:
                            fich.write(df_codigo_rutina_documentacion.iloc[ind, 0] + "\n")


        ####################################################################################################################################
        # MERGE EN BBDD FISICA
        ####################################################################################################################################

        try:
            subprocess.Popen("taskkill /f /im msaccess.exe")
        except:
            pass

        time.sleep(1)

        try:
            #se recupera el path del Access donde hacer el merge (BBDD_02 ) con dicc_codigos_bbdd
            #se reemplaza / por \ sino no va y no abre la bbdd en modo escritura
            path_bbdd = dicc_codigos_bbdd["BBDD_02"][tipo_bbdd_selecc]["PATH_BBDD"]
            path_bbdd = path_bbdd.replace('/', '\\')


            #se abre el access y se accede a su codigo VBA (se abre en modo exclusivo para poder realizar cambios OpenCurrentDatabase --> True)
            access_app = win32com.client.Dispatch("Access.Application")
            access_app.OpenCurrentDatabase(path_bbdd, False)

            vba_project = access_app.VBE.ActiveVBProject
            current_bbdd = access_app.CurrentDb()


        except Exception as Err_access_abrir_bbdd:

            #se informa del posible error al abrir la bbdd
            traceback_error = traceback.extract_tb(Err_access_abrir_bbdd.__traceback__)
            modulo_python = os.path.basename(traceback_error[0].filename)
            rutina_python = traceback_error[0].name
            linea_error = traceback_error[0].lineno

            lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

            dicc_errores_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                , "MODULO_PYTHON": modulo_python
                                , "RUTINA_PYTHON": rutina_python
                                , "TIPO_REPOSITORIO": None
                                , "REPOSITORIO": None
                                , "TIPO_OBJETO": None
                                , "NOMBRE_OBJETO": None
                                , "LINEA_ERROR": linea_error
                                , "ERRORES": str(Err_access_abrir_bbdd)
                                }

            if isinstance(lista_dicc_errores_migracion, list):
                lista_dicc_errores_migracion.append(dicc_errores_temp)
            else:
                lista_dicc_errores_migracion = [dicc_errores_temp]

            dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = lista_dicc_errores_migracion
            del dicc_errores_temp
            del lista_dicc_errores_migracion

            return #aqui es return (se para la rutina)


        else:

            ##############################################################################################
            #                TABLAS / VINCULOS
            ##############################################################################################

            #se migran las tablas / vinculos (si se han localizado cambios en estos tipos de objetos)
            #mediante bucle sobre la lista lista_dicc_tablas_y_vinculos_con_cambios calculada mas arriba
            #dentro del bucle se encapsula un bloque try except else para registrar para cada objeto de la iteracion
            #los que se han migrado correctamente (sentencia else) y los que no (sentencia except)
            #en ambos casos se genera un ficheros .txt
            if len(lista_dicc_tablas_y_vinculos_con_cambios) != 0:
                    
                for dicc_tablas_y_vinculos in lista_dicc_tablas_y_vinculos_con_cambios:

                    tipo_objeto = dicc_tablas_y_vinculos["TIPO_OBJETO"]
                    nombre_objeto = dicc_tablas_y_vinculos["NOMBRE_OBJETO"]
                    sentencia_create = dicc_tablas_y_vinculos["PARAMETROS_MERGE"]["SENTENCIA_CREATE"]
                    objeto_origen = dicc_tablas_y_vinculos["PARAMETROS_MERGE"]["OBJETO_ORIGEN"]

                    try:
                        #se elimina la tabla / vinculo
                        current_bbdd.TableDefs.Delete(nombre_objeto)
                    except:
                        pass#por si la tabla / vinculo no existe

                    finally:

                        try:
                            #se ejecuta la creacion de tablas / vinculos tan solo si tras eliminar el flag ELIMINADO
                            #la sentencia sentencia_create calculada mediante join es un string no None
                            if isinstance(sentencia_create, str):

                                #TABLA_LOCAL
                                if tipo_objeto == "TABLA_LOCAL":
                                    current_bbdd.Execute(sentencia_create)


                                #VINCULO_ODBC
                                elif tipo_objeto in ["VINCULO_ODBC", "VINCULO_OTRO"]:                   

                                    table_def = current_bbdd.CreateTableDef(nombre_objeto)
                                    table_def.Connect = sentencia_create
                                    table_def.SourceTableName = objeto_origen
                                    current_bbdd.TableDefs.Append(table_def)


                        except Exception as Err_access_tablas_y_vinculos:  
                            #objetos con errores migracion --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_ERRORES_MIGRACION)

                            traceback_error = traceback.extract_tb(Err_access_tablas_y_vinculos.__traceback__)
                            modulo_python = os.path.basename(traceback_error[0].filename)
                            rutina_python = traceback_error[0].name
                            linea_error = traceback_error[0].lineno

                            modulo_python = os.path.basename(modulo_python)

                            lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

                            dicc_errores_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                                , "MODULO_PYTHON": modulo_python
                                                , "RUTINA_PYTHON": rutina_python
                                                , "TIPO_REPOSITORIO": None
                                                , "REPOSITORIO": None
                                                , "TIPO_OBJETO": tipo_objeto
                                                , "NOMBRE_OBJETO": nombre_objeto
                                                , "LINEA_ERROR": linea_error
                                                , "ERRORES": str(Err_access_tablas_y_vinculos)}

                            if isinstance(lista_dicc_errores_migracion, list):
                                lista_dicc_errores_migracion.append(dicc_errores_temp)
                            else:
                                lista_dicc_errores_migracion = [dicc_errores_temp]

                            dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = lista_dicc_errores_migracion
                            del dicc_errores_temp
                            del lista_dicc_errores_migracion
                            pass #aqui es pass para que se encapsulen en el diccionario todos los errores generados


                        else:
                            #objetos migrados correctamente --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_OK_MIGRACION)
                            lista_dicc_ok_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"]

                            dicc_ok_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                            , "TIPO_REPOSITORIO": None
                                            , "REPOSITORIO": None
                                            , "TIPO_OBJETO": tipo_objeto
                                            , "NOMBRE_OBJETO": nombre_objeto
                                            }

                            if isinstance(lista_dicc_ok_migracion, list):
                                lista_dicc_ok_migracion.append(dicc_ok_temp)
                            else:
                                lista_dicc_ok_migracion = [dicc_ok_temp]

                            dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"] = lista_dicc_ok_migracion
                            del dicc_ok_temp
                            del lista_dicc_ok_migracion


            ##############################################################################################
            #                MODULOS VBA
            ##############################################################################################

            #se migran los modulos con cambios (si se han localizado cambios) usando los modulos reconstruidos
            #mediante bucle sobre la lista lista_dicc_modulos_con_cambios calculada mas arriba
            #dentro del bucle se encapsula un bloque try except else para registrar para cada objeto de la iteracion
            #los que se han migrado correctamente (sentencia else) y los que no (sentencia except)
            #en ambos casos se genera un ficheros .txt
            if len(lista_dicc_modulos_con_cambios) != 0:
                    
                for dicc_modulos in lista_dicc_modulos_con_cambios:

                    try:
                        tipo_repositorio = dicc_modulos["TIPO_REPOSITORIO"]
                        repositorio = dicc_modulos["NOMBRE_REPOSITORIO"]

                        df_codigo_repositorio = dicc_modulos["DF_CODIGO_REPOSITORIO"]
                        string_codigo_repositorio = "\n".join([df_codigo_repositorio.iloc[ind, df_codigo_repositorio.columns.get_loc("CODIGO")] for ind in df_codigo_repositorio.index])

                        #se conecta al modulo de la iteracion + se vacia el modulo + se agrega df_codigo al modulo
                        vba_module = vba_project.VBComponents(repositorio)
                        vba_module.CodeModule.DeleteLines(1, vba_module.CodeModule.CountOfLines)
                        vba_module.CodeModule.AddFromString(string_codigo_repositorio)


                    except Exception as Err_access_modulos_vba:
                        #modulos con errores migracion --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_ERRORES_MIGRACION)

                        traceback_error = traceback.extract_tb(Err_access_modulos_vba.__traceback__)
                        modulo_python = os.path.basename(traceback_error[0].filename)
                        rutina_python = traceback_error[0].name
                        linea_error = traceback_error[0].lineno

                        modulo_python = os.path.basename(modulo_python)

                        lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

                        dicc_errores_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                            , "MODULO_PYTHON": modulo_python
                                            , "RUTINA_PYTHON": rutina_python
                                            , "TIPO_REPOSITORIO": tipo_repositorio
                                            , "REPOSITORIO": repositorio
                                            , "TIPO_OBJETO": None
                                            , "NOMBRE_OBJETO": None
                                            , "LINEA_ERROR": linea_error
                                            , "ERRORES": str(Err_access_modulos_vba)}

                        if isinstance(lista_dicc_errores_migracion, list):
                            lista_dicc_errores_migracion.append(dicc_errores_temp)
                        else:
                            lista_dicc_errores_migracion = [dicc_errores_temp]

                        dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = lista_dicc_errores_migracion
                        del dicc_errores_temp
                        del lista_dicc_errores_migracion
                        pass #aqui es pass para que se encapsulen en el diccionario todos los errores generados


                    else:
                        #objetos migrados correctamente --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_OK_MIGRACION)
                        lista_dicc_ok_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"]

                        dicc_ok_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                        , "TIPO_REPOSITORIO": tipo_repositorio
                                        , "REPOSITORIO": repositorio
                                        , "TIPO_OBJETO": None
                                        , "NOMBRE_OBJETO": None
                                        }

                        if isinstance(lista_dicc_ok_migracion, list):
                            lista_dicc_ok_migracion.append(dicc_ok_temp)
                        else:
                            lista_dicc_ok_migracion = [dicc_ok_temp]

                        dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"] = lista_dicc_ok_migracion
                        del dicc_ok_temp
                        del lista_dicc_ok_migracion


        finally:
            access_app.CloseCurrentDatabase()
            access_app.Quit()
            access_app = None


    ####################################################################################################################################
    ####################################################################################################################################
    #           SQL SERVER
    ####################################################################################################################################
    ####################################################################################################################################

    elif tipo_bbdd_selecc == "SQL_SERVER":

        ####################################################################################################################################
        # CALCULOS
        ####################################################################################################################################

        #se crea la lista de diccionarios (uno por objeto) que sirve de base tanto para documentar como para realizar el merge
        #en esta lista se integran los tipos de objeto en este orden (TABLAS, FUNCIONES, VIEWS y STORED_PROCEDURES)
        #para evitar fallos en el merge en caso de dependencias de objetos con otros)
        lista_objetos_ordenados = ["TABLAS", "FUNCIONES", "VIEWS", "STORED_PROCEDURES"]

        lista_dicc_sql_server_objetos_con_cambios = []
        for tipo_objeto_ordenado in lista_objetos_ordenados:

            for dicc in lista_dicc_objetos_migrar_bbdd_fisica:
                tipo_objeto_subform = dicc["TIPO_OBJETO_SUBFORM"]
                tipo_objeto = func_dicc_control_versiones_tipo_objeto_buscar_en_dicc("TIPO_OBJETO_DESDE_SUBFORM", valor = tipo_objeto_subform)

                if tipo_objeto == tipo_objeto_ordenado:

                    repositorio = dicc["REPOSITORIO"]
                    nombre_objeto = dicc["NOMBRE_OBJETO"]
                    df_codigo = dicc["DF_CODIGO"]


                    #se crea el df de codigo para documentar
                    df_codigo_documentacion = df_codigo[["CODIGO_CON_NUM_LINEA", "CONTROL_CAMBIOS_ACTUAL"]].copy()

                    df_codigo_documentacion["ADD_DEL"] = (df_codigo_documentacion.apply(lambda x: "ADD\t" if x["CONTROL_CAMBIOS_ACTUAL"] == "AGREGADO" else "DEL\t" 
                                                                                        if x["CONTROL_CAMBIOS_ACTUAL"] == "ELIMINADO" else "---\t", 
                                                                                        axis = 1))
                    
                    df_codigo_documentacion["CODIGO"] = df_codigo_documentacion[["ADD_DEL", "CODIGO_CON_NUM_LINEA"]].astype(str).apply("".join, axis = 1)
                    df_codigo_documentacion = df_codigo_documentacion[["CODIGO"]]



                    #se crea la sentencia sentencia_sql el merge (se quitan las lineas con el flag ELIMINADO del df de codigo)
                    df_codigo_sin_flag_eliminado = df_codigo.loc[df_codigo["CONTROL_CAMBIOS_ACTUAL"] != "ELIMINADO", ["CODIGO"]]
                    df_codigo_sin_flag_eliminado.reset_index(drop = True, inplace = True)

                    sentencia_sql = "\n".join([df_codigo_sin_flag_eliminado.iloc[ind, 0] for ind in df_codigo_sin_flag_eliminado.index]) if len(df_codigo_sin_flag_eliminado) != 0 else None


                    #se crea un diccionario temporal que se agrega a la lista lista_dicc_sql_server_objetos_con_cambios
                    dicc_temp = {"REPOSITORIO": repositorio
                                , "TIPO_OBJETO": tipo_objeto
                                , "NOMBRE_OBJETO": nombre_objeto
                                , "DF_CODIGO_DOCUMENTACION": df_codigo_documentacion
                                , "SENTENCIA_SQL": sentencia_sql
                                }

                    lista_dicc_sql_server_objetos_con_cambios.append(dicc_temp)
                    del dicc_temp
                    del df_codigo_documentacion
                    del df_codigo_sin_flag_eliminado



        ####################################################################################################################################
        # AJUSTES PARA DOCUMENTACION Y MERGE EN BBDD FISICA
        ####################################################################################################################################

        #se crea la lista de los esquemas con cambios y se quitan los duplicados
        #(se usa tanto para la documentacion para crear subcarpetas como en el merge en bbdd fisica
        #para crear los esquemas si no existen en BBDD_02 antes de proceder a crear los objetos)
        lista_esquemas_con_cambios = [dicc["REPOSITORIO"] for dicc in lista_dicc_objetos_migrar_bbdd_fisica]
        lista_esquemas_con_cambios = list(dict.fromkeys(lista_esquemas_con_cambios))


        ####################################################################################################################################
        # DOCUMENTACION
        ####################################################################################################################################

        #se crea el directorio dentro de ruta_export
        nombre_carpeta_documentacion = tipo_bbdd_selecc + "_MIGRACION_" + now 
        ruta_carpeta_documentacion = str(ruta_export) + r'\%s' % nombre_carpeta_documentacion
        os.makedirs(ruta_carpeta_documentacion, exist_ok = True)


        #se crean las subcarpetas por esquemas
        if len(lista_esquemas_con_cambios) != 0:
            for esquema in lista_esquemas_con_cambios:

                ruta_carpeta_documentacion_esquemas = str(ruta_carpeta_documentacion) + r'\%s' % esquema
                os.makedirs(ruta_carpeta_documentacion_esquemas, exist_ok = True)

                #se localizan para el esquema de la iteracion que tipos de objeto tienen cambios
                lista_tipo_objetos_iteracion_esquemas = [dicc["TIPO_OBJETO"] for dicc in lista_dicc_sql_server_objetos_con_cambios if dicc["REPOSITORIO"] == esquema]
                lista_tipo_objetos_iteracion_esquemas = list(dict.fromkeys(lista_tipo_objetos_iteracion_esquemas))

                for tipo_objeto_esquema in lista_tipo_objetos_iteracion_esquemas:

                    ruta_carpeta_documentacion_esquemas_tipo_objetos = str(ruta_carpeta_documentacion_esquemas) + r'\%s' % tipo_objeto_esquema
                    os.makedirs(ruta_carpeta_documentacion_esquemas_tipo_objetos, exist_ok = True)

                    #se descargan los scripts de los objetos en ficheros .txt
                    #se crea la lista de los objetos con su df de las iteraciones por esquema y tipo de objeto
                    lista_descarga_objetos = [[dicc["NOMBRE_OBJETO"], dicc["DF_CODIGO_DOCUMENTACION"]]
                                              for dicc in lista_dicc_sql_server_objetos_con_cambios 
                                              if dicc["REPOSITORIO"] == esquema and dicc["TIPO_OBJETO"] == tipo_objeto_esquema]
                    
                    for nombre_objeto_documentacion, df_codigo_documentacion in lista_descarga_objetos:

                        nombre_fichero_txt = nombre_objeto_documentacion + ".txt"
                        ruta_fichero_txt = str(ruta_carpeta_documentacion_esquemas_tipo_objetos) + r'\%s' % nombre_fichero_txt

                        with open(ruta_fichero_txt, 'w') as fich:
                            for ind in df_codigo_documentacion.index:
                                fich.write(df_codigo_documentacion.iloc[ind, 0] + "\n")

                    del lista_descarga_objetos
                del lista_tipo_objetos_iteracion_esquemas



        ####################################################################################################################################
        # MERGE EN BBDD FISICA
        ####################################################################################################################################

        #el proceso se realiza por hitos:
        ################################### 1 --> se crean nuevos esquemas si se localizan
        ################################### 2 --> se crean las tablas (funciones, views y stored procedures pueden necesitar estas tablas)
        ################################### 3 --> se crean las funciones (views y stored procedures pueden necesitar estas funciones)
        ################################### 4 --> se crean las views
        ################################### 5 --> se crean los stored procedures
        ##### se almacenan en lista_errores los posibles errores de migracion

        try:
            #se establece la conexion a la bbdd sql server

            servidor_sql_server_merge = dicc_codigos_bbdd["BBDD_02"][tipo_bbdd_selecc]["SERVIDOR"]
            bbdd_sql_server_merge = dicc_codigos_bbdd["BBDD_02"][tipo_bbdd_selecc]["BBDD"]

            conn_str = dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["CONNECTING_STRING"].replace("REEMPLAZA_SERVIDOR", servidor_sql_server_merge).replace("REEMPLAZA_BBDD", bbdd_sql_server_merge)

            MiConexion = pyodbc.connect(conn_str)
            cursor = MiConexion.cursor()


        except Exception as Err_sql_server_conexion:
            #se informa del posible error al conectarse al servidor + bbdd

            traceback_error = traceback.extract_tb(Err_sql_server_conexion.__traceback__)
            modulo_python = os.path.basename(traceback_error[0].filename)
            rutina_python = traceback_error[0].name
            linea_error = traceback_error[0].lineno

            lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

            dicc_errores_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                , "MODULO_PYTHON": modulo_python
                                , "RUTINA_PYTHON": rutina_python
                                , "TIPO_REPOSITORIO": None
                                , "REPOSITORIO": None
                                , "TIPO_OBJETO": None
                                , "NOMBRE_OBJETO": None
                                , "LINEA_ERROR": linea_error
                                , "ERRORES": str(Err_sql_server_conexion)
                                }

            if isinstance(lista_dicc_errores_migracion, list):
                lista_dicc_errores_migracion.append(dicc_errores_temp)
            else:
                lista_dicc_errores_migracion = [dicc_errores_temp]

            dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = lista_dicc_errores_migracion
            del dicc_errores_temp
            del lista_dicc_errores_migracion
            return #aqui es return (se para la rutina)


        else:     

            #Paso 1 --> se crean los nuevos esquemas (si los hubiese)
            if len(lista_esquemas_con_cambios) != 0:
                for esquema in lista_esquemas_con_cambios:
                    try:
                        cadena_sql = "CREATE SCHEMA [" + esquema + "];"
                        cursor.execute(cadena_sql)
                        MiConexion.commit()

                    except:
                        pass#por si el esquema ya existe


            #Paso 2 a 5 --> tablas, funciones, views y stored procedures
            #se realiza el merge mediante bucle sobre la nueva lista ordenada lista_dicc_objetos_migrar_bbdd_fisica_ordenada
            for dicc in lista_dicc_sql_server_objetos_con_cambios:

                try:
                    repositorio = dicc["REPOSITORIO"]
                    tipo_objeto = dicc["TIPO_OBJETO"]
                    nombre_objeto = dicc["NOMBRE_OBJETO"]
                    sentencia_sql = dicc["SENTENCIA_SQL"]

                    repositorio_y_nombre_objeto = repositorio + "." + nombre_objeto

                    #se eliminan los objetos (si ya existian)
                    lista_objeto_id = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["TIPO_OBJETO"][tipo_objeto]["LISTA_OBJECT_ID"]
                    for objeto_id_tipo, objeto_id in lista_objeto_id:

                        cadena_sql = mod_sql_server.query_drop_if_exists_objeto.replace("REEMPLAZA_OBJETO_SHORT", repositorio_y_nombre_objeto).replace("REEMPLAZA_TIPO_OBJETO", objeto_id_tipo).replace("REEMPLAZA_ID_OBJETO", objeto_id)        
                        cursor.execute(cadena_sql)
                        MiConexion.commit()
     

                    #se crean los objetos (tan solo si sentencia_sql es string no None)
                    if isinstance(sentencia_sql, str):
                        cursor.execute(sentencia_sql)
                        MiConexion.commit()
 

                except Exception as Err_sql_server_objetos:
                    #objetos con errores migracion --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_ERRORES_MIGRACION)

                    traceback_error = traceback.extract_tb(Err_sql_server_objetos.__traceback__)
                    modulo_python = os.path.basename(traceback_error[0].filename)
                    rutina_python = traceback_error[0].name
                    linea_error = traceback_error[0].lineno

                    modulo_python = os.path.basename(modulo_python)

                    lista_dicc_errores_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"]

                    dicc_errores_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                        , "MODULO_PYTHON": modulo_python
                                        , "RUTINA_PYTHON": rutina_python
                                        , "TIPO_REPOSITORIO": None
                                        , "REPOSITORIO": repositorio
                                        , "TIPO_OBJETO": tipo_objeto_subform
                                        , "NOMBRE_OBJETO": nombre_objeto
                                        , "LINEA_ERROR": linea_error
                                        , "ERRORES": str(Err_sql_server_objetos)}

                    if isinstance(lista_dicc_errores_migracion, list):
                        lista_dicc_errores_migracion.append(dicc_errores_temp)
                    else:
                        lista_dicc_errores_migracion = [dicc_errores_temp]

                    dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_ERRORES_MIGRACION"] = lista_dicc_errores_migracion
                    del dicc_errores_temp
                    del lista_dicc_errores_migracion
                    pass #aqui es pass para que se encapsulen en el diccionario todos los errores generados


                else:
                    #objetos migrados correctamente --> se informa el diccionario dicc_control_versiones_tipo_objeto (LISTA_DICC_OK_MIGRACION)
                    lista_dicc_ok_migracion = dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"]

                    dicc_ok_temp = {"TIPO_BBDD": tipo_bbdd_selecc
                                    , "TIPO_REPOSITORIO": None
                                    , "REPOSITORIO": repositorio
                                    , "TIPO_OBJETO": tipo_objeto_subform
                                    , "NOMBRE_OBJETO": nombre_objeto
                                    }

                    if isinstance(lista_dicc_ok_migracion, list):
                        lista_dicc_ok_migracion.append(dicc_ok_temp)
                    else:
                        lista_dicc_ok_migracion = [dicc_ok_temp]

                    dicc_control_versiones_tipo_objeto[tipo_bbdd_selecc]["MERGE_BBDD_FISICA"]["LISTA_DICC_OK_MIGRACION"] = lista_dicc_ok_migracion
                    del dicc_ok_temp
                    del lista_dicc_ok_migracion


            MiConexion.close()

