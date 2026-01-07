import pandas as pd
import os
import re
import shutil
import traceback
import win32com.client
import xlwings as xw
import warnings
import datetime as dt

import APP_2_GENERAL as mod_gen


#################################################################################################################################################################################
##                     VARIABLES GENERALES
#################################################################################################################################################################################


#############################################
# query sistema access
#############################################

#query para recuperar de una bbdd Access el objeto origen de un vinculo (ODBC u otro)
query_access_vinculos = """SELECT Name AS NOMBRE_VINCULO_ACCESS, Connect AS LINK_ODBC, ForeignName AS OBJETO_ORIGEN
                            FROM MSysObjects
                        """


#############################################
# indicadores varios
#############################################

#numero maximo de parametros que VBA permite en las rutinas / funciones
num_max_parametros_rutina_vba_access = 60


#si la linea empieza por este indicador (previo trimeo) la linea de codigo completa esta comentada
indicador_comentario_vba_access = "'"



#si la linea (previo trimeo) acaba por este indicador indica que la linea de codigo esta truncada y sigue en la siguiente linea de codigo
#sirve para localizar las lineas de codigo que son bloques de declaraciones de parametros
indicador_linea_codigo_vba_truncada_a_linea_siguiente = " _"



#cuando se localizan los bloques de codigo que son declaraciones de parametros, este indicador permite splitear el codigo
#el nombre de la variable y su tipo de dato para informarlo en el diagnostico de dependencias
indicador_variable_vba_declaracion_tipo_dato = " As "


#es el indicador que pemite localizar si una variable eglobal es constante
#en el caso de las variables definidas por el usuario permite asignarle a cada sub-variable entre parentesis su valor 
indicador_variable_signo_igual = " = "


#sirve cuando en la GUI de inicio al pulsar el boton ADD para BBDD_01 y BBDD_02 abrir un filedialog que solo muestre los ficheros
#que son bbdd Access (mdb o accdb)
lista_GUI_askopenfilename_ms_access = [("Access Database Files", "*.mdb;*.accdb")]



#############################################
# variables relacionadas con objetos VBA
#############################################


dicc_objetos_vba = {"VARIABLE_PUBLICA_NORMAL":
                                {"STRING_PARA_LISTADO_EXCEL": "VARIABLE PUBLICA NORMAL"
                                , "LISTA_DECLARACION_INICIO_NO_ES_CONSTANTE": ["Public", "Global"]
                                , "LISTA_DECLARACION_INICIO_SI_ES_CONSTANTE": ["Public Const", "Global Const", "Const"]
                                }
                    # --> STRING_PARA_LISTADO_EXCEL: es el iteral que sale en el excel de diagnostico de dependencias columna D (hoja VARIABLES (LISTADO))
                    # --> LISTA_DECLARACION_INICIO_NO_ES_CONSTANTE: es una lista con todos los inicios de declaracion de variables publicas normales (que no son constantes)
                    # --> LISTA_DECLARACION_IINICIO_SI_ES_CONSTANTE: es una lista con todos los inicios de declaracion de variables publicas normales (que son constantes)

                    , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO":
                                                {"STRING_PARA_LISTADO_EXCEL": "VARIABLE PUBLICA DEFINIDA POR USUARIO"
                                                , "LISTA_DECLARACION_INICIO": ["Type", "Enum"]
                                                , "LISTA_DECLARACION_FIN": ["End", "End Type", "End Enum"]
                                                }
                    # --> STRING_PARA_LISTADO_EXCEL: es el iteral que sale en el excel de diagnostico de dependencias columna D (hoja VARIABLES (LISTADO))
                    # --> LISTA_DECLARACION_INICIO: es una lista con todos los inicios de declaracion de variables publicas normales (que no son constantes)
                    # --> LISTA_DECLARACION_FIN: es una lista con todos los finales de declaracion de variables publicas normales (que son constantes)

                    , "RUTINAS":
                                {"RUTINA_PRIVADA":
                                        {"TIPO_RUTINA": "RUTINA"
                                        , "TIPO_DECLARACION": "PRIVADA"
                                        , "LISTA_DECLARACION_INICIO": ["Private Sub"]
                                        , "LISTA_DECLARACION_FIN": ["End Sub"]
                                        }
                                , "RUTINA_PUBLICA":
                                        {"TIPO_RUTINA": "RUTINA"
                                        , "TIPO_DECLARACION": "PUBLICA"
                                        , "LISTA_DECLARACION_INICIO": ["Sub", "Public Sub", "Global Sub"]
                                        , "LISTA_DECLARACION_FIN": ["End Sub"]
                                        }
                                , "FUNCION_PRIVADA":
                                        {"TIPO_RUTINA": "FUNCION"
                                        , "TIPO_DECLARACION": "PRIVADA"
                                        , "LISTA_DECLARACION_INICIO": ["Private Function"]
                                        , "LISTA_DECLARACION_FIN": ["End Function"]
                                        }
                                , "FUNCION_PUBLICA":
                                        {"TIPO_RUTINA": "FUNCION"
                                        , "TIPO_DECLARACION": "PUBLICA"
                                        , "LISTA_DECLARACION_INICIO": ["Function", "Public Function", "Global Function"]
                                        , "LISTA_DECLARACION_FIN": ["End Function"]
                                        }
                                }
                    # --> LISTA_DECLARACION_INICIO: es una lista con todos los inicios de declaracion de rutinas / funciones
                    # --> LISTA_DECLARACION_FIN: es una lista con todos los finales de declaracion de rutinas / funciones
                                                
                    , "RUTINAS_VARIABLES_LOCALES":
                                        {"LISTA_DECLARACION_INICIO_NO_ES_CONSTANTE": ["Dim"]
                                        , "LISTA_DECLARACION_INICIO_SI_ES_CONSTANTE": ["Const"]
                                        } 
                    # --> LISTA_DECLARACION_INICIO: es una lista con todos los inicios de declaracion de variables locales definidas en rutinas / funciones
                          
                    }



#############################################
# variables relacionadas con el diagnostico de dependencias
#############################################

#diccionario que permite "limpiar" el codigo de las rutinas/funciones VBA para localizar cuando se usa una tabla local o vinculo en el codigo VBA
#permite preservar en el codigo VBA lo que se explica de LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL + lo que se explica de LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL
#los calculos se hacen en la funcion de este modulo Python func_diagnostico_access_tablas_codigo_preservar y los calculos se hacen por separado
#para sumar el resuldo de los 2 y obtener todos los codigos con sentencias donde poder buscar dependencias de tablas
#
# --> key LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL: para todas las instrucciones dentro del codigo VBA incluidos entre "", se listan todo lo parecido a una sentencia SQL
#                                                que es donde se puede considerar que se esta usando la tabla/vinculo en VBA
#                                                (en el df de codigo en una columna especial para diagnostica de tablas se quita todo lo no incluido entre "" y tambien
#                                                todo lo incluido entre "" que no contengan los items de la lista LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL)
#
# --> key LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL: es una lista de instrucciones VBA de manipulacion de tablas que preceden un "" donde y dentro del "" figura el nombre de una tabla/vinculo
#                                                (en la columna del df mencionado para LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL se converva tan solo el codigo VBA relacionado con los items de esta lista)

dicc_diagnostico_tablas_access_donde_buscar = {
                                                "LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL": 
                                                    ["SELECT", "FROM", "LEFT JOIN", "RIGHT JOIN", "INNER JOIN", "GROUP BY", "ON", "SET", "WHERE", "HAVING",
                                                    "CREATE TABLE", "DELETE", "DROP TABLE", "INSERT INTO", "UPDATE"]
                                                #es una lista de instrucciones SQL para preservar el codigo incluido entre comillas dobles "" y que se parezcan a sentencias SQL
                                                #(es indiferente poner en mayusculas ominusculas, el chequeo se hace poniendo tanto el codigo VBA de las rutinas como los items de la 
                                                #lista de esta key en mayusculas)
                                                    
                                                , "LISTA_TEXTOS_INSTRUCCIONES_VBA_MANIP_TABLAS":
                                                    [r"OpenRecordset\(", "DoCmd.OpenTable acTable,", r"DoCmd.OpenTable\(acTable,", "DoCmd.CloseacTable,", 
                                                     r"DoCmd.TransferSpreadsheet\(acImport, acSpreadsheetTypeExcel12,"]
                                                #es una lista de instrucciones VBA de uso de tablas / vinculos, si se localizan en la linea de codigo se preserva 
                                                #lo incluido entre comillas dobles "" inmediatamente posterior a la instruccion VBA

                                                #en caso de agregar una instruccion VBA que conlleve un parentesis de apertura ( hay que informar el item
                                                #de la forma siguiente: "OpenRecordset("" --> r"OpenRecordset\("
                                                #IMPORTANTE: no configurar en los items comillas dobles tipo "DoCmd.OpenTable\""  
                                                }




#############################################
# listas de columnas de df usados en el app
#############################################


#lista headers df_access_vinculos
lista_headers_df_access_vinculos = ["NOMBRE_VINCULO_ACCESS", "LINK_ODBC", "OBJETO_ORIGEN"]


#lista nuevas columnas calculadas para rutinas y variables publicas
lista_headers_nuevas_columnas_rutinas = ["INDICE", "TIPO_RUTINA", "TIPO_DECLARACION_RUTINA", "NOMBRE_RUTINA", "NUMERO_LINEA_CODIGO_RUTINA", "ORDEN_RUTINA_EN_MODULO", 
                                "PARAMETROS_RUTINA", "TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA"]

lista_headers_nuevas_columnas_variables_publicas = ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "ES_VARIABLE_PUBLICA", "CODIGO"]


#lista headers diagnostico importacion codigo VBA access
lista_headers_df_bbdd_codigo = ["TIPO_MODULO", "NOMBRE_MODULO", "NUMERO_LINEA_CODIGO_MODULO", "CODIGO"]


#lista headers para proceso diagnostico (df coddigo donde buscar)

lista_headers_df_diagnostico_donde_buscar_campos_usados = ["TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "NUMERO_LINEA_CODIGO_RUTINA", "CODIGO", "CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI", 
                                                           "CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO", "CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES", 
                                                           "TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA"]

lista_headers_df_diagnostico_donde_buscar_tablas = ["TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "NUMERO_LINEA_CODIGO_RUTINA", "CODIGO", "CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI", 
                                                    "CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO"]

lista_headers_df_diagnostico_donde_buscar_rutinas_y_variables = ["TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES"]
lista_headers_df_diagnostico_donde_buscar_variables_definidas_por_usuario = ["TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA"]

           
#lista headers para proceso diagnostico export excel
lista_headers_df_listado = ["TIPO_OBJETO", "TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_OBJETO", "TIPO_OBJETO_2", "TIPO_DECLARACION_RUTINA", "PARAMETROS_RUTINA", 
                                "TIPO_DATO_VARIABLE_PUBLICA", "VARIABLE_PUBLICA_ES_CONSTANTE", "VARIABLE_PUBLICA_CONSTANTE_VALOR", 
                                "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES", "CONNECTING_STRING_VINCULOS"]

lista_headers_df_dependencias = ["TIPO_OBJETO", "TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_OBJETO", "SE_USA_EN_TIPO_MODULO", "SE_USA_EN_NOMBRE_MODULO", "SE_USA_EN_NOMBRE_RUTINA"]
lista_headers_df_sin_dependencias = ["TIPO_OBJETO", "TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_OBJETO"]

lista_headers_df_check_manual_tablas = ["TIPO_TABLA", "NOMBRE_TABLA", "TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "NUMERO_LINEA_CODIGO_RUTINA", "CODIGO_LINEA"]



#################################################################################################################################################################################
##                     FUNCIONES MS ACCESS - CONTROL DE VERSIONES
#################################################################################################################################################################################


def func_codigo_vba_quitar_comentarios(codigo_linea):
    #funcion que permite quitar el comentario de una linea de codigo VBA
    #se realiza mediante funcion propia pq el framework re no dio los resultados esperados

    lista_codigo_sin_comentario = []
    esta_entre_comilla_simple = False
    esta_entre_comilla_doble = False
    
    i = 0
    while i < len(codigo_linea):
        caracter = codigo_linea[i]

        if caracter == '"' and not esta_entre_comilla_simple:
            esta_entre_comilla_doble = not esta_entre_comilla_doble
            lista_codigo_sin_comentario.append(caracter)

        elif caracter == "'" and not esta_entre_comilla_doble:
            if esta_entre_comilla_simple:
                esta_entre_comilla_simple = False
                lista_codigo_sin_comentario.append(caracter)
            else:
                break
        else:
            lista_codigo_sin_comentario.append(caracter)

        i += 1

    #resultado de la funcion
    return "".join(lista_codigo_sin_comentario).strip()



def func_diagnostico_access_tablas_df_create_table(opcion, **kwargs):
    #funcion que genera la sentencia CREATE TABLE (en forma de df) de la tabla Access y la connecting string para los vinculos ODBC u otros

    prop_tabla = kwargs.get("prop_tabla", None)
    nombre_objeto = kwargs.get("nombre_objeto", None)
    link_connect = kwargs.get("link_connect", None)
    nombre_vinculo = kwargs.get("nombre_vinculo", None)
    df_access_vinculos = kwargs.get("df_access_vinculos", None)


    if opcion == "TABLA_LOCAL":

        sentencia_sql = f"CREATE TABLE [{nombre_objeto}]\n(\n"

        cont = 0
        for campo in prop_tabla.Fields:

            cont += 1
            nombre_campo = campo.Name
            tipo_campo = campo.Type
            size_campo = campo.Size if tipo_campo in [1, 10] else None

            if tipo_campo == 1:
                temp = f"VARCHAR({size_campo})"
                
            elif tipo_campo in [2, 8, 10]:
                temp = "TEXT(255)"

            else:

                if tipo_campo == 3:
                    temp = "DOUBLE"

                elif tipo_campo == 4:
                    temp = "DATE"

                elif tipo_campo == 5:
                    temp = "DECIMAL(19,4)"

                elif tipo_campo == 6:
                    temp = "AUTOINCREMENT"

                elif tipo_campo == 7:
                    temp = "BOOLEAN"

                elif tipo_campo == 9: 
                    temp = "TEXT(255)"

                else:
                    temp = "TEXT(255)"

            sentencia_sql += f"\t[{nombre_campo}] {temp}\n" if cont == 1 else f"\t, [{nombre_campo}] {temp}\n"

        sentencia_sql = sentencia_sql + ")"

        lista_temp_1 = sentencia_sql.split("\n")
        lista_temp_2 = [[i] for i in lista_temp_1]
        df_resultado = pd.DataFrame(lista_temp_2, columns = ["CODIGO"])



    elif opcion == "VINCULO_ODBC":

        df_resultado = (df_access_vinculos.loc[(df_access_vinculos["LINK_ODBC"].isnull() == False) & (df_access_vinculos["NOMBRE_VINCULO_ACCESS"] == nombre_vinculo) & 
                                            (df_access_vinculos["OBJETO_ORIGEN"].isnull() == False), ["OBJETO_ORIGEN"]])
        
        df_resultado.reset_index(drop = True, inplace = True)

        lista_connect_temp = link_connect.split(";") + ["SOURCE: " + df_resultado.iloc[0, 0]] if len(df_resultado) != 0 else [link_connect, "ERROR: No se pudo establecer la conexion al vinculo y recuperar la tabla origen."]

        df_resultado = pd.DataFrame({"CODIGO": lista_connect_temp})



    elif opcion == "VINCULO_OTRO":
            
        df_resultado = (df_access_vinculos.loc[(df_access_vinculos["LINK_ODBC"].isnull() == True) & (df_access_vinculos["NOMBRE_VINCULO_ACCESS"] == nombre_vinculo) & 
                                            (df_access_vinculos["OBJETO_ORIGEN"].isnull() == False), ["OBJETO_ORIGEN"]])
        
        df_resultado.reset_index(drop = True, inplace = True)

        lista_connect_temp = [link_connect, "SOURCE: " + df_resultado.iloc[0, 0]] if len(df_resultado) != 0 else [link_connect, "ERROR: No se pudo establecer la conexion al vinculo y recuperar la tabla origen.."]

        df_resultado = pd.DataFrame({"CODIGO": lista_connect_temp})


    #resultado de la funcion
    return df_resultado


#################################################################################################################################################################################
##                     FUNCIONES MS ACCESS - DIAGNOSTICO
#################################################################################################################################################################################

def func_access_bbdd_codigo_calculos(opcion, **kwargs):
    #funcion que permite calcular listas por indicies del df df_bbdd_codigo_param (parametro de la funcion) y permiten calcular columnas adicionales en este df de codigo
    #realacionadas con rutinas y variables publicas
    #permite tambien crear la lista de diccionarias para las variables publicas con la misma estructura de diccionarios que el resto de tipo de objetos (tablas, vinculos y rutinas)
    #
    #la funcion tien 3 opciones:
    # --> LISTA_RUTINAS
    # --> LISTA_VARIABLES_PUBLICAS
    # --> LISTA_DICC_OBJETOS_VARIABLES_PUBLICAS


    #parametros kwargs
    df_bbdd_codigo_param = kwargs.get("df_bbdd_codigo_param", None)
    lista_calculos_variables_publicas = kwargs.get("lista_calculos_variables_publicas", None)


    warnings.filterwarnings("ignore")


    lista_resultado_funcion = []

    if opcion == "LISTA_RUTINAS":
        #la opcion permite crear una lista de listas (cuya longitud es la misma que el df df_bbdd_param, parametro de la funcion)
        #donde cada sublista contiene:
        # --> INDICE                                              indice del df df_bbdd_param
        # --> TIPO_RUTINA                                         rutina o funcion
        # --> TIPO_DECLARACION_RUTINA                             publica o privada
        # --> NOMBRE_RUTINA                                       nombre rutina / funcion
        # --> NUMERO_LINEA_CODIGO_RUTINA                          numero de linea del codigo dentro de la rutina / funcion
        # --> PARAMETROS_RUTINA                                   parametros de la rutina / funcion con los tipos de datos asociados
        #
        # --> TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA     recopila los tipos de dato de los parametros y variables locales de la rutina
        #                                                         se usa en el diagnostico de dependencias para localizar si variables publicas definidas por el usuario se usan en la rutina
        #                                                         (mismo valor para todas las lineas de la rutina)
        ############################
        #parametros kwargs usados --> df_bbdd_codigo_param


        ##################################################################################################
        ##################################################################################################
        #                AJUSTES PREVIOS
        ##################################################################################################
        ##################################################################################################

        #se crean las listas lista_declaracion_rutinas_vba_ini_ord_desc y lista_declaracion_rutinas_vba_fin
        #donde se fusionan todos los inicios y finales de declaracion de rutinas / funciones publicas
        # --> lista_declaracion_rutinas_vba_ini_ord_desc se ordena de longitud de instruccion declaracion inicial VBA de mayor a menor
        # --> lista_declaracion_rutinas_vba_fin no requiere ni el tipo de rutina ni el tipo de declaracion solo la instruccion declaracion final VBA
        lista_declaracion_rutinas_vba_ini_ord_desc = [[dicc_objetos_vba["RUTINAS"][key_2]["TIPO_RUTINA"], dicc_objetos_vba["RUTINAS"][key_2]["TIPO_DECLARACION"], key_3.strip() + " "] 
                                                        for key_2 in dicc_objetos_vba["RUTINAS"].keys() for key_3 in dicc_objetos_vba["RUTINAS"][key_2]["LISTA_DECLARACION_INICIO"]]

        lista_declaracion_rutinas_vba_ini_ord_desc.sort(key = lambda x: len(x[2]), reverse = True)

        tupla_declaracion_rutinas_vba_ini = tuple([item[2] for item in lista_declaracion_rutinas_vba_ini_ord_desc])


        lista_declaracion_rutinas_vba_fin = [key_3.strip() for key_2 in dicc_objetos_vba["RUTINAS"].keys() for key_3 in dicc_objetos_vba["RUTINAS"][key_2]["LISTA_DECLARACION_FIN"]]
        lista_declaracion_rutinas_vba_fin = list(dict.fromkeys(lista_declaracion_rutinas_vba_fin))#se quitan los duplicados

        tupla_declaracion_rutinas_vba_fin = tuple(lista_declaracion_rutinas_vba_fin)


        ##################################################################################################
        ##################################################################################################
        #                CALCULOS
        ##################################################################################################
        ##################################################################################################

        #se calculan las listas lista_indices_rutinas_ini y lista_indices_rutinas_fin que almacenan cada una sublistas
        #con el tipo y nombre del formulario y el indice de inicio / final de declaracion de una rutina / funcion VBA
        lista_campos_usados = ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE"]

        df_temp_declaracion_rutinas_vba_ini = df_bbdd_codigo_param.loc[df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_rutinas_vba_ini), lista_campos_usados]
        df_temp_declaracion_rutinas_vba_fin = df_bbdd_codigo_param.loc[df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_rutinas_vba_fin), lista_campos_usados]

        df_temp_declaracion_rutinas_vba_ini.reset_index(drop = True, inplace = True)
        df_temp_declaracion_rutinas_vba_fin.reset_index(drop = True, inplace = True)

        lista_indices_rutinas_ini = []
        lista_indices_rutinas_fin = []

        if len(df_temp_declaracion_rutinas_vba_ini) != 0:
            lista_indices_rutinas_ini = [[df_temp_declaracion_rutinas_vba_ini.iloc[ind, df_temp_declaracion_rutinas_vba_ini.columns.get_loc("TIPO_MODULO")]
                                        , df_temp_declaracion_rutinas_vba_ini.iloc[ind, df_temp_declaracion_rutinas_vba_ini.columns.get_loc("NOMBRE_MODULO")]
                                        , df_temp_declaracion_rutinas_vba_ini.iloc[ind, df_temp_declaracion_rutinas_vba_ini.columns.get_loc("INDICE")]]
                                        for ind in df_temp_declaracion_rutinas_vba_ini.index]
        
            lista_indices_rutinas_ini = sorted(lista_indices_rutinas_ini, key = lambda x: (x[0], x[1], x[2]))
            

        if len(df_temp_declaracion_rutinas_vba_fin) != 0:
            lista_indices_rutinas_fin = [[df_temp_declaracion_rutinas_vba_fin.iloc[ind, df_temp_declaracion_rutinas_vba_fin.columns.get_loc("TIPO_MODULO")]
                                        , df_temp_declaracion_rutinas_vba_fin.iloc[ind, df_temp_declaracion_rutinas_vba_fin.columns.get_loc("NOMBRE_MODULO")]
                                        , df_temp_declaracion_rutinas_vba_fin.iloc[ind, df_temp_declaracion_rutinas_vba_fin.columns.get_loc("INDICE")]]
                                        for ind in df_temp_declaracion_rutinas_vba_fin.index]
        
            lista_indices_rutinas_fin = sorted(lista_indices_rutinas_fin, key = lambda x: (x[0], x[1], x[2]))


        del df_temp_declaracion_rutinas_vba_ini
        del df_temp_declaracion_rutinas_vba_fin



        #se ejecuta el proceso de crear la lista lista_rutinas_modulo tan solo si se han localizado en los modulos VBA rutinas / funciones
        lista_dicc_datos_rutinas = []
        if len(lista_indices_rutinas_ini) != 0 and len(lista_indices_rutinas_ini) == len(lista_indices_rutinas_fin):
        

            #se extrae la lista lista_indices_y_codigo_inicio_rutina con el tipo y nombre de los modulos y todos los indices y codigo asociado al inicio de la rutina
            #se re-ordena la lista por tipo y nombre de los modulos y por indices
            lista_indices_ini_declar = [item[2] for item in lista_indices_rutinas_ini]
            df_indices_y_codigo_inicio_rutina = (df_bbdd_codigo_param.loc[df_bbdd_codigo_param["INDICE"].isin(lista_indices_ini_declar), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
    
            lista_indices_y_codigo_inicio_rutina = df_indices_y_codigo_inicio_rutina.values.tolist()
            lista_indices_y_codigo_inicio_rutina = sorted(lista_indices_y_codigo_inicio_rutina, key = lambda x: (x[0], x[1], x[2]))

            
            #se calcula la lista lista_dicc_datos_rutinas
            for tipo_modulo_ini, nombre_modulo_ini, indice_ini, linea_codigo_ini in lista_indices_y_codigo_inicio_rutina:


                #se calcula el indice de fin de declaracion de la rutina buscando en lista_indices_rutinas_fin el 1er indice de fin de declaracion
                #inmediatamente posterior a indice_ini siempre y cuando el tipo y nombre del modulo de lista_indices_y_codigo_inicio_rutina y lista_indices_rutinas_fin coincidan
                indice_fin = [item[2] for item in lista_indices_rutinas_fin if item[0] == tipo_modulo_ini and item[1] == nombre_modulo_ini and item[2] > indice_ini][0]


                #se calcula el tipo_rutina, tipo_declaracion_rutina y nombre_rutina
                #
                #se splitea la linea de codigo por el parentesis de apertura "(" para conservar solo el 1er item 
                #el 2ndo son los parametros (si los hubiese) que se calculan en este mismo bucle mas adelante
                linea_codigo_spliteada_parentesis = linea_codigo_ini.split("(")[0]

                tipo_rutina_localizada = None
                tipo_declaracion_rutina_localizada = None
                nombre_rutina_localizada = None

                for tipo_rutina, tipo_declaracion_rutina, ini_declar in lista_declaracion_rutinas_vba_ini_ord_desc:

                    if ini_declar in linea_codigo_spliteada_parentesis:
                        tipo_rutina_localizada = tipo_rutina
                        tipo_declaracion_rutina_localizada = tipo_declaracion_rutina
                        nombre_rutina_localizada = linea_codigo_spliteada_parentesis.replace(ini_declar, "").strip()

                        break

                #se calcula el numero de orden de la rutina dentro del modulo
                numero_orden_rutina_en_modulo = [item[2] for item in lista_indices_y_codigo_inicio_rutina if item[0] == tipo_modulo_ini and item[1] == nombre_modulo_ini].index(indice_ini) + 1



                ############################################################
                #      PARAMETROS
                ############################################################

                #las rutinas / funciones VBA pueden tener como maximo 60 parametros y estos pueden declararse en varias lineas (cada una acabando por "_")
                #se extrae del parametro de la presente funcion df_bbdd_codigo_param las 60 primeras lineas de la rutina (siempre y cuando la rutina tenga mas de 60 lineas
                #, en caso contrario se extrae todas las lineas del a rutina), se quitan las lineas vacias tambien
                indice_fin_extracc = indice_ini + num_max_parametros_rutina_vba_access if indice_ini + num_max_parametros_rutina_vba_access < indice_fin else indice_fin

                df_extracc_codigo = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["INDICE"] >= indice_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_fin_extracc) &
                                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.replace(" ", "") != indicador_linea_codigo_vba_truncada_a_linea_siguiente.strip()) & 
                                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                                ["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])            

                df_extracc_codigo.reset_index(drop = True, inplace = True)


                #se concatena en un string los 60 primeras lineas de codigo de la rutina / funcion
                #pq es el numero maximo de parametros que VBA permite en la declaracion de parametros que pueden realizarse en varias lineas
                #(las lineas que acaban por " _" se reemplaza el " _" por nada)
                string_60_primeras_lineas = "".join([df_extracc_codigo.iloc[ind, 0].strip()
                                                    if df_extracc_codigo.iloc[ind, 0][-2:] != indicador_linea_codigo_vba_truncada_a_linea_siguiente
                                                    else df_extracc_codigo.iloc[ind, 0][:-2].strip()
                                                    for ind in df_extracc_codigo.index])
                
                del df_extracc_codigo
  
                #se comprueba si la rutina / funcion tiene parametros (el 1er parentesis de cierre si es inmediatamente posterior 
                #al 1er parentesis de aperture es que la rutina no tiene parametros)
                lista_parentesis_apertura = list(re.finditer(r"\(", string_60_primeras_lineas))
                primer_parentesis_apertura = [lista_parentesis_apertura[ind].start() for ind, item in enumerate(lista_parentesis_apertura)][0]

                lista_parentesis_cierre = list(re.finditer(r"\)", string_60_primeras_lineas))
                primer_parentesis_cierre = [lista_parentesis_cierre[ind].start() for ind, item in enumerate(lista_parentesis_cierre)][0]

                rutina_tiene_parametros = "NO" if primer_parentesis_cierre == primer_parentesis_apertura + 1 else "SI"


                #si la rutina tiene parametros se calcula string_parametros donde en string_60_primeras_lineas se reemplaza previamente los "()" por "@@"
                #(@ es un caracter prohibido en VBA en declaraciones de parametros por lo que no hay riesgo de alterar el nombre o tipo de datos de un parametro)
                #y se vuelve a calcular las posiciones dentro del string del 1er parentesis de apertura y cierre
                #(los "()" dentro de la declaracion de un parametro corresponden a arrays, si se hace este replace por "@@" el calculo del string de parametros 
                #seria incorrecto si hubiese arrays)
                string_parametros = ""
                if rutina_tiene_parametros == "SI":
                    
                    string_parametros = string_60_primeras_lineas.replace("()", "@@")

                    lista_parentesis_apertura = list(re.finditer(r"\(", string_parametros))
                    primer_parentesis_apertura = [lista_parentesis_apertura[ind].start() for ind, item in enumerate(lista_parentesis_apertura)][0]

                    lista_parentesis_cierre = list(re.finditer(r"\)", string_parametros))
                    primer_parentesis_cierre = [lista_parentesis_cierre[ind].start() for ind, item in enumerate(lista_parentesis_cierre)][0]

                    string_parametros = string_parametros[primer_parentesis_apertura + 1:primer_parentesis_cierre].replace("@@", "()")

                del lista_parentesis_apertura
                del lista_parentesis_cierre


                ################################################################################################################
                #      STRING_TIPO_DATO_Y_VARIABLES_LOCALES
                ################################################################################################################

                #se genera el string con los tipos de datos usados en los parametros y variables locales para el 
                #diagnostico de dependencias de las variables publicas definidas por el usuario

                ####################################################
                #      PARAMETROS
                ####################################################

                #se crea la lista lista_parametros_rutina a partir de string parametros deonde previamente se quita todos lo encapsulado entre comillas dobles ""
                #y tambien todo lo encapsulado entre parentesis pq el string se splitea despues por la coma ", " (podria haber comas intercadas ahi)

                lista_parametros_rutina = []
                if isinstance(string_parametros, str):

                    #se quita lo encapsulado entre comillas dobles y lo encapsulado entre parantesis
                    string_parametros_corr = func_access_bbdd_codigo_calculos_varios("QUITAR_ENCAPSULADO_ENTRE_COMILLAS_DOBLES", string_param = string_parametros)
                    string_parametros_corr = func_access_bbdd_codigo_calculos_varios("QUITAR_ENCAPSULADO_ENTRE_PARENTESIS", string_param = string_parametros_corr)

                    #se crea la lista lista_parametros_rutina spliteada por la coma ","
                    lista_parametros_rutina = string_parametros_corr.split(",")


                ####################################################
                #      VARIABLES LOCALES
                ####################################################

                #se calcula lista_variables_locales_rutina_linea_codigo con la funcion func_access_bbdd_codigo_calculos_varios
                lista_variables_locales_rutina_linea_codigo = func_access_bbdd_codigo_calculos_varios("LISTA_DECLARACIONES_VARIABLES_LOCALES"
                                                                                                        , indice_param_ini = indice_ini
                                                                                                        , indice_param_fin = indice_fin
                                                                                                        , df_bbdd_codigo_param = df_bbdd_codigo_param
                                                                                                        )
                
                #se crea la lista lista_variables_locales_rutina
                lista_variables_locales_rutina = []

                if len(lista_variables_locales_rutina_linea_codigo) != 0:

                    for linea_codigo in lista_variables_locales_rutina_linea_codigo:

                        linea_codigo_corr = linea_codigo

                        #se quitan los " _" al final de la linea de codigo
                        if linea_codigo_corr[-2:] == indicador_linea_codigo_vba_truncada_a_linea_siguiente:
                            linea_codigo_corr = linea_codigo_corr[:-2]


                        #se quita lo encapsulado entre comillas dobles y lo encapsulado entre parantesis
                        linea_codigo_corr = func_access_bbdd_codigo_calculos_varios("QUITAR_ENCAPSULADO_ENTRE_COMILLAS_DOBLES", string_param = linea_codigo_corr)
                        linea_codigo_corr = func_access_bbdd_codigo_calculos_varios("QUITAR_ENCAPSULADO_ENTRE_PARENTESIS", string_param = linea_codigo_corr)


                        #se crea la lista linea_codigo_corr_spliteada_por_coma spliteada por la coma ","
                        linea_codigo_corr_spliteada_por_coma = linea_codigo_corr.split(",")


                        #se agregan los items de linea_codigo_corr_spliteada_por_coma a lista_variables_locales_rutina
                        for item_split_coma in linea_codigo_corr_spliteada_por_coma:
                            if len(item_split_coma) != 0:
                                lista_variables_locales_rutina.append(item_split_coma)


                del lista_variables_locales_rutina_linea_codigo


                ####################################################
                #      TIPO DATO PARAMETROS Y VARIABLES LOCALES
                ####################################################

                lista_parametros_y_variables_locales = lista_parametros_rutina + lista_variables_locales_rutina

                lista_tipo_dato_parametros_y_variables_locales = []
                for item in lista_parametros_y_variables_locales:

                    #se splitea el item por " As " y se conserva el 2ndo item de la lista resultante y se agrega a lista_tipo_dato_parametros_y_variables_locales
                    item_spliteado_por_as = item.strip().split(indicador_variable_vba_declaracion_tipo_dato)
                    tipo_dato_parametros_y_variables_locales = item_spliteado_por_as[1].strip() if len(item_spliteado_por_as) == 2 else ""

                    lista_tipo_dato_parametros_y_variables_locales.append(tipo_dato_parametros_y_variables_locales)

                #se quitan los duplicados de lista_tipo_dato_parametros_y_variables_locales en string y se convierte en string
                lista_tipo_dato_parametros_y_variables_locales = list(dict.fromkeys(lista_tipo_dato_parametros_y_variables_locales))
                string_tipo_dato_parametros_y_variables_locales = ", ".join(lista_tipo_dato_parametros_y_variables_locales)



                ############################################################
                #      SE AGREGAN LOS DATOS A lista_dicc_datos_rutinas
                ############################################################
                #se crea el diccionario temporal y se agrega a lista_dicc_datos_rutinas
                dicc_temp = {"TIPO_MODULO": tipo_modulo_ini
                            , "NOMBRE_MODULO": nombre_modulo_ini
                            , "INDICE_INICIO": indice_ini
                            , "INDICE_FIN": indice_fin
                            , "TIPO_RUTINA": tipo_rutina_localizada
                            , "TIPO_DECLARACION_RUTINA": tipo_declaracion_rutina_localizada
                            , "NOMBRE_RUTINA": nombre_rutina_localizada
                            , "NUMERO_ORDEN_RUTINA_EN_MODULO": numero_orden_rutina_en_modulo
                            , "RUTINA_TIENE_PARAMETROS": rutina_tiene_parametros
                            , "STRING_PARAMETROS_RUTINA": string_parametros
                            , "STRING_TIPO_DATO_Y_VARIABLES_LOCALES": string_tipo_dato_parametros_y_variables_locales
                            }

                lista_dicc_datos_rutinas.append(dicc_temp)
                del dicc_temp


            del lista_indices_ini_declar
            del df_indices_y_codigo_inicio_rutina
            del lista_indices_y_codigo_inicio_rutina


            ##################################################################################################
            ##################################################################################################
            #                CALCULO RESULTADO FUNCION
            ##################################################################################################
            ##################################################################################################

            #se localiza el 1er y ultimo indice del parametro df_bbdd_codigo_param
            primer_indice_df_bbdd_codigo_param = df_bbdd_codigo_param["INDICE"].min()
            ultimo_indice_df_bbdd_codigo_param = df_bbdd_codigo_param["INDICE"].max()


            #se insertan en la lista lista_resultado_funcion cada uno de los indices de codigo ubciados en la lista lista_dicc_datos_rutinas,
            #los que no encuuentra se completan los indices vacios
            for indice_df in range(primer_indice_df_bbdd_codigo_param, ultimo_indice_df_bbdd_codigo_param + 1, 1):

                check_rutina = 0
                for dicc in lista_dicc_datos_rutinas:

                    indice_ini = dicc["INDICE_INICIO"]
                    indice_fin = dicc["INDICE_FIN"]

                    numero_linea_codigo_rutina = 0
                    if indice_df >= indice_ini and indice_df <= indice_fin:

                        check_rutina = 1
                            
                        tipo_rutina = dicc["TIPO_RUTINA"]
                        tipo_declaracion = dicc["TIPO_DECLARACION_RUTINA"]
                        nombre_rutina = dicc["NOMBRE_RUTINA"]
                        numero_orden_rutina_en_modulo = dicc["NUMERO_ORDEN_RUTINA_EN_MODULO"]
                        rutina_tiene_parametros = dicc["RUTINA_TIENE_PARAMETROS"]
                        string_parametros_rutina = dicc["STRING_PARAMETROS_RUTINA"]
                        string_tipo_dato_parametros_y_variables_locales = dicc["STRING_TIPO_DATO_Y_VARIABLES_LOCALES"]

        
                        #se crean los numeros de linea de codigo dentro de la rutina / funcion
                        cont = 0
                        for cont in range(indice_ini, indice_df + 1, 1):
                            numero_linea_codigo_rutina = cont - indice_ini + 1


                        #se agregan los datos a lista_resultado_funcion
                        lista_resultado_funcion.append([indice_df
                                                , tipo_rutina
                                                , tipo_declaracion
                                                , nombre_rutina
                                                , numero_linea_codigo_rutina
                                                , numero_orden_rutina_en_modulo
                                                , string_parametros_rutina
                                                , string_tipo_dato_parametros_y_variables_locales
                                                ])                    


                #si indice_df no sale lista_dicc_datos_rutinas se añade con valores todos nulos
                if check_rutina == 0:
                    lista_resultado_funcion.append([indice_df, None, None, None, None, None, None, None])  




    elif opcion == "LISTA_VARIABLES_PUBLICAS":
        #la opcion permite crear una lista de listas (cuya longitud es la misma que el df df_bbdd_param, parametro de la funcion)
        #donde los items de cada sublista son:
        # --> tipo modulo
        # --> nombre modulo
        # --> el indice de linea dentro del df de codigo
        # --> indicador de si es variable publica y en caso de serlo coje 4 valores posibles:
        #               --> NO
        #               --> NO_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por "Public" o "Global")
        #               --> SI_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por ""Const", "Public Const" o "Global Const")
        #               --> DEFINIDA_POR_USUARIO_iii (para variables definidas por el usuario, se declaran en varias lineas, se crea un literal unico 
        #                                             para todas las lineas ue componen la declaracion de una misma variable, iii es un numero de variable unico 
        #                                             para cada una de estas variables)
        # --> la linea de codigo
        ############################
        #parametros kwargs usados --> df_bbdd_codigo_param


        #se extraen todas las lineas de codigo
        df_variables_publicas = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["ES_ENCABEZADO_MODULO"] == "NO") & (df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) & 
                                                            (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                            ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "NOMBRE_RUTINA", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
        
        df_variables_publicas.reset_index(drop = True, inplace = True)


        #se calcula la lista de modulos
        df_temp = df_variables_publicas.copy()
        df_temp.drop_duplicates(subset = ["TIPO_MODULO", "NOMBRE_MODULO"], keep = "last", inplace = True)
        df_temp.reset_index(drop = True, inplace = True)

        lista_modulos_con_variables_publicas = [[df_temp.iloc[ind, df_temp.columns.get_loc("TIPO_MODULO")]
                                                , df_temp.iloc[ind, df_temp.columns.get_loc("NOMBRE_MODULO")]]
                                                for ind in df_temp.index]
        del df_temp


        #se ejecuta el proceso tan solo si se localizan variables publicas en los modulos VBA,el calculo se hace por modulo
        lista_variables_publicas = []
        if len(lista_modulos_con_variables_publicas) != 0:

            for tipo_modulo, nombre_modulo in lista_modulos_con_variables_publicas:

                #se calcula el 1er y ultimo indice del modulo de la iteracion en df_variables_publicas
                primer_indice_df_bbdd_codigo_param_modulo = df_variables_publicas[(df_variables_publicas["TIPO_MODULO"] == tipo_modulo) & (df_variables_publicas["NOMBRE_MODULO"] == nombre_modulo)]["INDICE"].min()
                ultimo_indice_df_bbdd_codigo_param_modulo = df_variables_publicas[(df_variables_publicas["TIPO_MODULO"] == tipo_modulo) & (df_variables_publicas["NOMBRE_MODULO"] == nombre_modulo)]["INDICE"].max()


                #se calcula mediante la funcion func_access_bbdd_codigo_calculos_varios la lista de todas las lineas de codigo asociadas 
                #a variables publicas, es lista de lista donde cada sublista que la compone contiene:
                # --> tipo modulo
                # --> nombre modulo
                # --> indice linea codigo
                # --> identificador del tipo de variable publica:
                #        --> si es constante o no (si es variable publica normal)
                #        --> flag unico que afecta varias lines de codigo (si es variable publica definida por el usuario)
                # --> linea de codigo
                #
                #se agrega la lista a lista_variables_publicas tan solo si la lista resultante de la funcion no esta vacia 
                #(la lista lista_modulos_con_variables_publicas contiene todos los modulos con variables publicas sean normales 
                #o sean definidas por el usuario por lo que puede haber modulos donde solo aparezcan variables publicas normales 
                #y en este caso al calcular con la funcion la lista de variables publicas definidas por el usuario esta salga vacia 
                #y vice versa)


                # VARIABLES PUBLICAS NORMALES
                lista_variables_publicas_normales_modulo = func_access_bbdd_codigo_calculos_varios("LISTA_DECLARACIONES_VARIABLES_PUBLICAS_NORMALES"
                                                                                                        , indice_param_ini = primer_indice_df_bbdd_codigo_param_modulo
                                                                                                        , indice_param_fin = ultimo_indice_df_bbdd_codigo_param_modulo
                                                                                                        , df_bbdd_codigo_param = df_variables_publicas
                                                                                                        )

                if len(lista_variables_publicas_normales_modulo) != 0:

                    for item in lista_variables_publicas_normales_modulo:
                        lista_variables_publicas.append(item)


                # VARIABLES PUBLICAS DEFINIDAS POR USUARIO
                lista_variables_publicas_definidas_por_usuario_modulo = func_access_bbdd_codigo_calculos_varios("LISTA_DECLARACIONES_VARIABLES_PUBLICAS_DEFINIDAS_USUARIO"
                                                                                                                    , indice_param_ini = primer_indice_df_bbdd_codigo_param_modulo
                                                                                                                    , indice_param_fin = ultimo_indice_df_bbdd_codigo_param_modulo
                                                                                                                    , df_bbdd_codigo_param = df_variables_publicas
                                                                                                                    )

                if len(lista_variables_publicas_definidas_por_usuario_modulo) != 0:

                    for item in lista_variables_publicas_definidas_por_usuario_modulo:
                        lista_variables_publicas.append(item)



            #se crea lista_variables_publicas_indices que contiene solo los indices de lista_variables_publicas
            #se copia lista_variables_publicas para crear lista_resultado_funcion y se completan los huecos de indices del df de codigo 
            #que no estan en lista_variables_publicas por NO
            primer_indice_df_bbdd_codigo_param = df_bbdd_codigo_param["INDICE"].min()
            ultimo_indice_df_bbdd_codigo_param = df_bbdd_codigo_param["INDICE"].max()

            lista_variables_publicas_indices = [indice_linea for tipo_modulo, nombre_modulo, indice_linea, tipo_variable, linea_codigo in lista_variables_publicas]

            lista_resultado_funcion = lista_variables_publicas.copy()

            #se completan los huecos con los indices de codigo que no estan en lista_variables_publicas_indices
            for indice_df in range(primer_indice_df_bbdd_codigo_param, ultimo_indice_df_bbdd_codigo_param + 1, 1):

                if not indice_df in lista_variables_publicas_indices:
                    lista_resultado_funcion.append([tipo_modulo, nombre_modulo, indice_df, "NO", None])  

            del lista_variables_publicas_indices

        del lista_variables_publicas
        del df_variables_publicas




    elif opcion == "LISTA_DICC_OBJETOS_VARIABLES_PUBLICAS":
        #opcion permite preparar una lista de diccionarios que lista una a una todas las variables públicas normales y las que estan definidas por el usuario
        #(estas al declararse en varias lineas se unifican en 1 sola concatenando sus sub-variables con su tipo de dato o valor)
        #
        #la estructura de los diccionarios que componen esta lista es comun a todos los objetos MS ACCESS (tablas / vinculos, variables publicas VBA y rutinas / funciones VBA)
        #por lo que algunos items no aplican y se establecen a None
        #
        # --> TIPO_OBJETO                                           se usa (aqui VARIABLES_VBA)
        # --> TIPO_MODULO                                           se usa
        # --> NOMBRE_MODULO                                         se usa
        # --> NOMBRE_OBJETO                                         se usa
        # --> TIPO_RUTINA                                           no se usa (solo para rutinas / funciones VBA)
        # --> TIPO_DECLARACION_RUTINA                               no se usa (solo para rutinas / funciones VBA)
        # --> PARAMETROS_RUTINA                                     no se usa (solo para rutinas / funciones VBA)
        # --> TIPO_VARIABLE_PUBLICA                                 se usa
        # --> TIPO_DATO_VARIABLE_PUBLICA                            se usa
        # --> VARIABLE_PUBLICA_ES_CONSTANTE                         se usa
        # --> VARIABLE_PUBLICA_CONSTANTE_VALOR                      se usa
        # --> VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES   se usa
        # --> DF_CODIGO                                             no se usa (el codigo a nivel de variables va por tipo y nombre de modulo VBA, se calcula directamente
        #                                                                      en la rutina def_proceso_access_1_import)
        # --> CONNECTING_STRING_VINCULOS                            no se usa
        #
        #se calcula usando la lista lista_calculos_variables_publicas (parametro de la funcion) que se genera usando la presente funcion (opcion LISTA_VARIABLES_PUBLICAS)
        #en una llamada anterior a esta misma funcion en la rutina def_proceso_access_1_import
        #cada una de las sublistas que componen esta lista lista_calculos_variables_publicas contienen:
        # --> tipo modulo
        # --> nombre modulo0
        # --> el indice de linea dentro del df de codigo
        # --> indicador de si es variable publica y en caso de serlo coje 4 valores posibles:
        #               --> NO
        #               --> NO_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por "Public" o "Global")
        #               --> SI_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por ""Const", "Public Const" o "Global Const")
        #               --> DEFINIDA_POR_USUARIO_iii (para variables definidas por el usuario, se declaran en varias lineas, se crea un literal unico 
        #                                             para todas las lineas que componen la declaracion de una misma variable, iii es un numero de variable unico 
        #                                             para cada una de estas variables)
        # --> la linea de codigo
        ############################
        #parametros kwargs usados --> lista_calculos_variables_publicas




        ##################################################################################################
        # AJUSTES PREVIOS
        ##################################################################################################

        #se fusionan todos los inicios de declaracion en una lista donde cada item es el inicio de declaracion ordenado longitud mayor a menor)
        #(esta lista lista_variables_publicas_normales_declar_ini_ord_desc, que se usa tambien en la opcion DF_COLUMNAS_VARIABLES_PUBLICAS
        #de la presente funcion, no requiere aqui crear tambien una lista de sublistas, con el incio de declaracion es suficiente porque
        #se usa tan solo para hacer replace del item por nada)
        lista_variables_publicas_normales_declar_ini_ord_desc = [item.strip() + " " for key_2 in dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"].keys() 
                                        for item in dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"][key_2] if "LISTA_DECLARACION_INICIO" in key_2]

        lista_variables_publicas_normales_declar_ini_ord_desc.sort(key = lambda x: len(x[1]), reverse = True)


        #se crean la lista lista_variable_definida_por_usuario_ini que recupera la lista almacenada en subkey_3 (LISTA_DECLARACION_INICIO)
        #asociada a la key_1 (VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO) del diccionario dicc_objetos_vbay cada uno de sus items se trime y se 
        # le agrega un espacio en blanco al final (es por si el usuario ha configurado mal el diccionario)
        lista_variable_definida_por_usuario_ini = [item.strip() + " " for item in dicc_objetos_vba["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO"]["LISTA_DECLARACION_INICIO"]]


        ##################################################################################################
        # CALCULOS - VARIABLES PUBLICAS NORMALES
        ##################################################################################################

        #se crea la lista de modulos donde hay variables publicas (donde se quitan los duplicados) 911
        lista_modulos_con_variables_publicas_normales = [[tipo_modulo, nombre_modulo] for tipo_modulo, nombre_modulo, indice_linea, flag_variable_publica, linea_codigo in lista_calculos_variables_publicas
                                                            if flag_variable_publica in ["NO_ES_CONSTANTE", "SI_ES_CONSTANTE"]]

        lista_modulos_con_variables_publicas_normales = [sublista for i, sublista in enumerate(lista_modulos_con_variables_publicas_normales) if sublista not in lista_modulos_con_variables_publicas_normales[:i]]


        #se crea la lista lista_resultado_funcion resultado de la funcion mediante bucle sobre los distintos modulos con variables publicas
        #los calculos se realizan de forma distinata segun el tipo de variable publica
        for tipo_modulo_iteracion, nombre_modulo_iteracion in lista_modulos_con_variables_publicas_normales:

            #se recuperan el flag de tipo de variable y las lineas de codigo asociadas a este tipo de variable publica y para el modulo de la iteracion
            lista_variables_publicas_lineas_modulo = [[flag_variable_publica, linea_codigo] 
                                                        for tipo_modulo, nombre_modulo, indice_linea, flag_variable_publica, linea_codigo in lista_calculos_variables_publicas
                                                        if tipo_modulo == tipo_modulo_iteracion and nombre_modulo == nombre_modulo_iteracion and 
                                                        flag_variable_publica in ["NO_ES_CONSTANTE", "SI_ES_CONSTANTE"]]


            #mediante bucle por linea de codigo sobre la lista lista_variables_publicas_lineas_modulo:
            #--> se quitan los inicios de declaracion de este tipo de variables publicas
            #--> se calcula la lista de variables publicas 1 a 1 mediante la funcion func_access_bbdd_codigo_calculos_varios (opcion LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE)
            #--> mediante bucle sobre la lista del paso anterior se crea diccionario temporal que se agrega a la lista lista_resultado_funcion
            #previamente se quitan los inicios de declaracion de este tipo de variables publicas
            for flag_variable_publica, linea_codigo in lista_variables_publicas_lineas_modulo:

                #se quitan los inicios de declaracion segun el tipo de variable publica
                for item_replace in lista_variables_publicas_normales_declar_ini_ord_desc:
                    if item_replace in linea_codigo:
                        linea_codigo = linea_codigo.replace(item_replace, "")


                #se calcula la lista de variables publicas 1 a 1 mediante la funcion func_access_bbdd_codigo_calculos_varios
                opcion_funcion = "LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE" if flag_variable_publica == "NO_ES_CONSTANTE" else "LISTA_VARIABLES_PUBLICAS_SI_ES_CONSTANTE"
                lista_variables_publicas_1_a_1 = func_access_bbdd_codigo_calculos_varios(opcion_funcion, linea_codigo = linea_codigo)

                
                #mediante bucle sobre la lista del paso anterior se crea diccionario temporal que se agrega a la lista lista_resultado_funcion
                for item_variable_publica in lista_variables_publicas_1_a_1:
                    
                    variable_publica_nombre = item_variable_publica[0]
                    variable_publica_tipo = item_variable_publica[1]
                    variable_publica_tipo_dato = item_variable_publica[2]
                    variable_publica_es_constante = item_variable_publica[3]
                    variable_publica_es_constante_valor = item_variable_publica[4]
                    variable_publica_definida_usuario_subvariables = item_variable_publica[5]

                    dicc_temp = {"TIPO_OBJETO": "VARIABLES_VBA"
                                , "TIPO_MODULO": tipo_modulo_iteracion
                                , "NOMBRE_MODULO": nombre_modulo_iteracion
                                , "NOMBRE_OBJETO": variable_publica_nombre
                                , "TIPO_RUTINA": None                                              #key que no aplica
                                , "TIPO_DECLARACION_RUTINA": None                                  #key que no aplica
                                , "PARAMETROS_RUTINA": None                                        #key que no aplica
                                , "TIPO_VARIABLE_PUBLICA": variable_publica_tipo
                                , "TIPO_DATO_VARIABLE_PUBLICA": variable_publica_tipo_dato
                                , "VARIABLE_PUBLICA_ES_CONSTANTE": variable_publica_es_constante
                                , "VARIABLE_PUBLICA_CONSTANTE_VALOR": variable_publica_es_constante_valor
                                , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": variable_publica_definida_usuario_subvariables
                                , "DF_CODIGO": None                                                #key que no aplica por ahora (se asigna de otra forma, ver rutina def_proceso_access_2_control_versiones)
                                , "CONNECTING_STRING_VINCULOS": None                               #key que no aplica
                                }

                    lista_resultado_funcion.append(dicc_temp)
                    del dicc_temp


        ##################################################################################################
        # CALCULOS - VARIABLES PUBLICAS DEFINIDAS POR USUARIO
        ##################################################################################################
  
        #el calculo se hace aparte de las variables publicas normales pq aqui se realiza sobre el conjunto de lineas que conforman la variable publica definida por el usuario
        #en el bloque anterior (variable publica normal) el calculo se realiza por linea de codigo 1 a 1)
        ####################################################################################################

        #se crea la lista de modulos donde hay variables publicas definidas por el usuario (donde se quitan los duplicados)
        lista_modulos_con_variables_publicas_definidas_usuario = [[tipo_modulo, nombre_modulo] for tipo_modulo, nombre_modulo, indice_linea, flag_variable_publica, linea_codigo in lista_calculos_variables_publicas
                                                                    if "DEFINIDA_POR_USUARIO" in flag_variable_publica]

        lista_modulos_con_variables_publicas_definidas_usuario = [sublista for i, sublista in enumerate(lista_modulos_con_variables_publicas_normales) 
                                                                  if sublista not in lista_modulos_con_variables_publicas_normales[:i]]


        #se crea la lista lista_resultado_funcion resultado de la funcion mediante bucle sobre los distintos modulos con variables publicas
        #los calculos se realizan de forma distinata segun el tipo de variable publica
        for tipo_modulo_iteracion, nombre_modulo_iteracion in lista_modulos_con_variables_publicas_definidas_usuario:

            #se recuperan los distintos flags de este tipo de variable publica dentro del modulo de la iteracion donde se quitan los duplicados
            lista_variables_publicas_definidas_usuario_modulo = [flag_variable_publica
                                                                    for tipo_modulo, nombre_modulo, indice_linea, flag_variable_publica, linea_codigo in lista_calculos_variables_publicas
                                                                    if tipo_modulo == tipo_modulo_iteracion and nombre_modulo == nombre_modulo_iteracion and "DEFINIDA_POR_USUARIO" in flag_variable_publica]

            lista_variables_publicas_definidas_usuario_modulo = list(dict.fromkeys(lista_variables_publicas_definidas_usuario_modulo))


            #mediante bucle sobre cada item de lista_variables_publicas_definidas_usuario_modulo se recuperan las lineas de3 codigo asociadas
            #que se re-odenan por indice de linea y se ejecuta la funcion func_access_bbdd_codigo_calculos_varios para recuperar los datos de la variable
            #en el formato esperado
            for flag_variable_publica_modulo in lista_variables_publicas_definidas_usuario_modulo:

                lista_lineas_flag_variable_publica = [[indice_linea, linea_codigo]
                                                        for tipo_modulo, nombre_modulo, indice_linea, flag_variable_publica, linea_codigo in lista_calculos_variables_publicas
                                                        if tipo_modulo == tipo_modulo_iteracion and nombre_modulo == nombre_modulo_iteracion and flag_variable_publica == flag_variable_publica_modulo]

                lista_lineas_flag_variable_publica = sorted(lista_lineas_flag_variable_publica, key = lambda x: (x[0]))

                lista_variables_publicas_1_a_1 = func_access_bbdd_codigo_calculos_varios("LISTA_VARIABLES_PUBLICAS_DEFINIDAS_POR_USUARIO"
                                                                                            , lista_variable_definida_por_usuario_ini = lista_variable_definida_por_usuario_ini
                                                                                            , lista_lineas_flag_variable_publica = lista_lineas_flag_variable_publica)
  

                #mediante bucle sobre la lista del paso anterior se crea diccionario temporal que se agrega a la lista lista_resultado_funcion
                for item_variable_publica in lista_variables_publicas_1_a_1:
                    
                    variable_publica_nombre = item_variable_publica[0]
                    variable_publica_tipo = item_variable_publica[1]
                    variable_publica_tipo_dato = item_variable_publica[2]
                    variable_publica_es_constante = item_variable_publica[3]
                    variable_publica_es_constante_valor = item_variable_publica[4]
                    variable_publica_definida_usuario_subvariables = item_variable_publica[5]

                    dicc_temp = {"TIPO_OBJETO": "VARIABLES_VBA"
                                , "TIPO_MODULO": tipo_modulo_iteracion
                                , "NOMBRE_MODULO": nombre_modulo_iteracion
                                , "NOMBRE_OBJETO": variable_publica_nombre
                                , "TIPO_RUTINA": None                                              #key que no aplica
                                , "TIPO_DECLARACION_RUTINA": None                                  #key que no aplica
                                , "PARAMETROS_RUTINA": None                                        #key que no aplica
                                , "TIPO_VARIABLE_PUBLICA": variable_publica_tipo
                                , "TIPO_DATO_VARIABLE_PUBLICA": variable_publica_tipo_dato
                                , "VARIABLE_PUBLICA_ES_CONSTANTE": variable_publica_es_constante
                                , "VARIABLE_PUBLICA_CONSTANTE_VALOR": variable_publica_es_constante_valor
                                , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": variable_publica_definida_usuario_subvariables
                                , "DF_CODIGO": None                                                #key que no aplica por ahora (se asigna de otra forma, ver rutina def_proceso_access_2_control_versiones)
                                , "CONNECTING_STRING_VINCULOS": None                               #key que no aplica
                                }

                    lista_resultado_funcion.append(dicc_temp)
                    del dicc_temp


    #resultado de la funcion
    return lista_resultado_funcion




def func_access_bbdd_codigo_calculos_varios(opcion, **kwargs):
    #funcion que permite realizar tareas varias en la funcion func_access_df_bbdd_codigo_calculos_varios
    #se ponen en otra funcion pq son tareas que se usan en varios sitios, es para aligerar el codigo de func_access_df_bbdd_codigo_calculos_varios
    #
    #la funcion tiene 3 opciones:
    # --> QUITAR_ENCAPSULADO_ENTRE_COMILLAS_DOBLES                       crea un string donde se quita todo lo encapsulado entre commillas dobles dentro de un string
    # --> QUITAR_ENCAPSULADO_ENTRE_PARENTESIS                            crea un string donde se quita todo lo encapsulado entre parentesis dentro de un string
    #
    # --> LISTA_DECLARACIONES_VARIABLES_LOCALES                          crea lista que contiene las lineas de codigo asociadas a declaraciones 
    #                                                                    de variables locales de rutinas / funciones
    #
    # --> LISTA_DECLARACIONES_VARIABLES_PUBLICAS_NORMALES                crea lista de lista de variables publicas normales donde cada sublista contiene:
    #                                                                    --> tipo modulo
    #                                                                    --> nombre modulo
    #                                                                    --> el indice de linea dentro del df de codigo
    #                                                                    --> indicador de si es constante o no
    #                                                                    --> la linea de codigo
    #
    # --> LISTA_DECLARACIONES_VARIABLES_PUBLICAS_DEFINIDAS_USUARIO       crea lista de lista de variables publicas definidas por el usuario donde cada sublista contiene:
    #                                                                    --> tipo modulo
    #                                                                    --> nombre modulo
    #                                                                    --> el indice de linea dentro del df de codigo
    #                                                                    --> flag unico para este tipo de variable que se declara en varias lineas 
    #                                                                    --> la linea de codigo
    #
    # --> LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE                       a partir de una linea de codigo, lista las variables publicas que no son constantes
    #                                                                    (en una misma linea se pueden declara varias variables publicas)
    #                                                                    es lista de lista donde cada sublista contiene los items siguientes (algunos no aplican, se dejan para tener la misma 
    #                                                                    estructura de sublistas para todos los tipos de variables publicas):
    #                                                                    --> nombre variable publica
    #                                                                    --> tipo variable publica        
    #                                                                    --> tipo dato variable publica
    #                                                                    --> es constante
    #                                                                    --> valor constante
    #                                                                    --> subvariables variable definida por el usuario
    #
    # --> LISTA_VARIABLES_PUBLICAS_SI_ES_CONSTANTE                       hace lo mismo que la opcion LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE pero esta vez para las variables publicas
    #                                                                    que son constantes
    #
    # --> LISTA_VARIABLES_PUBLICAS_DEFINIDAS_POR_USUARIO                 unificar todas las lineas de codigo asociadas a una variable publica definida por el usuario (se declaran en varias lineas)
    #                                                                    y preparar una lista de lista en el mismo formato que las variables publicas normales (LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE y
    #                                                                    LISTA_VARIABLES_PUBLICAS_SI_ES_CONSTANTE)


    resultado_funcion = None


    #parametros kwargs
    string_param = kwargs.get("string_param", None)
    indice_param_ini = kwargs.get("indice_param_ini", None)
    indice_param_fin = kwargs.get("indice_param_fin", None)
    df_bbdd_codigo_param = kwargs.get("df_bbdd_codigo_param", None)
    linea_codigo = kwargs.get("linea_codigo", None)
    lista_variable_definida_por_usuario_ini = kwargs.get("lista_variable_definida_por_usuario_ini", None)
    lista_lineas_flag_variable_publica = kwargs.get("lista_lineas_flag_variable_publica", None)


    if opcion == "QUITAR_ENCAPSULADO_ENTRE_COMILLAS_DOBLES":
        #parametros kwargs usados --> string_param

        resultado_funcion =  re.compile(r'"([^"]*)"').sub('""', string_param)



    elif opcion == "QUITAR_ENCAPSULADO_ENTRE_PARENTESIS":
        #no se hace con el metodo  re como la opcion anterior pq no me dio los resultados esperados 
        #por lo que se hace de forma manual
        #########################
        #parametros kwargs usados --> string_param

        #se crean listas temporales que almacenan los indices de posicion de los parentesis de apertura y cierre en string_param
        lista_indices_parentesis_apertura = list(re.finditer(r"\(", string_param))
        lista_indices_parentesis_cierre= list(re.finditer(r"\)", string_param))

        lista_indices_parentesis_apertura = [lista_indices_parentesis_apertura[ind].start() for ind, item in enumerate(lista_indices_parentesis_apertura)]
        lista_indices_parentesis_cierre = [lista_indices_parentesis_cierre[ind].start() for ind, item in enumerate(lista_indices_parentesis_cierre)]

        #se revierten los items de lista_indices_parentesis_apertura
        lista_indices_parentesis_apertura_reversed = lista_indices_parentesis_apertura.copy()
        lista_indices_parentesis_apertura_reversed.reverse()


        #mediante bucle sobre la lista lista_indices_parentesis_apertura_reversed, se busca el item de lista_indices_parentesis_cierre
        #inmediatamente consecutivo que no se haya localizado en una iteracion anterior para otro item de apertura (de ahi el uso de la lista
        #lista_exclusion_parentesis_cierre donde en cada iteracion se agrega el item de cierre inmediatamente consecutivo localizado)
        #con el item de apertura y el item de cierre localizado se cre tupla que se agrega a la lista lista_tuplas_parentesis_replace
        lista_tuplas_parentesis_replace = []
        lista_exclusion_parentesis_cierre = []
        for item_apertura in lista_indices_parentesis_apertura_reversed:

            try:
                item_cierre_consecutivo = [item_cierre for item_cierre in lista_indices_parentesis_cierre if item_cierre > item_apertura and item_cierre not in lista_exclusion_parentesis_cierre][0]
                lista_exclusion_parentesis_cierre.append(item_cierre_consecutivo)

                tupla = (item_apertura, item_cierre_consecutivo, item_cierre_consecutivo - item_apertura)
                lista_tuplas_parentesis_replace.append(tupla)

            except:#por si el numero de parentesis de apertura y cierre no coincide
                pass


        #se ordena la lista lista_tuplas_parentesis_replace por longitud del substring (item 2 tupla - item 1 tupla) de mayor a menor
        lista_tuplas_parentesis_replace = sorted(lista_tuplas_parentesis_replace, key = lambda x: -x[2])

        #se hace el replace en el string para quitar lo encapsulado entre parentesis
        for tupla in lista_tuplas_parentesis_replace:

            item_replace = "(" + string_param[tupla[0] + 1:tupla[1]] + ")"
            string_param = string_param.replace(item_replace, "")

        resultado_funcion = string_param



    elif opcion == "LISTA_DECLARACIONES_VARIABLES_LOCALES":

        #la opcion permite calcular una lista que contiene las lineas de codigo asociadas con declaraciones de variables locales de rutinas / funciones
        #se calcula en base a las listas siguientes que se extraen del parametro df_bbdd_codigo_param:
        #
        # --> listas sino: son las lineas que empiezan por "Dim " o "Const " y que NO acaban por " _" 
        #                  se agregan directamente a la lista resultado de la funcion
        #
        # --> listas sisi: son las lineas que empiezan por "Dim " o "Const " que SI acaban por " _" 
        #                  se agregan directamente a la lista resultado de la funcion
        #
        # --> listas nosi: son las lineas que NO empiezan por "Dim " o "Const " que SI acaban por " _"
        #                  se agregan a la lista resultado de la funcion si el indice de la linea de codigo anterior esta en las listas sisi o nosi
        #
        # --> listas nono: son las lineas que NO empiezan por "Dim " o "Const " que NO acaban por " _"
        #                  se agregan a la lista resultado de la funcion si el indice de la linea de codigo anterior esta en las listas sisi o nosi
        #########################
        #parametros kwargs usados --> indice_param_ini
        #                         --> indice_param_fin
        #                         --> df_bbdd_codigo_param


        resultado_funcion = []


        ####################################################################
        # AJUSTES PREVIOS
        ####################################################################

        #para variables locales de rutinas se fusionan las listas de inicios de declaracion configurados en el diccionario dicc_objetos_vba 
        #y se converte en tupla (si no el metodo startswith no va en busquedas multiples), cada item se trimea y se agrega un espacio en blanco al final
        lista_temp = dicc_objetos_vba["RUTINAS_VARIABLES_LOCALES"]["LISTA_DECLARACION_INICIO_NO_ES_CONSTANTE"] + dicc_objetos_vba["RUTINAS_VARIABLES_LOCALES"]["LISTA_DECLARACION_INICIO_SI_ES_CONSTANTE"]
        tupla_declaracion_ini_variable_local_rutina = tuple([item.strip() + " " for item in lista_temp])
        

        #para el calculo de las lineas de codigo que son declaraciones de variables locales, se crean las listas de indices siguientes:
        # --> lista_indices_sino      son las listas listas sino (ver comentario mas arriba) donde los indices estan incluidos dentro de los de la rutina / funcion donde se calcula
        # --> lista_indices_sisi      son las listas listas sisi (ver comentario mas arriba) donde los indices estan incluidos dentro de los de la rutina / funcion donde se calcula
        # --> lista_indices_nosi      son las listas listas nosi (ver comentario mas arriba) donde los indices estan incluidos dentro de los de la rutina / funcion donde se calcula
        # --> lista_indices_nono      son las listas listas nono (ver comentario mas arriba) donde los indices estan incluidos dentro de los de la rutina / funcion donde se calcula
        #
        #la presente opcion LISTA_DECLARACIONES_VARIABLES_LOCALES de la funcion que se comenta se usa en la funcion func_access_bbdd_codigo_calculos_varios
        #a nivel de modulo y de rutina / funcion donde ya se conocen los indices de codigo correspondientes a declaraciones de inicio y fin de la rutina / funcion donde se calcula

        df_sino = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["INDICE"] > indice_param_ini) & (df_bbdd_codigo_param["INDICE"] < indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_local_rutina)) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])

        df_sisi = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["INDICE"] > indice_param_ini) & (df_bbdd_codigo_param["INDICE"] < indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_local_rutina)) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])

        df_nosi = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["INDICE"] > indice_param_ini) & (df_bbdd_codigo_param["INDICE"] < indice_param_fin) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_local_rutina)) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])

        df_nono = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["INDICE"] > indice_param_ini) & (df_bbdd_codigo_param["INDICE"] < indice_param_fin) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_local_rutina)) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])           

        df_sino.reset_index(drop = True, inplace = True)
        df_sisi.reset_index(drop = True, inplace = True)
        df_nosi.reset_index(drop = True, inplace = True)
        df_nono.reset_index(drop = True, inplace = True)


        lista_indices_sino = []#las completion list no van cuando el df esta vacio por lo que se cre la lista en varias etapas
        lista_indices_sisi = []#idem
        lista_indices_nosi = []#idem
        lista_indices_nono = []#idem

        if len(df_sino) != 0:
            lista_indices_sino = [[df_sino.iloc[ind, df_sino.columns.get_loc("INDICE")]
                                , df_sino.iloc[ind, df_sino.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_sino.index]

        if len(df_sisi) != 0:
            lista_indices_sisi = [[df_sisi.iloc[ind, df_sisi.columns.get_loc("INDICE")]
                                , df_sisi.iloc[ind, df_sisi.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_sisi.index]

        if len(df_nosi) != 0:
            lista_indices_nosi = [[df_nosi.iloc[ind, df_nosi.columns.get_loc("INDICE")]
                                , df_nosi.iloc[ind, df_nosi.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_nosi.index]

        if len(df_nono) != 0:
            lista_indices_nono = [[df_nono.iloc[ind, df_nono.columns.get_loc("INDICE")]
                                , df_nono.iloc[ind, df_nono.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_nono.index]

        del df_sino
        del df_sisi
        del df_nosi
        del df_nono


        #se re-ordenan las lista por el indice de codigo
        lista_indices_sino = sorted(lista_indices_sino, key = lambda x: x[0])
        lista_indices_sisi = sorted(lista_indices_sisi, key = lambda x: x[0])
        lista_indices_nosi = sorted(lista_indices_nosi, key = lambda x: x[0])
        lista_indices_nono = sorted(lista_indices_nono, key = lambda x: x[0])



        ####################################################################
        # CALCULOS
        ####################################################################

        #se agregan directamente los items de lista_indices_sino
        #donde se calcula previamente si declaracion de constante o no
        for indice_sino, linea_codigo_sino in lista_indices_sino:   
            resultado_funcion.append(linea_codigo_sino)


        #se agregan directamente los items de lista_indices_sisi donde se calcula previamente si declaracion de constante o no
        #y se inicia bucle while (hasta el el primer indice de linea que no empieza por el inicio de declaracion 
        #de variable publica o local, segun la opcion seleccionada, y que no acaba por " _")
        #donde se agregan todas las lineas que NO empiezan por el inicio de declaracion y acaban en " _" siempre y cuando el indice de linea
        #sea consecutivo

        for indice_sisi, linea_codigo_sisi in lista_indices_sisi:

            #se agregan los datos a la lista resultado de la funcion (la linea de codigo se neteea quitando el " _" final
            #siempre y cuando el resultado del neteo no sea de longitud = 0)
            if len(linea_codigo_sisi[:-2]) != 0:
                resultado_funcion.append(linea_codigo_sisi[:-2])

            #se realiza bucle desde el indice siguiente al de la lista indice_si_si hasta el 1er indice de la lista indice_linea_no_no
            #y se agregan los indices de indice_linea_no_si siempre y cuando el indice es inmediatamentente consecutivo al de la iteracion anterior
            #se finaliza el bucle cuando se llega al 1er indice de la lista indice_linea_no_no mencionado
            lista_indices_nosi_iteracion = [[indice_nosi, linea_codigo_nosi] for indice_nosi, linea_codigo_nosi in lista_indices_nosi if indice_nosi > indice_sisi]
            lista_indice_nono_despues_iteracion = [[indice_nono, linea_codigo_nono] for indice_nono, linea_codigo_nono in lista_indices_nono if indice_nono > indice_sisi]

            primer_indice_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][0]
            linea_codigo_primer_indice_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][1]

            indice_linea_siguiente = indice_sisi + 1
            while indice_linea_siguiente <= primer_indice_nono_despues_iteracion:

                for indice_nosi, linea_codigo_nosi in lista_indices_nosi_iteracion:

                    if indice_linea_siguiente == indice_nosi:

                        #se agregan los datos a la lista resultado de la funcion (la linea de codigo se neteea quitando el " _" final
                        #siempre y cuando el resultado del neteo no sea de longitud = 0)
                        if len(linea_codigo_nosi[:-2]) != 0:
                            resultado_funcion.append(linea_codigo_nosi[:-2])

                        break

                if indice_linea_siguiente == primer_indice_nono_despues_iteracion:

                    #se agregan los datos a la lista resultado de la funcion
                    resultado_funcion.append(linea_codigo_primer_indice_nono_despues_iteracion)
                    break

                indice_linea_siguiente = indice_linea_siguiente + 1

        del lista_indices_sino
        del lista_indices_sisi 
        del lista_indices_nosi
        del lista_indices_nono


 
    elif opcion == "LISTA_DECLARACIONES_VARIABLES_PUBLICAS_NORMALES":

        #la opcion permite calcular para las variables publicas normales una lista de listas donde cada sublista contiene 
        # --> tipo modulo
        # --> nombre modulo
        # --> indice linea de codigo
        # --> indicador de si la linea de codigo es de declaracion de constante o no
        # --> linea de codigo
        # 
        #se calcula en base a las listas siguientes que se extraen del parametro df_bbdd_codigo_param:
        #
        # --> listas sino: son las lineas que empiezan por incios de declaracion de variables publicas normales y que NO acaban por " _" 
        #                  se agregan directamente a la lista resultado de la funcion
        #
        # --> listas sisi: son las lineas que empiezan por incios de declaracion de variables publicas normales y que SI acaban por " _" 
        #                  se agregan directamente a la lista resultado de la funcion
        #
        # --> listas nosi: son las lineas que NO empiezan por incios de declaracion de variables publicas normales y que SI acaban por " _"
        #                  se agregan a la lista resultado de la funcion si el indice de la linea de codigo anterior esta en las listas sisi o nosi
        #
        # --> listas nono: son las lineas que NO empiezan por incios de declaracion de variables publicas normales y que NO acaban por " _"
        #                  se agregan a la lista resultado de la funcion si el indice de la linea de codigo anterior esta en las listas sisi o nosi
        #########################
        #parametros kwargs usados --> indice_param_ini
        #                         --> indice_param_fin
        #                         --> df_bbdd_codigo_param


        resultado_funcion = []


        ####################################################################
        # AJUSTES PREVIOS
        ####################################################################

        #se fusionan todos los inicios de declaracion en una lista de sublistas
        #donde cada sublista:
        # --> item 1: es la key del diccionario (NO_ES_CONSTANTE o SI_ES_CONSTANTE)
        # --> item 2: es la declaracion de inicio (se trimea y se agrega un espacio en blanco al final, por si el usuario ha configurado mal el diccionario)
        #se reordena la lista por el item 2 de cada sublista por longitud de mayor a menor
        lista_variables_publicas_normales_declar_ini_ord_desc = [[key_2.replace("LISTA_DECLARACION_INICIO_", ""), item.strip() + " "] 
                                                                for key_2 in dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"].keys() 
                                                                for item in dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"][key_2] if "LISTA_DECLARACION_INICIO" in key_2]

        lista_variables_publicas_normales_declar_ini_ord_desc.sort(key = lambda x: len(x[1]), reverse = True)


        #se convierte la lista lista_variables_publicas_normales_declar_ini_ord_desc (solo para el item 2 de las sublistas) en tupla 
        #(si no el metodo startswith no va en busquedas multiples en df), cada item se trimea y se agrega un espacio en blanco al final
        tupla_declaracion_ini_variable_publica_normal = tuple([item[1] for item in lista_variables_publicas_normales_declar_ini_ord_desc])


        #para el calculo de las lineas de codigo que son declaraciones de variables publicas normales, se crean las listas de indices siguientes:
        # --> lista_indices_si_no      son las listas listas sino (ver comentario mas arriba) donde los indices estan incluidos dentro del modulo donde se calcula (excluyendo las rutinas / funciones)
        # --> lista_indices_si_si      son las listas listas sisi (ver comentario mas arriba) donde los indices estan incluidos dentro del modulo donde se calcula (excluyendo las rutinas / funciones)
        # --> lista_indices_no_si      son las listas listas nosi (ver comentario mas arriba) donde los indices estan incluidos dentro del modulo donde se calcula (excluyendo las rutinas / funciones)
        # --> lista_indices_no_no      son las listas listas nono (ver comentario mas arriba) donde los indices estan incluidos dentro del modulo donde se calcula (excluyendo las rutinas / funciones)
        #
        #la presente opcion LISTA_DECLARACIONES_PUBLICAS_NORMALES de la funcion que se comenta se usa en la funcion func_access_bbdd_codigo_calculos_varios
        #a nivel de modulo donde ya se conocen los indices de codigo correspondientes al inicio y fin del modulo
        df_sino = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_publica_normal)) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
        
        df_sisi = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_publica_normal)) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
        
        df_nosi = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_publica_normal)) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
        
        df_nono = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_declaracion_ini_variable_publica_normal)) &
                                                (~df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.endswith(indicador_linea_codigo_vba_truncada_a_linea_siguiente)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
        
        df_sino.reset_index(drop = True, inplace = True)
        df_sisi.reset_index(drop = True, inplace = True)
        df_nosi.reset_index(drop = True, inplace = True)
        df_nono.reset_index(drop = True, inplace = True)


        lista_indices_sino = []#las completion list no van cuando el df esta vacio por lo que se cre la lista en varias etapas
        lista_indices_sisi = []#idem
        lista_indices_nosi = []#idem
        lista_indices_nono = []#idem

        if len(df_sino) != 0:
            lista_indices_sino = [[df_sino.iloc[ind, df_sino.columns.get_loc("TIPO_MODULO")]
                                    , df_sino.iloc[ind, df_sino.columns.get_loc("NOMBRE_MODULO")]
                                    , df_sino.iloc[ind, df_sino.columns.get_loc("INDICE")]
                                    , df_sino.iloc[ind, df_sino.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_sino.index]

        if len(df_sisi) != 0:
            lista_indices_sisi = [[df_sisi.iloc[ind, df_sisi.columns.get_loc("TIPO_MODULO")]
                                    , df_sisi.iloc[ind, df_sisi.columns.get_loc("NOMBRE_MODULO")]
                                    , df_sisi.iloc[ind, df_sisi.columns.get_loc("INDICE")]
                                    , df_sisi.iloc[ind, df_sisi.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_sisi.index]

        if len(df_nosi) != 0:
            lista_indices_nosi = [[df_nosi.iloc[ind, df_nosi.columns.get_loc("TIPO_MODULO")]
                                    , df_nosi.iloc[ind, df_nosi.columns.get_loc("NOMBRE_MODULO")]
                                    , df_nosi.iloc[ind, df_nosi.columns.get_loc("INDICE")]
                                    , df_nosi.iloc[ind, df_nosi.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_nosi.index]

        if len(df_nono) != 0:
            lista_indices_nono = [[df_nono.iloc[ind, df_nono.columns.get_loc("TIPO_MODULO")]
                                    , df_nono.iloc[ind, df_nono.columns.get_loc("NOMBRE_MODULO")]
                                    , df_nono.iloc[ind, df_nono.columns.get_loc("INDICE")]
                                    , df_nono.iloc[ind, df_nono.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]] for ind in df_nono.index]

        del df_sino
        del df_sisi
        del df_nosi
        del df_nono


        #se re-ordenan las lista por el tipo y nombre de modulo y por el indice de codigo
        lista_indices_sino = sorted(lista_indices_sino, key = lambda x: (x[0], x[1], x[2]))
        lista_indices_sisi = sorted(lista_indices_sisi, key = lambda x: (x[0], x[1], x[2]))
        lista_indices_nosi = sorted(lista_indices_nosi, key = lambda x: (x[0], x[1], x[2]))
        lista_indices_nono = sorted(lista_indices_nono, key = lambda x: (x[0], x[1], x[2]))


        ####################################################################
        # CALCULOS
        ####################################################################

        #se agregan directamente los items de lista_indices_sino
        #donde se calcula previamente si declaracion de constante o no
        for tipo_modulo_sino, nombre_modulo_sino, indice_sino, linea_codigo_sino in lista_indices_sino:

            #se calcula si la linea de declaracion de constante o no mediante la lista
            #lista_variables_publicas_normales_declar_ini_ord_desc, calculada en los ajustes
            es_constante = None
            for item_config in lista_variables_publicas_normales_declar_ini_ord_desc:
                es_constante_config = item_config[0]
                declar_ini_config = item_config[1]

                if declar_ini_config in linea_codigo_sino:
                    es_constante = es_constante_config
                    break
                
            #se agregan los datos de la iteracion a la lista resultado de la funcion
            resultado_funcion.append([tipo_modulo_sino
                                    , nombre_modulo_sino
                                    , indice_sino
                                    , es_constante
                                    , linea_codigo_sino])


        #se agregan directamente los items de lista_indices_sisi donde se calcula previamente si declaracion de constante o no
        #y se inicia bucle while (hasta el el primer indice de linea que no empieza por el inicio de declaracion 
        #de variable publica o local, segun la opcion seleccionada, y que no acaba por " _")
        #donde se agregan todas las lineas que NO empiezan por el inicio de declaracion y acaban en " _" siempre y cuando el indice de linea
        #sea consecutivo
        for tipo_modulo_sisi, nombre_modulo_sisi, indice_sisi, linea_codigo_sisi in lista_indices_sisi:

            #se calcula si la linea de declaracion de constante o no mediante la lista
            #lista_variables_publicas_normales_declar_ini_ord_desc, calculada en los ajustes
            es_constante = None
            for item_config in lista_variables_publicas_normales_declar_ini_ord_desc:
                es_constante_config = item_config[0]
                declar_ini_config = item_config[1]

                if declar_ini_config in linea_codigo_sisi:
                    es_constante = es_constante_config
                    break


            #se agregan los datos a la lista resultado de la funcion (la linea de codigo se neteea quitando el " _" final
            #siempre y cuando el resultado del neteo no sea de longitud = 0)
            if len(linea_codigo_sisi[:-2]) != 0:

                resultado_funcion.append([tipo_modulo_sisi
                                        , nombre_modulo_sisi
                                        , indice_sisi
                                        , es_constante
                                        , linea_codigo_sisi[:-2]])
                


            #se realiza bucle desde el indice siguiente al de la lista indice_sisi hasta el 1er indice de la lista indice_linea_nono
            #y se agregan los indices de indice_linea_nosi siempre y cuando el indice es inmediatamentente consecutivo al de la iteracion anterior
            #se finaliza el bucle cuando se llega al 1er indice de la lista indice_linea_n_no mencionado
            lista_indices_nosi_iteracion = [[tipo_modulo_nosi
                                            , nombre_modulo_nosi
                                            , indice_nosi
                                            , linea_codigo_nosi]
                                            for tipo_modulo_nosi, nombre_modulo_nosi, indice_nosi, linea_codigo_nosi in lista_indices_nosi
                                            if indice_nosi > indice_sisi]

            lista_indice_nono_despues_iteracion = [[tipo_modulo_nono
                                                    , nombre_modulo_nono
                                                    , indice_nono
                                                    , linea_codigo_nono]
                                                    for tipo_modulo_nono, nombre_modulo_nono, indice_nono, linea_codigo_nono in lista_indices_nono
                                                    if indice_nono > indice_sisi]


            tipo_modulo_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][0]
            nombre_modulo_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][1]
            primer_indice_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][2]
            linea_codigo_primer_indice_nono_despues_iteracion = lista_indice_nono_despues_iteracion[0][3]

            indice_linea_siguiente = indice_sisi + 1
            while indice_linea_siguiente <= primer_indice_nono_despues_iteracion:

                for tipo_modulo_nosi, nombre_modulo_nosi, indice_nosi, linea_codigo_nosi in lista_indices_nosi_iteracion:

                    if indice_linea_siguiente == indice_nosi:

                        #se agregan los datos a la lista resultado de la funcion (la linea de codigo se neteea quitando el " _" final
                        #siempre y cuando el resultado del neteo no sea de longitud = 0)
                        if len(linea_codigo_nosi[:-2]) != 0:

                            resultado_funcion.append([tipo_modulo_nosi
                                                    , nombre_modulo_nosi
                                                    , indice_nosi
                                                    , es_constante
                                                    , linea_codigo_nosi[:-2]])

                        break


                if indice_linea_siguiente == primer_indice_nono_despues_iteracion:

                    #se agregan los datos de la iteracion a la lista resultado de la funcion
                    resultado_funcion.append([tipo_modulo_nono_despues_iteracion
                                            , nombre_modulo_nono_despues_iteracion
                                            , primer_indice_nono_despues_iteracion
                                            , es_constante
                                            , linea_codigo_primer_indice_nono_despues_iteracion])
                    break

                indice_linea_siguiente = indice_linea_siguiente + 1

        del lista_indices_sino
        del lista_indices_sisi 
        del lista_indices_nosi
        del lista_indices_nono




    elif opcion == "LISTA_DECLARACIONES_VARIABLES_PUBLICAS_DEFINIDAS_USUARIO":
        #la opcion permite calcular para las variables publicas definidas por el usuario una lista de listas donde cada sublista contiene 
        # --> tipo modulo
        # --> nombre modulo
        # --> indice linea de codigo
        # --> flag unico indentificador de la variable publica definida por el usuario (todas las lineas de codigo que la componen
        #     tienen el mismo flag)
        # --> linea de codigo
        # 
        #se calcula en base a las listas siguientes que se extraen del parametro df_bbdd_codigo_param:
        # --> lista_variables_publicas_definidas_por_usuario_indices_ini    para un mismo modulo contiene todos los indices de codigo de los inicios de declaracion
        # --> lista_variables_publicas_definidas_por_usuario_indices_fin    para un mismo modulo contiene todos los indices de codigo de los finales de declaracion
        #
        #mediante bucle sobre la lista lista_variables_publicas_definidas_por_usuario_indices_ini se recupera el indice de codigo de inicio de declaracion y para este
        #se busca el indicie de codigo de fin de declaracion inmediatamente posterior en la lista lista_variables_publicas_definidas_por_usuario_indices_fin
        #disponiendo de los indices de inicio y fin se realiza extracion del df df_bbdd_codigo_param (parametro de la funcion) entre estos indices y mediante
        #bucle sobre los indices de este df extraido se agregan los datos a la lista resultado de la funcion (tipo y nombre de modulo, indice, linea de codigo y 
        #flag unico que se calcula mediante un contador dentro de la iteracion sobre el bucle df_bbdd_codigo_paraminicial linea de codigo)
        #########################
        #parametros kwargs usados --> indice_param_ini
        #                         --> indice_param_fin
        #                         --> df_bbdd_codigo_param

        resultado_funcion = []

        ####################################################################
        # AJUSTES PREVIOS
        ####################################################################

        #se calculan tuplas sobre las listas almacenadas en las subkeys_2 (LISTA_DECLARACION_INICIO y LISTA_DECLARACION_FIN) asociadas 
        #a la key_1 (VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO) del diccionario dicc_objetos_vbapara que los metodos startswith y isin de filtrado por varios elementos en df funcionen
        #los items de la tupla sobre los inicios de declaraciones se trimean y se les agrega un espacion en blaqnco al final (por si el usuario los configura mal en el diccionario)
        tupla_variable_definida_por_usuario_ini = tuple([item.strip() + " " for item in dicc_objetos_vba["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO"]["LISTA_DECLARACION_INICIO"]])
        tupla_variable_definida_por_usuario_fin = tuple([item.strip() for item in dicc_objetos_vba["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO"]["LISTA_DECLARACION_FIN"]])



        #se genera la lista lista_variables_publicas_definidas_por_usuario_indices_ini y lista_variables_publicas_definidas_por_usuario_indices_fin
        #que almacenan resperctivamente todos los inicios y final de declaracion de este tipo de variables publicas en los modulos en en el df de codigo
        df_temp_1 = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.startswith(tupla_variable_definida_por_usuario_ini)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE"]])#aqui es startswith
        
        df_temp_2 = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) &
                                                (df_bbdd_codigo_param["INDICE"] >= indice_param_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_param_fin) &
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].isin(tupla_variable_definida_por_usuario_fin)) & 
                                                (df_bbdd_codigo_param["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                                ["INDICE"]])#aqui es isin

        df_temp_1.reset_index(drop = True, inplace = True)
        df_temp_2.reset_index(drop = True, inplace = True)



        lista_variables_publicas_definidas_por_usuario_indices_ini = []#las completion list no van cuando el df esta vacio por lo que se cre la lista en varias etapas
        lista_variables_publicas_definidas_por_usuario_indices_fin = []#idem

        lista_variables_publicas_definidas_por_usuario_indices_ini = [df_temp_1.iloc[ind, df_temp_1.columns.get_loc("INDICE")] for ind in df_temp_1.index]
        lista_variables_publicas_definidas_por_usuario_indices_fin = [df_temp_2.iloc[ind, df_temp_2.columns.get_loc("INDICE")] for ind in df_temp_2.index]


        #se ordenan las listas por el indice de la linea de codigo
        lista_variables_publicas_definidas_por_usuario_indices_ini = sorted(lista_variables_publicas_definidas_por_usuario_indices_ini)
        lista_variables_publicas_definidas_por_usuario_indices_fin = sorted(lista_variables_publicas_definidas_por_usuario_indices_fin)

        del df_temp_1
        del df_temp_2


        ####################################################################
        # CALCULOS
        ####################################################################

        #el calculo se hace a nivel de modulo segun los indices de codigo de inicio y fin (parametros de la funcion)
        #cuando se ejecuta la opcion LISTA_DECLARACIONES_VARIABLES_PUBLICAS_DEFINIDAS_USUARIO de la presente funcion
        #en la funcion func_access_bbdd_codigo_calculos_varios
        for ind_ini, indice_ini in enumerate(lista_variables_publicas_definidas_por_usuario_indices_ini):

            #se calcula el indice de codigo correspondiente a fin de declaracion de variable publica definida por el usuario 
            #inmediatamente posterior al indice de inicio de declaracion
            indice_fin = [indice_fin for indice_fin in lista_variables_publicas_definidas_por_usuario_indices_fin if indice_fin > indice_ini][0]
            

            #se calcula el flag unico dentro de cada modula para cada variable publica definida por el usuario
            flag_variable_publica_definida_usuario = "DEFINIDA_POR_USUARIO_" + str(ind_ini + 1)


            #se extrae del parametro df_variables_publicas los indices y lineas de codigo entre los indices de inicio y fin de declaracion
            df_temp = (df_bbdd_codigo_param.loc[(df_bbdd_codigo_param["NOMBRE_RUTINA"].isnull() == True) & 
                                                (df_bbdd_codigo_param["INDICE"] >= indice_ini) & (df_bbdd_codigo_param["INDICE"] <= indice_fin), 
                                                ["TIPO_MODULO", "NOMBRE_MODULO", "INDICE", "CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]])
            
            df_temp.reset_index(drop = True, inplace = True)


            #mediante bucle sobre los indices del df extraido se completa la lista resultado de la funcion
            for ind in df_temp.index:

                resultado_funcion.append([df_temp.iloc[ind, df_temp.columns.get_loc("TIPO_MODULO")]
                                        , df_temp.iloc[ind, df_temp.columns.get_loc("NOMBRE_MODULO")]
                                        , df_temp.iloc[ind, df_temp.columns.get_loc("INDICE")]
                                        , flag_variable_publica_definida_usuario
                                        , df_temp.iloc[ind, df_temp.columns.get_loc("CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS")]])

            del df_temp

        del lista_variables_publicas_definidas_por_usuario_indices_ini
        del lista_variables_publicas_definidas_por_usuario_indices_fin



    elif opcion == "LISTA_VARIABLES_PUBLICAS_NO_ES_CONSTANTE":
        #la opcion permite calcular para una linea de declaracion listar en una lista todas las variables publicas normales 
        #que NO son constantes que la componen, esta lista se compone de sublistas:
        # --> nombre variable publica                        aplica
        # --> tipo variable publica                          valor subkey_2 (STRING_PARA_LISTADO_EXCEL) asociada a key_1 (VARIABLE_PUBLICA_NORMAL) del diccionario dicc_objetos_vba
        # --> tipo dato publica                              aplica
        # --> es constante                                   NO
        # --> valor constante                                None (no aplica)
        # --> subvariables variable definida por el usuario  None (no aplica)
        #
        #se realizan ajustes especificos en caso de declaracion de arrays, ver en los comentarios)
        #########################
        #parametros kwargs usados -> linea_codigo


        resultado_funcion = []

        #antes de splitear la linea de codigo por la coma ",", hay que localizar si se han declarado variable tipo arrays de mas de 1 dimension
        #pq estos vienen encapsulados entre parentesis y dentro hay una coma por lo que si se splitea directamente sin pasar por este ajuste
        #generaria errores en el diagnostico de dependencias
        #
        #los pasos a seguir son:
        #--> localizar en la linea de codigo si hay parentesis de apertura y cierre
        #--> en caso de que haya en ambos se prosigue el proceso sino se para y se pasa a splitear directamente por la coma
        #
        #los pasos a seguir de que haya ambos:
        #--> crear lista de tuplas donde cada 1 contiene los indices de posicion de los parentesis de apertura y los de los parentesis de cierre 
        #    inmediatamente consecutivos
        #--> localizar los indices de posicion de las comas ","
        #--> reemplazar las comas por @ cuyos indices de posicion se encuentren entre los items de las tuplas comentadas
        #    (@ es un caracter prohibido en la declaracion de variables publicas por lo que no hay riesgo de alterar ni el nombre ni en tipo de dato de ninguna)
        #--> a partir de aqui ya se puede splitear por la coma

        lista_indices_parentesis_apertura = list(re.finditer(r"\(", linea_codigo))
        lista_indices_parentesis_cierre= list(re.finditer(r"\)", linea_codigo))

        if len(lista_indices_parentesis_apertura) != 0 and len(lista_indices_parentesis_apertura) == len(lista_indices_parentesis_cierre):

            #en caso de que las 2 listas no son vacias y son de la misma longitud, mediante el metodo zip, se crea una lista de tuplas donde
            #cada tupla indica en el item 1 el indice de apertura y en el item 2 el indice de cierre
            lista_indices_parentesis_apertura = [lista_indices_parentesis_apertura[ind].start() for ind, item in enumerate(lista_indices_parentesis_apertura)]
            lista_indices_parentesis_cierre = [lista_indices_parentesis_cierre[ind].start() for ind, item in enumerate(lista_indices_parentesis_cierre)]


            #se crea la lista de tuplas con los indices de posicion de de los parentesis de apertura y cierre
            lista_tuplas_indices_parentesis = list(zip(lista_indices_parentesis_apertura, lista_indices_parentesis_cierre))


            #se localizan los indices de posicion de las comas que se han de reemplazar por @
            lista_indices_comas = list(re.finditer(",", linea_codigo))
            lista_indices_comas = [lista_indices_comas[ind].start() for ind, item in enumerate(lista_indices_comas)]

            lista_indices_comas_replace = [indice_coma for indice_coma in lista_indices_comas 
                                        for tupla_indices in lista_tuplas_indices_parentesis 
                                        if indice_coma >= tupla_indices[0] and indice_coma <= tupla_indices[1]]


            #se realiza el replace de la coma por @
            linea_codigo = "".join([caracter if ind not in lista_indices_comas_replace else "@" for ind, caracter in enumerate(linea_codigo)])


        #se splitea la linea de codigo ajustada por la coma y en cada item de la lista resultante se splitea por " As "  (excluyendo los items de longitud = 0)
        #para tener el nombre de la variable publica (1er item) y su tipo de dato (2ndo item)
        #para los casos de declaraciones de arrays se pasa lo que hay entre los parentesis "(" y ")" (incluyendo ambos)
        #del nombre de la variable publica hacia su tipo de dato (concatenandolo con el existente)
        linea_codigo_spliteada_por_coma = linea_codigo.split(",")

        for item_spliteado_por_coma in linea_codigo_spliteada_por_coma:

            if len(item_spliteado_por_coma) != 0:

                #se hace tambien el replace de @ por la coma
                item_spliteado_por_coma = item_spliteado_por_coma.replace("@", ",")


                #se localiza si item_spliteado_por_coma tiene parentesis "(" y ")" incluidos
                lista_indices_parentesis_apertura = list(re.finditer(r"\(", item_spliteado_por_coma))
                lista_indices_parentesis_cierre= list(re.finditer(r"\)", item_spliteado_por_coma))


                #se almacena lo incluido entre los parentesis "(" y ")" en el string string_tipo_dato_arrays
                #para poder concatenarlo al tipo de dato tras splitear item_spliteado_por_coma por " As " (mas adelante)
                string_tipo_dato_arrays = ""
                if len(lista_indices_parentesis_apertura) != 0 and len(lista_indices_parentesis_apertura) == len(lista_indices_parentesis_cierre):

                    indice_parentesis_apertura = [lista_indices_parentesis_apertura[ind].start() for ind, item in enumerate(lista_indices_parentesis_apertura)][0]
                    indice_parentesis_cierre = [lista_indices_parentesis_cierre[ind].start() for ind, item in enumerate(lista_indices_parentesis_cierre)][0]

                    string_tipo_dato_arrays = item_spliteado_por_coma[indice_parentesis_apertura:indice_parentesis_cierre + 1]

  
                #se splitea item_spliteado_por_coma por " As ", el 1er item de la lista resultante es el nombre de la variable y el 2ndo su tipo de dato
                #que se concatena con string_tipo_dato_arrays 
                #(si el array de de 1 una dimension, esdecir que sale"()" despues del nombre de variable, este no se concatena al tipo de dato)
                item_spliteado_por_coma_y_por_as = item_spliteado_por_coma.split(indicador_variable_vba_declaracion_tipo_dato)
                
                variable_publica_nombre = item_spliteado_por_coma_y_por_as[0].replace(string_tipo_dato_arrays, "").strip()

                variable_publica_tipo_dato = (item_spliteado_por_coma_y_por_as[1] + string_tipo_dato_arrays.replace("()", "")
                                                if len(item_spliteado_por_coma_y_por_as) == 2 else string_tipo_dato_arrays.replace("()", ""))


                #se agregan los datos a la lista resultado de la funcion
                variable_publica_tipo = dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"]["STRING_PARA_LISTADO_EXCEL"]
                variable_publica_es_constante = "NO"
                variable_publica_es_constante_valor = None
                variable_publica_definida_usuario_subvariables = None

                resultado_funcion.append([variable_publica_nombre
                                        , variable_publica_tipo
                                        , variable_publica_tipo_dato
                                        , variable_publica_es_constante
                                        , variable_publica_es_constante_valor
                                        , variable_publica_definida_usuario_subvariables])




    elif opcion == "LISTA_VARIABLES_PUBLICAS_SI_ES_CONSTANTE":
        #la opcion permite calcular para una linea de declaracion listar en una lista todas las variables publicas normales 
        #que SI son constantes que la componen, esta lista se compone de sublistas:
        # --> nombre variable publica                        aplica
        # --> tipo variable publica                          valor subkey_2 (STRING_PARA_LISTADO_EXCEL) asociada a key_1 (VARIABLE_PUBLICA_NORMAL) del diccionario dicc_objetos_vba
        # --> tipo dato publica                              None (no aplica)
        # --> es constante                                   SI
        # --> valor constante                                aplica
        # --> subvariables variable dinida por el usuario    None (no aplica)
        #
        #se realizan ajustes especificos en caso de declaracion de arrays, ver en los comentarios)
        #########################
        #parametros kwargs usados -> linea_codigo


        resultado_funcion = []

        #antes de splitear la linea de codigo por la coma "," y posteriormente por el signo "=", hay que localizar si se valores de constantes en formato string
        #(los que vienen encapsulados entre comillas dobles) pq dentro de estos valores podria haber comas "," y signos "=" lo que
        #generaria errores en el diagnostico de dependencias
        #
        #los pasos a seguir son:
        #--> localizar en la linea de codigo si hay valores de constantes encapsulados entre comillas dobles
        #--> en caso de que los haya se prosigue el proceso sino se para y se pasa a splitear directamente por la coma
        #
        #los pasos a seguir en caso de que los haya:
        #--> crear lista con todo lo encapsulado entre comillas dobles
        #--> para cada item de la lista del paso anterior se le asigna un valor de replace unico precedido de un @ + un indice
        #    el valor del replace incluye las comillas dobles, es decir "hola" se reemplazaria por ejemplo por @1 y no por "@1" pq @1 podria estar dentro del valor de la constante
        #--> reemplazar los valores de las constantes string por lo comentado en los 2 pasos anteriores en la linea de codigo
        #
        #--> a partir de aqui ya se puede splitear por la coma
        #--> con la lista resultante del paso anterior se splitea cada item por el signo "=" para obtener el nombre de la variable (1er item) 
        #    y su valor (2ndo item al cual se le hace replace de los valores string por su valor original)


        #se localizan si hay string encapsulados entre comillas dobles y se crea la lista con los valores replace
        lista_constantes_string = re.findall(r'"([^"]*)"', linea_codigo)
        lista_constantes_string_con_valor_replace = [[item, "@" + str(ind + 1)] for ind, item in enumerate(lista_constantes_string)]


        #se hace el replace (en caso de que haya valores de constantes duplicados en varias constantes declaradas en la misma linea
        #el valor replace del 1ero string que encuentra no se replica a los demas pq el metodo replace se limita al 1ero que encuentra no a todos, es el ", 1")
        for valor_constante_string, valor_replace in lista_constantes_string_con_valor_replace:
            linea_codigo = linea_codigo.replace("\"" + valor_constante_string + "\"", valor_replace, 1)


        #se splitea la linea de codigo ajustada por la coma y mediante bucle sobre la lista resultante (excluyendo los items de longitud = 0)
        #se splitea por " = "para tener el nombre de la variable publica (1er item) y su valor (2ndo item)
        #a este ultimo se le hace el replace usando la lista lista_constantes_string_con_valor_replace para
        #recuperar el valor del string original
        linea_codigo_spliteada_por_coma = linea_codigo.split(",")

        for item_spliteado_por_coma in linea_codigo_spliteada_por_coma:
            
            if len(item_spliteado_por_coma) != 0:

                linea_codigo_spliteada_por_signo_igual = item_spliteado_por_coma.split(indicador_variable_signo_igual)

                #se calculan los datos de la variable publica
                variable_publica_nombre = linea_codigo_spliteada_por_signo_igual[0].strip()
                variable_publica_es_constante_valor = linea_codigo_spliteada_por_signo_igual[1].replace("\"", "").strip() if len(linea_codigo_spliteada_por_signo_igual) == 2 else None


                #se hace el replace de los valores string de constantes para recuperar el valor original
                for item_original, item_replace in lista_constantes_string_con_valor_replace:
                    if item_replace in variable_publica_es_constante_valor:
                        variable_publica_es_constante_valor = variable_publica_es_constante_valor.replace(item_replace, item_original)
                        break


                #se agrega una comilla simple al principio del valor de la constante si este valor empieza por una comilla simple
                #esto es pq en Excel cuando una celda empieza por una comilla simple visualmente no la muestra en pantalla
                #(a pesar de que si sale en la barra de formulas)
                variable_publica_es_constante_valor = "'" + variable_publica_es_constante_valor if variable_publica_es_constante_valor[0:1] == "'" else variable_publica_es_constante_valor


                #se agregan los datos a la lista resultado de la funcion
                variable_publica_tipo = dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"]["STRING_PARA_LISTADO_EXCEL"]
                variable_publica_tipo_dato = None
                variable_publica_es_constante = "SI"
                variable_publica_definida_usuario_subvariables = None

                resultado_funcion.append([variable_publica_nombre
                                        , variable_publica_tipo
                                        , variable_publica_tipo_dato
                                        , variable_publica_es_constante
                                        , variable_publica_es_constante_valor
                                        , variable_publica_definida_usuario_subvariables])



    elif opcion == "LISTA_VARIABLES_PUBLICAS_DEFINIDAS_POR_USUARIO":
        #la opcion permite unificar todas las lineas de codigo asociadas a una variable publica definida por el usuario (se declaran en varias lineas)
        #y preparar una lista de de lista en el mismo formato que las variables publicas normales:
        # --> nombre variable publica                        aplica
        # --> tipo variable publica                          valor subkey_2 (STRING_PARA_LISTADO_EXCEL) asociada a key_1 (VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO) del diccionario dicc_objetos_vba
        # --> tipo dato publica                              aplica
        # --> es constante                                   -
        # --> valor constante                                None (no aplica)
        # --> subvariables variable dinida por el usuario    aplica
        #########################
        #parametros kwargs usados --> lista_variable_definida_por_usuario_ini
        #                         --> lista_lineas_flag_variable_publica (es lista de lista donde cada sublista contiene el indice de codigo y la linea de codigo)

        resultado_funcion = []


        #se recuperan las lineas de codigo para calcular el nombre de la variable y su tipo de dato (se hace con el 2ndo item de la primera sublista 
        #de lista_lineas_flag_variable_publica) y para calcular la concatenacion de sus subvariables asociadas (se hace con las sublistas 2 2 hasta la penultima)
        primer_item_lista_lineas_flag_variable_publica = lista_lineas_flag_variable_publica[0][1]
        lista_lineas_flag_variable_publica_para_subvariables = lista_lineas_flag_variable_publica[1:-1]


        #se calcula el nombre de la variable publica con el 1er item de lista_lineas_flag_variable_publica
        #y su tipo de dato mediante bucle sobre lista_variable_definida_por_usuario_ini
        variable_publica_nombre = None
        variable_publica_tipo_dato = None
        
        for item_config in lista_variable_definida_por_usuario_ini:

            if item_config in primer_item_lista_lineas_flag_variable_publica:
                variable_publica_nombre = primer_item_lista_lineas_flag_variable_publica.replace(item_config, "").strip()
                variable_publica_tipo_dato = item_config.strip()
                break


        #se calcula la concatenacion de las subvariables asociadas donde se encapsula el tipo de dato o su valor (en su defecto) entre parentesis
        variable_publica_definida_usuario_subvariables = ""

        for ind, item in enumerate(lista_lineas_flag_variable_publica_para_subvariables):

            #tan solo se extrae la linea de codigo (2ndo item de cada sublista de lista_lineas_flag_variable_publica_para_subvariables)
            linea_codigo  = item[1]


            #se splitea la linea de codigo por " As " y por " = " por separado
            #las 2 listas no pueden tener conjuntamente longitud = 2 por lo que el valor entre parentesis 
            #a continuacion del nombre de la subvariable viene del item 2 de la lista que tenga longitud 2
            #(en caso que las 2 listas tengan longitud = 1 se coje el 1er item de la lista spliteada por " As "
            #es el mismo que el de la lista spliteado por " =" )
            linea_codigo_spliteada_por_as = linea_codigo.split(indicador_variable_vba_declaracion_tipo_dato)
            linea_codigo_spliteada_por_signo_igual = linea_codigo.split(indicador_variable_signo_igual)

            sub_variable_iteracion = ""
            if len(linea_codigo_spliteada_por_as) == 1 and len(linea_codigo_spliteada_por_signo_igual) == 1:
                sub_variable_iteracion = linea_codigo_spliteada_por_as[0].strip()
            
            elif len(linea_codigo_spliteada_por_as) == 2:
                sub_variable_iteracion = linea_codigo_spliteada_por_as[0].strip() + " (" + linea_codigo_spliteada_por_as[1].strip() + ")"

            elif len(linea_codigo_spliteada_por_signo_igual) == 2:
                sub_variable_iteracion = linea_codigo_spliteada_por_signo_igual[0].strip() + " (" + linea_codigo_spliteada_por_signo_igual[1].strip() + ")"

            variable_publica_definida_usuario_subvariables = sub_variable_iteracion if ind == 0 else variable_publica_definida_usuario_subvariables + ", " + sub_variable_iteracion
   

        #se agregan los datos a la lista resultado de la funcion
        variable_publica_tipo = dicc_objetos_vba["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO"]["STRING_PARA_LISTADO_EXCEL"]
        variable_publica_es_constante = "-"
        variable_publica_es_constante_valor = None

        resultado_funcion.append([variable_publica_nombre
                                , variable_publica_tipo
                                , variable_publica_tipo_dato
                                , variable_publica_es_constante
                                , variable_publica_es_constante_valor
                                , variable_publica_definida_usuario_subvariables])


    #resultado de la funcion
    return resultado_funcion




def func_diagnostico_access_tablas_codigo_preservar(opcion, linea_codigo):
    #funcion que permite preservar en el codigo VBA de las rutinas / funciones tan solo los string encapsulados entre comillas dobles
    #para poder realizar el diagnostico de dependencias de los objetos tipo tablas / vinculos
    #funciona con 2 opciones:
    # --> PRESERVAR_SI     preserva solo el codigo necesario para localizar tablas / vinculos en una linea de codigo VBA se realiza mediante el diccionario dicc_diagnostico_tablas_access_donde_buscar
    #                      se fusionan las 2 listas temporales, se concatenan los item y el resultado de la funcion es la linea de codigo a preservar donde se busca la tabla / vinculo en el diagnostico
    #
    # --> PRESERVAR_NO    preserva todos los string entre comillas dobles que no se preservan en la opcion PRESERVAR_SI para que el usuario pueda chequear si considera que es uso de tabla en codigo VBA
    #                     el metodo usado mediante el diccionario dicc_diagnostico_tablas_access_donde_buscar con la opcion PRESERVAR_SI de esta funcion no es perfecto pq puede darse el caso que la tabla
    #                     se use en un string asociado a una variable publica o local y a posteriori se use dicha variable en un string tipo sentencia sql como las configuradas en el diccionario menciondo
    #                     (ejemplo mi_variable = "mi_tabla" y luego hay una instruccion VBA tipo misql = "SELECT * FROM " & mi_variable & "WHERE (CAMPO = 1)")


    #se crea lista de todo lo incluido entre comillas dobles
    lista_string_entre_comillas_dobles = re.findall(r'"([^"]*)"', linea_codigo)
    lista_string_entre_comillas_dobles = [i for i in lista_string_entre_comillas_dobles if len(i) != 0] if isinstance(lista_string_entre_comillas_dobles, list) else []


    #se crea la lista (lista_SI_preservar_codigo_sql) de codigo a preservar que parezca a sentencias SQL
    #con el framework re se preserva de la linea de codigo solo lo incluido entre ""
    lista_SI_preservar_codigo_sql = []
    for item_codigo in lista_string_entre_comillas_dobles:
        
        item_codigo_mayusc_trimeado = item_codigo.upper().strip()

        for item_config in dicc_diagnostico_tablas_access_donde_buscar["LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL"]:

            #segun que los items de la lista almacenada en dicc_diagnostico_tablas_access_donde_buscar["LISTA_TEXTOS_SIMILARES_SENTENCIAS_SQL"]
            #contengan espacios en blanco de por medio se realiza el chequeo de unaforma u otra
            # --> si hay espacios en blanco de por medio se realiza el matcheo por "contiene"
            # --> en caso contario (es decir es una palabra sola) se realiza el matcheo con el framework re para buscar solo la palabra sola

            item_config_mayusc_trimeado = item_config.upper().strip()
            item_config_split_espacio_blanco_intermedio = item_config_mayusc_trimeado.split(" ")

            check_item_config = 0
            if len(item_config_split_espacio_blanco_intermedio) == 1:
                                
                check_si_item_config_es_palabra_sola = fr'(?<![\w_-]){re.escape(item_config_mayusc_trimeado)}(?![\w-])'
                check_si_item_config_es_palabra_sola = bool(re.search(check_si_item_config_es_palabra_sola, item_codigo_mayusc_trimeado))

                if check_si_item_config_es_palabra_sola:
                    check_item_config = 1

            else:
                if item_config_mayusc_trimeado in item_codigo_mayusc_trimeado:
                    check_item_config = 1

            if check_item_config == 1: 
                lista_SI_preservar_codigo_sql.append(item_codigo)
                break


    #se crea lista (lista_SI_preservar_codigo_instruccion_vba) de lo incluido entre comillas dobles "" inmediatamente posterior a los items de la lista 
    #contenida en la key LISTA_TEXTOS_INSTRUCCIONES_VBA_MANIP_TABLAS del diccionario dicc_diagnostico_tablas_access_donde_buscar
    #para ello se crean 2 listas:
    #     lista_indices_localiz_comillas_dobles    --> es una lista de tuplas donde cada tupla contiene el indice de inicio y el de fin de lo incluido entre comillas dobles "" (metodo zip)

    #     lista_indices_fin_instruccion_vba        --> es una lista que contiene el indice de fin de la instruccion VBA de todos los items de la lista 
    #                                                  contenida en la key LISTA_TEXTOS_INSTRUCCIONES_VBA_MANIP_TABLAS del diccionario dicc_diagnostico_tablas_access_donde_buscar
    #
    #iterando por cada item de lista_indices_fin_instruccion_vba (que se reordena previamente de menor a mayor) se localiza en lista_indices_localiz_comillas_dobles
    #que tupla tiene el indice de inicio inmediatamente posterior al item de lista_indices_fin_instruccion_vba y con la tupla se calcula el texto a preservar en la linea de codigo
    #y se almacena en lista_SI_preservar_codigo_instruccion_vba

    linea_codigo_mayusc = linea_codigo.upper()

    #lista_indices_localiz_comillas_dobles
    lista_indices_localiz_comillas_dobles_temp_1 = list(re.finditer("\"", linea_codigo_mayusc))
    lista_indices_localiz_comillas_dobles_temp_2 = [lista_indices_localiz_comillas_dobles_temp_1[ind].start() for ind, item in enumerate(lista_indices_localiz_comillas_dobles_temp_1)]

    lista_indices_localiz_comillas_dobles_temp_3_1 = [item for ind, item in enumerate(lista_indices_localiz_comillas_dobles_temp_2) if int(ind / 2) == ind / 2] 
    lista_indices_localiz_comillas_dobles_temp_3_2 = [item for ind, item in enumerate(lista_indices_localiz_comillas_dobles_temp_2) if int(ind / 2) != ind / 2]
    lista_tuplas_indices_localiz_comillas_dobles = list(zip(lista_indices_localiz_comillas_dobles_temp_3_1, lista_indices_localiz_comillas_dobles_temp_3_2))

    del lista_indices_localiz_comillas_dobles_temp_1
    del lista_indices_localiz_comillas_dobles_temp_2
    del lista_indices_localiz_comillas_dobles_temp_3_1
    del lista_indices_localiz_comillas_dobles_temp_3_2


    #lista_indices_fin_instruccion_vba
    lista_indices_fin_instruccion_vba = []
    for item_config in dicc_diagnostico_tablas_access_donde_buscar["LISTA_TEXTOS_INSTRUCCIONES_VBA_MANIP_TABLAS"]:

        item_config_mayusc_trimeado = item_config.upper().strip()

        lista_indices_fin_instruccion_vba_temp = list(re.finditer(item_config_mayusc_trimeado, linea_codigo_mayusc))

        if len(lista_indices_fin_instruccion_vba_temp) != 0:                                       
            for ind, item in enumerate(lista_indices_fin_instruccion_vba_temp):
                lista_indices_fin_instruccion_vba.append(lista_indices_fin_instruccion_vba_temp[ind].end())

        del lista_indices_fin_instruccion_vba_temp
 
    lista_indices_fin_instruccion_vba = sorted(lista_indices_fin_instruccion_vba)


    #lista_SI_preservar_codigo_instruccion_vba
    lista_SI_preservar_codigo_instruccion_vba = []
    for indice_fin_instruccion_vba in lista_indices_fin_instruccion_vba:
        tupla_indices = [tupla for tupla in lista_tuplas_indices_localiz_comillas_dobles if tupla[0] > indice_fin_instruccion_vba][0]

        indice_ini_entre_comillas_dobles = tupla_indices[0]
        indice_fin_entre_comillas_dobles = tupla_indices[1]

        #para que el texto encapsulado entre comillas dobles pueda ser considerado con instruccucion VBA de manipulacion de tablas
        #el indice_ini_entre_comillas_dobles ha de ser inmediatamente posterior al indice_fin_instruccion_vba de la iteracion 
        # del bucle sobre lista_indices_fin_instruccion_vba
        #se agrega el string entre indice_ini_entre_comillas_dobles y indice_fin_entre_comillas_dobles, al cual se le quitan las comillas dobles "
        if indice_ini_entre_comillas_dobles == indice_fin_instruccion_vba + 1:
            codigo_preservar = linea_codigo[indice_ini_entre_comillas_dobles:indice_fin_entre_comillas_dobles].replace("\"", "")
            lista_SI_preservar_codigo_instruccion_vba.append(codigo_preservar)

        del tupla_indices
        del indice_fin_entre_comillas_dobles
        del codigo_preservar


    lista_SI_preservar_codigo = lista_SI_preservar_codigo_sql + lista_SI_preservar_codigo_instruccion_vba
    lista_NO_preservar_codigo = [item for item in lista_string_entre_comillas_dobles if item not in lista_SI_preservar_codigo]


    #resultado de la funcion
    return "".join(lista_SI_preservar_codigo) if opcion == "PRESERVAR_SI" else "".join(lista_NO_preservar_codigo)


#################################################################################################################################################################################
##                     RUTINA MS ACCESS - COMUN A CONTROL DE VERSIONES Y AL DIAGNOSTICO
#################################################################################################################################################################################

def def_proceso_access_1_import(proceso_id, opcion_bbdd, path_bbdd):
    #realiza la importacion del codigo VBA del access seleccionado (opcion_bbdd) en un df (df_bbdd_codigo) 
    #que se almacena en dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]
    #para poder reutilizarlo en otras rutinas donde segun que se opte por el control de versiones o el diagnostico se le agragaran 
    #mas o menos columnas

    warnings.filterwarnings("ignore")

    try:

        #se crea la lista temporal lista_dicc_objetos que sirve de base para construir el df que se almacena 
        #en dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["lista_dicc_objetos"]
        #(es comuna a todos los tipos de objeto MS ACCESS y a los procesos de control de versiones y diagnostico de dependencias)
        #se usa varias veces en esta rutina
        lista_dicc_objetos = []


        #se abre el access y se accede a su codigo VBA
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.OpenCurrentDatabase(path_bbdd)

        current_bbdd = access_app.CurrentDb()
        vba_project = access_app.VBE.ActiveVBProject


        #se listan las librerias DLL usadas en el codigo VBA y se almacenan en dicc_codigos_bbdd
        lista_librerias_dll = [dll.Description for dll in vba_project.References]
        lista_librerias_dll = sorted(lista_librerias_dll)

        mod_gen.dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["LISTA_LIBRERIAS_DLL"] = lista_librerias_dll if len(lista_librerias_dll) != 0 else None



        ##################################################################################################################################################
        # se calcula el df df_bbdd_codigo resultante de la importacion de todos los codigos VBA del access seleccionado
        ##################################################################################################################################################        

        #se concatenan los codigos VBA de los distintos modulos en un df (df_bbdd_codigo) con las columnas siguientes:
        # --> TIPO_MODULO
        # --> NOMBRE_MODULO
        # --> NUMERO_LINEA_CODIGO_MODULO
        # --> CODIGO
        df_bbdd_codigo = pd.DataFrame()
        cont = 0
        for component in vba_project.VBComponents:

            cont += 1

            lambda_tipo_modulo = lambda x: "Estandar" if x == 1 else "Clase" if x == 2 else "UserForm" if x == 3 else "Formulario/Reporte" if x == 100 else None
            tipo_modulo = lambda_tipo_modulo(component.type)

            nombre_modulo = component.Name
            code_module = component.CodeModule
            num_lines = code_module.CountOfLines
            code_lines = code_module.Lines(1, num_lines)

            lista_temp_1 = code_lines.split("\n")
            lista_temp_2 = [[tipo_modulo, nombre_modulo, ind + 1, linea] for ind, linea in enumerate(lista_temp_1)]

            df_temp = pd.DataFrame(lista_temp_2, columns = lista_headers_df_bbdd_codigo)
            df_bbdd_codigo = df_temp if cont == 1 else pd.concat([df_bbdd_codigo, df_temp])

            del df_temp

        df_bbdd_codigo.reset_index(drop = True, inplace = True)


        ##################################################################################################################################################
        # se generan los datos para lista_dicc_objetos de los objetos 
        # TABLA_LOCAL, VINCULO_ODBC y VINCULO_OTRO
        ##################################################################################################################################################        

        #se calcula el df df_access_vinculos para almacenar los datos de los vinculos si los hubiese
        #se realiza mediante recordset sobre la query query_access_vinculos
        MiRecordset = current_bbdd.OpenRecordset(query_access_vinculos)

        lista_temp = []
        while not MiRecordset.EOF:
            nombre_vinculo = MiRecordset.Fields("NOMBRE_VINCULO_ACCESS").Value
            link_connect = MiRecordset.Fields("LINK_ODBC").Value
            objeto_origen = MiRecordset.Fields("OBJETO_ORIGEN").Value

            lista_temp.append([nombre_vinculo, link_connect, objeto_origen])
            
            MiRecordset.MoveNext()

        MiRecordset.Close()

        df_access_vinculos = pd.DataFrame(lista_temp, columns = lista_headers_df_access_vinculos) if len(lista_temp) != 0 else None


        #se agregan los datos de las tablas / vinculos a la lista lista_objetos_con_o_sin_codigo
        #(se hace aqui pq requiere que la base de datos MS ACCESS este abierta)
        tipo_objeto = None
        nombre_objeto = None
        df_codigo = pd.DataFrame(columns = ["CODIGO"])
        for tabledef in current_bbdd.TableDefs:
            nombre_objeto = tabledef.Name

            if not tabledef.Name.startswith("MSys"):
                
                #TABLAS LOCALES
                if not tabledef.Connect:
                    tipo_objeto = "TABLA_LOCAL"
                    link_connect = None

                    prop_tabla = current_bbdd.TableDefs(nombre_objeto)

                    df_codigo = func_diagnostico_access_tablas_df_create_table(tipo_objeto, prop_tabla = prop_tabla, nombre_objeto = nombre_objeto) if proceso_id == "PROCESO_01" else None

                    dicc_temp = {"TIPO_OBJETO": tipo_objeto
                                , "TIPO_MODULO": None                                            #key que no aplica
                                , "NOMBRE_MODULO": None                                          #key que no aplica
                                , "NOMBRE_OBJETO": nombre_objeto
                                , "TIPO_RUTINA": None                                            #key que no aplica
                                , "TIPO_DECLARACION_RUTINA": None                                #key que no aplica
                                , "PARAMETROS_RUTINA": None                                      #key que no aplica
                                , "TIPO_VARIABLE_PUBLICA": None                                  #key que no aplica
                                , "TIPO_DATO_VARIABLE_PUBLICA": None                             #key que no aplica
                                , "VARIABLE_PUBLICA_ES_CONSTANTE": None                          #key que no aplica
                                , "VARIABLE_PUBLICA_CONSTANTE_VALOR": None                       #key que no aplica
                                , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": None    #key que no aplica
                                , "DF_CODIGO": df_codigo
                                , "CONNECTING_STRING_VINCULOS": None                             #key que no aplica
                                }

                    lista_dicc_objetos.append(dicc_temp)
                    del dicc_temp


                # VINCULOS
                elif tabledef.Connect:

                    tipo_objeto = "VINCULO_ODBC" if tabledef.Connect.upper().startswith("ODBC") else "VINCULO_OTRO"
                    link_connect = tabledef.Connect

                    df_codigo = func_diagnostico_access_tablas_df_create_table(tipo_objeto, link_connect = link_connect, nombre_vinculo = nombre_objeto, df_access_vinculos = df_access_vinculos)
                    connecting_string_vinculos = "\n".join([df_codigo.iloc[ind, 0] for ind in df_codigo.index])

                    dicc_temp = {"TIPO_OBJETO": tipo_objeto
                                , "TIPO_MODULO": None                                            #key que no aplica
                                , "NOMBRE_MODULO": None                                          #key que no aplica
                                , "NOMBRE_OBJETO": nombre_objeto
                                , "TIPO_RUTINA": None                                            #key que no aplica
                                , "TIPO_DECLARACION_RUTINA": None                                #key que no aplica
                                , "PARAMETROS_RUTINA": None                                      #key que no aplica
                                , "TIPO_VARIABLE_PUBLICA": None                                  #key que no aplica
                                , "TIPO_DATO_VARIABLE_PUBLICA": None                             #key que no aplica
                                , "VARIABLE_PUBLICA_ES_CONSTANTE": None                          #key que no aplica
                                , "VARIABLE_PUBLICA_CONSTANTE_VALOR": None                       #key que no aplica
                                , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": None    #key que no aplica
                                , "DF_CODIGO": df_codigo
                                , "CONNECTING_STRING_VINCULOS": connecting_string_vinculos
                                }

                    lista_dicc_objetos.append(dicc_temp)
                    del dicc_temp

        del df_access_vinculos


        ##################################################################################################################################################
        # se cierra el access
        ##################################################################################################################################################        

        access_app.CloseCurrentDatabase()
        access_app.Quit()
        access_app = None



        ###################################################################################################################################################################################################
        # se calcula en df_bbdd_codigo columnas adicionales necesarias que son COMUNES tanto para el control de versiones y el diagnostico de dependencias
        ###################################################################################################################################################################################################

        #ademas de las columnas resultantes de la importacion anterior, se calculan las columnas siguientes:
        #
        # --> INDICE                                              es el indice del df (se usa para localizar el nombre de la rutina / funcion y poder localizar su codigo asociado en CODIGO)
        # --> CODIGO_MAYUSC_SIN_TAB_TRIMEADO                      es la columna CODIGO puesta en mayusculas donde se quitan las tabulaciones y se trimea
        #
        # --> CODIGO_SIN_COMENTARIOS                              es la columna CODIGO a la que se le borran todos los comentarios para poder realizar el diagnostico
        #                                                         (los que son lineas completas de comentarios y las que estan a la derecha del codigo)
        #
        # --> CODIGO_SIN_COMENTARIOS_MAYUSC                       es la columna CODIGO_SIN_COMENTARIOS puesta en mayusculas 
        #                                                         (se usa para calcular las columnas CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI y CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO)
        #
        # --> CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS             es la columna CODIGO donde se quitan las tabulaciones y se trimea y al cual se le quitan los comentarios
        #                                                         (los que son lineas completas de comentarios y las que estan a la derecha del codigo)
        #
        # --> ES_ENCABEZADO_MODULO                                es la columna CODIGO_MAYUSC_SIN_TAB_TRIMEADO donde se filtra todo lo que empieza por OPTION que son las sentencias de modulo (SI / NO)
        #
        # --> TIPO_RUTINA                                         rutina o funcion
        # --> TIPO_DECLARACION_RUTINA                             publica o privada
        # --> NOMBRE_RUTINA                                       nombre rutina / funcion
        # --> NUMERO_LINEA_CODIGO_RUTINA                          numero de linea del codigo dentro de la rutina / funcion
        # --> PARAMETROS_RUTINA                                   parametros de la rutina / funcion con los tipos de datos asociados
        #
        # --> VARIABLES_LOCALES_RUTINA                            indica que lineas de codigo dentro de la rutina son declaraciones de variables locales
        #
        # --> TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA     recopila los tipos de dato de los parametros y variables locales de la rutina
        #                                                         se usa en el diagnostico de dependencias para localizar si variables publicas definidas por el usuario se usan en la rutina
        #
        # --> ES_VARIABLE_PUBLICA                                 4 valores distintos: NO, NO_ES_CONSTANTE, SI_ES_CONSTANTE y DEFINIDA_POR_USUARIO_iii
        #                                                         --> NO_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por "Public" o "Global")
        #                                                         --> SI_ES_CONSTANTE (para variables que se declaran en 1 sola linea, las que empiezan por ""Const", "Public Const" o "Global Const")
        #                                                         --> DEFINIDA_POR_USUARIO_iii (para variables definidas por el usuario, se declaran en varias lineas, se crea un literal unico 
        #                                                                                       para todas las lineas que componen la declaracion de una misma variable, iii es un numero de variable unico 
        #                                                                                       para cada una de estas variables)
        #
        # --> ES_VARIABLE_PUBLICA_CONTROL_VERSIONES               es la columna ES_VARIABLE_PUBLICA != NO donde no se eliminan los comentarios /se agregan las lineas de codigo que son todo comentarios
        #                                                         y que no estan asociados a ninguna rutina / funcion
        #
        # --> ES_VARIABLE_PUBLICA_DIAGNOSTICO                     es la columna ES_VARIABLE_PUBLICA != NO donde si se eliminan todos los comentarios


        #########################################################
        #     df_bbdd_codigo["INDICE"]
        #     df_bbdd_codigo["CODIGO_MAYUSC_SIN_TAB_TRIMEADO"]
        #     df_bbdd_codigo["CODIGO_SIN_COMENTARIOS"]
        #     df_bbdd_codigo["CODIGO_SIN_COMENTARIOS_MAYUSC"]
        #     df_bbdd_codigo["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"]
        #     df_bbdd_codigo["ES_ENCABEZADO_MODULO"]
        #########################################################

        #para quitar los comentarios (los que corresponden a toda la linea comentada o los que estan a la derecha de una una de codigo activa)
        #se usa la funcion func_codigo_vba_quitar_comentarios

        df_bbdd_codigo["CODIGO"] = df_bbdd_codigo["CODIGO"].apply(lambda x: x.replace("\r", "").replace("\\'", indicador_comentario_vba_access))
        df_bbdd_codigo["INDICE"] = df_bbdd_codigo.index
        df_bbdd_codigo["CODIGO_MAYUSC_SIN_TAB_TRIMEADO"] = df_bbdd_codigo["CODIGO"].apply(lambda x: x.upper().replace("\t", "").strip())

        df_bbdd_codigo["CODIGO_SIN_COMENTARIOS"] = df_bbdd_codigo["CODIGO"].apply(lambda x: func_codigo_vba_quitar_comentarios(str(x)))
        df_bbdd_codigo["CODIGO_SIN_COMENTARIOS_MAYUSC"] = df_bbdd_codigo["CODIGO_SIN_COMENTARIOS"].apply(lambda x: x.upper())

        df_bbdd_codigo["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"] = df_bbdd_codigo["CODIGO"].apply(lambda x: func_codigo_vba_quitar_comentarios(str(x).replace("\t", "").strip()))

        df_bbdd_codigo["ES_ENCABEZADO_MODULO"] = df_bbdd_codigo["CODIGO_MAYUSC_SIN_TAB_TRIMEADO"].apply(lambda x: "SI" if len(x) >= 6 and x.upper()[0:6] == "OPTION" else "NO")


        #########################################################
        #     df_bbdd_codigo["TIPO_RUTINA"]
        #     df_bbdd_codigo["TIPO_DECLARACION_RUTINA"]
        #     df_bbdd_codigo["NOMBRE_RUTINA"]
        #     df_bbdd_codigo["NUMERO_LINEA_CODIGO_RUTINA"]
        #     df_bbdd_codigo["PARAMETROS_RUTINA"]
        #     df_bbdd_codigo["VARIABLES_LOCALES_RUTINA"]
        #     df_bbdd_codigo["TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA"]

        #     df_bbdd_codigo["ES_VARIABLE_PUBLICA"]
        #     df_bbdd_codigo["ES_VARIABLE_PUBLICA_CONTROL_VERSIONES"]
        #     df_bbdd_codigo["ES_VARIABLE_PUBLICA_DIAGNOSTICO"]
        #########################################################

        #se crean las columnas adicionales en el df df_bbdd_codigo para las RUTINAS usando la lista 
        #que devuelve la funcion func_access_bbdd_codigo_calculos que se convierte en df y se realiza merge 
        #con el df df_bbdd_codigo por la columna INDICE
        lista_calculos_rutinas = func_access_bbdd_codigo_calculos("LISTA_RUTINAS", df_bbdd_codigo_param = df_bbdd_codigo)

        df_temp_rutinas = pd.DataFrame(lista_calculos_rutinas, columns = lista_headers_nuevas_columnas_rutinas)

        df_bbdd_codigo = (pd.merge(df_bbdd_codigo[[i for i in df_bbdd_codigo.columns]], df_temp_rutinas[[i for i in df_temp_rutinas.columns]], 
                            how = "left", left_on = "INDICE", right_on = "INDICE"))



        #se crean las columnas adicionales en el df df_bbdd_codigo para las VARIABLES PUBLICAS usando la lista 
        #que devuelve la funcion func_access_bbdd_codigo_calculos que se convierte en df y se realiza merge 
        #con el df df_bbdd_codigo por la columna INDICE
        lista_calculos_variables_publicas = func_access_bbdd_codigo_calculos("LISTA_VARIABLES_PUBLICAS", df_bbdd_codigo_param = df_bbdd_codigo)

        df_temp_variables_publicas = pd.DataFrame(lista_calculos_variables_publicas, columns = lista_headers_nuevas_columnas_variables_publicas)

        df_bbdd_codigo = (pd.merge(df_bbdd_codigo[[i for i in df_bbdd_codigo.columns]], 
                                   df_temp_variables_publicas[[i for i in df_temp_variables_publicas.columns if i not in ["TIPO_MODULO", "NOMBRE_MODULO", "CODIGO"]]], 
                            how = "left", left_on = "INDICE", right_on = "INDICE"))



        # ES_VARIABLE_PUBLICA_CONTROL_VERSIONES
        # ES_VARIABLE_PUBLICA_DIAGNOSTIC
        df_bbdd_codigo["ES_VARIABLE_PUBLICA_CONTROL_VERSIONES"] = (df_bbdd_codigo.apply(lambda x: "SI" 
                                                                                    if (x["CODIGO_MAYUSC_SIN_TAB_TRIMEADO"][0:1] == indicador_comentario_vba_access and x["NOMBRE_RUTINA"] == None) 
                                                                                    or x["ES_VARIABLE_PUBLICA"] != "NO" else "NO", axis = 1))
    
        df_bbdd_codigo["ES_VARIABLE_PUBLICA_DIAGNOSTICO"] = df_bbdd_codigo.apply(lambda x: "SI" if x["ES_VARIABLE_PUBLICA"] != "NO" else "NO", axis = 1)



        ##################################################################################################################################################
        #se generan los datos para lista_dicc_objetos de los objetos RUTINAS_VBA
        #se suma a lista_dicc_objetos ya calculados mas arriba para tablas y vinculos
        ##################################################################################################################################################        

        lista_campos_rutina = ["TIPO_MODULO", "NOMBRE_MODULO", "NOMBRE_RUTINA", "TIPO_RUTINA", "TIPO_DECLARACION_RUTINA", "PARAMETROS_RUTINA"]

        df_rutinas = df_bbdd_codigo.loc[df_bbdd_codigo["NOMBRE_RUTINA"].isnull() == False, lista_campos_rutina]
        df_rutinas.drop_duplicates(subset = lista_campos_rutina, keep = "last", inplace = True)
        df_rutinas.reset_index(drop = True, inplace = True)


        for ind in df_rutinas.index:
        
            tipo_modulo = df_rutinas.iloc[ind, df_rutinas.columns.get_loc("TIPO_MODULO")]
            nombre_modulo = df_rutinas.iloc[ind, df_rutinas.columns.get_loc("NOMBRE_MODULO")]
            nombre_rutina = df_rutinas.iloc[ind, df_rutinas.columns.get_loc("NOMBRE_RUTINA")]


            #si el proceso seleccionado es el control de versiones (PROCESO_01) se agrega el codigo sin alteraciones (columna CODIGO)
            #en caso de seleccionar el diagnostico de dependencias (PROCESO_02) no se agrega el codigo
            if proceso_id == "PROCESO_01":
                df_codigo = (df_bbdd_codigo.loc[(df_bbdd_codigo["TIPO_MODULO"] == tipo_modulo) & (df_bbdd_codigo["NOMBRE_MODULO"] == nombre_modulo) & 
                                                (df_bbdd_codigo["NOMBRE_RUTINA"] == nombre_rutina), ["CODIGO"]])

                df_codigo.reset_index(drop = True, inplace = True)

            elif proceso_id == "PROCESO_02":
                df_codigo = pd.DataFrame(columns = ["CODIGO"])


            #se agregan los datos a la lista lista_dicc_objetos
            dicc_temp = {"TIPO_OBJETO": "RUTINAS_VBA"
                        , "TIPO_MODULO": tipo_modulo
                        , "NOMBRE_MODULO": nombre_modulo
                        , "NOMBRE_OBJETO": nombre_rutina
                        , "TIPO_RUTINA": df_rutinas.iloc[ind, df_rutinas.columns.get_loc("TIPO_RUTINA")]
                        , "TIPO_DECLARACION_RUTINA": df_rutinas.iloc[ind, df_rutinas.columns.get_loc("TIPO_DECLARACION_RUTINA")]
                        , "PARAMETROS_RUTINA": df_rutinas.iloc[ind, df_rutinas.columns.get_loc("PARAMETROS_RUTINA")]
                        , "TIPO_VARIABLE_PUBLICA": None                                  #key que no aplica
                        , "TIPO_DATO_VARIABLE_PUBLICA": None                             #key que no aplica
                        , "VARIABLE_PUBLICA_ES_CONSTANTE": None                          #key que no aplica
                        , "VARIABLE_PUBLICA_CONSTANTE_VALOR": None                       #key que no aplica
                        , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": None    #key que no aplica
                        , "DF_CODIGO": df_codigo
                        , "CONNECTING_STRING_VINCULOS": None                             #key que no aplica
                        }

            lista_dicc_objetos.append(dicc_temp)

            del dicc_temp
            del df_codigo

        del df_rutinas


        ##################################################################################################################################################
        # se generan los datos para lista_dicc_objetos de los objetos VARIABLES_VBA
        ##################################################################################################################################################        

        #se generan los datos para lista_dicc_objetos de los objetos VARIABLES_VBA con la funcion func_access_df_bbdd_codigo_calculos_varios 
        #pq requieren muchos calculos intermedios para poder presentarlas en forma de listado (es para aligerar el codigo de la presente rutina)
        #se usa la lista temporal de variables publicas generada más arriba con la misma funcion (opcion LISTA_VARIABLES_PUBLICAS)
        lista_dicc_objetos_variables_publicas = func_access_bbdd_codigo_calculos("LISTA_DICC_OBJETOS_VARIABLES_PUBLICAS", lista_calculos_variables_publicas = lista_calculos_variables_publicas)


        ##################################################################################################################################################
        # se actualiza el diccionario dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["LISTA_DICC_OBJETOS"]
        # con lista_dicc_objetos
        ##################################################################################################################################################        

        #se fusiona lista_dicc_objetos con lista_dicc_objetos_variables_publicas
        lista_dicc_objetos = lista_dicc_objetos + lista_dicc_objetos_variables_publicas


        #se almacena lista_dicc_objetos en dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["LISTA_OBJETOS_CON_O_SIN_CODIGO"]
        mod_gen.dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["LISTA_DICC_OBJETOS"] = lista_dicc_objetos

        del lista_dicc_objetos
        del lista_dicc_objetos_variables_publicas



        ##################################################################################################################################
        #se agregan a df_bbdd_codigo las columnas necesarias al diagnostico de dependencias

        # --> CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI             es la columna CODIGO_SIN_COMENTARIOS_MAYUSC donde se excluye todo segun que no se adapte 
        #                                                                   al diccionario dicc_diagnostico_tablas_access_donde_buscar y sirve para el calculo de dependencias 
        #                                                                   de tablas / vinculos con la funcion func_diagnostico_access_tablas_codigo_preservar

        # --> CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO             es la columna CODIGO_SIN_COMENTARIOS_MAYUSC con todo lo incluido entre comillas dobles
        #                                                                   y que no este incluido en CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI
        #                                                                   se usa en el diagnostico para informar de las tablas que se usan (o no) fuera del ambito 
        #                                                                   de las logicas configuradas en el diccionario dicc_diagnostico_tablas_access_donde_buscar
        #                                                                   es el usuario quien tiene que chequear manualmente desde el excel final si se ha de considerar dependencia o no
        #                                                                   (es una lista con el numero de linea de la rutina / funcion y el codigo de la linea)

        # --> CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES           es la columna CODIGO_SIN_COMENTARIOS_MAYUSC donde se quita todo lo incuido entre "" (usar un objeto llamado aaa no es lo mismo que usar un "aaa")
        #                                                      sirve para el calculo de dependencias de rutinas VBA y variables VBA
        ##################################################################################################################################

        if proceso_id == "PROCESO_02":

            #se calcula CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_SI y CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO usando
            #la funcion func_diagnostico_access_tablas_codigo_preservar
            df_bbdd_codigo["CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI"] = (df_bbdd_codigo.apply(lambda x: 
                                                                                    func_diagnostico_access_tablas_codigo_preservar("PRESERVAR_SI", x["CODIGO_SIN_COMENTARIOS_MAYUSC"])
                                                                                    , axis = 1))
            

            df_bbdd_codigo["CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO"] = (df_bbdd_codigo.apply(lambda x: 
                                                                                    func_diagnostico_access_tablas_codigo_preservar("PRESERVAR_NO", x["CODIGO_SIN_COMENTARIOS_MAYUSC"])
                                                                                    , axis = 1))


            #se calcula CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES
            #usando CODIGO_SIN_COMENTARIOS_MAYUSC se quita todo lo incluido entre "" para evitar que localice en el diagnostico objetos tipo rutinas o variables 
            #incluidos en string (ejemplo: usar un objeto llamado aaa no es lo mismo que usar un "aaa")
            df_bbdd_codigo["CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES"] = (df_bbdd_codigo["CODIGO_SIN_COMENTARIOS_MAYUSC"].apply(lambda x: 
                                                                                        func_access_bbdd_codigo_calculos_varios("QUITAR_ENCAPSULADO_ENTRE_COMILLAS_DOBLES", string_param = x.upper())))


        ################################################################################################################
        # se almacena dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]
        ################################################################################################################

        mod_gen.dicc_codigos_bbdd[opcion_bbdd]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"] = df_bbdd_codigo
        del df_bbdd_codigo



    except Exception as Err:

        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos[proceso_id]["MS_ACCESS"]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd]

        dicc_errores_temp = {"TIPO_BBDD": "MS_ACCESS"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos[proceso_id]["MS_ACCESS"]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd] = lista_dicc_errores_migracion
        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso de import



#################################################################################################################################################################################
##                     RUTINA MS ACCESS - CONTROL DE VERSIONES
#################################################################################################################################################################################

def def_proceso_access_2_control_versiones():
    #rutina que permite realizar los calculos necesarios para el control de versiones

    warnings.filterwarnings("ignore")

    try:

        #se recuperan los datos necesarios del diccionario dicc_codigos_bbdd
        lista_dicc_objetos_bbdd_1 = mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["LISTA_DICC_OBJETOS"]
        lista_dicc_objetos_bbdd_2 = mod_gen.dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["LISTA_DICC_OBJETOS"]

        df_bbdd_codigo_1 = mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]
        df_bbdd_codigo_2 = mod_gen.dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]


        ###############################################################
        #       AJUSTES PREVIOS
        ###############################################################

        #se han de ajustar las listas lista_dicc_objetos_bbdd_1 y lista_dicc_objetos_bbdd_2 recuperadas del diccionario dicc_codigos_bbdd
        #para el control de versiones de las variables publicas VBA no se tira de lista_dicc_objetos_bbdd_1 y lista_dicc_objetos_bbdd_2 pq aqui 
        #las variables publicas estan listadas 1 a 1, en el control de versiones para este tipo de objetos el calculo se hace a nivel de modulo no de objeto
        #
        # 1 --> mediante el codigo almacenado en el diccionario dicc_codigos_bbdd se extraen los tipos y nombre de modulos donde se han localizado variables publicas 
        #       para el control de versiones (columna ES_VARIABLE_PUBLICA_CONTROL_VERSIONES = SI) y se almacenan estos tipos y nombres de modulos en las
        #       listas temporales lista_temp_1 y lista_temp_2 a las cuales se les quita los duplicados

        # 2 --> mediante bucle lista_temp_1 y lista_temp_2 se reconstruyen las listas lista_dicc_objetos_bbdd_1 y lista_dicc_objetos_bbdd_2
        #       solo para VARIABLES_VBA, en cada iteracion se extrae de los df de codigo (donde columna ES_VARIABLE_PUBLICA_CONTROL_VERSIONES = SI 
        #       y donde el tipo y nombre del modulo del df de codigo corresponde con los de la iteracion), en cada diccionario que se agrega a
        #       estas listas reconstruidas tan solo se informa las keys TIPO_OBJETO, TIPO_MODULO, NOMBRE_MODULO y DF_CODIGO
        #
        # 3 --> las listas definitivas lista_dicc_objetos_bbdd_1 y lista_dicc_objetos_bbdd_2 son la fusion entre las listas originales (sin VARIABLES_VBA)
        #       y de estas listas reconstruidas

        #############################
        # VARIABLES_VBA
        #############################

        #se extraen los codigos por tipo y nombre de modulo de los df df_bbdd_codigo_1 y df_bbdd_codigo_2
        df_temp_1 = df_bbdd_codigo_1.loc[df_bbdd_codigo_1["ES_VARIABLE_PUBLICA_CONTROL_VERSIONES"] == "SI", ["TIPO_MODULO", "NOMBRE_MODULO", "CODIGO"]]
        df_temp_2 = df_bbdd_codigo_2.loc[df_bbdd_codigo_2["ES_VARIABLE_PUBLICA_CONTROL_VERSIONES"] == "SI", ["TIPO_MODULO", "NOMBRE_MODULO", "CODIGO"]]
    
        df_temp_1.reset_index(drop = True, inplace = True)
        df_temp_2.reset_index(drop = True, inplace = True)


        #se listan los tipos y nombres de modulos en cada base de datos y se quitan los duplicados
        lista_temp_1 = []
        lista_temp_2 = []

        if len(df_temp_1) != 0:
            lista_temp_1 = [[df_temp_1.iloc[ind, df_temp_1.columns.get_loc("TIPO_MODULO")]
                            , df_temp_1.iloc[ind, df_temp_1.columns.get_loc("NOMBRE_MODULO")]] for ind in df_temp_1.index]

            lista_temp_1 = [sublista for i, sublista in enumerate(lista_temp_1) if sublista not in lista_temp_1[:i]]


        if len(df_temp_2) != 0:
            lista_temp_2 = [[df_temp_2.iloc[ind, df_temp_2.columns.get_loc("TIPO_MODULO")]
                            , df_temp_2.iloc[ind, df_temp_2.columns.get_loc("NOMBRE_MODULO")]] for ind in df_temp_2.index]
        
            lista_temp_2 = [sublista for i, sublista in enumerate(lista_temp_2) if sublista not in lista_temp_2[:i]]


        #se reconstruyen listas para VARIABLES_VBA consolidadas por tipo y nombre de modulo con el codigo del modulo (excluyendo rutinas / funciones)
        lista_dicc_objetos_variables_publicas_bbdd_1 = []
        if len(lista_temp_1) != 0:

            for tipo_modulo, nombre_modulo in lista_temp_1:

                df_codigo = df_temp_1.loc[(df_temp_1["TIPO_MODULO"] == tipo_modulo) & (df_temp_1["NOMBRE_MODULO"] == nombre_modulo), ["CODIGO"]]
                df_codigo.reset_index(drop = True, inplace = True)

                df_codigo = df_codigo if len(df_codigo) != 0 else None


                dicc_temp = {"TIPO_OBJETO": "VARIABLES_VBA"
                            , "TIPO_MODULO": tipo_modulo
                            , "NOMBRE_MODULO": nombre_modulo
                            , "NOMBRE_OBJETO": None                                           #key que no aplica aqui
                            , "TIPO_RUTINA": None                                             #key que no aplica aqui
                            , "TIPO_DECLARACION_RUTINA": None                                 #key que no aplica aqui
                            , "PARAMETROS_RUTINA": None                                       #key que no aplica aqui
                            , "TIPO_VARIABLE_PUBLICA": None                                   #key que no aplica aqui
                            , "TIPO_DATO_VARIABLE_PUBLICA": None                              #key que no aplica aqui
                            , "VARIABLE_PUBLICA_ES_CONSTANTE": None                           #key que no aplica aqui
                            , "VARIABLE_PUBLICA_CONSTANTE_VALOR": None                        #key que no aplica aqui
                            , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": None     #key que no aplica aqui
                            , "DF_CODIGO": df_codigo
                            }

                lista_dicc_objetos_variables_publicas_bbdd_1.append(dicc_temp)
                del dicc_temp
                del df_codigo



        lista_dicc_objetos_variables_publicas_bbdd_2 = []
        if len(lista_temp_2) != 0:

            for tipo_modulo, nombre_modulo in lista_temp_2:

                df_codigo = df_temp_2.loc[(df_temp_2["TIPO_MODULO"] == tipo_modulo) & (df_temp_2["NOMBRE_MODULO"] == nombre_modulo), ["CODIGO"]]
                df_codigo.reset_index(drop = True, inplace = True)

                df_codigo = df_codigo if len(df_codigo) != 0 else None


                dicc_temp = {"TIPO_OBJETO": "VARIABLES_VBA"
                            , "TIPO_MODULO": tipo_modulo
                            , "NOMBRE_MODULO": nombre_modulo
                            , "NOMBRE_OBJETO": None                                           #key que no aplica aqui
                            , "TIPO_RUTINA": None                                             #key que no aplica aqui
                            , "TIPO_DECLARACION_RUTINA": None                                 #key que no aplica aqui
                            , "PARAMETROS_RUTINA": None                                       #key que no aplica aqui
                            , "TIPO_VARIABLE_PUBLICA": None                                   #key que no aplica aqui
                            , "TIPO_DATO_VARIABLE_PUBLICA": None                              #key que no aplica aqui
                            , "VARIABLE_PUBLICA_ES_CONSTANTE": None                           #key que no aplica aqui
                            , "VARIABLE_PUBLICA_CONSTANTE_VALOR": None                        #key que no aplica aqui
                            , "VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES": None     #key que no aplica aqui
                            , "DF_CODIGO": df_codigo
                            }

                lista_dicc_objetos_variables_publicas_bbdd_2.append(dicc_temp)
                del dicc_temp
                del df_codigo

        del df_temp_1
        del df_temp_2
        del lista_temp_1
        del lista_temp_2


        #############################
        # listas ajustadas
        #############################

        lista_dicc_objetos_bbdd_1 = [dicc for dicc in lista_dicc_objetos_bbdd_1 if dicc["TIPO_OBJETO"] != "VARIABLES_VBA"] + lista_dicc_objetos_variables_publicas_bbdd_1
        lista_dicc_objetos_bbdd_2 = [dicc for dicc in lista_dicc_objetos_bbdd_2 if dicc["TIPO_OBJETO"] != "VARIABLES_VBA"] + lista_dicc_objetos_variables_publicas_bbdd_2

        del lista_dicc_objetos_variables_publicas_bbdd_1
        del lista_dicc_objetos_variables_publicas_bbdd_2


        ###############################################################
        #       CALCULO CONTROL VERSIONES
        ###############################################################

        #se crea la lista lista_objetos_conso qe permite consolidar los datos por objeto de una base de datos a otra
        #se compone de  sblistas donde:
        # 0 --> tipo objeto
        # 1 --> tipo modulo
        # 2 --> nombre modulo
        # 3 --> nombre objeto
        # 4 --> indicador de si el objeto esta en BBDD_01 (de inicio se marca a 0, se calcula despues)
        # 5 --> indicador de si el objeto esta en BBDD_02 (de inicio se marca a 0, se calcula despues)
        # 6 --> df con el codigo del objeto en BBDD_01 (de inicio se marca a None, se calcula despues)
        # 7 --> df con el codigo del objeto en BBDD_02 (de inicio se marca a None, se calcula despues)

        indice_tipo_objeto = 0
        indice_tipo_modulo = 1
        indice_nombre_modulo = 2
        indice_nombre_objeto = 3
        indice_esta_en_bbdd_1 = 4
        indice_esta_en_bbdd_2 = 5
        indice_df_codigo_1 = 6
        indice_df_codigo_2 = 7

        lista_objetos_conso_1 = [[dicc["TIPO_OBJETO"]
                                , dicc["TIPO_MODULO"]
                                , dicc["NOMBRE_MODULO"]
                                , dicc["NOMBRE_OBJETO"]
                                , 0
                                , 0
                                , None
                                , None] for dicc in lista_dicc_objetos_bbdd_1]

        lista_objetos_conso_2 = [[dicc["TIPO_OBJETO"]
                                , dicc["TIPO_MODULO"]
                                , dicc["NOMBRE_MODULO"]
                                , dicc["NOMBRE_OBJETO"]
                                , 0
                                , 0
                                , None
                                , None] for dicc in lista_dicc_objetos_bbdd_2]

        lista_objetos_conso = lista_objetos_conso_1 + lista_objetos_conso_2
        lista_objetos_conso = [sublista for i, sublista in enumerate(lista_objetos_conso) if sublista not in lista_objetos_conso[:i]]#se quitan los duplicados


        #se informan los indices (4 a 7) de las sublistas de lista_objetos_conso
        #la busqueda de los items de cada sublista de lista_objetos_conso se hace en lista_dicc_objetos_bbdd_1 (para ver si el objeto existe en BBDD_01)
        #y en lista_dicc_objetos_bbdd_2 (para var si el objeto existe en BBDD_02)
        #los criterios de busqueda y macheo varian por tipo de objetos:
        # --> tablas / vinculos              el macheo se hace por tipo de objeto y nombre de objeto (las tablas / vinculos no son objetos VBA por lo que no se almacean en modulos)
        # --> rutinas VBA                    el macheo se hace por tipo de objeto, tipo y nombre de modulo VBA y nombre de objeto
        # --> variables publicas VBA         el macheo se hace por tipo de objeto, tipo y nombre de modulo VBA (como se comenta mas arriba el control de versiones de las variables publicas 
        #                                    es a nivel de modulo no de objeto)
        for indice_lista, item in enumerate(lista_objetos_conso):

            tipo_objeto_conso = item[indice_tipo_objeto]
            tipo_modulo_conso = item[indice_tipo_modulo]
            nombre_modulo_conso = item[indice_nombre_modulo]
            nombre_objeto_conso = item[indice_nombre_objeto]


            #BBDD_01 (se informan los indices 4 y 6)
            cont_1 = 0
            for dicc in lista_dicc_objetos_bbdd_1:

                tipo_objeto_seek = dicc["TIPO_OBJETO"]
                tipo_modulo_seek = dicc["TIPO_MODULO"]
                nombre_modulo_seek = dicc["NOMBRE_MODULO"]
                nombre_objeto_seek = dicc["NOMBRE_OBJETO"]
                df_codigo_seek = dicc["DF_CODIGO"]

                if tipo_objeto_conso in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:
                    if tipo_objeto_conso == tipo_objeto_seek and nombre_objeto_conso == nombre_objeto_seek:
                        cont_1 = 1
                        break

                elif tipo_objeto_conso == "RUTINAS_VBA":
                    if tipo_objeto_conso == tipo_objeto_seek and tipo_modulo_conso == tipo_modulo_seek and nombre_modulo_conso == nombre_modulo_seek and nombre_objeto_conso == nombre_objeto_seek:
                        cont_1 = 1
                        break

                elif tipo_objeto_conso == "VARIABLES_VBA":
                    if tipo_objeto_conso == tipo_objeto_seek and tipo_modulo_conso == tipo_modulo_seek and nombre_modulo_conso == nombre_modulo_seek:
                        cont_1 = 1
                        break

            if cont_1 == 1:
                lista_objetos_conso[indice_lista][indice_esta_en_bbdd_1] = 1
                lista_objetos_conso[indice_lista][indice_df_codigo_1] = df_codigo_seek



            #BBDD_02 (se informan los indices 5 y 7)
            cont_2 = 0
            for dicc in lista_dicc_objetos_bbdd_2:

                tipo_objeto_seek = dicc["TIPO_OBJETO"]
                tipo_modulo_seek = dicc["TIPO_MODULO"]
                nombre_modulo_seek = dicc["NOMBRE_MODULO"]
                nombre_objeto_seek = dicc["NOMBRE_OBJETO"]
                df_codigo_seek = dicc["DF_CODIGO"]

                if tipo_objeto_conso in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:
                    if tipo_objeto_conso == tipo_objeto_seek and nombre_objeto_conso == nombre_objeto_seek:
                        cont_2 = 1
                        break

                elif tipo_objeto_conso == "RUTINAS_VBA":
                    if tipo_objeto_conso == tipo_objeto_seek and tipo_modulo_conso == tipo_modulo_seek and nombre_modulo_conso == nombre_modulo_seek and nombre_objeto_conso == nombre_objeto_seek:
                        cont_2 = 1
                        break

                elif tipo_objeto_conso == "VARIABLES_VBA":
                    if tipo_objeto_conso == tipo_objeto_seek and tipo_modulo_conso == tipo_modulo_seek and nombre_modulo_conso == nombre_modulo_seek:
                        cont_2 = 1
                        break

            if cont_2 == 1:
                lista_objetos_conso[indice_lista][indice_esta_en_bbdd_2] = 1
                lista_objetos_conso[indice_lista][indice_df_codigo_2] = df_codigo_seek



        #se calcula el control de versiones
        lista_control_versiones_tablas_access = []
        lista_control_versiones_vinculos_odbc = []
        lista_control_versiones_vinculos_otros = []
        lista_control_versiones_rutinas_vba = []
        lista_control_versiones_variables_vba = []

        for tipo_objeto_conso, tipo_modulo_conso, nombre_modulo_conso, nombre_objeto_conso, en_bbdd_1_conso, en_bbdd_2_conso, df_codigo_1_conso, df_codigo_2_conso in lista_objetos_conso:

            #segun se localice si el objeto esta en las 2 bbdd o solo en 1 se le asigna el valor correspondiente de la key
            #del diccionario dicc_control_versiones_tipo_concepto
            if en_bbdd_1_conso == 1 and en_bbdd_2_conso == 1:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["YA_EXISTE"]

            elif en_bbdd_1_conso == 1 and en_bbdd_2_conso == 0:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_01"]

            elif en_bbdd_1_conso == 0 and en_bbdd_2_conso == 1:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_02"]


            #se calcula el control de versiones usando la funcion func_dicc_control_versiones que tiene parametros kwargs
            #segun el tipo de objeto, el resultado de la funcion es un diccionario que se almacena en listas temporales
            if tipo_objeto_conso in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:

                dicc_temp = mod_gen.func_dicc_control_versiones(tipo_bbdd = "MS_ACCESS"
                                                            , check_objeto = check_objeto
                                                            , tipo_objeto = tipo_objeto_conso
                                                            , nombre_objeto = nombre_objeto_conso
                                                            , df_codigo_bbdd_1 = df_codigo_1_conso
                                                            , df_codigo_bbdd_2 = df_codigo_2_conso
                                                            )
    
                if isinstance(dicc_temp, dict):
                    if tipo_objeto_conso == "TABLA_LOCAL":
                        lista_control_versiones_tablas_access.append(dicc_temp)

                    elif tipo_objeto_conso == "VINCULO_ODBC":
                        lista_control_versiones_vinculos_odbc.append(dicc_temp)

                    elif tipo_objeto_conso == "VINCULO_OTRO":
                        lista_control_versiones_vinculos_otros.append(dicc_temp)



            elif tipo_objeto_conso == "RUTINAS_VBA":

                dicc_temp = mod_gen.func_dicc_control_versiones(tipo_bbdd = "MS_ACCESS"
                                                            , check_objeto = check_objeto
                                                            , tipo_objeto = tipo_objeto_conso
                                                            , tipo_repositorio = tipo_modulo_conso
                                                            , repositorio = nombre_modulo_conso
                                                            , nombre_objeto = nombre_objeto_conso
                                                            , df_codigo_bbdd_1 = df_codigo_1_conso
                                                            , df_codigo_bbdd_2 = df_codigo_2_conso
                                                            )
                
                if isinstance(dicc_temp, dict):
                    lista_control_versiones_rutinas_vba.append(dicc_temp)



            elif tipo_objeto_conso == "VARIABLES_VBA":

                dicc_temp = mod_gen.func_dicc_control_versiones(tipo_bbdd = "MS_ACCESS"
                                                            , check_objeto = check_objeto
                                                            , tipo_objeto = tipo_objeto_conso
                                                            , tipo_repositorio = tipo_modulo_conso
                                                            , repositorio = nombre_modulo_conso
                                                            , df_codigo_bbdd_1 = df_codigo_1_conso
                                                            , df_codigo_bbdd_2 = df_codigo_2_conso
                                                            )
                
                if isinstance(dicc_temp, dict):
                    lista_control_versiones_variables_vba.append(dicc_temp)


        del lista_objetos_conso
        del lista_dicc_objetos_bbdd_1
        del lista_dicc_objetos_bbdd_2


        #se almacenan las listas de diccionarios temporales en el diccionario dicc_control_versiones_tipo_objeto dicc_control_versiones_tipo_objeto
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["TABLA_LOCAL"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_tablas_access if len(lista_control_versiones_tablas_access) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VINCULO_ODBC"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_vinculos_odbc if len(lista_control_versiones_vinculos_odbc) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VINCULO_OTRO"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_vinculos_otros if len(lista_control_versiones_vinculos_otros) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["RUTINAS_VBA"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_rutinas_vba if len(lista_control_versiones_rutinas_vba) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["VARIABLES_VBA"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_variables_vba if len(lista_control_versiones_variables_vba) != 0 else None

        del lista_control_versiones_tablas_access
        del lista_control_versiones_vinculos_odbc
        del lista_control_versiones_vinculos_otros
        del lista_control_versiones_rutinas_vba
        del lista_control_versiones_variables_vba



        #se almacena la lista en dicc_control_versiones_tipo_objeto en TODOS (es para el combobox opcion TODOS)
        lista_tipo_objetos = list(mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"].keys())

        lista_control_versiones_todo = []
        for tipo_objeto in lista_tipo_objetos:
            if tipo_objeto != "TODOS":

                try:
                    for dicc in mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]:
                        lista_control_versiones_todo.append(dicc)

                except:#puede haber tipo de objetos que no tengan cambios
                    pass
        
        mod_gen.dicc_control_versiones_tipo_objeto["MS_ACCESS"]["TIPO_OBJETO"]["TODOS"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_todo if len(lista_control_versiones_todo) != 0 else None

        del lista_control_versiones_todo


        #se vacia parte de dicc_codigos_bbdd para liberar memoria
        mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["LISTA_OBJETOS_CON_O_SIN_CODIGO"] = None
        mod_gen.dicc_codigos_bbdd["BBDD_02"]["MS_ACCESS"]["LISTA_OBJETOS_CON_O_SIN_CODIGO"] = None

        mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"] = None #se conserva DF_CODIGO_CALCULADO_TRAS_IMPORT 
                                                                                                    #(para BBDD_02 pq es por defecto la bbdd donde se hace el merge en bbdd fisica
                                                                                                    #(se usa a la hora de hacer el merge en bbdd fisica y saber donde incrustar 
                                                                                                    #las rutinas / funciones VBA dentro del modulo para mentener la misma estructura 
                                                                                                    # del modulo de inicio, es decir a nivel de orden de las rutinas dentro del modulo))



    except Exception as Err:

        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos["PROCESO_01"]["MS_ACCESS"]["LISTA_DICC_ERRORES_CALCULO"]

        dicc_errores_temp = {"TIPO_BBDD": "MS_ACCESS"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos["PROCESO_01"]["MS_ACCESS"]["LISTA_DICC_ERRORES_CALCULO"] = lista_dicc_errores_migracion
        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso



#################################################################################################################################################################################
##                     RUTINA MS ACCESS - DIAGNOSTICO
#################################################################################################################################################################################

def def_proceso_access_2_diagnostico(ruta_destino_excel_diagnostico_access):
    #rutina que permite realizar los calculos necesarios para el diagnostico y su exportacion a excel


    warnings.filterwarnings("ignore")

    try:

        #se recuperan los datos necesarios del diccionario dicc_codigos_bbdd 
        #(por defecto BBDD_01 el la bbdd donde se hace el diagnostico)
        lista_dicc_objetos = mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["LISTA_DICC_OBJETOS"]
        df_bbdd_codigo = mod_gen.dicc_codigos_bbdd["BBDD_01"]["MS_ACCESS"]["DF_CODIGO_CALCULADO_TRAS_IMPORT"]


        #########################################################################################################
        #        AJUSTE - se crea lista donde buscar los objetos en el codigo VBA de rutinas y funciones
        #########################################################################################################

        #se extraen del df df_bbdd_codigo las columnas necesecarias donde realizar las busquedas en el codigo VBA de las rutinas / funciones
        # --> tablas y vinculos          la busqueda se realiza en las columnas CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI y CODIGO_DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO
        # --> rutinas  VBA               la busqueda se realiza en la columna CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES
        #
        # --> variables VBA              variables normales                    la busqueda se realiza en la columna CODIGO_DIAGNOSTICO_RUTINAS_Y_VARIABLES 
        #                                variables definidas por el usuario    la busqueda se realiza en la columna TIPO_DATO_PARAMETROS_Y_VARIABLES_LOCALES_RUTINA 
        #
        #se crea la lista de diccionarios lista_lista_dicc_rutinas_codigo_donde_buscar donde cada diccionario contiene:
        # --> tipo modulo
        # --> nombre modulo
        # --> nombre rutina
        # --> diccionario con las listas con las lineas de codigo de la rutina donde item:
        #                   --> codigo para diagnostico de tablas y vinculos
        #                   --> codigo para diagnostico de rutinas y variables
        #                   --> tipo datos de parametros y variables locales de las rutinas (es para el diagnostico de variables publicas definidas por el usuario)
        #
        #se crea esta lista de diccionarios pq en mas rapido en ejecucion iterar por listas en el diagnostico que extraer, en cada iteracion (ver mas adelante) de los
        #objetos que se buscan, un df que contenga el nombre del objeto

        df_temp = (df_bbdd_codigo.loc[(df_bbdd_codigo["NOMBRE_RUTINA"].isnull() == False) & (df_bbdd_codigo["CODIGO_SIN_TAB_TRIMEADO_SIN_COMENTARIOS"].str.len() != 0), 
                                        lista_headers_df_diagnostico_donde_buscar_campos_usados])
        
        df_temp.reset_index(drop = True, inplace = True)


        lista_codigo_para_tablas_y_vinculos_temp = df_temp[lista_headers_df_diagnostico_donde_buscar_tablas].values.tolist()
        lista_codigo_para_rutinas_y_variables_temp = df_temp[lista_headers_df_diagnostico_donde_buscar_rutinas_y_variables].values.tolist()
        lista_codigo_para_variables_definidas_por_usuario_temp = df_temp[lista_headers_df_diagnostico_donde_buscar_variables_definidas_por_usuario].values.tolist()

        del df_temp


        lista_dicc_rutinas_codigo_donde_buscar = []
        for dicc in lista_dicc_objetos:

            if dicc["TIPO_OBJETO"] == "RUTINAS_VBA":

                tipo_modulo = dicc["TIPO_MODULO"]
                nombre_modulo = dicc["NOMBRE_MODULO"]
                nombre_rutina = dicc["NOMBRE_OBJETO"]


                #se crean las listas lista_codigo_para_tablas_y_vinculos_preservar_si y lista_codigo_para_tablas_y_vinculos_preservar_no
                #para la busqueda de objetos tipo tablas /vinculos
                #
                # --> lista_codigo_para_tablas_y_vinculos_preservar_si   es lista normal donde cada item son los string encapsulados entre comillas dobles (concatenados)
                #                                                        que parecen a sentencias SQL o de manipulacion de tablas
                #
                # --> lista_codigo_para_tablas_y_vinculos_preservar_no   es lista de sublistas con el tipo modulo, nombre modulo, nombre rutina, numero linea codigo en la rutina,
                #                                                        linea codigo original y los string encapsulados entre comillas dobles (concatenados)
                #                                                        que NO parecen a sentencias SQL o de manipulacion de tablas 
                lista_codigo_para_tablas_y_vinculos_preservar_si = []
                lista_codigo_para_tablas_y_vinculos_preservar_no = []
                for tipo_modulo_codigo, nombre_modulo_codigo, nombre_rutina_codigo, numero_linea_codigo_rutina, codigo_original, codigo_preservar_si, codigo_preservar_no in lista_codigo_para_tablas_y_vinculos_temp:

                    if tipo_modulo == tipo_modulo_codigo and nombre_modulo == nombre_modulo_codigo and nombre_rutina == nombre_rutina_codigo:

                        lista_codigo_para_tablas_y_vinculos_preservar_si.append(codigo_preservar_si)

                        lista_codigo_para_tablas_y_vinculos_preservar_no.append([tipo_modulo_codigo
                                                                                , nombre_modulo_codigo
                                                                                , nombre_rutina_codigo
                                                                                , numero_linea_codigo_rutina
                                                                                , codigo_original
                                                                                , codigo_preservar_no
                                                                                ])

                

                #se crea la lista lista_codigo_para_rutinas_y_variables para la busqueda de objetos tipo rutinas (o funciones) o variables publicas normales
                lista_codigo_para_rutinas_y_variables = []
                for tipo_modulo_codigo, nombre_modulo_codigo, nombre_rutina_codigo, codigo_linea in lista_codigo_para_rutinas_y_variables_temp:
                    
                    if tipo_modulo == tipo_modulo_codigo and nombre_modulo == nombre_modulo_codigo and nombre_rutina == nombre_rutina_codigo:
                        lista_codigo_para_rutinas_y_variables.append(codigo_linea)


                #se crea la lista lista_codigo_para_variables_definidas_por_usuario para la busqueda de ariables publicas definidas por el usuario
                lista_codigo_para_variables_definidas_por_usuario = []
                for tipo_modulo_codigo, nombre_modulo_codigo, nombre_rutina_codigo, codigo_linea in lista_codigo_para_variables_definidas_por_usuario_temp:
                    
                    if tipo_modulo == tipo_modulo_codigo and nombre_modulo == nombre_modulo_codigo and nombre_rutina == nombre_rutina_codigo:
                        lista_codigo_para_variables_definidas_por_usuario.append(codigo_linea)



                dicc_temp = {"TIPO_MODULO": tipo_modulo
                            , "NOMBRE_MODULO": nombre_modulo
                            , "NOMBRE_RUTINA": nombre_rutina
                            , "LISTA_CODIGO":
                                            {"DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI": lista_codigo_para_tablas_y_vinculos_preservar_si
                                            , "DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO": lista_codigo_para_tablas_y_vinculos_preservar_no
                                            , "DIAGNOSTICO_RUTINAS_Y_VARIABLES": lista_codigo_para_rutinas_y_variables
                                            , "DIAGNOSTICO_VARIABLES_DEFINIDAS_POR_USUARIO": lista_codigo_para_variables_definidas_por_usuario
                                            }
                            }

                lista_dicc_rutinas_codigo_donde_buscar.append(dicc_temp)


                del dicc_temp
                del lista_codigo_para_tablas_y_vinculos_preservar_si
                del lista_codigo_para_tablas_y_vinculos_preservar_no
                del lista_codigo_para_rutinas_y_variables
                del lista_codigo_para_variables_definidas_por_usuario


        del lista_codigo_para_tablas_y_vinculos_temp
        del lista_codigo_para_rutinas_y_variables_temp
        del lista_codigo_para_variables_definidas_por_usuario_temp
        del df_bbdd_codigo


        #################################################################################################################################
        #         CALCULO DIAGNOSTICO DEPENDENCIAS
        #################################################################################################################################

        #se calculan las listas lista_diagnostico_dependencias y lista_diagnostico_sin_dependencias y se crean los df para exportar a excel
        lista_diagnostico_dependencias = []
        lista_diagnostico_sin_dependencias = []

        for dicc_objetos_a_buscar in lista_dicc_objetos:

            tipo_objeto_a_buscar = dicc_objetos_a_buscar["TIPO_OBJETO"]
            tipo_modulo_a_buscar = dicc_objetos_a_buscar["TIPO_MODULO"]
            nombre_modulo_a_buscar = dicc_objetos_a_buscar["NOMBRE_MODULO"]
            nombre_objeto_a_buscar = dicc_objetos_a_buscar["NOMBRE_OBJETO"]
            tipo_variable_publica_a_buscar  = dicc_objetos_a_buscar["TIPO_VARIABLE_PUBLICA"]

            nombre_objeto_mayusc_a_buscar = nombre_objeto_a_buscar.upper() 


            #se localiza cual es la lista de lineas de codigo a usar como referencia segun el tipo de objeto
            if tipo_objeto_a_buscar in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:
                key_codigo_donde_buscar = "DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_SI"

            elif tipo_objeto_a_buscar == "RUTINAS_VBA":
                key_codigo_donde_buscar = "DIAGNOSTICO_RUTINAS_Y_VARIABLES"


            elif tipo_objeto_a_buscar == "VARIABLES_VBA":

                if tipo_variable_publica_a_buscar  == dicc_objetos_vba["VARIABLE_PUBLICA_NORMAL"]["STRING_PARA_LISTADO_EXCEL"]:
                    key_codigo_donde_buscar = "DIAGNOSTICO_RUTINAS_Y_VARIABLES"
                    
                elif tipo_variable_publica_a_buscar  == dicc_objetos_vba["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO"]["STRING_PARA_LISTADO_EXCEL"]:
                    key_codigo_donde_buscar = "DIAGNOSTICO_VARIABLES_DEFINIDAS_POR_USUARIO"


            check_objeto = 0
            for dicc_rutinas_codigo_donde_buscar in lista_dicc_rutinas_codigo_donde_buscar:

                tipo_modulo_donde_buscar = dicc_rutinas_codigo_donde_buscar["TIPO_MODULO"]
                nombre_modulo_donde_buscar = dicc_rutinas_codigo_donde_buscar["NOMBRE_MODULO"]
                nombre_rutina_donde_buscar = dicc_rutinas_codigo_donde_buscar["NOMBRE_RUTINA"]
                lista_codigo_donde_buscar = dicc_rutinas_codigo_donde_buscar["LISTA_CODIGO"][key_codigo_donde_buscar]


                #se excluye de la busqueda el propio objeto (cuando este es rutina, las funciones no aplica porque el resultado de la funcion en VBA 
                #usa el nombre de la función para declararla)
                if not (tipo_objeto_a_buscar  == "RUTINAS_VBA" and tipo_modulo_a_buscar == tipo_modulo_donde_buscar and nombre_modulo_a_buscar == nombre_modulo_donde_buscar and 
                        nombre_objeto_a_buscar == nombre_rutina_donde_buscar):

                    check_codigo = 0
                    for linea_codigo in lista_codigo_donde_buscar:
                        
                        if linea_codigo != None:
                            if nombre_objeto_mayusc_a_buscar in linea_codigo:

                                #se considera que el objeto se usa en otra rutina si nombre_objeto_mayusc esta incluido en la linea de codigo de la iteracion sobre
                                #la lista lista_codigo_donde_buscar y que ademas nombre_objeto_mayusc sea una palabra suelta no un substring incluido en otra palabra,
                                #salvo que este seguido de un parentesis de apertura "(" (casos de llamadas de rutinas)
                                # ejemplo: rutina Sub_Calc en las lineas de codigo "Call Sub_Calc()" "Call Sub_Calc_1()" se consideraria que Sub_Calc 
                                #se usa en la 1era linea pero no en la 2nda)
                                check_si_nombre_objeto_es_palabra_sola = fr'(?<![\w_-]){re.escape(nombre_objeto_mayusc_a_buscar)}(?![\w-])'
                                check_si_nombre_objeto_es_palabra_sola = bool(re.search(check_si_nombre_objeto_es_palabra_sola, linea_codigo))

                                check_objeto = 1 if check_si_nombre_objeto_es_palabra_sola == True else 0
                                check_codigo = 1 if check_si_nombre_objeto_es_palabra_sola == True else 0

                                break


                    if check_codigo == 1:
                        lista_diagnostico_dependencias.append([tipo_objeto_a_buscar
                                                                , tipo_modulo_a_buscar
                                                                , nombre_modulo_a_buscar
                                                                , nombre_objeto_a_buscar
                                                                , tipo_modulo_donde_buscar
                                                                , nombre_modulo_donde_buscar
                                                                , nombre_rutina_donde_buscar])
                        
                                                                                    
            if check_objeto == 0:
                lista_diagnostico_sin_dependencias.append([tipo_objeto_a_buscar
                                                            , tipo_modulo_a_buscar
                                                            , nombre_modulo_a_buscar
                                                            , nombre_objeto_a_buscar])



        #se calcula el diagnostico de check manual para tablas / vinculos donde los string encapsulados entre comillas dobles
        #contienen nombres de tablas / vinculos pero el string no se parece a sentencias SQL o de manipulacion de tablas via VBA pe
        #se hace en otro bucle sobre lista_dicc_objetos para aligerar el codigo y hacerlo mas legible
        lista_diagnostico_tablas_check_manual = []
        for dicc_objetos_a_buscar in lista_dicc_objetos:

            tipo_objeto_a_buscar = dicc_objetos_a_buscar["TIPO_OBJETO"]

            #se busca tan solo para las tablas / vinculos
            if tipo_objeto_a_buscar in ["TABLA_LOCAL", "VINCULO_ODBC", "VINCULO_OTRO"]:

                nombre_objeto_a_buscar = dicc_objetos_a_buscar["NOMBRE_OBJETO"]
                nombre_objeto_mayusc_a_buscar = nombre_objeto_a_buscar.upper()

                #mediante bucle sobre lista_dicc_rutinas_codigo_donde_buscar se recupera la lista almacenada en la subkey_2
                #DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO
                for dicc_rutinas_codigo_donde_buscar in lista_dicc_rutinas_codigo_donde_buscar:

                    tipo_modulo_donde_buscar_check = dicc_rutinas_codigo_donde_buscar["TIPO_MODULO"]
                    nombre_modulo_donde_buscar_check = dicc_rutinas_codigo_donde_buscar["NOMBRE_MODULO"]
                    nombre_rutina_donde_buscar_donde_buscar_check = dicc_rutinas_codigo_donde_buscar["NOMBRE_RUTINA"]
                    lista_codigo_donde_buscar_check = dicc_rutinas_codigo_donde_buscar["LISTA_CODIGO"]["DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO"]

                    #se chequea si para la tabla de la iteracion (nombre_objeto_a_buscar) no se localizo anterioremente una dependencia 
                    #con la rutina nombre_rutina_donde_buscar_donde_buscar_check
                    #en caso de que NO sea el caso se prsigue con el matcho con la lista contenida en la subkey_2 DIAGNOSTICO_TABLAS_Y_VINCULOS_PRESERVAR_NO
                    check_si_dependencia_localizada = 0
                    for tipo_objeto_depend, tipo_modulo_depend, nombre_modulo_depend, nombre_objeto_depend, tipo_modulo_depend, nombre_modulo_depend, nombre_rutina_depend in lista_diagnostico_dependencias:

                        if (tipo_objeto_a_buscar == tipo_objeto_depend and nombre_objeto_a_buscar == nombre_objeto_depend and tipo_modulo_donde_buscar_check == tipo_modulo_depend and
                            nombre_modulo_donde_buscar_check == nombre_modulo_depend and nombre_rutina_donde_buscar_donde_buscar_check == nombre_rutina_depend):

                            check_si_dependencia_localizada = 1
                            break

                    if check_si_dependencia_localizada == 0:

                        #mediante bucle sobre la lista lista_codigo_donde_buscar_check se realizan los matcheos de si se encuentra exactamente el nombre de la tabla 
                        #(como palabra sola y no como substring dentro de otra palabra mas larga) se agrega a la lista lista_diagnostico_tablas_check_manual la info
                        #necesaria para determinar si se ha de considerar o no usade tabla en la rutina correspondiente
                        for tipo_modulo_check, nombre_modulo_check, nombre_rutina_check, numero_linea_codigo_rutina_check, codigo_linea_original_check, string_concat_check in lista_codigo_donde_buscar_check:
                                                                                                                                                                            
                            if nombre_objeto_mayusc_a_buscar in string_concat_check:

                                check_manual_si_nombre_objeto_no_en_otra_palabra = fr'(?<![\w_-]){re.escape(nombre_objeto_mayusc_a_buscar)}(?![\w-])'
                                check_manual_si_nombre_objeto_no_en_otra_palabra = bool(re.search(check_manual_si_nombre_objeto_no_en_otra_palabra, string_concat_check))

                                if check_manual_si_nombre_objeto_no_en_otra_palabra == True:
                                    lista_diagnostico_tablas_check_manual.append([tipo_objeto_a_buscar
                                                                                , nombre_objeto_a_buscar
                                                                                , tipo_modulo_check
                                                                                , nombre_modulo_check
                                                                                , nombre_rutina_check
                                                                                , numero_linea_codigo_rutina_check
                                                                                , codigo_linea_original_check])





        #######################################################
        # CREACION DF PARA EXPORT EXCEL
        #######################################################

        #se crean los df df_diagnostico_dependencias y df_diagnostico_sin_dependencias para exportar a excel
        df_diagnostico_dependencias = pd.DataFrame(lista_diagnostico_dependencias, columns = lista_headers_df_dependencias)
        df_diagnostico_dependencias = (df_diagnostico_dependencias[lista_headers_df_dependencias].sort_values([i for i in lista_headers_df_dependencias]
                                                                                , ascending = [True for i in lista_headers_df_dependencias]))

        df_diagnostico_sin_dependencias = pd.DataFrame(lista_diagnostico_sin_dependencias, columns = lista_headers_df_sin_dependencias)
        df_diagnostico_sin_dependencias = (df_diagnostico_sin_dependencias[lista_headers_df_sin_dependencias].sort_values([i for i in lista_headers_df_sin_dependencias]
                                                                                , ascending = [True for i in lista_headers_df_sin_dependencias]))


        #se crean el df df_diagnostico_tablas_check_manual
        df_diagnostico_tablas_check_manual = pd.DataFrame(lista_diagnostico_tablas_check_manual, columns = lista_headers_df_check_manual_tablas)
        df_diagnostico_tablas_check_manual = (df_diagnostico_tablas_check_manual[lista_headers_df_check_manual_tablas].sort_values([i for i in lista_headers_df_check_manual_tablas]
                                                                                , ascending = [True for i in lista_headers_df_check_manual_tablas]))


        #se crea el df de listado de objetos para poder exportarlo a excel (mas abajo)
        lista_para_df_listado = []
        for dicc in lista_dicc_objetos:
            
            tipo_objeto_2 = dicc["TIPO_RUTINA"] if dicc["TIPO_OBJETO"] == "RUTINAS_VBA" else dicc["TIPO_VARIABLE_PUBLICA"] if dicc["TIPO_OBJETO"] == "VARIABLES_VBA" else None

            lista_para_df_listado.append([dicc["TIPO_OBJETO"]
                                        , dicc["TIPO_MODULO"]
                                        , dicc["NOMBRE_MODULO"]
                                        , dicc["NOMBRE_OBJETO"]
                                        , tipo_objeto_2
                                        , dicc["TIPO_DECLARACION_RUTINA"]
                                        , dicc["PARAMETROS_RUTINA"]
                                        , dicc["TIPO_DATO_VARIABLE_PUBLICA"]
                                        , dicc["VARIABLE_PUBLICA_ES_CONSTANTE"]
                                        , dicc["VARIABLE_PUBLICA_CONSTANTE_VALOR"]
                                        , dicc["VARIABLE_PUBLICA_DEFINIDA_POR_USUARIO_SUB_VARIABLES"]
                                        , dicc["CONNECTING_STRING_VINCULOS"]
                                        ])

        df_listado_objetos = pd.DataFrame(lista_para_df_listado, columns = lista_headers_df_listado)
        df_listado_objetos = df_listado_objetos[lista_headers_df_listado].sort_values(lista_headers_df_listado, ascending = [True for i in lista_headers_df_listado])



        #######################################################
        # PARA EXPORT EXCEL
        #######################################################

        now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
        saveas = str(ruta_destino_excel_diagnostico_access) + r"\DEPENDENCIAS_MS_ACCESS_"  + now + ".xlsb"

        shutil.copyfile(mod_gen.ruta_plantilla_diagnostico_access_xls, saveas)

        app = xw.App(visible = False)
        wb = app.books.open(saveas, update_links = False)

        dicc_export_excel = {"LISTADO": 
                                        {"HOJA_EXCEL": "LISTADO"
                                        , "DF_EXPORT": df_listado_objetos
                                        }
                            , "DEPENDENCIAS": 
                                        {"HOJA_EXCEL": "DEPENDENCIAS"
                                        , "DF_EXPORT": df_diagnostico_dependencias
                                        }
                            , "SIN_DEPENDENCIAS": 
                                        {"HOJA_EXCEL": "SIN DEPENDENCIAS"
                                        , "DF_EXPORT": df_diagnostico_sin_dependencias
                                        }
                            , "TABLAS_CHECK_MANUAL": 
                                        {"HOJA_EXCEL": "TABLAS (CHECK MANUAL)"
                                        , "DF_EXPORT": df_diagnostico_tablas_check_manual
                                        }
                            }

        for key in dicc_export_excel.keys():

            hoja_excel = dicc_export_excel[key]["HOJA_EXCEL"]
            df_export = dicc_export_excel[key]["DF_EXPORT"]

            if len(df_export) != 0:
                ws = wb.sheets[hoja_excel]
                ws.range("A2:FF65000").clear_contents()
                ws["A2"].options(pd.DataFrame, header = 0, index = False, expand = "table").value = df_export



        #se guarda en el excel, se cierra y se re-abre en 1er plano        
        wb.save(saveas)
        wb.close()
        app = xw.App(visible = True)
        app.quit()

        wb = xw.Book(saveas, update_links = False)



        del lista_dicc_objetos
        del lista_dicc_rutinas_codigo_donde_buscar
        del lista_diagnostico_dependencias
        del lista_diagnostico_sin_dependencias
        del lista_diagnostico_tablas_check_manual
        del df_listado_objetos
        del df_diagnostico_dependencias
        del df_diagnostico_sin_dependencias
        del df_diagnostico_tablas_check_manual
        del dicc_export_excel



    except Exception as Err:

        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos["PROCESO_02"]["MS_ACCESS"]["LISTA_DICC_ERRORES_CALCULO"]

        dicc_errores_temp = {"TIPO_BBDD": "MS_ACCESS"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos["PROCESO_02"]["MS_ACCESS"]["LISTA_DICC_ERRORES_CALCULO"] = lista_dicc_errores_migracion
        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso



