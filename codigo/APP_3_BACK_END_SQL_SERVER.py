import pandas as pd
import traceback
import pyodbc
import os
import re
import shutil
import xlwings as xw
import datetime as dt
import warnings

import APP_2_GENERAL as mod_gen


#################################################################################################################################################################################
##                     VARIABLES GENERALES
#################################################################################################################################################################################


#############################################
# connecting string SQL Server y relacionados
#############################################

#conn_str_sql_server_windows_authentication + conn_str_sql_server_login_password_authentication se usan al seleccionar en la GUI un servidor
#--> se intenta la conexion 1ero mediante la connecting string de windows authentication (donde se reemplaza previamente REEMPLAZA_SERVIDOR y REEMPLAZA_BBDD 
#    por el servidor seleccionado y master respectivamente)
#--> si el bloque da error (es un try except) se abre la GUI de SQL Server authentication  para poder registrar login y password y se prueba
#    la connecting string conn_str_sql_server_login_password_authentication donde se reemplaza REEMPLAZA_SERVIDOR, REEMPLAZA_BBDD, REEMPLAZA_LOGIN y REEMPLAZA_PASSWORD
#    por el servidor seleccionado, master, el login registrado y el password registrado
#--> este proceso permite informar la variable conn_str_sql_server para almacenarla en el diccionario dicc_codigos_bbdd (modulo general) para posterior uso en las distintas
#    rutinas del app cuando se requiera

driver_sql_server = "SQL Server Native Client 11.0"
conn_str_sql_server_windows_authentication = (r'Driver={' + driver_sql_server + '};'r'Server=REEMPLAZA_SERVIDOR;'r'Database=REEMPLAZA_BBDD;'r'Trusted_Connection=yes;Connection Timeout=0')
conn_str_sql_server_login_password_authentication = (r'Driver={' + driver_sql_server + '};'r'Server=REEMPLAZA_SERVIDOR;'r'Database=REEMPLAZA_BBDD;'r'UID=REEMPLAZA_LOGIN;'r'PWD=REEMPLAZA_PASSWORD;'r'Connection Timeout=0')
conn_str_sql_server = None #se establece a None de inicio (se actualiza al intentar realizar la conexion al servidor)

#la variable codigo_error_pyodbc_connexion_por_windows_authentication almacena el numero de error que debe dar el bloque try except cuando se intenta conectar 
#por windows authentication cuando la conexion esta en realidad configurada por SQL Server authentication
codigo_error_pyodbc_connexion_por_windows_authentication = 28000



#############################################
# listas para combobox de la GUI
#############################################

#lista que se usa en la GUI de inicio y permite seleccionar los servidores que se quieran que contenga el app
lista_GUI_sql_server_servidor = []

#lista que se usa en la GUI de diagnostico de dependencias y permite seleccionar el tipo de proceso a ejecutar
#IMPORTANTE: no cambiar el orden de los indices sino habra que cambiarlos donde se usa esta lista
lista_GUI_diagnostico_combobox_sql_server = ["Realizar diagnostico", "Descargar códigos T-SQL"]



#############################################
# queries SQL Server sobre tablas de sistema
#############################################

#query que permite determinar si el usuario tiene para una bbdd (a la que se ha conectado previamente mediante conecting string con el servidor seleccionado y esta misma bbdd)
#permisos de acceso al codigo T-SQL del objeto (permission_name = 'VIEW DEFINITION' que aplica no solo a views sino tambien a tablas, stored procedures y funciones)
query_sql_server_bbdd_permisos = """SELECT permission_name
                                    FROM fn_my_permissions(NULL, 'DATABASE')
                                    WHERE permission_name = 'VIEW DEFINITION'"""


#query que permite listar todas las bbdd de un servidor sea cual sea el permiso que tiene el usuario sobre ellas (excluyendo las de sistema como master, tempdb, model y msdb)
query_lista_bbdd_sql_server = """SELECT UPPER(name) FROM master.dbo.sysdatabases
                            WHERE name not in ('master', 'tempdb', 'model', 'msdb')
                            ORDER BY name"""


#query que permite listar para un servidor y bbdd determinado todos los objetos que contiene
#filtrando solo por tablas (U), views (V), stored procedures (P) y funciones (FN, IF, TF y AF) 
#cuando se usa se ha de hacer replace de REEMPLAZA_BBDD por la bbdd
#
#el app se extralimita a tablas, views, stored procedures y funciones cuando en realidad hay mas objetos
#
#IMPORTANTE: los nombres que toma TIPO_OBJETO (ver el CASE WHEN) han de ser los mismos que los que aparecen en
#             el diccionario dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"] (modulo general) --> TABLAS, VIEWS, STORED_PROCEDURES y FUNCIONES
query_objetos_sql_server = """SELECT
                                'REEMPLAZA_BBDD' AS BBDD
                                , T2.name AS REPOSITORIO
                                , CASE WHEN T1.type = 'U' THEN 'TABLAS'
                                       WHEN T1.type = 'V' THEN 'VIEWS'
                                       WHEN T1.type = 'P' THEN 'STORED_PROCEDURES'
                                       WHEN T1.type IN ('FN', 'IF', 'TF', 'AF') THEN 'FUNCIONES'
                                END AS TIPO_OBJETO
                                , T1.name AS NOMBRE_OBJETO
                                , 'REEMPLAZA_BBDD' + '.' + T2.name + '.' + T1.name AS OBJETO_LONG
                                , T2.name + '.' + T1.name AS OBJETO_SHORT

                                FROM REEMPLAZA_BBDD.sys.objects T1
                                LEFT JOIN REEMPLAZA_BBDD.sys.schemas T2 on T1.schema_id = T2.schema_id

                                WHERE UPPER(T1.type) IN ('U', 'V', 'P', 'FN', 'IF', 'TF', 'AF')

                                ORDER BY T1.type, T2.name, T1.name
                            """

query_objetos_sql_server_parametros = """SELECT
                                            'REEMPLAZA_BBDD' AS BBDD
                                            , T3.name AS REPOSITORIO
                                            , CASE WHEN T2.type = 'U' THEN 'TABLAS'
                                                    WHEN T2.type = 'V' THEN 'VIEWS'
                                                    WHEN T2.type = 'P' THEN 'STORED_PROCEDURES'
                                                    WHEN T2.type IN ('FN', 'IF', 'TF', 'AF') THEN 'FUNCIONES'
                                            END AS TIPO_OBJETO
                                            , T2.name AS NOMBRE_OBJETO 
                                            , CASE WHEN UPPER(T4.name) LIKE '%CHAR%' THEN CONCAT(T1.name, ' ', UPPER(T4.name), '(', CAST(T1.max_length AS VARCHAR(100)), ')') 
                                            ELSE CONCAT(T1.name, ' ', UPPER(T4.name))
                                            END AS PARAMETROS

                                            FROM REEMPLAZA_BBDD.sys.parameters T1
                                            LEFT JOIN REEMPLAZA_BBDD.sys.objects T2 ON T1.object_id = T2.object_id
                                                LEFT JOIN REEMPLAZA_BBDD.sys.schemas T3 on T2.schema_id = T3.schema_id

                                            LEFT JOIN REEMPLAZA_BBDD.sys.types T4 ON T1.user_type_id = T4.user_type_id

                                            WHERE T2.type IN ('P', 'FN', 'IF', 'TF') AND LEN(COALESCE(T1.name, '')) <>0
                                            ORDER BY T2.name, T1.parameter_id
                                        """



#query que permite calcular el tipo de dato asociado a los campos de una tabla en un servidor y bbdd determindado
#se usa para poder calcular la sentencia CREATE TABLE del objeto tabla relacionado
#cuando se usa se debe hacer replace de REEMPLAZA_BBDD y REEMPLAZA_OBJETO_LONG por la bbdd y por el nombre de objeto
#en formato (bbdd.esquema.nombre_tabla)
query_campos_tablas_sql_server = """SELECT 
                                    CASE WHEN CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN CONCAT('[', COLUMN_NAME, '] ', UPPER(DATA_TYPE), '(', CHARACTER_MAXIMUM_LENGTH, ')')
                                    ELSE CONCAT('[', COLUMN_NAME, '] ', UPPER(DATA_TYPE))
                                    END AS CAMPO_CON_TIPO_DATO
                                    FROM REEMPLAZA_BBDD.INFORMATION_SCHEMA.COLUMNS
                                    WHERE TABLE_CATALOG + '.' + TABLE_SCHEMA + '.' + TABLE_NAME = 'REEMPLAZA_OBJETO_LONG'
                                    ORDER BY TABLE_CATALOG + '.' + TABLE_SCHEMA + '.' + TABLE_NAME, ORDINAL_POSITION
                                """


#query que permite encapsular en un pandas df el codigo T-SQL de un objeto (views, stored procedures y funciones --> no va con las tablas, se hace de otra forma)
#la query es un mini-proceso en tiempo de ejecucion que:
#--> elimina de la bbdd de sistema tempdb la tabla temporal #TEMP (si existe)
#--> crea la tabla temporal #TEMP con 2 campos NUMERO_LINEA_CODIGO (que es un INT autoincremental para almacenar el numero de linea del codigo) y CODIGO
#--> se agrega a esta tabla temporal el CODIGO resultante de ejecutar la sentencia sp_helptext
#
#cuando se usa se debe hacer replace de REEMPLAZA_BBDD y REEMPLAZA_OBJETO_LONG por la bbdd y por el nombre de objeto
#en formato (bbdd.esquema.nombre_tabla)
query_helptext_sql_server = """USE REEMPLAZA_BBDD

                                IF OBJECT_ID('tempdb.dbo.#TEMP', 'U') IS NOT NULL
                                    DROP TABLE #TEMP;

                                CREATE TABLE #TEMP (
                                    NUMERO_LINEA_CODIGO INT IDENTITY(1,1)
                                    , CODIGO NVARCHAR(MAX)
                                );

                                INSERT INTO #TEMP (CODIGO)
                                EXEC sp_helptext 'REEMPLAZA_OBJETO_LONG';
                                """


#query que permite eliminar objetos (tablas, views, stored procedures y funciones) en el proceso de merge en bbdd fisica
#si previamente ya existen para poder reemplazarlos por el objeto resultante de los cambios en el codigo realizados por el usuario
#
#cuando se usa se debe hacer replace de REEMPLAZA_TIPO_OBJETO, REEMPLAZA_OBJETO_SHORT y REEMPLAZA_ID_OBJETO por el tipo de objeto, el nombre de objeto en formato (esquema.nombre_tabla) y por el ID de objeto 
#(U para tablas, V para views, P para stored procedures, ['FN', 'IF', 'TF', 'AF'] para funciones) - no se hace con DROP TABLE IF EXISTS pq solo funciona en versiones recientes de SQL SERVER
query_drop_if_exists_objeto = """IF OBJECT_ID('REEMPLAZA_OBJETO_SHORT', 'REEMPLAZA_ID_OBJETO') IS NOT NULL
                                    BEGIN
                                        DROP REEMPLAZA_TIPO_OBJETO REEMPLAZA_OBJETO_SHORT;
                                    END
                                """


#df con el encabezado de codigo para la descarga de codigos en ficheros .sql (opcion diagnostico)
df_encabezado_descarga_codigo_objetos = pd.DataFrame({"CODIGO": ["USE [REEMPLAZA_BBDD]\n", "GO\n\n", "SET ANSI_NULLS ON\n", "GO\n", "SET QUOTED_IDENTIFIER ON\n", "GO\n\n"]})




#############################################
# indicadores de codigo comentado en T-SQL
#############################################

#si la linea empieza por este indicador (previo trimeo) la linea de codigo completa esta comentada
indicador_sql_server_comentario_linea = "--"

#indicadores de comentarios por bloques de linea (ini es el de apertura y fin el de cierre)
indicador_sql_server_comentario_bloque_linea_ini = "/*"
indicador_sql_server_comentario_bloque_linea_fin = "*/"



#############################################
# listas de columnas de df usados en el app
#############################################

#lista headers de df_codigo asociado al objeto
lista_headers_campos_df_codigo_objetos = ["NUMERO_LINEA_CODIGO", "CODIGO"]


#lista headers para exportar a excel el diagnostico de dependencias
lista_headers_sql_server_diagnostico_df_bbdd_selecc = ["SERVIDOR", "BBDD"]
lista_headers_sql_server_diagnostico_df_listado_objetos = ["BBDD", "REPOSITORIO", "TIPO_OBJETO", "NOMBRE_OBJETO", "PARAMETROS"]

lista_headers_sql_server_diagnostico_df_dependencias = ["BBDD", "REPOSITORIO", "TIPO_OBJETO", "NOMBRE_OBJETO", "SE_USA_EN_BBDD", "SE_USA_EN_REPOSITORIO", "SE_USA_EN_TIPO_OBJETO", "SE_USA_EN_NOMBRE_OBJETO"]

lista_headers_sql_server_diagnostico_df_sin_dependencias = ["BBDD", "REPOSITORIO", "TIPO_OBJETO", "NOMBRE_OBJETO"]



#################################################################################################################################################################################
##                     FUNCIONES
#################################################################################################################################################################################

def func_sql_server_tipo_conexion_servidor(servidor):
    #funcion que permite determinar el tipo de conexión a SQL Server (WINDOWS_AUTHENTICATION o SQL_SERVER_AUTHENTICATION)

    try:
        tipo_connecting_string = None
        conn_str = conn_str_sql_server_windows_authentication.replace("REEMPLAZA_SERVIDOR", servidor).replace("REEMPLAZA_BBDD", "master")
        MiConexion = pyodbc.connect(conn_str)

    except pyodbc.Error as Err:
        if Err.args[0] == codigo_error_pyodbc_connexion_por_windows_authentication:
            tipo_connecting_string = "SQL_SERVER_AUTHENTICATION"
        else:
            tipo_connecting_string = None
        
        pass

    else:
        tipo_connecting_string = "WINDOWS_AUTHENTICATION"
        MiConexion.close()

    return tipo_connecting_string



def func_sql_server_df_codigo_objeto(opcion, **kwargs):
    #funcion con 2 opciones que permite:
    # --> DF_CODIGO_ORIGINAL        devuelve un df con el codigo T-SQL asociado al objeto (en el caso de las tablas es la sentencia CREATE TABLE)
    #                               para los demas objetos que no son tablas se hace con la instrucion sp_helptext incluida en la query query_helptext_sql_server
    #
    # --> LISTA_CODIGO_LIMPIO       devuelve una lista donde cada item son las lineas con el codigo original donde se pone todo en mayuscula, se trimea, se quitan las tabulaciones, 
    #                               se quitan los corchetes "[" y "] y se quitan las lineas vacias, se quitan tambien todos los comentarios sean de linea (los que empiezan por "--") 
    #                               o sean por bloques (los que vienen encapsulados por /* y */)

    #parametros kwargs
    tipo_objeto = kwargs.get("tipo_objeto", None)
    MiConexion = kwargs.get("MiConexion", None)
    bbdd = kwargs.get("bbdd", None)
    objeto_long = kwargs.get("objeto_long", None)
    cursor = kwargs.get("cursor", None)
    objeto_short = kwargs.get("objeto_short", None)
    df_codigo_original = kwargs.get("df_codigo_original", None)


    ########################################################################################
    # DF_CODIGO_ORIGINAL
    ########################################################################################
    if opcion == "DF_CODIGO_ORIGINAL":

        if tipo_objeto == "TABLAS":

            query_temp = query_campos_tablas_sql_server.replace("REEMPLAZA_BBDD", bbdd).replace("REEMPLAZA_OBJETO_LONG", objeto_long)
            df_temp = pd.read_sql_query(query_temp, MiConexion)


            #se crea el df con la sentencia CREATE TABLE (se agrega el numero de linea de codigo)
            lista_df = [[1, "CREATE TABLE " + objeto_short], 
                        [2, "\t("]]
            
            for ind in df_temp.index:
                lista_temp = [ind + 3, df_temp.iloc[ind, 0]] if ind == 0 else [ind + 3, ", " + df_temp.iloc[ind, 0]]
                lista_df.append(lista_temp)
                del lista_temp

            lista_df.append([len(lista_df) + 3, ")"])
            df_codigo_original = pd.DataFrame(lista_df, columns = lista_headers_campos_df_codigo_objetos)

            del lista_df


        elif tipo_objeto != "TABLAS":

            #la query query_helptext_sql_server ya viene con el NUMERO_LINEA autoincremental
            query_temp = query_helptext_sql_server.replace("REEMPLAZA_BBDD", bbdd).replace("REEMPLAZA_OBJETO_LONG", objeto_long)

            cursor.execute(query_temp)
            MiConexion.commit()

            df_codigo_original = pd.read_sql_query("select * from #TEMP", MiConexion)
            df_codigo_original["CODIGO"] = df_codigo_original["CODIGO"].apply(lambda x: x.replace("\n", ""))#sp_helptext viene con 2 saltos de linea (se quita 1)


    ########################################################################################
    # LISTA_CODIGO_LIMPIO
    ########################################################################################
    elif opcion == "LISTA_CODIGO_LIMPIO":

        #se hace limpieza del codigo donde se pone en mayusculas (para evitar casos de case sensitive para cuando se realice la busqueda de dependencias)
        #y se quitan las tabulaciones, saltos de linea (los df que se calculan con la instruccion sp_helptext vienen con 2 saltos de linea por lo que se quita 1), se trimea, 
        # se quitan los corchetes "[" y "]", se quitan las lineas vacias
        df_codigo_original["CODIGO"] = df_codigo_original["CODIGO"].apply(lambda x: x.replace("\r\n", "").replace("\t", "").replace("[", "").replace("]", "").upper().strip())

        df_codigo_original = df_codigo_original.loc[df_codigo_original["CODIGO"].str.len() != 0, ["CODIGO"]]
        df_codigo_original.reset_index(drop = True, inplace = True)


        #se crea el string del codigo
        string_codigo = "\n".join([df_codigo_original.iloc[ind, 0] for ind in df_codigo_original.index])


        #se localizan los string encapsulados entre comillas dobles para hacer un replace temporal de sus contenidos 
        #(que pueden contener "/*" y "*/" y tambien "--" aqui no serian aperturas y/o cierres de comentario)
        #el replace se hace por @@@ (+ 1 indice) pq @@@ no se puede declarar en codigo T-SQL fuera de comillas simples
        #por lo que no hay riesgo de remplazar una sentencia ya existente
        lista_string_entre_comillas_simples = re.findall(r"'([^']*)'", string_codigo)
        lista_string_entre_comillas_simples = [["'" + item + "'", "@@@" + str(ind + 1)] for ind, item in enumerate(lista_string_entre_comillas_simples)]

        for item, item_replace in lista_string_entre_comillas_simples:
            string_codigo.replace(item, item_replace, 1)


        #se quitan los comentarios por bloques (/* y */)
        lista_codigo_caracteres_sin_commentarios = []
        dentro_comentario = False
        caracter = 0

        while caracter < len(string_codigo):

            #se detectan las aperturas de comentario en bloque "/*"
            if not dentro_comentario and string_codigo[caracter:caracter + 2] == indicador_sql_server_comentario_bloque_linea_ini:
                dentro_comentario = True
                caracter += 2 # es para saltar los caracteres /*


            #se detectan los cierres de comentario en bloque "*/"
            elif dentro_comentario and string_codigo[caracter:caracter + 2] == indicador_sql_server_comentario_bloque_linea_fin:
                dentro_comentario = False
                caracter += 2 # es para saltar los caracteres */


            #se agregan solo los caracteres fuera de los comentarios
            elif not dentro_comentario:
                lista_codigo_caracteres_sin_commentarios.append(string_codigo[caracter])
                caracter += 1
            else:
                caracter += 1

        lista_codigo_sin_commentarios_bloques = "".join(lista_codigo_caracteres_sin_commentarios).split("\n")
        del lista_codigo_caracteres_sin_commentarios



        #se quitan los comentarios de linea ("--")
        #para cada item de lista_codigo_sin_commentarios_bloques, se splitea por "--" para recuperar tan solo el 1er item de la lista resultante
        lista_codigo_sin_commentarios = []
        for linea_codigo_sin_commentarios_bloques in lista_codigo_sin_commentarios_bloques:

            linea_codigo_ajust_sin_comentario_linea = linea_codigo_sin_commentarios_bloques.replace(indicador_sql_server_comentario_bloque_linea_fin, "").replace(indicador_sql_server_comentario_bloque_linea_ini, "")
            linea_codigo_ajust_sin_comentario_linea = linea_codigo_ajust_sin_comentario_linea.split(indicador_sql_server_comentario_linea)[0]

            if len(linea_codigo_ajust_sin_comentario_linea) != 0:
                lista_codigo_sin_commentarios.append(linea_codigo_ajust_sin_comentario_linea)

        del lista_codigo_sin_commentarios_bloques



        #se hacen los replace de lo encapsulado entre comillas simples para recuperar el codigo correcto
        string_codigo_limpiado = "\n".join(lista_codigo_sin_commentarios)
        del lista_codigo_sin_commentarios

        for item, item_replace in lista_string_entre_comillas_simples:
            string_codigo_limpiado.replace(item, item_replace)
        
        del lista_string_entre_comillas_simples


        #para los casos donde objeto long y/o objeto short esten escritos en varias lineas se restablece todo en la misma linea
        #y se crea la lista lista_codigo_limpiado
        lista_codigo_limpiado = string_codigo_limpiado.replace("\n.\n", ".").replace("\n.", ".").replace(".\n", ".").split("\n")


    #resultado de la funcion
    return df_codigo_original if opcion == "DF_CODIGO_ORIGINAL" else lista_codigo_limpiado


#################################################################################################################################################################################
##                     RUTINAS ASOCIADAS A PERMISOS EN BBDD
#################################################################################################################################################################################

def def_sql_server_servidor_permisos(opcion_bbdd, path_servidor):
    #permite localizar segun el servidor seleccionado que permisos tiene el usuario (se usa la variable global global_acceso_servidor_selecc = SI o NO)
    #tambien segun el servidor seleccionado se informa la variable global global_servidor_bbdd_permisos_acceso_codigo donde se listan todas las bbdd
    #del servidor seleccionado donde hay permisos de acceso al codigo

    global global_acceso_servidor_selecc
    global global_servidor_bbdd_permisos_acceso_codigo

    warnings.filterwarnings("ignore")

    global_acceso_servidor_selecc = None
    global_servidor_bbdd_permisos_acceso_codigo = None

    if mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"] != None:
        
        try:#se chequea si el usuario puede acceder al servidor
            conn_str = mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"].replace("REEMPLAZA_SERVIDOR", path_servidor).replace("REEMPLAZA_BBDD", "master")
            MiConexion_1 = pyodbc.connect(conn_str)
            
        except:#en caso de que no se informa global_acceso_servidor_selecc = NO 413
            global_acceso_servidor_selecc = "NO"
            pass

        else:
            global_acceso_servidor_selecc = "SI"

            #se listan todas las bbdd del servidor seleccionado (excluyendo las de master, tempdb, model y msdb)
            query_temp = query_lista_bbdd_sql_server
            df_temp = pd.read_sql_query(query_temp, MiConexion_1)
            MiConexion_1.close()

            lista_bbdd_servidor = [df_temp.iloc[i, 0] for i in df_temp.index] if len(df_temp) != 0 else []

            #para el servidor seleccionado se listan las bbdd con permisos de acceso al codigo (variable permisos_sql_accesso_codigos)
            #en global_servidor_bbdd_permisos_acceso_codigo para poder actualizar los combobox de bbdd en la GUI de inicio
            global_servidor_bbdd_permisos_acceso_codigo = []
            if len(lista_bbdd_servidor) != 0:

                for bbdd in lista_bbdd_servidor:

                    try:#se conecta a la bbdd de la iteracion
                        conn_str = mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"].replace("REEMPLAZA_SERVIDOR", path_servidor).replace("REEMPLAZA_BBDD", bbdd)
                        MiConexion_2 = pyodbc.connect(conn_str)

                    except:
                        pass

                    else:
                        #si se establece la conexion a la bbdd y si la query query_sql_server_bbdd_permisos filtrada por permisos_sql_accesso_codigos
                        #devuelve un df no vacio se agrega la bbdd a global_servidor_bbdd_permisos_acceso_codigo
                        query_permiso = query_sql_server_bbdd_permisos
                        df_permiso = pd.read_sql_query(query_permiso, MiConexion_2)

                        if len(df_permiso) != 0:
                            global_servidor_bbdd_permisos_acceso_codigo.append(bbdd)

                        MiConexion_2.close()



#################################################################################################################################################################################
##                     RUTINA SQL SERVER - COMUN A CONTROL DE VERSIONES Y AL DIAGNOSTICO 456
#################################################################################################################################################################################

def def_proceso_sql_server_1_import(proceso_id, opcion_bbdd, sql_server_servidor, lista_sql_server_bbdd):
    #realiza la importacion del codigo T-SQL de los objetos de la bbdd sql server seleccionada (opcion_bbdd)
    #se crea la lista de diccionarios lista_dicc_objetos_sql_server donde se lista en cada diccionario que compone los distintos los objetos
    #las keys son las siguientes:
    # --> SERVIDOR                   es el nombre del servidor
    # --> BBDD                       es el nombre de la base datos
    # --> REPOSITORIO                es el nombre del esquema
    # --> TIPO_OBJETO                TABLAS, VIEWS, STORED_PROCEDURES y FUNCIONES (son los nombres de las subkeys_3 del diccionario dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"])
    # --> NOMBRE_OBJETO              nombre del objeto
    # --> OBJETO_LONG                nombre del objeto en formato nombre_bbdd.nombre_esquema.nombre_objeto
    # --> OBJETO_SHORT               nombre del objeto en formato nombre_esquema.nombre_objeto
    #
    # --> DF_CODIGO_ORIGINAL         es el df con el codigo T-SQL del objeto (tiene 2 columnas NUMERO_LINEA_CODIGO y CODIGO)
    #                                (se usa tanto en el control de versiones como en el diagnostico de dependencias)
    #
    # --> LISTA_CODIGO_LIMPIADO      es una lista donde cada item es una linea con el codigo ajustado para poder calcular el diagnostico de dependencias donde 
    #                                se quitan los comentarios, puesto que el diagnostico al igual que para MS_ACCESS se realiza con codigo activo, 
    #                                en T-SQL los comentarios pueden ir por bloques (encapsulados entre /* y */) y pueden afectar varias lineas que se netean aqui
    #                                (ES NECESARIO que sea un string y no un df para poder localizar objetos donde las combinaciones de objeto long o short pueden ir en varias lineas)


    warnings.filterwarnings("ignore")


    try:

        #se crea la conexion con el servidor seleccionado
        conn_str = mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["CONNECTING_STRING"].replace("REEMPLAZA_SERVIDOR", sql_server_servidor).replace("REEMPLAZA_BBDD", "master")
        MiConexion = pyodbc.connect(conn_str)
        cursor = MiConexion.cursor()


        #se crea la lista de bbdd que conforman el servidor (se usa solo en el diagnostico de dependencias)
        df_temp = pd.read_sql_query(query_lista_bbdd_sql_server, MiConexion)
        lista_sql_server_bbdd_servidor = [df_temp.iloc[ind, 0] for ind in df_temp.index] if proceso_id == "PROCESO_03" else None
        del df_temp


        #se crea el df con los objetos de las distintas bbdd de la lista lista_sql_server_bbdd
        #se crea tambien el df df_objetos_sql_server_parametros con los distintos parametros de objetos tipo stored procedures y funciones
        #(solo aplica si se opta por el diagnostico de dependencias)
        df_objetos_sql_server = None
        df_objetos_sql_server_parametros = None
        for ind, bbdd in enumerate(lista_sql_server_bbdd):

            df_temp_1 = pd.read_sql_query(query_objetos_sql_server.replace("REEMPLAZA_BBDD", bbdd), MiConexion)
            df_objetos_sql_server = df_temp_1 if ind == 0 else pd.concat([df_objetos_sql_server, df_temp_1])
            del df_temp_1


            if proceso_id == "PROCESO_03":

                df_temp_2 = pd.read_sql_query(query_objetos_sql_server_parametros.replace("REEMPLAZA_BBDD", bbdd), MiConexion)
                df_objetos_sql_server_parametros = df_temp_2 if ind == 0 else pd.concat([df_objetos_sql_server_parametros, df_temp_2])
                del df_temp_2

        df_objetos_sql_server.reset_index(drop = True, inplace = True)
        
        if proceso_id == "PROCESO_03":
            df_objetos_sql_server_parametros.reset_index(drop = True, inplace = True)



        #se crea la lista de diccionarios lista_dicc_objetos_sql_server con todos los objetos de las distintas bbdd seleccionadas
        lista_dicc_objetos_sql_server = []
        for ind in df_objetos_sql_server.index:

            bbdd = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("BBDD")]
            repositorio = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("REPOSITORIO")]
            tipo_objeto = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("TIPO_OBJETO")]
            nombre_objeto = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("NOMBRE_OBJETO")]
            objeto_long = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("OBJETO_LONG")]
            objeto_short = df_objetos_sql_server.iloc[ind, df_objetos_sql_server.columns.get_loc("OBJETO_SHORT")]


            df_codigo_original = None
            lista_codigo_limpiado = None
            string_parametros = None

            #se crea el df con el codigo del objeto usando la funcion func_sql_server_df_codigo_objeto
            df_codigo_original = func_sql_server_df_codigo_objeto("DF_CODIGO_ORIGINAL", tipo_objeto = tipo_objeto, MiConexion = MiConexion, bbdd = bbdd, objeto_long = objeto_long, 
                                                                                        cursor = cursor, objeto_short = objeto_short)


            #se limpia el df de codigo original usando la funcion func_sql_server_df_codigo_objeto (aplica solo para el diagnostico de dependencias)
            lista_codigo_limpiado = func_sql_server_df_codigo_objeto("LISTA_CODIGO_LIMPIO", df_codigo_original = df_codigo_original) if proceso_id == "PROCESO_03" else None


            #se crea el string_parametros (aplica solo para el diagnostico de dependencias y para las bbdd seleccionadas)
            if proceso_id == "PROCESO_03":


                df_parametros_temp = (df_objetos_sql_server_parametros.loc[(df_objetos_sql_server_parametros["BBDD"] == bbdd) & 
                                                                            (df_objetos_sql_server_parametros["REPOSITORIO"] == repositorio) & 
                                                                            (df_objetos_sql_server_parametros["NOMBRE_OBJETO"] == nombre_objeto) & 
                                                                            (df_objetos_sql_server_parametros["TIPO_OBJETO"] == tipo_objeto), ["PARAMETROS"]])
            
                df_parametros_temp.reset_index(drop = True, inplace = True)
                string_parametros = ", ".join([df_parametros_temp.iloc[ind, 0] for ind in df_parametros_temp.index]) if len(df_parametros_temp) != 0 else None
                del df_parametros_temp


            #se crea el diccionario que se agrega a lista_dicc_objetos_sql_server
            dicc_temp = {"SERVIDOR": sql_server_servidor
                        , "BBDD": bbdd
                        , "REPOSITORIO": repositorio
                        , "TIPO_OBJETO": tipo_objeto
                        , "NOMBRE_OBJETO": nombre_objeto
                        , "OBJETO_LONG": objeto_long
                        , "OBJETO_SHORT": objeto_short
                        , "PARAMETROS": string_parametros
                        , "DF_CODIGO_ORIGINAL": df_codigo_original
                        , "LISTA_CODIGO_LIMPIADO": lista_codigo_limpiado
                        }

            lista_dicc_objetos_sql_server.append(dicc_temp)

            del df_codigo_original
            del dicc_temp


        del df_objetos_sql_server


        #se cierra la conexion con el servidor seleccionado
        MiConexion.close()


        ######################################################################################################################################################
        #se almacena lista_dicc_objetos_sql_server en dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["LISTA_DICC_OBJETOS"]
        ######################################################################################################################################################

        mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["LISTA_DICC_OBJETOS"] = lista_dicc_objetos_sql_server if len(lista_dicc_objetos_sql_server) != 0 else None
        mod_gen.dicc_codigos_bbdd[opcion_bbdd]["SQL_SERVER"]["LISTA_BBDD_SERVIDOR"] = lista_sql_server_bbdd_servidor



    except Exception as Err:

        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos[proceso_id]["SQL_SERVER"]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd]

        dicc_errores_temp = {"TIPO_BBDD": "SQL_SERVER"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos[proceso_id]["SQL_SERVER"]["LISTA_DICC_ERRORES_IMPORTACION_" + opcion_bbdd] = lista_dicc_errores_migracion
        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso de import

    
    

#################################################################################################################################################################################
##                     RUTINA SQL SERVER - CONTROL DE VERSIONES
#################################################################################################################################################################################

def def_proceso_sql_server_2_control_versiones():
    #rutina que permite realizar los calculos necesarios para el control de versiones

    warnings.filterwarnings("ignore")

    try:

        #se unifican los objetos de las 2 bbdd
        lista_dicc_objetos_bbdd_1 = mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["LISTA_DICC_OBJETOS"]
        lista_dicc_objetos_bbdd_2 = mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["LISTA_DICC_OBJETOS"]


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
        indice_repositorio = 1
        indice_nombre_objeto = 2
        indice_esta_en_bbdd_1 = 3
        indice_esta_en_bbdd_2 = 4
        indice_df_codigo_1 = 5
        indice_df_codigo_2 = 6

        lista_objetos_conso_1 = [[dicc["TIPO_OBJETO"]
                                , dicc["REPOSITORIO"]
                                , dicc["NOMBRE_OBJETO"]
                                , 0
                                , 0
                                , None
                                , None] for dicc in lista_dicc_objetos_bbdd_1]

        lista_objetos_conso_2 = [[dicc["TIPO_OBJETO"]
                                , dicc["REPOSITORIO"]
                                , dicc["NOMBRE_OBJETO"]
                                , 0
                                , 0
                                , None
                                , None] for dicc in lista_dicc_objetos_bbdd_2]

        lista_objetos_conso = lista_objetos_conso_1 + lista_objetos_conso_2
        lista_objetos_conso = [sublista for i, sublista in enumerate(lista_objetos_conso) if sublista not in lista_objetos_conso[:i]]#se quitan los duplicados



        #indices en las sublistas de lista_objetos
        indice_tipo_objeto = 0
        indice_repositorio = 1
        indice_nombre_objeto = 2
        indice_esta_en_bbdd_1 = 3
        indice_esta_en_bbdd_2 = 4
        indice_df_codigo_1 = 5
        indice_df_codigo_2 = 6


        #se localizan que objetos estan en BBDD_01 + se asigna el df_codigo
        for ind, item in enumerate(lista_objetos_conso):

            tipo_objeto_conso = item[indice_tipo_objeto]
            repositorio_conso = item[indice_repositorio]
            nombre_objeto_conso = item[indice_nombre_objeto]

            #BBDD_01
            if isinstance(lista_dicc_objetos_bbdd_1, list):
                for dicc in lista_dicc_objetos_bbdd_1:

                    tipo_objeto_seek = dicc["TIPO_OBJETO"]
                    repositorio_seek = dicc["REPOSITORIO"]
                    nombre_objeto_seek = dicc["NOMBRE_OBJETO"]

                    if tipo_objeto_conso == tipo_objeto_seek and repositorio_conso == repositorio_seek and nombre_objeto_conso == nombre_objeto_seek:
                        lista_objetos_conso[ind][indice_esta_en_bbdd_1] = 1
                        lista_objetos_conso[ind][indice_df_codigo_1] = dicc["DF_CODIGO_ORIGINAL"]

                        break

            
            #BBDD_02
            if isinstance(lista_dicc_objetos_bbdd_2, list):
                for dicc in lista_dicc_objetos_bbdd_2:

                    tipo_objeto_seek = dicc["TIPO_OBJETO"]
                    repositorio_seek = dicc["REPOSITORIO"]
                    nombre_objeto_seek = dicc["NOMBRE_OBJETO"]

                    if tipo_objeto_conso == tipo_objeto_seek and repositorio_conso == repositorio_seek and nombre_objeto_conso == nombre_objeto_seek:
                        lista_objetos_conso[ind][indice_esta_en_bbdd_2] = 1
                        lista_objetos_conso[ind][indice_df_codigo_2] = dicc["DF_CODIGO_ORIGINAL"]

                        break


        del lista_dicc_objetos_bbdd_1
        del lista_dicc_objetos_bbdd_2



        #se calcula el control de versiones
        lista_control_versiones_sql_server_tablas = []
        lista_control_versiones_sql_server_views = []
        lista_control_versiones_sql_server_stored_procedures = []
        lista_control_versiones_sql_server_funciones = []

        for tipo_objeto_conso, repositorio_conso, nombre_objeto_conso, en_bbdd_1_conso, en_bbdd_2_conso, df_codigo_1_conso, df_codigo_2_conso in lista_objetos_conso:

            if en_bbdd_1_conso == 1 and en_bbdd_2_conso == 1:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["YA_EXISTE"]

            elif en_bbdd_1_conso == 1 and en_bbdd_2_conso == 0:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_01"]

            elif en_bbdd_1_conso == 0 and en_bbdd_2_conso == 1:
                check_objeto = mod_gen.dicc_control_versiones_tipo_concepto["SOLO_EN_BBDD_02"]



            dicc_temp = mod_gen.func_dicc_control_versiones(tipo_bbdd = "SQL_SERVER"
                                                        , check_objeto = check_objeto
                                                        , tipo_objeto = tipo_objeto_conso
                                                        , repositorio = repositorio_conso
                                                        , nombre_objeto = nombre_objeto_conso
                                                        , df_codigo_bbdd_1 = df_codigo_1_conso
                                                        , df_codigo_bbdd_2 = df_codigo_2_conso
                                                        )
            

            if isinstance(dicc_temp, dict):
                
                if tipo_objeto_conso == "TABLAS":
                    lista_control_versiones_sql_server_tablas.append(dicc_temp)

                if tipo_objeto_conso == "VIEWS":
                    lista_control_versiones_sql_server_views.append(dicc_temp)

                if tipo_objeto_conso == "STORED_PROCEDURES":
                    lista_control_versiones_sql_server_stored_procedures.append(dicc_temp)

                if tipo_objeto_conso == "FUNCIONES":
                    lista_control_versiones_sql_server_funciones.append(dicc_temp)

            del dicc_temp


        #se almacenan lo referente a Access en la lista en dicc_control_versiones_tipo_objeto
        mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"]["TABLAS"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_sql_server_tablas if len(lista_control_versiones_sql_server_tablas) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"]["VIEWS"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_sql_server_views if len(lista_control_versiones_sql_server_views) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"]["STORED_PROCEDURES"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_sql_server_stored_procedures if len(lista_control_versiones_sql_server_stored_procedures) != 0 else None
        mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"]["FUNCIONES"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_sql_server_funciones if len(lista_control_versiones_sql_server_funciones) != 0 else None

        del lista_control_versiones_sql_server_tablas
        del lista_control_versiones_sql_server_views
        del lista_control_versiones_sql_server_stored_procedures
        del lista_control_versiones_sql_server_funciones



        #se almacena la lista en dicc_control_versiones_tipo_objeto en TODOS (es para el combobox opcion TODOS)
        lista_tipo_objetos = list(mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"].keys())

        lista_control_versiones_todo = []
        for tipo_objeto in lista_tipo_objetos:
            if tipo_objeto != "TODOS":

                try:
                    for dicc in mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"][tipo_objeto]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"]:
                        lista_control_versiones_todo.append(dicc)

                except:#puede haber tipo de objetos que no tengan cambios
                    pass
        
        mod_gen.dicc_control_versiones_tipo_objeto["SQL_SERVER"]["TIPO_OBJETO"]["TODOS"]["LISTA_DICC_OBJETOS_CONTROL_VERSIONES"] = lista_control_versiones_todo if len(lista_control_versiones_todo) != 0 else None
        del lista_control_versiones_todo


        #se vacia parte de dicc_codigos_bbdd para liberar memoria
        mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["CONTROL_VERSIONES_LISTA_DICC_OBJETOS"] = None
        mod_gen.dicc_codigos_bbdd["BBDD_02"]["SQL_SERVER"]["CONTROL_VERSIONES_LISTA_DICC_OBJETOS"] = None



    except Exception as Err:
        
        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos["PROCESO_01"]["SQL_SERVER"]["LISTA_DICC_ERRORES_CALCULO"]

        dicc_errores_temp = {"TIPO_BBDD": "SQL_SERVER"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos["PROCESO_01"]["SQL_SERVER"]["LISTA_DICC_ERRORES_CALCULO"] = lista_dicc_errores_migracion
        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso





#################################################################################################################################################################################
##                     RUTINA SQL SERVER - DIAGNOSTICO
#################################################################################################################################################################################

def def_proceso_sql_server_2_diagnostico(opcion, servidor_selecc, lista_bbdd_selecc, ruta_destino):
    #rutina que permite realizar los calculos necesarios para el diagnostico y su exportacion a excel

    warnings.filterwarnings("ignore")


    try:
        
        #se recuperan los datos del diccionario dicc_codigos_bbdd para la key BBDD_01 (por defecto es la bbdd que se usa en el app
        #para el diagnostico de dependencias)
        lista_dicc_objetos_sql_server = mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["LISTA_DICC_OBJETOS"]
        lista_sql_server_bbdd_servidor = mod_gen.dicc_codigos_bbdd["BBDD_01"]["SQL_SERVER"]["LISTA_BBDD_SERVIDOR"]


        ##################################################################################
        #opcion --> "Realizar diagnostico"
        ##################################################################################
        if opcion == lista_GUI_diagnostico_combobox_sql_server[0]:


            #se crean las listas sobre las cuales realizar el calculo del diagnostico de dependencias
            lista_sql_server_objetos = []
            lista_sql_server_objetos_dependencias = []

            for dicc in lista_dicc_objetos_sql_server:

                bbdd = dicc["BBDD"]
                repositorio = dicc["REPOSITORIO"]
                tipo_objeto = dicc["TIPO_OBJETO"]
                nombre_objeto = dicc["NOMBRE_OBJETO"]
                objeto_long_mayusc = dicc["OBJETO_LONG"].upper()
                objeto_short_mayusc = dicc["OBJETO_SHORT"].upper()
                lista_codigo_limpiado = dicc["LISTA_CODIGO_LIMPIADO"]


                #se crean las listas para poder realizar el calculo del diagnostico de dependencias sobre la totalidad de objetos
                #de todas las bbdd que componen el servidor seleccionado
                lista_sql_server_objetos.append([bbdd
                                                , repositorio
                                                , tipo_objeto
                                                , nombre_objeto
                                                , objeto_long_mayusc
                                                , objeto_short_mayusc
                                                , lista_codigo_limpiado
                                                , []    #sirve para almacenar la lista de objetos donde se usa el objeto (se establece como lista vacia de inicio, se informa mas adelante)
                                                ])

                lista_sql_server_objetos_dependencias.append([bbdd
                                                            , repositorio
                                                            , tipo_objeto
                                                            , nombre_objeto
                                                            , objeto_short_mayusc
                                                            ])


                
            #indices de las sublistas de lista_sql_server_objetos
            indice_bbdd = 0
            indice_repositorio = 1
            indice_tipo_objeto = 2
            indice_nombre_objeto = 3
            indice_objeto_long = 4
            indice_objeto_short = 5
            indice_lista_codigo = 6
            indice_lista_dependencias = 7

            indice_dependencias_bbdd = 0
            indice_dependencias_repositorio = 1
            indice_dependencias_tipo_objeto = 2
            indice_dependencias_nombre_objeto = 3
            indice_dependencias_objeto_short = 4


            ######################################################################################
            #se calcula el diagnostico de dependencias
            ######################################################################################

            for indice_lista, item_1 in enumerate(lista_sql_server_objetos):

                bbdd = item_1[indice_bbdd]
                repositorio = item_1[indice_repositorio]
                tipo_objeto = item_1[indice_tipo_objeto]
                nombre_objeto = item_1[indice_nombre_objeto]
                objeto_long_mayusc = item_1[indice_objeto_long]
                objeto_short_mayusc = item_1[indice_objeto_short]
                lista_codigo = item_1[indice_lista_codigo]


                lista_dependencias_objeto = []
                for item_2 in lista_sql_server_objetos_dependencias:

                    bbdd_dependencias = item_2[indice_dependencias_bbdd]
                    repositorio_dependencias = item_2[indice_dependencias_repositorio]
                    tipo_objeto_dependencias = item_2[indice_dependencias_tipo_objeto]
                    nombre_objeto_dependencias = item_2[indice_dependencias_nombre_objeto]
                    objeto_short_mayusc_dependencias = item_2[indice_dependencias_objeto_short]


                    #se excluye de la busqueda el propio objeto para evitar localizar dependencias del objeto con su propio codigo
                    if not (bbdd == bbdd_dependencias and repositorio == repositorio_dependencias and 
                            tipo_objeto == tipo_objeto_dependencias and nombre_objeto == nombre_objeto_dependencias):
                        

                        for linea_codigo in lista_codigo:

                            #se excluye de la busqueda la sentencia de declaracion de la rutina / funcion
                            if objeto_long_mayusc not in linea_codigo and objeto_short_mayusc not in linea_codigo:


                                #se extralimita el macheo cuando se localiza que el objeto short esta en la linea de codigo
                                if objeto_short_mayusc_dependencias in linea_codigo:


                                    #se chequea si el objeto_short_mayusc_dependencias es una palabra sola (no incluida en otra mas larga, que podria ser otro objeto)
                                    check_si_objeto_short_es_palabra_sola = fr'(?<![\w_.-]){re.escape(objeto_short_mayusc_dependencias)}(?![\w_.-])'
                                    check_si_objeto_short_es_palabra_sola = bool(re.search(check_si_objeto_short_es_palabra_sola, linea_codigo))

                                    check_objeto_short = 1 if check_si_objeto_short_es_palabra_sola == True else 0

                                    #si se localiza objeto_short_mayusc_dependencias como palabra sola
                                    #en caso de macheo se agregan los datos de la iteracion sobre lista_sql_server_objetos_dependencias
                                    #pero asignandoles la bbdd donde esta guardado el script
                                    if check_objeto_short == 1:

                                        lista_temp = [bbdd, repositorio_dependencias, tipo_objeto_dependencias, nombre_objeto_dependencias]

                                        if lista_temp not in lista_dependencias_objeto:
                                            lista_dependencias_objeto.append(lista_temp)

                                    #si NO se localiza objeto_short_mayusc_dependencias como palabra sola
                                    #se pasa, mediante bucle sobre las bbdd del servidor, a reconstruir el objeto long
                                    #y se comprueba en cada iteracion si aparece como palabra sola (no incluida en otra mas larga, que podria ser otro objeto)
                                    elif check_objeto_short == 0:

                                        for bbdd_servidor in lista_sql_server_bbdd_servidor:
                                            objeto_long_mayusc_bbdd_servidor = bbdd_servidor.upper() + "." + objeto_short_mayusc_dependencias

                                            #se extralimita el macheo cuando se localiza que el objeto short esta en la linea de codigo
                                            if objeto_long_mayusc_bbdd_servidor in linea_codigo:

                                                check_si_objeto_long_es_palabra_sola = fr'(?<![\w_.-]){re.escape(objeto_long_mayusc_bbdd_servidor)}(?![\w_.-])'
                                                check_si_objeto_long_es_palabra_sola = bool(re.search(check_si_objeto_long_es_palabra_sola, linea_codigo))

                                                check_objeto_long = 1 if check_si_objeto_long_es_palabra_sola == True else 0


                                                #en caso de macheo se agregan los datos de la iteracion sobre lista_sql_server_objetos_dependencias
                                                #asignandoles la bbdd del servidor localizada
                                                if check_objeto_long == 1:

                                                    lista_temp = [bbdd_servidor, repositorio_dependencias, tipo_objeto_dependencias, nombre_objeto_dependencias]

                                                    if lista_temp not in lista_dependencias_objeto:
                                                        lista_dependencias_objeto.append(lista_temp)


                #se actualiza la lista de dependencias en lista_sql_server_objetos
                lista_sql_server_objetos[indice_lista][indice_lista_dependencias] = lista_dependencias_objeto if len(lista_dependencias_objeto) != 0 else None


            #se crean las listas de dependencias y sin dependencias
            lista_sql_server_para_df_dependencias_objetos = []
            lista_sql_server_para_df_sin_dependencias_objetos = []

            for indice_lista, item in enumerate(lista_sql_server_objetos):

                bbdd = item[indice_bbdd]
                repositorio = item[indice_repositorio]
                tipo_objeto = item[indice_tipo_objeto]
                nombre_objeto = item[indice_nombre_objeto]
                lista_dependencias_objeto = item[indice_lista_dependencias]

                if bbdd in lista_bbdd_selecc:

                    if isinstance(lista_dependencias_objeto, list):
                        for bbdd_dependencias, repositorio_dependencias, tipo_objeto_dependencias, nombre_objeto_dependencias in lista_dependencias_objeto:

                            if bbdd_dependencias in lista_bbdd_selecc:
                                lista_sql_server_para_df_dependencias_objetos.append([bbdd_dependencias
                                                                                    , repositorio_dependencias
                                                                                    , tipo_objeto_dependencias
                                                                                    , nombre_objeto_dependencias
                                                                                    , bbdd
                                                                                    , repositorio
                                                                                    , tipo_objeto
                                                                                    , nombre_objeto])

                    else:
                        lista_sql_server_para_df_sin_dependencias_objetos.append([bbdd
                                                                                , repositorio
                                                                                , tipo_objeto
                                                                                , nombre_objeto])            

    
            #se exporta a excel (se crea el diccionario temporal  para poder realizar la exportacion por bucle)
            #se crea la lista de bbdd selecionadas para su exportacion a excel

            lista_sql_server_para_df_lista_bbdd = [[servidor_selecc, bbdd] for bbdd in lista_bbdd_selecc]

            lista_sql_server_para_df_listado_objetos = [[dicc["BBDD"], dicc["REPOSITORIO"], dicc["TIPO_OBJETO"], dicc["NOMBRE_OBJETO"], dicc["PARAMETROS"]] 
                                                        for dicc in lista_dicc_objetos_sql_server if dicc["BBDD"] in lista_bbdd_selecc]

            dicc_objetos_sql_server = {"BBDD_SELECCIONADAS":
                                                            {"HOJA_EXCEL": "BBDD SELECCIONADAS"
                                                            , "LISTA_PARA_DF": lista_sql_server_para_df_lista_bbdd
                                                            , "LISTA_HEADERS": lista_headers_sql_server_diagnostico_df_bbdd_selecc
                                                            }

                                        , "LISTADO_OBJETOS":
                                                            {"HOJA_EXCEL": "LISTADO"
                                                            , "LISTA_PARA_DF": lista_sql_server_para_df_listado_objetos
                                                            , "LISTA_HEADERS": lista_headers_sql_server_diagnostico_df_listado_objetos
                                                            }

                                        , "DEPENDENCIAS":
                                                            {"HOJA_EXCEL": "DEPENDENCIAS"
                                                            , "LISTA_PARA_DF": lista_sql_server_para_df_dependencias_objetos
                                                            , "LISTA_HEADERS": lista_headers_sql_server_diagnostico_df_dependencias
                                                            }

                                        , "SIN_DEPENDENCIAS":
                                                            {"HOJA_EXCEL": "SIN DEPENDENCIAS"
                                                            , "LISTA_PARA_DF": lista_sql_server_para_df_sin_dependencias_objetos
                                                            , "LISTA_HEADERS": lista_headers_sql_server_diagnostico_df_sin_dependencias
                                                            }
                                        }



            now = str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
            saveas = str(ruta_destino) + r"\DEPENDENCIAS_SQL_SERVER_"  + now + ".xlsb"

            shutil.copyfile(mod_gen.ruta_plantilla_diagnostico_sql_server_xls, saveas)

            app = xw.App(visible = False)
            wb = app.books.open(saveas, update_links = False)


            for key in dicc_objetos_sql_server.keys():

                hoja_excel = dicc_objetos_sql_server[key]["HOJA_EXCEL"]
                lista_para_df = dicc_objetos_sql_server[key]["LISTA_PARA_DF"]
                lista_headers = dicc_objetos_sql_server[key]["LISTA_HEADERS"]
                

                if len(lista_para_df) != 0:

                    df_temp = pd.DataFrame(lista_para_df, columns = lista_headers)
                    df_temp = df_temp[lista_headers].sort_values(lista_headers, ascending = [True for i in lista_headers])
                    df_temp.reset_index(drop = True, inplace = True)

                    ws = wb.sheets[hoja_excel]
                    ws.range("A2:FF65000").clear_contents()
                    ws["A2"].options(pd.DataFrame, header = 0, index = False, expand = "table").value = df_temp


            #se guarda en el excel, se cierra y se re-abre en 1er plano        
            wb.save(saveas)
            wb.close()
            app = xw.App(visible = True)
            app.quit()

            wb = xw.Book(saveas, update_links = False)


            del dicc_objetos_sql_server
            del lista_sql_server_para_df_listado_objetos
            del lista_sql_server_para_df_dependencias_objetos
            del lista_sql_server_para_df_sin_dependencias_objetos



        ##################################################################################
        #opcion --> "Descargar códigos T-SQL"
        ##################################################################################
        elif opcion == lista_GUI_diagnostico_combobox_sql_server[1]:

            now = "CODIGOS_T_SQL_" + servidor_selecc.replace(r"\"", "_") + str(dt.datetime.now()).replace("-", "").replace(" ", "_").replace(":", "")[0:15]
            ruta_export = ruta_destino + r"\%s" % now
            os.makedirs(ruta_export, exist_ok = True)
    

            #se crean las subcarpetas por bbdd
            for bbdd in lista_bbdd_selecc:
                ruta_export_bbdd = ruta_export + r"\%s" % bbdd
                os.makedirs(ruta_export_bbdd, exist_ok = True)

                lista_esquemas = [dicc["REPOSITORIO"] for dicc in lista_dicc_objetos_sql_server if dicc["BBDD"] == bbdd]
                lista_esquemas = list(dict.fromkeys(lista_esquemas))


                #se crean las subcarpetas por esquemas
                for esquema in lista_esquemas:
                    ruta_export_esquema = ruta_export_bbdd + r"\%s" % esquema
                    os.makedirs(ruta_export_esquema, exist_ok = True)

                    lista_tipo_objetos = [dicc["TIPO_OBJETO"] for dicc in lista_dicc_objetos_sql_server if dicc["BBDD"] == bbdd and dicc["REPOSITORIO"] == esquema]
                    lista_tipo_objetos = list(dict.fromkeys(lista_tipo_objetos))


                    #se crean las subcarpetas por tipo_objeto
                    for tipo_objeto in lista_tipo_objetos:
                        ruta_export_tipo_objeto = ruta_export_esquema + r"\%s" % tipo_objeto
                        os.makedirs(ruta_export_tipo_objeto, exist_ok = True)

                        lista_objetos = [[dicc["NOMBRE_OBJETO"], dicc["DF_CODIGO_ORIGINAL"]]
                                        for dicc in lista_dicc_objetos_sql_server if dicc["BBDD"] == bbdd and dicc["REPOSITORIO"] == esquema and dicc["TIPO_OBJETO"] == tipo_objeto]                           


                        #se crean los ficheros .sql
                        for nombre_objeto, df_codigo in lista_objetos:
                            nombre_objeto_corr = nombre_objeto + ".sql"
                            nombre_objeto_corr = ruta_export_tipo_objeto + r"\%s" % nombre_objeto_corr

                            with open(nombre_objeto_corr, 'w') as fich:

                                #se concatena el encabezado de codigo con el df de codigo (y se reemplaza la bbdd en el 1er registro del df resultante)
                                df_codigo_corr = pd.concat([df_encabezado_descarga_codigo_objetos, df_codigo])
                                df_codigo_corr.reset_index(drop = True, inplace = True)

                                df_codigo_corr.iloc[0, 0] = df_codigo_corr.iloc[0, 0].replace("REEMPLAZA_BBDD", bbdd)

                                for ind in df_codigo_corr.index:
                                    fich.write(df_codigo_corr.iloc[ind, df_codigo_corr.columns.get_loc("CODIGO")] + "\n")


        del lista_dicc_objetos_sql_server


    except Exception as Err:
        
        traceback_error = traceback.extract_tb(Err.__traceback__)
        modulo_python = os.path.basename(traceback_error[0].filename)
        rutina_python = traceback_error[0].name
        linea_error = traceback_error[0].lineno

        lista_dicc_errores_migracion = mod_gen.dicc_errores_procesos["PROCESO_03"]["SQL_SERVER"]["LISTA_DICC_ERRORES_CALCULO"]

        dicc_errores_temp = {"TIPO_BBDD": "SQL_SERVER"
                            , "MODULO_PYTHON": modulo_python
                            , "RUTINA_PYTHON": rutina_python
                            , "LINEA_ERROR": linea_error
                            , "ERRORES": str(Err)
                            }

        if isinstance(lista_dicc_errores_migracion, list):
            lista_dicc_errores_migracion.append(dicc_errores_temp)
        else:
            lista_dicc_errores_migracion = [dicc_errores_temp]

        mod_gen.dicc_errores_procesos["PROCESO_03"]["SQL_SERVER"]["LISTA_DICC_ERRORES_CALCULO"] = lista_dicc_errores_migracion

        del dicc_errores_temp
        del lista_dicc_errores_migracion
        pass #es pass para que se localicen todos los posibles errores en el proceso


