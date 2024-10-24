
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar control de versiones de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos). 

Adicionalmente, ofrece un diagnóstico de dependencias de objetos de bases de datos MS Access o de un servidor SQL Server:

El aplicativo es una alternativa humilde a herramientas de control de versiones que requieren licencias de pago, especialmente en entornos empresariales.

  __Funcionamiento del Control de Versiones__
      
1. Trás seleccionar 2 bbdd MS Access distintas y/o 2 bases de datos SQL Server, el app distingue entre 4 tipos de objetos y cada uno se distingue en 3 tipos de concepto donde se han localizado cambios en los scripts (VBA y/o T-SQL):
    
      * __Objetos__:
          * tablas MS Access (pueden ser tablas locales, vinculos ODBC o vinculos hacia otras fuentes externas)
        
          * rutinas MS Access VBA de los distintos módulos de la base de datos
          
          * variables públicas MS Access VBA de los distintos módulos de la base de datos
        
          * objetos SQL Server (tablas, stored procedures, funciones o views)
            
      * __Conceptos__
          * Objetos ya existentes en las 2 BBDD's
          * En BBDD_01 pero no en BBDD_02
          * En BBDD_02 pero no en BBDD_01
      
2. Tras seleccionar el tipo de objeto y el tipo de concepto, se localizan los objetos asociados y al seleccionar cada uno se muestra en la misma pantalla los scripts de las 2 bases de datos marcando de color aquellas lineas de código que son nuevas o ya existentes pero con cambios.

    En la misma interfaz esta habilitada la posibilidad de hacer MERGE de un script a otro:
    * migrando por completo los scripts que no figuran en la bbdd donde se quiere hacer el merge
    * quitando por completo los scripts que no interesa que sigan en la bbdd donde se quiere hacer el merge
    * desplazar lineas de código de un script a otro (marcandolas de color naranja para guardar trazabilidad de los cambios)
    * revertir los cambios en caso de cometer errores
      
3. El app permite aplicar los cambios en las BBDD fisicas sean MS Access o SQL Server. Al pulsar en la opción aparece otra interfaz que permite seleccionar que objetos se van a migrar.
   
   En una ruta que el usuario indique se genera documentación del proceso de MERGE indicando que objetos afecta y para cada uno que se ha cambiado.
   
   * __MS ACCESS__: los módulos VBA se modifican con los cambios aportados (los cambios que se hagan en los formularios como añadir widgets es una tarea que no hace el app por lo que esta parte es manual)
     
   * __SQL SERVER__: se crean en la bbdd MERGE las tablas, vistas, funciones y views (se crean nuevos esquemas en caso que la bbdd NO_MERGE tenga alguno que la bbdd MERGE no tenga)
  
5. Por último, el app permite acceder a un historico de cambios por fecha y usuario (es un sistema en SQLite)
 
 
 
  __Funcionamiento del Diagnostico de dependencias de objetos__
  
  * __MS Access__: genera un informe en Excel que detalla dónde se usan las tablas, rutinas VBA y variables públicas VBA e identifica los objetos que no son utilizados en ninguna rutina VBA.

  * __SQL Server__: genera el mismo tipo de informe en Excel con la opción de poder seleccionar varias bbdd para localizar sus dependencias. Incluye también la posibilidad de descargar todos los scripts de objetos en ficheros .sql en la ruta que indique el usuario.


## __DEMO__

  A falta de incorporar un video de demo, he colgado en este repositorio un powerpoint con pantallazos y 
  explicaciones de lo que se ve en pantalla.

 
## __REQUISITOS FUNCIONAMIENTO DEL APP__

__MS Access__:
  * deshabilitar los password VBA de los MS Access que se vayan a usar
  * deshabilitar la macro AutoExec si existiese (cambiadole el nombre de forma temporal por ejemplo)

__SQL Server__: 
  * La conexión esta configurada por Windows Authentication.
    Para cambiarla, ir a la variable __conn_str_sql_server__ (fila 39) del módulo __APP_CONTROL_VERSIONES_2_GENERAL_POO__

  * Para cambiar la lista de los servidores ir al módulo __APP_CONTROL_VERSIONES_2_GENERAL_POO__ a la variable __lista_GUI_sql_server_servidor__ (fila 41)


## __REQUISITOS LIBRERIAS PYTHON__

El app se ha desarrollado y probado en entorno Windows (10) usando la versión 3.9.5 de Python

Librerias que requieren instalación (pip install):

    numpy                     1.22.0
    pandas                    1.5.0
    pyinstaller               4.7
    pyodbc                    4.0.32
    thread6                   0.2.0
    xlwings                   0.27.15

Librerias nativas Python:

    datetime
    difflib
    os
    pathlib
    re
    sqlite3
    sys
    tkinter
    warnings
    win32com.client

## __ORGANIZACIÓN DEL PROYECTO__

El proyecto se organiza en 4 módulos .py (ver carpeta codigo):

  * __APP_CONTROL_VERSIONES_1_GUI__: contiene la interfaz de usuario con sus distintos widgets y rutinas asociadas

  * __APP_CONTROL_VERSIONES_2_GENERAL_POO__: contiene variables globales que se usan en los 2 otros módulos y también clases propias creadas para el proyecto

  * __APP_CONTROL_VERSIONES_3_BACK_END__: contiene todas las rutinas back-end del proyecto

  * __APP_CONTROL_VERSIONES_4_DIAGNOSTICO_SQL_SERVER__: contiene todo lo referente al diagnostico de un servidor SQL Server

El app necesita 3 templates para su correcto funcionamiento que se han de integrar en el ejecutable cuando se compile (ver carpeta templates):
  * __ico_app__: fichero .ico
  * __ruta_plantilla_control_versiones_xls__: plantilla excel para poder descargar todos los objetos con cambios
  * __ruta_plantilla_diagnostico_access_xls__: plantilla excel para el diagnostico de una base de datos MS Access
  * __ruta_plantilla_diagnostico_sql_server_xls__: plantilla excel para el diagnostico de un servidor SQL Server

## CONSIDERACIONES SOBRE LA OPCIÓN DIAGNOSTICO SERVIDOR SQL SERVER

En el app existe la posibilidad de realizar un diagnostico de un servidor SQL Server.
Inicialmente esto iba a ser otro app pero opte por incorporarlo también aqui puesto que ya ofrezco la posibilidad de hacer lo mismo con bbdd MS Access.

Por si interesa separarlo de este app para montar uno propio, he aislado el código en un unico módulo __APP_CONTROL_VERSIONES_4_DIAGNOSTICO_SQL_SERVER__
sin dependencias de variables globales como la connecting string puesto que la declaro de nuevo en este módulo.

Al final del script he encapsulado una mini interfaz despues de la sentencia __if __name__ == "__main__":__por si interesa no compilarlo en ejecutable y ejecutarlo desde la consola.
      
## FASE DEL PROYECTO (actualizado a 2024-10-18)

Tengo pendiente limpiar código encapsulando todavia más en clases para aligerar el código

Tengo también pendiente realizar un desarollo funcional del app pero solo cuando lo de por acabado a nivel de código

Tengo pendiente decidir si incluir un diagnostico sobre base de datos SQL Server al igual que el comento en el 1er parrafo. Empece a desarrollarlo en otro app anterior a este por lo que no tengo claro si agregarlo al app de control de versiones que presento aqui o al anterior.















