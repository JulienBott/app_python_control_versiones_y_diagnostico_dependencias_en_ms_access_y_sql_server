
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar control de versiones de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos). Adicionalmente, ofrece un diagnóstico de bases de datos MS Access, generando un informe en Excel que detalla dónde se usan las tablas, rutinas VBA y variables públicas VBA e identifica los objetos que no son utilizados en ninguna rutina VBA.

El aplicativo es una alternativa humilde a herramientas de control de versiones que requieren licencias de pago, especialmente en entornos empresariales.

  __Funcionamiento del Control de Versiones__
      
1. Trás seleccionar 2 bbdd MS Access distintas y/o 2 bases de datos SQL Server, el app distingue entre 4 tipos de objetos y cada uno se distingue en 3 tipos de concepto donde se han localizado cambios en los scripts (VBA y/o T-SQL):
    
      * __Objetos__:
          * tablas MS Access (pueden ser tablas locales, vinculos ODBC o vinculos hacia otras fuentes externas)
        
          * rutinas MS Access VBA de los distintos módulos de la base de datos
          
          * variables públicas MS Access VBA de los distintos módulos de la base de datos
        
          * objetos SQL Server (tablas, stored procedures o views)
            
      * __Conceptos__
          * Objetos ya existentes en las 2 BBDD's
          * En BBDD_01 pero no en BBDD_02
          * En BBDD_02 pero no en BBDD_01
      
2. Tras seleccionar el tipo de objeto y el tipo de concepto, se localizan los objetos asociados y al seleccionar cada uno se muestra en la misma pantalla los scripts de las 2 bases de datos marcando de color verde aquellas lineas de código que son nuevas o ya existentes pero con cambios.
      
3. El app permite también realizar el MERGE de una base de datos a otra de forma manual y automatica:
   * __Manual__: para los objetos incluidos en "Objetos ya existentes en las 2 BBDD's" se pueden traspasar lineas de código de un script a otro, revertir cambios o guardar
   * __Automatica__: para los objetos que NO figuran en la base de datos donde se hace el merge pero que no figuren en esta última
      
5. Por último, al validar el MERGE se crea en una ruta que el usuario indique documentación del proceso de MERGE indicando que objetos afecta y para cada uno que se ha cambiado.


## __DEMO__

  A falta de incorporar un video de demo, he colgado en este repositorio un powerpoint con pantallazos y 
  explicaciones de lo que se ve en pantalla.

 
## __REQUISITOS FUNCIONAMIENTO DEL APP__

__MS Access__:
  * deshabilitar los password VBA de los MS Access que se vayan a usar
    
  * deshabilitar la macro AutoExec si existiese (cambiadole el nombre de forma temporal por ejemplo)

__SQL Server__: 
  * La conexión esta configurada por Windows Authentication.
    Para cambiarla, ir a la variable __conn_str_sql_server__ (fila 39) del módulo APP_CONTROL_VERSIONES_2_GENERAL_POO


## __REQUISITOS LIBRERIAS PYTHON__

    thread6                   0.2.0
    pandas                    1.5.0
    numpy                     1.22.0
    pyodbc                    4.0.32
    xlwings                   0.27.15
    pyinstaller               4.7

## __ORGANIZACIÓN DEL PROYECTO__

El proyecto se organiza en 3 módulos .py (ver carpeta codigo):

  * __APP_CONTROL_VERSIONES_1_GUI__: contiene la interfaz de usuario con sus distintos widgets y rutinas asociadas

  * __APP_CONTROL_VERSIONES_2_GENERAL_POO__: contiene variables globales que se usan en los 2 otros módulos y también clases propias creadas para el proyecto

  * __APP_CONTROL_VERSIONES_3_BACK_END__: contiene todas las rutinas back-end del proyecto

El app necesita 3 templates para su correcto funcionamiento que se han de integrar en el ejecutable cuando se compile (ver carpeta templates):
  * __ico_app__: fichero .ico
  * __ruta_plantilla_diagnostico_xls__: plantilla excel para el diagnostico de una base de datos MS Access
  * __ruta_plantilla_control_versiones_xls__: plantilla excel para poder descargar todos los objetos con cambios
      
## FASE DEL PROYECTO (actualizado a 2024-10-18)

Los puntos 3 y 4 que figuran en la descripción estan en desarrollo (espero tenerlos a lo largo de la semana que viene)

Tengo pendiente limpiar código encapsulando todavia más en clases para aligerar el código

Tengo también pendiente realizar un desarollo funcional del app pero solo cuando lo de por acabado a nivel de código














