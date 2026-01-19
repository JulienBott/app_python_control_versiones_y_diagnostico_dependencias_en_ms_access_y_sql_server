
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar __control de versiones__ de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos).

En ambos entornos de bases de datos:

  * localiza para los mismos scripts presentes en cada base de dato las lineas de código donde hay diferencias, las marca de un color.
    Marca tambien de otro color los caracteres que han cambiado. Tambien localiza los scripts que tan solo estan en una de las dos bases de datos.
    
  * el usuario tiene la posibilidad de realizar de forma sencilla y agil migraciones de lineas de código de un script de una base datos a otra.
    
  * el usuario tambien puede ejecutar el merge en base de datos fisica, documentar el proceso y acceder a logs en caso de errores de migración que detallen el porque de los fallos, también genera logs para los objetos migrados correctamente.

Adicionalmente, ofrece un __diagnóstico de dependencias de objetos__ de bases de datos MS ACCESS o de una o varias bases de datos de un mismo servidor SQL SERVER. Permite para cada objeto localizar en que scripts de otros objetos se usan (excluye de la busqueda de dependencias los comentarios dentro del código para focalizarse solo en código activo). Asimismo, localiza aquellos objetos que no dependen de ningún otro. El output final es un fichero Excel para que sea más agil y flexible para el usuario realizar sus analisis. En el caso de SQL Server, da la posibilidad tambien de descargar todo los códigos de objetos de bases de datos en ficheros .sql.

Para __MS ACCESS__, los tipos de objeto se han limitado a las tablas, vinculos ODBC, vinculos hacia otras fuentes externas, variables públicas VBA y rutinas / funciones VBA. 

Para __SQL SERVER__, los objetos se han limitado a tablas, views, funciones y stored procedures. 

En el diseño funcional, colgado tambien en el repositorio GitHub, en el apartado "Limitaciones del app", se explican los pasos a seguir en caso de querer agregar más tipos de objeto.

A falta de un video de demostración, ir a la guia de usuario ubicada en la carpeta documentacion_otra para poder ver el proposito del aplicativo.

## __CONTENIDO DEL REPOSITORIO GITHUB__:

Nada más acceder al repositorio, se encuentra el README que estas leyendo ahora mismo acompañado de un contrato MIT Licence donde autorizo cualquier tipo de uso del app y de su código asociado sea a nivel particular o empresarial siempre y cuando se me reconozca la autoria original del app. Dicho contrato de MIT Licence tiene clausulas añadidas.

El resto del repositorio se divide por subcarpetas:

* __codigo__

  Contiene los 5 módulos de código Python:

  * __APP_1_GUI__
  * __APP_2_GUI_UTILS__
  * __APP_3_GENERAL__
  * __APP_4_BACK_END_MS_ACCESS__
  * __APP_5_BACK_END_SQL_SERVER__
 
* __documentacion_herencia_clases_tkinter__

  Contiene los 3 documentos correspondientes a la v1.0 de mi otro proyecto Github (tkinter_utils / v1.0):

  https://github.com/JulienBott/python_tools_modulares.git


  * __tkinter_utils_v1_0__: fichero .py que contiene el sistema que se ha de incorporar en los proyectos Python donde se quiera usar.
  * __EJEMPLO_USO_tkinter_utils_v1_0__: fichero .py con los ejemplos que se documentan en el manual que se comenta a continuación.
  * __MANUAL_tkinter_utils_v1_0__: fichero pdf que explica el sistema y lo ilustra con ejemplos documentados de como implementarlo en otros proyectos.

  Se adjuntan asismismo 2 archivos para que funcionen los ejemplos:
  * __ico_tapar_pluma_tkinter__: es un archivo .ico que se usa en el módulo .py de ejemplos.
  * __png_para_boton__: es un archivo .png que se usa en el módulo .py de ejemplos.


* __documentacion_otra__
    
   Ahi se encuentra un unico fichero llamado __GUIA_USUARIO_v1.1__ y contiene la guia de usuario que explica como operar a base pantallazos además de otras explicaciones.


* __documentacion_tecnica__

    Ahi se encuentra un unico fichero llamado __DISEÑO_FUNCIONAL_v1.1__ y contiene el diseño funcional del app donde se explica la arquitectura usada y se entra también muy en detalle del código complementandolo con ejemplos para entender su alcance..

* __templates__

  Contiene los archivos que son necesarios para poder ejecutar el app:
  
  * __ico_app__: fichero .ico
  * __PLANTILLA_CONTROL_VERSIONES__: plantilla excel para poder descargar todos los objetos con cambios
  * __PLANTILLA_DIAGNOSTICO_MS_ACCESS__: plantilla excel para el diagnostico de una base de datos MS Access
  * __PLANTILLA_DIAGNOSTICO_SQL_SERVER__: plantilla excel para el diagnostico de un servidor SQL Server
  * __GUIA_USUARIO_V1.1__: guia de usuario (pdf)
  * __img_guia_usuario__: archivo png
  * __img_boton_procesos__: archivo png
  * __img_boton_add__: archivo png
  * __img_boton_clear__: archivo png
  * __img_boton_sql_server_authentication__: archivo png
  * __img_seleccionar_all_none__: archivo png
  * __img_boton_dependencias_sql_server__: archivo png
  * __img_control_versiones_boton_ver__: archivo png
  * __img_control_versiones_boton_excel__: archivo png
  * __img_control_versiones_boton_migrar_lineas_codigo__: archivo png
  * __img_control_versiones_boton_merge_bbdd_fisica__: archivo png

  Contiene, asimismo, el fichero __APP_1_GUI.spec__ que se ha de usar para poder compilar el app en .exe (ver el manual __MANUAL_PARA_COMPILAR_EN_EXE_v1.1__)
  
## __REQUISITOS FUNCIONAMIENTO DEL APP__

__MS ACCESS__:
  * deshabilitar los password VBA de los MS ACCESS que se vayan a usar.
  * deshabilitar la macro AutoExec si existiese (cambiadole el nombre de forma temporal por ejemplo).

__SQL SERVER__: 
  * hay que configurar la lista de los servidores deseada en el módulo __APP_3_BACK_END_SQL_SERVER__ en la variable __lista_GUI_sql_server_servidor__ (fila 48).

__EJECUCIÓN DEL APP DESDE LA INTERFAZ DE PROGRAMACIÓN__:

Para ejecutar el app desde la consola de la interfaz de programación que se use hay que guardar en una misma carpeta en el PC los archivos de las carpetas codigo y templates mencionadas en este README. Una vez guardados, hay que ejecutar el módulo APP_1_GUI.py.

## __REQUISITOS SISTEMA Y LIBRERIAS PYTHON__

El app se ha desarrollado y probado en entorno Windows (10) usando la versión 3.9.5 de Python. No se ha probado con otros sistemas operativos por lo que podria haber errores.

Librerias que requieren instalación (pip install):

<img width="161" height="173" alt="image" src="https://github.com/user-attachments/assets/89525d30-495b-490f-8885-bd54653a5752" />


Librerias nativas Python:

<img width="95" height="369" alt="image" src="https://github.com/user-attachments/assets/2695c2b4-135e-49e1-95fe-6ddeb788c17d" />





















