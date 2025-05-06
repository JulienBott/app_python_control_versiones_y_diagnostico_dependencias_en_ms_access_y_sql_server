
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar __control de versiones__ de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos).

En ambos entornos de bases de datos:

  * localiza para los mismos scripts presentes en cada base de dato las lineas de código donde hay diferencias, las marca de un color. Tambien localiza los scripts que tan solo estan en una de las dos bases de datos.
    
  * el usuario tiene la posbilidad de realizar de forma sencilla y agil migraciones de lineas de código de un script de una base datos a otra.
    
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

  Contiene los 4 módulos de código Python:

  * __APP_1_GUI__
  * __APP_2_GENERAL__
  * __APP_3_BACK_END_MS_ACCESS__
  * __APP_3_BACK_END_SQL_SERVER__

* __documentacion_otra__

  Contiene 2 documentos pdf:

  * __GUIA_USUARIO_V1__: es la guia de usuario que explica como operar a base pantallazos de la GUI y algunas que otras explicaciones.
    
  * __MANUAL_PARA_COMPILAR_EN_EXE__: es un manual para compilar el código del app junto con sus templates (ver subcarpeta templates) en archivo .exe para poder usar el app sin necesidad de tener Python instalado en el PC en el que se use.

* __documentacion_tecnica__

    Ahi se encuentra un unico fichero llamado __DISEÑO_FUNCIONAL_V1__ y contiene el diseño funcional del app donde se explica la arquitectura usada y se entra también muy en detalle del código complementandolo con ejemplos para entender su alcance..

* __templates__

  Contiene los archivos que son necesarios para poder ejecutar el app:
  
  * __ico_app__: fichero .ico
  * __PLANTILLA_CONTROL_VERSIONES__: plantilla excel para poder descargar todos los objetos con cambios
  * __PLANTILLA_DIAGNOSTICO_MS_ACCESS__: plantilla excel para el diagnostico de una base de datos MS Access
  * __PLANTILLA_DIAGNOSTICO_SQL_SERVER__: plantilla excel para el diagnostico de un servidor SQL Server

  Contiene, asimismo, el fichero __APP_1_GUI.spec__ que se ha de usar para poder compilar el app en .exe (ver el manual __MANUAL_PARA_COMPILAR_EN_EXE__)

  
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

![image](https://github.com/user-attachments/assets/cb4ba9c1-2c59-4b5f-a28b-b6071087ae9f)


Librerias nativas Python:

![image](https://github.com/user-attachments/assets/dcc35e0d-3720-4505-af8d-be9b12515737)


## FASE DEL PROYECTO (actualizado a 2025-04-20)

La versión del app es la 1.0. En el diseño funcional, se enumera una lista de posibles futuros desarrollos.















