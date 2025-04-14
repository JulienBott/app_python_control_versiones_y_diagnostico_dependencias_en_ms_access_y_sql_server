
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar __control de versiones__ de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos).

En ambos entornos de bases de datos:

  * localiza para los mismos scripts presentes en cada base de dato las lineas de código donde hay diferencias, las marca de un color. Tambien localiza los scripts que tan solo estan en una de las dos bases de datos.
    
  * el usuario tiene la posbilidad de realizar de forma sencilla y agil migraciones de lineas de código de un script de una base datos a otra.
    
  * el usuario tambien puede ejecutar el merge en base de datos fisica, documentar el proceso y acceder a logs en caso de errores de migración que detallen el porque de los fallos, también genera logs para los objetos migrados correctamente.

Adicionalmente, ofrece un __diagnóstico de dependencias de objetos__ de bases de datos MS ACCESS o de una o varias bases de datos de un mismo servidor SQL SERVER. Permite para cada objeto localizar en que scripts de otros objetos se usan (excluye de la busqueda de dependencias los comentarios dentro del código para focalizarse solo en código activo). Asimismo, localiza aquellos objetos que no dependen de ningún otro. El output final es un fichero Excel para que sea más agil y flexible PARA el usuario realizar sus analisis. En el caso de SQL Server, da la posibilidad tambien de descargar todo los códigos de objetos de bases de datos en ficheros .sql.

Para __MS ACCESS__, los tipos de objeto se han limitado a las tablas, vinculos ODBC, vinculos hacia otras fuentes externas, variables públicas VBA y rutinas / funciones VBA. 

Para __SQL SERVER__, los objetos se han limitado a tablas, views, funciones y stored procedures. 

En el diseño funcional, colgado tambien en el repositorio GitHub, en el apartado "Limitaciones del app", se explican los pasos a seguir en caso de querer agregar más tipos de objeto.

## __CONTENIDO DEL REPOSITORIO GITHUB__:

Nada más acceder al repositorio, se encuentra el README que estas leyendo ahora mismo acompañado de un contrato MIT Licence donde autorizo cualquier tipo de uso del app y de su código asociado sea a nivel particular o empresarial siempre y cuando se me reconozca la autoria del app. Dicho contrato de MIT Licence tiene clausulas añadidas.

El resto del repositorio se divide por subcarpetas:

* __documentacion_tecnica__

    Contiene el diseño funcional del app. Ahi se encuentran 4 ficheros en formato pdf (en realidad es el mismo documento pero por tamaño he tenido que partirlos). Estos ficheros pdf son:

    * __DF_V1_GUIA_ARQUITECTURA__: se explica aqui la arquitectura, como se estructuran los módulos, como interactuan los distintos objetos entre si (sean clases, rutinas / funciones o variables globales). Para cada proceso que permite ejecutar el app, se explica a nivel de código como se ejecuta el proceso. Se listan tambien las limitaciones del app y tambien se enumera una lista de posibles futuros desarrollos.
 
    * __DF_V1_ANEXOS__PARTE_1__: contiene explicaciones mucho más en detalle del código de las rutinas relacionadas con los procesos de importación del código tanto VBA como T-SQL.
 
    * __DF_V1_ANEXOS__PARTE_2__: contiene explicaciones mucho más en detalle del código de algunas rutinas / funciones relacionadas con el proceso de __control de versiones y merge en base de datos fisica__. En la mayoria de los casos se adjuntan ejemplos para ilustrar las explicaciones de los distintos bloques de código.
 
    * __DF_V1_ANEXOS__PARTE_3__: contiene explicaciones mucho más en detalle del código de algunas rutinas / funciones relacionadas con el proceso de __diagnostico de dependencias__. En la mayoria de los casos se adjuntan ejemplos para ilustrar las explicaciones de los distintos bloques de código.

* __documentacion_otra__

  Contiene 2 documentos pdf:

  * __GUIA_USUARIO_V1__: es la guia de usuario que explica como operar a base pantallazos de la GUI y algunas que otras explicaciones.
    
  * __MANUAL_PARA_COMPILAR_EN_EXE__: es un manual para compilar el código del app junto con sus templates (ver subcarpeta Templates) en archivo .exe para poder usar el app sin necesidad de tener Python instalado en el PC en el que se use.

* __codigo__

  Contiene los 4 módulos de código Python:

  * __APP_1_GUI__
  * __APP_2_GENERAL__
  * __APP_3_BACK_END_MS_ACCESS__
  * __APP_3_BACK_END_SQL_SERVER__
 
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


## __REQUISITOS SISTEMA Y LIBRERIAS PYTHON__

El app se ha desarrollado y probado en entorno Windows (10) usando la versión 3.9.5 de Python. No se ha probado con otros sistemas operativos por lo que podria haber errores.

Librerias que requieren instalación (pip install):

![image](https://github.com/user-attachments/assets/cb4ba9c1-2c59-4b5f-a28b-b6071087ae9f)


Librerias nativas Python:

![image](https://github.com/user-attachments/assets/98739493-b3c8-4894-b89c-ccfb70216c57)


## FASE DEL PROYECTO (actualizado a 2025-04-12)

La versión del app es la 1.0. En el diseño funcional, se enumera una lista de posibles futuros desarrollos.















