
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python permite realizar __control de versiones__ de código VBA entre 2 bases de datos MS Access y también de código T-SQL entre 2 bases de datos SQL Server (pueden estar en servidores distintos).

En ambos entornos de bases de datos:

  * localiza para los mismos scripts presentes en cada base de dato las lineas de código donde hay diferencias, las marca de un color. Tambien localiza los scripts que tan solo estan en una de las dos bases de datos.
    
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

* *__v1.0__*

   Corresponde a la __versión 1.0__ publicada el 20/04/2025.
  La subcarpeta contiene nada más acceder el documento __README_v1.0__ que explica el contenido de la subcarpeta.

* *__v1.1__*

   Corresponde a la __versión 1.1__ publicada el 07/01/2026.
  La subcarpeta contiene nada más acceder el documento __README_v1.1__ que explica el contenido de la subcarpeta.

  Esta versión corresponde a una refactorización de la versión anterior de la GUI basada en herencias de clases tkinter para poder configurarla dinamicamente mediante el uso de kwargs.
  Asimismo, se han hecho pequeños ajustes de sintaxis en el back-end que no afectan las funcionalidades ya desarrolladas en la versión anterior.

  El desarrollo está __finalizado y testeado__ y el código está disponible en el repositorio Github.

  Tareas pendientes (que se actualizarán en los próximos dias):
  
  * *__Adaptación del diseño funcional.__*
  * *__Manual para usar el sistema de herencias de clases de tkinter en otros proyectos.__*
