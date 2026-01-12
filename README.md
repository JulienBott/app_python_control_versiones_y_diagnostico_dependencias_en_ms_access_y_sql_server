
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

   Corresponde a la versión 1.1, publicada inicialmente el 07/01/2026.
   La subcarpeta contiene, al acceder, el documento README_v1.1, donde se explica en detalle el contenido y organización de esta versión.
   
   Esta versión corresponde a una refactorización completa de la GUI respecto a la versión anterior, basada en un sistema de herencias de clases en Tkinter, que permite su configuración dinámica mediante el uso de kwargs.
   
   Asimismo, se han realizado ajustes menores de sintaxis en el back-end, que no afectan a las funcionalidades existentes respecto a la versión 1.0.
   
   El desarrollo está finalizado y testeado, y el código se encuentra disponible en este repositorio de GitHub.


   __Actualización 07/01/2026__

     Estado inicial de la versión 1.1 tras su publicación.
     
     Tareas pendientes (en curso):
     
     * __Adaptación del diseño funcional__
     * __Manual para el uso del sistema de herencias de clases de Tkinter en otros proyectos__


   __Actualización 12/01/2026__

    Aunque la versión publicada sigue siendo la v1.1, se han realizado pequeñas modificaciones adicionales en el front-end para mejorar y ampliar el sistema de herencias de clases de la GUI.
    Estas modificaciones no alteran el comportamiento funcional del aplicativo.

  Módulos afectados:

     * __APP_1_GUI__ (pequeños cambios)
     * __APP_2_GUI_UTILS__ (pequeños cambios)
     * __APP_3_GENERAL__ (sin cambios)
     * __APP_4_BACK_END_MS_ACCESS__ (sin cambios)
     * __APP_5_BACK_END_SQL_SERVER__ (sin cambios)

   La subcarpeta contiene, al acceder, el documento README_v1.1, donde se detalla el contenido de esta versión.
   
   Además, ya se encuentra disponible el manual técnico que explica cómo utilizar el sistema de herencias de clases de Tkinter en otros proyectos, acompañado de un fichero .py con ejemplos prácticos.
   
   Tareas pendientes:

     * __Adaptación del diseño funcional__
     (mientras tanto, utilizar el diseño funcional publicado en la versión 1.0).
