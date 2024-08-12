# Hub de libros - Quinto Sprint
Proyecto hecho por Lizeth Consuelo Bañuelos Ruelas.


# Descripción
Hub de lectura donde el usuario puede gestionar libros que lee, le gustaron, así como favoritos y verlos en una lista completa, y añadir imagenes de portada de los lbros.
Además agregar y mantener una lista de libros diversos.

# Objetivos
- Hacer un programa para Windows de un hub de lectura con lo siguiente:
    - Catálogo de libros
    - Libros que ya leyó el usuario
    - Libros que quiere leer
    - Libros que no le gustaron
    - Géneros favoritos
    - Libros recomendados

# Dependencias y bibliotecas
  - Visual Basic 6
  - SQL Server


# Captura de pantalla del proyecto
 - ![Inicio](https://github.com/user-attachments/assets/81cb0515-0ada-47fe-a254-51ee9fbd6d0d)
 - ![iniciodesesion](https://github.com/user-attachments/assets/cd8b5163-8724-47cb-b37c-544a08dfce93)
 - ![iniciodesesion2](https://github.com/user-attachments/assets/84ce3f55-e43a-4635-ad82-0627bf86e286)

# Instrucciones
Para comenzar, ya se debe de tener instalado Visual Basic 6 así como SQL server.
Despues, se debe de clonar el repositorio.
  - Haz clic en el botón verde que dice "code"
  - Copia el URL del repositorio que aparece, ya sea HTTP o SSH.
  - Abre la terminal y ejecuta el comando `git clone [url del repositorio]` (asegurate de estar en la carpeta donde quieras que se clone).
  - Abre Visual Basic 6
  - Ve a File, después a "Open Project" y selecciona el archivo `.vbp` que se encuentra dentro de la carpeta del repositorio clonado.
  - También se debe de configurar la base de datos y realizar la conexión.
Una vez hecho todo lo anterior, se puede compilar y ejecutar el proyecto.

# Descripción de como se hizo
Para comenzar a entender sobre visual basic, comencé a buscar videos en Youtube con ejemplos sobre su uso y explicación de las diferentes controles y componentes, conforme veía los videos obtenía ideas de cómo podría implementarlo en el proyecto.
Comencé con realizar el CRUD de libros para que el usuario pudiera agregarlos, eliminarlos , editarlos y verlos (incluyendo recorrer los libros guardados dentro de ese form, mostrar el primero y último).
Después realicé la manera en que el usuario vería sus favoritos, leídos y que no le gustaron mediante una lista, teniendo la manera de agregarlos a estas listas a través de botones al estar buscando y viendo los libros.


# Diagrama entidad-relación
![diagramaER](https://github.com/user-attachments/assets/5c9b7338-ad04-4642-aee9-0b9fc98dea2a)


# Problemas conocidos
 - Al cargar la imagen al form donde se agregan o editan, la imagen tiene que tener el tamaño de 194x295 pixeles para poder adecuarse al picture box, si no se le da ese tamaño se ve solo una parte de la imagen guardada.
 - En el listview no se aprecian imagenes de la portada de los libros.
-  Al general el instalable me marcaba que tres dependencias no se encontraban en los archivos, sin embargo, al buscarlas si se encontraban ahí y no logré solucionar el problema.
# Retrospectiva

## ¿Qué hice bien?
 - Aprender cómo utilizar las funciones básicas de Visual Basic 6 y aplicarlas en un form para realizar un CRUD.
 - Entender como hacer y almacenar las imágenes como dato varbinary, así como mostrarlas ya que era algo que no había realizado antes.

## ¿Qué no salió bien?
 - Mostrar la imagen de portada del libro en un listview
 - La imagen del picture box no se ajusta automáticamente aunque pudiera realizar con un imagebox y utilizar la propiedad stretch
 - Visualizar géneros favoritos, necesito aprender sobre storage procedures para poder implementarlos.

## ¿Qué puedo hacer diferente?
 - Investigar más en foros sobre los listview para aprender a mostrar imagenes en ellas, así como el ajuste de imagenes en picture box.
 - La contraseña del usuario guardarla de manera correcta y no directamente en la base de datos.
 - Investigar más sobre la interfaz de usuario en vb6, ya que considero que hice una que necesita mejoría.
