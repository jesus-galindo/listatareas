import './styles.css';

// import { Todo } from './classes/todo.class.js';
// import { TodoList } from './classes/todo-list.class.js';
/** para no estar importando todo a esta archicvo, se crea un archivo 
 * index.js en la carpeta classes y ahi se importan todas las clases
 * y aqui solo se imorta el index.  */

/** las llaves vacias indica quew buscara el index.js por defecto ,
 * pero de ttodas formas hay que poner los nombres de las clases
 * 
*/
import { Todo, TodoList } from './classes';
//importamos el archivo con el html a insertar
import { crearTodoHtml } from './js/componentes';

/** se crea una instancia y un array vacio */
export const todoList = new TodoList();


//para cargar las tareas de localstorage al html
/** se hace un for each donde por cada elemento todo
 * se manda a crearTodoHTML este metodo crea el html de cada tarea
 */
//todoList.todos.forEach(todo => crearTodoHtml(todo));

/** En el caso de de que el argumento que se envia, es el unico
 *  se puede quitar el argumento y dejar solo el metodo que se llama
  */
//todoList.todos.forEach(todo => crearTodoHtml(todo));
todoList.todos.forEach( crearTodoHtml );



/** creamos nueva tarea para comprobar si las instancias estan creadas */
// const newTodo = new Todo('123 Tarea desde el codigo');
// todoList.nuevoTodo( newTodo );
//newTodo.imprimirClase();

/** tratamos de llamar al metodo imprimirClase  */
todoList.todos[5].imprimirClase();


//muestra todas las tareas
console.log( 'TAREAS', todoList.todos);






//const tarea = new Todo('Aprender javascript - jesus');
//console.log(tarea);

/** en este metodo se agrega la tarea con sus atributos 
 * al array
 */
//todoList.nuevoTodo(tarea);

//const tarea2 = new Todo('Comprar tarea');
//todoList.nuevoTodo(tarea2);

//console.log( todoList);

//tarea.completado = false;

//crearTodoHtml( tarea );

/** ejemplos de localstorage */

//agregar un dato
// localStorage.setItem('mi-key', 'valor de dato12');

// //borrar la info en 3 segundos
// setTimeout( () =>{

//    // localStorage.removeItem('mi-key');

// }, 3000 );


// /** ejemplo de localstorage */
// sessionStorage.setItem('dato1', 'valor de dato 1');