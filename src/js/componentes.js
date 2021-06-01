/** importamos al clase Todo para poder crear tareas
 * no es necesario poner el /todo-class.js, porque
 * pusimos un index en esa carpeta y ahi se exporta todo
 */
import { Todo } from '../classes'

//importamos para poder crear instancias de la lista de tareas
// en src- index.js
import { todoList } from '../index';

/** referencias en el html */

const divTodoList = document.querySelector('.todo-list');
//input donde se escriben las tareas
const txtInput  = document.querySelector('.new-todo');
//clase del link de "borrar completados"
const btnBorrar = document.querySelector('.clear-completed');
//clase de la lista para filtrar las tareas, por pendientes, hechas, ets
const ulFilters = document.querySelector('.filters');


//esta funcion crea el html para insertarlo

export const crearTodoHtml = ( todo ) => {

    /** se inserta un operador ternario, si es true, 
     * se agrega la clase cpompleted, si no, se deja vacio
     * para mostrar tachado o no la tarea
     */
    const htmlTodo = `    
        <li class=" ${ (todo.completado) ? 'completed' : '' }" data-id="${ todo.id }">
            <div class="view">
                <input class="toggle" type="checkbox" ${ (todo.completado) ? 'checked' : ''} >
                <label>${ todo.tarea }</label>
                <button class="destroy"></button>
            </div>
            <input class="edit" value="Create a TodoMVC template">
        </li> `

    const div = document.createElement('div');
    div.innerHTML = htmlTodo;

    divTodoList.append( div.firstElementChild );
   // divTodoList.append( div );

}

//Eventos

/** cuando se presiona una tecla 
 * el event mostrara la tecla que se presiono 
*/
txtInput.addEventListener('keyup', (event) => {

    //detecta si la tecla levantada fue un enter y que no este vacio
    if ( event.keyCode === 13 && txtInput.value.length > 0) {
        //creamos una nueva tarea
        //console.log(txtInput.value);  
        const nuevoTodo = new Todo( txtInput.value );
        //console.log(nuevoTodo);   
        //agrega la tarea con sus propiedades en la lista
        todoList.nuevoTodo( nuevoTodo );   
        //console.log(todoList);
        /** se inserta la tarea en html */
        crearTodoHtml( nuevoTodo);

        //vaciamos el input
        txtInput.value = '';

    }

   // console.log(event);
});

divTodoList.addEventListener('click', (event) => {

    //console.log('hiciste click');
    // para ubicar a que elemento se le dio click
    //console.log(event);
   // console.log(event.target);
   /** obtenemos el nombre de la eqiqueta a la que se le hizo click */
    console.log(event.target.localName); // input, label, button
    const nombreElemento = event.target.localName; // input, label, button

    /** recuperamos el codigo li de la tarea */
    const todoElemento = event.target.parentElement.parentElement; // input, label, button
    console.log(todoElemento);

    /** obtenemos el id del atributo data-id del li */
    const todoId = todoElemento.getAttribute('data-id');
    console.log("el id es ", todoId);

    /**para marcar como completado al dar click en el check  */
    if ( nombreElemento.includes('input')){ 
        //todoList es la instancia de la lista de tareas
        todoList.marcarCompletado( todoId );
        //marcar completado pone el contrario del estado actual
        /**  */
        todoElemento.classList.toggle('completed');
     
    } else if ( nombreElemento.includes('button') ){ 
        /** si se selecciona el button que es la X a la derecha */
        todoList.eliminarTodo( todoId );

        /** quitamos la tarea del html */
        divTodoList.removeChild( todoElemento );

     }

    console.log( todoList );
});

/** evento para eliminar las tareas completadas */
btnBorrar.addEventListener('click', ( ) => {

    todoList.eliminarCompletados();

    for ( let i = divTodoList.children.length -1; i >= 0; i-- ){

        const elemento = divTodoList.children[i];

        //buscamos si el elemento contiene la clase "completed"
        if( elemento.classList.contains('completed')){

            //elimina el li del html
            divTodoList.removeChild(elemento);
        }

        console.log(elemento);
    }



   // console.log('click en borrar');
   console.log( todoList );
});

/** evento para filtrar las tareas por estado */
ulFilters.addEventListener( 'click', ( event ) => {

    //console.log('ulfilters', event.target.text); //todos, pendientes, undefined
    const filtro = event.target.text;

    //si filtro es undefined, sale del evento
    if ( !filtro ) { return; }

   for( const elemento of divTodoList.children ){
       
       // console.log( elemento );
       
       /** usamos la clase hidden que esta configurada en el css
        *  y la agregamos y la quitamos segin sea el caso
        */
        elemento.classList.remove('hidden');

        //obtenemos las tareas que esten completadas
        const completado = elemento.classList.contains('completed');

        /** */
        switch ( filtro ){

            /** si se presiono mostrar tareas pendientes, 
             * a las tareas completadas se le agrega la clase hidden 
             * para ocultarlas en el html
             */
            case 'Pendientes':
                if( completado ) {
                    elemento.classList.add('hidden');
                }
            break;

            case 'Completados':
                if( !completado ){
                    elemento.classList.add('hidden');
                }
            break;
            
        }
    
    } 



});

