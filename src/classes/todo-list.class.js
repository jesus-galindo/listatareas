import { Todo } from "./todo.class";

export class TodoList {

    constructor() {
        
        //this.todos = [];
        this.cargarLocalStorage();
    }


    /** clase parra agregar tareas */
    nuevoTodo ( todo ) {

        this.todos.push( todo );
        /** guarda el array de las tareas en localstorage 
         * en forma de json
         */
        this.guardarLocalStorage(); 
    }

    eliminarTodo ( id ) {
        /** se usa el metodo filter   https://developer.mozilla.org/es/docs/Web/JavaScript/Reference/Global_Objects/Array/filter 
         *  este metodo filtra el array con la codicion especificada y devuelve un nuevo 
         * arreglo con los elementos filtrados, en este caso, 
         * filtrará las tareas cuyo id NO sea el enviado, es decir quitara 
         * la tarea del id indicado, y devolvera un nuevo array,
         * ese nuevo array sera asignado a la variable que contiene el array original, 
         * susituyendolo.
         *  se pone != porque son diferentes tipos y pueda compararlos.
        */
         this.todos = this.todos.filter( todo => todo.id != id );
         this.guardarLocalStorage();

    }

    marcarCompletado( id ) {

        for ( const todo of this.todos){
            
            //mandamos a consola para comparar
            console.log(id, todo.id);

            if ( todo.id == id){
                //negacion de completado, si es true, se vuelve falso 
                //y viceversa
                todo.completado = !todo.completado;
                this.guardarLocalStorage();
                break;
            }

        }

       
       


    }

    eliminarCompletados() {
        /** hacemos un filtrado del array 
         * en este caso dejaremos solo los NO completados
         * de esta manera se quitaran del array los completados
         */
        
        // todo.completado = true , tareas completadas
        // !todo.completado = false, tareas no completadas
        this.todos = this.todos.filter( todo => !todo.completado);
        this.guardarLocalStorage();
    }

    /** agregamos metodos para local storage */
    guardarLocalStorage(){

        /** local storage solo acepta string, asi que se debe 
         * transformar el objeto en un json para poderlo almacenar
         * cada tarea sera un nodo del json
         */
        localStorage.setItem('todo', JSON.stringify( this.todos ));

    }

    cargarLocalStorage(){

                //si existe en el local storage la key 'todo'
        // if( localStorage.getItem('todo')){

        //     /** obtiene los datos del localstorage y los convierte
        //      * de json a un objeto
        //      */
        //     this.todos = JSON.parse( localStorage.getItem( 'todo'));

        // } else {
        //     this.todos =[];
        //     console.log('en if', this.todos);
        // }

        //conviertiendo lo de arriba en una condicion ternaria
        this.todos = (localStorage.getItem('todo')) 
                        ? JSON.parse( localStorage.getItem( 'todo'))
                        :  this.todos =[];

        
        /** se llama a fromJson para convertirlos a instancias de 
         * la clase.
         * se usa map, para mutar un array y regresarlo mutado
         */
                                //para cada objeto lo manda a fromjson
                                // y lo regresa como instancia
        //this.todos = this.todos.map( obj => Todo.fromJson( obj ) );
        
        //como hay 1 argumento que se envia, se puede quitar
        this.todos = this.todos.map( Todo.fromJson );

    }

}