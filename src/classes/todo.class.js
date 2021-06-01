
export class Todo{

    /** video 116 se crea metodo para recibir el objeto
     * de localstorage y reconstruir las instancias de la 
     * clase      */

                  //desestucturacion de los datos del objeto
    static fromJson( {id, tarea, completado, creado } ) {

        const tempTodo = new Todo( tarea );

        tempTodo.id = id;
        tempTodo.completado = completado;
        tempTodo.creado = creado;

        return tempTodo;

    }

    constructor( tarea ){
        
        // recibe la tarea por hacer
        this.tarea = tarea;

        /** Atributos nuevos de la clase para manejar 
         * las tareas
         */
        //crea un id unico por la fecha y hora creada
        this.id = new Date().getTime(); // regresa 123633455
        this.completado = false;
        this.creado = new Date();

    }

    /** creamos el metodo para probar la reconstruccion de
     * las instancias
     */

    imprimirClase(){
        console.log(`imprimirClase: ${this.tarea} - ${ this.id} `);
    }
}