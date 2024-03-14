# DEDIE


###############################
#                             #
#   1. ACTUALIZAR SCRIPT      #
#                             #
###############################

// ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
// ;;  Función para asignar texto a elemento *nombre elemento*      ;;
//;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
function asignar*NombreElemento*() {
  Word.run(async (context) => {
    const selection = context.document.getSelection();                         
    selection.load("text, parentBody, parentContentControlOrNullObject");
    await context.sync();
    const textoSeleccionado = selection.text; 
    const *nombreElemento* = ` *Remplazar el contenido del <div> hasta que este se encuentre con algun <p>* ${textoSeleccionado} *Remplazar con el resto del contenido del <div> respetando cierre de elementos </> *`; 

    if (!selection.parentContentControlOrNullObject.isNullObject) {     
      selection.parentContentControlOrNullObject.insertText(*nombreElemento*, "Replace"); 
    } 
    else {
      const range = selection.getRange("Start");
      range.insertText(*nombreElemento*, "Replace");
      range.font.color = "*cambiar el color a uno que les parezca bien, mientras no se repitan jaja*";
    }
    selection.delete();
    await context.sync();
  })
}


##########################
#                        #
#   2.ACTUALIZAR HTML    #
#                        #
##########################


<div>
  <button id="btn*Elemento*" class="ms-Button">*NombreElemento*</button>
</div> <br>



#################################################
#                                               #
#   3. DAR DE 'ALTA' LA FUNCION EN TRYCATCH()   #
#                                               #
#################################################

$("#btn*Elemento*").on("click", () => tryCatch(asignar*NombreElemento*));

