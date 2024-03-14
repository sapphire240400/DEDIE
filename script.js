$("#boton").on("click", () => tryCatch(asignarGreenText));
$("#btnRed").on("click", () => tryCatch(asignarImportantText));
$("#btnText").on("click", () => tryCatch(asignarText));




// ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
// ;;  Función para manejar errores de forma general ;;
//;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}


// ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
// ;;  Función para asignar texto a elemento "GREEN_TEXT"  ;;;;
//;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
function asignarGreenText() {
  Word.run(async (context) => {
    const selection = context.document.getSelection();                         // Obtener la selección actual del documento
    selection.load("text, parentBody, parentContentControlOrNullObject");
    await context.sync();

    const textoSeleccionado = selection.text; // Obtener el texto seleccionado y el párrafo que lo contiene

    const greenTextHtml = `<div class="green_text"><p>${textoSeleccionado}</p></div>`; // Construir el texto formateado

    if (!selection.parentContentControlOrNullObject.isNullObject) {     
      selection.parentContentControlOrNullObject.insertText(greenTextHtml, "Replace"); // Insertar el texto formateado en lugar del texto seleccionado
    } else {
      const range = selection.getRange("Start");
      range.insertText(greenTextHtml, "Replace");
      range.font.color = "#10C246";
    }
    selection.delete(); // Eliminar el texto seleccionado
    await context.sync();
  })
  
}


// ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
// ;;  Función para asignar texto a elemento "RED_TEXT"  ;;;;
//;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
function asignarImportantText() {
  Word.run(async (context) => {
    const selection = context.document.getSelection();                         // Obtener la selección actual del documento
    selection.load("text, parentBody, parentContentControlOrNullObject");
    await context.sync();

    const textoSeleccionado = selection.text; // Obtener el texto seleccionado y el párrafo que lo contiene

    const redTextHtml = `<div class="secc_important_text"><div class="important_text"><p>${textoSeleccionado}</p><img src="img/red_text.png" style="max-width: 60px;"> </div> </div>`; // Construir el texto formateado

    if (!selection.parentContentControlOrNullObject.isNullObject) {     
      selection.parentContentControlOrNullObject.insertText(redTextHtml, "Replace"); // Insertar el texto formateado en lugar del texto seleccionado
    } else {
      const range = selection.getRange("Start");
      range.insertText(redTextHtml, "Replace");
      range.font.color = "#FC4C6B";
    }
    selection.delete(); // Eliminar el texto seleccionado
    await context.sync();
  })
  
}




// ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
// ;;  Función para asignar texto a elemento "RED_TEXT"  ;;;;
//;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
function asignarText() {
  Word.run(async (context) => {
    const selection = context.document.getSelection();                         // Obtener la selección actual del documento
    selection.load("text, parentBody, parentContentControlOrNullObject");
    await context.sync();

    const textoSeleccionado = selection.text; // Obtener el texto seleccionado y el párrafo que lo contiene

    const TextHtml = `<div class="secc_text"><div class="titulo_text"><img src="img/soporte.png" style="max-width: 60px;"> <h1>${textoSeleccionado}</h1></div><p> </p>
</div>`; // Construir el texto formateado

    if (!selection.parentContentControlOrNullObject.isNullObject) {     
      selection.parentContentControlOrNullObject.insertText(TextHtml, "Replace"); // Insertar el texto formateado en lugar del texto seleccionado
    } else {
      const range = selection.getRange("Start");
      range.insertText(TextHtml, "Replace");
      range.font.color = "#FC4C6B";
    }
    selection.delete(); // Eliminar el texto seleccionado
    await context.sync();
  })
  
}
