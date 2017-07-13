/**
Para probar estas ejecuciones debes:
- Cambiar el nombre de la funcion a doGet.
- Poner IDs de hoja de Google Sheet sobre los que tengas permiso
**/

function ejemploInvocacion_obtenerJSONDesdeSpreadsheetAgrupadoPor()
{
  var columnasDeseadasYNoDeseadas = {columnasDeseadas: ['*'], columnasNoDeseadas: []};
  var objetoADevolver = obtenerJSONDesdeSpreadsheetAgrupadoPor('1d0zQXl7zgW07fpNcu5fJ6tq3NsearNnrTEVZJMcFDR4',
                                                    '0',
                                                    columnasDeseadasYNoDeseadas, 
                                                    function(elemento) { return (elemento == 'AD04');},
                                                    "horarios-20016",
                                                    false,
                                                    true,
                                                    'profesor');
  Logger.log(JSON.stringify(objetoADevolver));
                                            
  var medianteJSONP = false;
  var prefijo = ''; 
  
  if (medianteJSONP) {
    return ContentService.createTextOutput(prefijo + '(' + JSON.stringify(objetoADevolver) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  else {
    return ContentService.createTextOutput(JSON.stringify(objetoADevolver)).setMimeType(ContentService.MimeType.JSON);     
  }   
}





function ejemploInvocacion_addFilaDesdeJSON() {
  addFilaDesdeJSON('17G_AALniqiYQoehpXpHYHI5UVuvdIkaSaNaTFuN7_gQ',
    '0', {
      anno: 2016,
      evaluacion: 1
    });
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
}






function ejemploInvocacion_obtenerJSONDesdeSpreadsheet() {
  var columnasDeseadasYNoDeseadas = {
    columnasDeseadas: ['codigo', 'materia', 'abreviada'],
    columnasNoDeseadas: []
  };
  var texto = "Miguel";
  var objetoADevolver = obtenerJSONDesdeSpreadsheet('19lMEa0hYsRX0eznW6x1udX5TVcdOp42aOewSbbMb4vY',
    '198577462',
    columnasDeseadasYNoDeseadas,
    null,
    "materias",
    false,
    true);

  return ContentService.createTextOutput(objetoADevolver).setMimeType(ContentService.MimeType.JSON);
}






function ejemploInvocacion_obtenerJSONDesdeSpreadsheet() {
  var columnasDeseadasYNoDeseadas = {
    columnasDeseadas: ['*'],
    columnasNoDeseadas: []
  };
  var texto = "Miguel";
  var objetoADevolver = obtenerJSONDesdeSpreadsheet('17G_AALniqiYQoehpXpHYHI5UVuvdIkaSaNaTFuN7_gQ',
    '0',
    columnasDeseadasYNoDeseadas,
    function(elemento) {
      return (elemento.fecha.getDate() == 5);
    },
    "",
    false,
    true);

  return ContentService.createTextOutput(objetoADevolver).setMimeType(ContentService.MimeType.JSON);
}
