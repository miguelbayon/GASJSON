function ejemplo1() {
  addFilaDesdeJSON('17G_AALniqiYQoehpXpHYHI5UVuvdIkaSaNaTFuN7_gQ',
    '0', {
      anno: 2016,
      evaluacion: 1
    });
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
}






function ejemplo2() {
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






function ejemplo3() {
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
