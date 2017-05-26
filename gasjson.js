/**
Version 1.03

ChangeLog:

1.03
- Ahora se puede elegir si la funcion devuelve el objeto como JSON o como texto

1.02
- Modificado este comentario

1.01
- Ahora se puede elegir si se toma o no el resultado de la cache o si se pisa dicho resultado

1.00
- Version inicial de obtenerJSONDesdeSpreadsheet


Funcion para obtener un JSON en formato cadena de caracteres a partir de una hoja de calculo
de Google Drive.

- idArchivo: el ID del archivo de Google Drive
- idHoja: el ID de la hoja en concreto del archivo
- columnasADevolver: es un objeto del tipo:
    {
      columnasDeseadas: ['camelCase1', 'camelCase2', 'camelCase2'...],
      columnaNoDeseadas: ['camelCase1', 'camelCase2', 'camelCase2'...]
    }
    o del tipo:
    {
      columnasDeseadas: ['*'],
      columnaNoDeseadas: ['camelCase1', 'camelCase2', 'camelCase2'...]
    }  
- funcionDeFiltrado: o null (si no se quiere filtrar) o una funcion que devuelve true para los elementos
que se deseen mostrar o false para los que no. Algo del tipo:
                       function(elemento) {
                         return elemento.id == "AA22";
                       }
- nombreConsulta: especifica un nombre para la consulta con el fin de guardarla en la cache. Si es "" entonces no
                  se guarda la consulta en la cache.
                  
- tomarDeCacheSiEsPosible: si es true se intenta tomar la consulta de la caché.
- devolverEnJsonNativo: si es true devuelve un JSON, si no, devuelve el JSON en texto
**/
function obtenerJSONDesdeSpreadsheet(idArchivo, idHoja, columnasADevolver, funcionDeFiltrado, nombreConsulta, tomarDeCacheSiEsPosible, devolverEnJsonNativo) {
  var cache = CacheService.getScriptCache();
  // Si procede, miramos a ver si la consula ya esta en la cache  
  if (nombreConsulta != "" && tomarDeCacheSiEsPosible) {
    var resultadoEnCache = cache.get(nombreConsulta);
    if (resultadoEnCache != null) {
      return resultadoEnCache;
    }
  }


  // Abre el archivo de hoja de cálculo y obtiene todas las hojas de la misma
  var hojas = SpreadsheetApp.openById(idArchivo).getSheets();

  // Guarda en hoja la hoja que se esta buscando 
  var hoja = undefined;
  var i = 0;
  while (i < hojas.length) {
    if (hojas[i].getSheetId() == idHoja) {
      hoja = hojas[i];
    }
    i++;
  }

  // Cargamos en un array todos los datos de la hoja
  var arrayDatos = hoja.getDataRange().getValues();

  // Eliminamos los encabezados
  var encabezadosLenguajeNatural = arrayDatos.shift();
  var encabezadosCamelCase = arrayDatos.shift();
  var encabezados3 = arrayDatos.shift();

  // Creamos el objeto que vamos a devolver   
  // En la siguiente linea podemos ver la estructura del archivo JSON que obtenemos
  var objetoDevuelto = {
    encabezadosLenguajeNatural: {},
    encabezadosCamelCase: {},
    encabezados3: {},
    numerosColumna: {},
    datos: []
  };

  // Obtenemos las correspondencia encabezadoLenguajeNatural -> encabezadoCamelCase
  for (var j = 0; j < encabezadosLenguajeNatural.length; j++) {
    objetoDevuelto.encabezadosLenguajeNatural[encabezadosLenguajeNatural[j]] = encabezadosCamelCase[j];
  }

  // Obtenemos las correspondencia encabezadoCamelCase -> encabezadoLenguajeNatural
  for (var j = 0; j < encabezadosLenguajeNatural.length; j++) {
    objetoDevuelto.encabezadosCamelCase[encabezadosCamelCase[j]] = encabezadosLenguajeNatural[j];
  }

  // Obtenemos las correspondencia encabezadoCamelCase -> numeroColumna
  for (var j = 0; j < encabezadosLenguajeNatural.length; j++) {
    objetoDevuelto.numerosColumna[encabezadosCamelCase[j]] = j + 1;
  }

  // Obtenemos los numeros de columna que interesan a quien invoca el metodo  
  // Vamos a obtener un array llamado columnasInteresantes del tipo, por ejemplo, [1, 3, 5, 6]
  var columnasInteresantes = []
  if (columnasADevolver.columnasDeseadas[0] == "*") {
    var numeroColumna = 0;
    for (var j in objetoDevuelto.numerosColumna) {
      columnasInteresantes.push(numeroColumna);
      numeroColumna++;
    }
  }
  else {
    for (var i = 0; i < columnasADevolver.columnasDeseadas.length; i++) {
      var nombreColumnaDeseada = columnasADevolver.columnasDeseadas[i];
      var numeroColumnaDeseada = objetoDevuelto.numerosColumna[nombreColumnaDeseada];
      columnasInteresantes.push(numeroColumnaDeseada - 1);
    }
  }
  for (var i = 0; i < columnasADevolver.columnasNoDeseadas.length; i++) {
    var nombreColumnaCamelCaseABorrar = columnasADevolver.columnasNoDeseadas[i];
    var numeroDeColumnaABorrar = objetoDevuelto.numerosColumna[nombreColumnaCamelCaseABorrar] - 1;
    var posicion = columnasInteresantes.indexOf(numeroDeColumnaABorrar);
    if (posicion > -1) {
      columnasInteresantes.splice(posicion, 1);
    }
  }

  objetoDevuelto.columnasInteresantes = columnasInteresantes;

  // Exportamos los datos
  // Recorremos las filas
  for (var i = 0; i < arrayDatos.length; i++) {
    var objetoActual = {};
    // Recorremos las columnas
    for (var j = 0; j < columnasInteresantes.length; j++) {
      var claveParaElObjetoActual = encabezadosCamelCase[columnasInteresantes[j]];
      var valorParaElObjetoActual = arrayDatos[i][columnasInteresantes[j]];
      if (claveParaElObjetoActual != "") {
        objetoActual[claveParaElObjetoActual] = valorParaElObjetoActual;
      }
    }
    objetoDevuelto.datos.push(objetoActual);
  }


  // Filtramos los datos conforme la funcion que nos envien
  if (funcionDeFiltrado != null) {
    objetoDevuelto.datos = objetoDevuelto.datos.filter(funcionDeFiltrado);
  }

  var objetoDevueltoDefinitivo;

  if (!devolverEnJsonNativo) {
    // Convertimos a cadena el objeto JSON a devolver  
    objetoDevueltoDefinitivo = JSON.stringify(objetoDevuelto);
  }
  else {
    //Dejamos el objeto en JSON nativo
    objetoDevueltoDefinitivo = objetoDevuelto;
  }

  // Si procede, guardamos la consulta en la caché
  // Si es demasiado grande no pasa nada
  if (nombreConsulta != "") {
    try {
      cache.put(nombreConsulta, objetoDevueltoDefinitivo, 21600); // cache para 6 horas
    }
    catch (error) {
      Logger.log(error.message);
    }
  }

  return objetoDevueltoDefinitivo;
}


/**
Version 1.00

ChangeLog:

1.00
- Version inicial de addFilaDesdeJSON

Funcion que permite insertar una nueva linea con datos dentro de una hoja de calculo a
partir de un objeto JSON.

Parámetros:
- idArchivo: el ID del archivo de Google Drive
- idHoja: el ID de la hoja en concreto del archivo
- jsonAEscribir: objeto JSON a escribir.
***/
function addFilaDesdeJSON(idArchivo, idHoja, jsonAEscribir) {
  // Abre el archivo de hoja de cálculo y obtiene todas las hojas de la misma
  var hojas = SpreadsheetApp.openById(idArchivo).getSheets();

  // Guarda en hoja la hoja que se esta buscando 
  var hoja = undefined;
  var i = 0;
  while (i < hojas.length) {
    if (hojas[i].getSheetId() == idHoja) {
      hoja = hojas[i];
    }
    i++;
  }

  // Cargamos en un array todos los datos de la hoja
  var arrayDatos = hoja.getDataRange().getValues();

  Logger.log(arrayDatos);

  // Eliminamos los encabezados
  var encabezadosLenguajeNatural = arrayDatos.shift();
  var encabezadosLenguajeNatural = arrayDatos.shift();
  var encabezados3 = arrayDatos.shift();

  var arrayAEscribir = [];
  for (var i = 0; i < encabezadosLenguajeNatural.length; i++) {
    if (jsonAEscribir[encabezadosLenguajeNatural[i]] != undefined) {
      arrayAEscribir.push(jsonAEscribir[encabezadosLenguajeNatural[i]]);
    }
    else {
      arrayAEscribir.push("");
    }
  }

  hoja.appendRow(arrayAEscribir);
}


/**
Version 1.00

ChangeLog:

1.00
- Version inicial
**/
function getDatosFilasSeleccionadas(hoja) {
  var rangoSeleccionado = hoja.getActiveRange();
  var numeroFilaInicial = rangoSeleccionado.getRowIndex();
  var numeroFilasSeleccionadas = rangoSeleccionado.getHeight();
  var datos = hoja.getRange(numeroFilaInicial, 0, numeroFilasSeleccionadas, hoja.getDataRange().getWidth());
  return datos;
}
