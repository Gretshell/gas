function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function processForm(formObject) {
  if (formObject.searchtext) {
    var query = formObject.searchtext.trim();

    if (query.toLowerCase() === '!info') {
      return { type: 'info', data: getInfo() };
    }

    var precioMatch = query.match(/^\$(\d+(\.\d+)?)$/);
    if (precioMatch) {
      var maxPrecio = parseFloat(precioMatch[1]);
      return { type: 'games', data: searchByPrice(maxPrecio) };
    }

    return { type: 'games', data: search(query) };
  }
  return { type: 'games', data: [] };
}

function search(searchtext) {
  var spreadsheetId = '1pp4rpGfhnnB7dqmFhYLtRlbdi5Jdd5vFU6hG9uef0kI';   // ← CAMBIA ESTO
  var dataRange     = 'Almacen!B2:K';
  var data = SpreadsheetApp.openById(spreadsheetId)
               .getRange(dataRange).getValues();
  var ar = [];
  var queryLower = searchtext.toLowerCase();
  data.forEach(function(f) {
    if (String(f[0]).toLowerCase().includes(queryLower)) {
      ar.push(f);
    }
  });
  return ar;
}

function searchByPrice(maxPrecio) {
  var spreadsheetId = '1pp4rpGfhnnB7dqmFhYLtRlbdi5Jdd5vFU6hG9uef0kI';   // ← CAMBIA ESTO
  var dataRange     = 'Almacen!B2:K';
  var data = SpreadsheetApp.openById(spreadsheetId)
               .getRange(dataRange).getValues();
  var ar = [];
  data.forEach(function(f) {
    var precio = parseFloat(String(f[5]).replace(/[^0-9.]/g, ''));
    if (!isNaN(precio) && precio < maxPrecio) {
      ar.push(f);
    }
  });
  return ar;
}

function getInfo() {
  var spreadsheetId = '1pp4rpGfhnnB7dqmFhYLtRlbdi5Jdd5vFU6hG9uef0kI';   // ← CAMBIA ESTO
  var ss   = SpreadsheetApp.openById(spreadsheetId);
  var hoja = ss.getSheetByName('Info');
  if (!hoja) return null;
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return null;
  return {
    nombre:    datos[1][0],
    semestre:  datos[1][1],
    carrera:   datos[1][2],
    matricula: datos[1][3]
  };
}
