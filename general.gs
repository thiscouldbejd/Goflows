var REGEX_METAS = ["\\","^","$",".","|","?","*","+","(",")","[","{"];

function _safeRegex(text) {
  for (var i = 0; i < REGEX_METAS.length; ++i) {
    var replace = "\\" + REGEX_METAS[i];
    var regEx = new RegExp(replace,"g");
    text = text.replace(regEx, replace);    
  }
  return text;
}

function _templateRegex(text, prefix_Value, value, prefix) {

  if (prefix && prefix !== "") prefix = prefix + ".";
  var value_Keys = Object.keys(value);
       
  for (var i = 0; i < value_Keys.length; i++) {
       
    var regEx;
    
    if (prefix_Value) {
      regEx = new RegExp(_safeRegex("{{" + prefix + prefix_Value[value_Keys[i]] + "}}"),"gi");
    } else {
      regEx = new RegExp(_safeRegex("{{" + prefix + value_Keys[i] + "}}"),"gi");
    }
    text = text.replace(regEx, value[value_Keys[i]]);

  }
  
  return text;
       
}

function _docRegex(doc, prefix_Value, value, prefix) {

  if (prefix && prefix !== "") prefix = prefix + ".";
  var value_Keys = Object.keys(value);
  var doc_Body = doc.getBody();
  
  for (var i = 0; i < value_Keys.length; i++) {
       
    var replace;
    
    if (prefix_Value) {
      replace = _safeRegex("{{" + prefix + prefix_Value[value_Keys[i]] + "}}");
    } else {
      replace = _safeRegex("{{" + prefix + value_Keys[i] + "}}");
    }
    doc_Body.replaceText(replace, value[value_Keys[i]]);
    
  }
  
  return doc;
  
}

function _getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (_isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

function _isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

function _autoResizeColumns(sheet, last_Col) {
  if (!last_Col) last_Col = sheet.getLastColumn();
  for (var i = 1; i <= last_Col; i++) {
    sheet.autoResizeColumn(i);
  }
}

function _formatSheet(sheet) {
  var last_Col = sheet.getLastColumn();
  sheet.setFrozenRows(1); // Freeze the First Row
  sheet.getRange(1,1,1,last_Col).setBorder(null, null, true, null, false, false); // Bottom Border for Row 1
  _autoResizeColumns(sheet, last_Col); // Resize all Columns
}

String.prototype.endsWith = function(suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
};
