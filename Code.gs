/*
localizable-sheet-script
A Google Sheets script that will take a sheet in a specific format and return iOS and Android localization files.
https://github.com/cobeisfresh/localizable-sheet-script
Created by COBE http://cobeisfresh.com/ Copyright 2017 COBE
License: MIT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

// Configurable properties

/*
   The number of languages you support. Please check the README.md for more
   information on column positions.
*/
var NUMBER_OF_LANGUAGES = 1;

/* 
   The script expects two columns for iOS and Android identifiers, respectively,
   and a column after that with all of the string values. This is the position of
   the iOS column.
*/
var FIRST_COLUMN_POSITION = 1;

/*
   The position of the header containing the strings "Identifier iOS" and "Identifier Android"
*/
var HEADER_ROW_POSITION = 1;

/*
   True if iOS output should contain a `Localizable` `enum` that contains all of
   the keys as string constants.
*/
var IOS_INCLUDES_LOCALIZABLE_ENUM = true;


// Constants

var LANGUAGE_IOS      = 'iOS';
var LANGUAGE_ANDROID  = 'Android';
var DEFAULT_LANGUAGE = LANGUAGE_IOS;


// Export

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Export')
      .addItem('iOS', 'exportForIos')
      .addItem('Android', 'exportForAndroid')
      .addToUi();
}

function exportForIos() {
  var e = {
    parameter: {
      language: LANGUAGE_IOS
    }
  };
  exportSheet(e);
}

function exportForAndroid() {
  var e = {
    parameter: {
      language: LANGUAGE_ANDROID
    }
  };
  exportSheet(e);
}

/*
   Fetches the active sheet, gets all of the data and displays the
   result strings.
*/
function exportSheet(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var options = getExportOptions(e)
  var rowsData = getRowsData_(sheet, options);

  var strings = [];
  for (var i = 0; i < NUMBER_OF_LANGUAGES; i++) {
    strings = strings.concat(makeString(rowsData, i, options));
  }
  return displayTexts_(strings, options);
}

function getExportOptions(e) {

  var options = {};
  options.language = e && e.parameter.language || DEFAULT_LANGUAGE;  
  return options;
}


// UI Elements

function makeLabel(app, text, id) {
  var lb = app.createLabel(text);
  if (id) lb.setId(id);
  return lb;
}

function makeListBox(app, name, items) {
  var listBox = app.createListBox().setId(name).setName(name);
  listBox.setVisibleItemCount(1);
  
  var cache = CacheService.getPublicCache();
  var selectedValue = cache.get(name);
  Logger.log(selectedValue);
  for (var i = 0; i < items.length; i++) {
    listBox.addItem(items[i]);
    if (items[1] == selectedValue) {
      listBox.setSelectedIndex(i);
    }
  }
  return listBox;
}

function makeButton(app, parent, name, callback) {
  var button = app.createButton(name);
  app.add(button);
  var handler = app.createServerClickHandler(callback).addCallbackElement(parent);;
  button.addClickHandler(handler);
  return button;
}

function makeTextBox(id, content) {
  var textArea = '<textarea rows="10" cols="80" id="' + id + '">' + content + '</textarea>';
  return textArea;
}

function displayTexts_(texts, options) {
  
  var app = HtmlService.createHtmlOutput().setWidth(1200).setHeight(800);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var headersRange = sheet.getRange(HEADER_ROW_POSITION, FIRST_COLUMN_POSITION + 2, 1, sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];

  for (var i = 0, j = 0; i < texts.length; i++, j++) {
    if(options.language == LANGUAGE_IOS) {
      if(i == 0) {
        app.append(makeTextBox("export_" + i, "Enums:", texts[i]))
        i++;
      }
      app.append(makeTextBox("export_" + i, headers[j] + " plist", texts[i]))
      i++;
    }
    app.append(makeTextBox("export_" + i, headers[j], texts[i]))
  }
  
  SpreadsheetApp.getUi().showModalDialog(app, "Translations");

  return app; 
}


// Creating iOS and Android strings

function makeString(object, textIndex, options) {
  switch (options.language) {
    case LANGUAGE_ANDROID:
      return makeXmlString(object, textIndex, options);
      break;
    case LANGUAGE_IOS:
      return makeIosString(object, textIndex, options);
      break;
    default:
      break;
  }
}

/*
   Creates the strings.xml file for Android.
*/

function makeXmlString(object, textIndex, options) {

  var prevIdentifier = "";
  
  var root = XmlService.createElement('resources');
  var stringArray;
  var pluralArray = undefined;
  var pluralFormat;
  var pluralArgument;
  
  for(var i=0; i<object.length; i++) {
    
    var o = object[i];
    var identifier = o.identifierAndroid;
    
    var text = o.texts[textIndex];
    
    
    if(identifier == "") {
      continue;
    }
    
//    identifier = identifier.replace(/\./g, "_"); use formula in sheet for the [Identifier Android]:
//=SUBSTITUTE(B2; "."; "_")


    if (text == undefined || text == "") {
      if(pluralArray == undefined) {
        continue;
      } else {
        var arrayPosition = identifier.indexOf("[]")
        if(arrayPosition > 0) {
          var formatString = identifier.substr(0, arrayPosition)
          var arrayIdentifier = identifier.substr(arrayPosition + 2)

          pluralArgument = arrayIdentifier;
        }
        continue;
      }
    }

    if(typeof text === 'string') {
      text = text.replace(/\n/g, "\\n");
      text = text.replace(/&/g, "&amp;");
      text = text.replace(/\'/g, "\\'");
      text = text.replace(/</g, "&lt;");
      text = text.replace(/>/g, "&gt;");
      text = text.replace(/"/g, "\\\"");
    }
  
    if(pluralArray == undefined) {
      
      if(identifier != prevIdentifier && prevIdentifier != "") {
        root.addContent(stringArray);
        stringArray = undefined
        prevIdentifier = "";
      }
      
      var arrayPosition = identifier.indexOf("[]");
      if(arrayPosition > 0) {
        var arrayIdentifier = identifier.substr(0, arrayPosition);
        
        if(identifier != prevIdentifier) {
          stringArray = XmlService.createElement('string-array').setAttribute('name', arrayIdentifier);
        }
        var item = XmlService.createElement('item').setText(text);
        stringArray.addContent(item)
        
        prevIdentifier = identifier;
        
      } else {
        arrayPosition = identifier.indexOf("[p]");
        if(arrayPosition > 0) {
          var arrayIdentifier = identifier.substr(0, arrayPosition)
          pluralArray = XmlService.createElement('plurals').setAttribute('name', arrayIdentifier);
          pluralFormat = text
        } else {
          var item = XmlService.createElement('string').setAttribute('name', identifier).setText(text);
          
          root.addContent(item)
        }
      }
    } else {
      var replaceCriteria = "%#@" + pluralArgument + "@"
      var text = pluralFormat.replace(replaceCriteria, text)
      var item = XmlService.createElement('item').setAttribute('quantity', identifier).setText(text);
      pluralArray.addContent(item)
      if(identifier == "other") {
          root.addContent(pluralArray);
          pluralArray = undefined
          prevIdentifier = "";
      }
    }
  }

  if(prevIdentifier != "") {
    root.addContent(stringArray);
  }

  var document = XmlService.createDocument(root);

  var xml = XmlService.getPrettyFormat().setEncoding('UTF-8').format(document);

  return xml;
}

/*
   Creates the Localizable.strings file and a Localizable enum for iOS.
*/

//http://ramblings.mcpher.com/Home/excelquirks/gassnips/es6shim
var Set = cEs6Shim.Set;
var Map = cEs6Shim.Map;

function makeIosString(object, textIndex, options) {

  Logger.clear()
  var exportString = "";
  
  var plist = XmlService.createElement('plist').setAttribute('version', "1.0");
  var root = XmlService.createElement('dict');
  var stringArray, paramArray;
  var iOSConstants = new Map();
  var iOSConstantToStrings = new Map();
  
  for(var i=0; i<object.length; i++) {
    var o = object[i];
    var identifier = o.identifierIos;
    var text = o.texts[textIndex];
    
   if(identifier == "" || identifier == undefined) {
      continue;
    }
    
    if (text == undefined || text == "") {
      if (stringArray == undefined) {
        exportString += returnIfComment(identifier)
        continue;
      } else {
        var arrayPosition = identifier.indexOf("[]")
        if(arrayPosition > 0) {
          var formatString = identifier.substr(0, arrayPosition)
          var arrayIdentifier = identifier.substr(arrayPosition + 2)
          
          var key = XmlService.createElement('key').setText(arrayIdentifier);
          stringArray.addContent(key)
          
          paramArray = XmlService.createElement('dict');
          
          key = XmlService.createElement('key').setText("NSStringFormatSpecTypeKey");
          paramArray.addContent(key)
          var string = XmlService.createElement('string').setText("NSStringPluralRuleType");
          paramArray.addContent(string);
          key = XmlService.createElement('key').setText("NSStringFormatValueTypeKey");
          paramArray.addContent(key)
          string = XmlService.createElement('string').setText(formatString);
          paramArray.addContent(string);
          continue;
        } else {
         //fallthrough 
          if(paramArray == undefined && stringArray != undefined) {
            root.addContent(stringArray)
            stringArray = undefined
          }
        }
      }
    }
        
    if(typeof text === 'string') {
      text = text.replace(/"/g, "\\\"");
      text = stringFormatToSwift("", text)
    }
    
    if(paramArray != undefined) {
      if(text == undefined || text == "") {
        exportString += returnIfComment(identifier)
        continue;
      }
      var key = XmlService.createElement('key').setText(identifier);
      paramArray.addContent(key)
      var string = XmlService.createElement('string').setText(text);
      paramArray.addContent(string);
      if(identifier == "other") {
        stringArray.addContent(paramArray)
        paramArray = undefined
      }
      continue;
    }
    
    if(stringArray != undefined) {
      root.addContent(stringArray)
      stringArray = undefined
    }
    
    var arrayPosition = identifier.indexOf("[p]");
    if(arrayPosition > 0) {
      var arrayIdentifier = identifier.substr(0, arrayPosition)
      
      if(textIndex === 0) {
        fillConstants(iOSConstants, arrayIdentifier, iOSConstantToStrings, text)
      }
      exportString += '"' + arrayIdentifier + '" = "' + text + "\";\n";
      
      var key = XmlService.createElement('key').setText(arrayIdentifier);
      root.addContent(key)
      
      stringArray = XmlService.createElement('dict');
      
      key = XmlService.createElement('key').setText("NSStringLocalizedFormatKey");
      stringArray.addContent(key);
      var string = XmlService.createElement('string').setText(text);
      stringArray.addContent(string);

    } else {
      if(text != undefined && text !== "") {
        if(textIndex === 0) {
          fillConstants(iOSConstants, identifier, iOSConstantToStrings, text)
        }
        exportString += '"' + identifier + '" = "' + text + "\";\n";
      } else {
        exportString += returnIfComment(identifier)
      }
    }
  }
  
  if(stringArray != undefined) {
    root.addContent(stringArray);
  }
  plist.addContent(root);

  var document = XmlService.createDocument(plist);
  var docType = XmlService.createDocType("plist");
  docType.setPublicId("-//Apple//DTD PLIST 1.0//EN");
  docType.setSystemId("http://www.apple.com/DTDs/PropertyList-1.0.dtd")
  document = document.setDocType(docType);
  
//  exportString += "\n" + XmlService.getPrettyFormat().setEncoding('UTF-8').format(document);
  
  Logger.log("Finishing export")
//  var iOSEnum = "enum Localize: String, Localizable {\n" + ConstructEnum(iOSConstants, "  ") + "}\n"
  var resultArray = []
  if(textIndex === 0) {
    var iOSEnum = "enum Localized" + ConstructEnum(iOSConstants, "\t", iOSConstantToStrings)
    resultArray.push(iOSEnum);
  }
  resultArray.push(XmlService.getPrettyFormat().setEncoding('UTF-8').format(document))
  resultArray.push(exportString)
  return resultArray
}

function returnIfComment(identifier) {
  var trimmedIdentifier = identifier.toString().trim();
  if(trimmedIdentifier.indexOf('/*') == 0 && trimmedIdentifier.indexOf('*/', identifier.length-2) > 0) {
    return "\n/* " + trimmedIdentifier.substr(2,identifier.length-4) + "*/\n\n";
  }
  return "";
}

function capitalizeFirstLetter(aString) {
  var firstChar = aString.charAt(0).toUpperCase()
  if(!isNaN(firstChar)) {
    firstChar = "_" + firstChar
  }
  return firstChar + aString.slice(1).toLowerCase();
}

function stringToSwiftKey(aString) {
  var firstChar = aString.charAt(0).toUpperCase()
  if(!isNaN(firstChar)) {
    firstChar = "_" + firstChar
  }
  return firstChar + aString.slice(1).toLowerCase();
}

function stringToSwiftValue(aString) {
  var firstChar = aString.charAt(0)
  var prefix = ""
  var lowerCasedString = aString.toLowerCase()
  if(!isNaN(firstChar)) {
    prefix = "_"
  } else if(lowerCasedString === "default") {
    prefix = "_"
  }
  return prefix + lowerCasedString;
}


function stringFormatToSwift(identifier, value) {
  return value.replace(/([^%]|\b|^|%%)(%{1})(\d+\$){0,1}s{1}/g, "$1$2$3@")
}
             

function fillConstants(iOSConstants, identifier, valuesMap, stringValue) {

//  Logger.log("iOSConstants, start:" + Array.from(iOSConstants.keys()))
  var slicedIdentifier = identifier.split(".")
  var value = slicedIdentifier[slicedIdentifier.length - 1]
  var object = iOSConstants;
  if(value != undefined) {
    value = stringToSwiftValue(value)
    var existingObject = object
    var debugOutput = ""
    var keysArray = slicedIdentifier.slice(0, -1)
//    var keysArrayLowercased = slicedIdentifier.slice(0, -1)
//    var keysArray = Array.from(keysArrayLowercased , function (d) { return d.toUpperCase(); })
    for(index in keysArray) {
      var key = stringToSwiftKey(slicedIdentifier[index])
      existingObject = object.get(key);
      if(existingObject == undefined) {
        existingObject = new Map()
        object.set(key, existingObject)
        debugOutput += key + "(new),"
      } else {
        debugOutput += key + ","
      }
      object = existingObject
    }
    
    Logger.log("Got " + identifier + " sliced: " + debugOutput + " == " + keysArray + ", value = " + value)
    
//    var objectCopy = object
//    Logger.log("Equal " + (object === objectCopy).toString() + " type = " + typeof(object))
    
//    if(existingObject == undefined) {
//      existingObject = new Map()
//      object.set(key, existingObject)
//    }
    object.set(value, identifier)
    valuesMap.set(identifier, stringValue)
//    Logger.log("Equal after set " + (object === objectCopy).toString() + " object keys: " + Array.from(object.keys()) + " values: " + Array.from(object.values()))
//    Logger.log("After set object keys: " + Array.from(object.keys()) + " values: " + Array.from(object.values()))
  }
//  Logger.log("iOSConstants:" + Array.from(iOSConstants.keys()))
//  if(slicedIdentifier.length > 1) {
//    Logger.log("Map value after set(login):  " + iOSConstants.get("login") )
//    Logger.log("object keys:" + Array.from(object.keys()) + ", values:" + Array.from(object.values()))
//  }
}

function ConstructEnum(constantsMap, offset, constantToStrings) {

  if(constantsMap == undefined) {
    return "{}\n"
  }
//  if(typeof(constantsMap) == "string") {
//    return offset + "static let item = " + constantsMap
//  }
  if(typeof(constantsMap) !== "object") {
    Logger.log("ConstructEnum: unexpected type of constantsMap:" + typeof(constantsMap))
    return "{}\n"
  }
  var enumHasCases = false
  var enumHasFormats = false
  var enumContent = ""
  var keys = Array.from(constantsMap.keys())
  Logger.log("ConstructEnum for keys array " + Array.from(constantsMap.keys()))
  for(var index in keys) {
    var key = keys[index]
//    if(keysMap.keys().lenght < 10) {
      Logger.log("ConstructEnum for key " + key + ", value = " + constantsMap.get(key))
//    }
    
    var mapResult = ""
    var value = constantsMap.get(key)
    if(typeof(value) === "string") {
      enumHasCases = true
      var stringValue = constantToStrings.get(value)
////      stringValue.replace("(?<ArgumentPosition>([^%]|\b|^|%%)%{1}(\d+\$){0,1})(?<StringCharacter>s{1})", "${ArgumentPosition}@") 
//      stringValue = stringValue.replace(/([^%]|\b|^|%%)(%{1})(\d+\$){0,1}s{1}/g, "$1$2$3@")
////      stringValue += " replaced"
////      stringValue.replace(new RegExp("\w*", 'g'), '$&rep')
////      stringValue = stringValue.replace(/\w*/g, '$&rep')
      mapResult = offset + "case " + key + " = \"" + value + "\"\t/*" + stringValue + "*/\n"
      if(value.endsWith("format")) {
        enumHasFormats = true
        
//        
      }
    } else if(typeof(constantsMap) === "object") {
      var enumObject = constantsMap.get(key)
      
      mapResult = offset + "enum " + key + ConstructEnum(constantsMap.get(key), offset + "\t", constantToStrings)
//      mapResult += offset + "}\n"
    } else {
      Logger.log("ConstructEnum: unexpected type of constantsMap:" + typeof(constantsMap))
    }
    enumContent += mapResult
  }

  var resultString = ""
  if(enumHasCases) {
    if(enumHasFormats) {
      resultString = ": String, LocalizableFormat"
    } else {
      resultString = ": String, Localizable"
    }
  }
  resultString += " {\n" + enumContent + offset.slice(0, -1) + "}\n"
  return resultString
}

// Data fetching

/*
   Gets the titles for the first row from the speadsheet, in lower case and without spaces.
   - returns: a string array of the headers
*/
function getNormalizedHeaders(sheet, options) {
  var headersRange = sheet.getRange(1, FIRST_COLUMN_POSITION, HEADER_ROW_POSITION, sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  return normalizeHeaders(headers);
}

/*
   Removes all empty cells from the headers string array, and normalizes the rest into camelCase.
   - returns: a string array containing a list of normalized headers
*/
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

/*
   Converts a header string into a camelCase string.
    - returns a string in camelCase
*/
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

/*
   Gets all of the data from the sheet.
    - returns an array of objects containing all the necessary data for display.
*/
function getRowsData_(sheet, options) {
  
  var dataRange = sheet.getRange(HEADER_ROW_POSITION + 1, FIRST_COLUMN_POSITION, sheet.getMaxRows(), sheet.getMaxColumns());
  var headers = getNormalizedHeaders(sheet, options);
  var objects = getObjects(dataRange.getValues(), headers);
  
  return objects;
}

/*
   Gets the objects for the cell data. For each cell, the keys are the headers and the value is the
   data inside the cell.
   - returns: an array of objects with data for displaying the final string
*/
function getObjects(data, keys) {
  
  var objects = [];
  
  for (var i = 0; i < data.length; ++i) {
    
    var object = {
      "texts": []
    };
    
    var hasData = false;
    
    for (var j = 0; j < data[i].length; ++j) {
      
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        cellData = "";
      }
      
      if (keys[j] != "identifierIos" && keys[j] != "identifierAndroid") {
        object["texts"].push(cellData);
      } else {
        object[keys[j]] = cellData;
      }
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}


// Utils

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose_(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}