const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
var _ = require('lodash');
const config = require('../config.json');
const types = require('./types');
const parser = require('./parser');

const DataType = types.DataType;

// class StringBuffer {
//   constructor(str) {
//     this._str_ = [];
//     if (str) {
//       this.append(str);
//     }
//   }

//   toString() {
//     return this._str_.join("");
//   }

//   append(str) {
//     this._str_.push(str);
//   }
// }


/**
 * save workbook
 */
function serializeWorkbook(parsedWorkbook, settings, dest) {

  for (let name in parsedWorkbook) {
    let setting = settings[name];
    let sheet = parsedWorkbook[name];
    let resultJson = JSON.stringify(sheet, null, config.json.uglify ? 0 : 2); //, null, 2
    let pathList = [];
    /**这里重置文件路径 */
    let route = dest;
    if (setting) {
      if (setting.path == "s") { setting.path = "server" }
      if (setting.path == "w") { setting.path = "web" };
      if (setting.path == "p") {
        pathList.push(
          "server"
        );
        pathList.push(
          "web"
        );
      }
      //=================================================
      //如果路径是public 缩写 p 俩个文件夹都存放
      if (pathList.length > 0) {
        pathList.forEach(v=>{
         let pathll = route + `/${v}`.trim();
          creatorFile(pathll,name,resultJson);
        })
      }else{
        route = route + `/${setting.path}`.trim();
        creatorFile(route,name,resultJson);
      }
    }

  }

function creatorFile(route,name,resultJson){
  if (!fs.existsSync(route)) {
    fs.mkdirSync(route);
  }
  let dest_file = path.resolve(route, name + ".json");
  fs.writeFile(dest_file, resultJson, err => {
    if (err) {
      console.error("error：", err);
      throw err;
    }
    console.log('exported json  -->  ', path.basename(dest_file));
  });
}

  /**
   * save dts
   */
  function serializeDTS(dest, fileName, settings) {

    let dts = "";

    for (let name in settings) {
      dts += formatDTS(name, settings[name]);
    }

    let dest_file = path.resolve(dest, fileName + ".d.ts");
    fs.writeFile(dest_file, dts, err => {
      if (err) {
        console.error("error：", err);
        throw err;
      }
      console.log('exported t.ds  -->  ', path.basename(dest_file));
    });

  }


  /**
   * 
   * @param {String} name the excel file name will be use on create d.ts
   * @param {Object} head the excel head will be the javescript field
   */
  function formatDTS(name, setting) {
    let className = _.capitalize(name);
    let strHead = "interface " + className + " {\r\n";
    for (let i = 0; i < setting.head.length; ++i) {
      let head = setting.head[i];
      if (head.name.startsWith('!')) {
        continue;
      }
      let typesDes = "any";
      switch (head.type) {
        case DataType.NUMBER:
          {
            typesDes = "number";
            break;
          }
        case DataType.STRING:
          {
            typesDes = "string";
            break;
          }
        case DataType.BOOL:
          {
            typesDes = "boolean";
            break;
          }
        case DataType.ID:
          {
            typesDes = "string";
            break;
          }
        case DataType.ARRAY:
          {
            typesDes = "any[]";
            break;
          }
        case DataType.OBJECT:
          {
            typesDes = "any";
            break;
          }
        case DataType.UNKNOWN:
          {
            typesDes = "any";
            break;
          }
        default:
          {
            typesDes = "any";
          }
      }
      strHead += "\t" + head.name + ": " + typesDes + "\r\n";
    }

    setting.slaves.forEach(slave_name => {
      strHead += "\t" + slave_name + ": " + _.capitalize(slave_name) + "\r\n";
    });

    strHead += "}\r\n";

    return strHead;
  }}

  module.exports = {

    /**
     * convert xlsx file to json and save it to file system.
     * @param  {String} src path of .xlsx files.
     * @param  {String} dest       directory for exported json files.
     * @param  {Number} headIndex      index of head line.
     * @param  {String} separator      array separator.
     * @param  {String} frontEndTable     directory for exported json files.
     * @param  {String} backendTable      directory for exported json files.
     * excel structure
     * workbook > worksheet > table(row column)
     */
    toJson: function (src, dest, frontEnd, backEnd) {

      if (!fs.existsSync(dest)) {

        fs.mkdirSync(dest);
      }
      let parsed_src = path.parse(src);

      let workbook = xlsx.parse(src);

      console.log("parsing excel:", parsed_src.base);

      let settings = parser.parseSettings(workbook);
      //====================add========================
      // let parsed_workbook = parseWorkbook(workbook, dest, headIndex, path.join(dest, parsed_src.name));
      let parsed_workbook = parser.parseWorkbook(workbook, settings);

      serializeWorkbook(parsed_workbook, settings, dest);

      if (config.ts) {
        serializeDTS(dest, parsed_src.name, settings);
      }
    }
  }