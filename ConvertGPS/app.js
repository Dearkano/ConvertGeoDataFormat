'use strict';
var fs = require("fs");
var read_xlsx = require("read_xlsx");
var excelBuffer = "";
var gcoord = require('gcoord');
var xlsx = require('node-xlsx');
var join = require('path').join;

 st();
 async function st (){
    var rs = findSync("./source/");
    var data = [];
    for (var item of rs) {
        data =data.concat(await convert(fs.readFileSync("./source/"+item)));
        console.log(data);
     }
   
    var buffer = xlsx.build([
        {
            name: 'Sheet1',
            data: data
        }
    ]);
    fs.appendFileSync('test2.xlsx', buffer, { 'flag': 'w' });
}
async function convert(excelBuffer) {

    let data = [];
    //返回Promise对象
    var workbook = await read_xlsx.getWorkbook(excelBuffer);
        //获得所有工作簿名称
        var sheetNames = workbook.getSheetNames();
        // console.log(sheetNames);
        //获得名称为Sheet1的工作簿
        var sheet = await workbook.getSheet("Sheet1")
            //获得总行数
            var rowLen = sheet.getRows();
            //获得总列数
            var cellLen = sheet.getColumns();
            //遍历所有单元格
           

            for (var i = 1; i < rowLen; i++) {
                var lat;
                var lng;
                for (var k = 0; k < cellLen; k++) {
                    var cell = sheet.getCell(i, k);
                    //If the cell is empty, it is possible that the cell does not exist return null!
                    if (cell !== null) {
                        //打印单元格内容
                        if (k === 0) {
                            //console.log(cell.getContents());
                        } else if (k === 1) {
                            lat = cell.getContents();
                            //console.log(lat);
                        } else {
                            lng = cell.getContents();
                            //console.log(lng)
                            var result = gcoord.transform([lat, lng], gcoord.BD09, gcoord.WGS84);
                            data.push(result);
                           // console.log(lat, lng);
                          //  console.log(result);
                        }

                    }
                }
            }
          //  console.log("11");
           // console.log(data);
            return data;
   
}
function findSync(startPath) {
    let result = [];
    function finder(path) {
        let files = fs.readdirSync(path);
        files.forEach((val, index) => {
            let fPath = join(path, val);
            let stats = fs.statSync(fPath);
            if (stats.isDirectory()) finder(fPath);
            if (stats.isFile()) result.push(fPath);
        });

    }
    finder(startPath);
    var rs = [];
    for (var item of result) {
        var s = item.split("\\");
        item = s[1];
        rs.push(item);
    }
    //console.log(rs);
    return rs;
}
//let fileNames = findSync('./');