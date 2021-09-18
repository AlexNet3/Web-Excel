////////////////////////////////////////////////////// web server part //////////////////////////////////////////////////////////////
const express = require('express');
const app = express();
const server =  require('http').createServer(app);
const port = 8000
const hostname = 'localhost';
var fs = require('fs');

app.use(express.static('public'));
app.use('/css', express.static(__dirname + '/css'))
app.use('/js', express.static(__dirname + '/js'))
app.use('/img', express.static(__dirname + '/img'))

app.set('views', './');
app.set('view engine', 'ejs');

app.get('', (req, res) => {
    res.render('index')
});

server.listen(port, () => console.info(`App listening on port ${hostname+':'+port}  `))

////////////////////////////////////////////////////// web server part //////////////////////////////////////////////////////////////

///////////////////////////// DB part //////////////////////////////////////////////////////////////////////////////////////////
var mysql = require("mysql");
var con = mysql.createConnection({
    host: "localhost",  
    user: "server",
    password: "adminushka",
    database: "mainDB"
});
///////////////////////////// DB part //////////////////////////////////////////////////////////////////////////////////////////

const io = require('socket.io')(server, { cors: { origin: "*" }})
var formuls = ['SELECTLIST','CONCATENATE','SUBSTITUTE','INDIRECT','ADDRESS','ROUNDUP','ROUND','ISERROR','MATCH','VALUE','ISERR','FIND','TRIM','MID','LEN','ROW','SUM','AND','MIN','MAX','NOT','OR','IF','='];
var alphaBet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];     
var excelData = {};  
var replaceall = require("replaceall");
var newData = {};

var registerMainCell = {};
var registerCalcCell = {};
var registerChangesCell = {};
var mapExecute = {};
var loadDataIndex = {};
var executeList = {};
var executeNow = {};
var lastChanged = {};

io.on('connection', socket => {
    excelData[socket.id+''] = {};
    excelData[socket.id+'']['main'] =  {};
    excelData[socket.id+'']['calc'] =  {};

    newData[socket.id+''] = {};
    newData[socket.id+'']['main'] =  {};
    newData[socket.id+'']['calc'] =  {};

    executeNow[socket.id+''] = '';
    registerMainCell[socket.id+''] = {};
    registerCalcCell[socket.id+''] = {};
    mapExecute[socket.id+''] = '';
    registerChangesCell[socket.id+''] = {};
    loadDataIndex[socket.id+''] = '';
    executeList[socket.id+''] = [];
    lastChanged[socket.id+''] = [];

    socket.on('insertInBigData', data => {
        if(excelData[socket.id+''][data.table][data.cell]){
            excelData[socket.id+''][data.table][data.cell][data.type] = data.data;
        } else {
            var subType;
                if(data.type == 'front'){
                    subType = 'back';
                } else if(data.type == 'back') {
                    subType = 'front';
                }
            var tmpObj = {};
                tmpObj[subType] = '';
                tmpObj[data.type] = data.data;
            excelData[socket.id+''][data.table][data.cell] = tmpObj;
        }
        if(data.typeIns == 'selectlist'){
            socket.emit('cb_change_selList', {'cell1':data.cell, 'cell2':data.tbr, 'type':'wait','tpar':data.table});
        }                                                                               
    })
    socket.on('recalcFunction', async function(data){
        if(excelData[socket.id+''][data.tpar][data.cell1]){
            if(executeNow[socket.id+'']){
                executeList[socket.id+''].push(data.cell1+data.cell2);
            } else {
                registerChangesCell[socket.id+''] = {};   
                registerChangesCell[socket.id+''][data.cell1+data.cell2] = {};
                mapExecute[socket.id+''] = data.cell1+data.cell2;
                        
                executeNow[socket.id+''] = data.cell1+data.cell2;
                console.time('startExecute');
                var backD = excelData[socket.id+''][data.tpar][data.cell1]['back'];
                var ckSelectList = backD.indexOf('SELECTLIST');
                if(backD && ckSelectList == -1){
                    var cell1 = backD;
                    var cell2 = [];
                        cell2.push(data.cell1);
                        cell2.push(data.cell2);
                    var res = await calсBackData(cell1,cell2,socket.id+'');
                } else {
                    var res = await recalcFunction(data.cell1,data.cell2,data.type,socket.id+'');
                }
                for(;;){
                    if(res.type == 'wait'){
                        res = await recalcFunction(res.cell1,res.cell2,res.type,res.socketid);
                    } else if(res.type == 'calcBack'){
                        res = await calсBackData(res.cell1,res.cell2,res.socketid);
                    } else if(res.type == 'Finish'){
                        console.timeEnd('startExecute');
                        if(newData[socket.id+'']['main'] || newData[socket.id+'']['calc']){
                            socket.emit('cb_loadFrontData', newData[socket.id+'']);

                            newData[socket.id+'']['main'] = {};
                            newData[socket.id+'']['calc'] = {};
                        }
                        break
                    }
                } 
            }
        }     
    })
    socket.on('disconnect', function() {
      delete excelData[socket.id+''];
      delete registerMainCell[socket.id+''];
      delete registerCalcCell[socket.id+''];
      delete registerChangesCell[socket.id+''];
      delete mapExecute[socket.id+''];
      delete loadDataIndex[socket.id+''];
      delete executeList[socket.id+''];
      delete executeNow[socket.id+''];
    })




    ///////////////////////////////////////////////////////////DB part
    function loadStyleCells(socketid){
    con.query("SELECT s.style , c.`char`,c.`number`, c.`table` FROM style_cell s INNER JOIN cells c ON c.id = s.fk_id AND s.style IS NOT NULL", function (err, rows) {
            if (err)
            throw err;
            io.sockets.connected[socketid].emit('setStyleCells',rows);
        })
};

function insertPartCell(id,type){
    if(type == 'calc'){
        con.query("INSERT INTO `calc_excel`.`calc_back` (`id`) VALUES (?);",  [id])
        con.query("INSERT INTO `calc_excel`.`calc_front` (`id`) VALUES (?);",  [id])
    } else if(type == 'main'){
        con.query("INSERT INTO `calc_excel`.`main_back` (`id`) VALUES (?);",  [id])
        con.query("INSERT INTO `calc_excel`.`main_front` (`id`) VALUES (?);",  [id])
    }
        con.query("INSERT INTO `calc_excel`.`style_cell` (`fk_id`) VALUES (?);",  [id])
}

function loadMainBack(socketid){
    con.query("SELECT c.`char`, c.`number`, c_f.`value` AS 'backData'  FROM cells c INNER JOIN main_back c_f ON c_f.id = c.id WHERE c_f.`value` IS NOT NULL AND c_f.`value` NOT LIKE '%SELECTLIST%'", async function (err, rows) {
                if (err)
                throw err;

                for(var item of rows){
                    if(excelData[socketid]['main'][item.char+item.number]){
                        excelData[socketid]['main'][item.char+item.number]['back'] = item.backData;
                    } else {
                        var tmpObj = {};
                            tmpObj['back'] = item.backData;
                            tmpObj['front'] = '';
                        excelData[socketid]['main'][item.char+item.number] = tmpObj;
                    }
                }
                loadDataIndex[socketid] = 'mainContentExcel';
                        var first = false;
                        for(var item of rows){
                            if(item.backData){
                                var cell = [];
                                    cell.push(item.char);
                                    cell.push(item.number);
                                var tableEs = 'main';
                                    await registerFunctionCell(item.backData,tableEs,cell,socketid);
                            }
                        }
                        io.sockets.connected[socketid].emit('cb_loadFrontData', excelData[socketid]);  
    })
}

function loadCalcBack(socketid){
                con.query("SELECT c.`char`, c.`number`, c_f.`value` AS 'backData'  FROM cells c INNER JOIN calc_back c_f ON c_f.id = c.id WHERE c_f.`value` IS NOT NULL", async function (err, rows) {
                    if (err)
                    throw err;

                    for(var item of rows){
                        if(excelData[socketid]['calc'][item.char+item.number]){
                            excelData[socketid]['calc'][item.char+item.number]['back'] = item.backData;
                        } else {
                            var tmpObj = {};
                                tmpObj['back'] = item.backData;
                                tmpObj['front'] = '';
                            excelData[socketid]['calc'][item.char+item.number] = tmpObj;
                        }
                    }

                    loadDataIndex[socketid] = 'calcContentExcel';
                        var first = false;
                        for(var item of rows){
                            if(item.backData){
                                var cell = [];
                                    cell.push(item.char);
                                    cell.push(item.number);
                                var tableEs = 'calc';
                                    await registerFunctionCell(item.backData,tableEs,cell,socketid);
                            }
                        }
                        loadMainBack(socketid);
                })
        } 

    socket.on('addCell', data => {
        con.query("INSERT INTO `calc_excel`.`cells` (`char`, `number`, `table`) VALUES (?, ?,'main');", [data.char,data.number], function (err, rows) {
            if (err)
            throw err;
            var id = rows.insertId;
            
            insertPartCell(id,'main')
        })
        
        con.query("INSERT INTO `calc_excel`.`cells` (`char`, `number`, `table`) VALUES (?, ?,'calc');", [data.char,data.number], function (err, rows) {
            if (err)
            throw err;
            var id = rows.insertId;
            insertPartCell(id,'calc')
        })  
    })

    socket.on('loadFrontDataExcel', data => {
        con.query("SELECT c.`char`, c.`number`, c_f.`value` AS 'frontData'  FROM cells c INNER JOIN calc_front c_f ON c_f.id = c.id WHERE c_f.`value` IS NOT NULL", function (err, rows) {
            if (err)
            throw err;

            for(var item of rows){
                var tmpObj = {};
                    tmpObj['front'] = item.frontData;
                    tmpObj['back'] = '';
                excelData[socket.id+'']['calc'][item.char+item.number] = tmpObj;
            }
            startLoadSelectList()
        })
        function startLoadSelectList(){
                con.query("SELECT c.`char`, c.`number`, c_b.`value` AS 'backData' ,c_f.`value` AS 'frontData'  FROM cells c INNER JOIN main_back c_b ON c_b.id = c.id INNER JOIN main_front c_f ON c_f.id = c.id WHERE c_b.`value` LIKE '%SELECTLIST%'", async function (err, rows) {
                    if (err)
                    throw err;
                    for(var item of rows){
                        var tmpObj = {};
                            tmpObj['back'] = item.backData;
                            if(item.frontData){
                                tmpObj['front'] = item.frontData; 
                            } else {
                                tmpObj['front'] = '';
                            }
                        excelData[socket.id+'']['main'][item.char+item.number] = tmpObj;

                        var cell = [];
                            cell.push(item.char);
                            cell.push(item.number);
                        var tableEs = 'main';
                        await registerFunctionCell(item.backData,tableEs,cell,socket.id+'');
                        var cellArr = [];
                            cellArr.push(item.char+item.number)
                            cellArr.push('cellsMain')

                      var res = await calсBackData(item.backData,cellArr,socket.id+'');
                    }
                    loadStyleCells(socket.id+'');
                    console.time('loadData');
                    loadFrontMain();   
                })
            }
        function loadFrontMain(){
                con.query("SELECT c.`char`, c.`number`, c_f.`value` AS 'frontData'  FROM cells c INNER JOIN main_front c_f ON c_f.id = c.id WHERE c_f.`value` IS NOT NULL", function (err, rows) {
                    if (err)
                    throw err;

                    for(var item of rows){
                        var tmpObj = {};
                            tmpObj['front'] = item.frontData;
                            tmpObj['back'] = '';
                        excelData[socket.id+'']['main'][item.char+item.number] = tmpObj;
                    }
                    loadCalcBack(''+socket.id);
                })
        }

    })
    /////////////////////////DB part
})

function registerFunctionCell(data,table,funcCell,socketId){
        var group = funcCell[0];
        var line = funcCell[1];
        var id;
        if(table == 'main'){
            id = 'cellsMain';
        } else if(table == 'calc'){
            id = 'cellsCalc';
        }
    var funcAdress = group+line+'_'+id;
    var count = 0;
    var openQuotation = false;
    var cells = [];
    var openCell = '';
    for(var i = 0;i < data.length;i++){
        if(data.charAt(i) == '{' && openQuotation == false){
            count += 1;
            openCell = cells.length;

            var arr = [];
                arr.push(i);
            cells.push(arr);    

        } else if(data.charAt(i) == '}' && openQuotation == false){
            cells[openCell].push(i);
        } else if(data.charAt(i) == '"' && openQuotation == false){
            openQuotation = true;
        } else if(data.charAt(i) == '"' && openQuotation == true){
            openQuotation = false;
        }
    }
    if(count > 0){
        for(var i = 0;i < cells.length;i++){
            var start = cells[i][0]+1;
            var finish = cells[i][1];

            var cell = getCellAndRange(data,start,finish,table);
                for(var i2 = 0;i2 < cell.length;i2++){
                    if(cell[i2][0].indexOf('.') > -1){
                        var ntm = cell[i2][0].split('.');
                        cell[i2][0] = ntm[1];
                        cell[i2][2] = ntm[0];
                    }
                    if(cell[i2][2] == 'main'){
                            if(Array.isArray(registerMainCell[socketId][cell[i2][0]+cell[i2][1]])){
                                if(registerMainCell[socketId][cell[i2][0]+cell[i2][1]].indexOf(funcAdress) == -1){
                                    registerMainCell[socketId][cell[i2][0]+cell[i2][1]].push(funcAdress);
                                }
                            } else {
                                var tmpArr = [];
                                tmpArr.push(funcAdress);
                                registerMainCell[socketId][cell[i2][0]+cell[i2][1]] = tmpArr;
                            }
                    } else if(cell[i2][2] == 'calc'){
                            if(Array.isArray(registerCalcCell[socketId][cell[i2][0]+cell[i2][1]])){
                                if(registerCalcCell[socketId][cell[i2][0]+cell[i2][1]].indexOf(funcAdress) == -1){
                                    registerCalcCell[socketId][cell[i2][0]+cell[i2][1]].push(funcAdress);
                                }
                            } else {
                                var tmpArr = [];
                                tmpArr.push(funcAdress);
                                registerCalcCell[socketId][cell[i2][0]+cell[i2][1]] = tmpArr;
                            }
                    }
                }   
        }
        return true;
    }
}

async function calсBackData(data,cellName,socketid){
    var arrayCalc = {};
    var index = -1;
    var indexOpen = -1;
    var parent = -1;
    var open = -1;
    var quotes = false;

    var str = data;
        for(var i = 0; i < str.length;i++){
            var char = str.charAt(i);
            if(char == '(' && quotes == false){
                if(parent == -1){
                    var arr = {};
                        var finish = getFinishFunction(i,data);

                        var format = getFormatData(i,data);

                        var arrItem = {};
                            arrItem['index'] = '0';
                            arrItem['char'] = i;
                            arrItem['function'] = format+data.substring(i,finish);

                        indexOpen = Object.keys(arr).length; 
                        arr[indexOpen]= arrItem; //<-- psuh array

                    arrayCalc[0] = arr;

                        index = 0;
                        open = 0;
                        parent = 0;

                } else if(open > -1){
                    parent = open;
                    index = Number(Object.keys(arrayCalc[parent]).length) - 1; 
                    open++;

                    if(arrayCalc[open]){
                        var finish = getFinishFunction(i,data);

                        var format = getFormatData(i,data);

                        var arrItem = {};
                            arrItem['index'] = parent+':'+index;
                            arrItem['char'] = i;
                            if(format){
                                arrItem['function'] = format+data.substring(i,finish);
                            } else {
                                arrItem['function'] = data.substring(i,finish);
                            }

                        indexOpen = Object.keys(arrayCalc[open]).length;
                        arrayCalc[open][indexOpen] = arrItem;

                        arrayCalc[parent][index]['var-'+open+'-'+indexOpen] = false;

                        var oldFunction = arrayCalc[parent][index]['function'];
                            if(format){
                                var func = format+data.substring(i,finish);
                            } else {
                                var func = data.substring(i,finish);
                            }
                        
                            oldFunction = oldFunction.replace(func, '|'+'var-'+open+'-'+indexOpen+'|');
                            arrayCalc[parent][index]['function'] = oldFunction;

                    } else {
                        var finish = getFinishFunction(i,data);

                        var format = getFormatData(i,data);

                        var arr = {};
                        var arrItem = {};
                            arrItem['index'] = parent+':'+index;
                            arrItem['char'] = i;
                            if(format){
                                arrItem['function'] = format+data.substring(i,finish);
                            } else {
                                arrItem['function'] = data.substring(i,finish);
                            }
                            

                            indexOpen = Object.keys(arr).length
                            arr[indexOpen]= arrItem;
                        arrayCalc[open] = arr;

                        arrayCalc[parent][index]['var-'+open+'-'+indexOpen] = false;

                        var oldFunction = arrayCalc[parent][index]['function'];
                        if(format){
                            var func = format+data.substring(i,finish);
                        } else {
                            var func = data.substring(i,finish);
                        }
                            oldFunction = oldFunction.replace(func, '|'+'var-'+open+'-'+indexOpen+'|');
                            arrayCalc[parent][index]['function'] = oldFunction;
                    }
                }       
            } else if(char == ')' && quotes == false){
                if(open > 0){
                    index = Number(Object.keys(arrayCalc[parent]).length) - 1;

                    parent--;
                    open--;

                    if(parent < 0 )parent = 0;

                    index = Number(Object.keys(arrayCalc[parent]).length) - 1;
                }
            } else if(char == '"'){
                if(quotes == false){
                    quotes = true;
                } else {
                    quotes = false;
                }
            }
        }
        return await calcFunction(arrayCalc,cellName,socketid);
}

function charСount(str, letter){
    var letterCount = 0;
    for (var position = 0; position < str.length; position++){
        if (str.charAt(position) == letter){
            letterCount += 1;
          }
    }
    return letterCount;
}


function getFinishFunction(start,data){
    var open = 1;
    var finish = 0;
    var gal = 0;
    for(var i = start+1;i < data.length;i++){
        if(data.charAt(i) == '(' && gal == 0){
            open++;
        } else if(data.charAt(i) == ')' && gal == 0){
            open--;
        } else if(data.charAt(i) == '"' && gal == 0){
            gal = 1;
        } else if(data.charAt(i) == '"' && gal == 1){
            gal = 0;
        }
        if(open == 0){
            finish = i;
            break; 
        }
    }
    return finish+1;
}

function getFormatData(i,data){
    for(var item of formuls){
        var ckFormul = data.substring(i-(item.length), i);
            if(ckFormul == item){
                return item;
            }
    }
}

async function calcFunction(data,cellName,socketid){
    var indexCalc = {
    nr: 0,
    index: 0  
    };

    indexCalc['nr'] = Number(Object.keys(data).length) - 1;
    indexCalc['index'] = Number(Object.keys(data[indexCalc['nr']]).length) - 1;

    var nr = indexCalc['nr'];
    var index = indexCalc['index'];

    var func = data[indexCalc['nr']][indexCalc['index']]['function'];
        var tb;
        if(cellName[1] == 'cellsCalc'){
            tb = 'calc';
        } else if(cellName[1] == 'cellsMain') {
            tb = 'main';
        }
    var res = await calcFormul(index, nr, func, tb,cellName,socketid,data);
    for(;;){
        if(res.type == 'calcFormul'){
          res = await calcFormul(res.index, res.nr, res.func, res.table, res.cellName, res.socketid, res.data)
        } else if(res.type == 'calcBack' || res.type == 'Finish'){
            return res;
            break;
        }
    }
}

function cb_calcFormul(result , nr, index, table,cellName,socketid,data) {
    var parent = data[nr][index]['index'];
        parent = parent.split(':');
        var indexParent = parent[0];
        var nrParent = parent[1];

        data[indexParent][nrParent]['var-'+nr+'-'+index] = result;
        if(data[indexParent][nrParent]['function']){
            data[indexParent][nrParent]['function'] = data[indexParent][nrParent]['function'].replace('|var-'+nr+'-'+index+'|', result);
        }
    
        if((nr > 0 || nr == 0) && index > 0){
            index = index-1;
        } else if(index == 0 && nr > 0){
            nr = nr-1;
            index = Number(Object.keys(data[nr]).length)-1;
        }

        if(nr >= 0 && index >= 0){
            var func = data[nr][index]['function'];
            var dts = {}
           return {'index':index,'nr': nr,'func': func,'table': table,'cellName': cellName,'socketid': socketid,'data' :data, 'type':'calcFormul'}
        }
}

async function calcFormul(index, nr, func, table,cellName,socketid,dataCalcs){
    var useFunc = {};
        var func = func;
        var typeFomul;
        
        for(var item of formuls){
            var ck = func.indexOf(item);
            if(ck > -1){
                typeFomul = item;
                break;
            }
        }
        if(!typeFomul){
            if(useFunc[table+cellName[0]]){
                useFunc[table+cellName[0]] += 1;
            } else {
                useFunc[table+cellName[0]] = 1;
            }

            typeFomul = 'simpleCalc';
        }
        /////////////////////////////////////////////////////значение\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        if(typeFomul == 'FIND'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('FIND','');
            func = func.substring(1,func.length-1);

            func = func.split('|');
            var srcStr = func[0];
                srcStr = srcStr.replace(/"/g,'');
            var str = func[1];
            
            var result = getValueCellInstring(str,table,'',socketid);
                if(func[2]){
                    func[2] =  getValueCellInstring(func[2],table,'',socketid)
                    func[2] = calcEval(func[2]);
                    result = result.indexOf(srcStr,func[2])+1;
                } else {
                    result = result.indexOf(srcStr)+1;
                }
                if(result > 0){
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell(result,cellName,socketid);
                    } else {
                        return await cb_calcFormul(result,nr,index,table,cellName,socketid,dataCalcs);
                    }
                } else {
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell('#error#',cellName,socketid);
                    } else {
                        return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                    }
                }
        } else if(typeFomul == 'VALUE'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('VALUE','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);
               if(func.indexOf('#error#') > -1){
                    var result = '#error#';
               } else {
                    var result = func.replace(/[^0-9]/g, '');
                    var str = func.indexOf(',');
                        result = Number(result);
                        if(str > -1){
                            func = func.replace(',', '.');
                            result = Number(func);
                            //if(func[1].length != 3){
                              //  result = '#error#';
                            //}
                        }
               }
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(result,cellName,socketid);
                } else {
                    return await cb_calcFormul(result,nr,index,table,cellName,socketid,dataCalcs);
                }
            
        } else if(typeFomul == 'MID'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('MID','');
            func = func.substring(1,func.length-1);   

            func = func.split(';');
            var main = func[0];
                main = getValueCellInstring(main,table,'',socketid);

            var startSplit = func[1];
                startSplit = getValueCellInstring(startSplit,table,'',socketid);
                startSplit = calcEval(startSplit);

            var continueSplit = func[2];
                continueSplit = getValueCellInstring(continueSplit,table,'',socketid);
                continueSplit = calcEval(continueSplit);

            var ck1 = checkToString(main),
                ck2 = checkToString(startSplit),
                ck3 = checkToString(continueSplit);


            if(!ck2 && !ck3 && main.length > 0){
                startSplit = startSplit-1;
                continueSplit = continueSplit + startSplit;
                if(startSplit > (main.length)-1){
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell('#error#',cellName,socketid);
                    } else {
                        return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                    }
                } else {
                    var result = main.substring(startSplit, continueSplit);
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell(result,cellName,socketid);
                    } else {
                        return await cb_calcFormul(result,nr,index,table,cellName,socketid,dataCalcs);
                    }
                }
            } else {
                if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell('#error#',cellName,socketid);
                    } else {
                        return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                    }
            }

        } else if(typeFomul == 'LEN'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('LEN','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            var result = func.length;
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(result,cellName,socketid);
                } else {
                    return await cb_calcFormul(result,nr,index,table,cellName,socketid,dataCalcs);
                }
        } else if(typeFomul == 'SUM'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('SUM','');
            func = func.substring(1,func.length-1);
            var sum = 0;

            var ck_err = getValueCellRange(func,table,socketid);
            if(ck_err.indexOf('#error#') > -1){
                sum = '#error#';
            } else {
                funcM = func.split(';');
                for(var i2 = 0;i2 < funcM.length;i2++){
                    if(funcM[i2].indexOf(':') > -1){
                        func = getValueCellRange(funcM[i2],table,socketid);

                        for(var i = 0;i < func.length;i++){
                            var val = Number(func[i]);
                            if(!isNaN(val)){
                                sum += val;
                            }  
                        }
                    } else {
                        func = getValueCellInstring(funcM[i2],table,'',socketid);
                        var val = Number(func);
                        if(!isNaN(val)){
                            sum += val;
                        } 
                    }
                }
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(sum,cellName,socketid);
            } else {
                return await cb_calcFormul(sum,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'MAX'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('MAX','');
            func = func.substring(1,func.length-1);
            if(func.indexOf(':') > -1){
                func = getValueCellRange(func,table,socketid);
            } else if (func.indexOf(';') > -1){
                func = getValueCellInstring(func,table,'',socketid);
                func = func.split(';');
            }
            var max = 0;
            for(var i = 0;i < func.length;i++){
                var val = Number(func[i]);
                if(!isNaN(val)){
                    if(val >= max){
                        max = val;
                    }
                } 
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(max,cellName,socketid);
            } else {
                return await cb_calcFormul(max,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'MIN'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('MIN','');
            func = func.substring(1,func.length-1);
            func = getValueCellRange(func,table,socketid);

            var min = 0;
            for(var i = 0;i < func.length;i++){
                var val = Number(func[i]);
                if(func[i] != '' && !isNaN(val) && val != 0){
                    if(i == 0){
                        min = val
                    } else {
                        if(val <= min){
                            min = val;
                        }
                    }
                } 
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(min,cellName,socketid);
            } else {
                return await cb_calcFormul(min,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'NOT'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('NOT','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            var res;
            if(Number(func) == 0){
                res = 'TRUE';
            } else if(Number(func) > 0){
                res = 'FALSE';
            } else if(func == 'FALSE'){
                res = 'TRUE';
            } else if(func == 'TRUE'){
                res = 'FALSE';
            } else if(isNaN(Number(func))){
                res = '#error#';
            }

            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(res,cellName,socketid);
            } else {
                return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'ISERR' || typeFomul == 'ISERROR'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace(typeFomul,'');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            if(func == '#error#'){
                func = 'TRUE';
            } else {
                func = 'FALSE';
            }
            
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'SUBSTITUTE'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('SUBSTITUTE','');
            func = func.substring(1,func.length-1);

            func = replaceall('"','', func);

            func = func.split(';');
            if(func.length == 3){
                for(var i = 0;i < func.length;i++){
                    var ck = func[i].indexOf('{');
                    var ck2 = func[i].indexOf('}');
                    if(ck > -1 && ck2 > -1){
                        func[i] = await getValueCellInstring(func[i],table,'',socketid);
                    }
                }
                
                func = replaceall(func[1],func[2], func[0]);
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(func,cellName,socketid);
                } else {
                    return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
                }
            } else {
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell('#error#',cellName,socketid);
                } else {
                    return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                }
            }
        } else if(typeFomul == 'CONCATENATE'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('CONCATENATE','');
            func = func.substring(1,func.length-1);

            func = getValueCellInstring(func,table,'',socketid);
            func = func.split(';');
            var structure = true;
                if(structure == true){
                    var txt = '';
                    for(var i = 0;i < func.length;i++){
                        if(func[i])txt += func[i];
                    }
                    
                    txt = replaceall('"','', txt);
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell(txt,cellName,socketid);
                    } else {
                        return await cb_calcFormul(txt,nr,index,table,cellName,socketid,dataCalcs);
                    }
                }
        } else if(typeFomul == 'ROUNDUP'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('ROUNDUP','');
            func = func.substring(1,func.length-1);

            func = func.split(';');
            if(func[1]){
                func[0] = getValueCellInstring(func[0],table,'number',socketid);
                func[1] = getValueCellInstring(func[1],table,'number',socketid);
                func[0] = calcEval(func[0]);
                if(Number(func[0]) == 0){
                    func = 0;
                } else {
                    func = roundup(func[0],Number(func[1]))
                }
            } else {
                func[0] = getValueCellInstring(func[0],table,'',socketid);
                func[0] = calcEval(func[0]);
                if(func[0] == 0){
                    func = 0;
                } else {
                    func = roundup(func[0])
                }
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'ROUND'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('ROUND','');
            func = func.substring(1,func.length-1);

            func = func.split(';');
            if(func[1]){
                func[0] = getValueCellInstring(func[0],table,'number',socketid);
                func[1] = getValueCellInstring(func[1],table,'number',socketid);
                func[0] = calcEval(func[0]);
                func = excelRound(func[0],func[1])
            } else {
                func[0] = getValueCellInstring(func[0],table,'number',socketid);
                func[0] = calcEval(func[0]);
                func = excelRound(func[0])
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'AND'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('AND','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            func = replaceall('"','',func);
            func = func.split(';');

            var AllTrue = true;
            for(var i = 0;i < func.length;i++){
                var ck = searchOperation(func[i]);
                if(ck){
                    if(func[i].indexOf('$+$') > -1 || 
                       func[i].indexOf('$-$') > -1 || 
                       func[i].indexOf('$*$') > -1 ||
                       func[i].indexOf('$/$') > -1){
                        if(index == 0 && nr == 0){
                            return await applicateFrontDataInCell('#error#',cellName,socketid);
                        } else {
                            return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                        }
                        AllTrue = false;
                        break;
                    } else {
                        var result = getResultFuncAND(ck,func[i],table,socketid);
                        if(result == 'TRUE' || result == 'FALSE'){
                            func[i] = result;
                        } else {
                            if(index == 0 && nr == 0){
                                return await applicateFrontDataInCell('#error#',cellName,socketid);
                            } else {
                                return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                            }
                            AllTrue = false;
                            break;
                        }
                    }
                } else {
                    if(index == 0 && nr == 0){
                        return await applicateFrontDataInCell('#error#',cellName,socketid);
                    } else {
                        return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                    }
                    AllTrue = false;
                    break;
                }
            }
            if(AllTrue == true){
                var res = 'TRUE';
                for(var i = 0;i < func.length;i++){
                    if(func[i] == 'FALSE'){
                        res = 'FALSE';
                        break;
                    }
                }
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(res,cellName,socketid);
                } else {
                    return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
                }
            }   
        } else if(typeFomul == 'OR'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('OR','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);
            func = replaceall('"','',func);
            func = func.split(';');

            var AllTrue = false;
            var error = false;
            for(var i = 0;i < func.length;i++){
                if(func[i] == 'TRUE'){
                    AllTrue = true;
                    break;
                } else if(func[i] == 'FALSE'){
                    continue;
                } else {
                    var ck = searchOperation(func[i]);
                    if(ck){
                        if(func[i].indexOf('$+$') > -1 || 
                           func[i].indexOf('$-$') > -1 || 
                           func[i].indexOf('$*$') > -1 ||
                           func[i].indexOf('$/$') > -1){
                            if(index == 0 && nr == 0){
                                return await applicateFrontDataInCell('#error#',cellName,socketid);
                            } else {
                                return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                            }
                            AllTrue = false;
                            error = true;
                            break;
                        } else {
                            var result = getResultFuncAND(ck,func[i],table,socketid);
                            if(result == 'TRUE'){
                                AllTrue = true;
                                break;
                            } else if(result == 'FALSE'){
                                continue;
                            } else {
                                if(index == 0 && nr == 0){
                                    return await applicateFrontDataInCell('#error#',cellName,socketid);
                                } else {
                                    return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                                }
                                AllTrue = false;
                                error = true;
                                break;
                            }
                        }
                    } else {
                        if(index == 0 && nr == 0){
                            return await applicateFrontDataInCell('#error#',cellName,socketid);
                        } else {
                            return await cb_calcFormul('#error#',nr,index,table,cellName,socketid,dataCalcs);
                        }
                        AllTrue = false;
                        error = true;
                        break;
                    }
                }
            }
            if(error == false){
                if(AllTrue == true){
                    var res = 'TRUE';
                } else {
                    var res = 'FALSE';
                }

                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(res,cellName,socketid);
                } else {
                    return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
                }
            }
        } else if(typeFomul == 'TRIM'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('TRIM','');
            func = func.substring(1,func.length-1);

            var ck_and = func.indexOf('&');
            if(ck_and > -1){
                func = func.split('&');
                for(var i = 0;i < func.length;i++){
                    func[i] = getValueCellInstring(func[i],table,'',socketid);
                }
                var data = '';
                for(var i = 0;i < func.length;i++){
                    data += func[i];
                }
                for(;;){
                    if(data.indexOf('  ') > -1){
                        data = replaceall('  ',' ',data);
                    } else {
                        break;
                    }
                }
                if(index == 0 && nr == 0){
                    return await applicateFrontDataInCell(data,cellName,socketid);
                } else {
                    return await cb_calcFormul(data,nr,index,table,cellName,socketid,dataCalcs);
                }
            } else {
                func = getValueCellInstring(func,table,'',socketid);
            }
        } else if(typeFomul == 'ROW'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('ROW','');
            func = func.substring(1,func.length-1);

            var res = getRowCell(func);
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(res,cellName,socketid);
            } else {
                return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'ADDRESS'){

            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('ADDRESS','');
            func = func.substring(1,func.length-1);

            var res = exAdressFunc(func,table,socketid);
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(res,cellName,socketid);
            } else {
                return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'INDIRECT'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('INDIRECT','');
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'MATCH'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('MATCH','');
            func = func.substring(1,func.length-1);
            var res = matchFuncExcel(func,table,socketid);

            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(res,cellName,socketid);
            } else {
                return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'IF'){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('IF','');
            func = func.substring(1,func.length-1);

            var res = funcIfExcel(func,table,socketid);
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(res,cellName,socketid);
            } else {
                return await cb_calcFormul(res,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == 'SELECTLIST'){
            var fn = '=('+func+')'
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.replace('SELECTLIST','');
            func = func.substring(1,func.length-1);
            var res = getValueCellRange(func,table,socketid);
            return await applicateFrontListInCell(res,cellName,socketid,fn);
        } else if(typeFomul == 'simpleCalc'){
            func = getValueCellInstring(func,table,'number',socketid);
            var quotesOpen = false;
            var calcCk = false;
            for(var i = 0;i < func.length;i++){
                if(func.indexOf('$+$') > -1 || func.indexOf('$-$') > -1 || func.indexOf('$*$') > -1 || func.indexOf('$/$') > -1 || func.indexOf('$=$') > -1){
                    if(quotesOpen == false){
                        calcCk = true;
                    }
                } else if(func.charAt(i) == '"' && quotesOpen == false){
                    quotesOpen == true;
                } else if(func.charAt(i) == '"' && quotesOpen == true){
                    quotesOpen == false;
                }
            }
            if(calcCk == true){
                func = calcEval(func);
            }
            if(func){
                func = func+'';
                var ck_and = func.indexOf('&');
               if(ck_and > -1){
                  func = func.split('&');
                  var data = '';
                  for(var i = 0;i < func.length;i++){
                     data += func[i];
                  }
                  for(;;){
                     if(data.indexOf('  ') > -1){
                        data = replaceall('  ',' ',data);
                    } else {
                        break;
                     }
                  }
                  data = replaceall('"','',data);
                  func = data;
               }
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        } else if(typeFomul == '='){
            if(func.charAt(0) == '='){
                func = func.substring(1);
            }
            func = func.substring(1,func.length-1);
            func = getValueCellInstring(func,table,'',socketid);

            var quotesOpen = false;
            var calcCk = false;
            for(var i = 0;i < func.length;i++){
                if(func.indexOf('$+$') > -1 || func.indexOf('$-$') > -1 || func.indexOf('$*$') > -1 || func.indexOf('$/$') > -1 || func.indexOf('$=$') > -1){
                    if(quotesOpen == false){
                        calcCk = true;
                    }
                } else if(func.charAt(i) == '"' && quotesOpen == false){
                    quotesOpen == true;
                } else if(func.charAt(i) == '"' && quotesOpen == true){
                    quotesOpen == false;
                }
            }
            if(calcCk == true){
                func = calcEval(func);
            }
            if(func){
                func = func+'';
                var ck_and = func.indexOf('&');
               if(ck_and > -1){
                  func = func.split('&');
                  var data = '';
                  for(var i = 0;i < func.length;i++){
                     data += func[i];
                  }
                  for(;;){
                     if(data.indexOf('  ') > -1){
                        data = replaceall('  ',' ',data)
                    } else {
                        break;
                     }
                  }
                  data = replaceall('"','',data)
                  func = data;
               }
            }
            if(index == 0 && nr == 0){
                return await applicateFrontDataInCell(func,cellName,socketid);
            } else {
                return await cb_calcFormul(func,nr,index,table,cellName,socketid,dataCalcs);
            }
        }
/////////////////////////////////////////////////////значение\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
}

async function applicateFrontDataInCell(data,cellName,socketid){
    data = data+'';
    var table;
    if(cellName[1] == 'cellsMain'){
        table = 'main';
    } else if(cellName[1] == 'cellsCalc'){
        table = 'calc';
    }

        if(cellName){
            var num = cellName[0].replace(/[^0-9]/g, '');
            var char = cellName[0].indexOf(num);
                char = cellName[0].substring(0, char);
            var table;
                if(cellName[1] == 'cellsMain'){
                    table = 'main';
                } else if(cellName[1] == 'cellsCalc'){
                    table = 'calc';
                }
               
                if(excelData[socketid][table][char+num]){
                   var frontData = excelData[socketid][table][char+num]['front'];
                    var ckLast =  ckLastChange(socketid,cellName[0]+cellName[1]);
                   if(frontData != data || frontData == data && ckLast == true){
                        if(newData[socketid][table][cellName[0]]){
                            newData[socketid][table][cellName[0]]['front'] = data;
                        } else {
                            var tmpObj = {};
                                tmpObj['back'] = '';
                                tmpObj['front'] = data;
                                    newData[socketid][table][cellName[0]] = tmpObj;
                                }

                        excelData[socketid][table][char+num]['front'] = data;

                       var res = await recalcFunction(cellName[0],cellName[1],'wait',socketid);
                       for(;;){
                           if(res.type == 'wait'){
                                res = await recalcFunction(res.cell1,res.cell2,res.type,res.socketid);
                           } else if(res.type == 'calcBack' || res.type == 'Finish'){
                                return res;
                                break
                           }
                       }
                   } else {
                        var res = await recalcFunction(cellName[0],cellName[1],'next',socketid)
                        for(;;){
                           if(res.type == 'wait'){
                                res = await recalcFunction(res.cell1,res.cell2,res.type,res.socketid);
                           } else if(res.type == 'calcBack' || res.type == 'Finish'){
                                return res;
                                break
                           }
                       }
                   }
                } else {
                    var tmpObj = {};
                        tmpObj['front'] = data;
                        tmpObj['back'] = '';
                    excelData[socketid][table][char+num] = tmpObj;
                    var res = await recalcFunction(cellName[0],cellName[1],'wait',socketid);
                     for(;;){
                           if(res.type == 'wait'){
                                res = await recalcFunction(res.cell1,res.cell2,res.type,res.socketid);
                           } else if(res.type == 'calcBack' || res.type == 'Finish'){
                                return res;
                                break
                           }
                       }
                }
        }
}

function ckLastChange(socketid,cell){
    lastChanged[socketid]
    var mapArr = mapExecute[socketid].split('_');
    if(mapArr.length == 1){
        lastChanged[socketid] = cell;
        return true
    } else if(mapArr.length > 1 && lastChanged[socketid] == ''){
        lastChanged[socketid] = cell;
        return true
    } else if(mapArr.length > 1 && lastChanged[socketid]){
        if(mapArr[mapArr.length - 2]){
            if(mapArr[mapArr.length - 2] == lastChanged[socketid]){
                lastChanged[socketid] = cell;
                return true
            } else {
                lastChanged[socketid] = '';
                return false
            }
        } else {
            lastChanged[socketid] = '';
            return false
        }
    }
}

function applicateFrontListInCell(data, cellName,socketid,fn){
    var table;
    if(cellName[1] == 'cellsMain'){
        table = 'main';
    } else if(cellName[1] == 'cellsCalc'){
        table = 'calc';
    }
    if(excelData[socketid][table][cellName[0]]){
        if(excelData[socketid][table][cellName[0]]['front'] == ''){
            excelData[socketid][table][cellName[0]]['front'] = data[0];
        }
    } else {
        var tmpObj = {};
            tmpObj['back'] = '';
            tmpObj['front'] = data[0];
        excelData[socketid][table][cellName[0]] = tmpObj;
    }
    io.sockets.connected[socketid].emit('applicateFrontListInCell', {data, cellName,socketid, bigData:excelData[socketid], deff: excelData[socketid][table][cellName[0]]['front'], backDD: fn });
    return {'type':'Finish'};
}

function funcIfExcel(data,table,socketid){
    data = data.split(';');
    var condition = data[0];
    var trueResult = data[1];
    var elseResult = '';
    if(data[2])elseResult = data[2];

    var ck_1 =  getValueCellInstring(condition,table,'',socketid);
    var ck_2 =  getValueCellInstring(trueResult,table,'',socketid);
    var ck_3 =  getValueCellInstring(elseResult,table,'',socketid);

    if(ck_1.indexOf('#error#') > -1){
        return '#error#';
    } else {
        /////check condition /////
        var equals = ['<>','=<','=>','<=','>=','=','>','<'];
        var registerEqual = false;
        var evalCondition;
        for(var i = 0;i < equals.length;i++){
            if(condition.indexOf(equals[i]) > -1){
                registerEqual = equals[i];
                break;
            }
        }
        if(registerEqual){
            condition = condition.split(registerEqual);
            condition[0] = getValueCellInstring(condition[0],table,'',socketid);
            condition[1] = getValueCellInstring(condition[1],table,'',socketid);
            condition[0] = replaceall('"','',condition[0]);
            condition[1] = replaceall('"','',condition[1]);
            evalCondition = manualEvalBool(condition[0],condition[1],registerEqual);
        } else {
            condition = getValueCellInstring(condition,table,'',socketid);
            if(condition == 'TRUE' || condition != ''  && condition.indexOf('#error#') < 0 && condition != 'FALSE' && condition != '0'){
                evalCondition = true;
            } else if(condition == 'FALSE'){
                evalCondition = false;
            } else {
                evalCondition = false;
            }
        }
        if(evalCondition == true){
            if(ck_2.indexOf('#error#') > -1){
                return '#error#';
            } else {
                trueResult = replaceall('&','',trueResult);
                trueResult = replaceall('"','',trueResult);
                trueResult = getValueCellInstring(trueResult,table,'',socketid);

                if(trueResult.indexOf('$-$') > -1 ||trueResult.indexOf('$+$') > -1 ||
                    trueResult.indexOf('$/$') > -1 ||trueResult.indexOf('$*$') > -1 ){
                    trueResult = calcEval(trueResult);
                }

                return trueResult;
            }
        } else if(evalCondition == false){
            if(ck_3.indexOf('#error#') > -1){
                return '#error#';
            } else {
                if(elseResult){
                    elseResult = replaceall('&','',elseResult);
                    elseResult = replaceall('"','',elseResult);
                    elseResult = getValueCellInstring(elseResult,table,'',socketid);
                    if(elseResult.indexOf('$-$') > -1 ||elseResult.indexOf('$+$') > -1 ||
                       elseResult.indexOf('$/$') > -1 ||elseResult.indexOf('$*$') > -1 ){
                       elseResult = calcEval(elseResult);
                    }
                    return elseResult;
                } else {
                    return 'FALSE';
                }
            }
        }
    }
}

function manualEvalBool(val,val2,eq){
    if(eq != '='){
        var tst = Number(val);
        var tst2 = Number(val2);
        if(!isNaN(tst)){
            val = Number(val);
        }
        if(!isNaN(tst2)){
            val2 = Number(val2);
        }
    }
    if(eq == '='){
        if(val == val2 || val == '' && val2 == 0 || val == 0 && val2 == ''){
            return true;
        } else {
            return false;
        }
    } else if(eq == '<'){
        if(val < val2){
            return true
        } else {
            return false
        }
    } else if(eq == '>'){
        if(val > val2){
            return true
        } else {
            return false
        }
    } else if(eq == '=<'){
        if(val <= val2){
            return true
        } else {
            return false
        }
    } else if(eq == '=>'){
        if(val >= val2){
            return true
        } else {
            return false
        }
    } else if(eq == '<='){
        if(val <= val2){
            return true
        } else {
            return false
        }
    } else if(eq == '>='){
        if(val >= val2){
            return true
        } else {
            return false
        }
    } else if(eq == '<>'){
        if(val != val2){
            return true
        } else {
            return false
        }
    }
}

function matchFuncExcel(data,table,socketid){
    data = data.split(';');
    var search = data[0];
        search = getValueCellInstring(search,table,'',socketid);
        search = replaceall('"','',search);

    var radius = data[1];
    var precision;
    if(data.length > 2){
        precision = data[2];
    } else {
        precision = 1;
    }
    radius = getValueCellRange(radius,table,socketid);

    if(precision == 0){
        var srcData = radius.indexOf(search);
        if(srcData == -1){
            return '#error#';
        } else {
            return srcData+1;
        }
    } else if(precision == 1){
        var ck = radius.indexOf(search);
        if(ck == -1){
            search = Number(search);
            if(!isNaN(search) && search != 0){
                var min = '#error#';
                for(var i = 0;i < radius.length;i++){
                    if(Number(radius[i]) < search){
                        min = i+1;
                    }
                }
                return min; 
            } else{
                if(search == '#error#'){
                    return '#error#';
                } else {
                    return radius.length-1;
                }
            }
        } else {
            return ck+1;
        }
    }
}

function GetChar(nr){  
    var circle = -1;
    var count = 26;
    var curent = 0;

    var needChar;

    for(var i = 0;i < nr ;i++){
        var groupIndex = '';
        if(circle == -1){
            needChar = alphaBet[curent];
        } else if(circle > -1){
            needChar = alphaBet[circle]+alphaBet[curent];
        }

        count--;
        curent++;

        if(count == 0){
            count = 26;
            circle++;
            curent = 0;
        }

    }
    return needChar;
}

function exAdressFunc(data,table,socketid){
    data = data.split(';');
    var ck = getValueCellInstring(data[0],table,'',socketid);
    var ck2 = getValueCellInstring(data[1],table,'',socketid);
        ck = calcEval(ck);
        ck2 = calcEval(ck2);
        ck = Number(ck);
        ck2 = Number(ck2);

    if(!isNaN(ck) && !isNaN(ck2)){
            if(ck2 >=1 && ck > 0 ){
                var char = '';    
                    char = GetChar(ck2); 
                if(table == 'cellsMain' ){
                    table = 'main';
                } else if(table == 'cellsCalc'){
                    table = 'calc';
                }
                return '{'+table+'.'+char+ck+'}';
            } else {
                return '#error#';
            }
    } else {
        return '#error#';
    }
}

function getRowCell(data){
    if(data == ''){
        console.log('-----> line 1273 function getRowCell need transform')
    } else {
        var tmpData = data;
        var next = true;
        
        var ck = tmpData.indexOf('{');
        var ck2 = tmpData.indexOf('}');
        if(ck > -1 && ck2 > -1){
            var remove = tmpData.substring(ck, ck2+1);
            tmpData = tmpData.replace(remove,'');
        } else {
            next = false;
        }
        if(tmpData.length > 1){
            next = false;
        }
        if(next == true){
            var cut = data.substring(ck+1, ck2);
            if(tmpData.length == 0){
                if(cut.indexOf(':') > -1){
                    cut = cut.split(':');
                    cut = cut[0];
                }
                return cut.replace(/[^0-9]/g, '');
            } else {
                return "#error#";
            }
        } else {
            return "#error#";
        }
    }
}

function getResultFuncAND(ck,data,table,socketid){
    data = data.split(ck);
    data[0] = getValueCellInstring(data[0],table,'',socketid);
    data[1] = getValueCellInstring(data[1],table,'',socketid);
    if(data[0] == '')data[0] = 0;
    if(data[1] == '')data[1] = 0;
    if(data[0] == '#error#' || data[0] == '#error#'){
        return '#error#';
    } else {
        if(!isNaN(Number(data[0])))data[0] = Number(data[0]);
        if(!isNaN(Number(data[1])))data[1] = Number(data[1]);

        var result = '';
        if(ck == '>'){
            result = eval(data[0] > data[1]);
        } else if(ck == '<'){
            result = eval(data[0] < data[1]);
        } else if(ck == '<>'){
            result = eval(data[0] > data[1]);
            if(result == false){
                result = eval(data[0] < data[1]);
            }
        } else if(ck == '=='){
            result = data[0]==data[1];
        } else {
            data = data[0]+ck+data[1];
            result = eval(data);
        }
        if(result == false){
            return 'FALSE';
        } else if(result == true) {
            return 'TRUE';
        } else {
            return '#error#';
        }
    }
}

function searchOperation(cell){
    if(cell.indexOf('>=') > -1){
        return '>=';
    } else if(cell.indexOf('=>') > -1){
        return '=>';
    } else if(cell.indexOf('=<') > -1){
        return '=<';
    } else if(cell.indexOf('<=') > -1){
        return '<=';
    } else if(cell.indexOf('==') > -1){
        return '==';
    } else if(cell.indexOf('<>') > -1){
        return '<>';
    } else if(cell.indexOf('>') > -1){
        return '>';
    } else if(cell.indexOf('<') > -1){
        return '<';
    } else {
        return false;
    }
}

function roundup(x, digits) {
    if(isNaN(Number(x)) || x == ''){
        return '#error#';
    } else {
            x = x.toFixed(8);
        if(!digits)digits = 0;
        var m=Math.pow(10, digits);
            m = Math.ceil(x*m)/m;
            x = x+'';
        var len = x.split('.');
            if(len[1]){
                len = len[1].length;
            } else {
                len = 0;
            }
            if(len < digits)digits = len;
            m = m.toFixed(digits);
            m = Number(m);
            if(isNaN(m))m = '#error#';
        return m;
    }
}

function excelRound(val, num) {
    var coef = Math.pow(10, num);
    return (Math.round(val * coef))/coef
}

function getCell(str,start,finish,table){
    var cell = str.substring(start, finish);
        
        var ckOtherTable = cell.indexOf('.');
        if(ckOtherTable > -1){
            cell = cell.split('.');
            table = cell[0]; 
            cell = cell[1]; 
        }        
    var num = cell.replace(/[^0-9]/g, '');
    var char = cell.indexOf(num);
        char = cell.substring(0, char);

    var arr = [];
        arr.push(char);
        arr.push(num);
        arr.push(table);
    return arr;
}

function getCellAndRange(str,start,finish,table){
    var cell = str.substring(start, finish);
        if(cell.indexOf(':') > -1){
            return getCellCharRange(cell,table);
        } else { 
            var num = cell.replace(/[^0-9]/g, '');
            var char = cell.indexOf(num);
                char = cell.substring(0, char);

            var arr = [];
                arr.push(char);
                arr.push(num);
                arr.push(table);
            var arrParent = [];
                arrParent.push(arr);
            return arrParent;
        }
}

function getValueCell(cell,socketid){
    var group = cell[0];
    var line = cell[1];
    var table = cell[2];

    if(table == 'cellsMain' ){
        table = 'main';
    } else if(table == 'cellsCalc'){
        table = 'calc';
    }

    var cellItem = '';
    if(line && group){
        
        if(line > 0 ){
            if(excelData[socketid][table][group+line]){
               return excelData[socketid][table][group+line]['front']; 
            } else {
                return '';
            } 
        } else {
            return '';
        }
    } else {
        return '';
    }
}


function getValueCellRange(data,table,socketid){
    var count = 0;
    var openQuotation = false;
    var cells = [];
    var openCell = '';
    for(var i = 0;i < data.length;i++){
        if(data.charAt(i) == '{' && openQuotation == false){
            count += 1;
            openCell = cells.length;

            var arr = [];
                arr.push(i);
            cells.push(arr);    

        } else if(data.charAt(i) == '}' && openQuotation == false){
            cells[openCell].push(i);
        } else if(data.charAt(i) == '"' && openQuotation == false){
            openQuotation = true;
        } else if(data.charAt(i) == '"' && openQuotation == true){
            openQuotation = false;
        }
    }

    
    data = data.substring(cells[0][0]+1, cells[0][1]);
    cells = data.split(':');
    var cell_1 = [];
        var ckOtherTable = cells[0].indexOf('.');
        if(ckOtherTable > -1){
            var cell = cells[0].split('.');
            table = cell[0];
            cells[0] = cell[1];
        }
        var num = cells[0].replace(/[^0-9]/g, '');
        var char = cells[0].indexOf(num);
            char = cells[0].substring(0, char);
        cell_1.push(char);
        cell_1.push(num);
    var cell_2 = [];
        var ckOtherTable = cells[1].indexOf('.');
        if(ckOtherTable > -1){
            var cell = cells[1].split('.');
            table = cell[0];
            cells[1] = cell[1];
        }
        var num = cells[1].replace(/[^0-9]/g, '');
        var char = cells[1].indexOf(num);
            char = cells[1].substring(0, char);
        cell_2.push(char);
        cell_2.push(num);

    cells = [];
    cells.push(cell_1);
    cells.push(cell_2);

    var result;

    if(cells[0][0] != cells[1][0]){
      result = valueInRange(cells,'char',table,socketid);
    } else if(cells[0][1] != cells[1][1]){
      result = valueInRange(cells,'number',table,socketid);
    }
    return result;
}

function valueInRange(cells,type,table,socketid){
    var sum = 0;
    var arr = [];
    if(type == 'number'){
        var char = cells[0][0];
        for(var i = Number(cells[0][1]);i <= Number(cells[1][1]);i++){
            if(excelData[socketid][table][cells[0][0]+i]){
                var cell = excelData[socketid][table][cells[0][0]+i]['front'];
            } else {
                var cell = '';
            }
            arr.push(cell)
        }
    } else if(type == 'char'){
        // alphaBet
        var children = cells[0][1];
        var nrChar1 = alphaBet.indexOf(cells[0][0]);
        var nrChar2 = alphaBet.indexOf(cells[1][0]);
        for(var i = nrChar1;i <= nrChar2;i++){
            var char = alphaBet[i];
            if(excelData[socketid][table][char+children]){
                var cell = excelData[socketid][table][char+children]['front'];
            } else {
                var cell = '';
            }
            arr.push(cell)
        }
    }
    return arr;
}


function getCellCharRange(data,table){
    var count = 0;
    var cells = [];

    cells = data.split(':');
    var cell_1 = [];
        var ckOtherTable = cells[0].indexOf('.');
        if(ckOtherTable > -1){
            var cell = cells[0].split('.');
            table = cell[0];
            cells[0] = cell[1];
        }
        var num = cells[0].replace(/[^0-9]/g, '');
        var char = cells[0].indexOf(num);
            char = cells[0].substring(0, char);
        cell_1.push(char);
        cell_1.push(num);
        cell_1.push(table);
    var cell_2 = [];
        var ckOtherTable = cells[1].indexOf('.');
        if(ckOtherTable > -1){
            var cell = cells[1].split('.');
            table = cell[0];
            cells[1] = cell[1];
        }
        var num = cells[1].replace(/[^0-9]/g, '');
        var char = cells[1].indexOf(num);
            char = cells[1].substring(0, char);
        cell_2.push(char);
        cell_2.push(num);
        cell_2.push(table);
    cells = [];
    cells.push(cell_1);
    cells.push(cell_2);

    var result;

    if(cells[0][0] != cells[1][0]){
      result = charInRange(cells,'char',table);
    } else if(cells[0][1] != cells[1][1]){
      result = charInRange(cells,'number',table);
    }
    return result;
}

function charInRange(cells,type,table){
    var sum = 0;
    var arr = [];
    if(type == 'number'){
        for(var i = Number(cells[0][1]);i <= Number(cells[1][1]);i++){
            var arrTmp = [];
                arrTmp.push(cells[0][0]);
                arrTmp.push(i);
                arrTmp.push(table);
            arr.push(arrTmp);
        }
    } else if(type == 'char'){
        // alphaBet
        var children = cells[0][1];
        var nrChar1 = alphaBet.indexOf(cells[0][0]);
        var nrChar2 = alphaBet.indexOf(cells[1][0]);
        for(var i = nrChar1;i <= nrChar2;i++){
            var char = alphaBet[i];
            var arrTmp = [];
                arrTmp.push(char);
                arrTmp.push(children);
                arrTmp.push(table);
            arr.push(arrTmp);
        }
    }
    return arr;
}


function getValueCellInstring(data,table,type,socketid){
    if(data.indexOf('#error#') > -1){
        return '#error#';
    } else {
        var count = 0;
        var openQuotation = false;
        var cells = [];
        var openCell = '';
        for(var i = 0;i < data.length;i++){
            if(data.charAt(i) == '{' && openQuotation == false){
                count += 1;
                openCell = cells.length;

                var arr = [];
                    arr.push(i);
                cells.push(arr);    

            } else if(data.charAt(i) == '}' && openQuotation == false){
                cells[openCell].push(i);
            } else if(data.charAt(i) == '"' && openQuotation == false){
                openQuotation = true;
            } else if(data.charAt(i) == '"' && openQuotation == true){
                openQuotation = false;
            }
        }
        if(count > 0){
            for(var i = 0;i < cells.length;i++){
                var start = cells[i][0]+1;
                var finish = cells[i][1];
                var cell = getCell(data,start,finish,table);
                var result = getValueCell(cell,socketid);
                    if(result == '' && type == 'number'){
                        cells[i].push('0');
                    } else {
                        cells[i].push(result);
                    }
            }
            for(var i = (Number(cells.length)-1);i >= 0;i--){
                var start = cells[i][0];
                var finish = cells[i][1]+1;

                data = replaceRange(data, start, finish, cells[i][2]);
            }
        }
        return data;
    }
}


function replaceRange(s, start, end, substitute) {
    return s.substring(0, start) + substitute + s.substring(end);
}

function calcEval(data){
    if(data.length == 0){
        return '#error#';
    } else {
        if(data.indexOf('$=$') > -1){
            data = data.replace(/"/g, '');
            data = data.split('$=$');

            try {
                data = eval(data[0]==data[1]);
                if(data == true){
                    data = 'TRUE';
                } else {
                    data = 'FALSE';
                }
                return data;
            } catch (e) {
                if (e instanceof SyntaxError || e instanceof ReferenceError) {
                    return '#error#';
                }
            }
        } else {
            data = data.replace(/"/g, "`");
            data = data.replace(/"/g, '"');
            data = data.replace(/&/g, "+");
            data = data.replace(/=/g, "==");
            try {
                data = replaceall('$-$','-',data);
                data = replaceall('$+$','+',data);
                data = replaceall('$*$','*',data);
                data = replaceall('$/$','/',data);
                data = eval(data);
                return data;
            } catch (e) {
                if (e instanceof SyntaxError || e instanceof ReferenceError) {
                    return '#error#';
                }
            }
        }
    }
}

function checkToString(data){
    var  ck = false;
    if(data && data != '#error#'){
        for(var i = 0;i < data.length;i++){
            var item = Number(data.charAt(i));
            if(isNaN(item)){
                ck = true;
                break;
            }
        }
    } else if(data == '#error#'){
        ck = true;
    }   
    return ck;
}

function recalcFunction(cellName,tableN,action,socketid){     
        var childrenInsert = false;
        if(action != 'next'){
            if(tableN == 'cellsMain'){
                if(registerMainCell[socketid][cellName]){
                    for(var i = 0;i < registerMainCell[socketid][cellName].length;i++){
                        var cell = registerMainCell[socketid][cellName][i].split('_');
                        var table = cell[1];

                        if(table == 'cellsCalc'){
                            tb = 'calc';
                        } else if(table == 'cellsMain'){
                            tb = 'main';
                        }
                        if(excelData[socketid][tb][cell[0]]){
                            var backdata = excelData[socketid][tb][cell[0]]['back'];

                            if(backdata){
                                childrenInsert = true;
                                var partExecuted = getNeedCellToMap(socketid);
                                partExecuted[cell[0]+table] = {};
                            }
                        }
                    }
                }
            } else if(tableN == 'cellsCalc'){
                if(registerCalcCell[socketid][cellName]){
                    for(var i = 0;i < registerCalcCell[socketid][cellName].length;i++){
                        var cell = registerCalcCell[socketid][cellName][i].split('_');
                        var table = cell[1];

                        if(table == 'cellsCalc'){
                            tb = 'calc';
                        } else if(table == 'cellsMain'){
                            tb = 'main';
                        }
                        if(excelData[socketid][tb][cell[0]]){
                            var backdata = excelData[socketid][tb][cell[0]]['back'];

                            if(backdata){
                                childrenInsert = true;
                                var partExecuted = getNeedCellToMap(socketid);
                                partExecuted[cell[0]+table] = {};
                            }
                        }
                    }
                }
            }
        }
        if(childrenInsert == true){
            index = Object.keys(partExecuted)[0];
            mapExecute[socketid] += '_'+index;
            cell = '';
            cell = getCellCalcBackData(index,socketid);
            if(cell){
               return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'calcBack'};
            } else {
                console.log(cell)
            }
        } else {
            cell = '';
            cell = getNeedCellToMap(socketid);
            var mapArr = mapExecute[socketid].split('_');
            mapArr.pop();
            mapExecute[socketid] = registerMapRenew(mapArr,socketid);

            index = getNeedExecuteCell(mapArr,socketid);
            if(mapArr.length == 0 && action == 'wait'){
                action = '';
            }
            if(index){
                mapExecute[socketid] += '_'+index;
                cell = '';
                cell = getCellCalcBackData(index,socketid);
                if(cell){
                    return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'calcBack'};
                }
            } else {
                    if(executeList[socketid].length > 0){
                        if(action != 'wait'){
                            executeNow[socketid] = executeList[socketid][0];
                            executeList[socketid].shift();
                            cell = getCellCalcBackData(executeNow[socketid],socketid);
                            if(cell){
                                registerChangesCell[socketid] = {};   
                                registerChangesCell[socketid][cell[1][0]+cell[1][1]] = {};
                                mapExecute[socketid] = cell[1][0]+cell[1][1];

                                if(cell[0]){
                                    return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'calcBack'};
                                } else {
                                    return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'wait'};
                                }
                            }
                            else {
                                registerChangesCell[socketid] = {};   
                                registerChangesCell[socketid][cell[1][0]+cell[1][1]] = {};
                                mapExecute[socketid] = cell[1][0]+cell[1][1];

                                if(cell[0]){
                                    return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'calcBack'};
                                } else {
                                    return {'cell1': cell[0],'cell2':cell[1],'socketid': socketid,'type':'wait'};
                                }
                            }
                        }
                    } else {
                        executeNow[socketid] = '';
                        return {'type':'Finish'};
                    }
                
            }
        }   
}

function getNeedCellToMap(socketid){
    var mapArr = mapExecute[socketid].split('_');
    var needPart;
    for(var i = 0; i < mapArr.length;i++){
        if(needPart){
            needPart = needPart[mapArr[i]];
        } else {
            needPart = registerChangesCell[socketid][mapArr[i]]; 
        }
    }
    return needPart;
}

function registerMapRenew(mapArr,socketid){
    delete mapExecute[socketid];
    mapExecute[socketid] = '';
    for(var i= 0;i < mapArr.length;i++){
        if(mapArr[i]){
            if(mapExecute[socketid]){
                mapExecute[socketid] += '_'+mapArr[i];
            } else {
                mapExecute[socketid] += mapArr[i];
            }
        }
    }
    return mapExecute[socketid]
}

function getCellCalcBackData(index,socketid){
    var cell,table;
    if(index.indexOf('cellsCalc') > 1){
        cell = index.replace('cellsCalc','');
        tb = 'calc';
        table = 'cellsCalc';
    } else if(index.indexOf('cellsMain') > 1){
        cell = index.replace('cellsMain','');
        tb = 'main';
        table = 'cellsMain';
    }


    var backdata = excelData[socketid][tb][cell]['back'];
        var arr = [];
            arr.push(cell);
            arr.push(table);
        var result = [];
            result.push(backdata);
            result.push(arr);
        return  result;
}

function getNeedExecuteCell(mapArr,socketid){
    var needCell,needPart;
    for(;;){
        for(var i = 0; i < mapArr.length;i++){
            if(!needPart){
                needPart = registerChangesCell[socketid][mapArr[i]];  
            } else {
                needPart = needPart[mapArr[i]];
            }
        }
        if(needPart){
            if(mapArr.length > 0){
                if(Object.keys(needPart).length <= 1){
                    delete needPart[Object.keys(needPart)[0]];
                    mapArr.pop();
                    mapExecute[socketid] = registerMapRenew(mapArr);
                    needPart = '';
                } else {
                    delete needPart[Object.keys(needPart)[0]];
                    needCell = Object.keys(needPart)[0];
                    break;
                }
            } else {
                break;
            }
        } else {
            break;
        }
    }
    return needCell;
}
