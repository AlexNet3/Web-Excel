var excelData;

const socketE = io('http://localhost:8000');   
var firstLoad = false;

var tableExcel = 'main';
var registerMainCell = {};
var registerCalcCell = {};

var executeNow = '';
var executeList = [];

var lastAction = {
	'main':{},
	'calc':{}
}

var tables = ['mainContentExcel','calcContentExcel'];

///////////// costum  stye  column ////////
/*
var styleWidth = {
	'A' : 108,
	'D' : 101,
	'G' : 184,
	'K' : 122
};
*/
///////////// costum  stye  column ////////
function addLine(){
	for(var iM = 0;iM < tables.length;iM++){
		var main = document.getElementById(tables[iM]);

		var listNumberExcel = main.getElementsByClassName('listNumberExcel')[0];

		for(var i = 1;i < 267 ;i++){
			var lineNr = document.createElement('div');
				lineNr.classList.add('lineNr');
				lineNr.classList.add('line_'+i);
				lineNr.classList.add('childLine_'+i);
				lineNr.setAttribute('grouptype','line');
				lineNr.setAttribute('lineNr',i);

				lineNr.innerText = i;
				var resizeLine = document.createElement('div');
					resizeLine.classList.add('resizeLine');
				lineNr.append(resizeLine);	
				
			listNumberExcel.append(lineNr);
		}
	}
}

var alphaBet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
function addGoup(){
	for(var iM = 0;iM < tables.length;iM++){
		var main = document.getElementById(tables[iM]);

		var listNumberExcel = main.getElementsByClassName('listGroupExcel')[0];
		
		var circle = -1;
		var count = 26;
		var curent = 0;
		for(var i = 0;i < 23 ;i++){
			var groupNr = document.createElement('div');
				groupNr.classList.add('groupNr');

			var groupIndex = '';
			if(circle == -1){
			 	groupNr.innerText = alphaBet[curent];
			 	groupNr.classList.add('group_'+alphaBet[curent]);
			 	groupNr.setAttribute('groupName',alphaBet[curent]);
			 	groupNr.setAttribute('grouptype','group');
			 	groupIndex = alphaBet[curent];
			 	if(iM === 0){
			 	///////////// costum  stye  column ////////
			 	/*
			 	var ckStyle = styleWidth[alphaBet[curent]];
			 		if(ckStyle){
			 			groupNr.style.width = (ckStyle-1)+'px';
			 		}
			 	*/
			 	///////////// costum  stye  column ////////
			 	}
			} else if(circle > -1){
				groupNr.innerText =  alphaBet[circle]+alphaBet[curent];
				groupNr.classList.add('group_'+alphaBet[circle]+alphaBet[curent]);
				groupNr.setAttribute('groupName',alphaBet[circle]+alphaBet[curent]);
				groupIndex = alphaBet[circle]+alphaBet[curent];
			}
				var resize = document.createElement('div');
					resize.classList.add('resizeGroup');
				groupNr.append(resize);	

			var listCels = main.getElementsByClassName('listCels')[0];
			var group = document.createElement('div');
				group.classList.add('groupCellExcel');
				group.classList.add('childrenGroup_'+groupIndex);
				if(iM === 0){
					///////////// costum  stye  column ////////
					/*
						var ckStyle = styleWidth[groupIndex];
				 		if(ckStyle){
				 			group.style.width = ckStyle+'px';
				 		}
					*/
					///////////// costum  stye  column ////////
					}
				for (var i2 = 1;i2 < 267;i2++){
						var cell = document.createElement('div');
							cell.classList.add('cellExcel');
							cell.classList.add('childrenLine_'+i2);
							cell.setAttribute('line', i2);
							cell.setAttribute('group', groupIndex);
							cell.setAttribute('backData','');
							var txt = document.createElement('div');
								txt.classList.add('txtCellExcel');
								txt.setAttribute('front','');
								cell.append(txt);
							group.append(cell);
					listCels.append(group);
				}

			count--;
			curent++;

			if(count == 0){
				count = 26;
				circle++;
				curent = 0;
			}

			listNumberExcel.append(groupNr);
		}
	}
	selectCellFunc();
	///////////////////////////// DB part ////////////////////////////////////////////////////////////////////////////
	 socketE.emit('loadFrontDataExcel',tables);
	///////////////////////////// DB part ////////////////////////////////////////////////////////////////////////////
}
addLine();
addGoup();

const parentListCels = document.getElementsByClassName('parentListCels');
	for(var i = 0;i < parentListCels.length;i++){
		parentListCels[i].addEventListener("scroll", function () {
			var mainBodySearch = document.getElementById(tableExcel+'ContentExcel');

	        var left = this.scrollLeft;
	        var top = this.scrollTop;
				var parentListGroupExcel = mainBodySearch.getElementsByClassName('parentListGroupExcel')[0];
					parentListGroupExcel.scrollLeft = left;

				var parentListNumberExcel = mainBodySearch.getElementsByClassName('parentListNumberExcel')[0];
					parentListNumberExcel.scrollTop = top;
		});
	}


function getValueAtribute(t,type){
    for(var i = 0;i < t.length;i++){
        if(t[i].name == type){
            return t[i].value;
        }
    }
}
function setValueAtribute(t,type,value){
    for(var i = 0;i < t.length;i++){
        if(t[i].name == type){
            t[i].value = value;
        }
    }
}   

function selectCellFunc(){
	var parentListCels = document.getElementsByClassName('parentListCels');
		var cellExcel = document.querySelectorAll(".cellExcel");
		for(var i = 0;i < cellExcel.length;i++){
			cellExcel[i].addEventListener("click", function () {
				var item = this.parentNode.parentNode.parentNode.parentNode.parentNode;
				var listSelect = item.getElementsByClassName('selectCellExcel');
					if(listSelect.length > 0){
						var lineOld = getValueAtribute(listSelect[0].attributes,'line');
						var charOld	= getValueAtribute(listSelect[0].attributes,'group');

						var lineNew = getValueAtribute(this.attributes,'line');
						var charNew	= getValueAtribute(this.attributes,'group');

						if(charOld+lineOld != charNew+lineNew){
							if(listSelect[0].children.length > 1){
								listSelect[0].children[0].classList.remove('slectListExcel');
								listSelect[0].children[0].classList.add('slectListExcel');
								listSelect[0].children[2].style.display = '';
							    listSelect[0].children[2].style.width = '';
							    listSelect[0].children[2].style.marginTop = '';
							    listSelect[0].children[2].style.position = '';
							    listSelect[0].children[1].textContent = "▼";
							    listSelect[0].children[1].style.display = 'none';
							}
						}
						listSelect[0].classList.remove('selectCellExcel');
					}

				var selectLineNr = item.getElementsByClassName('selectLineNr');
					if(selectLineNr.length > 0){
						selectLineNr[0].classList.remove('selectLineNr');
					}

				var selectGroupNr = item.getElementsByClassName('selectGroupNr');
					if(selectGroupNr.length > 0){
						selectGroupNr[0].classList.remove('selectGroupNr');
					}

				var editTxtCellExcel = item.getElementsByClassName('editTxtCellExcel');
					if(editTxtCellExcel.length > 0){
						editTxtCellExcel[0].removeAttribute('contenteditable');

						var char = getValueAtribute(editTxtCellExcel[0].parentNode.attributes,'group');
						var number = getValueAtribute(editTxtCellExcel[0].parentNode.attributes,'line');	
						var backData = getValueAtribute(editTxtCellExcel[0].parentNode.attributes,'backdata');
						var frontData = getValueAtribute(editTxtCellExcel[0].attributes,'front');

						var lgt = Object.keys(lastAction[tableExcel]).length;

						lastAction[tableExcel][lgt] = {};
						lastAction[tableExcel][lgt]['front'] = frontData;
						lastAction[tableExcel][lgt]['back'] = backData;
						lastAction[tableExcel][lgt]['cell'] = editTxtCellExcel[0].parentNode;
						lastAction[tableExcel][lgt]['indx'] = char+number;

						setValueAtribute(editTxtCellExcel[0].attributes,'front',editTxtCellExcel[0].innerText);	
						var lineNew = getValueAtribute(this.attributes,'line');
						var charNew	= getValueAtribute(this.attributes,'group');

						if(char+number != charNew+lineNew){
							if(editTxtCellExcel[0].parentNode.children.length > 1){
								editTxtCellExcel[0].parentNode.children[0].classList.remove('slectListExcel');
								editTxtCellExcel[0].parentNode.children[0].classList.add('slectListExcel');
								editTxtCellExcel[0].parentNode.children[2].style.display = '';
							    editTxtCellExcel[0].parentNode.children[2].style.width = '';
							    editTxtCellExcel[0].parentNode.children[2].style.marginTop = '';
							    editTxtCellExcel[0].parentNode.children[2].style.position = '';
							    editTxtCellExcel[0].parentNode.children[1].textContent = "▼";
							    editTxtCellExcel[0].parentNode.children[1].style.display = 'none';
							}
						} else {
							if(editTxtCellExcel[0].parentNode.children.length > 1){
								if (editTxtCellExcel[0].parentNode.children[2].style.display === '') {
								    editTxtCellExcel[0].parentNode.children[2].style.display = 'block';
								    editTxtCellExcel[0].parentNode.children[2].style.width = editTxtCellExcel[0].parentNode.clientWidth+'px';
								    var ckClass= checkClass(editTxtCellExcel[0].parentNode.children[0].classList,'slectListExcel');
								    if(ckClass){
								    	editTxtCellExcel[0].parentNode.children[2].style.marginTop = editTxtCellExcel[0].parentNode.clientHeight+'px';
								    }
								    editTxtCellExcel[0].parentNode.children[2].style.position = 'absolute';
								} else {
									editTxtCellExcel[0].parentNode.children[2].style.display = '';
								    editTxtCellExcel[0].parentNode.children[2].style.width = '';
								    editTxtCellExcel[0].parentNode.children[2].style.marginTop = '';
								    editTxtCellExcel[0].parentNode.children[2].style.position = '';
								    editTxtCellExcel[0].parentNode.children[1].textContent = "▼";
								}
							}
						}

						var backD = getValueAtribute(this.attributes,'backdata');

						if(backD){
							socketE.emit('insertInBigData', {table :tableExcel, cell:char+number, data: backD, type: 'back'});
						} else {
							socketE.emit('insertInBigData', {table :tableExcel, cell:char+number, data: editTxtCellExcel[0].innerText, type: 'front'});
						} 


						editTxtCellExcel[0].classList.remove('editTxtCellExcel');
					}

				var editCellExcel = item.getElementsByClassName('editCellExcel');
					if(editCellExcel.length > 0){
						editCellExcel[0].classList.remove('editCellExcel');
					}

				var shiftSelect_visual = item.getElementsByClassName('shiftSelect_visual');
					if(shiftSelect_visual.length > 0){
						for(var i = Number(shiftSelect_visual.length)-1;Number(shiftSelect_visual.length)-1 >= 0;i--){
							shiftSelect_visual[i].classList.remove('shiftSelect_visual');
						}
					}

				var selectGroupNr = item.getElementsByClassName('selectGroupNr');
					if(selectGroupNr.length > 0){
						for(var i = Number(selectGroupNr.length)-1;Number(selectGroupNr.length)-1 >= 0;i--){
							selectGroupNr[i].classList.remove('selectGroupNr');
						}
					}			
				
				this.classList.add('selectCellExcel');
				if(this.children.length > 1){
					this.children[1].style.display = '';
				}

				shiftSelectCell_visual[tableExcel] = [this]; 

				var line = getValueAtribute(this.attributes,'line');
				var group = getValueAtribute(this.attributes,'group');

				var index = group+line;
				shiftSelectCell[tableExcel] = [index];

				var lineNr = item.getElementsByClassName('line_'+line)[0];
					lineNr.classList.add('selectLineNr');

				var groupNr = item.getElementsByClassName('group_'+group)[0];
					groupNr.classList.add('selectGroupNr');

				var num =  Number(this.children[0].innerText);
				var summ_infoSummCellExcell = document.getElementById('summ_infoSummCellExcell');

					if(isNaN(num)){
						summ_infoSummCellExcell.innerText = 0;
					} else {
						summ_infoSummCellExcell.innerText = num;
					}

				var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');
					var backData = getValueAtribute(this.attributes,'backdata');
						if(backData){
							infoBackCellDataExcel.value = backData;
						} else {
							infoBackCellDataExcel.value = this.children[0].innerText;
						} 

				var acceptEditFormTableExcel = document.getElementById('acceptEditFormTableExcel');
					if(acceptEditFormTableExcel.style.display == 'flex'){
						acceptEditCell();
					}	
			});
		}
}

function checkClass(t,type){
	var list = [];
	for(var i = 0;i < t.length; i++){
		list.push(t[i]);
	}
    var find = list.indexOf(type);
    if(find > -1){
    	return true;
    } else {
    	return false;
    }
}

var registerChangesCell = {};
var lastIndexRegister = 0;
var mapExecute;

var loadDataIndex = '';


window.addEventListener('keydown', clickKeyboard);
window.addEventListener('keyup', keyUpKeyboard);
window.addEventListener('dblclick', clickKeyboard);

var shiftSelectCell_visual = {};
var shiftSelectCell = {};

function keyUpKeyboard(e){
	var arrayKey2 = [8,32,48,49,50,51,52,53,54,55,56,57,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,96,97,98,99,100,101,102,103,104,105,106,107,109,110,111,186,187,188,189,190,191,192,219,220,221,222,226];

    var arrayKeyLock = [13,20,37,38,39,40];
	var find = arrayKey2.indexOf(event.keyCode);
    		
    if(find > -1 && event.ctrlKey == false){
    	var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');

    	if(event.keyCode == 32){
    		if(event.target.id != 'infoBackCellDataExcel'){
    			infoBackCellDataExcel.value += '\xa0';
    		} else {
    			var editTxtCellExcel = document.getElementsByClassName('editTxtCellExcel');
    			if(editTxtCellExcel.length > 0){
    				editTxtCellExcel[0].innerText += '\xa0';
    			}
    		}
    	} else {
    		if(event.target.id != 'infoBackCellDataExcel'){
    			if(event.target.localName != 'body'){
    				infoBackCellDataExcel.value = event.target.innerText;
    			}
    		} else {
    			var editTxtCellExcel = document.getElementsByClassName('editTxtCellExcel');
    			if(editTxtCellExcel.length > 0){
    				editTxtCellExcel[0].innerText = event.target.value;
    			}
    		}
    	}
    }
}

async function clickKeyboard(e){
	if(event.type == 'dblclick' || event.keyCode == 9){
		event.preventDefault();
	}
	var bodyM = document.getElementById(tableExcel+'ContentExcel');
    var selectCellExcel = bodyM.getElementsByClassName('selectCellExcel');
    var arrayKey = [48,49,50,51,52,53,54,55,56,57,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,96,97,98,99,100,101,102,103,104,105,106,107,109,110,111,186,187,188,189,190,191,192,219,220,221,222,226];

    	if(selectCellExcel.length > 0){
    		var find = arrayKey.indexOf(event.keyCode);
    		var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');
    			
    		if(find > -1 && event.shiftKey == false && event.ctrlKey == false){
	    		selectCellExcel[0].classList.add('editCellExcel');
	    		focusCellInfoBack();

	    		selectCellExcel[0].children[0].setAttribute('contenteditable','true');
	    		selectCellExcel[0].children[0].classList.add('editTxtCellExcel');

	    		if(selectCellExcel[0].children.length > 1){
		    		selectCellExcel[0].children[0].classList.remove('slectListExcel');
		    		selectCellExcel[0].children[2].style.display = 'none';
	    		}
	    		if(event.target.id != 'infoBackCellDataExcel'){
	    			selectCellExcel[0].children[0].focus();

		    		selectCellExcel[0].children[0].innerText = '';
		    		setValueAtribute(selectCellExcel[0].attributes,'backdata','')
					infoBackCellDataExcel.value = ''; 
	    		}
	    		selectCellExcel[0].classList.remove('selectCellExcel');
    		}
    		if(event.keyCode == 8 && event.shiftKey == false && event.ctrlKey == false ){
    			var char = getValueAtribute(selectCellExcel[0].attributes,'group');
				var number = getValueAtribute(selectCellExcel[0].attributes,'line');

    			if(shiftSelectCell[tableExcel].length > 1){
    				for(var i = 0; i < shiftSelectCell[tableExcel].length;i++){
    					shiftSelectCell_visual[tableExcel][i].children[0].innerText = '';
    				}
    			} else {
    				var backData = getValueAtribute(selectCellExcel[0].attributes,'backdata');

    				var lgt = Object.keys(lastAction[tableExcel]).length;
					lastAction[tableExcel][lgt] = {};
					lastAction[tableExcel][lgt]['front'] =  selectCellExcel[0].children[0].innerText;
					lastAction[tableExcel][lgt]['back'] = backData;
					lastAction[tableExcel][lgt]['cell'] = selectCellExcel[0];
					lastAction[tableExcel][lgt]['indx'] = char+number;

    				selectCellExcel[0].children[0].innerText = '';
    				infoBackCellDataExcel.value = '';
    			}
				socketE.emit('insertInBigData', {table :tableExcel, cell:char+number, data: "", type: 'back'});	
				socketE.emit('insertInBigData', {table :tableExcel, cell:char+number, data: "", type: 'front'});

				if(tableExcel == 'main'){
				   var tableN = 'cellsMain';
				} else if(tableExcel == 'calc'){
				   var tableN = 'cellsCalc';
				}
				socketE.emit('recalcFunction',{'cell1':char+number, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel}) 

    			calcSummCell(shiftSelectCell_visual[tableExcel]);
    		}

    	}
    var arrayKey2 = [8,32,48,49,50,51,52,53,54,55,56,57,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,96,97,98,99,100,101,102,103,104,105,106,107,109,110,111,186,187,188,189,190,191,192,219,220,221,222,226];

    	var arrayKeyLock = [13,20,37,38,39,40];
    	var find2 = arrayKeyLock.indexOf(event.keyCode);
    	var target = checkClass(event.target.classList,'editTxtCellExcel');
    		if(find2 > -1){
    			event.preventDefault();
    		}

    	if(event.keyCode == 9 && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}
    		var line = getValueAtribute(mainCell.attributes,'line');
    		var lineNr = line;
    			line = Number(line - 1);
    		mainCell.parentNode.nextSibling.children[line].click();

    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+lineNr, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}
    	if(event.keyCode == 13 && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}	
    		mainCell.nextSibling.click();

    		var line = getValueAtribute(mainCell.attributes,'line');
    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+line, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}																						

    	if(event.key == 'ArrowUp' && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}	
    		mainCell.previousSibling.click();

    		var line = getValueAtribute(mainCell.attributes,'line');
    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+line, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}
    	if(event.key == 'ArrowDown' && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}	
    		mainCell.nextSibling.click();

    		var line = getValueAtribute(mainCell.attributes,'line');
    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+line, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}
    	if(event.key == 'ArrowLeft' && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}
    		var line = getValueAtribute(mainCell.attributes,'line');
    			line = Number(line - 1);
    		mainCell.parentNode.previousSibling.children[line].click();

    		var lineNR = getValueAtribute(mainCell.attributes,'line');
    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+lineNR, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}
    	if(event.key == 'ArrowRight' && event.shiftKey == false && event.ctrlKey == false){
    		var mainCell = selectCellExcel[0];
    		var target2 = checkClass(event.target.classList,'editTxtCellExcel');
    			if(target2 == true){
    				mainCell = event.target.parentNode;
    			}
    		var line = getValueAtribute(mainCell.attributes,'line');
    			line = Number(line - 1);
    		mainCell.parentNode.nextSibling.children[line].click();

    		var lineNr = getValueAtribute(mainCell.attributes,'line');
    		var groupNr = getValueAtribute(mainCell.attributes,'group');
    		var tableN;
    		if(tableExcel == 'main'){
    			tableN = 'cellsMain';
    		} else if(tableExcel == 'calc'){
    			tableN = 'cellsCalc';
    		}
			socketE.emit('recalcFunction',{'cell1':groupNr+lineNr, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    	}
    	////////////////////////////////////////////////////////////////////////////////////////////////////////
    	if(event.key == 'ArrowUp' && event.shiftKey == true){
    		var tempArr = shiftSelectCell_visual[tableExcel];

    			var lastLineItem = tempArr[Number(tempArr.length)-1];
	    			var	lineLLI = getValueAtribute(lastLineItem.previousSibling.attributes,'line');
	    			var groupLLI =  getValueAtribute(lastLineItem.previousSibling.attributes,'group');
	    		var indexLLI = groupLLI  + lineLLI;

	    		if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(indexLLI) < 0){
	    			var lengthSum = shiftSelectCell_visual[tableExcel].length;
	    			for(var i = 0;i < lengthSum; i++){
			    		var ckClass = checkClass(shiftSelectCell_visual[tableExcel][i].classList,'shiftSelect_visual');
							if(ckClass == false){
								shiftSelectCell_visual[tableExcel][i].classList.add('shiftSelect_visual');
							}

			    		var lineI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].previousSibling.attributes,'line');
						var groupI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].previousSibling.attributes,'group');

						colorBlock(lineI,'line','color');

						var index = groupI+lineI;

						if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(index) < 0){
							
							shiftSelectCell_visual[tableExcel][i].previousSibling.classList.add('shiftSelect_visual');

			    			tempArr.push(shiftSelectCell_visual[tableExcel][i].previousSibling);

			    			shiftSelectCell[tableExcel].push(index);
						}
	    			}
	    			shiftSelectCell_visual[tableExcel] = tempArr;

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    			
	    		} else {
	    			var tempColectIndexNeedRemove = [];
	    				for(var i = 0;i < shiftSelectCell_visual[tableExcel].length;i++){
	    					var lineLast = getValueAtribute(lastLineItem.attributes,'line');
	    					var lineTmp =  getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'line');
	    					if(lineTmp == lineLast){
	    						tempColectIndexNeedRemove.push(i);
	    					}
	    				}
	    			tempColectIndexNeedRemove.sort(function(a, b){return b-a});

	    			for(var i = 0;i < tempColectIndexNeedRemove.length;i++){
	    				var lineDellItem =  getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'line');
						var groupDellItem = getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'group'); 

						var valueIndexDell = groupDellItem+lineDellItem;
						var indexBackSelect = shiftSelectCell[tableExcel].indexOf(valueIndexDell);

						shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].classList.remove('shiftSelect_visual');

	    				delete shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]];
	    				delete shiftSelectCell[tableExcel][indexBackSelect];

	    				colorBlock(lineDellItem,'line','remove');
	    			}

	    			shiftSelectCell_visual[tableExcel] = renameKeyInArray(shiftSelectCell_visual[tableExcel]);
	    			shiftSelectCell[tableExcel] = renameKeyInArray(shiftSelectCell[tableExcel]);

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		}
    	}
    	if(event.key == 'ArrowDown' && event.shiftKey == true){
    		var tempArr = shiftSelectCell_visual[tableExcel];
    			
    			var lastLineItem = tempArr[Number(tempArr.length)-1];
	    			var	lineLLI = getValueAtribute(lastLineItem.nextSibling.attributes,'line');
	    			var groupLLI =  getValueAtribute(lastLineItem.nextSibling.attributes,'group');
	    		var indexLLI = groupLLI  + lineLLI;

	    		if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(indexLLI) < 0){
		    		var lengthSum = shiftSelectCell_visual[tableExcel].length;
	    			for(var i = 0;i < lengthSum; i++){
			    		var ckClass = checkClass(shiftSelectCell_visual[tableExcel][i].classList,'shiftSelect_visual');
							if(ckClass == false){
								shiftSelectCell_visual[tableExcel][i].classList.add('shiftSelect_visual');
							}

			    		var lineI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].nextSibling.attributes,'line');
						var groupI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].nextSibling.attributes,'group');

						colorBlock(lineI,'line','color');

						var index = groupI+lineI;

						if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(index) < 0){
							
							shiftSelectCell_visual[tableExcel][i].nextSibling.classList.add('shiftSelect_visual');

			    			tempArr.push(shiftSelectCell_visual[tableExcel][i].nextSibling);

			    			shiftSelectCell[tableExcel].push(index);
						}
	    			}
	    			shiftSelectCell_visual[tableExcel] = tempArr;

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		} else {
	    			var tempColectIndexNeedRemove = [];
	    				for(var i = 0;i < shiftSelectCell_visual[tableExcel].length;i++){
	    					var lineLast = getValueAtribute(lastLineItem.attributes,'line');
	    					var lineTmp =  getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'line');
	    					if(lineTmp == lineLast){
	    						tempColectIndexNeedRemove.push(i);
	    					}
	    				}
	    			tempColectIndexNeedRemove.sort(function(a, b){return b-a});

	    			for(var i = 0;i < tempColectIndexNeedRemove.length;i++){
	    				var lineDellItem =  getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'line');
						var groupDellItem = getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'group'); 

						var valueIndexDell = groupDellItem+lineDellItem;
						var indexBackSelect = shiftSelectCell[tableExcel].indexOf(valueIndexDell);

						shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].classList.remove('shiftSelect_visual');
						
	    				delete shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]];
	    				delete shiftSelectCell[tableExcel][indexBackSelect];

	    				colorBlock(lineDellItem,'line','remove');
	    			}

	    			shiftSelectCell_visual[tableExcel] = renameKeyInArray(shiftSelectCell_visual[tableExcel]);
	    			shiftSelectCell[tableExcel] = renameKeyInArray(shiftSelectCell[tableExcel]);

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		}
    	} 
    	if(event.key == 'ArrowLeft' && event.shiftKey == true){
    	    var tempArr = shiftSelectCell_visual[tableExcel];
    			
    	    	var lastLineItem = tempArr[Number(tempArr.length)-1];
	    			var	lineLLI = getValueAtribute(lastLineItem.attributes,'line');
	    				var lineT = Number(lineLLI) - 1;
	    			var groupLLI =  getValueAtribute(lastLineItem.parentNode.previousSibling.children[lineT].attributes,'group');
	    		var indexLLI = groupLLI  + lineLLI;

	    		colorBlock(groupLLI,'group','color');

	    		if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(indexLLI) < 0){
	    			var lengthSum = shiftSelectCell_visual[tableExcel].length;
	    			for(var i = 0;i < lengthSum; i++){
	    				var ckClass = checkClass(shiftSelectCell_visual[tableExcel][i].classList,'shiftSelect_visual');
							if(ckClass == false){
								shiftSelectCell_visual[tableExcel][i].classList.add('shiftSelect_visual');
							}

	    				var line = getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'line');
			    			line = Number(line - 1);
			    		
			    		var lineI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].parentNode.previousSibling.children[line].attributes,'line');
						var groupI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].parentNode.previousSibling.children[line].attributes,'group');

						var index = groupI+lineI;

						if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(index) < 0){
							
							shiftSelectCell_visual[tableExcel][i].parentNode.previousSibling.children[line].classList.add('shiftSelect_visual');

			    			tempArr.push(shiftSelectCell_visual[tableExcel][i].parentNode.previousSibling.children[line]);

			    			shiftSelectCell[tableExcel].push(index);
						}
	    			}
	    			shiftSelectCell_visual[tableExcel] = tempArr;	

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		} else {
	    			var tempColectIndexNeedRemove = [];
	    				for(var i = 0;i < shiftSelectCell_visual[tableExcel].length;i++){
	    					var groupLast = getValueAtribute(lastLineItem.attributes,'group');
	    					var groupTmp =  getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'group');
	    					if(groupTmp == groupLast){
	    						tempColectIndexNeedRemove.push(i);
	    					}
	    				}
	    			tempColectIndexNeedRemove.sort(function(a, b){return b-a});

	    			for(var i = 0;i < tempColectIndexNeedRemove.length;i++){
	    				var lineDellItem =  getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'line');
						var groupDellItem = getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'group'); 

						var valueIndexDell = groupDellItem+lineDellItem;
						var indexBackSelect = shiftSelectCell[tableExcel].indexOf(valueIndexDell);

						shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].classList.remove('shiftSelect_visual');
						
	    				delete shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]];
	    				delete shiftSelectCell[tableExcel][indexBackSelect];

	    				colorBlock(groupDellItem,'group','remove');
	    			}

	    			shiftSelectCell_visual[tableExcel] = renameKeyInArray(shiftSelectCell_visual[tableExcel]);
	    			shiftSelectCell[tableExcel] = renameKeyInArray(shiftSelectCell[tableExcel]);

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		}
    	}
    	if(event.key == 'ArrowRight' && event.shiftKey == true){
    		var tempArr = shiftSelectCell_visual[tableExcel];
    			
    			var lastLineItem = tempArr[Number(tempArr.length)-1];
	    			var	lineLLI = getValueAtribute(lastLineItem.attributes,'line');
	    				var lineT = Number(lineLLI) - 1;
	    			var groupLLI =  getValueAtribute(lastLineItem.parentNode.nextSibling.children[lineT].attributes,'group');
	    		var indexLLI = groupLLI  + lineLLI;

	    		colorBlock(groupLLI,'group','color');

	    		if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(indexLLI) < 0){
	    			var lengthSum = shiftSelectCell_visual[tableExcel].length;
	    			for(var i = 0;i < lengthSum; i++){
	    				var ckClass = checkClass(shiftSelectCell_visual[tableExcel][i].classList,'shiftSelect_visual');
							if(ckClass == false){
								shiftSelectCell_visual[tableExcel][i].classList.add('shiftSelect_visual');
							}

	    				var line = getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'line');
			    			line = Number(line - 1);
			    		
			    		var lineI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].parentNode.nextSibling.children[line].attributes,'line');
						var groupI = getValueAtribute(shiftSelectCell_visual[tableExcel][i].parentNode.nextSibling.children[line].attributes,'group');

						var index = groupI+lineI;

						if(shiftSelectCell[tableExcel] && shiftSelectCell[tableExcel].indexOf(index) < 0){
							
							shiftSelectCell_visual[tableExcel][i].parentNode.nextSibling.children[line].classList.add('shiftSelect_visual');

			    			tempArr.push(shiftSelectCell_visual[tableExcel][i].parentNode.nextSibling.children[line]);

			    			shiftSelectCell[tableExcel].push(index);
						}
	    			}
	    			shiftSelectCell_visual[tableExcel] = tempArr;

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		} else {
	    			var tempColectIndexNeedRemove = [];
	    				for(var i = 0;i < shiftSelectCell_visual[tableExcel].length;i++){
	    					var groupLast = getValueAtribute(lastLineItem.attributes,'group');
	    					var groupTmp =  getValueAtribute(shiftSelectCell_visual[tableExcel][i].attributes,'group');
	    					if(groupTmp == groupLast){
	    						tempColectIndexNeedRemove.push(i);
	    					}
	    				}
	    			tempColectIndexNeedRemove.sort(function(a, b){return b-a});

	    			for(var i = 0;i < tempColectIndexNeedRemove.length;i++){
	    				var lineDellItem =  getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'line');
						var groupDellItem = getValueAtribute(shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].attributes,'group'); 

						var valueIndexDell = groupDellItem+lineDellItem;
						var indexBackSelect = shiftSelectCell[tableExcel].indexOf(valueIndexDell);

						shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]].classList.remove('shiftSelect_visual');
						
	    				delete shiftSelectCell_visual[tableExcel][tempColectIndexNeedRemove[i]];
	    				delete shiftSelectCell[tableExcel][indexBackSelect];

	    				colorBlock(groupDellItem,'group','remove');
	    			}

	    			shiftSelectCell_visual[tableExcel] = renameKeyInArray(shiftSelectCell_visual[tableExcel]);
	    			shiftSelectCell[tableExcel] = renameKeyInArray(shiftSelectCell[tableExcel]);

	    			calcSummCell(shiftSelectCell_visual[tableExcel]);
	    		}
    	}
    	if(event.key == 'c' && event.ctrlKey == true){
    		var target = event.target.id;
    		if(selectCellExcel.length > 0 && target != 'infoBackCellDataExcel'){
    			copyToClipboard(selectCellExcel[0])
    		}
    	}
    	if(event.key == 'v' && event.ctrlKey == true){
    		if(selectCellExcel.length > 0 && target != 'infoBackCellDataExcel'){
    			const text = await navigator.clipboard.readText();
    			var char = getValueAtribute(selectCellExcel[0].attributes,'group');
				var number = getValueAtribute(selectCellExcel[0].attributes,'line');
				var backD = getValueAtribute(selectCellExcel[0].attribute,'backdata');


				var lgt = Object.keys(lastAction[tableExcel]).length;
				lastAction[tableExcel][lgt] = {};
				lastAction[tableExcel][lgt]['front'] = selectCellExcel[0].children[0].innerText;
				lastAction[tableExcel][lgt]['back'] = backD;
				lastAction[tableExcel][lgt]['cell'] = selectCellExcel[0];
				lastAction[tableExcel][lgt]['indx'] = char+number;

    			selectCellExcel[0].children[0].innerText = text;
	
				socketE.emit('insertInBigData', {table :tableExcel, cell:char+number, data: text, type: 'front'});

				if(tableExcel == 'main'){
				    var tableN = 'cellsMain';
				} else if(tableExcel == 'calc'){
				    var tableN = 'cellsCalc';
				}
				socketE.emit('recalcFunction',{'cell1':char+number, 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})

    		}
    	}
    	if(event.key == 'z' && event.ctrlKey == true){
    		var elemtn = lastAction[tableExcel][Number(Object.keys(lastAction[tableExcel]).length)-1];
    		elemtn['cell'].click();
    		elemtn['cell'].children[0].innerText = elemtn['front'];

    		setValueAtribute(elemtn['cell'].children[0].attributes, 'front',elemtn['front']);
    		socketE.emit('insertInBigData', {table :tableExcel, cell:elemtn['indx'], data: elemtn['front'], type: 'front'});
    		socketE.emit('insertInBigData', {table :tableExcel, cell:elemtn['indx'], data: elemtn['back'], type: 'back'});

    		if(elemtn['back'] && elemtn['back'].indexOf('SELECTLIST') < 0){
    			if(tableExcel == 'main'){
				    var tableN = 'cellsMain';
				} else if(tableExcel == 'calc'){
				    var tableN = 'cellsCalc';
				}
				socketE.emit('recalcFunction',{'cell1':elemtn['indx'], 'cell2':tableN, 'type':'wait', 'tpar':tableExcel})
    		}
    		
    		delete lastAction[tableExcel][Number(Object.keys(lastAction[tableExcel]).length)-1];
    	}
    	if(event.key == 'd' && event.ctrlKey == true){
    		if(selectCellExcel[0].length == 3){
    			selectCellExcel[0].children[2].style.display = 'block';
    			selectCellExcel[0].children[2].style.marginTop = '19px';
    			selectCellExcel[0].children[2].style.position = 'absolute';	
    		}
    	}

    	var chkEditCell = bodyM.getElementsByClassName('editTxtCellExcel');

    	if(event.shiftKey == true && (event.key == 'ArrowLeft' || event.key == 'ArrowRight' || event.key == 'ArrowDown' || event.key == 'ArrowUp') && chkEditCell.length > 0){
    		chkEditCell[0].parentNode.classList.remove('editCellExcel');
    		chkEditCell[0].parentNode.classList.add('selectCellExcel');
    		chkEditCell[0].classList.remove('editTxtCellExcel');
    	}
}

function copyToClipboard(t) {
  const str = t.children[0].innerText
  const el = document.createElement('textarea')
  el.value = str
  el.setAttribute('readonly', '')
  el.style.position = 'absolute'
  el.style.left = '-9999px'
  document.body.appendChild(el)
  el.select()
  document.execCommand('copy')
  document.body.removeChild(el)
}

function calcSummCell(arr){
	var summ = 0;
	if(arr){
		for(var i = 0; i < arr.length ;i++){
			if(isNaN(Number(arr[i].innerText))){
				summ += 0;
			} else {
				summ += Number(arr[i].innerText);

				var n = summ.toString();
					n = n.split('.');
					if(n.length > 1){
						n = n[1];
						n = n.length;
						summ = Number(summ.toFixed(n));	
					}
				
			}		
		}
	}
	var summ_infoSummCellExcell = document.getElementById('summ_infoSummCellExcell');
		summ_infoSummCellExcell.innerText = summ;
}

function renameKeyInArray(arr){
	var temp = [];
	for(var i = 0; i < arr.length;i++){
		if(!arr[i]){
			continue;
		} else {
			temp.push(arr[i]);
		}
	}
	return temp;
}

function colorBlock(t,type,action){
	var bodyM = document.getElementById(tableExcel+'ContentExcel');
	if(action == 'color'){
		if(type == 'line'){
			var item = bodyM.getElementsByClassName('childLine_'+t);
				if(item.length > 0){
					var ckClass = checkClass(item[0].classList,'selectGroupNr');
						if(ckClass == false){
							item[0].classList.add('selectGroupNr');
						}
				}
		}
		if(type == 'group'){
			var item = bodyM.getElementsByClassName('group_'+t);
				if(item.length > 0){
					var ckClass = checkClass(item[0].classList,'selectGroupNr');
						if(ckClass == false){
							item[0].classList.add('selectGroupNr');
						}
				}
		}
	}
	if(action == 'remove'){
		if(type == 'line'){
			var item = bodyM.getElementsByClassName('childLine_'+t);
				if(item.length > 0){
					var ckClass = checkClass(item[0].classList,'selectGroupNr');
						if(ckClass == true){
							item[0].classList.remove('selectGroupNr');
						}
				}
		}
		if(type == 'group'){
			var item = bodyM.getElementsByClassName('group_'+t);
				if(item.length > 0){
					var ckClass = checkClass(item[0].classList,'selectGroupNr');
						if(ckClass == true){
							item[0].classList.remove('selectGroupNr');
						}
				}
		}
	}
}


var resizeGroup = document.getElementsByClassName('resizeGroup');
	for(var i = 0;i <  resizeGroup.length;i++){
		resizeGroup[i].addEventListener('mousedown', initDrag, false);
	}

var resizeLine = document.getElementsByClassName('resizeLine');
	for(var i = 0;i <  resizeLine.length;i++){
		resizeLine[i].addEventListener('mousedown', initDrag, false);
	}


var mousePosition = [0,0];
var doDragBool = false;
var dragObject = '';

document.addEventListener('mousemove', function(e){
	mousePosition[0] = e.clientX;
	mousePosition[1] = e.clientY;

	if(doDragBool == true){
		doDrag(dragObject);
		var type = getValueAtribute(dragObject.attributes, 'grouptype');
		if(type == 'group'){
			document.documentElement.style.cursor = 'col-resize';
		}
		if(type == 'line'){
			document.documentElement.style.cursor = 'row-resize';
		}
	}
});


var startX, startY, startWidth, startHeight;

function initDrag() {
   var t = this.parentNode;

   startX = mousePosition[0];
   startY = mousePosition[1];

   startWidth = parseInt(document.defaultView.getComputedStyle(t).width, 10);
   startHeight = parseInt(document.defaultView.getComputedStyle(t).height, 10);

   doDragBool = true;
   dragObject = t;
   document.addEventListener('mouseup', stopDrag, false);
}

function doDrag(t) {
	var type = getValueAtribute(t.attributes, 'grouptype');
	var bodyM = document.getElementById(tableExcel+'ContentExcel');
		if(type == 'group'){
			var IdGroup = getValueAtribute(t.attributes, 'groupname');
		   	var childrenG = bodyM.getElementsByClassName('childrenGroup_'+IdGroup);
		   	t.style.width = (startWidth + mousePosition[0] - startX) + 'px';
		   	childrenG[0].style.width = (startWidth + mousePosition[0] - startX + 1) + 'px';
		}
		if(type == 'line'){
			t.style.height = (startHeight + mousePosition[1] - startY) + 'px';
			t.style.lineHeight = (startHeight + mousePosition[1] - startY) + 'px';

			var lineNr = getValueAtribute(t.attributes, 'linenr');
			var childrenLine = bodyM.getElementsByClassName('childrenLine_'+lineNr);

			for(var i = 0;i < childrenLine.length ;i++){
				childrenLine[i].style.height = (startHeight + mousePosition[1] - startY - 4) + 'px';
			}
		}
}

function stopDrag() {
	doDragBool = false;
	dragObject = '';
	document.documentElement.style.cursor = 'context-menu';
    document.removeEventListener('mouseup', stopDrag, false);
}

var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');
	infoBackCellDataExcel.addEventListener('focus', focusCellInfoBack);

function focusCellInfoBack(){
	var acceptEditFormTableExcel = document.getElementById('acceptEditFormTableExcel');
	var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');
		var data = getValueAtribute(infoBackCellDataExcel.attributes, 'oldData');
		if(!data){
			infoBackCellDataExcel.setAttribute('oldData',infoBackCellDataExcel.value);
		}
		acceptEditFormTableExcel.style.display = 'flex';
}

var cancelEditFormTableExcel = document.getElementById('cancelEditFormTableExcel');
	cancelEditFormTableExcel.addEventListener('click', cancelEditCell);

function cancelEditCell(){
	var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel');
	var oldData = getValueAtribute(infoBackCellDataExcel.attributes,'olddata');
		infoBackCellDataExcel.value = oldData;	
	var editCellExcel = document.getElementsByClassName('editCellExcel');
		if(editCellExcel.length > 0){
			editCellExcel[0].children[0].innerText = oldData;
			editCellExcel[0].click();
		}
	var acceptEditFormTableExcel = document.getElementById('acceptEditFormTableExcel');
		acceptEditFormTableExcel.style.display = 'none';
}

var yesEditFormTableExcel = document.getElementById('yesEditFormTableExcel');
	yesEditFormTableExcel.addEventListener('click', acceptEditCell);

// var calcExcelFin = [];

function acceptEditCell(){
	var acceptEditFormTableExcel = document.getElementById('acceptEditFormTableExcel');
		acceptEditFormTableExcel.style.display = 'none';

		var editCellExcel = document.getElementsByClassName('editCellExcel');
		if(editCellExcel.length > 0){
			var infoBackCellDataExcel = document.getElementById('infoBackCellDataExcel').value;
			var firstChar = infoBackCellDataExcel.charAt(0);
				var tableN;
		    	if(tableExcel == 'main'){
		    		tableN = 'cellsMain';
		    	} else if(tableExcel == 'calc'){
		    		tableN = 'cellsCalc';
		    	}

		    	var line = getValueAtribute(editCellExcel[0].attributes,'line');
				var group = getValueAtribute(editCellExcel[0].attributes,'group');

				if(firstChar == '='){
					setValueAtribute(editCellExcel[0].attributes,'backdata',infoBackCellDataExcel);
					registerFunctionCell(infoBackCellDataExcel,tableExcel,editCellExcel[0]);

					var cellName = [];
						cellName.push(group+line);
						cellName.push(tableN);
					editCellExcel[0].click();
					socketE.emit('recalcFunction',{'cell1':cellName[0], 'cell2':cellName[1], 'type':'wait', 'tpar':tableExcel})
					// calcExcelFin = [];
				} else {
					editCellExcel[0].click();
					socketE.emit('recalcFunction',{'cell1':group+line, 'cell2':tableN, 'type':'calcBack', 'tpar':tableExcel})
				}
		}
}



///////////////////////////////////////////////////////////////////////////////////////function part//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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
		var  element = document.getElementsByClassName('childrenGroup_'+cells[0][0]);
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


function registerFunctionCell(data,table,funcCell){
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
		        			if(Array.isArray(registerMainCell[cell[i2][0]+cell[i2][1]])){
				        		if(registerMainCell[cell[i2][0]+cell[i2][1]].indexOf(funcAdress) == -1){
				        			registerMainCell[cell[i2][0]+cell[i2][1]].push(funcAdress);
				        		}
				        	} else {
				        		var tmpArr = [];
					        	tmpArr.push(funcAdress);
					        	registerMainCell[cell[i2][0]+cell[i2][1]] = tmpArr;
				        	}
	        		} else if(cell[i2][2] == 'calc'){
	        				if(Array.isArray(registerCalcCell[cell[i2][0]+cell[i2][1]])){
				        		if(registerCalcCell[cell[i2][0]+cell[i2][1]].indexOf(funcAdress) == -1){
				        			registerCalcCell[cell[i2][0]+cell[i2][1]].push(funcAdress);
				        		}
				        	} else {
				        		var tmpArr = [];
					        	tmpArr.push(funcAdress);
					        	registerCalcCell[cell[i2][0]+cell[i2][1]] = tmpArr;
				        	}
	        		}
	        	}   
		}
	}
}


socketE.on('applicateFrontDataInCell', data => {
    var cellName = data.cellName;
    var arrIndex = data.arrIndex;
    var arrIndex2 = data.arrIndex2;
    var dataA = data.data;
    applicateFrontDataInCell(dataA,cellName,arrIndex,arrIndex2);
})

socketE.on('applicateFrontListInCell', data => {
    var cellName = data.cellName;
    var arrIndex = data.arrIndex;
    var arrIndex2 = data.arrIndex2;
    var dataA = data.data;
    var deff = data.deff;
    var backDD = data.backDD;

    applicateFrontListInCell(dataA,cellName,arrIndex,arrIndex2,deff,backDD);
})

function applicateFrontListInCell(data,cellName,arrIndex,arrIndex2,deff,backDD){
	var select = document.createElement('select');
		select.setAttribute('multiple','');
		select.classList.add('listExcelCell');
		for(var i = 0;i < data.length;i++){
			if(data[i]){
				var option = document.createElement('option');
					option.value = data[i];
					option.innerText = data[i];
				select.append(option);
			}
		}

		var button = document.createElement('button');
			button.classList.add('selectListBtnExcel');
			button.innerText = '▼';
			button.style.display = 'none';
			button.addEventListener('click', function(){
				this.parentNode.children[0].classList.remove('slectListExcel');
				this.parentNode.children[0].classList.add('slectListExcel');
				var dataL = this.parentNode.children[2];
				var Sel = this.parentNode.children[2].children[0];
				var options = Sel.options;
				if (dataL.style.display === '') {
				    dataL.style.display = 'block';
				    dataL.style.width = this.parentNode.clientWidth+'px';
				    var ckClass= checkClass(this.parentNode.children[0].classList,'slectListExcel');
				    if(ckClass){
				    	dataL.style.marginTop = this.parentNode.clientHeight+'px';
				    }
				    dataL.style.position = 'absolute';
				} else {
					dataL.style.display = '';
				    dataL.style.width = '';
				    dataL.style.marginTop = '';
				    dataL.style.position = '';
				    this.textContent = "▼";
				}
			})

		var datalist = document.createElement('datalist');
			datalist.classList.add('datalist');
			datalist.append(select)
			
	if(cellName){
		var num = cellName[0].replace(/[^0-9]/g, '');
    	var char = cellName[0].indexOf(num);
        	char = cellName[0].substring(0, char);
        var table = document.getElementById(cellName[1]);
        var group =	table.getElementsByClassName('childrenGroup_'+char)[0];
        var cellHtml = group.getElementsByClassName('childrenLine_'+num)[0];
        	setValueAtribute(cellHtml.attributes, 'backdata', backDD)
			if(cellHtml.children.length > 1){
				cellHtml.children[2].outerHTML ='';
				cellHtml.children[2].append(datalist);
			} else {
				cellHtml.append(button)
				cellHtml.append(datalist)
			}
			cellHtml.children[0].classList.remove('slectListExcel');
			cellHtml.children[0].classList.add('slectListExcel');
			if(deff){
				setValueAtribute(cellHtml.attributes, 'front', deff)
				cellHtml.children[0].innerHTML = deff;
			}
	} else {
		var editCellExcel = document.getElementsByClassName('editCellExcel')[0];
			if(editCellExcel.children.length > 1){
				editCellExcel.children[2].outerHTML ='';
				editCellExcel.children[2].append(datalist);
			} else {
				editCellExcel.append(button)
				editCellExcel.append(datalist)
			}
			editCellExcel.children[0].classList.remove('slectListExcel');
			editCellExcel.children[0].classList.add('slectListExcel');
			if(deff){
				setValueAtribute(editCellExcel.attributes, 'front', deff)
				editCellExcel.children[0].innerHTML = deff;
			}
	}
	select.addEventListener('change', function() {
		var previous = this.parentNode.parentNode.children[0].innerText;
		this.parentNode.parentNode.children[0].innerText = this.value;
		setValueAtribute(this.parentNode.parentNode.children[0].attributes,'front',this.value)
		this.parentNode.style.display = '';
		this.parentNode.style.width = '';
		this.parentNode.style.marginTop = '';
		this.parentNode.style.position = '';
		this.parentNode.parentNode.children[1].innerHTML = "▼";
		this.parentNode.parentNode.children[0].classList.remove('slectListExcel');
		this.parentNode.parentNode.children[0].classList.add('slectListExcel');
				var lineN = getValueAtribute(this.parentNode.parentNode.attributes ,'line');
				var groupN = getValueAtribute(this.parentNode.parentNode.attributes ,'group'); 
				var tableN;
	    		if(cellName[1] == 'cellsMain'){
	    			tableN = 'main';
	    		} else if(cellName[1] == 'cellsCalc'){
	    			tableN = 'calc';
	    		}

	    	if(excelData[tableN][char+num]){
                excelData[tableN][char+num]['front'] = this.value;
	        } else {
		    var tmpObj = {};
		        tmpObj['back'] = '';
		        tmpObj['front'] = this.value;
	        excelData[tableN][char+num] = tmpObj;
	    }
	    socketE.emit('insertInBigData', {table :tableN, cell:char+num, data: this.value, type: 'front','typeIns':'selectlist','tbr':cellName[1]});

	    var backData = getValueAtribute(this.parentNode.parentNode.attributes,'backdata');
	    var lgt = Object.keys(lastAction[tableExcel]).length;
	    lastAction[tableExcel][lgt] = {};
		lastAction[tableExcel][lgt]['front'] = previous;
		lastAction[tableExcel][lgt]['back'] = backData;
		lastAction[tableExcel][lgt]['cell'] = this.parentNode.parentNode;
		lastAction[tableExcel][lgt]['indx'] = char+num;
	})
}

socketE.on('cb_change_selList', data => {
    socketE.emit('recalcFunction',{'cell1':data.cell1, 'cell2':data.cell2, 'type':data.type, 'tpar':data.tpar})
})

/////////////////////////////////////////////////button  tab////////////////////////////////////

var btnsTab = document.querySelectorAll(".btnSelectTabExcel");
for(var iT = 0;iT < btnsTab.length;iT++){
	btnsTab[iT].addEventListener("click", function() {
	  		var type = getValueAtribute(this.attributes,'cell');
	  		if(type == 'main'){
	  			var mainContentExcel = document.getElementById('mainContentExcel');
	  				mainContentExcel.style.display = '-webkit-inline-box';

	  			var calcContentExcel = document.getElementById('calcContentExcel');
	  				calcContentExcel.style.display = 'none';

	  			tableExcel = 'main';
	  		} else if(type == 'calc'){
	  			var mainContentExcel = document.getElementById('mainContentExcel');
	  				mainContentExcel.style.display = 'none';

	  			var calcContentExcel = document.getElementById('calcContentExcel');
	  				calcContentExcel.style.display = '-webkit-inline-box';

	  			tableExcel = 'calc';
	  		}
	  		calcSummCell(shiftSelectCell_visual[tableExcel]);
	  		var activeTabExcel = document.getElementsByClassName('activeTabExcel')[0];
	  			activeTabExcel.classList.add('inactiveTabExcel');
	  			activeTabExcel.classList.remove('activeTabExcel');

	  		this.classList.remove('inactiveTabExcel');
	  		this.classList.add('activeTabExcel');
	});
}

socketE.on('cb_loadFrontData', data => {
	if(firstLoad == false){
		excelData = data;
		var d = excelData.main;
	} else {
		var d = data.main;
	}
	var table = document.getElementById('cellsMain');
	for(var item in d){
		var src = d[item]['back'].indexOf('SELECTLIST');
		if(src == -1){
			var num = item.replace(/[^0-9]/g, '');
	    	var char = item.indexOf(num);
	        	char = item.substring(0, char);
	        var group =	table.getElementsByClassName('childrenGroup_'+char)[0];
	        var cellHtml = group.getElementsByClassName('childrenLine_'+num)[0];
	        	cellHtml.children[0].innerText = d[item]['front']; 
		}
	}
	if(firstLoad == false){
		var	d2 = excelData.calc;
	} else {
		var d2 = data.calc;
	}
	var	table2 = document.getElementById('cellsCalc');
	for(var item in d2){
		// if(d2[item]['front'] != ''){
			var num = item.replace(/[^0-9]/g, '');
	    	var char = item.indexOf(num);
	        	char = item.substring(0, char);
	        var group =	table2.getElementsByClassName('childrenGroup_'+char)[0];
	        var cellHtml = group.getElementsByClassName('childrenLine_'+num)[0];
	        	cellHtml.children[0].innerText = d2[item]['front']; 
        // }
	}
	 if(firstLoad == false){
		loadBackData();
	 }
})
function loadBackData(){
	var d = excelData.main;
	var table = document.getElementById('cellsMain');
	for(var item in d){
		if(d[item]['back']){
			var num = item.replace(/[^0-9]/g, '');
	    	var char = item.indexOf(num);
	        	char = item.substring(0, char);
	        var group =	table.getElementsByClassName('childrenGroup_'+char)[0];
	        var cellHtml = group.getElementsByClassName('childrenLine_'+num)[0];
	        	setValueAtribute(cellHtml.attributes,'backdata',d[item]['back']);
		}
	}
	var d2 = excelData.calc;
	var table2 = document.getElementById('cellsCalc');
	for(var item in d2){
		if(d2[item]['back']){
			var num = item.replace(/[^0-9]/g, '');
	    	var char = item.indexOf(num);
	        	char = item.substring(0, char);
	        var group =	table2.getElementsByClassName('childrenGroup_'+char)[0];
	        var cellHtml = group.getElementsByClassName('childrenLine_'+num)[0];
	        	setValueAtribute(cellHtml.attributes,'backdata',d2[item]['back']);
        }
	}
	if(firstLoad == false){
		firstLoad = true;
	}
}
socketE.on('setStyleCells', data => {
	for(var item of data){
		if(item.table == 'main'){
			var table2 = document.getElementById('cellsMain');
		} else if(item.table == 'calc'){
			var table2 = document.getElementById('cellsCalc');
		}
	    var group =	table2.getElementsByClassName('childrenGroup_'+item.char)[0];
	    var cellHtml = group.getElementsByClassName('childrenLine_'+item.number)[0];

	    cellHtml.setAttribute('style',item.style);

	}
})