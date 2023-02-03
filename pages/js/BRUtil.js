//=====================================================================
//依照BankType的值,顯示銀行清單
//=====================================================================
function loadBankList(frm, frmbank)
{
	var arr, i;
	var sltname, bankelemt;

	i = frm.BankType.selectedIndex;
	sltname = frm.BankType.options[i].value;

	frm.BankListSrc.options.length = 0;
	frm.BankListDst.options.length = 0;

	for (var j =0; j < frmbank.length; j++){
		if ((frmbank.elements[j].type == "select-one") && (frmbank.elements[j].name == sltname)){
			bankelemt = frmbank.elements[j]
			for (var k = 0; k < bankelemt.length; k++){
				frm.BankListSrc.options[k] = new Option(bankelemt.options[k].text, bankelemt.options[k].value);
			}
			break;
		}
	}
}

//=====================================================================
//從Destination刪除一筆資料
//=====================================================================
function deleteOneInDst(ListDst)
{
	for (var i = ListDst.options.length -1; i >= 0 ; i--){
		if ((ListDst.options[i] != null) && (ListDst.options[i].selected) == true)
			ListDst.options[i] = null;
		}
}
//=====================================================================
//新增一筆Source到Destination
//=====================================================================
function addOneToDst(ListSrc, ListDst)
{
	var i, found;
	var arr = new Array();

//	add selected source to array
	for (i = 0; i < ListSrc.options.length; i++){
		found = false;
		if (ListSrc.options[i].selected == true){
			for (var j = 0; j < ListDst.options.length; j++){
				if (ListSrc.options[i].value == ListDst.options[j].value){
					found = true;
					break;
				}
			}
			if (!found){
				arr[arr.length] = ListSrc.options[i].value + '+' + ListSrc.options[i].text;
			}
		}
	}
//	add all destination to array,delete all item then sort array
	if (arr.length > 0){
		for (i = ListDst.options.length -1; i >= 0 ; i--){
			arr[arr.length] = ListDst.options[i].value + '+' + ListDst.options[i].text;
			ListDst.options[i] = null;
		}
		arr.sort();
		var s = arr.join();
		var a = s.split(',');

		for (i = 0; i < a.length; i++){
			var b = a[i].split('+');
			ListDst.options[i] = new Option(b[1], b[0]);
		}
	}
}

//=====================================================================
//新增一筆Source到Destination,但是將新增的資料直接加到Dst的最後
//=====================================================================
function addOneToDst2(ListSrc, ListDst)
{
	var i, found;
	var arr = new Array();

//	add selected source to array
	for (i = 0; i < ListSrc.options.length; i++){
		found = false;
		if (ListSrc.options[i].selected == true){
			for (var j = 0; j < ListDst.options.length; j++){
				if (ListSrc.options[i].value == ListDst.options[j].value){
					found = true;
					break;
				}
			}
			if (!found){
				arr[arr.length] = ListSrc.options[i].value + '+' + ListSrc.options[i].text;
			}
		}
	}
//	add all destination to array,delete all item then sort array
	if (arr.length > 0){
		var s = ListDst.options.length;
//		var a = s.split(',');

		for (i = 0; i < arr.length; i++){
			var b = arr[i].split('+');
			ListDst.options[ListDst.options.length] = new Option(b[1], b[0]);
		}
	}
}

//=====================================================================
//將所有的ListDst中的item放到hidden button(value+text,value+text)
//=====================================================================
function MoveSelectToBtn(btn, ListDst)
{
	btn.value = '';
	for (var i =0; i < ListDst.options.length; i++){
		if (i == 0)
			btn.value = ListDst.options[i].value + '+' + ListDst.options[i].text;
		else
			btn.value = btn.value + ',' + ListDst.options[i].value + '+' + ListDst.options[i].text;
	}	
}
//=====================================================================
//將所定義的報表欄位顯示在畫面上
//=====================================================================
function fn_loadFieldList(form)
{
	var arr, arrFd;
	var table;
	var checkAdd = false;
	var addCount = 0;	
	form.FieldListSrc.options.length = 0;
	//form.FieldListDst.options.length = 0;
	arr = form.FieldList.value.split(',');    
    
	for (var j =0; j < arr.length; j++){
		arrFd = arr[j].split('+');				
		checkAdd=false;		
		for(var i =0;i<form.FieldListDst.length;i++){
			//alert('form.FieldListDst.options[i].value='+form.FieldListDst.options[i].value);		
			//alert('form.FieldListDst.options[i].text='+form.FieldListDst.options[i].text);						
			//alert('arrFd[2]='+arrFd[2]);						
			if(form.FieldListDst.options[i].text == arrFd[2]){		    
			   checkAdd = true;			       
		    }   
	    }
	    if(checkAdd == false){
	       //alert('add item='+arrFd[2]);
	       form.FieldListSrc.options[addCount] = new Option(arrFd[2], arrFd[0] + '.' + arrFd[1]);	 
	       addCount++;
	    }	
	}
}