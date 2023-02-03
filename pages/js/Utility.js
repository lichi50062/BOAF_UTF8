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
//新增全部的Source到Destination
//=====================================================================
function gfn_addALLToDst(ListSrc, ListDst)
{
	ListDst.options.length == 0;
	for (i = 0; i < ListSrc.options.length; i++){
		ListDst.options[i] = new Option(ListSrc.options[i].text, ListSrc.options[i].value);
		}
}
//=====================================================================
//移除Destination全部的item
//=====================================================================
function gfn_deleteALLInDst(ListDst)
{
	ListDst.options.length = 0;
}

//=====================================================================
//從Destination刪除一筆資料
//=====================================================================
function gfn_deleteOneInDst(ListDst)
{
	for (var i = ListDst.options.length -1; i >= 0 ; i--){
		if ((ListDst.options[i] != null) && (ListDst.options[i].selected) == true)
			ListDst.options[i] = null;
		}
}
//=====================================================================
//新增一筆Source到Destination
//=====================================================================
function gfn_addOneToDst(ListSrc, ListDst)
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
function gfn_addOneToDst2(ListSrc, ListDst)
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
//將所有的ListDst中的item放到hidden button
//=====================================================================
function gfn_MoveSelectToBtn(btn, ListDst, strDelimiter)
{
	btn.value = '';
	for (var i =0; i < ListDst.options.length; i++){
		if (i == 0)
			btn.value = strDelimiter + ListDst.options[i].value + strDelimiter;
		else
			btn.value = btn.value + ',' + strDelimiter + ListDst.options[i].value + strDelimiter;
	}
}
//=====================================================================
//將所有的ListDst中的item放到hidden button(value+text,value+text)
//=====================================================================
function gfn_MoveSelectToBtn2(btn, ListDst)
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
//將所有的ListDst中的item放到hidden button(value+text+index,value+text+index)
//this fun for SortList use,把欄位出現的位置也存起來供Order by 1,3,4使用
//=====================================================================
function gfn_MoveSelectToBtn3(btn, ListSrc, ListDst)
{
	var idx;
	btn.value = '';
	for (var i =0; i < ListDst.options.length; i++){
		for (var j=0; j < ListSrc.options.length; j++){
			if (ListSrc.options[j].value == ListDst.options[i].value){
				idx = j + 1;
				break;
			}
		}
		if (i == 0)
			btn.value = ListDst.options[i].value + '+' + ListDst.options[i].text + '+' + idx ;
		else
			btn.value = btn.value + ',' + ListDst.options[i].value + '+' + ListDst.options[i].text + '+' + idx;
	}
}
//=====================================================================
//將字串中的空白移除
//=====================================================================
//function trimString(inString)
//{
//	var outString;
//	var startPos;
//	var endPos;
//	var ch;
//
//	// where do we start?
//	startPos = 0;
//	test = 0;
//	ch = inString.charAt(startPos);
//	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
//		startPos++;
//		if ( ch==" " ) {
//			test++;
//		}
//		ch = inString.charAt(startPos);
//	}
//     if  ( test==inString.length )
//     	flag = true;
//     else
//     	flag = false;
//	endPos = inString.length - 1;
//	ch = inString.charAt(endPos);
//	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
//		endPos--;
//		ch = inString.charAt(endPos);
//	}
//
//	// get the string
//	outString = inString.substring(startPos, endPos + 1);
//	if ( flag==true )
//		return "";
//	else
//		return outString;
//}
function gfn_Up(ListDst){

	var oldPos, newPos, value, text, oldLen;
	var arr = new Array();

	if ((ListDst.selectedIndex == -1) || (ListDst.selectedIndex == 0)){
		return;}
	oldLen = ListDst.length;
	oldPos = ListDst.selectedIndex;
	newPos = ListDst.selectedIndex - 1;
	value = ListDst.options[oldPos].value;
	text = ListDst.options[oldPos].text;

	ListDst.options[oldPos] = null;
	var j =0;
	for (var i=0; i < oldLen; i++){
		if (i == newPos){
			arr[i] = value + ',' + text;
		}
		else{
			arr[i] = ListDst.options[j].value + ',' + ListDst.options[j].text;
			j ++;
		}
	}
	ListDst.options.length = 0;
	var arrNew;

	for (i=0; i < arr.length; i++){
		arrNew = arr[i].split(',');
		ListDst.options[i] = new Option(arrNew[1], arrNew[0]);
		}
	ListDst.selectedIndex = newPos
}
function gfn_Down(ListDst){

	var oldPos, newPos, value, text, oldLen;
	var arr = new Array();

	if ((ListDst.selectedIndex == -1) || (ListDst.selectedIndex == ListDst.length -1)){
		return;}
	oldLen = ListDst.length;
	oldPos = ListDst.selectedIndex;
	newPos = ListDst.selectedIndex + 1;
	value = ListDst.options[oldPos].value;
	text = ListDst.options[oldPos].text;

	ListDst.options[oldPos] = null;
	var j =0;
	for (var i=0; i < oldLen; i++){
		if (i == newPos){
			arr[i] = value + ',' + text;
		}
		else{
			arr[i] = ListDst.options[j].value + ',' + ListDst.options[j].text;
			j ++;
		}
	}
	ListDst.options.length = 0;
	var arrNew;

	for (i=0; i < arr.length; i++){
		arrNew = arr[i].split(',');
		ListDst.options[i] = new Option(arrNew[1], arrNew[0]);
		}
	ListDst.selectedIndex = newPos
}