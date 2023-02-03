var mainPg = "/pages/TM002W.jsp";

function goEdit(form,acc_Tr_Type,acc_Tr_Name){
	form.action= mainPg+"?act=Edit&acc_Tr_Type="+acc_Tr_Type+"&acc_Tr_Name="+acc_Tr_Name+"&test=nothing";	
	form.submit();
}
function AskReset(form){
	if(confirm("確定要回復原有資料嗎?")){
		form.acc_Tr_Name.value = "";
		form.itemListDst1.options.length = 0;
		form.itemListDst2.options.length = 0;
		form.BankListDst.options.length = 0;
		fn_changeListSrc('01');
		fn_changeListSrc('02');
		fn_changeListSrc('03');
	}
}
function doSubmit(form,fun) {
	flg = true;
	if('Update'==fun){
		if(form.acc_Tr_Name.value==''){
			alert("請輸入協助措施名稱");
			flg = false;
		}
		getSelListDst(this.document.forms[0].selSubList1, this.document.forms[0].itemListDst1);
		getSelListDst(this.document.forms[0].selSubList2, this.document.forms[0].itemListDst2);
		getSelListDst(this.document.forms[0].selBankList, this.document.forms[0].BankListDst);
		if(form.selSubList1.value=='' && form.selSubList2.value==''){
			alert("請選擇舊貸展延需求或新貸需求");
			flg = false;
		}
	}
	if(flg){
		form.action= mainPg+"?act="+fun+"&test=nothing";		
		form.submit();	
	}
	
}
function getSelListDst(sel, ListDst){
	sel.value = '';
	for (var i =0; i < ListDst.options.length; i++){
		if (i == 0)
			sel.value = ListDst.options[i].value;
		else
			sel.value = sel.value + ';' + ListDst.options[i].value;
	}	
}