/*
 * Created on 94.11.24 by lilic0c0 4183
 *
 */
//==========================================================
function doSubmit(form,cnd,parameter){
    alert(form.debtname.value);
	form.action="/pages/FX008AW.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing"; 
    calculusDate(form);
      		
	if(cnd == "Update") {
		if(!checkUpdate(form)) 
	  		return;
	  	
	  	if(!AskUpdate(form))
	    	return;
	    	
		if(!confirm("�Ҧ���Ƴ��T�w�F��"))
    		return;
    	
    	form.submit();
    }
    if(cnd == "load") {
			alert(parameter);		
    }
}	

//==========================================================
function checkUpdate(form) 
{	
	var str1 = "�֭���";
	var str2 = "������";
	
	if( trimString(form.DocNo.value) == "" ){
		alert("�֭�帹���i�ť�");
		form.DocNo.focus();
		return false;
	}
	else if (parseInt(form.Audit_Delay_y.value)>4 ){
		form.Audit_Delay_y.focus();
		return confirm("�֭����W�L4�~�A�T�w�ܡH");
	}
	else if( isNaN(Math.abs(form.ApplyOK_y.value)) ){
		alert("�֭������i����r");
		form.ApplyOK_y.focus();
		return false;
	}
	else if(!checkDate(form.ApplyOK_y,form.ApplyOK_m,form.ApplyOK_d,str1,0)){
		return false;
	}
	else if(checkCurrentDate(form.ApplyOK_y,form.ApplyOK_m,form.ApplyOK_d,str1)){
		return false;
	}
	else if( trimString(form.BOAF_DocNo.value) == "" ){
		alert("����帹���i�ť�");
		form.BOAF_DocNo.focus();
		return false;
	}
	else if( isNaN(Math.abs(form.BOAF_y.value)) ){
		alert("���������i����r");
		form.BOAF_y.focus();
		return false;
	}
	else if(!checkDate(form.BOAF_y,form.BOAF_m,form.BOAF_d,str2,0)){
		return false;
	}
	else if(checkCurrentDate(form.BOAF_y,form.BOAF_m,form.BOAF_d,str2)){
		return false;
	}
	else if( trimString(form.damage_yn.value) == "" &&
		trimString(form.disposal_fact_yn.value) == "" &&
		trimString(form.disposal_plan_yn.value) == "" &&
		trimString(form.auditresult_yn.value) == ""){
			
		form.damage_yn.focus();
		return confirm("�f�ֿﶵ�槡���ťաA�U�����N�i�H�~��ק��ơA�T�w�ܡH");
	}
	else if( trimString(form.auditresult_yn.value) != "Y" &&
		trimString(form.auditresult_yn.value) != "N"){
		alert("�f�ֵ��G���i�ť�");
		form.auditresult_yn.focus();
		return false;
	}
	else if( trimString(form.damage_yn.value) != "Y" &&
		trimString(form.damage_yn.value) != "N"){
		alert("�O�_�����Ʃ�l���l�����i�ť�");
		form.damage_yn.focus();
		return false;
	}
	else if( trimString(form.disposal_fact_yn.value) != "Y" &&
		trimString(form.disposal_fact_yn.value) != "N"){
		alert("�O�_���n���B�����ƹꤣ�i�ť�");
		form.disposal_fact_yn.focus();
		return false;
	}
	
	else if( trimString(form.disposal_plan_yn.value) != "Y" &&
		trimString(form.disposal_plan_yn.value) != "N"){
		alert("���ӳB���p���O�_�X�z�i�椣�i�ť�");
		form.disposal_plan_yn.focus();
		return false;
	}

	else if((trimString(form.damage_yn.value) == "N" ||
			 trimString(form.disposal_fact_yn.value) == "N" ||
			 trimString(form.disposal_plan_yn.value) == "N" )&&
			 trimString(form.auditresult_yn.value) == "Y"      ){
		form.auditresult_yn.focus();
		return confirm("�f�ֿﶵ�榳�ﶵ���uN�v�A�ӡu�f�ֵ��G�v�o���uY�v�A�T�w(Y/N)�H");
	}
	
	else if( trimString(form.damage_yn.value) == "Y" &&
			 trimString(form.disposal_fact_yn.value) == "Y" &&
			 trimString(form.disposal_plan_yn.value) == "Y" &&
			 trimString(form.auditresult_yn.value) == "N"){
		form.auditresult_yn.focus();
		return confirm("�f�ֿﶵ�槡���uY�v�A�ӡu�f�ֵ��G�v�o���uN�v�A�T�w(Y/N)�H");
	}
	
	return true;

}

//==========================================================
function loadUpperDate(form){
	form.Audit_Delay_y.value = form.apply_year.value;
	form.Audit_Delay_m.value = form.apply_month.value;
}

//==========================================================
function calculusDate(form){
	
	var accept_y = parseInt(form.accept_year.value);
	var accept_m = parseInt(form.accept_month.value);
	var accept_d = parseInt(form.accept_day.value);
	var delay_y = parseInt(form.Audit_Delay_y.value);
	var delay_m = parseInt(form.Audit_Delay_m.value);
	var result_y = 0; 
	var result_m = 0;
	var result_d = 0;
	
	result_y = accept_y + delay_y;
	result_m = accept_m + delay_m;
	result_d = accept_d;
	 
	if(result_m > 12){//����W�L�@�~
		result_m-=12;
		result_y++;
	}
	
	//�p�G��X�ӬO31���o�S�J�����O�p�� 
	if(result_d==31){
		
		switch(result_m){
			case 2:
				result_d=28;
				if(leapYear(parseInt(result_y,10)+1911))
					result_d=29;
				break;
			case 4:
			case 6:
			case 9:
			case 11:
				result_d=30;
				break;
		}//end of switch
	}//end of if
	
	form.Cnt_year.value	= result_y;
	form.Cnt_month.value= result_m;
	form.Cnt_date.value	= result_d;
}

//====�ۿ��̪�@����������======================================================
function LoadRecentData(form,xml){
	
	var xmlDoc = document.all(xml).XMLDocument; //���oxml data
	var flag =  xmlDoc.getElementsByTagName("flag");
	
	if(flag.item(0) == null){
		alert("�u�L�̪�@����ơA�L�q�ۿ��Цۦ��J�C�v");
		form.DocNo.focus();
		return;
	}
	else{//�i��ۿ�
		//���h���oxml������
		var docno =  xmlDoc.getElementsByTagName("docno");
		var boafdocno =  xmlDoc.getElementsByTagName("boafdocno");
		var apoky =  xmlDoc.getElementsByTagName("apoky");
		var apokm =  xmlDoc.getElementsByTagName("apokm");
		var apokd =  xmlDoc.getElementsByTagName("apokd");
		var boafy =  xmlDoc.getElementsByTagName("boafy");
		var boafm =  xmlDoc.getElementsByTagName("boafm");
		var boafd =  xmlDoc.getElementsByTagName("boafd");
		
		//���ۿ����ʧ@
		form.DocNo.value = docno.item(0).firstChild.nodeValue;
		form.BOAF_DocNo.value = boafdocno.item(0).firstChild.nodeValue;
		
		form.ApplyOK_y.value = apoky.item(0).firstChild.nodeValue;
		setSelect(form.ApplyOK_m,apokm.item(0).firstChild.nodeValue);
		setSelect(form.ApplyOK_d,apokd.item(0).firstChild.nodeValue);
		form.BOAF_y.value = boafy.item(0).firstChild.nodeValue;
		setSelect(form.BOAF_m,boafm.item(0).firstChild.nodeValue);
		setSelect(form.BOAF_d,boafd.item(0).firstChild.nodeValue);
	}

}
//=========================================================
function checkMonth(month){
	if(month.value < 0 ){		
		alert("������i�p��s");
		return false;
	}
	else if(month.value > 12){
		alert("������i�j��12�Ӥ�");
		return false;
	}

	return true;

}
//=========================================================
function setSelect(S1,bankid) {

    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==bankid)    	{
        	S1.options[i].selected=true;
        	break;
    	}
    }
}

//=========================================================

function checkCurrentDate(year,month,day,txt){
	var today = new Date();
	var nowYear = today.getYear()-1911; //change Anno Domini(A.D.) to ROC date
	var nowMonth = today.getMonth()+1;//begin with 0
	var nowDate = today.getDate();
	var totalNowDate = nowYear*12*30+nowMonth*30+nowDate+0;
	var totalTestDate = parseInt(year.value,10)*12*30+parseInt(month.value,10)*30+parseInt(day.value,10)+0;
	
	//alert( totalTestDate);
	if(totalTestDate > totalNowDate){
		alert("��J���y"+txt+"�z�j�󥻤����A�ЦA���T�{");
		return true;
	}
	
	return false;
}

	