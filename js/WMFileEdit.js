// 94.01.04 add �W�[bank_type by 2295
// 94.03.15 add M06,M07 �p���I���p�p��Ʀ۰ʭp��B�z  by EGG 
// 94.03.17 add 91060P=910400/910500����p�Ʋ�2��(�ĤT��|�ˤ��J) by 2295
// 94.03.17 add �֥[���I�ʸ겣�B��910500 by 2295
// 94.07.14 fix acc_code=99141Y�̪�M��~�ץu���J3��,�[�Wacc_code�ӧP�_  			
// 94.07.20 add 920710.920720.920730.920740.920750 �]�@�֭p�⭷�I�v�� by 2295
// 94.11.15 add �W���ƶפJ�B�O�_�T�w�欰'Y'��,�~�פJ�W���� by 2295
// 95.01.17 fix Insert/Update/Delete�[�Wbank_type by 2295
// 95.04.12 add A06 970000�X�p-->�۰ʲ֥[ by 2295
// 			   �p��O���v by 2295
// 95.05.11 add ��910203(2_1)�S�w�l���Ҵ��C���Ʃ�b�b�B�l���ǳƤ���~�ǳ�		
// 95.05.16 fix �Y��120601/120602/120603/120604���֥[ by 2295		   	
//          fix �Y�X�{�ƾǸ�e��,��v�H0.000���N by 2295
// 95.05.16 add �O��X�p���B���i�j��A01��ڦX�p by 2295
// 95.05.26 add A99 check 992300�H�γ����u�H�� by 2295
// 95.05.29 add 990220 <= 990210 by 2295
// 95.06.01 add �s�W/�ק��,���n�ˬd992300�H�γ����u�H�� by 2295
// 95.06.01 add 970000�U�O����ƦX�p�Ƥ��i�p�� 0 by 2295
// 95.06.01 add 150200/960500/970000���Τ���O�_��A01��ڦX�p�p by 2295
// 95.06.01 970000.A01����"0",�~�p��O���v by 2295
// 95.06.01 960500��� by 2295
// 95.06.01 fix 910500���I�ʸ겣�B������.�|�B���J	
// 95.06.20 fix �O���v��ܦ�m���� by 2295
// 95.08.10 fix ����960500.�X�p��m16�令15 by 2295
// 95.08.21 fix A01���ƧC.�β֭p���¤@�߬����� by 2295
// 96.07.10 add A08.�s�ڱb���`���(E)�۰ʲ֥[ by 2295
// 96.12.10 �۰ʭp�������`�B����v(A-B)/(C-D) by 2295
// 97.06.13 add A10.�۰ʭp��X�p by 2295
// 97.07.07 fix �Yc-d���s��,��v��0 by 2295
//100.06.27 fix A99.992110�A(��)�|����b�ȥ����ӳ�,�B���i��0 by 2295
//100.08.04 fix A99.992100�A(��)�|����겣�`�B�����ӳ�,�B���i��0/�t�� by 2295
//101.06.29 add 992300�H�γ����u�H�ƻݻP�`���c�q�ƫH�η~�ȭ��u�H�Ƥ@�P,�~�i�s�W by 2295
//101.08.16 add A02.�W�[990811/990812/990813/990814 by 2295
function fillText(form,val){	
	form.amt320300.value=val;	
}	

//94.03.16 add �֥[���I�ʸ겣�B��910500 by 2295
//95.5.11 add ��910203(2_1)�S�w�l���Ҵ��C���Ʃ�b�b�B�l���ǳƤ���~�ǳ�		
//95.06.01 fix 910500���I�ʸ겣�`�B������.�|�B���J
function addText(form){	          
    //alert('addText begin'); 
     		
	var a = 0;		
	if(form.assumed_amt[0].value != '' && !isNaN(changeVal(form.assumed_amt[0]))){//920101	
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[0])));	   	   
	   //alert('920101='+a);
	   form.amt[0].value=changeStr(form.amt[0]);
	   //alert(form.amt[0].value);
	   //alert('a='+a);
	}		
	if(form.assumed_amt[1].value != '' && !isNaN(changeVal(form.assumed_amt[1]))){//920201		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[1])));	   
	   //alert('920201='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[2].value != '' && !isNaN(changeVal(form.assumed_amt[2]))){//920301	
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[2])));
	   //alert('920301='+a);
	   //alert('a='+a);
	}
	
	if(form.assumed_amt[3].value != '' && !isNaN(changeVal(form.assumed_amt[3]))){//920401
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[3])));
	   //alert('920401='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[4].value != '' && !isNaN(changeVal(form.assumed_amt[4]))){//920501		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[4])));
	   //alert('920501='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[5].value != '' && !isNaN(changeVal(form.assumed_amt[5]))){//920601		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[5])));
	   //alert('920601='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[6].value != '' && !isNaN(changeVal(form.assumed_amt[6]))){//920710		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[6])));
	   //alert('920710='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[9].value != '' && !isNaN(changeVal(form.assumed_amt[9]))){//920720		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[9])));
	   //alert('920720='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[12].value != '' && !isNaN(changeVal(form.assumed_amt[12]))){//920730		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[12])));
	   //alert('920730='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[15].value != '' && !isNaN(changeVal(form.assumed_amt[15]))){//920740		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[15])));
	   //alert('920740='+a);
	   //alert('a='+a);
	}
	if(form.assumed_amt[18].value != '' && !isNaN(changeVal(form.assumed_amt[18]))){//920750		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[18])));
	   //alert('920750='+a);
	   //alert('a='+a);
	}
	
	if(form.assumed_amt[21].value != '' && !isNaN(changeVal(form.assumed_amt[21]))){//920801		
	   a = a + parseFloat(eval(changeVal(form.assumed_amt[21])));
	   //alert('920801='+a);
	   //alert('a='+a);
	}
	if(form.amt[36].value != '' && !isNaN(changeVal(form.amt[36]))){//910203(2_1)�S�w�l���Ҵ��C���Ʃ�b�b�B�l���ǳƤ���~�ǳ�		
	   //alert('910203='+form.amt[36].value);
	   a = a - parseFloat(eval(changeVal(form.amt[36])));
	   //alert('920801='+a);
	   //alert('a='+a);
	}	
	
	//alert('a='+a);	
	//alert('a.round='+Math.round(a));
	//97.03.14 add 910204����910400.910500.91060p��m
	form.amt[45].value= Math.round(a) ;	//95.06.01 fix ������.�|�B���J910500	
	form.amt[45].value=changeStr(form.amt[45]);
	//alert('form.amt[42].value='+form.amt[42].value);
    
    //94.03.17 add 91060P=910400/910500����p�Ʋ�2��(�ĤT��|�ˤ��J) by 2295
    var tmp = 0;        
    //alert('form.amt[42].value='+form.amt[42].value);
    if((form.amt[44].value != '' && !isNaN(changeVal(form.amt[44]))) && (form.amt[45].value != '' && form.amt[45].value != '0')){//910500	    
    	//alert('count 91060P');    	
    	//alert(parseFloat(eval(changeVal(form.amt[41]))) / parseFloat(eval(changeVal(form.amt[42]))));
        tmp = parseFloat(eval(changeVal(form.amt[44]))) / parseFloat(eval(changeVal(form.amt[45])));    
        //alert(tmp);
        tmp = Math.round(tmp * 100000);    
        tmp = tmp / 1000;
        form.amt[46].value=tmp    
    }	
}


//94.03.11 add �֥[��910500 by 2295
/*
function addText(form){	
     
    94.03.10 have ??
	i=0;
	assum=0;
	for(i=0;i<form.assumed_amt.length;i++){
		tmpstr=changeVal(form.assumed_amt[i]);
		if(tmpstr.length!=0 && !isNaN(tmpstr)){
			assum+=parseInt(tmpstr);
		}
	}
	form.amt[42].value=assum;
	form.amt[42].value=changeStr(form.amt[42]);
	 
     
     		
	var a =0;		
	if(form.amt[0].value != '' && !isNaN(changeVal(form.amt[0]))){//920101	
	   a = a + parseInt(eval(changeVal(form.amt[0])));	   
	   //alert('920101='+a);
	}		
	if(form.amt[1].value != '' && !isNaN(changeVal(form.amt[1]))){//920201		
	   a = a + parseInt(eval(changeVal(form.amt[1])));	   
	   //alert('920201='+a);
	}
	if(form.amt[2].value != '' && !isNaN(changeVal(form.amt[2]))){//920301	
	   a = a + parseInt(eval(changeVal(form.amt[2])));
	   //alert('920301='+a);
	}
	if(form.amt[3].value != '' && !isNaN(changeVal(form.amt[3]))){//920401
	   a = a + parseInt(eval(changeVal(form.amt[3])));
	   //alert('920401='+a);
	}
	if(form.amt[4].value != '' && !isNaN(changeVal(form.amt[4]))){//920501		
	   a = a + parseInt(eval(changeVal(form.amt[4])));
	   //alert('920501='+a);
	}
	if(form.amt[5].value != '' && !isNaN(changeVal(form.amt[5]))){//920601		
	   a = a + parseInt(eval(changeVal(form.amt[5])));
	   //alert('920601='+a);
	}
	if(form.amt[6].value != '' && !isNaN(changeVal(form.amt[6]))){//920710		
	   a = a + parseInt(eval(changeVal(form.amt[6])));
	   //alert('920710='+a);
	}
	if(form.amt[9].value != '' && !isNaN(changeVal(form.amt[9]))){//920720		
	   a = a + parseInt(eval(changeVal(form.amt[9])));
	   //alert('920720='+a);
	}
	if(form.amt[12].value != '' && !isNaN(changeVal(form.amt[12]))){//920730		
	   a = a + parseInt(eval(changeVal(form.amt[12])));
	   //alert('920730='+a);
	}
	if(form.amt[15].value != '' && !isNaN(changeVal(form.amt[15]))){//920740		
	   a = a + parseInt(eval(changeVal(form.amt[15])));
	   //alert('920740='+a);
	}
	if(form.amt[18].value != '' && !isNaN(changeVal(form.amt[18]))){//920750		
	   a = a + parseInt(eval(changeVal(form.amt[18])));
	   //alert('920750='+a);
	}
	
	if(form.amt[21].value != '' && !isNaN(changeVal(form.amt[21]))){//920801		
	   a = a + parseInt(eval(changeVal(form.amt[21])));
	   //alert('920801='+a);
	}
	//alert('a='+a);
	form.amt[42].value= a ;	//910500
	//alert('form.amt[42].value='+form.amt[42].value);
	
}
*/
function doSubmitStatus(form,Report_no, Report_name,bank_type) {		
	//alert(form);
	//var form = document.frmWMFileEdit;
	
	//�Ybank_type="2" && Report_no == 'A11'�A���ɨ�Bank_List.jsp ���obank_no �A�A�ɦ^WMFileEdit_A11.jsp?act=List
	if(Report_no == 'A11'){
		//if(bank_type=='2'){
			//�A�����i�J
		//	form.action="/pages/WMFileEdit_A11.jsp?act=List";
	//	}else{
			form.action="/pages/WMFileEdit_A11.jsp?act=List&bank_type="+bank_type;	
	//	}
	}else{
		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";	
	}
	form.submit();
}

//95.01.17 fix Insert/Update/Delete�[�Wbank_type by 2295
function doSubmit(form,fun,Report_no,S_YEAR,S_MONTH,bank_type) {	
	form.act.value=fun;			
	//alert(fun)
	if(fun == 'Edit'){	   
	   form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&bank_code=123456&test=nothing";		
	   form.submit();	   
	}
	if(fun == 'new'){	   	   
	   if(Report_no == 'F01' && !checkShowInsert_F01(form, fun)) return;//2005.11.9 add F01�bnew��.�ˮ֦~�� by 2295	   	  	   	
	   form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	   form.submit();	   
	}
	if(fun == 'getLastMonthData' && (Report_no == 'F01' || Report_no == 'A06' || Report_no == 'A10')){	   	          
	   //94.11.15 add �W���ƶפJ�B�O�_�T�w�欰'Y'��,�~�פJ�W����	
	   if(confirm("�O�_�T�w(�O/�_)??")){	   	  	   	   
	   	  form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	      form.submit();	   	   	
	   }else{
	   	  return;
 	   }	
    }	
	if(fun == 'Insert'){		
		if(Report_no == 'A01' && (checkShowInsert_A01(form, fun))){/* modify block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'A02' && (checkShowInsert_A02(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'A03' && (checkShowInsert_A03(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'A04' && (checkShowInsert_A04(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'A05' && (checkShowInsert_A05(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'A06' && (checkShowInsert_A06(form, fun))){/* Add block by 2295 2006.04.10 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}	
	   	}else if(Report_no == 'A08' && (checkShowInsert_A08(form, fun))){/* Add block by 2295 2007.07.10 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}	
	   	}else if(Report_no == 'A09' && (checkShowInsert_A09(form, fun))){/* Add block by 2295 2007.12.10 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}	
	   	}else if(Report_no == 'A10' && (checkShowInsert_A10(form, fun))){/* Add block by 2295 2008.06.13 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}			
	   	}else if(Report_no == 'A99' && (checkShowInsert_A99(form, fun))){/* Add block by 2295 2006.05.18 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}		
	   	}else if(Report_no == 'B01' && (checkShowInsert_B01(form, fun))){/* Add block by jei 2004.12.14 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'B02' && (checkShowInsert_B02(form, fun))){/* Add block by jei 2004.12.14 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'B03' && (checkShowInsert_B03(form, fun))){/* Add block by egg 2004.12.15 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M01' && (checkShowInsert_M01(form, fun))){/* Add block by egg 2004.12.10 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M02' && (checkShowInsert_M02(form, fun))){/* Add block by jei 2004.12.10 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M03' && (checkShowInsert_M03(form, fun))){/* Add block by egg 2004.12.11 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M04' && (checkShowInsert_M04(form, fun))){/* Add block by jei 2004.12.11 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}	
	   	}else if(Report_no == 'M05' && (checkShowInsert_M05(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M06' && (checkShowInsert_M06(form, fun))){/* Add block by egg 2004.12.13 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M07' && (checkShowInsert_M07(form, fun))){/* Add block by egg 2004.12.13 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'M08' && (checkShowInsert_M08(form, fun))){/* Add block by egg 2004.12.16 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}else if(Report_no == 'F01' && (checkShowInsert_F01(form, fun))){/* add block by 2295 2005.11.09 */
	   		if(AskInsert(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	   		}
	   	}
	}
	if(fun == 'Update') {		
		if( Report_no == 'A01' && (checkShowInsert_A01(form, fun))){/* modify block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'A02' && (checkShowInsert_A02(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'A03' && (checkShowInsert_A03(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'A04' && (checkShowInsert_A04(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'A05' && (checkShowInsert_A05(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	    }else if( Report_no == 'A06' && (checkShowInsert_A06(form, fun))){/* Add block by 2295 2006.04.10 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}   
	    }else if( Report_no == 'A08' && (checkShowInsert_A08(form, fun))){/* Add block by 2295 2007.07.10 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}     		
	    }else if( Report_no == 'A09' && (checkShowInsert_A09(form, fun))){/* Add block by 2295 2007.12.10 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}    
	    }else if( Report_no == 'A10' && (checkShowInsert_A10(form, fun))){/* Add block by 2295 2008.06.13 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}      	 		   	
	    }else if( Report_no == 'A99' && (checkShowInsert_A99(form, fun))){/* Add block by 2295 2006.05.18 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}      	
	   	}else if( Report_no == 'B01' && (checkShowInsert_B01(form, fun))){/* Add block by jei 2004.12.14 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'B02' && (checkShowInsert_B02(form, fun))){/* Add block by jei 2004.12.14 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'B03' && (checkShowInsert_B03(form, fun))){/* Add block by egg 2004.12.15 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M01' && (checkShowInsert_M01(form, fun))){/* Add block by egg 2004.12.10 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M02' && (checkShowInsert_M02(form, fun))){/* Add block by jei 2004.12.10 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M03' && (checkShowInsert_M03(form, fun))){/* Add block by egg 2004.12.13 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	    }else if( Report_no == 'M04' && (checkShowInsert_M04(form, fun))){/* Add block by jei 2004.12.13 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M05' && (checkShowInsert_M05(form, fun))){/* Add block by Winnin.chu 2004.11.24 ver1.0*/
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M06' && (checkShowInsert_M06(form, fun))){/* Add block by egg 2004.12.13 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}else if( Report_no == 'M07' && (checkShowInsert_M07(form, fun))){/* Add block by egg 2004.12.13 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
		}else if( Report_no == 'M08' && (checkShowInsert_M08(form, fun))){/* Add block by egg 2004.12.16 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
		}else if( Report_no == 'F01' && (checkShowInsert_F01(form, fun))){/* add block by 2295 2005.11.9 */
	   		if(AskUpdate(form)){	
	       		form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       		form.submit();
	       	}
	   	}	   	
	}	
	if(fun == 'Delete'){
		if(AskDelete(form)){	
	       form.action="/pages/WMFileEdit.jsp?act="+form.act.value+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type+"&test=nothing";		
	       form.submit();
	   } 
	}	
}
//95.08.21 fix �ƧC.�β֭p���¤@�߬����� by 2295
function checkShowInsert_A01(form, fun) {		
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}   
    for (var i = 0; i < form.amt.length; i++) {
    	if(!checkNumber(form.amt[i])){
		   return false;
	    }
    	//�ƧC.�β֭p���¤@�߬�����
    	//alert(form.acc_name[i].value);
    	//alert(form.acc_name[i].value.indexOf("�Ʃ�"));
    	if((form.acc_name[i].value.indexOf("�Ʃ�") != -1) || (form.acc_name[i].value.indexOf("�֭p����") != -1)){    	    	
	        if(form.amt[i].value.indexOf("-") != -1){
	           form.amt[i].focus();	
	           alert(form.acc_name[i].value+'���i���t��!!');	
	           return false;
	        }	
        }
    }    	
	return true;
}


/* Add Method by Winnin.chu 2004.11.24 ver1.0*/
//95.05.29 add 990220 <= 990210 by 2295
//95.06.01 add 990510=990511+990512 by 2295
//101.08.16 add 990811/990812/990813/990814 by 2295
function checkShowInsert_A02(form, fun) {
	var amt990210=0;
	var amt990220=0;
	var amt990510=0;
	var amt990511=0;
	var amt990512=0;	
	var amt990811=0;
	var amt990812=0;
	var amt990813=0;
	var amt990814=0;
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i = 0; i < form.amt.length; i++) {
		if (!checkNumber(form.amt[i]))
			return false;
		//94.07.14 fix acc_code=99141Y�̪�M��~�ץu���J3��  			
		if (  ((i==34 && form.acc_code[i].value=='99141Y') || (i==39 && form.acc_code[i].value=='99141Y')) && (parseInt(form.amt[i].value)<0 || parseInt(form.amt[i].value)>999)){  //�̫ᵲ��~��
			//alert(form.amt[i].value);
			//alert(form.acc_code[i].value);
			alert('�̪�M��~�ץu���J3��(����~)');
			return false;
		}
		
		if( i==5 && form.acc_code[i].value=='990210' && trimString(form.amt[i].value) != "") {
			amt990210 = parseInt(form.amt[i].value); 
	    }	
		if( i==6 && form.acc_code[i].value=='990220' && trimString(form.amt[i].value) != ""){
			amt990220 = parseInt(form.amt[i].value); 			
	    }	
	    
	    //95.06.01 =====================================
		if( form.acc_code[i].value=='990510' && trimString(form.amt[i].value) != ""){			
			amt990510 = parseInt(form.amt[i].value); 
	    }	
		if( form.acc_code[i].value=='990511' && trimString(form.amt[i].value) != ""){
			amt990511 = parseInt(form.amt[i].value); 
	    }
	    if( form.acc_code[i].value=='990512' && trimString(form.amt[i].value) != ""){
			amt990512 = parseInt(form.amt[i].value); 
	    }
	    //101.08.16 add 990811/990812/990813/990814============
	    if(form.acc_code[i].value=='990811'){
	    	if(form.chk990811.checked == true){ 
	    		form.amt[i].value = "1";
	    		amt990811 = 1;
	    	}else{
	    		form.amt[i].value = "0";
	    		amt990811 = 0;
	    	}
		}	
		if(form.acc_code[i].value=='990812'){
	    	if(form.chk990812.checked == true){ 
	    		form.amt[i].value = "1";
	    		amt990812 = 1;
	    	}else{
	    		form.amt[i].value = "0";
	    		amt990812 = 0;
	    	}
		}	
		if(form.acc_code[i].value=='990813'){
	    	if(form.chk990813.checked == true){ 
	    		form.amt[i].value = "1";
	    		amt990813 = 1;
	    	}else{
	    		form.amt[i].value = "0";
	    		amt990813 = 0;
	    	}
		}
		if(form.acc_code[i].value=='990814'){
	    	if(form.chk990814.checked == true){ 
	    		form.amt[i].value = "1";
	    		amt990814 = 1;
	    	}else{
	    		form.amt[i].value = "0";
	    		amt990814 = 0;
	    	}
		}
		//=================================================================		
	}

	//95.05.29 add 990220 <= 990210 =====================================		
	if( amt990220 > amt990210){
	    alert('�i990220�j�ݤp(��)��i990210�j\n�ˮֵo�{990220('+amt990220+'��)�j��990210('+amt990210+'��)');
	    return false;
	}	
	//==================================================================
	//alert('990510='+amt990510);
	//alert('990511='+amt990511);
	//alert('990512='+amt990512);
	//95.06.01 add 990510=990511+990512 =====================================		
	if( amt990510 != 0 || amt990511 != 0 || amt990512 != 0){	  	
	   	amt990511 = amt990511 + amt990512;
	   	if(amt990510 != amt990511){
	       alert('�i990510�j=�i990511�j+�i990512�j\n�ˮֵo�{990510(��'+amt990510+'��)������990511+990512('+amt990511+'��)');
	       return false;
	    }
	}	
	//==================================================================
	//101.08.16 add 990811�Y���Ŀ��,���i�Ŀ�990812/990813/990814============
	if(amt990811 == 1 && (amt990812 == 1 || amt990813 == 1 || amt990814 ==1)){
      alert('�Y�w�Ŀ�H�γ��T�w�겣�b�ȡA�L�W�L��b��,�h���i�Ŀ�H�γ��T�w�겣�b�B�A���o�W�L��b�Ȧ��U�C���Τ��@���ﶵ!!');
      return false;
    }	
	//=======================================================================
	return true;
}

/* Add Method by Winnin.chu 2004.11.24 ver1.0*/
function checkShowInsert_A03(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}

	return true;
}

/* Add Method by Winnin.chu 2004.11.24 ver1.0*/
function checkShowInsert_A04(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}

	return true;
}

/* Add Method by Winnin.chu 2004.11.24 ver1.0*/
function checkShowInsert_A05(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}
//95.06.01 add 970000�U�O����ƦX�p�Ƥ��i�p�� 0 by 2295
function checkShowInsert_A06(form, fun) {			
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && (!checkNumber(form.elements[i]))){
			return false;
		}	
	}
	if(parseInt(form.amt_3month[15].value) < 0){
	   alert('�ˮֵo�{����3�Ӥ몺�X�p���B('+form.amt_3month[15].value+'��)�p��0');	
	   return false;
    }	
    if(parseInt(form.amt_6month[15].value) < 0){
	   alert('�ˮֵo�{3�Ӥ�~����6�Ӥ몺�X�p���B('+form.amt_6month[15].value+'��)�p��0');	
	   return false;
    }
    if(parseInt(form.amt_1year[15].value) < 0){
	   alert('�ˮֵo�{6�Ӥ�~����1�~���X�p���B('+form.amt_1year[15].value+'��)�p��0');	
	   return false;
    }
    if(parseInt(form.amt_2year[15].value) < 0){
	   alert('�ˮֵo�{1�~~����2�~���X�p���B('+form.amt_2year[15].value+'��)�p��0');	
	   return false;
    }
    if(parseInt(form.amt_over2year[15].value) < 0){
	   alert('�ˮֵo�{2�~�H�W���X�p���B('+form.amt_over2year[15].value+'��)�p��0');	
	   return false;
    }
    if(parseInt(form.amt_total[15].value) < 0){
	   alert('�ˮֵo�{�O����ƦX�p���X�p���B('+form.amt_total[15].value+'��)�p��0');	
	   return false;
    }
	return true;
}
// 95.05.26 add check 992300�H�γ����u�H�� by 2295
// 95.06.01 add �s�W/�ק��,���n�ˬd992300�H�γ����u�H�� by 2295
//100.06.27 fix A99.992110�A(��)�|����b�ȥ����ӳ�,�B���i��0 by 2295
//100.08.04 fix A99.992100�A(��)�|����겣�`�B�����ӳ�,�B���i��0/�t�� by 2295
//101.06.29 add 992300�H�γ����u�H�ƻݻP�`���c�q�ƫH�η~�ȭ��u�H�Ƥ@�P,�~�i�s�W by 2295
function checkShowInsert_A99(form, fun) {			
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}
	if (fun == 'Insert' || fun == 'Update') {
		//alert('992300='+form.amt[10].value);
		//alert('wlx01.credit_staff_num='+form.wlx01_credit_staff_num.value);
		//alert('���`���c�򥻸�ƺ��@���q�ƫH�η~�ȭ��u�H��='+form.wlx01_credit_staff_num.value);		
        if(form.amt[10].value < 0){		   
           alert('992300�H�γ����u�H�Ƥ��i���t��');	
		   return false;
		}else if(form.amt[10].value == 0){
		   alert('992300�H�γ����u�H�Ƥ��i��0');	
		   return false;		   
		}else if(form.amt[10].value != form.wlx01_credit_staff_num.value){
		   alert('992300�H�γ����u�H��('+form.amt[10].value+')�P���`���c�򥻸�ƺ��@���q�ƫH�η~�ȭ��u�H��('+form.wlx01_credit_staff_num.value+')���@�P!!�Y�ݭץ��`���c�򥻸�ƺ��@���q�ƫH�η~�ȭ��u�H��,�Ц��`���c�򥻸�ƺ��@(FX001W)�i��ץ�!!');		   	  
		   return false;			  
	  }
	    	
	    //100.08.04 fix A99.992100�A(��)�|����겣�`�B�����ӳ�,�B���i��0/�t��
	    if(form.amt[0].value == ''){
	       alert('�A(��)�|����겣�`�B���i���ŭ�');
	       return false;
 	    }else if(form.amt[0].value == 0){
	       alert('�A(��)�|����겣�`�B���i��0');
	       return false;
	    }else if(form.amt[0].value < 0){
	       alert('�A(��)�|����겣�`�B���i���t��');
	       return false;
	    }	
	    //100.06.27 fix A99.992110�A(��)�|����b�ȥ����ӳ�,�B���i��0
	    if(form.amt[1].value == ''){
	       alert('�A(��)�|����b�Ȥ��i���ŭ�');
	       return false;
 	    }else if(form.amt[1].value == 0){
	       alert('�A(��)�|����b�Ȥ��i��0');
	       return false;
	    }		   	    			
	}
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && (!checkNumber(form.elements[i]))){
			return false;
		}	
	}
	return true;
}
//96.07.10 add �s�ڱb���`���(E)�۰ʲ֥[
function addA08(form){	          
     		
	var a = 0;		
	if(form.WarnAccount_Cnt.value != '' && !isNaN(changeVal(form.WarnAccount_Cnt))){//WarnAccount_Cnt-A	
	   a = a + parseInt(eval(changeVal(form.WarnAccount_Cnt)));	   	   
	}
	if(form.LimitAccount_Cnt.value != '' && !isNaN(changeVal(form.LimitAccount_Cnt))){//LimitAccount_Cnt-B	
	   a = a + parseInt(eval(changeVal(form.LimitAccount_Cnt)));	   	   
	}
	if(form.ErrorAccount_Cnt.value != '' && !isNaN(changeVal(form.ErrorAccount_Cnt))){//ErrorAccount_Cnt-C	
	   a = a + parseInt(eval(changeVal(form.ErrorAccount_Cnt)));	   	   
	}
	if(form.OtherAccount_Cnt.value != '' && !isNaN(changeVal(form.OtherAccount_Cnt))){//OtherAccount_Cnt-D	
	   a = a + parseInt(eval(changeVal(form.OtherAccount_Cnt)));	   	   
	}
	form.DepositAccount_TCnt.value = a;
	form.DepositAccount_TCnt.value=changeStr(form.DepositAccount_TCnt);		
}		
//96.07.10 add �s�W/�ק��,���n�ˬdE=A+B+C+D
function checkShowInsert_A08(form, fun) {			
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}	
	var a = 0;		
	var b = 0;
	if(form.WarnAccount_Cnt.value != '' && !isNaN(changeVal(form.WarnAccount_Cnt))){//WarnAccount_Cnt-A	
	   a = a + parseInt(eval(changeVal(form.WarnAccount_Cnt)));	   	   
	}
	if(form.LimitAccount_Cnt.value != '' && !isNaN(changeVal(form.LimitAccount_Cnt))){//LimitAccount_Cnt-B	
	   a = a + parseInt(eval(changeVal(form.LimitAccount_Cnt)));	   	   
	}
	if(form.ErrorAccount_Cnt.value != '' && !isNaN(changeVal(form.ErrorAccount_Cnt))){//ErrorAccount_Cnt-C	
	   a = a + parseInt(eval(changeVal(form.ErrorAccount_Cnt)));	   	   
	}
	if(form.OtherAccount_Cnt.value != '' && !isNaN(changeVal(form.OtherAccount_Cnt))){//OtherAccount_Cnt-D	
	   a = a + parseInt(eval(changeVal(form.OtherAccount_Cnt)));	   	   
	}
	if(form.DepositAccount_TCnt.value != '' && !isNaN(changeVal(form.DepositAccount_TCnt))){//DepositAccount_TCnt-E	
	   b = parseInt(eval(changeVal(form.DepositAccount_TCnt)));	   	   
	}
	if(a != b){
      alert('�s�ڱb���`���(E)�G�Pĵ�ܱb��(A)�B�l�ͺި�b��(B)�B�ۦ�z�靈���`�äv�ĸ���y�X�ި�I���s�ڱb��(C)�B��L�b��(D)����ƥ[�`����');
      return false;
    }	
	return true;
}

//96.12.10 �۰ʭp�������`�B����v(A-B)/(C-D)
//97.07.07 fix �Yc-d���s��,��v��0
function countA09(form){	          
     		
	var a_b = 0;
	var c_d = 0;	
	var tmp = 0;	
	if(form.over_amt.value != '' && !isNaN(changeVal(form.over_amt))){//over_amt-�Ѿl���B(A)
	   a_b = a_b + parseInt(eval(changeVal(form.over_amt)));	   	   
	}
	
	if(form.PUSH_over_amt.value != '' && !isNaN(changeVal(form.PUSH_over_amt))){//PUSH_over_amt-�Ѿl���B-�ʦ���(B)
	   a_b = a_b - parseInt(eval(changeVal(form.PUSH_over_amt)));	   	   
	}
	
	if(form.totalamt.value != '' && !isNaN(changeVal(form.totalamt))){//totalamt-���|��X�`���B(C)	
	   c_d = c_d + parseInt(eval(changeVal(form.totalamt)));	   	   
	}
	if(form.PUSH_totalamt.value != '' && !isNaN(changeVal(form.PUSH_totalamt))){//PUSH_totalamt-���|��X�`���B-�ʦ���(D)	
	   c_d = c_d - parseInt(eval(changeVal(form.PUSH_totalamt)));	   	   
	}
	//����p�Ʋ�2��(�ĤT��|�ˤ��J)
	if(c_d != 0){//97.07.07 fix �Yc-d���s��,��v��0
	  tmp = parseFloat(a_b) / parseFloat(c_d);    
      //alert(tmp);
      tmp = Math.round(tmp * 10000);    
      tmp = tmp / 100;
	}else{
	  tmp = 0;
	}
    form.Over_total_rate.value=tmp    	
}

//96.12.10 
function checkShowInsert_A09(form, fun) {			
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}			
	return true;
}		


//97.06.13 A10.�۰ʭp��X�p
function addA10(form){	          
     		
	var a = 0;		
	if(form.loan2_amt.value != '' && !isNaN(changeVal(form.loan2_amt))){//loan2_amt
	   a = a + parseInt(eval(changeVal(form.loan2_amt)));	   	   
	}
	if(form.loan3_amt.value != '' && !isNaN(changeVal(form.loan3_amt))){//loan2_amt
	   a = a + parseInt(eval(changeVal(form.loan3_amt)));	   	   
	}
	if(form.loan4_amt.value != '' && !isNaN(changeVal(form.loan2_amt))){//loan2_amt
	   a = a + parseInt(eval(changeVal(form.loan4_amt)));	   	   
	}
	form.loan_sum.value = a;
	form.loan_sum.value=changeStr(form.loan_sum);	
	
	a = 0;		
	if(form.invest2_amt.value != '' && !isNaN(changeVal(form.invest2_amt))){//invest2_amt
	   a = a + parseInt(eval(changeVal(form.invest2_amt)));	   	   
	}
	if(form.invest3_amt.value != '' && !isNaN(changeVal(form.invest2_amt))){//invest3_amt
	   a = a + parseInt(eval(changeVal(form.invest3_amt)));	   	   
	}
	if(form.invest4_amt.value != '' && !isNaN(changeVal(form.loan2_amt))){//invest4_amt
	   a = a + parseInt(eval(changeVal(form.invest4_amt)));	   	   
	}
	form.invest_sum.value = a;
	form.invest_sum.value=changeStr(form.invest_sum);		
	
	a = 0;		
	if(form.other2_amt.value != '' && !isNaN(changeVal(form.other2_amt))){//other2_amt
	   a = a + parseInt(eval(changeVal(form.other2_amt)));	   	   
	}
	if(form.other3_amt.value != '' && !isNaN(changeVal(form.other3_amt))){//other2_amt
	   a = a + parseInt(eval(changeVal(form.other3_amt)));	   	   
	}
	if(form.other4_amt.value != '' && !isNaN(changeVal(form.other4_amt))){//other2_amt
	   a = a + parseInt(eval(changeVal(form.other4_amt)));	   	   
	}
	form.other_sum.value = a;
	form.other_sum.value=changeStr(form.other_sum);	
	
	a = 0;		
	if(form.loan2_amt.value != '' && !isNaN(changeVal(form.loan2_amt))){//loan2_amt
	   a = a + parseInt(eval(changeVal(form.loan2_amt)));	   	   
	}
	if(form.invest2_amt.value != '' && !isNaN(changeVal(form.invest2_amt))){//invest2_amt
	   a = a + parseInt(eval(changeVal(form.invest2_amt)));	   	   
	}
	if(form.other2_amt.value != '' && !isNaN(changeVal(form.other2_amt))){//other2_amt
	   a = a + parseInt(eval(changeVal(form.other2_amt)));	   	   
	}
	form.type2_sum.value = a;
	form.type2_sum.value=changeStr(form.type2_sum);	
	
	a = 0;		
	if(form.loan3_amt.value != '' && !isNaN(changeVal(form.loan3_amt))){//loan3_amt
	   a = a + parseInt(eval(changeVal(form.loan3_amt)));	   	   
	}
	if(form.invest3_amt.value != '' && !isNaN(changeVal(form.invest3_amt))){//invest3_amt
	   a = a + parseInt(eval(changeVal(form.invest3_amt)));	   	   
	}
	if(form.other3_amt.value != '' && !isNaN(changeVal(form.other3_amt))){//other3_amt
	   a = a + parseInt(eval(changeVal(form.other3_amt)));	   	   
	}
	form.type3_sum.value = a;
	form.type3_sum.value=changeStr(form.type3_sum);	
	
	a = 0;		
	if(form.loan4_amt.value != '' && !isNaN(changeVal(form.loan4_amt))){//loan4_amt
	   a = a + parseInt(eval(changeVal(form.loan4_amt)));	   	   
	}
	if(form.invest4_amt.value != '' && !isNaN(changeVal(form.invest4_amt))){//invest4_amt
	   a = a + parseInt(eval(changeVal(form.invest4_amt)));	   	   
	}
	if(form.other4_amt.value != '' && !isNaN(changeVal(form.other4_amt))){//other4_amt
	   a = a + parseInt(eval(changeVal(form.other4_amt)));	   	   
	}
	form.type4_sum.value = a;
	form.type4_sum.value=changeStr(form.type4_sum);	
	
	a = 0;		
	if(form.type2_sum.value != '' && !isNaN(changeVal(form.type2_sum))){//type2_sum
	   a = a + parseInt(eval(changeVal(form.type2_sum)));	   	   
	}
	if(form.type3_sum.value != '' && !isNaN(changeVal(form.type3_sum))){//type3_sum
	   a = a + parseInt(eval(changeVal(form.type3_sum)));	   	   
	}
	if(form.type4_sum.value != '' && !isNaN(changeVal(form.type4_sum))){//type4_sum
	   a = a + parseInt(eval(changeVal(form.type4_sum)));	   	   
	}
	form.type_sum.value = a;
	form.type_sum.value=changeStr(form.type_sum);	
}		

//97.06.13 
function checkShowInsert_A10(form, fun) {			
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}			
	return true;
}
/* Add Method by jei 2004.12.14 ver1.0*/
function checkShowInsert_B01(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by jei 2004.12.14 ver1.0*/
function checkShowInsert_B02(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by egg 2004.12.15 */
function checkShowInsert_B03(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by egg 2004.12.10 ver1.0*/
function checkShowInsert_M01(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by jei 2004.12.10 ver1.0*/
function checkShowInsert_M02(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by egg 2004.12.11 */
function checkShowInsert_M03(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by jei 2004.12.11 */
function checkShowInsert_M04(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by Winnin.chu 2004.11.24 ver1.0*/
function checkShowInsert_M05(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

/* Add Method by egg 2004.12.13 */
function checkShowInsert_M06(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}
	
	for (var i = 0; i < form.guarantee_no_month.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_month.length; i++) {
		if (!checkNumber(form.guarantee_amt_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_month.length; i++) {
		if (!checkNumber(form.loan_amt_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_no_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_no_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_no.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_amt.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_p.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_bal.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	
	return true;
}

/* Add Method by egg 2004.12.13 */
function checkShowInsert_M07(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i = 0; i < form.guarantee_no_month.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_month.length; i++) {
		if (!checkNumber(form.guarantee_amt_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_month.length; i++) {
		if (!checkNumber(form.loan_amt_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_no_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_year.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_no_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_amt_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_amt_totacc.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_no.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_amt.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.guarantee_bal_p.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}
	for (var i = 0; i < form.loan_bal.length; i++) {
		if (!checkNumber(form.guarantee_no_month[i]))
			return false;
	}

	return true;
}

/* Add Method by egg 2004.12.16 */
function checkShowInsert_M08(form, fun) {
	//if (!Check_Maintain(form))
	//	return false;
	
	//if (trimString(form.maintain_email.value) == "") {
	//	alert("�ӿ��E_MAIL������J")
	//	form.maintain_email.focus();
	//	return false;
	//}
	
	if (fun == 'Insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	/* 
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkFloatNumber(form.amt[i]))
			return false;
	}
	*/

	return true;
}

function checkShowInsert_G02(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
		if (!CheckYear(form.beg_y))
			return false;
		if (!CheckYear(form.over_y))
			return false;
		if (trimString(form.cust_form.value) == "") {
			alert("�Τ@�s���Ψ����Ҧr��������J")
			return false;
		}
		if (trimString(form.cust_name.value) == "") {
			alert("�j�B�«H�Ȥ�W�٥�����J")
			return false;
		}
		if (trimString(form.beg_y.value) == "") {
			alert("�}�l�ʥΤ�(�~)������J")
			return false;
		}
		if (trimString(form.beg_m.value) == "") {
			alert("�}�l�ʥΤ�(��)������J")
			return false;
		}
		if (trimString(form.beg_d.value) == "") {
			alert("�}�l�ʥΤ�(��)������J")
			return false;
		}
		if (trimString(form.over_y.value) == "") {
			alert("�}�l�n���(�~)������J")
			return false;
		}
		if (trimString(form.over_m.value) == "") {
			alert("�}�l�n���(��)������J")
			return false;
		}
		if (trimString(form.over_d.value) == "") {
			alert("�}�l�n���(��)������J")
			return false;
		}
		var chkDate;
		chkDate =  '' + (parseInt(form.beg_y.value) + 1911) + '/' + form.beg_m.value + '/' + form.beg_d.value;
		if (fnValidDate(chkDate) != true) {
			alert('�ҿ�J���}�l�ʥΤ鬰�L�Ĥ��!!');
    	    form.beg_d.focus();
        	return false;
   		}
		chkDate =  '' + (parseInt(form.over_y.value) + 1911) + '/' + form.over_m.value + '/' + form.over_d.value;
		if (fnValidDate(chkDate) != true) {
			alert('�ҿ�J���}�l�n��鬰�L�Ĥ��!!');
    	    form.over_d.focus();
        	return false;
   		}
	} //End of if

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt')	{
			if (!checkNumber(form.elements[i]))
				return false;
		}
	}
	form.Function.value = fun;
	if (fun == 'update')
		form.submit();

	return true;
}
function checkShowInsert_G05(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
		if (trimString(form.cust_form.value) == "") {
			alert("�Τ@�s���Ψ����Ҧr��������J")
			return false;
		}
		if (trimString(form.cust_name.value) == "") {
			alert("�j�B�«H�Ȥ�W�٥�����J")
			return false;
		}
	} //End of if

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt')	{
			if (!checkNumber(form.elements[i]))
				return false;
		}
	}
	form.Function.value = fun;
	if (fun == 'update')
		form.submit();
	return true;
}
function checkShowInsert_H02(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H03(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
//		if ((form.elements[i].name.substr(7) == 'amt') || (form.elements[i].name.substr(7) == 'amt1')) {
		if ((form.elements[i].name == 'amt') || (form.elements[i].name == 'amt1')) {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
//		if (form.elements[i].name.substr(7) == 'rate') {
		if (form.elements[i].name == 'rate') {
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("�п�J�Ʀr");
		   	    return false;
		    }
			if (!checkRate00(form.elements[i]))
				return false;
		}

	}
	return true;
}

function checkShowInsert_H04(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H05(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H06(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H07(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
		if (form.elements[i].name == 'rate') {
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("�п�J�Ʀr");
		   	    return false;
		    }
			if (!checkRate00(form.elements[i]))
				return false;
		}
	}
	return true;
}
//=========================================================
function checkShowInsert_F01(form, fun) {	
	if (fun == 'new') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && (!checkNumber(form.elements[i]))){
			return false;
		}	
	}

	return true;
}
function checkRate000(Rate1) {

	var rate = changeVal(Rate1)

    if (isNaN(Math.abs(rate))) {
// 	    alert("�п�J�Ʀr");
   	    return;
    }
    if (rate > 100.000 || rate < -99.999) {
        alert("�Q�v���i�j�� 100.000, �]���i�p�� -99.999")
        return;
    }
    if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 3) {
            alert("�p���I��u�঳�T�Ӧ��");
            return;
        }
    }
    return true;
}
//=========================================================
function checkRate00(Rate1) {

	var rate = changeVal(Rate1)

    if (isNaN(Math.abs(rate))) {
// 	    alert("�п�J�Ʀr");
   	    return;
    }
    if (rate > 100.00 || rate < -99.99) {
        alert("�Q�v���i�j�� 100.00, �]���i�p�� -99.99")
        return;
    }
    if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 2) {
            alert("�p���I��u�঳�G�Ӧ��");
            return;
        }
    }
    return true;
}

/* Add Method by Winnin.chu 2004.11.24*/
/* modify by 2354 2004.12.23*/
//94.07.20 add 920710.920720.920730.920740.920750 �]�@�֭p�⭷�I�v��
function changeStr_A05(amt_obj,amt_index,form){
	/* �W�C�H�~�̳W�w,���I�v�ƥ��F100%���겣*/	
	if(amt_obj.value=="") return;	
	if(parseInt(form.acc_code[amt_index].value.substring(0,5))>=92071 &&
	   parseInt(form.acc_code[amt_index].value.substring(0,5))<=92075 &&
	   form.acc_code[amt_index].value.indexOf('N')==-1){// &&
	   //form.acc_code[amt_index].value.indexOf('P')!=-1){	   
	    
		if(isNaN(amt_obj.value)==true){
			alert('�п�J�Ʀr');
			amt_obj.value="";
			amt_obj.focus();
			return ;
		}
		
		if(form.acc_code[amt_index].value.indexOf('P')!=-1){ 
		   form.amt[amt_index-2].value=changeVal(form.amt[amt_index-2]);	
		   //alert(form.amt[amt_index-2].value);	
		   if(isNaN(form.amt[amt_index-2].value)){
			  alert('�п�J�b�����B�Ʀr');
			  form.amt[amt_index-2].focus();
			  return ;
		   }		   
		   tmpassumed=parseInt(form.amt[amt_index-2].value)*parseFloat(amt_obj.value);
		   tmpassumed=tmpassumed+"";
		   
		   /* �Ȩ���p���I���3��*/
		   if(tmpassumed.indexOf('.')!=-1) tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+3);
		   form.assumed_amt[amt_index-2].value=tmpassumed;
		   form.assumed_amt[amt_index-2].value=changeStr(form.assumed_amt[amt_index-2]);
		   document.all["div_assumed_amt"][amt_index-2].innerText=form.assumed_amt[amt_index-2].value;
		   form.amt[amt_index-2].value=changeStr(form.amt[amt_index-2]);
	    }else{//94.07.20 add 920710.920720.920730.920740.920750 
		   form.amt[amt_index+2].value=changeVal(form.amt[amt_index+2]);		
		   //alert(form.amt[amt_index+2].value);
		   if(isNaN(form.amt[amt_index+2].value)){
			  alert('�п�J���I�v��');
			  form.amt[amt_index+2].focus();
			  return ;
		   }		   
		   tmpassumed=parseInt(amt_obj.value)*parseFloat(form.amt[amt_index+2].value);
		   tmpassumed=tmpassumed+"";		   
		   /* �Ȩ���p���I���3��*/
		   if(tmpassumed.indexOf('.')!=-1) tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+3);
		   form.assumed_amt[amt_index].value=tmpassumed;
		   form.assumed_amt[amt_index].value=changeStr(form.assumed_amt[amt_index]);
		   document.all["div_assumed_amt"][amt_index].innerText=form.assumed_amt[amt_index].value;
		   form.amt[amt_index].value=changeStr(form.amt[amt_index]);	 
	    }	
	}else{
	    amt_obj.value=changeStr(amt_obj);			
		if( form.assumed[amt_index] != null && form.assumed[amt_index].value!="" && !isNaN(form.assumed[amt_index].value)){				
			amt_obj.value=changeVal(amt_obj);
			tmpassumed=parseInt(amt_obj.value)*parseFloat(form.assumed[amt_index].value);
			tmpassumed=tmpassumed+"";
			/* �Ȩ���p���I���3��*/
			if(tmpassumed.indexOf('.')!=-1) tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+3);
			form.assumed_amt[amt_index].value=tmpassumed;
			form.assumed_amt[amt_index].value=changeStr(form.assumed_amt[amt_index]);
			document.all["div_assumed_amt"][amt_index].innerText=form.assumed_amt[amt_index].value;
			amt_obj.value=changeStr(amt_obj);				
		}
	}
	
	addText(form);
	/*94.03.10 have ??
	i=0;
	assum=0;
	for(i=0;i<form.assumed_amt.length;i++){
		tmpstr=changeVal(form.assumed_amt[i]);
		if(tmpstr.length!=0 && !isNaN(tmpstr)){
			assum+=parseInt(tmpstr);
		}
	}
	form.amt[42].value=assum;
	form.amt[42].value=changeStr(form.amt[42]);
	*/
}

/* 
Add Method by EGG 2005.02.16
M06,M07 �p�p��Ʀ۰ʭp��B�z
*/
function amtSub(form,changefield)
{
	var outInteger=0;

	for(var i=0;i<form.area_no.length;i++){
		if(form.area_no[i].value != 'A' &&
		   form.area_no[i].value != 'E' &&
		   form.area_no[i].value != '1' &&
		   form.area_no[i].value != '2' &&
		   form.area_no[i].value != '0')
		{
			//���D�ťծɤp�p�֭p�B�z
			if (trimString(changefield[i].value) != "") {
				//alert("changefield[i]="+trimString(changefield[i].value));
				outInteger=outInteger+parseInt(changeVal(changefield[i]));
			}
			//alert("outInteger="+outInteger);
			//outInteger=outInteger+parseInt(changeVal(changefield[i]));
			//alert("1234="+outInteger=outInteger+parseInt(checkNumber(changeVal(changefield[i]))));
			//alert("1234="+(outInteger+parseInt(checkNumber(changeVal(changefield[i]) ) ) ) );
		}else if(form.area_no[i].value=='2'){
			changefield[i].value=outInteger;
			//�ഫ","�B�z
			changefield[i].value=changeStr(changefield[i]);
		}
	}
}

/* 
Add Method by EGG 2005.03.15
M06,M07 �p�p��Ʀ۰ʭp��B�z
*/
function amtSub_1(form,changefield)
{
	var outInteger=0;
	//alert("amtSub_1");
	for(var i=0;i<form.area_no.length;i++){
		if(form.area_no[i].value != 'A' &&
		   form.area_no[i].value != 'E' &&
		   form.area_no[i].value != '1' &&
		   form.area_no[i].value != '2' &&
		   form.area_no[i].value != '0')
		{
			//���D�ťծɤp�p�֭p�B�z
			if (trimString(changefield[i].value) != "") {
				//alert("changefield[i]="+trimString(changefield[i].value));
				outInteger=outInteger+eval((changefield[i].value)*100);
			}
		}else if(form.area_no[i].value=='2'){
			//alert("outInteger="+outInteger);
			changefield[i].value=(outInteger/100);
			//�ഫ","�B�z
			//changefield[i].value=changeStr(changefield[i]);
		}
	}
}


 
//94.11.10 ���멳���.�W�멳�l�B.����s�J.���봣�X�p�p�ʭp��B�z by 2295
function subsum(form,varx5n)
{	
    var x5n=0
    //alert(varx5n.name);
	for(var i=0;i<form.elements.length;i++){								
        if(form.elements[i].name == varx5n.name.substr(0,1) + "1" + varx5n.name.substr(2,1)) {						   					            
           //alert('find it ');
           if (trimString(form.elements[i].value) != "" && !isNaN(changeVal(form.elements[i]))) {				
               x5n = x5n + parseInt(eval(changeVal(form.elements[i]))) ;             
           }
           if (trimString(form.elements[i+5].value) != "" && !isNaN(changeVal(form.elements[i+5]))) {				
               x5n = x5n + parseInt(eval(changeVal(form.elements[i+5]))) ;               
           }
           if (trimString(form.elements[i+10].value) != "" && !isNaN(changeVal(form.elements[i+10]))) {				
               x5n = x5n + parseInt(eval(changeVal(form.elements[i+10]))) ;               
           }
           if (trimString(form.elements[i+15].value) != "" && !isNaN(changeVal(form.elements[i+15]))) {				
               x5n = x5n + parseInt(eval(changeVal(form.elements[i+15]))) ;               
           }
           
           eval("form."+varx5n.name.substr(0,1)+"5"+varx5n.name.substr(2,1)+".value="+x5n);
		   eval("form."+varx5n.name.substr(0,1)+"5"+varx5n.name.substr(2,1)+".value=changeStr(form."+varx5n.name.substr(0,1)+"5"+varx5n.name.substr(2,1)+")");		    	 		       
		   //alert('set it');
		   return;
        }	
	}//end of form length	
}



/* 
94.11.08 add ���멳�l�B�۰ʭp��B�z by 2295             
             ex:A15=A12+A13-A14
				A25=A12+A13-A14
				A35=A32+A33-A34
				A45=A42+A43-A44
				�H������B.C.D.E
*/
function amtXn5(form,Varx5){    
    var Xn5=0;    
    //alert(Varx5.name.substr(0,2));
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && trimString(form.elements[i].value) != "" && !isNaN(changeVal(form.elements[i]))
		   && form.elements[i].name.substr(0,2) == Varx5.name.substr(0,2)) {						   			
			//alert('count'+form.elements[i].name);
			//alert(form.elements[i].name.substr(2,1));
			if(form.elements[i].name.substr(2,1) == '2' && form.elements[i].name.substr(1,1) != '5'){				
				Xn5 = Xn5 + parseInt(eval(changeVal(form.elements[i])));				
			   //alert('add 2: Xn5-->'+Xn5);
			}   
			if(form.elements[i].name.substr(2,1) == '3' && form.elements[i].name.substr(1,1) != '5'){
			   Xn5 = Xn5+ parseInt(eval(changeVal(form.elements[i])));			   
			   //alert('add 3:Xn5-->'+Xn5);
		    }
			if(form.elements[i].name.substr(2,1) == '4' && form.elements[i].name.substr(1,1) != '5'){
			   Xn5 = Xn5 - parseInt(eval(changeVal(form.elements[i])));
			   //alert('minus 4 Xn5-->'+Xn5);
			}		    
		    Varx5.value=Xn5;
		    Varx5.value=changeStr(Varx5);	 		 		    
		}
	}//end of form length
}

/* 
94.11.09 add ���멳�l�B�۰ʭp��B�z by 2295             
             ex:A55=A52+A53-A54
				B55=B52+B53-B54
				C55=C52+C53-C54
				D55=D52+D53-D54
				E55=E52+E53-E54
*/
function amtX55(form,Varx5){    	
    var X55=0;        
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && trimString(form.elements[i].value) != "" && !isNaN(changeVal(form.elements[i]))
		   && form.elements[i].name.substr(1,1) == '5' && form.elements[i].name.substr(2,1) != '1'
		   && form.elements[i].name.substr(0,2) == Varx5.name.substr(0,2)   			
		   ) {						   			
			//alert('count'+form.elements[i].name);
			//alert(form.elements[i].name.substr(2,1));
			if(form.elements[i].name.substr(2,1) == '2'){				
				X55 = X55 + parseInt(eval(changeVal(form.elements[i])));				
			   //alert('add 2: Xn5-->'+Xn5);
			}   
			if(form.elements[i].name.substr(2,1) == '3'){
			   X55 = X55+ parseInt(eval(changeVal(form.elements[i])));			   
			   //alert('add 3:Xn5-->'+Xn5);
		    }
			if(form.elements[i].name.substr(2,1) == '4'){
			   X55 = X55 - parseInt(eval(changeVal(form.elements[i])));
			   //alert('minus 4 Xn5-->'+Xn5);
			}
			Varx5.value=X55;
		    Varx5.value=changeStr(Varx5);	 				    
		    //eval("form."+form.elements[i].name.substr(0,2)+"5.value="+X55);
		    //eval("form."+form.elements[i].name.substr(0,2)+"5.value=changeStr(form."+form.elements[i].name.substr(0,2)+"5)");		    	 		    
		}
	}//end of form length
}

//95.04.12 add A06 970000�X�p-->�۰ʲ֥[ by 2295
//95.05.16 fix �Y��120601/120602/120603/120604���֥[ by 2295	
//95.05.25 fix �Y��960500���֥[ by 2295	  
//95.06.01 fix 960500��� by 2295	
//95.08.10 fix ����960500.�X�p��m16�令15 by 2295
function A06_TOTAL(form,sumname){    	
    var sum=0;        
    var amt=0;
    var idx=0;
    //alert(sumname);    
    eval("form."+sumname+"[15].value=0");
    sumLoop:
	for(var i=0;i<form.elements.length;i++){
		if(form.elements[i].type=='text' && trimString(form.elements[i].value) != "" && !isNaN(changeVal(form.elements[i]))		   
		   && form.elements[i].name == sumname   			
		   ) {		
		   	//�Y��120601/120602/120603/120604���֥[		   	
		    if(sumname == 'amt_3month'){  
		       idx = i-2;
		    }else if(sumname == 'amt_6month'){     	
		       idx = i-3;
		    }else if(sumname == 'amt_1year'){  
		       idx = i-4;	
		    }else if(sumname == 'amt_2year'){  	
		       idx = i-5;	
		    }else if(sumname == 'amt_over2year'){  
		       idx = i-6;	
            }else if(sumname == 'amt_total'){  		    	
               idx = i-7;
		    }	   	
		    //alert(form.elements[idx].value);
		   	if(form.elements[idx].value == '120601' || form.elements[idx].value == '120602' ||
		       form.elements[idx].value == '120603' || form.elements[idx].value == '120604' )
		    {
		       continue;	
		    }
		    	
			amt=parseInt(eval(changeVal(form.elements[i])));
			/*
			if(form.elements[idx].value == '960500'){//95.06.01 960500���
		       sum = sum - parseInt(eval(changeVal(form.elements[i])));						    
		    }else{
			   sum = sum + parseInt(eval(changeVal(form.elements[i])));						    
			} */
		    sum = sum + parseInt(eval(changeVal(form.elements[i])));						    
			//alert('add : sum-->'+amt);
			//alert('sum='+sum);
		}			
	}//end of form length
	eval("form."+sumname+"[15].value="+sum);
    eval("form."+sumname+"[15].value=changeStr(form."+sumname+"[15])");		    	 		    	 	
}

//95.04.12 �p��O���v by 2295
//95.05.03 add 970000���O���v by 2295
//95.05.16 fix �Y�X�{�ƾǸ�e��,��v�H0.000���N by 2295
//95.08.10 fix ����960500.�X�p��m16�令15 by 2295
function A06_per(form,amt_index){    	    	
	if(form.amt_total[amt_index].value == '' || form.A01[amt_index].value == '' || form.A01[amt_index].value == 0) return;
    form.amt_total[amt_index].value=changeVal(form.amt_total[amt_index]);	
    form.A01[amt_index].value=changeVal(form.A01[amt_index]);	
    //alert(form.amt_total[amt_index].value);	
    if(isNaN(form.amt_total[amt_index].value)){
	   alert('�п�J�O��X�p�Ʀr');
	   form.amt_total[amt_index].focus();
	   return ;
    }		       
    //tmpassumed=0.0;
    //alert('amt_index='+amt_index);
    //alert('amt_total='+form.amt_total[amt_index].value);
    //alert('A01='+form.A01[amt_index].value);
    //alert('A01per='+form.A01per[amt_index].value);
    
	tmpassumed=(parseInt(form.amt_total[amt_index].value) / parseInt(form.A01[amt_index].value)) * 100;
	tmpassumed=tmpassumed+"";
	//alert('tmpassumed='+tmpassumed);	   
	/* �Ȩ���p���I���3��*/
	if(tmpassumed.indexOf('e') != -1){//�X�{�ƾǲŸ�
	   //alert(tmpassumed);
	   tmpassumed = "0.000";//�Y�X�{�ƾǸ�e��,��v�H0.000���N
	}else if(tmpassumed.indexOf('.')!=-1){
  	   //alert(tmpassumed);
	   tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+4);
    }	
	
	form.A01per[amt_index].value=tmpassumed;
	form.A01per[amt_index].value=changeStr(form.A01per[amt_index]);
	document.all["div_A01per"][amt_index].innerText=form.A01per[amt_index].value;
	form.amt_total[amt_index].value=changeStr(form.amt_total[amt_index]);
	form.A01[amt_index].value=changeStr(form.A01[amt_index]);
    
	//add 970000���O���v by 2295
	//fix ����960500.�X�p��m16�令15 by 2295
	if(parseInt(form.A01[15].value) != 0){//95.06.01 970000.A01����"0",�~�p��O���v by 2295
		form.amt_total[15].value=changeVal(form.amt_total[15]);	
    	form.A01[15].value=changeVal(form.A01[15]);	
		tmpassumed=(parseInt(form.amt_total[15].value) / parseInt(form.A01[15].value)) * 100;
		tmpassumed=tmpassumed+"";
		//alert('tmpassumed='+tmpassumed);	   
		/* �Ȩ���p���I���3��*/
		if(tmpassumed.indexOf('e') != -1){//�X�{�ƾǲŸ�
	   		//alert(tmpassumed);
	   		tmpassumed = "0.000";//�Y�X�{�ƾǸ�e��,��v�H0.000���N
		}else if(tmpassumed.indexOf('.')!=-1){
	   		//alert(tmpassumed);
	   		tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+4);
    	}	
		form.A01per[15].value=tmpassumed;
		form.A01per[15].value=changeStr(form.A01per[15]);
		document.all["div_A01per"][15].innerText=form.A01per[15].value;
		form.amt_total[15].value=changeStr(form.amt_total[15]);
		form.A01[15].value=changeStr(form.A01[15]);
    }
}
//95.04.17 �p��O���v by 2295
//95.05.16 fix �Y�X�{�ƾǸ�e��,��v�H0.000���N by 2295
//95.06.20 fix �O���v��ܦ�m���� by 2295
function A06_per_onload(form){    	    	
	var A01_idx = -1;
	for(var i=0;i<form.elements.length;i++){								
		if(form.elements[i].type=='text' && form.elements[i].name == 'amt_total') {		
			//alert(i);
			++A01_idx;		
			if(form.elements[i].value == '' || form.elements[i+1].value == '' || form.elements[i+1].value == 0){			   
			   //alert('A01_idx='+A01_idx);	
			   continue;
			}   			
		   	
		   	form.elements[i].value=changeVal(form.elements[i]);	
		   	form.elements[i+1].value=changeVal(form.elements[i+1]);	
		   	
		   	//alert('A01_idx='+A01_idx);
		   	//alert(form.elements[i].value);
		   	//alert(form.elements[i+1].value);
			
			if(parseInt(form.elements[i+1].value) != 0){
				tmpassumed=(parseInt(form.elements[i].value) / parseInt(form.elements[i+1].value)) * 100;
				//alert('tmpassumed='+tmpassumed);	   
	        	tmpassumed=tmpassumed+"";	
	        
		        /* �Ȩ���p���I���3��*/
		        if(tmpassumed.indexOf('e') != -1){//�X�{�ƾǲŸ�
	    	       //alert(tmpassumed);
	        	   tmpassumed = "0.000";//�Y�X�{�ƾǸ�e��,��v�H0.000���N
	        	}else if(tmpassumed.indexOf('.')!=-1){
	           	   //alert(tmpassumed);
	           	   tmpassumed=tmpassumed.substring(0,tmpassumed.indexOf('.')+4);	           
            	}				
				form.elements[i+2].value=tmpassumed;
				//alert(form.elements[i+2].value);
				form.elements[i+2].value=changeStr(form.elements[i+2]);
				//alert(form.elements[i+2].value);
				document.all["div_A01per"][A01_idx].innerText=form.elements[i+2].value;
				form.elements[i].value=changeStr(form.elements[i]);		
				form.elements[i+1].value=changeStr(form.elements[i+1]);				
			}else{
				form.elements[i].value=changeStr(form.elements[i]);		
				form.elements[i+1].value=changeStr(form.elements[i+1]);			
			}	
		}			
	}//end of form length		 		    	 	   
}

//95.05.16 add �O��X�p���B���i�j��A01��ڦX�p
//95.06.01 add 150200/960500/970000���Τ�� by 2295
function A06_checkA01(acc_code,acc_name,form,idx){    	    		
	//alert(acc_code+'='+form.amt_total[idx].value);
	//alert(acc_code+'.A01='+form.A01[idx].value);
	//alert(parseInt(form.amt_total[idx].value));
	//alert(changeVal(form.A01[idx]));
	if(acc_code == '150200' || acc_code == '960500' || acc_code == '970000' ) return;
	if(parseInt(form.amt_total[idx].value) > changeVal(form.A01[idx])){
	   alert(acc_code+acc_name+'�O��X�p���B��'+changeStr(form.amt_total[idx])+'���i�j��A01��ڦX�p'+form.A01[idx].value+'!!');
	   return;
    } 			 		    	 	   
}