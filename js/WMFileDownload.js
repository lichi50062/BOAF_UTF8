//94.05.25 add�[�Wbank_type���Ѽ�
//94.11.15 add �Y��F01��,�u�}��d�� by 2295
function doSubmit(form, myfun) {
	if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) return;
	
//    form.Function.value = myfun;

	var _id = form.Report_no.value;
	var bank_code = form.Bank_Code_S.value;
	/*
	if (_id != 'F01-1' && _id != 'F01-2' && _id != 'F01-3' && _id != 'F01-4' &&
		_id != 'F02' && _id != 'F03' && _id != 'F04' && _id != 'F07' && _id != 'F20') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}
    */
    
    //modify by 2354 2004.12.23 == begin
	var DLId = form.Report_no.value.substr(0, 1);
	if(myfun=='Download' && (form.Report_no.value=='M08' || DLId=='B')){
		alert('���������ѤU���\��');
		return;
	}
    /*
	if (myfun == 'download') {
		if (DLId == 'A')						//����
			form.action = 'WM004W_DownLoad1';
		else if (DLId == 'B')					//�~��
			form.action = 'WM004W_DownLoad3';
		else if (DLId == 'E')					//�H�X��
			form.action = 'WM004W_DownLoad5';
//		else if (DLId == 'D')					//�A�|
//			form.action = 'WM004W_DownLoad6';
//		else if (DLId == 'E')					//���|
//			form.action = 'WM004W_DownLoad7';
		else if (DLId == 'F')					//����
			form.action = 'WM004W_DownLoadF';
		else if (DLId == 'G')					//����
			form.action = 'WM004W_DownLoad8';
		else if (DLId == 'H')					//�H�U
			form.action = 'WM004W_DownLoad0';
		form.submit();
	}
	else
	if (myfun == 'query') {
		if (bank_code == 'ALL') {
			alert('�d�߮ɽп�ܳ�@���ľ��c�W��!');
			return false;
		}
		if (DLId == 'F') {
			if (
				(_id != 'F05') && (_id != 'F06') && (_id != 'F10') && (_id!= 'F11') &&
				(_id != 'F12') && (_id != 'F13') && (_id != 'F14') && (_id!= 'F15') &&
				(_id != 'F16') && (_id != 'F17')) {
				alert('������ХѰ򥻸�ƺ��@�d��');
				return false;
			}
		}
		form.DL_NAME.value = form.DLId.options[form.DLId.selectedIndex].text;
		if (DLId == 'A')						//����
			form.action = 'WM004W_ShowQuery1';
		else if (DLId == 'B')					//�~��
			form.action = 'WM004W_ShowQuery3';
		else if (DLId == 'E')					//�H�X��
			form.action = 'WM004W_ShowQuery5';
//		else if (DLId == 'D')					//�A�|
//			form.action = 'WM004W_ShowQuery6';
//		else if (DLId == 'E')					//���|
//			form.action = 'WM004W_ShowQuery7';
		else if (DLId == 'F')					//����
			form.action = 'WM004W_ShowQueryF';
		else if (DLId == 'G')					//����
			form.action = 'WM004W_ShowQuery8';
		else if (DLId == 'H')					//�H�U
			form.action = 'WM004W_ShowQuery0';
		form.submit();
	}
	*/
	
	//form.act.value="Upload";	
	
	//form.action="/pages/WMFileDownload.jsp?act="+myfun+"&M_YEAR="+form.S_YEAR.value+"&M_MONTH="+form.S_MONTH+"&test=nothing";
	form.action="/pages/WMFileDownload.jsp?act="+myfun+"&bank_type="+form.bank_type.value+"&test=nothing";
	form.submit();
	
	
}

//94.11.15 add �Y��F01��,�u�}��d�� by 2295
function hiddenDownload(form,b)
{
	if(form.Report_no.value=='F01'){
       if (b == 'hidden'){        
           document.all.Downloadbtn.style.display='none';        
       }else{       
           document.all.Downloadbtn.style.display='block';        
       }    
    }else{
       document.all.Downloadbtn.style.display='block';        	
    }   
}