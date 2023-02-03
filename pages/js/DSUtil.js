//=====================================================================
//依照BankType的值,顯示銀行清單
//95.08.30 若為農漁會時,則不可挑選金融機構代號 by 2295
//95.09.06 add 所選的金融機構類別為"ALL"全部時,且農/漁會別 != 農漁會 BankListSrc設成disabled by 2295
//95.11.20 add 報表格式fuction by 2295
//95.11.20 add 欲儲存的格式名稱不可有底線符號 by 2295
//95.12.04 add 檢查查詢年月-起始日期.不可大於結束日期 by 2295
//99.03.30 add 根據查詢年度.縣市別.改變總機構名稱 by 2295
//99.12.21 add BR基本報表用的報表格式doSubmit_RptStyle_BR(report_no,cnd) by 2295
//99.12.21 add BR基本報表使用將所定義的報表欄位顯示在畫面上function fn_loadFieldList_BR by 2295
//99.12.21 add BR基本報表用的清除資料function ResetAllData_BR by 2295
//100.03.26 add 100.03.26 add chInput 由舊制年列印 至 新制年 檢核 by 2479
//103.01.21 add 地方主管機關可看到農漁會別所屬轄區資料 by 2295
//103.04.30 add changeOption_loan()提出共用專案農貸.彈性報表使用 by 2295
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
    //alert(arr.length);
	for (var j =0; j < arr.length; j++){
		arrFd = arr[j].split('+');	
		//alert(arrFd[1]+arrFd[0]);			
		checkAdd=false;		
		for(var i =0;i<form.FieldListDst.length;i++){
			//alert('form.FieldListDst.options[i].value='+form.FieldListDst.options[i].value);		
			//alert('form.FieldListDst.options[i].text='+form.FieldListDst.options[i].text);						
			//alert('arrFd[1]='+arrFd[1]);						
			if(form.FieldListDst.options[i].text == arrFd[1]){		    
			   checkAdd = true;			       
		    }   
	    }
	    if(checkAdd == false){
	       //alert('add item='+arrFd[2]);
	       form.FieldListSrc.options[addCount] = new Option(arrFd[1], arrFd[0]);	 
	       addCount++;
	    }	
	}
}


//=====================================================================
//99.12.21將所定義的報表欄位顯示在畫面上(BR基本報表使用)
//=====================================================================
function fn_loadFieldList_BR(form)
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

function fn_changeBankType(form)
{
	if(form.bank_type.value == 'ALL'){		
	   form.BankListSrc.disabled = true;
	   form.BankListDst.disabled = true;	   
	}else{
	   form.BankListSrc.disabled = false;
	   form.BankListDst.disabled = false;
    }	
} 

//95.09.06 add 所選的金融機構類別為"ALL"全部時,且農/漁會別 != 農漁會 BankListSrc設成disabled	
function fn_changeBankListSrc(form)
{
	var checkALL=false;
	for(var k =0;k<form.BankListDst.length;k++){			
		if(form.BankListDst.options[k].value == 'ALL'){		    
		   checkALL = true;			       
		}   
	}
	//alert(checkALL);
	//alert(form.bank_type.value);
	//所選的金融機構類別為"ALL"全部時,且農/漁會別 != 農漁會	
	if(checkALL == true && form.bank_type.value != 'ALL'){		    
	   form.BankListSrc.disabled = true;	      
	}else{
	   form.BankListSrc.disabled = false;	      
	}	    					
} 

// 95.09.04 add 農漁會.預設全部放在已選欄位 by 2295
//          add 農/漁會.全部 by 2295
// 99.03.30 add 根據查詢年度.縣市別.改變總機構名稱 by 2295
//103.01.21 add 地方主管機關可看到農漁會別所屬轄區資料 by 2295
//103.04.30 add 專案農貸.彈性報表使用 by 2295
function changeOption(form,cnd){		
	var agri_loan = document.getElementsByName("agri_loan");
	//alert(agri_loan.length)
    if(agri_loan.length > 0){  	
	   changeOption_loan(form,cnd);//103.04.30 add 專案農貸.彈性報表使用
	}else{   
       var myXML,nodeType,nodeType1,nodeValue, nodeName,bankType,nodeYear;
       var m_year = form.S_YEAR.value;
       //alert(m_year);    
       if(m_year >= 100){
          m_year = 100;
       }else{
          m_year = 99;
       }	
       myXML = document.all("TBankXML").XMLDocument;
       form.BankListSrc.length = 0;
       if(cnd == 'change') form.BankListDst.length = 0;
       BnType = myXML.getElementsByTagName("BnType");//bn_type    
       bankType = myXML.getElementsByTagName("BankType");//bank_type    
	   nodeType = myXML.getElementsByTagName("HsienId");//hsien_id
	   nodeValue = myXML.getElementsByTagName("bankValue");//bank_no
	   nodeName = myXML.getElementsByTagName("bankName");//bank_no+bank_name
	   nodeYear = myXML.getElementsByTagName("BankYear");//m_year
	   var oOption;
       var checkAdd = false;	
       fn_changeBankType(form);	
       //alert('now.m_year='+m_year); 
       //alert('HSIEN_ID='+form.HSIEN_ID.value);   
       //alert('form.bank_type='+form.bank_type.value);
	   for(var i=0;i<nodeType.length ;i++)
	   {
	   	//alert('bankType='+bankType.item(i).firstChild.nodeValue);		
	   	//alert('nodeYear='+nodeYear.item(i).firstChild.nodeValue);
	   	//103.01.21 地方主管機關可看到農漁會別所屬轄區資料 
	   	if(form.bank_type.value !='B' && bankType.item(i).firstChild.nodeValue != form.bank_type.value) continue;//農漁會別
	   	if(nodeYear.item(i).firstChild.nodeValue != 'ALL'){
	   		if(nodeYear.item(i).firstChild.nodeValue != m_year){				
	   		 continue;//顯示查詢年度的新機構名稱
	   		} 
	   	} 
	   	
	   	if(form.HSIEN_ID.value == 'ALL'){		
	   		oOption = document.createElement("OPTION");
	   		if(form.CANCEL_NO.value == 'N'){//營運中				
	   		   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  			   	
	   		      oOption.text=nodeName.item(i).firstChild.nodeValue;
  	   		      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	    			     
	   		   }		
	   	    }else{//已裁撤		    
	   	       if(BnType.item(i).firstChild.nodeValue == '2'){		       
	   		      oOption.text=nodeName.item(i).firstChild.nodeValue;
  	   		      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
	   		   }
	   	    }				  		
  	   		checkAdd=false;
	   		for(var k =0;k<form.BankListDst.length;k++){			
	   			if(form.BankListDst.options[k].text == oOption.text){		    
	   		   	   checkAdd = true;			       
	   	    	}   
	       	}
	       	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	  
	       		//alert('add '+oOption.text);
	       		//alert('add '+oOption.value);	    		
	       		if(form.bank_type.value == 'ALL'){//農漁會.預設全部放在已選欄位
  	   			   form.BankListDst.add(oOption); 
  	   			}else{
  	   			   form.BankListSrc.add(oOption); 
  	   			}
  	   		}	
	       }else if (nodeType.item(i).firstChild.nodeValue == form.HSIEN_ID.value){
  	   		oOption = document.createElement("OPTION");
  	   		if(form.CANCEL_NO.value == 'N'){//營運中  				
	   		   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  
	   		      oOption.text=nodeName.item(i).firstChild.nodeValue;
  	   		      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
	   		   }		
	   	    }else{//已裁撤		    
	   	       if(BnType.item(i).firstChild.nodeValue == '2'){
	   		      oOption.text=nodeName.item(i).firstChild.nodeValue;
  	   		      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
	   		   }
	   	    }			      		
	   		checkAdd=false;
	   		for(var k =0;k<form.BankListDst.length;k++){			
	   			if(form.BankListDst.options[k].text == oOption.text){		    
	   		   	   checkAdd = true;			       
	   	    	}   
	       	}
	       	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	       
  	   			form.BankListSrc.add(oOption); 
  	   		}  			
       	}       	  		
       }
    }//end of 非專案農貸
}
//103.04.30 專案農貸彈性報表使用(DS056W~DS065W) by 2968
function changeOption_loan(form,cnd){	
    var myXML,nodeType,nodeType1,nodeValue, nodeName,bankType,nodeYear;
    var m_year = form.S_YEAR.value;
    if(m_year >= 100){
       m_year = 100;
    }else{
       m_year = 99;
    }	
    myXML = document.all("TBankXML").XMLDocument;
    form.BankListSrc.length = 0;
    if(cnd == 'change') form.BankListDst.length = 0;
    BnType = myXML.getElementsByTagName("BnType");//bn_type    
    bankType = myXML.getElementsByTagName("BankType");//bank_type    
	nodeType = myXML.getElementsByTagName("HsienId");//hsien_id
	nodeValue = myXML.getElementsByTagName("bankValue");//bank_no
	nodeName = myXML.getElementsByTagName("bankName");//bank_no+bank_name
	nodeYear = myXML.getElementsByTagName("BankYear");//m_year
	var oOption;
    var checkAdd = false;	
	for(var i=0;i<nodeType.length ;i++){
		if( form.bank_type.value!='ALL' && (bankType.item(i).firstChild.nodeValue != form.bank_type.value))continue;//農漁會別
		if(nodeYear.item(i).firstChild.nodeValue != m_year)	continue;//顯示查詢年度的新機構名稱
		if(form.HSIEN_ID.value == 'ALL'){	
			oOption = document.createElement("OPTION");
			if(form.CANCEL_NO.value == 'N'){//營運中				
			   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  			   	
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	    			     
			   }		
		    }else{//已裁撤		    
		       if(BnType.item(i).firstChild.nodeValue == '2'){		       
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }
		    }				  		
  			checkAdd=false;
			for(var k =0;k<form.BankListDst.length;k++){			
				if(form.BankListDst.options[k].text == oOption.text){		    
			   	   checkAdd = true;			       
		    	}   
	    	}
	    	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	  
	    		/*if(form.bank_type.value == 'ALL'){//農漁會.預設全部放在已選欄位
  				   form.BankListDst.add(oOption); 
  				}else{*/
  				   form.BankListSrc.add(oOption); 
  				//}
  			}	
	    }else if (nodeType.item(i).firstChild.nodeValue == form.HSIEN_ID.value){
  			oOption = document.createElement("OPTION");
  			if(form.CANCEL_NO.value == 'N'){//營運中  				
			   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }		
		    }else{//已裁撤		    
		       if(BnType.item(i).firstChild.nodeValue == '2'){
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }
		    }			      		
			checkAdd=false;
			for(var k =0;k<form.BankListDst.length;k++){			
				if(form.BankListDst.options[k].text == oOption.text){		    
			   	   checkAdd = true;			       
		    	}   
	    	}
	    	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	       
  				form.BankListSrc.add(oOption); 
  			}  			
    	}       	  		
    }
}

//99.03.30 add 根據查詢年月.改變縣市別名稱
function changeCity(xml, target, source, form) {
      var myXML,nodeType,nodeValue, nodeName,nodeYear,m_year;      
      m_year = source.value;
      if(m_year >= 100){
         m_year = 100;
      }else{
         m_year = 99;
      }	
      
      target.length = 0;      
      var oOption;
     
      myXML = document.all(xml).XMLDocument;
      nodeType = myXML.getElementsByTagName("cityType");//hsien_id
      nodeYear = myXML.getElementsByTagName("cityYear");//m_year
	  nodeValue = myXML.getElementsByTagName("cityValue");//hsien_id
	  nodeName = myXML.getElementsByTagName("cityName");//hsien_name
		
	  oOption = document.createElement("OPTION");
	  oOption.text='全部';
  	  oOption.value='ALL';
  	  target.add(oOption);
  	  
	  for(var i=0;i<nodeType.length ;i++)	{	  	
  	     if (nodeYear.item(i).firstChild.nodeValue == m_year)  {
  		    oOption = document.createElement("OPTION");
		    oOption.text=nodeName.item(i).firstChild.nodeValue;
  		    oOption.value=nodeValue.item(i).firstChild.nodeValue;
  		    target.add(oOption);
   	     }
      }
      form.HSIEN_ID[0].selected=true;      
      changeOption(this.document.forms[0],'');
}

//103.04.30 add 根據查詢年月.改變縣市別名稱(專案農貸彈性報使用)
function changeCity_loan(xml, target, source, form) {
      var myXML,nodeType,nodeValue, nodeName,nodeYear,m_year;      
      m_year = source.value;
      if(m_year >= 100){
         m_year = 100;
      }else{
         m_year = 99;
      }	
      
      target.length = 0;      
      var oOption;
     
      myXML = document.all(xml).XMLDocument;
      nodeType = myXML.getElementsByTagName("cityType");//hsien_id
      nodeYear = myXML.getElementsByTagName("cityYear");//m_year
	  nodeValue = myXML.getElementsByTagName("cityValue");//hsien_id
	  nodeName = myXML.getElementsByTagName("cityName");//hsien_name
		
	  oOption = document.createElement("OPTION");
	  oOption.text='全部';
  	  oOption.value='ALL';
  	  target.add(oOption);
  	  
	  for(var i=0;i<nodeType.length ;i++)	{	  	
  	     if (nodeYear.item(i).firstChild.nodeValue == m_year)  {
  		    oOption = document.createElement("OPTION");
		    oOption.text=nodeName.item(i).firstChild.nodeValue;
  		    oOption.value=nodeValue.item(i).firstChild.nodeValue;
  		    target.add(oOption);
   	     }
      }
      form.HSIEN_ID[0].selected=true;      
      changeOption_loan(this.document.forms[0],'');
}

function setSelect(S1, bankid) {
    if(S1 == null)
    	return;
    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==bankid)    	{
        	S1.options[i].selected=true;
        	break;
    	}
    }
}
//95.11.20 add 報表格式fuction
//95.12.04 add 起始日期不可大於結束日期
function doSubmit_RptStyle(report_no,cnd){//for 報表格式用
   if(cnd == 'createRpt'){//產生報表
      if(this.document.forms[0].BankList.value == ''){        
         alert('金融機構代碼必須選擇');
         return;
      }
      if(!chInput(this.document.forms[0])) return;//95.12.04 add 起始日期不可大於結束日期
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }else if(cnd == 'SaveRpt'){//儲存格式檔   
      if(this.document.forms[0].template.value == ''){
         alert('欲儲存格式之名稱不可為空白!!');
         return;
      }
      //95.11.20 add 欲儲存的格式名稱不可有底線符號 by 2295
      if(this.document.forms[0].template.value.indexOf('_') != -1 ){
         alert('欲儲存格式之名稱不可為有底線[ _ ]符號!!');
         return;
      }
      if(!confirm('是否確定儲存本格式名稱 ??')){         
         return;
      }  
   }else if(cnd == 'DeleteRpt' || cnd == 'ReadRpt'){//刪除/讀取格式檔   
      //alert(this.document.forms[0].template_list.length);
      var flag = false;  
      for(var i = 0 ; i < this.document.forms[0].template_list.length; i++) {    
          if(this.document.forms[0].template_list[i].selected == true ) {
             flag = true;
          }    
      }
      if (flag == false) {     
         alert('請至少選擇一筆範本名稱!!');
         return;
      }
      if(cnd == 'DeleteRpt'){
         if(!confirm('是否確定刪除本格式名稱 ??')){
            return;
         }
      }
      if(cnd == 'ReadRpt'){   
	     if(!confirm('是否確定讀取本格式名稱 ??')){  
           return;
        }	   
      }         
   }
   fn_ShowPanel(report_no,cnd);      
}


//99.12.21 add BR基本報表用的報表格式
function doSubmit_RptStyle_BR(report_no,cnd){//for 報表格式用
   if(cnd == 'createRpt'){//產生報表
      if(this.document.forms[0].BankList.value == ''){        
         alert('金融機構代碼必須選擇');
         return;
      }
      
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }       
   }
   fn_ShowPanel(report_no,cnd);      
}
//95.12.04 add 檢查查詢年月-起始日期.不可大於結束日期
//100.03.26 add 由舊制年列印 至 新制年 檢核
function chInput(form,cnd){
	//alert(form.S_YEAR.value+form.S_MONTH.value);
	//alert(form.E_YEAR.value+form.E_MONTH.value);
    if(cnd == 'BankList'){//金融機構類別Panel時,才需檢查此項   
	   if(!checkDateS(form.S_YEAR,form.S_MONTH,'查詢年月 -起始日期','1')) return false;
	   if(!checkDateS(form.E_YEAR,form.E_MONTH,'查詢年月 -結束日期','1')) return false;
	
	   if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期")) return false;
       if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期")) return false;    
    }
    //新舊無法同時列印
    if(parseInt(form.S_YEAR.value)<=99 && parseInt(form.E_YEAR.value)>=100){
        alert('縣市於100年起改用新制，新舊制無法同時輸出，請重新輸入結束日期');
        return false;
    }
    
	if(form.S_YEAR.value+form.S_MONTH.value > form.E_YEAR.value+form.E_MONTH.value ){
		alert('起始日期不可大於結束日期');
		return false;
	}
    if(cnd == 'BankList'){//金融機構類別Panel時,才需檢查此項
	   if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!=""){
		  if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value)){
    		 alert("起始查詢年月不可大於結束查詢年月");
    		 return false;
    	  }    	   
       }    
    }
	return true;
}

function ResetAllData(cnd){
    if(confirm("確定要清除已選定的資料嗎？")){  
    	if(cnd == 'BankList'){//金融機構	
           this.document.forms[0].BankListDst.length = 0;
           this.document.forms[0].HSIEN_ID[0].selected=true;	   
           changeOption(this.document.forms[0],'');
        }else if (cnd == 'RptColumn'){//報表欄位
           this.document.forms[0].FieldListDst.length = 0;        
           fn_loadFieldList(this.document.forms[0]);//顯示所有的報表欄位名稱	
    	}else if (cnd == 'RptOrder'){//排序欄位
    	   this.document.forms[0].SortListDst.length = 0;
           getbtnFieldList();   
        } 	
        //clearBankList();95.12.07
	}
	return;	
}     
//99.12.21 add BR基本報表用的清除資料
function ResetAllData_BR(cnd){
    if(confirm("確定要清除已選定的資料嗎？")){  
    	if(cnd == 'BankList'){//金融機構	    	
           this.document.forms[0].BankListDst.length = 0;
           this.document.forms[0].HSIEN_ID[0].selected=true;	   
           changeOption(this.document.forms[0],'');
        }else if (cnd == 'RptColumn'){//報表欄位
           this.document.forms[0].FieldListDst.length = 0;        
           fn_loadFieldList_BR(this.document.forms[0]);//顯示所有的報表欄位名稱	
    	}else if (cnd == 'RptOrder'){//排序欄位
    	   this.document.forms[0].SortListDst.length = 0;
           getbtnFieldList();   
        } 	
        //clearBankList();95.12.07
	}
	return;	
}

function ResetAllData_AgriBank(cnd){
    if(confirm("確定要清除已選定的資料嗎？")){  
    	if (cnd == 'RptColumn'){//報表欄位    	   
           this.document.forms[0].FieldListDst.length = 0;
           fn_loadFieldList(this.document.forms[0]);//顯示所有的報表欄位名稱
    	}        
	}
	return;	
}

function fn_ShowPanel(report_no,cnd){   
	//act=BankList/RptColumn/RptOrder/RptType	
	this.document.forms[0].action = "/pages/"+report_no+".jsp?act="+cnd;	
    this.document.forms[0].target = '_self';
	this.document.forms[0].submit();   	  
}

<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->