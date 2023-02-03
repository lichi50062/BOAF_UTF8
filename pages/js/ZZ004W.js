//94.04.07 fix 若科別為空值時,顯示空白 by 2295
function changeOption_NotCreate(form,cnd){	
    var myXML,nodeType,nodeValue_bank_type, nodeValue_tbank_no, nodeValue_bank_no, nodeValue_subdep_id;
    myXML = document.all("NotCreateXML").XMLDocument;
    
	nodeType = myXML.getElementsByTagName("muser_id");
	nodeValue_bank_type = myXML.getElementsByTagName("bank_type");
	nodeValue_tbank_no = myXML.getElementsByTagName("tbank_no");
	nodeValue_bank_no = myXML.getElementsByTagName("bank_no");
	nodeValue_subdep_id = myXML.getElementsByTagName("subdep_id");
	
	var oOption;
    var checkAdd = false;	
	
	for(var i=0;i<nodeType.length ;i++)
	{	
		if (nodeType.item(i).firstChild.nodeValue == form.MUSER_ID_NotCreate.value){									
	    	form.BANK_TYPE_NotCreate.value = nodeValue_bank_type.item(i).firstChild.nodeValue;	        
	    	form.TBANK_NO_NotCreate.value = nodeValue_tbank_no.item(i).firstChild.nodeValue;
	    	
	    	if(nodeValue_bank_no.item(i).firstChild.nodeValue != 'null' ){    		    	
	    	   form.BANK_NO_NotCreate.value = nodeValue_bank_no.item(i).firstChild.nodeValue;	    	   
	        }else{
	           form.BANK_NO_NotCreate.value = '';
	        }
	        
	        if(nodeValue_subdep_id.item(i).firstChild.nodeValue != 'null'){    		           
	    	   form.SUBDEP_ID_NotCreate.value = nodeValue_subdep_id.item(i).firstChild.nodeValue;	    	   
	        }else{
	           form.SUBDEP_ID_NotCreate.value = '';
	        }
    	}
    }
}
function changeOption_Create(form,cnd){	
    var myXML,nodeType,nodeValue_bank_type, nodeValue_tbank_no, nodeValue_bank_no, nodeValue_subdep_id;
    myXML = document.all("CreateXML").XMLDocument;
    
	nodeType = myXML.getElementsByTagName("muser_id");
	nodeValue_bank_type = myXML.getElementsByTagName("bank_type");
	nodeValue_tbank_no = myXML.getElementsByTagName("tbank_no");
	nodeValue_bank_no = myXML.getElementsByTagName("bank_no");
	nodeValue_subdep_id = myXML.getElementsByTagName("subdep_id");
	
	var oOption;
    var checkAdd = false;	
	
	for(var i=0;i<nodeType.length ;i++)
	{
		if (nodeType.item(i).firstChild.nodeValue == form.MUSER_ID_Create.value){				
	    	form.BANK_TYPE_Create.value = nodeValue_bank_type.item(i).firstChild.nodeValue;	        
	    	form.TBANK_NO_Create.value = nodeValue_tbank_no.item(i).firstChild.nodeValue;
	    	
	    	if(nodeValue_bank_no.item(i).firstChild.nodeValue != 'null'){    	
	    	   form.BANK_NO_Create.value = nodeValue_bank_no.item(i).firstChild.nodeValue;
	        }else{
	           form.BANK_NO_Create.value = '';
	        }
	        if(nodeValue_subdep_id.item(i).firstChild.nodeValue != 'null'){    	
	    	   form.SUBDEP_ID_Create.value = nodeValue_subdep_id.item(i).firstChild.nodeValue;
	        }else{
	           form.SUBDEP_ID_Create.value = '';
	        }
  			  			
    	}
    }
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




function doSubmit(form,cnd){	     	    	  	     
	    form.action="/pages/ZZ004W.jsp?act="+cnd+"&test=nothing";	    
	    if( cnd == "Insert" && (!checkData_Insert(form))) return;	   	    
	    if( cnd == "Update" && (!checkData_Update(form))) return;	 
	    if( cnd == "Delete" && (!checkData_Delete(form))) return;	   	    
	    if(cnd == "new" || cnd == "del" || cnd == "upd" || cnd == "Qry" || cnd == "DelQry") form.submit();	    	
	    if((cnd == "Insert") && AskInsert(form)){ 
	    	setData_Insert(form); 	    	
	    	form.action="/pages/ZZ004W.jsp?act="+cnd+"&inserUser="+form.MUSER_ID_NotCreate.value+"&test=nothing";	  
	    	form.submit();	    	    
	    }	
	    if((cnd == "Update") && AskUpdate(form)){ 
	    	setData_Insert(form); 	    	
	    	form.action="/pages/ZZ004W.jsp?act="+cnd+"&updateUser="+form.MUSER_ID_Create.value+"&test=nothing";	  
	    	form.submit();	    	    
	    }
	    if((cnd == "Delete") && AskDelete(form)){
	    	setData_Delete(form); 	    	
	    	form.action="/pages/ZZ004W.jsp?act="+cnd+"&deleteUser="+form.MUSER_ID_Create.value+"&test=nothing";	  
	    	form.submit();	    	    
	    }	 
}	

function getData(form,cnd,choose){	     	    	    
	    form.action="/pages/ZZ004W.jsp?act=getData&nowact="+cnd+"&choose_type="+choose+"&test=nothing";
	    form.submit();	    
}

function getUser(form,cnd){	     	    	    
	    form.action="/pages/ZZ004W.jsp?act=getUser&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}
function checkData_Insert(form) 
{
  if (trimString(form.MUSER_ID_NotCreate.value) =="" ){
	  alert("欲新增建檔的使用者帳號不可空白");
	  form.MUSER_ID_NotCreate.focus();
	  return false;
  }
  
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {        
   	if(form.elements[i].type=='checkbox' &&  form.elements[i].name.substr(0,2) == 'P_' && form.elements[i].checked == true) {	      	
        flag = true;
    }    
  }
  if (flag == false) {         
    alert('請至少選擇一筆欲新增的基本功能資料!');
    return false;
  }
  	
  return true;
}
function checkData_Update(form) 
{
  if (trimString(form.MUSER_ID_Create.value) =="" ){
	  alert("欲修改的使用者帳號不可空白");
	  form.MUSER_ID_Create.focus();
	  return false;
  }  
  
  	
  return true;
}

function checkData_Delete(form) 
{
  if (trimString(form.MUSER_ID_Create.value) =="" ){
	  alert("欲刪除建檔的使用者帳號不可空白");
	  form.MUSER_ID_Create.focus();
	  return false;
  }
  
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {        
   	if(form.elements[i].type=='checkbox' &&  form.elements[i].name.substr(0,2) == 'P_' && form.elements[i].checked == true) {	      	
        flag = true;
    }    
    if((form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,15) == 'upload_isModify' && form.elements[i].checked == true) {	
      	flag = true;
    }	    
    if((form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,17) == 'download_isModify' && form.elements[i].checked == true) {	
      	flag = true;
    }	  
    if((form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,13) == 'edit_isModify' && form.elements[i].checked == true) {	
      	flag = true;
    }	  
    if((form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,14) == 'query_isModify' && form.elements[i].checked == true) {	
      	flag = true;
    }    
  }
  if (flag == false) {         
    alert('請至少選擇一筆欲刪除的基本功能/上傳/下載/編輯/查詢資料!');
    return false;
  }
  	
  return true;
}

function selectAll(form,kind) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if((kind == "upload") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,15) == 'upload_isModify') {	
      	form.elements[i].checked = true;
      }	    
      if((kind == "download") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,17) == 'download_isModify') {	
      	form.elements[i].checked = true;
      }	  
      if((kind == "edit") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,13) == 'edit_isModify') {	
      	form.elements[i].checked = true;
      }	  
      if((kind == "query") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,14) == 'query_isModify') {	
      	form.elements[i].checked = true;
      }
      if((kind == "program_id") && (form.elements[i].type=='checkbox') && form.elements[i].disabled == false && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].checked = true;
      }	  
  }
  return;
}

function selectNo(form,kind) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if((kind == "upload") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,15) == 'upload_isModify') {	
      	form.elements[i].checked = false;
      }	    
      if((kind == "download") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,17) == 'download_isModify') {	
      	form.elements[i].checked = false;
      }	  
      if((kind == "edit") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,13) == 'edit_isModify') {	
      	form.elements[i].checked = false;
      }	  
      if((kind == "query") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,14) == 'query_isModify') {	
      	form.elements[i].checked = false;
      }   
      if((kind == "program_id") && (form.elements[i].type=='checkbox') && form.elements[i].disabled == false && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].checked = false;
      }	
  }
  return;
}

function setData_Insert(form) {  
  for ( var i = 0; i < form.elements.length; i++) {      
      if(form.elements[i].type=='checkbox' && form.elements[i].checked ==true && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].value = 'Y';
      }
      if(form.elements[i].type=='checkbox' && form.elements[i].checked ==false && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].value = 'N';
      }	
  }
  return;
}

function setData_Delete(form) {  
  for ( var i = 0; i < form.elements.length; i++) {      
      if(form.elements[i].type=='checkbox' && form.elements[i].checked ==true && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].value = 'N';
      }
      if(form.elements[i].type=='checkbox' && form.elements[i].disabled ==true && form.elements[i].name.substr(0,2) == 'P_') {	
      	form.elements[i].value = 'N';
      }
      if(form.elements[i].type=='checkbox' && form.elements[i].checked ==false && form.elements[i].name.substr(0,2) == 'P_') {	      	
      	form.elements[i].checked = true;
      	form.elements[i].value = 'Y';
      }	
  }
  return;
}