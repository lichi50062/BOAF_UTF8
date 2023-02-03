function doSubmit(form,cnd,program_id,rpt_no){	 
	    form.action="/pages/ZZ006W.jsp?act="+cnd+"&test=nothing";		     
	    if(cnd == "Qry" && chkData(form)) form.submit();
	    if(cnd == "Edit" ){
	       form.action="/pages/ZZ006W.jsp?act="+cnd+"&program_id="+program_id+"&rpt_no="+rpt_no+"&test=nothing";	    	
	       form.submit();
	    }
	    if(((cnd == "Update") && AskUpdate(form)) || ((cnd == "Delete")&& AskDelete(form))){
	    	form.action="/pages/ZZ006W.jsp?act="+cnd+"&test=nothing";	    	
	    	form.submit();	    
	    }
}	
function chkData(form){
    /*
	if(form.sn_docno0.value=="" && form.sn_docno2.value==""){
		alert("請輸入報表名稱或報表欄位名稱");
		return false;
	}else{
		return true;
	}
	*/
	return true;
}

