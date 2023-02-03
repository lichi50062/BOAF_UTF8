<%
// 94/10/17 first design by 4180
// 95/04/04 依增修功能案會議記錄fix by 2495
// 95.07.21 修改取得FX005AW權限 by 2295
//100.01.26 fix cd02區域別.區分100年度/99年度 by 2295
//          fix 日期顯示格式 by 2295
//101.01.10 add 機器品名使用代碼檔(cdshareno.cmuse_div='037')代入 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	List WLX05_ATM_SETUP= (request.getAttribute("WLX05_ATM_SETUP")==null)?null:(List)request.getAttribute("WLX05_ATM_SETUP");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	 																										
	Properties permission = ( session.getAttribute("FX005AW")==null ) ? new Properties() : (Properties)session.getAttribute("FX005AW");
	if(permission == null){
       System.out.println("FX005AW_Edit.permission == null");
    }else{
       System.out.println("FX005AW_Edit.permission.size ="+permission.size());
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

    String seq_no="";
    String site_name="";
    String property_no="";
    String machine_name="";
    String hsien_id="";
    String area_id="";
    String hsien_name="";
    String area_name="";
    String addr="";
    String setup_date="";
    String cancel_type="";
    String cancel_date="";
    String comment_m="";
    String per_machine_name="";
    String CANCEL_DATE="";
    
    String SETUP_DATE_Y="";
    String SETUP_DATE_M="";
    String SETUP_DATE_D="";
    	
    String CANCEL_DATE_Y="";
    String CANCEL_DATE_M="";
    String CANCEL_DATE_D="";
    String loaddata="";
    List date_tmp = null;//100.01.26    
    
	cancel_type =  ( request.getParameter("cancel_type")==null ) ? "" : (String)request.getParameter("cancel_type");								 
	site_name =  ( request.getParameter("site_name")==null ) ? "" : (String)request.getParameter("site_name");						   	 
   	addr =  ( request.getParameter("addr")==null ) ? "" : (String)request.getParameter("addr");						      
    SETUP_DATE_Y =  ( request.getParameter("setup_date_y")==null ) ? "" : (String)request.getParameter("setup_date_y");						      
	SETUP_DATE_M =  ( request.getParameter("setup_date_m")==null ) ? "" : (String)request.getParameter("setup_date_m");						   
	SETUP_DATE_D =  ( request.getParameter("setup_date_d")==null ) ? "" : (String)request.getParameter("setup_date_d");						    
	property_no =  ( request.getParameter("property_no")==null ) ? "" : (String)request.getParameter("property_no");						   
    per_machine_name =  ( request.getParameter("machine_name")==null ) ? "" : (String)request.getParameter("machine_name");						  	    
    seq_no =  ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");						   
    loaddata =  ( request.getParameter("loaddata")==null ) ? "" : (String)request.getParameter("loaddata");						      
    CANCEL_DATE =  ( request.getParameter("CANCEL_DATE")==null ) ? "" : (String)request.getParameter("CANCEL_DATE");						      
	CANCEL_DATE_Y =  ( request.getParameter("CANCEL_DATE_Y")==null ) ? "" : (String)request.getParameter("CANCEL_DATE_Y");						   	
	CANCEL_DATE_M =  ( request.getParameter("CANCEL_DATE_M")==null ) ? "" : (String)request.getParameter("CANCEL_DATE_M");						   
	CANCEL_DATE_D =  ( request.getParameter("CANCEL_DATE_D")==null ) ? "" : (String)request.getParameter("CANCEL_DATE_D");						   
	 	 
    if(WLX05_ATM_SETUP!=null)
    {
      seq_no = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("seq_no")); 
    	site_name = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("site_name")); 
    	property_no = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("property_no")); 
    	machine_name = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("machine_name")); 
    	hsien_id = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("hsien_id")); 
    	area_id = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("area_id")); 
    	hsien_name = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("hsien_name")); 
    	area_name = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("area_name")); 
    	addr = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("addr")); 
    	setup_date = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("setup_date")); 
    	cancel_type = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("cancel_type")); 
    	cancel_date = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("cancel_date")); 
    	comment_m = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("comment_m")); 
    	per_machine_name = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(0)).getValue("machine_name")); 
    	System.out.println("sssssssssss per_machine_name ="+per_machine_name);
    	//per_machine_name="OMRON";
        setup_date = (setup_date.equals("null"))?"-":Utility.getCHTdate(setup_date.substring(0, 10),0);     				
        cancel_date = (cancel_date.equals("null"))?"-":Utility.getCHTdate(cancel_date.substring(0, 10),0);     	
    
        //100.01.26 fix 日期顯示格式   				
    	if(setup_date.length() > 1){
    	   date_tmp = Utility.getStringTokenizerData(setup_date,"/");//100.01.26 add
    	   SETUP_DATE_Y = (String)date_tmp.get(0);
    	   SETUP_DATE_M = (String)date_tmp.get(1);
    	   SETUP_DATE_D = (String)date_tmp.get(2);
    	}
    	   
    	//100.01.26 fix 日期顯示格式   				
    	if(cancel_date.length() > 1){
    	   date_tmp = Utility.getStringTokenizerData(cancel_date,"/");//100.01.26 add
    	   CANCEL_DATE_Y = (String)date_tmp.get(0);
    	   CANCEL_DATE_M = (String)date_tmp.get(1);
    	   CANCEL_DATE_D = (String)date_tmp.get(2);
    	}	 
    
    }//end of WLX05_ATM_SETUP!=null

    System.out.println("act ="+act);
	System.out.println("cancel_type ="+cancel_type);
    System.out.println("site_name ="+site_name);
    System.out.println("addr ="+addr);
	System.out.println("SETUP_DATE_Y ="+SETUP_DATE_Y);
    System.out.println("SETUP_DATE_M ="+SETUP_DATE_M);   
	System.out.println("SETUP_DATE_D ="+SETUP_DATE_D);
	System.out.println("property_no ="+property_no);	
    System.out.println("per_machine_name ="+per_machine_name); 
    System.out.println("seq_no ="+seq_no);
    System.out.println("loaddata ="+loaddata);
    System.out.println("CANCEL_DATE ="+CANCEL_DATE);	
    System.out.println("CANCEL_DATE_Y ="+CANCEL_DATE_Y);
    System.out.println("CANCEL_DATE_M ="+CANCEL_DATE_M);   
	System.out.println("CANCEL_DATE_D ="+CANCEL_DATE_D);

    comment_m=(comment_m.equals("null"))?"":comment_m;
    String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ? "cd01" : "cd01_99" ;	  
    String cd02Table = Integer.parseInt(Utility.getYear())>99 ?"cd02" : "cd02_99" ;     
    String sqlcmd = "Select CD01.HSIEN_ID,CD01.HSIEN_NAME,CD02.AREA_ID,CD02.AREA_NAME From "+cd01Table+" CD01, "+cd02Table+" CD02 "+  
    				 " Where  CD01.HSIEN_ID=CD02.HSIEN_ID Order by CD01.HSIEN_ID, CD02.AREA_ID ";
    				 
    List hsien_id_area_id = DBManager.QueryDB_SQLParam(sqlcmd,null,"");
    String HSIEN_ID_AREA_ID = "";
    HSIEN_ID_AREA_ID=((hsien_id.equals("null"))?"":hsien_id)+"/"+((area_id.equals("null"))?"":area_id);
	
%>

<script language="javascript" src="js/FX005AW.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>各農漁會ATM裝設紀錄</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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

function Enable(form)
{
	if(form.cancel_type.value!=""){
		form.CANCEL_DATE_Y.disabled=false;
		form.CANCEL_DATE_M.disabled=false;
		form.CANCEL_DATE_D.disabled=false;
	}
	else
	{
		form.CANCEL_DATE_Y.value="";
		form.CANCEL_DATE_M.value="";
		form.CANCEL_DATE_D.value="";
		form.CANCEL_DATE_Y.disabled=true;
		form.CANCEL_DATE_M.disabled=true;
		form.CANCEL_DATE_D.disabled=true;
	}
}
//-->
</script>
</head>
<body leftmargin="15" topmargin="0">

<%if(act.equals("modify")){%>
<script language="javascript" type="text/JavaScript">
alert("已有該年月資料，將進行修改動作!");
</script>
<%}%>

<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr>
   		 <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
      <tr>
        <td bgcolor="#FFFFFF" width="600">  
           <tr>
                <td width="600" height="18">
                 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="300" align="center"><b><font size="4">各農漁會ATM裝設紀錄</font></b></td>
                      <td width="150" align="right"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="600" height="50">
                <div align="left">
                <table width="600" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="600" >               
                      </td>
                    </tr>
                    <tr>
                      
                       <div>
                       <jsp:include page="getLoginUser.jsp?width=600" flush="true" />             			
                       </div>			 							
			 		   
                    </tr>
                <tr>
                <td width="600" height="300">
                <form method="post" name="outpushdata">
                <table width="600" border="0" align="center" cellpadding="0" cellspacing="0" height="2">
                    <tr>
                      <td  class="sbody" width="600" height="76">
                        <table height="300"  bordercolor="#3A9D99" border="1" width="600">
                        
                     <%
        						if(!seq_no.equals("")){
        						%>
        						<input type="hidden" name="editseq_no" value="<%=seq_no%>">
        						<%}%>
                        
                            <tr class="sbody" >
                            <td align="left" width="171" bgColor="#d8efee" height="21">金融機構代號</td>
                            <td width="411" bgColor="#e7e7e7" height="21"><%=bank_no%></td>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >裝設地點名稱<font size=2 color=red>*</font>
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">
                            <input maxLength="100" size="30" name="site_name" 
                            value="<%=site_name%>">&nbsp;&nbsp;&nbsp;&nbsp; 
                           <% if(WLX05_ATM_SETUP==null)
								{%>
                            <input type="button" name="load" value="載入已申報資料" onClick="javascript:doSubmit(this.document.forms[0],'load');"> 
                            <%}%> 
                            </td>   
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="60" >裝設地址<font size=2 color=red>*</font>
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="60">
                             <select name='HSIEN_ID_AREA_ID'>
						 	 <%for(int i=0;i<hsien_id_area_id.size();i++){%>
						  	<option value="<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")%>"
						 	 <%if((HSIEN_ID_AREA_ID).equals(((String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id"))+"/"+((String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")))) out.print("selected");%>
						  	><%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_name")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_name")%></option>
						 	 <%}%>
						  	</select>
                            <input maxLength="100" size="47" name="addr" value="<%=addr%>">
                            </td>   
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="30" >裝設日期<font size=2 color=red>*</font>
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="30">
                            <input type='hidden' name="SETUP_DATE" value="">
                            <input type='text' name="SETUP_DATE_Y" value="<%=SETUP_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年 
        						<select  name="SETUP_DATE_M">
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(SETUP_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(SETUP_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>月 
        						<select id="hide1" name="SETUP_DATE_D">
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(SETUP_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(SETUP_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>日</font> 
                            </td>   
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="30" >機器編號<font size=2 color=red>*</font>
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="30">
                            <input maxLength="100" size="20" name="property_no" value="<%if(property_no.equals("null")){out.print("");}
                               else{out.print(property_no);}
                      	    %>">
                            </td>   
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="30" >機器品名<font size=2 color=red>*</font>
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="30">
                             <% List machine_name_list = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='037' order by input_order",null,""); //101.10.13 add%>
							<select name='machine_name'>
						 	 <%for(int i=0;i<machine_name_list.size();i++){%>
                             <option value="<%=(String)((DataObject)machine_name_list.get(i)).getValue("cmuse_name")%>"
                             <%if(machine_name.equals((String)((DataObject)machine_name_list.get(i)).getValue("cmuse_name"))) out.print("selected");%>
                             ><%=(String)((DataObject)machine_name_list.get(i)).getValue("cmuse_name")%></option>                            
	                         <%}%>						
							</select>                           
                            
                            </td>   
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee"  height="45">   
                             備註</td>
                            <td width="411" bgColor="#e7e7e7" height="45">
                            <input maxLength="100" size="47" name="comment_m" value="<%=comment_m%>"></td>
                          	
                          </tr>
                            <tr class="sbody">
                            <td align="left" width="96" bgColor="#d8efee" height="40" rowSpan="2">
                              <p align="center">遷移/裁撤 日期</p> 
                           
                            </td>
                            
                            <td width="411" bgColor="#e7e7e7" height="40">
                             <select name='cancel_type' onChange='Enable(this.document.forms[0])'>
                               <option></option>
                               <option value='1' <% out.print((cancel_type.equals("1"))?"selected":""); %> >遷移</option>
                               <option value='2' <% out.print((cancel_type.equals("2"))?"selected":""); %>>裁撤</option>
                             </select>
                             &nbsp;&nbsp;&nbsp;&nbsp; 
                              <input type='hidden' name="CANCEL_DATE" value="" >
                        	 <input type='text' name="CANCEL_DATE_Y" value="<%=CANCEL_DATE_Y%>" size='3' maxlength='3' <% out.print((cancel_type.equals(""))?"disabled":""); %> onblur='CheckYear(this)' > 
        						<font color='#000000'>年 
        						<select id="hide1" <% out.print((cancel_type.equals(""))?"disabled":""); %> name="CANCEL_DATE_M">
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CANCEL_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CANCEL_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>月 
        						<select id="hide1" <% out.print((cancel_type.equals(""))?"disabled":""); %> name="CANCEL_DATE_D">
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CANCEL_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CANCEL_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>日</font> 
                            </td>        
                            </tr>
                            
                        </table>
                      </td>
                    </tr>
                    <td width="600" height="21"><div align="right"><div align="right">
                    <div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>
                 </td>
             </table>
             </td></tr>
             <tr>
                <td width="660" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody" height="176">
                    <tr>
                      <td colspan="2" width="583" height="41">
                      <div align="center">
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                       <%if(act.equals("new")){
                       if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add
                      %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				        <% } }else{
				        if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify','<%=bank_no%>','<%=seq_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
				         <% }}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
                        <td width="93"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'returnList');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                    </table>
                  </div>
                    </td>
                    </tr>
                    <tr>
                      <td colspan="2" width="583" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明      
                        :</font></font> </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
                      <ul>
                      	  <li class="sbody" ><font color="#FF0000" >按[載入已申報資料]鈕方可進行載入作業&nbsp;</li>
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>      
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>      
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>      
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報的數字欄位仍請填「0」</li> 
                          
                        </ul>
                            </font>
                            </font>
                          </font>
                       </font>
                      </td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</table>
<script language="javascript" type="text/JavaScript">
Enable(this.document.forms[0])
</script>

</body>
