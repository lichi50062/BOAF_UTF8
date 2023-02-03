<%
//103.01.16 created by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	String nowDay = Utility.getDateFormat("yyyy/MM/dd");//當日日期
	String haveCancel = ( request.getParameter("haveCancel")==null ) ? "N" : (String)request.getParameter("haveCancel");		
   	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
   	List WLX01 = (List)request.getAttribute("WLX01");
   	List WLX01_Audit = (List)request.getAttribute("WLX01_Audit");   	
   	
	//取得FX001W的權限====================================================================================================
	Properties permission = ( session.getAttribute("FX001W")==null ) ? new Properties() : (Properties)session.getAttribute("FX001W"); 
	if(permission == null){
       System.out.println("FX001W_EditAudit.permission == null");
    }else{
       System.out.println("FX001W_EditAudit.permission.size ="+permission.size());               
    }   	
   	//=======================================================================================================================	
	String SETUP_DATE_Y=""; 
	String SETUP_DATE_M="";
	String SETUP_DATE_D="";
	String SETUP_DATE="";
	String PART_TIME_DATE_Y="";
	String PART_TIME_DATE_M="";
	String PART_TIME_DATE_D="";
	String PART_TIME_DATE="";
	String AUDIT_HSIEN_ID_AREA_ID = "";
	String FULL_TIME = "";
	String PART_TIME = "";
	String HSIEN_ID_AREA_ID = "";
	String ADDR = "";
	List tmpDate=null;
	
	String cd01Table = Integer.parseInt(Utility.getYear())> 99 ?"cd01" : "cd01_99" ;
	String cd02Table = Integer.parseInt(Utility.getYear())> 99 ?"cd02" : "cd02_99" ;
	String sqlcmd = "Select CD01.HSIEN_ID,CD01.HSIEN_NAME,CD02.AREA_ID,CD02.AREA_NAME From "+cd01Table+" CD01, "+cd02Table+" CD02 "+
			 " Where  CD01.HSIEN_ID=CD02.HSIEN_ID Order by CD01.HSIEN_ID, CD02.AREA_ID ";
	List hsien_id_area_id = DBManager.QueryDB_SQLParam(sqlcmd,null,"");
	if(WLX01 != null && WLX01.size() != 0){
	    HSIEN_ID_AREA_ID =((((DataObject)WLX01.get(0)).getValue("hsien_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("hsien_id"))	   
				  +"/"+ ((((DataObject)WLX01.get(0)).getValue("area_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("area_id"));	   
		ADDR = ((((DataObject)WLX01.get(0)).getValue("addr")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("addr"));
		System.out.println("WLX01.size()="+WLX01.size());	
	}else{
		System.out.println("WLX01 == null");
	}
	if(WLX01_Audit != null && WLX01_Audit.size() != 0){
	    int i = 0;
		if(((DataObject)WLX01_Audit.get(0)).getValue("setup_date") != null){
		   SETUP_DATE = Utility.getCHTdate((((DataObject)WLX01_Audit.get(0)).getValue("setup_date")).toString().substring(0, 10), 0);		 
		   //95.08.21 顯示日期用getStringTokenizerData來拆日期字串
		   tmpDate = Utility.getStringTokenizerData(SETUP_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		       SETUP_DATE_Y = (String)tmpDate.get(0);
		       SETUP_DATE_M = (String)tmpDate.get(1);
		       SETUP_DATE_D = (String)tmpDate.get(2);		      
		   }
		}		
		if(((DataObject)WLX01_Audit.get(0)).getValue("part_time_date") != null){
		   PART_TIME_DATE = Utility.getCHTdate((((DataObject)WLX01_Audit.get(0)).getValue("part_time_date")).toString().substring(0, 10), 0);
		   tmpDate = Utility.getStringTokenizerData(PART_TIME_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		       PART_TIME_DATE_Y = (String)tmpDate.get(0);
		       PART_TIME_DATE_M = (String)tmpDate.get(1);
		       PART_TIME_DATE_D = (String)tmpDate.get(2);		      
		   }
		}
		AUDIT_HSIEN_ID_AREA_ID =((((DataObject)WLX01_Audit.get(0)).getValue("hsien_id") == null) ?"":(String)((DataObject)WLX01_Audit.get(0)).getValue("hsien_id"))	   
				  +"/"+ ((((DataObject)WLX01_Audit.get(0)).getValue("area_id") == null) ?"":(String)((DataObject)WLX01_Audit.get(0)).getValue("area_id"));	   
		FULL_TIME = ((((DataObject)WLX01_Audit.get(0)).getValue("full_time") == null) ?"":(String)((DataObject)WLX01_Audit.get(0)).getValue("full_time"));
		PART_TIME = ((((DataObject)WLX01_Audit.get(0)).getValue("part_time") == null) ?"":(String)((DataObject)WLX01_Audit.get(0)).getValue("part_time"));
		System.out.println("WLX01_Audit.size="+WLX01_Audit.size());
	}else{
	   	System.out.println("WLX01_Audit == null");
	}
	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX001W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>總機構稽核人員基本資料維護</title>
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
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">  
<input type="hidden" name="seq_no" value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0){out.print((((DataObject)WLX01_Audit.get(0)).getValue("seq_no")).toString());}%>">
<input type="hidden" name="nowDay" value="<%=nowDay%>">
<input type="hidden" name="tmpStr1" value="<%=HSIEN_ID_AREA_ID%>"> 
<input type="hidden" name="tmpStr2" value="<%=ADDR%>"> 
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		  <tr> 
   		   <td><img src="images/space_1.gif" width="12" height="12"></td>
  		  </tr>
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="250"><font color='#000000' size=4><b> 
                        <center>總機構稽核人員基本資料維護 </center>
                        </b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                          						
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>稽核人員姓名</td>
						<td width='70%' colspan=3 bgcolor='e7e7e7'>
						    <input type='text' name='NAME' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( (((DataObject)WLX01_Audit.get(0)).getValue("name") == null )?"":  (String)((DataObject)WLX01_Audit.get(0)).getValue("name"));%>" maxlength='20' >
						    <font color='red' size=4>*</font>    
						</td>
        			    </tr>
					
        			    <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>隸屬部門</td>						                       
						<td width='70%' colspan=3 bgcolor='e7e7e7'>
                            <input type='text' name='DEPARTMENT' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( (((DataObject)WLX01_Audit.get(0)).getValue("department") == null )?"":  (String)((DataObject)WLX01_Audit.get(0)).getValue("department"));%>" size='55' maxlength='80' >
                        </td>                      
                        </tr>       
                        <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>主管機關核准日期</td>
						  <td width='70%' colspan=3 bgcolor='e7e7e7'>
						    <input type='hidden' name='SETUP_DATE' value="">
                            <input type='text' name='SETUP_DATE_Y' value="<%=SETUP_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=SETUP_DATE_M>
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
        						<select id="hide1" name=SETUP_DATE_D>
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
						  <td width='30%' align='left' bgcolor='#D8EFEE'>主管機關核准文號</td>
						  <td width='70%' colspan=3 bgcolor='e7e7e7'>
                            <input type='text' name='SETUP_NO' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( ((DataObject)WLX01_Audit.get(0)).getValue("setup_no") == null ?"":(((DataObject)WLX01_Audit.get(0)).getValue("setup_no")).toString());%>" size='55' maxlength=120>
        					(含字號，另號碼以阿拉伯數字列載)
                          </td>
	  					</tr>
	  					<tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>專任與否</td>						                       
						<td width='70%' colspan=3 bgcolor='e7e7e7'>
                            <select name='FULL_TIME'> 
                            <option value="" >請選擇</option>
                            <option value="Y" <%if("Y".equals(FULL_TIME)) out.print("selected");%>>是</option>
                            <option value="N" <%if("N".equals(FULL_TIME)) out.print("selected");%>>否</option>    
                            </select>
                        </td>  
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>主管機關核准兼任與否</td>						                       
						<td width='70%' colspan=3 bgcolor='e7e7e7'>
                            <select name='PART_TIME'> 
                            <option value="" >請選擇</option>
                            <option value="Y" <%if("Y".equals(PART_TIME)) out.print("selected");%>>是</option>
                            <option value="N" <%if("N".equals(PART_TIME)) out.print("selected");%>>否</option>    
                            </select>
                        </td> 
                        <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>主管機關核准兼任日期</td>
						  <td width='70%' colspan=3 bgcolor='e7e7e7'>
						    <input type='hidden' name='PART_TIME_DATE' value="">
                            <input type='text' name='PART_TIME_DATE_Y' value="<%=PART_TIME_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=PART_TIME_DATE_M>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(PART_TIME_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(PART_TIME_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=PART_TIME_DATE_D>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(PART_TIME_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(PART_TIME_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日</font>
                            </td>
	  					</tr>
	  					<tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>主管機關核准兼任文號</td>
						  <td width='70%' colspan=3 bgcolor='e7e7e7'>
                            <input type='text' name='PART_TIME_NO' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( ((DataObject)WLX01_Audit.get(0)).getValue("part_time_no") == null ?"":(((DataObject)WLX01_Audit.get(0)).getValue("part_time_no")).toString());%>" size='55'  maxlength=120>
        					(含字號，另號碼以阿拉伯數字列載)
                          </td>
	  					</tr>
	  					<tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>稽核人員電話</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='TELNO' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( (((DataObject)WLX01_Audit.get(0)).getValue("telno")==null)?"": (String)((DataObject)WLX01_Audit.get(0)).getValue("telno"));%>" size='20' maxlength='20' >
        				 (含區域號碼，並以"-"區隔) 
        			    </td>
        			    </tr>
        			    <tr class="sbody">
							<td width='30%' bgcolor='#D8EFEE' align='left'>稽核單位地址</td>
							<td width='70%' colspan=2 bgcolor='e7e7e7'>						
								<select name='HSIEN_ID_AREA_ID'>
								<%for(int i=0;i<hsien_id_area_id.size();i++){%>
									<option value="<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")%>"
									<%if((AUDIT_HSIEN_ID_AREA_ID).equals(((String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id"))+"/"+((String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")))) out.print("selected");%>
									><%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_name")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_name")%></option>
									<%}%>						
								</select>
			       				<input type='button' name='ToSameAddr' value="同總機構地址" onClick="javascript:setAddr(form,'AUDIT_ADDR');">                            
								<br><span id="spanA" style="display:"><%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( (((DataObject)WLX01_Audit.get(0)).getValue("area_id")==null)?"": (String)((DataObject)WLX01_Audit.get(0)).getValue("area_id"));%></span>
									<span id="spanB" style="display:none;"><%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("area_id")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("area_id"));%></span>
								<input type='text' name='ADDR' value="<%if(WLX01_Audit != null && WLX01_Audit.size() != 0) out.print( (((DataObject)WLX01_Audit.get(0)).getValue("addr")==null)?"": (String)((DataObject)WLX01_Audit.get(0)).getValue("addr"));%>" size='57' maxlength='80' >                            
	        			    </td>
        			    </tr>
                        </Table></td>
                    </tr>                 
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>   
                        <%if(haveCancel.equals("N")){//94.04.06總機構未裁撤時才能執行修改.刪除%>
                        
   						<%if(act.equals("newAudit")){%>
   						   <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'InsertAudit','Audit');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				           <%}%>
				      	<%}else if(act.equals("EditAudit")){%>  
				      	   <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>                   	        	                                   		     				        
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'UpdateAudit','Audit');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>                        				        
				           <%}%>
				           <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){//Delete %>                   	        	                                   		      
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'DeleteAudit','Audit');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>                        				        				        
				           <%}%>
				      	<%}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>                        
                        <%}//end of 未裁撤%>
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>   
              
  <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                      	  <li>確認輸入資料無誤後, 按<font color="#666666">【確定】</font>即將本表上的資料, 於資料庫中建檔。</li>
                      	  <li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>                         
                          <li>欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>
        				  <li>如放棄修改或無修改之資料需輸入, 按<font color="#666666">【回上一頁】]即離開本程式。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
</table>
</form>
</body>
</html>
