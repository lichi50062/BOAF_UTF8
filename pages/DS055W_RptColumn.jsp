<%
//102.07.09 created N年內及目前月份經營統計資料 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS055W";	
	String act = Utility.getTrimString(dataMap.get("act"));	
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	String menuBank_type = (session.getAttribute("menuBank_type")==null)?"":(String)session.getAttribute("menuBank_type");
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");	
	System.out.println(report_no+"_RptColumn.szExcelAction="+szExcelAction);		
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");
	String S_YEAR1 = (session.getAttribute("S_YEAR1")==null)?"":(String)session.getAttribute("S_YEAR1");
   	String Unit = (session.getAttribute("Unit")==null)?"":(String)session.getAttribute("Unit");   	   	
	//取得登入者資訊=================================================================================================
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
    //==============================================================================================================    	    
%>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--

function doSubmit(report_no,cnd){   
   if(cnd == 'createRpt'){      
      if(!chInput(this.document.forms[0],'')) return;//95.12.04 add 起始日期不可大於結束日期
      if(this.document.forms[0].S_YEAR1.value==""){
  		alert("請輸入N年內");
  		return;
  	  }
      if(this.document.forms[0].FieldListDst.length == 0){        
      	 alert('報表欄位必須選擇');
      	 return;
      }
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }   
   
   MoveSelectToBtn(this.document.forms[0].btnFieldList, this.document.forms[0].FieldListDst);		         
   fn_ShowPanel(report_no,cnd);      
}
//所有欄位只能輸入數字
function IsNum(){ 
       return ((event.keyCode >= 48) && (event.keyCode <= 57)); 
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#'>
<input type='hidden' name='menuBank_type' value='<%=bank_type %>' >
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="750" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td><table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
                <td width="*" class="title_font"><%=Utility.getPgName(report_no)%></td>
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><img src="images/space_1.gif" width="8" height="8"></td>
        </tr>        
        <tr> 
          <td><table width="750" border="1" align="center" cellpadding="0" cellspacing="0" class="bordercolor">
              <tr> 
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3"><table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> <div align="right">
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                          <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %> 
                      	  <a href="javascript:doSubmit('<%=report_no%>','createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData_AgriBank('RptColumn');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody">                            
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">1.報表欄位</font></a>
                            </td>
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptStyle')"><font color='black'>2.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr class="sbody">
					  <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
					     <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' onkeypress="event.returnValue=IsNum()">年                             
							<select id="hide1" name=S_MONTH>        						
							<%
								for (int j = 1; j <= 12; j++) {
								if (j < 10){%>        	
								<option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
								<%}else{%>
								<option value=<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
								<%}%>
							<%}%>
							</select><font color='#000000'>月</font>
							<input type=hidden name=S_DATE value=''>
					    </td>
					</tr>
                    <tr class="sbody">
                         <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">農(漁)會別 :</span>
                            <select size="1" name="bank_type">                            
                              <%if(menuBank_type.equals("6") || bank_type.equals("6")){//95.11.10 有農會的menu時,才可顯示農會%>
                              <option value ='6' <%if((!bank_type.equals("")) && bank_type.equals("6")) out.print("selected");%>>農會</option>                                                            
                              <%}%>
                              <%if(menuBank_type.equals("7") || bank_type.equals("7")){//95.11.10 有漁會的menu時,才可顯示漁會%>
                              <option value ='7' <%if((!bank_type.equals("")) && bank_type.equals("7")) out.print("selected");%>>漁會</option>                              
                              <%}%>
                              <%if(!bank_type.equals("") && (muser_bank_type.equals("2") || muser_id.equals("A111111111"))){
                              	    //95.11.10 登入者為A111111111 or 農金局時,才可顯示農漁會%>                              
                              <option value ='ALL' <%if((!bank_type.equals("")) && bank_type.equals("ALL")) out.print("selected");%>>農漁會</option>                              
                              <%}%>
                            </select>
		                  </td>
                    </tr>  
                    <%@include file="./include/DS_Unit.include" %><!-- 金額單位-->
                    <tr class="sbody">
                         <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">N年內 :</span>
		                  <input type='text' name='S_YEAR1' value="<%=S_YEAR1%>" size='6' maxlength='3' onkeypress="event.returnValue=IsNum()">
		                  </td>
                    </tr>  
                    <%@include file="./include/DS_RptColumn_AgriBank.include" %><!-- 報表欄位.可挑選項目-->
                  </table></td>
              </tr>
               
            </table></td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=FieldList value='<%=request.getAttribute("FieldList")%>'><!--//FieldList儲存所有的報表欄位名稱-->
<INPUT type="hidden" name=btnFieldList><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構代碼-->
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.04儲存已勾選的結束日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.04儲存已勾選的結束日期-月-->
</form>
<script language="JavaScript" >
<!-- 
    
<%//從session裡把勾選的報表欄位讀出來.放在FieldListDst 	
if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){ 
   System.out.println("DS055W_RptColumn.btnFieldList="+(String)session.getAttribute("btnFieldList"));
%>
	var bnlist;
	bnlist = '<%=(String)session.getAttribute("btnFieldList")%>';
	<%
	session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位
    %>
	var a = bnlist.split(',');
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		this.document.forms[0].FieldListDst.options[i] = new Option(j[1], j[0]);
	}
<%}%>

fn_loadFieldList(this.document.forms[0]);//顯示所有的報表欄位名稱

-->
</script>

</body>
</html>
