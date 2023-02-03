<%
//95.11.27 create by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");					
	//營運中/已裁撤===================================================================================
	String cancel_no = ( session.getAttribute("CANCEL_NO")==null ) ? "N" : (String)session.getAttribute("CANCEL_NO");				
	//========================================================================================================
	String bank_type = (session.getAttribute("nowbank_type")==null)?"6":(String)session.getAttribute("nowbank_type");
	//String title=(bank_type.equals("6"))?"農會":"漁會";	
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");		
	System.out.println("FR0066WB.szExcelAction="+szExcelAction);
	String hsien_id = ( session.getAttribute("HSIEN_ID")==null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");				
	System.out.println("FR0066WB.hsien_id="+hsien_id);
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");	
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");            
   	String Unit = (session.getAttribute("Unit")==null)?"":(String)session.getAttribute("Unit");   	
   	//95.11.14 取得登入者資訊.DS_bank_type=================================================================================================
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
   	String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
    //==============================================================================================================    	       	
	System.out.println("S_MONTH="+S_MONTH);
	System.out.println("bank_type="+bank_type);
	System.out.println("DS_bank_type="+DS_bank_type);			
	String sqlCmd = " select bn01.bn_type,bn01.bank_type,wlx01.hsien_id,bn01.bank_no,bn01.bank_name "
	              + " from bn01  LEFT JOIN WLX01 on bn01.bank_no = WLX01.bank_no "
	              + "                            and bn01.bank_type in('6','7')"
	              + " order by wlx01.hsien_id,bn01.bank_no ";
    List tbankList = DBManager.QueryDB_SQLParam(sqlCmd,null,"");
	
    //取得FR0066WB的權限
	Properties permission = ( session.getAttribute("FR0066WB")==null ) ? new Properties() : (Properties)session.getAttribute("FR0066WB"); 
	if(permission == null){
       System.out.println("FR0066WB.permission == null");
    }else{
       System.out.println("FR0066WB.permission.size ="+permission.size());               
    }	
    
    // XML Ducument for 總機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
    out.println("<datalist>");     
    if(tbankList != null){   
    	for(int i=0;i< tbankList.size(); i++) {
        	DataObject bean =(DataObject)tbankList.get(i);
        	out.println("<data>");        
        	out.println("<BnType>"+bean.getValue("bn_type")+"</BnType>");
        	out.println("<BankType>"+bean.getValue("bank_type")+"</BankType>");
        	out.println("<HsienId>"+bean.getValue("hsien_id")+"</HsienId>");
        	out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
        	out.println("<bankName>"+bean.getValue("bank_no")+bean.getValue("bank_name")+"</bankName>");
        	out.println("</data>");
        	//System.out.println("<option>"+bean.getValue("bank_no")+"&nbsp;"+bean.getValue("bank_name")+"</option>");
    	}
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 總機構代碼 end 	
%>



<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
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

function doSubmit(cnd){
    this.document.forms[0].action = "/pages/FR0066WB_Excel.jsp?act="+cnd;	
    this.document.forms[0].target = '_self';
	this.document.forms[0].submit();   	     
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='BankListfrm'>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="600" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
                <td width="*" class="title_font">法定比率分析統計表 </td>          
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><img src="images/space_1.gif" width="8" height="8"></td>
        </tr>
        <tr> 
          <td><table width="600" border="1" align="center" cellpadding="0" cellspacing="0" class="bordercolor">
              <tr> 
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3"><table width="600" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> <div align="right">
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                      	  <%if(permission != null && permission.get("P") != null && permission.get("P").equals("Y")){//Print %>                   	        	                                   		     			        
                      	  <a href="javascript:doSubmit('createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image41" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image51" width="66" height="25" border="0" id="Image51"></a> 
                          <!--a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image61','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image61" width="80" height="25" border="0" id="Image61"></a--> 
                        </div></td>
                    </tr>
                    <!--tr> 
                      <td class="menu_bgcolor"> <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">1.金融機構</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('RptColumn')"><font color='black'>2.報表欄位</font></a></td>                            
                            <td width="100"><a href="javascript:doSubmit('RptOrder')"><font color='black'>3.排序欄位</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr-->                    
                     
                    <tr> 
                      <td class="body_bgcolor"> <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
                             <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'><font color='#000000'>年                            
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
                            </td>
                         </tr>    
                    	 
                    	 <tr>
                         <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">農(漁)會別 :</span>
                            <select size="1" name="bank_type" onchange="javascript:changeOption(document.forms[0],'change');">                                                          
                              <option value ='6' <%if((!bank_type.equals("")) && bank_type.equals("6")) out.print("selected");%>>農會</option>                                                                                          
                              <option value ='7' <%if((!bank_type.equals("")) && bank_type.equals("7")) out.print("selected");%>>漁會</option>                                                            
                             </select>
		                  </td>
                         </tr>  
                    
                         <tr>
                         <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">金額單位 :</span>
                            <select size="1" name="Unit">
                              <option value ='1' <%if((!Unit.equals("")) && Unit.equals("1")) out.print("selected");%>>元</option>
                              <option value ='1000' <%if((!Unit.equals("")) && Unit.equals("1000")) out.print("selected");%>>千元</option>
                              <option value ='10000' <%if((!Unit.equals("")) && Unit.equals("10000")) out.print("selected");%>>萬元</option>
                              <option value ='1000000' <%if((!Unit.equals("")) && Unit.equals("1000000")) out.print("selected");%>>百萬元</option>
                              <option value ='10000000' <%if((!Unit.equals("")) && Unit.equals("10000000")) out.print("selected");%>>千萬元</option>
                              <option value ='100000000' <%if((!Unit.equals("")) && Unit.equals("100000000")) out.print("selected");%>>億元</option>
                            </select>
		                  </td>
                          </tr>          
                          <tr>
                          <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">營運中/裁撤別 :</span> 	 
                           <select name='CANCEL_NO' onchange="javascript:changeOption(document.forms[0],'change');">
                           	<option  value="N" <%if((!cancel_no.equals("")) && cancel_no.equals("N")) out.print("selected");%>>營運中</option>
                           	<option  value="Y" <%if((!cancel_no.equals("")) && cancel_no.equals("Y")) out.print("selected");%>>已裁撤</option>
                           </select>
                          </td>
                          </tr>
                          
                          <tr> 
                          <%
                          List hsien_id_data = DBManager.QueryDB_SQLParam("select distinct hsien_id,hsien_name from cd01",null,""); 
                          %>
                            <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">縣市別 :</span>                                
                               <select name="HSIEN_ID" onchange="javascript:changeOption(document.forms[0],'');">                               
                                <option value="ALL">全部</option>
                                <%for(int i=0;i<hsien_id_data.size();i++){%>                                
                                <option value="<%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")%>"
                                <%if(((String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")).equals(hsien_id)) out.print("selected");%>
                                ><%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_name")%></option>                            
                                <%}%>
                              </select>
                            </td>
                          </tr>
                          
                          <tr>
                          <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">金融機構 :</span> 	 
                              <select name="BankListSrc" ondblclick="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);fn_changeBankListSrc(document.forms[0]);" style="width: 292;">							                           
                              <INPUT type="hidden" name="BankListDst">		
                          </td>
                          </tr>
                          
                          <tr> 
                          <td class="body_bgcolor"><div align="center">
                            <table width="555" border="0" cellpadding="1" cellspacing="1">
                              <tr> 
                              <td width="34"><img src="images/print_1.gif" width="34" height="34"></td>
                              <td width="222"><font color="#CC6600">本報表採用A4紙張直印</font></td>
                              <!--td width="293" align=right><font color="red" size=2>註:Office 2003版Excel現僅提供「下載報表」功能</font></td-->                              
                              </tr>                            
                              </table>
                           </div></td>
                        </tr>
                        </table></td>
                    </tr>                    
                  </table></td>
              </tr>
            </table></td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=BankList><!--//BankList儲存已勾選的金融機構代碼-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=clearbtnFieldList><!--//儲存是否清除btnFieldList-->
</form>
<script language="JavaScript" >
<!--

<%
//從session裡把勾選的金融機構代碼讀出來.放在BankListDst
if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){ 
   System.out.println("DS003W_BankList.BankList="+(String)session.getAttribute("BankList"));
%>
var bnlist;
bnlist = '<%=(String)session.getAttribute("BankList")%>';
var a = bnlist.split(',');
for (var i =0; i < a.length; i ++){
	var j = a[i].split('+');
	this.document.forms[0].BankListDst.options[i] = new Option(j[1], j[0]);
}
<%}%>

setSelect(this.document.forms[0].HSIEN_ID,"<%=hsien_id%>");
setSelect(this.document.forms[0].CANCEL_NO,"<%=cancel_no%>");
changeOption(this.document.forms[0],'');
function clearBankList(){
 <%
	session.setAttribute("BankList",null);//清除已勾選的BankList
 %>
}
-->
</script>

</body>
</html>
