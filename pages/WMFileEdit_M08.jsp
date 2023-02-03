<%
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
//94.04.21 fix text field靠右 by 2295
//99.10.06 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>


<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
   	
	//String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");	
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");
	
	System.out.println("[*** egg test file WMFileEdit_M08.jsp -- Begin ***]");		
	System.out.println("Input Report_no="+Report_no);
	System.out.println("Input S_YEAR="+S_YEAR);
	System.out.println("Input S_MONTH="+S_MONTH);
	
	//宣告List div01
	List data_div01 = null;
	if(act.equals("new")){
		StringBuffer sqlCmd = new StringBuffer();
		//宣告div01
		sqlCmd.append(" select a.id_no,a.id_name,b.data_range,b.data_range_name from m00_id_item a,m00_data_range_item b ");
		sqlCmd.append(" where ((a.id_no<>'0' and b.data_range_type<>'T') or (a.id_no='0' and b.data_range_type='T'))");
		sqlCmd.append(" and b.report_no='M08' order by a.input_order,b.input_order");
		data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");	//保留數值型態欄位資料	
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
	}
	System.out.println("data_div01.size=" + data_div01.size());
%>

<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
</style>

<title>線上編輯申報資料</title>


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
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">

<script language="JavaScript" src="js/menu.js"></script>
<!--
不需使用浮動視窗
<script language="JavaScript" src="js/menucontext_M08.js"></script> 
<script language="JavaScript">
showToolbar();
</script>
-->
<script language="JavaScript">
function UpdateIt(){
if (document.all){
document.all["MainTable"].style.top = document.body.scrollTop;
setTimeout("UpdateIt()", 200);
}
}
UpdateIt();
</script>
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
 <input type="hidden" name="act" value="">  
<table width="650" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#297A76">
  <tr>
    <td bordercolor="#FFFFFF"><table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
<!--        <tr> 
          <td><img src="images/topbanner_1.gif" width="780" height="103"></td>
        </tr>
-->       
        <tr> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
<table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="106"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="*"><b> 
                        <center>
                          <b> 
                          <center>
                          	<%if(act.equals("Query")){%>
                          	 	<font color='#000000' size=4>申報資料查詢</font> 
                          	<%}else{%>
                            	<font color='#000000' size=4>線上編輯</font><font color="#CC0000">【保證案件</font></font></font><font size="4" color="#CC0000">_身份別_</font><font color='#000000' size=4><font color="#CC0000">月報表】</font></font><font color='#000000' size=4></font> 
                          	<%}%>
                          </center>
                          </b> 
                        </center>
                        </b> </td>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp?width=650" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><Table width=650 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          
                          <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>基準日</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" <%if(act.equals("Edit")) out.print("disabled");%> size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年 
        						<select id="hide1" name=S_MONTH <%if(act.equals("Edit")) out.print("disabled");%>>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        								<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            							<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
                            </font>
                            </td>
                          </tr>
                          
						<table width=650 border='1' align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99" class="sbody">
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC"> <div align=center>區分<BR>保證項目</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>保證<BR>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>融資金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>保證金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>保證餘額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>保證餘額<BR>(結構比%)</div></td>
                          </tr>
 			              <tr bgcolor='e7e7e7' class="sbody"><td bgcolor="#D8EFEE" colspan=9><div align=left>保證案件身份別 </div></td></tr>
 						<% 	//表身資料處理
 						int i = 0 ;
 						System.out.println("data_div01.size=[" + data_div01.size() + "]");
						while( i < data_div01.size()){
		    				String tmpData_Range = ((String)((DataObject)data_div01.get(i)).getValue("data_range")).trim();
		    				tmpData_Range = tmpData_Range.substring(tmpData_Range.length()-2);
		    				
		    				if(tmpData_Range.equals("MM") ||tmpData_Range.equals("MT") ){%>
 			              		<tr bgcolor='e7e7e7' class="sbody"><td bgcolor="#D8EFEE" colspan=9><div align=left><b><%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("id_name")) == null ? "":(((DataObject)data_div01.get(i)).getValue("id_name"))).toString())%></div></td></b></div></td></tr>
 			              	<%}%>
 			              		<tr bgcolor='e7e7e7' class="sbody">
			               			<td bgcolor="#D8EFEE"><div align=left><%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("data_range_name")) == null ? "":(((DataObject)data_div01.get(i)).getValue("data_range_name"))).toString())%></div></td>
			                			<input type=hidden name=id_no value="<%=(String)((DataObject)data_div01.get(i)).getValue("id_no")%>">
			                			<input type=hidden name=data_range value="<%=(String)((DataObject)data_div01.get(i)).getValue("data_range")%>">
			                		<%if( ((DataObject)data_div01.get(i)).getValue("guarantee_no_month") == null ||  (((DataObject)data_div01.get(i)).getValue("guarantee_no_month") != null && ((((DataObject)data_div01.get(i)).getValue("guarantee_no_month")).toString()).equals("0")) ){%>
			                			<td><div align=left><input type='text' name='guarantee_no_month' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}else{%>
			                			<td><div align=left><input type='text' name='guarantee_no_month' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("guarantee_no_month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("guarantee_no_month"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}%>
			                		<%if( ((DataObject)data_div01.get(i)).getValue("loan_amt_month") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_month") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_month")).toString()).equals("0")) ){%>
			                			<td><div align=left><input type='text' name='loan_amt_month' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}else{%>
			                			<td><div align=left><input type='text' name='loan_amt_month' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_month"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}%>
			                		<%if( ((DataObject)data_div01.get(i)).getValue("guarantee_amt_month") == null ||  (((DataObject)data_div01.get(i)).getValue("guarantee_amt_month") != null && ((((DataObject)data_div01.get(i)).getValue("guarantee_amt_month")).toString()).equals("0")) ){%>
			                			<td><div align=left><input type='text' name='guarantee_amt_month' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}else{%>
			                			<td><div align=left><input type='text' name='guarantee_amt_month' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("guarantee_amt_month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("guarantee_amt_month"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}%>
			                		<%if( ((DataObject)data_div01.get(i)).getValue("guarantee_bal_month") == null ||  (((DataObject)data_div01.get(i)).getValue("guarantee_bal_month") != null && ((((DataObject)data_div01.get(i)).getValue("guarantee_bal_month")).toString()).equals("0")) ){%>
			                			<td><div align=left><input type='text' name='guarantee_bal_month' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}else{%>
			               	 			<td><div align=left><input type='text' name='guarantee_bal_month' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("guarantee_bal_month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("guarantee_bal_month"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}%>
			                		<%if( ((DataObject)data_div01.get(i)).getValue("guarantee_bal_p") == null ||  (((DataObject)data_div01.get(i)).getValue("guarantee_bal_p") != null && ((((DataObject)data_div01.get(i)).getValue("guarantee_bal_p")).toString()).equals("0")) ){%>
			                			<td><div align=left><input type='text' name='guarantee_bal_p' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}else{%>
			                			<td><div align=left><input type='text' name='guarantee_bal_p' class="small" value="<%=Utility.getPercentNumber(((((DataObject)data_div01.get(i)).getValue("guarantee_bal_p")) == null ? "":(((DataObject)data_div01.get(i)).getValue("guarantee_bal_p"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'></div></td>
			                		<%}%>
			              		</tr>
						<%	i++;
						}
						%>

                    </table><BR>
              <tr>                  
                	<td><div align="right"><jsp:include page="getMaintainUser.jsp?width=650" flush="true" /></div></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                      	<%if(act.equals("new")){%>       
                       		<td width="74"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','M08','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image91','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image91" width="66" height="25" border="0" id="Image91"></a></div></td>
         				<%}%>
         				<%if(act.equals("Edit")){%>
				        	<td width="74"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','M08','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image91','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image91" width="66" height="25" border="0" id="Image91"></a></div></td>
							<td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','M08','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
						<%}%>				
         				<%if(!act.equals("Query")){%>       
                        	<td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                		<%}%>			 	 
                        	<td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image81" width="80" height="25" border="0" id="Image81"></a></div></td>
                     </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><table width="650" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="650" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="650" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明  
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增保證案件月報表。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="650" height="12"></div></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
<% System.out.println("[*** egg test file WMFileEdit_M08.jsp -- End ***]");%>
</body>
</html>
