<%
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
//94.04.21 fix text field靠右 by 2295
//99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
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
   	
	System.out.println("WMFileEdit_M05.jsp");	
	//String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	System.out.println("S_MONTH="+S_MONTH);
	//String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	//System.out.println("Report="+Report);
	List data_div01 = null;
	List data_div02 = null;
	List data_div03 = null;
	if(act.equals("new")){
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
        sqlCmd.append(" select a.loan_unit_no as \"loan_unit_no\" ,a.loan_unit_name,b.data_range ,b.data_range_name,substr(b.data_range,1,3) as \"period_no\",substr(b.data_range,4,5) as \"item_no\" ");
        sqlCmd.append("   from m00_loan_unit a,m00_data_range_item b ");
        sqlCmd.append("  where b.report_no=? and b.data_range_type=? ");
        sqlCmd.append("  order by a.input_order,b.input_order 		");
        paramList.add("M05");
        paramList.add("C");
        
	    data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"update_date");//各貸款機構之小統計 	    
	    
	    sqlCmd.delete(0,sqlCmd.length());
		sqlCmd.append(" select loan_unit_no,loan_unit_name,'0ET' as \"fix_no\" ");
    	sqlCmd.append("   from m00_loan_unit ");
    	sqlCmd.append("  order by input_order ");
	    data_div02 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"update_date");//各貸款機構之累計保證項目 	    
	    
		sqlCmd.delete(0,sqlCmd.length());
		paramList = new ArrayList();
		sqlCmd.append(" select data_range as \"note_no\" from m00_data_range_item where report_no='M05' and data_range_type='N' order by input_order ");
		paramList.add("M05");
		paramList.add("N");
		System.out.println("sqlCmd="+sqlCmd);
		data_div03 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"update_date");//備註欄
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
		data_div02 = (List)request.getAttribute("data_div02");
		data_div03 = (List)request.getAttribute("data_div03");
	}
	System.out.println("data_div01.size="+data_div01.size());
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_M05.permission == null");
    }else{
       System.out.println("WMFileEdit_M05.permission.size ="+permission.size());
               
    }
%>
<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
input.small {font-size:8pt}
</style>
<%if(act.equals("Query")){%>
<title>申報資料查詢</title>
<%}else{%>
<title>線上編輯申報資料</title>
<%}%>

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

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" background="images/bg_1.gif" leftmargin="0">
<script language="JavaScript" src="js/menu.js"></script>
<!--
不需使用浮動視窗
<script language="JavaScript" src="js/menucontext_M05.js"></script> 
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
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="1250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="1250" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="450"><img src="images/banner_bg1.gif" width="450" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("4", "M05")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="450"><img src="images/banner_bg1.gif" width="450" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="1250" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp?width=1250" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><Table width=	1250 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>M05&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("4", "M05")%></td>
                          </tr>  
                          <%}%>
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="112"> <div align=left>基準日</div></td>
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
                            </td>
                          </tr>
                        <table width=1250 border='1' align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99" class="sbody">
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>貸款機構</div></td>
                            <td bgcolor="#B1DEDC" colspan=4><div align=center>&nbsp;</div></td>
                            <td bgcolor="#B1DEDC" colspan=12><div align=center>代償量值及原因</div></td>
                          </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>累計保證</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>代償案件</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>經營不善</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>週轉失靈</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>疾病或死亡</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>災歉</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>兼業失利</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>其他</div></td>
                          </tr>
                          
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>件數</div></td>
                            <td bgcolor="#B1DEDC"> <div align=center>金額</div></td>
                          </tr>
 <% 	int i = 0 ;
 		int loan_index = 0 ;
		boolean fontbold=false;
		String bgcolor="#F2F2F2";
		//String fontsize="2";		
		while( i < data_div01.size()){
			bgcolor="#F2F2F2";
			fontbold=false;
			//fontsize="2";
			//if(((String)((DataObject)data_div02.get(loan_index)).getValue("loan_unit_no")).equals("0")) bgcolor = "#FFFFE6";
		    String tmpData_Range = ((String)((DataObject)data_div01.get(i)).getValue("data_range")).trim();		  
		    if(tmpData_Range.equals("0MTRP")){%>
 			              <tr bgcolor='<%=bgcolor%>' class="sbody">
			                <td bgcolor="<%=bgcolor%>"><div align=left><b>
			                  <%=(String)((DataObject)data_div02.get(loan_index)).getValue("loan_unit_name")%>
			                  <input type=hidden value="<%=(String)((DataObject)data_div02.get(loan_index)).getValue("loan_unit_no")%>" name=loan_unit_no_c>
			                  <input type=hidden value="<%=(String)((DataObject)data_div02.get(loan_index)).getValue("fix_no")%>" name=fix_no_c>
			                </b></div></td>
			                <%if( ((DataObject)data_div02.get(loan_index)).getValue("guarantee_no_totacc") == null ||  (((DataObject)data_div02.get(loan_index)).getValue("guarantee_no_totacc") != null && ((((DataObject)data_div02.get(loan_index)).getValue("guarantee_no_totacc")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='guarantee_no_totacc' rowspan=10 class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='guarantee_no_totacc' rowspan=10 class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div02.get(loan_index)).getValue("guarantee_no_totacc")) == null ? "":(((DataObject)data_div02.get(loan_index)).getValue("guarantee_no_totacc"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div02.get(loan_index)).getValue("guarantee_amt_totacc") == null ||  (((DataObject)data_div02.get(loan_index)).getValue("guarantee_amt_totacc") != null && ((((DataObject)data_div02.get(loan_index)).getValue("guarantee_amt_totacc")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='guarantee_amt_totacc' rowspan=10 class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='guarantee_amt_totacc' rowspan=10 class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div02.get(loan_index)).getValue("guarantee_amt_totacc")) == null ? "":(((DataObject)data_div02.get(loan_index)).getValue("guarantee_amt_totacc"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			              </tr>
		<%
				loan_index=loan_index+1;
			}
		%>
			              <tr bgcolor='<%=bgcolor%>' class="sbody">
			                <td bgcolor="<%=bgcolor%>"><div align=left>
			                  <%=(String)((DataObject)data_div01.get(i)).getValue("data_range_name")%>
			                  <input type=hidden name=period_no value="<%=(String)((DataObject)data_div01.get(i)).getValue("period_no")%>">
			                  <input type=hidden name=item_no value="<%=(String)((DataObject)data_div01.get(i)).getValue("item_no")%>">
			                  <input type=hidden name=loan_unit_no value="<%=(String)((DataObject)data_div01.get(i)).getValue("loan_unit_no")%>">
			                </div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <td><div align=left>&nbsp;</div></td>
			                <%if( ((DataObject)data_div01.get(i)).getValue("repay_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("repay_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("repay_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='repay_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='repay_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("repay_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("repay_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("repay_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("repay_amt") != null && ((((DataObject)data_div01.get(i)).getValue("repay_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='repay_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='repay_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("repay_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("repay_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("run_notgood_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("run_notgood_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("run_notgood_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='run_notgood_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='run_notgood_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("run_notgood_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("run_notgood_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("run_notgood_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("run_notgood_amt") != null && ((((DataObject)data_div01.get(i)).getValue("run_notgood_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='run_notgood_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='run_notgood_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("run_notgood_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("run_notgood_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("turn_out_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("turn_out_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("turn_out_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='turn_out_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='turn_out_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("turn_out_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("turn_out_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("turn_out_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("turn_out_amt") != null && ((((DataObject)data_div01.get(i)).getValue("turn_out_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='turn_out_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='turn_out_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("turn_out_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("turn_out_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("diease_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("diease_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("diease_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='diease_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='diease_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("diease_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("diease_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("dieaserepay_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("dieaserepay_amt") != null && ((((DataObject)data_div01.get(i)).getValue("dieaserepay_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='dieaserepay_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='dieaserepay_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("dieaserepay_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("dieaserepay_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("disaster_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("disaster_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("disaster_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='disaster_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='disaster_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("disaster_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("disaster_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("disaster_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("disaster_amt") != null && ((((DataObject)data_div01.get(i)).getValue("disaster_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='disaster_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='disaster_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("disaster_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("disaster_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("corun_out_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("corun_out_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("corun_out_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='corun_out_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='corun_out_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("corun_out_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("corun_out_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("corun_out_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("corun_out_amt") != null && ((((DataObject)data_div01.get(i)).getValue("corun_out_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='corun_out_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='corun_out_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("corun_out_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("corun_out_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("other_cnt") == null ||  (((DataObject)data_div01.get(i)).getValue("other_cnt") != null && ((((DataObject)data_div01.get(i)).getValue("other_cnt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='other_cnt' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='other_cnt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("other_cnt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("other_cnt"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			                <%if( ((DataObject)data_div01.get(i)).getValue("other_amt") == null ||  (((DataObject)data_div01.get(i)).getValue("other_amt") != null && ((((DataObject)data_div01.get(i)).getValue("other_amt")).toString()).equals("0")) ){%>
			                	<td><div align=left><input type='text' name='other_amt' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}else{%>
			                	<td><div align=left><input type='text' name='other_amt' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("other_amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("other_amt"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'></div></td>
			                <%}%>
			              </tr>
	<%
			i++;	
		}
	%>
			              <tr bgcolor='#e7e7e7' class="sbody">
			                <td colspan=19>註：代償案件經追索全案收回者計
			                <%if( ((DataObject)data_div03.get(0)).getValue("note_amt_rate") == null ||  (((DataObject)data_div03.get(0)).getValue("note_amt_rate") != null && ((((DataObject)data_div03.get(0)).getValue("note_amt_rate")).toString()).equals("0")) ){%>
			                	<input type='text' name='note_amt_rate' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}else{%>
			                	<input type='text' name='note_amt_rate' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(0)).getValue("note_amt_rate")) == null ? "":(((DataObject)data_div03.get(0)).getValue("note_amt_rate"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}%>
			                <input type='hidden' name='note_no' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(0)).getValue("note_no")) == null ? "":(((DataObject)data_div03.get(0)).getValue("note_no"))).toString())%>">
			                ，獲償金額
			                <%if( ((DataObject)data_div03.get(1)).getValue("note_amt_rate") == null ||  (((DataObject)data_div03.get(1)).getValue("note_amt_rate") != null && ((((DataObject)data_div03.get(1)).getValue("note_amt_rate")).toString()).equals("0")) ){%>
			                	<input type='text' name='note_amt_rate' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}else{%>
			                	<input type='text' name='note_amt_rate' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(1)).getValue("note_amt_rate")) == null ? "":(((DataObject)data_div03.get(1)).getValue("note_amt_rate"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}%>
			                <input type='hidden' name='note_no' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(1)).getValue("note_no")) == null ? "":(((DataObject)data_div03.get(1)).getValue("note_no"))).toString())%>">
			                ，另部份收回者計
			                <%if( ((DataObject)data_div03.get(2)).getValue("note_amt_rate") == null ||  (((DataObject)data_div03.get(2)).getValue("note_amt_rate") != null && ((((DataObject)data_div03.get(2)).getValue("note_amt_rate")).toString()).equals("0")) ){%>
			               		<input type='text' name='note_amt_rate' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			               	<%}else{%>
			               		<input type='text' name='note_amt_rate' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(2)).getValue("note_amt_rate")) == null ? "":(((DataObject)data_div03.get(2)).getValue("note_amt_rate"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			               	<%}%>
			                <input type='hidden' name='note_no' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(2)).getValue("note_no")) == null ? "":(((DataObject)data_div03.get(2)).getValue("note_no"))).toString())%>">
			                ，獲償金額
			                <%if( ((DataObject)data_div03.get(3)).getValue("note_amt_rate") == null ||  (((DataObject)data_div03.get(3)).getValue("note_amt_rate") != null && ((((DataObject)data_div03.get(3)).getValue("note_amt_rate")).toString()).equals("0")) ){%>
			                	<input type='text' name='note_amt_rate' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}else{%>
			                	<input type='text' name='note_amt_rate' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(3)).getValue("note_amt_rate")) == null ? "":(((DataObject)data_div03.get(3)).getValue("note_amt_rate"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			                <%}%>
			                <input type='hidden' name='note_no' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div03.get(3)).getValue("note_no")) == null ? "":(((DataObject)data_div03.get(3)).getValue("note_no"))).toString())%>">
			                元。
			                </td>
			              </tr>
	          </Table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp?width=1250" flush="true" /></div></td>              
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','M05','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','M05','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','M05','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
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
          <td bgcolor="#FFFFFF"><table width="1250" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="1250" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="1250" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "M05")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="1250" height="12"></div></td>
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
</body>

</html>
