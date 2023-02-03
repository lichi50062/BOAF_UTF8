<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<%
    List paramList = new ArrayList();
	String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");		
	String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	paramList.add(Report);
	paramList.add(Acc_Div);
	List data = DBManager.QueryDB_SQLParam("select * from ncacno where acc_tr_type = ? and acc_div= ?",paramList,"update_date");
	
	if(data != null){
		System.out.println("data != null");
	}else{
		System.out.println("data == null");
	}
%>


<HTML>
<HEAD>
<TITLE>線上編輯申報資料</TITLE>
</HEAD>
<BODY bgColor=white>
<font color='#000000' size=4><b><center>線上編輯申報資料</center></b></font>
<Form name='frmWM002W' method=post action='/pages/WMFileEdit.jsp' onSubmit='return checkShowInsert_A01(this,'insert');'>
<input type=hidden name=Function value=insert>
<table width=550 border=1 align='center'>
<div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div>  
</table>

<Table border=1 width=550 align=center>
    <tr bgcolor='#E7E7E7'>
		<td align='left'><div align=left>基準日</div></td>
		<td colspan=2>
        <input type='text' name='S_YEAR' size='3' maxlength='3' onblur='CheckYear(this)'>
        年
        <select name=S_MONTH>
        <option></option>
        <%
        for (int j = 1; j <= 12; j++) {
        	if (j < 10){%>        	
        	<option value=0<%=j%>>0<%=j%></option>        		
            <%}else{%>
            <option value=<%=j%>><%=j%></option>
            <%}%>
        <%}%>
        </select>月
        </td>
    </tr>

	<tr bgcolor='e7e7e7'>
		<td width=18%><div align=left>科目代碼</div></td>
		<td><div align=left>科目名稱</div></td>
		<td width=20%><div align=left>項目數值</div></td>
	</tr>
	<%  int i = 0 ;
		while( i <data.size()){%>
	<tr bgcolor='e7e7e7'>
		<td>
		<div align=left><%=(String)((DataObject)data.get(i)).getValue("acc_code")%></div>
		<input type=hidden name=acc_code value="<%=(String)((DataObject)data.get(i)).getValue("acc_code")%>"></td>
		<td>
		<div align=left><%=(String)((DataObject)data.get(i)).getValue("acc_name")%></div></td>
		<td><input type='text' name='amt' size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)'></td>
	</tr>
	<%	    i++;	
		}
	%>
	<tr bgcolor='#E7E7E7'>
	<td align='left'>申報者姓名</td>
	<td colspan=2>
	<input type='text' size=20 name=INPUT_NAME>
	</td>
	</tr>

	<tr bgcolor='#E7E7E7'>
	<td align='left'>申報者電話</td>
	<td colspan=2>
	<input type='text' size=20 name=INPUT_TEL>(區域號碼以"-"區隔,分機以"#"區隔)
	</td>
	</tr>
</Table>
<table width='550' border='1' align='center'>
     <div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>     
</table>

<table border=0 align=center width=550>
	<tr><th>
	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%>    
    <input type='submit' value='確定'>
    <input type='reset' value='取消'>    
    <input type='button' value='回上一頁' onClick='history.back()'>    
    </th>
    </tr>
</Table>

</Form>


<hr>
<p><font color='#FF0000'><strong>使用說明:</strong></font></p>
<ul>
<li>本網頁提供新增<%=ListArray.getDLIdName("1", Report)%>.</li>
<li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知.</li>
<li>確認資料無誤後, 按[確定]即將本網頁上的資料, 於資料庫中新增. </li>
<li>按[取消]即重新輸入資料. </li>
<li>點選所列之[回上一頁]則放棄資料, 回至前一畫面.</li>
</ul><hr>
</BODY>
</HTML>

