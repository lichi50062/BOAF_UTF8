<%@ page contentType="text/html;charset=Big5" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%
	
	String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");		
	//String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	System.out.println("Report="+Report);
	List data_div01 = null;
	List data_div02 = null;
	List paramList = new ArrayList();
	
	if(act.equals("new")){
	    paramList.add(Report);
	    data_div01 = DBManager.QueryDB_SQLParam("select * from ncacno where acc_tr_type = ? and acc_div='01'",paramList,"update_date");//資負表 	    
		data_div02 = DBManager.QueryDB_SQLParam("select * from ncacno where acc_tr_type = ? and acc_div='02'",paramList,"update_date");//損益表 	    	
	}else if(act.equals("Edit")){
		data_div01 = (List)request.getAttribute("data_div01");
		data_div02 = (List)request.getAttribute("data_div02");
	}
	System.out.println("data_div01.size="+data_div01.size());
	System.out.println("data_div02.size="+data_div02.size());
%>
<html><head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: red;}
</style>
<title>線上編輯申報資料</title>
</head>
<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" >
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
<script language="JavaScript" src="js/menu.js"></script>
<script language="JavaScript" src="js/menucontext_A01.js"></script>
<script language="JavaScript">
showToolbar();
</script>
<script language="JavaScript">
function UpdateIt(){
if (document.all){
document.all["MainTable"].style.top = document.body.scrollTop;
setTimeout("UpdateIt()", 200);
}
}
UpdateIt();
</script>
<font color='#000000' size=4><br><b><center>線上編輯申報資料</center></b></font>
<Form name='frmFileEdit' method=post action='/pages/WMFileEdit.jsp'>
 <input type="hidden" name="act" value="">  
<table width=550 border=1 align='center'>
<div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div>  
</table>

<Table border=1 width=550 align=center>
    
    <tr bgcolor='#E7E7E7'>
		<td align='left'><div align=left>基準日</div></td>
		<td colspan=2>
        <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        年
        <select id="hide1" name=S_MONTH>
        <option></option>
        <%
        for (int j = 1; j <= 12; j++) {
        	if (j < 10){%>        	
        	<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            <%}else{%>
            <option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            <%}%>
        <%}%>
        </select>月
        </td>
    </tr>
    
	<tr bgcolor='e7e7e7'>		
		<td colspan=3><font size="5"><b><div align=center><a name="A01_div01">資產負債表</a></div></b></font></td>		
	</tr>
	<tr bgcolor='e7e7e7'>
		<td width=18%><div align=left>科目代碼</div></td>
		<td><div align=left>科目名稱</div></td>
		<td width=20%><div align=left>項目數值</div></td>
	</tr>
	<%  int i = 0 ;
		boolean fontbold=false;
		String fontsize="2";		
		while( i < data_div01.size()){

			fontbold=false;
			fontsize="2";
		    String tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
			
			//資產負債表
			if(tmpAcc_Code.equals("110000")/*流動資產*/       || tmpAcc_Code.equals("110300")/*存放行庫--合計*/
			|| tmpAcc_Code.equals("120000")/*放款*/           || tmpAcc_Code.equals("130000")/*基金及出資*/
			|| tmpAcc_Code.equals("140000")/*固定資產*/       || tmpAcc_Code.equals("150000")/*其他資產*/
			|| tmpAcc_Code.equals("160000")/*往來*/           || tmpAcc_Code.equals("210000")/*流動負債*/
			|| tmpAcc_Code.equals("220000")/*存款*/           || tmpAcc_Code.equals("240000")/*長期負債*/
			|| tmpAcc_Code.equals("250000")/*其他負債*/       || tmpAcc_Code.equals("260000")/*往來*/
			|| tmpAcc_Code.equals("310000")/*事業資金及公積*/ || tmpAcc_Code.equals("320000")/*盈虧及損益*/			
			){ 			
				fontbold=true;
				fontsize="4";
			}
			
			if(tmpAcc_Code.equals("200000")/*負債合計*/ || tmpAcc_Code.equals("300000")/*淨值合計*/
			|| tmpAcc_Code.equals("100000")/*資產合計*/ || tmpAcc_Code.equals("600000")/*負債及淨值合計*/
			){ 			
				fontbold=true;
				fontsize="5";
			}
			
			if(tmpAcc_Code.equals("110310")/*存放行庫--合作金庫--小計*/ || tmpAcc_Code.equals("110320")/*存放行庫--全國農業金庫--小計*/
			|| tmpAcc_Code.equals("120600")/*農業發展基基放款--小計*/   || tmpAcc_Code.equals("210200")/*透支行庫--小計*/
			|| tmpAcc_Code.equals("210400")/*短期借款--小計*/           || tmpAcc_Code.equals("240100")/*長期借款--合計*/
			|| tmpAcc_Code.equals("240210")/*借入農業發展基金放款資金--借入農建放款資金--小計*/ 
			|| tmpAcc_Code.equals("240200")/*借入農業發展基金放款資金--合計*/
			|| tmpAcc_Code.equals("240300")/*借入專案放款資金--合計*/
			){ 			
				fontbold=true;
				fontsize="3";
			}

			
	%>	
	<tr bgcolor='e7e7e7'>
	
		<td><font size="<%=fontsize%>">			
		<%if(fontbold){%><b><%}%>				
		<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
		<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
		<input type=hidden name=acc_div value="01">
		</font>
		<font size="<%=fontsize%>">	
		</td>
		<td>
		<font size="<%=fontsize%>">	
		<%if(fontbold){%><b><%}%>	

		<div align=left><%if((((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
						  || (((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
		<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>		
		</div>
		<%if(fontbold){%></b><%}%>		
		</font>
		</td>
		<font size="<%=fontsize%>">			
		<td><a name="<%=tmpAcc_Code%>">

		<%
		
		%>
		<input type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)'>
		</a>

		</font>
		</td>		
	</font>	
	</tr>
	
		
	<%	    i++;	
		}
	%>
	<tr bgcolor='white'><td colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
	<tr bgcolor='white'><td colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
	
	<tr bgcolor='white'><td colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>

	<tr bgcolor='e7e7e7'>		
		<td colspan=3><font size="5"><b><div align=center><a name="A01_div02">損益表</a></div></b></font></td>		
	</tr>
	<tr bgcolor='e7e7e7'>
		<td width=18%><div align=left>科目代碼</div></td>
		<td><div align=left>科目名稱</div></td>
		<td width=20%><div align=left>項目數值</div></td>
	</tr>
	<%  i = 0 ;
		fontbold=false;
		fontsize="2";
		
		while( i < data_div02.size()){
			fontbold=false;
			fontsize="2";

		    String tmpAcc_Code = ((String)((DataObject)data_div02.get(i)).getValue("acc_code")).trim();
		    
		    //損益表
			if(tmpAcc_Code.equals("520000")/*業務支出*/ || tmpAcc_Code.equals("522000")/*業務外支出*/
			|| tmpAcc_Code.equals("420000")/*業務收入*/	|| tmpAcc_Code.equals("422000")/*業務外收入*/ 
			||(tmpAcc_Code.equals("320300") && ((String)((DataObject)data_div02.get(i)).getValue("acc_div")).trim().equals("02"))/*本期損益*/){
				fontbold=true;
				fontsize="4";
			}
			if(tmpAcc_Code.equals("500000")/*合計*/ || tmpAcc_Code.equals("400000")/*合計*/){ 			
				fontbold=true;
				fontsize="5";
			}
			if(tmpAcc_Code.equals("520100")/*存款利息支出--合計*/ || tmpAcc_Code.equals("520200")/*借款利息支出--合計*/
			|| tmpAcc_Code.equals("420100")/*放款利息收入--合計*/ || tmpAcc_Code.equals("420300")/*存儲利息收入--合計*/){ 			
				fontbold=true;
				fontsize="3";
			}
	%>	
	<tr bgcolor='e7e7e7'>
	
		<td><font size="<%=fontsize%>">			
		<%if(fontbold){%><b><%}%>		

		<div align=left><%=(String)((DataObject)data_div02.get(i)).getValue("acc_code")%></div>
		<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div02.get(i)).getValue("acc_code")%>">
		<input type=hidden name=acc_div value="01">
		</font>
		<font size="<%=fontsize%>">	
		</td>
		<td>
		<font size="<%=fontsize%>">	
		<%if(fontbold){%><b><%}%>	
		
		<div align=left><%if((((String)((DataObject)data_div02.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
						  || (((String)((DataObject)data_div02.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
		<%=(String)((DataObject)data_div02.get(i)).getValue("acc_name")%>		
		</div>
		<%if(fontbold){%></b><%}%>		
		</font>
		</td>
		<font size="<%=fontsize%>">			
		<td><a name="<%=tmpAcc_Code%>">	

		<input type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div02.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div02.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)'>

		<a>	
		</font>
		</td>		
	</font>	
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
	<%if(act.equals("new")){%>   
	<input type='button' value='確定' onClick=doSubmit(this.form,'Insert','A01')>     
	<%}else{%>
	<input type='button' value='確定' onClick=doSubmit(this.form,'Update','A01')>     
	<input type='button' value='刪除' onClick=doSubmit(this.form,'Delete','A01')>  
	<%}%>
    <input type='button' value='取消' onClick=AskReset(this.form)>    
    
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
</body></html> 
