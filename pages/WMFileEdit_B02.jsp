<%
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
//99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.List" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Calendar" %>
<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
   	
	System.out.println("WMFileEdit_B02.jsp -Begin");	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	System.out.println("S_MONTH="+S_MONTH);
	
	List data_div01 = null;
	if(act.equals("new")){
		StringBuffer sqlCmd = new StringBuffer();
		sqlCmd.append(" select * from b00_run_item order by input_order ");
        data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");	
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
	}
	System.out.println("data_div01.size="+data_div01.size());
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
<script language="JavaScript" src="js/menucontext_A03.js"></script> 
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
<table width="780" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#297A76">
  <tr>
    <td bordercolor="#FFFFFF"><table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
<!--        <tr> 
          <td><img src="images/topbanner_1.gif" width="780" height="103"></td>
        </tr>
-->       
        <tr> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
<table width="850" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="850" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000"><font size=4>【農業發展基金及農業天然災害救助基金貸放餘額統計】</font></font><font color='#000000' size=4></font> 
                          <%}%>
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="850" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                        <div align="right"><jsp:include page="getLoginUser.jsp?width=850" flush="true" /></div>
                    </tr>
                    <tr> 
                      <td><Table width=850 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>B01&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("3", "B02")%></td>
                          </tr>  
                          <%}%>                          
                          <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>基準日</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' <%if(act.equals("Edit")) out.print("disabled");%>>
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
                        <table width=850 border='1' align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99" class="sbody">
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC" rowspan=4><div align=center>業務別</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>當年度貸放</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>累計貸放</div></td>
                            <td bgcolor="#B1DEDC" colspan=4><div align=center>貸放餘額</div></td>
                          </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>數　量</div></td>
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>金　額</div></td>
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>數　量</div></td>
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>金　額</div></td>
                            <td bgcolor="#B1DEDC" rowspan=3><div align=center>數　量</div></td>
                            <td bgcolor="#B1DEDC" colspan=3><div align=center>金　額</div></td>
                          </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC" rowspan=2><div align=center>小計</div></td>
                            <td bgcolor="#B1DEDC" colspan=2><div align=center>按資金來源分</div></td>
                          </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td bgcolor="#B1DEDC"><div align=center>基金</div></td>
                            <td bgcolor="#B1DEDC"><div align=center>經辦機構</div></td>
                          </tr>
 <% 	int i = 0 ;
 		int loan_index = 0 ;
		boolean fontbold=false;
		//String fontsize="2";
		
		while( i < data_div01.size()){
 %>
 			              <tr bgcolor='e7e7e7' class="sbody">
			                <td bgcolor="#D8EFEE"><div align=left>
			                <%= ((String)((DataObject)data_div01.get(i)).getValue("fund_name")).trim()%></div>
			                <input type=hidden value="<%=(String)((DataObject)data_div01.get(i)).getValue("run_master_no")%>" name='run_master_no_c'>
			                <input type=hidden value="<%=(String)((DataObject)data_div01.get(i)).getValue("run_sub_no")%>" name='run_sub_no_c'>
			                <input type=hidden value="<%=(String)((DataObject)data_div01.get(i)).getValue("run_next_no")%>" name='run_next_no_c'>
			                </td>
			                <% if( ((DataObject)data_div01.get(i)).getValue("loan_cnt_year") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_cnt_year") != null && ((((DataObject)data_div01.get(i)).getValue("loan_cnt_year")).toString()).equals("0")) ){%>
			                       <td><div align=left><input type='text' name='loan_cnt_year' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_cnt_year' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_cnt_year")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_cnt_year"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_amt_year") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_year") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_year")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_amt_year' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>           
			                        <td><div align=left><input type='text' name='loan_amt_year' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_year")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_year"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_cnt_totacc") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_cnt_totacc") != null && ((((DataObject)data_div01.get(i)).getValue("loan_cnt_totacc")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_cnt_totacc' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_cnt_totacc' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_cnt_totacc")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_cnt_totacc"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_amt_totacc") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_totacc") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_totacc")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_amt_totacc' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_amt_totacc' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_totacc")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_totacc"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_cnt_bal") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_cnt_bal") != null && ((((DataObject)data_div01.get(i)).getValue("loan_cnt_bal")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_cnt_bal' class="small" value="" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_cnt_bal' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_cnt_bal")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_cnt_bal"))).toString())%>" size=5 maxlength=6 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_amt_bal_subtot") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_bal_subtot") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_subtot")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_amt_bal_subtot' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_amt_bal_subtot' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_subtot")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_bal_subtot"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_amt_bal_fund") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_bal_fund") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_fund")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_amt_bal_fund' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_amt_bal_fund' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_fund")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_bal_fund"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }
			                            
			                if( ((DataObject)data_div01.get(i)).getValue("loan_amt_bal_bank") == null ||  (((DataObject)data_div01.get(i)).getValue("loan_amt_bal_bank") != null && ((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_bank")).toString()).equals("0")) ){%>
			                    <td><div align=left><input type='text' name='loan_amt_bal_bank' class="small" value="" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% }else{ %>
			                        <td><div align=left><input type='text' name='loan_amt_bal_bank' class="small" value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("loan_amt_bal_bank")) == null ? "":(((DataObject)data_div01.get(i)).getValue("loan_amt_bal_bank"))).toString())%>" size=13 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)'></div></td>
			                <% } %>
			              </tr>
                <%	i++;
		}%> 			              
                    </table><BR>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr>
              </table></td>
              </tr>
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp?width=850" flush="true" /></div></td>
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
                        <td width="74"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','B02','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image91','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image91" width="66" height="25" border="0" id="Image91"></a></div></td>
         		<%}%>
         		<%if(act.equals("Edit")){%>         		
				        <td width="74"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','B02','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image91','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image91" width="66" height="25" border="0" id="Image91"></a></div></td>         		
						<td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','B02','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
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
          <td bgcolor="#FFFFFF"><table width="850" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="850" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="850" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#985000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增農業發展基金及農業天然災害救助基金貸放餘額統計。<%=ListArray.getDLIdName("1", "B02")%></li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="850" height="12"></div></td>
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
