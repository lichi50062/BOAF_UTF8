<%
//94.11.04 add 在台無住所外國人新台幣存款 by 2295
//94.11.14 add 自動計算的text field用灰色顯示 by 2295
//94.12.25 fix 按新增時.檢查ok而且不是INI者.直接執行"上月資料匯入" by 2295
//95.01.16 fix 合計-小計列,0也要顯示 by 2295
//95.01.17 fix add 加上Insert/Update/Delete 加上 bank_type by 2295
//95.09.08 add 拿掉新增/修改/刪除 by 2295 
//99.10.06 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>

<%	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");		
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	//農金局F01的起始年月
	String F01_APPLY_INI = ( request.getParameter("F01_APPLY_INI")==null ) ? "false" : (String)request.getParameter("F01_APPLY_INI");		
	//下月申報資料
	String NextMonthCreated = ( request.getParameter("NextMonthCreated")==null ) ? "false" : (String)request.getParameter("NextMonthCreated");		
	
	System.out.println("WMFileEdit_F01.act="+act);
	System.out.println("S_YEAR="+S_YEAR);	
	System.out.println("S_MONTH="+S_MONTH);	
	System.out.println("bank_type="+bank_type);		
	System.out.println("F01_APPLY_INI="+F01_APPLY_INI);		
	System.out.println("NextMonthCreated="+NextMonthCreated);		
	List data_div01 = null;	
	//if(!act.equals("new")){
		data_div01 = (List)request.getAttribute("data_div01");		
		if(data_div01 != null) System.out.println("data_div01.size="+data_div01.size());			
	//}
	
	//取得WMFileEdit的權限
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_F01.permission == null");
    }else{
       System.out.println("WMFileEdit_F01.permission.size ="+permission.size());
    }
%>
<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
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

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
<input type="hidden" name="act" value="">  
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="80"><img src="images/banner_bg1.gif" width="80" height="17"></td>
                      <td width="440"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("6", "F01")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="80"><img src="images/banner_bg1.gif" width="80" height="17"></td>
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
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="112"> <div align=left>申報年月</div></td>
                            <td colspan=2>                            
                            <input type='hidden' name="S_YEAR" value="<%=S_YEAR%>">
                            <input type='hidden' name="S_MONTH" value="<%=S_MONTH%>">
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' disabled>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH disabled>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%> ><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月份資料 </font>
        						<!--
        						<%//if(act.equals("new") && F01_APPLY_INI.equals("false")){//不為農金局設定之起始年月才可匯入上月資料%>
        						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        						<input type="button" name="LastMonthDataBtn" value="上月資料匯入" onclick="javascript:doSubmit(this.document.forms[0],'getLastMonthData','F01','','');">&nbsp;&nbsp; 
        						是否確定
        						<select name="LastMonthDataYN" size="1">
                              		<option selected>N</option>
                              		<option>Y</option>
                              	</select-->	
        						<%//}%>
                            </td>
                          </tr>
                          </table>
                          
                          <table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                                                                              
                          <tr bgcolor='e7e7e7'>
                           <td width="127" rowspan="2" bgcolor="#B1DEDC" class="sbody"><div align=left>&nbsp;&nbsp; 存款/帳戶類型</div>
                           </td>
                           <td width="80" rowspan="2" bgcolor="#B1DEDC" class="sbody"><div align="center">本月底戶數</div></td>
                           <td colspan="4" bgcolor="#B1DEDC" class="sbody"><div align="center">金額（元）</div></td>
                          </tr>
                          <tr bgcolor='e7e7e7'>
                           <td width="90" bgcolor="#B1DEDC" class="sbody"><div align="center">上月底餘額</div></td>
                           <td width="90" bgcolor="#B1DEDC" class="sbody"><div align="center">本月存入</div></td>
                           <td width="90" bgcolor="#B1DEDC" class="sbody"><div align="center">本月提出</div></td>
                           <td width="90" bgcolor="#B1DEDC" class="sbody"><div align="center">本月底餘額</div></td>
                          </tr>                          
                        <% 
                           String subtitle[] = {"個人","公司、行號、團體","外國專業投資機構","外國銀行","小計"};
                           String title[] = {"活期存款","活期儲蓄存款","定期存款","支票存款","合計"};
                           String titleidx[] = {"A","B","C","D","E"};
                           int d=0;
                           String tmpVar="";
                           String amt="";
                           String[] column={"acct_cnt_tm","bal_lm","dep_tm","wtd_tm","bal_tm"};                          
                        %>  
                        
                        <%for(int k=0;k<5;k++){%> 
                        <tr bgcolor='e7e7e7' class="sbody">
                          <td colspan="6" bgcolor="#D8EFEE"><div align="justify"><strong><%=title[k]%></strong></div></td>
                        </tr>
                        <%for(int j=0;j<5;j++){
                        //若為農金局設定之起始年月,則上月底餘額開放補輸入(INI)
                        //不為農金局設定之起始年月.則上月底餘額不開放輸入,只能使用上月資料匯入(不為INI)                        
                        %>
                        <tr bgcolor='#FFFFE6' class="sbody">                          
                          <td bgcolor="#FFFFE6"><div align=left><%=subtitle[j]%></div></td>                          
                          <%                                                    
                          for(int i=1;i<=5;i++){
                              //System.out.println("d="+d);
                              //System.out.println("get["+column[i-1]+"]");
                              tmpVar = ";amtXn5(this.form,this.form."+titleidx[k]+(j+1)+"5)";                                                    
                              //if(!act.equals("new") && ((DataObject)data_div01.get(d)).getValue(column[i-1]) != null ){
                              if( data_div01 != null && ((DataObject)data_div01.get(d)).getValue(column[i-1]) != null ){
                                 amt = (((DataObject)data_div01.get(d)).getValue(column[i-1])).toString();                                 
                              }                              
                          %>
                          <td>
                          
                          <input name="<%=titleidx[k]%><%=j+1%><%=i%>" type="text" value="<%/*95.01.16 合計-小計列,0也要顯示*/if(j+1==5 && titleidx[k].equals("E") && (!amt.equals(""))){out.print(Utility.setCommaFormat(amt));} else if(amt.equals("") || amt.equals("0")){out.print("");}else{out.print(Utility.setCommaFormat(amt));}%>" size="<%if(i==1)out.print("10"); else out.print("11");%>"  <%if(F01_APPLY_INI.equals("false")/*不為農金局設定之起始年月*/ && i==2) out.print("readonly ");%>  <%if(i==5 || j+1==5) out.print("readonly ");%> 
                          style='text-align: right; <%if((F01_APPLY_INI.equals("false")&& i==2) || (i==5 || j+1==5)) out.print(" color:#808080; background-color:#FFFFE6");%>'  
                          
                          onFocus='this.value=changeVal(this)' 
                          onBlur='checkPoint_focus(this);this.value=changeStr(this)<%if(j+1 <= 4 && i <=4) out.print(";subsum(this.form,this.form."+titleidx[k]+(j+1)+i+")");%><%if(j+1 <5 && i >1 && i< 5) out.print(tmpVar);%>;amtX55(this.form,this.form.<%=titleidx[k]%>55)'>                          
                          </td>
                          <%}
                          d++;
                          %>                          
                        </tr>
                        <%}//end of j ->個人,公司、行號、團體,外國專業投資機構,外國銀行,小計%>
                        <%}//end of k ->活期存款,活期儲蓄存款,定期存款,支票存款,合計%>
	          </table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
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
			 	<% 
			 	//如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消
			 	//95.09.08 拿掉新增/修改/刪除 by 2295
			 	%> 
			 	<%if(act.equals("getLastMonthData")) act="new";%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','F01','','','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit") && NextMonthCreated.equals("false")){//94.11.07無下月申報資料時,才可修改%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','F01','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','F01','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
                <%}%>        
                        <td width="80"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a></div></td>
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
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
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
                          <%if(act.equals("Edit") && NextMonthCreated.equals("true")){//94.11.07有下月申報資料時,僅可查詢%>
                          <li><font color="#FF0000">次一個月的申報資料已建檔,該申報年月的資料僅提供查詢功能</font></li>
                          <%}else{%>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("6", "F01")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【確定】</font>或<font color="#666666">【修改】</font>時,會一併執行線上檢核,需耗時5-7秒。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                          <%}%>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
            </table></td>
        </tr>        
      </table></td>
  </tr>
</table>
</form>
</body>

</html>
