<%
//95.05.18 add 增加A99法定比率資料 by 2295
//95.05.30 add 992440,修改畫面格式 by 2295
//95.09.25 add 992510~992640,修改畫面格式 by 2295
//95.11.07 add 992710~992730,修改畫面格式 by 2295
//97.02.18 add 992550 無擔保消費性貸款中之逾期放款 
//             992650 無擔保消費性貸款中之應予觀察放款  by 2295
//95.12.27 fix A99 title by 2295
//99.10.06 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//100.03.28 fix wlx01 加入M_year條件，查詢年大於等於100 帶入100 by2479 
//101.02.01 fix wlx01 加入M_year條件，查詢年大於等於100 帶入100 by 2295
//101.06.29 add 992300增加說明(依農金局95.2.7  農授金字第  0955010504  號函之定義) by 2295
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
    
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	
	String m_year = YEAR;
	
	//登入者的tbank_no
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no"); 
	//所點選的nowtbank_no
	//若有點選的tbank_no,則顯示所點選的tbank_no
	tbank_no = ( session.getAttribute("nowtbank_no")==null ) ? tbank_no : (String)session.getAttribute("nowtbank_no");		
	
    //農金局F01的起始年月
	String F01_APPLY_INI = ( request.getParameter("F01_APPLY_INI")==null ) ? "false" : (String)request.getParameter("F01_APPLY_INI");			
	System.out.println("WMFileEdit_A01.act="+act);
	System.out.println("S_YEAR="+S_YEAR);	
	S_MONTH =String.valueOf(Integer.parseInt(S_MONTH));
	System.out.println("S_MONTH="+S_MONTH);	
	System.out.println("bank_type="+bank_type);	
	System.out.println("F01_APPLY_INI="+F01_APPLY_INI);				
	List data_div01 = null;
	
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String ncacno = "ncacno";
	sqlCmd.append("select credit_staff_num from wlx01 where bank_no = ? and m_year=?");
	paramList.add(tbank_no);
	m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
	paramList.add(m_year);
	List wlx01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"credit_staff_num");
	System.out.println("wlx01.size="+wlx01.size());
	
	if(act.equals("new")){
	    if(bank_type.equals("7")){//漁會	       	       	       
	       ncacno = "ncacno_7";	    
	    }
	    data_div01 = Utility.getAcc_Code(ncacno,"A99","99"); ///法定比率延申表
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
	}
	System.out.println("data_div01.size="+data_div01.size());
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A99.permission == null");
    }else{
       System.out.println("WMFileEdit_A99.permission.size ="+permission.size());
               
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
<!--script language="JavaScript" src="js/menu.js"></script>
<script language="JavaScript" src="js/menucontext_A02.js"></script>
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
</script-->
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
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A99")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
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
                      <td><Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=3 bgcolor='e7e7e7'>A99&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A99")%></td>
                          </tr>  
                          <%}%>
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="112"> <div align=left>基準日</div></td>
                            <td colspan=3 bgcolor='e7e7e7'>
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
        						</select><font color='#000000'>月 </font>        						
                            </td>
                          </tr>
                          
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td width=111 bgcolor="#D8EFEE"> <div align=left>項目代碼</div></td>
                            <td colspan="2" width="300"> <div align=left>項目名稱</div></td>
                            <td width=177 bgcolor="#B1DEDC"> <div align=left>項目數值</div></td>
                          </tr>
 <% 	int i = 0 ;
		boolean fontbold=false;
		String bgcolor="#F2F2F2";
		//String fontsize="2";		
		while( i < data_div01.size()){
			bgcolor="#F2F2F2";
			fontbold=false;
			//fontsize="2";
		    String tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
			
			//每個item變粗體字
			if(tmpAcc_Code.equals("992300")/*信用部員工人數*/			
			){ 			
				fontbold=false;
				//fontsize="4";
				bgcolor = "#FFFFE6";			 
			}
%>
			<tr bgcolor='<%=bgcolor%>' class="sbody">			
			<td bgcolor="<%=bgcolor%>">			
			<%if(fontbold){%><b><%}%>				
			<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
			<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
			<input type=hidden name=acc_div value="01">
			<%if(tmpAcc_Code.equals("992300")){//從事信用業務之員工人數
				 //if(!act.equals("new")){%>
				    <input type=hidden name=wlx01_credit_staff_num value="<%out.print(((DataObject)wlx01.get(0)).getValue("credit_staff_num") == null ?"":(((DataObject)wlx01.get(0)).getValue("credit_staff_num")).toString());%>">
			<%	 //}   				      
			  }
			%>
			</td>
									
			<%if(fontbold){%><b><%}%>	
		    <%if(!(tmpAcc_Code.equals("992410") || tmpAcc_Code.equals("992420") || tmpAcc_Code.equals("992440")
		         ||tmpAcc_Code.equals("992510") || tmpAcc_Code.equals("992520") || tmpAcc_Code.equals("992530") || tmpAcc_Code.equals("992540") || tmpAcc_Code.equals("992550") 
                 ||tmpAcc_Code.equals("992610") || tmpAcc_Code.equals("992620") || tmpAcc_Code.equals("992630") || tmpAcc_Code.equals("992640") || tmpAcc_Code.equals("992650")
                 ||tmpAcc_Code.equals("992710") || tmpAcc_Code.equals("992720") || tmpAcc_Code.equals("992730")
		    )){//acc_code not in ('992410','992420','992440','992510','992520','992530','992540','992610','992620','992630','992640')%>
		    <td colspan="2" width="293" height="34" bgcolor="<%=bgcolor%>">		
			<div align=left><%if((((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
							  || (((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>		
			<%if(tmpAcc_Code.equals("992300")){//從事信用業務之員工人數
				     //if(!act.equals("new")){
				        if(wlx01 != null && wlx01.size() != 0){
				           out.print("<font color='red'><br>(依農金局95.2.7  農授金字第  0955010504  號函之定義)</font>");//101.06.29 add 992300增加說明 
				           out.print("<br>原總機構基本資料維護之從事信用業務員工人數:");
				           out.print(((DataObject)wlx01.get(0)).getValue("credit_staff_num") == null ?"":(((DataObject)wlx01.get(0)).getValue("credit_staff_num")).toString());				  
				        }   
				     //} 
				  }
				%>
			</div>
			<%if(fontbold){%></b><%}%>		
			<%}else{%>			
			<%if(tmpAcc_Code.equals("992410")){%>
			<td rowspan="3" width="19" height="59">
			<div align=left>		
            <p align="center">應提流動準備					
			</div>
			</td>			
			<td width="268" height="6">			
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>										
			<%}//end of 992410%>
			<%if(tmpAcc_Code.equals("992510")){//95.09.25 add 992510%>
			<td rowspan="5" width="19" height="80">
			<div align=left>		
            <p align="center">逾期放款					
			</div>
			</td>			
			<td width="268" height="6">			
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>											
			<%}//end of 992510%>
			<%if(tmpAcc_Code.equals("992610")){//95.09.25 add 992610%>
			<td rowspan="5" width="19" height="80">
			<div align=left>		
            <p align="center">應予觀察放款					
			</div>
			</td>			
			<td width="268" height="6">			
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>									
			<%}//end of 992610%>						
			<%if(tmpAcc_Code.equals("992710")){//95.11.7 add 992710%>
			<td rowspan="3" width="19" height="80">
			<div align=left>		
            <p align="center">建築放款				
			</div>
			</td>			
			<td width="268" height="6">			
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>									
			<%}//end of 992610%>	
			<%if(tmpAcc_Code.equals("992440") || tmpAcc_Code.equals("992520") || tmpAcc_Code.equals("992620") || tmpAcc_Code.equals("992720")){%>
			<td width="268" height="36">
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>	
			</td>							
			<%}//end of 992440.992520.992620.992720%>
			
			<%if(tmpAcc_Code.equals("992420")|| tmpAcc_Code.equals("992530") || tmpAcc_Code.equals("992540") || tmpAcc_Code.equals("992550") || tmpAcc_Code.equals("992630") || tmpAcc_Code.equals("992640") || tmpAcc_Code.equals("992650") || tmpAcc_Code.equals("992730")){%>
			<td width="268" height="21">
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>	
			</td>			
			<%}//end of 992420.992530.992540.992630.992640%>
			<%}//end of acc_code in ('992410','992420','992440','992510','992520','992530','992540','992550','992610','992620','992630','992640','992650') %>
			</td>
			
			<td width="175" height="34" ><a name="<%=tmpAcc_Code%>">
			<%if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){%>
				<input type='text' name='amt' value="<%if(tmpAcc_Code.equals("992300")){//從事信用業務之員工人數
				     if(act.equals("new")){
				        System.out.println("992300");
				        if(wlx01 != null && wlx01.size() != 0){
				          System.out.println("wlx01 != null");
				         out.print( ((DataObject)wlx01.get(0)).getValue("credit_staff_num") == null ?"":(((DataObject)wlx01.get(0)).getValue("credit_staff_num")).toString());
				        } 
				     } 
				  }
				%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			<%}else{%>
				<input type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			<%}%>
				</a>				
			</td>		
		
		</tr>
	
		
	<%	    i++;	
		}
	%>
	          </Table></td>
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
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A99','','','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A99','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A99','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
                <%}%>        
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','images/bt_05b.gif',1)"><img src="images/bt_05.gif" name="Image81" width="80" height="25" border="0" id="Image81"></a></div></td>
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
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A99")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>                          
                          <li><font color="red">【992100】農(漁)會全體資產總額：需要扣除抵銷科目後之資產總額</font></li>
                          <li><font color="red">【992110】農(漁)會全體淨值＝基準日"綜合資產負債表"中<br>【事業資金及公積】＋【盈虧及損益】</font></li>
                          <li><font color="red">【992120】農(漁)會全體事業本期損益：農(漁)會全體事業本期損益＝基準日" 綜合資產負<br>&nbsp;&nbsp;債表"中【本期損益】之科目</font></li>
                          <li><font color="red">【992150】無擔保消費性貸款：消費性貸款係指對於房屋修繕、耐久性消費品（包括汽車）、<br>&nbsp;&nbsp;支付學費及其他個人之小額貸款。（不含內部融資）</font></li>
                          <li><font color="red">【992170】以轄區外縣市土地建物為擔(副)保品之戶數、【992180】以轄區外縣市土地建物為<br>&nbsp;&nbsp;擔(副)保品之金額：轄區外縣市土地建物為擔(副)保品係指所轄縣市以外土地建物為擔(副)保品</font></li>
                          <li><font color="red">【992300】信用部員工人數：依農金局95.2.7農授金字第0955010504號函之定義</font></li>                          
                          <li><font color="red">逾期放款比率項目中，科目【992510】+【992520】+【992530】+【992540】之金額總和需與A01_資產負債及損益與逾期放款資料中之逾期放款金額科目【990000】作跨表檢核。</font></li>                          
                          <li><font color="red">應予觀察放款比率項目中，科目【992610】+【992620】+【992630】+【992640】之金額總和，需與A04_資產品質分析資料中之應予觀察放款金額項次科目【840710】+【840720】+【840731】+【840732】+【840733】+【840734】+【840735】之金額總和作跨表檢核。</font></li>                          
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
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
