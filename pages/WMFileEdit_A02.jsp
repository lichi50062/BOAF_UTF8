<%
//93.12.17 add 權限檢核 by 2295
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
//94.03.24 fix text field靠右 by 2295
//95.01.17 fix add 加上Insert/Update/Delete 加上 bank_type by 2295
//97.01.29 add 【990610】非會員授信總額-->申報總額.只有在計算比率時.才需扣除990611/990612
//				1.需 包含【990611】對直轄市、縣（市）政府、離島地區鄉(鎮、市)公所辦理之授信總額
//			           及【990612】非會員政策性農業專案貸款 
//97.04.18 add【990620】及【990310】申報時,皆需包含公庫存款  by 2295
//99.03.03 add 990610】非會員授信總額申報時,不含內部融資 by 2295
//99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.08.16 add 990811信用部固定資產淨值，無超過其淨值(未上線)
//				990812因購置或汰換安全維護或營業相關設備，經中央主管機關核准
//				990813因固定資產重估增值
//				990814因淨值降低 by 2295
//102.01.15 add 990421已申請符合逾放比率低於百分之一、其資本適足率高於百分之十且備抵保帳覆蓋率高於百分之一百經主管機關同意,農金局回文函號
//              990621已申請符合逾放比率低於百分之二且資本適足率高於百分之八經主管機關同意,農金局回文函號 by 2295
//102.04.02 fix 農金局回文文號改成農委會回文函號 by 2295
//104.02.09 add 990422已申請符合逾放比率低於百分之一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意
//              990622已申請符合逾放比率低於百分之一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意 by 2295
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

	System.out.println("S_MONTH="+S_MONTH);
	//String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	//System.out.println("Report="+Report);
	List data_div01 = null;
	String ncacno="ncacno";
	
	if(act.equals("new")){
	   if(bank_type.equals("7")){//漁會	      
	   	   ncacno = "ncacno_7";	       
	   }
	   data_div01 = Utility.getAcc_Code(ncacno,"A02","04");//法定比率表
	}else{
	   data_div01 = (List)request.getAttribute("data_div01");
	}
	System.out.println("data_div01.size="+data_div01.size());
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A02.permission == null");
    }else{
       System.out.println("WMFileEdit_A02.permission.size ="+permission.size());
               
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
<script language="JavaScript" src="js/menu.js"></script>
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
</script>
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
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A02")%>】</font></font><font color='#000000' size=4></font> 
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
                            <td colspan=2 bgcolor='e7e7e7'>A02&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A02")%></td>
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
                          
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td width=111 bgcolor="#D8EFEE"><a name="990110a"><div align=left>項目代碼</div></a></td>
                            <td width="316"> <div align=left>項目名稱</div></td>
                            <td width=160 bgcolor="#B1DEDC"> <div align=left>項目數值</div></td>
                          </tr>
 <% 	int i = 0 ;
		boolean fontbold=false;
		String bgcolor="#F2F2F2";
		//String fontsize="2";		
		String tmpAcc_Code_Amt = "";
		DataObject bean = null;
		while( i < data_div01.size()){
			bgcolor="#F2F2F2";
			fontbold=false;
			//fontsize="2";
			bean = (DataObject)data_div01.get(i);
		    String tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
		    
			tmpAcc_Code_Amt = bean.getValue("amt") == null ? "0":(bean.getValue("amt")).toString();
			//每個item變粗體字
			if(tmpAcc_Code.equals("990110")/*1.最近一年平均存放比率", "1.最近一年平均存放比率*/
			|| tmpAcc_Code.equals("990210")/*2.農會信用部對農會經濟事業門融通資金之限制*/
			|| tmpAcc_Code.equals("990310")/*3.非會員存款之額度限制*/
			|| tmpAcc_Code.equals("990410")/*4.贊助會員授信總額占贊助會員存款總額之比率*/
			|| tmpAcc_Code.equals("990510")/*5.辦理非會員無擔保銷費性貸款（1,000千元以下）之限制*/
			|| tmpAcc_Code.equals("990610")/*6.非會員授信總額占非會員存款總額之比率*/
			|| tmpAcc_Code.equals("990710")/*7.辦理自用住宅放款限額*/
			|| tmpAcc_Code.equals("990810")/*8.信用部固定資產淨額*/
			|| tmpAcc_Code.equals("990910")/*9.外弊風險之限制*/
			|| tmpAcc_Code.equals("991010")/*10.對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者為擔保授信限制*/
			|| tmpAcc_Code.equals("991110")/*11.對每一會員（含同戶家屬）及同一關係人放款最高限額*/
			|| tmpAcc_Code.equals("991210")/*12.對每一贊助會員及同一關係人之授信限額*/
			|| tmpAcc_Code.equals("991310")/*13.對同一非會員及同一關係人之授信限額*/
			|| tmpAcc_Code.equals("991410")/*14.最近決算年度*/
			){ 			
				fontbold=false;
				//fontsize="4";
				bgcolor = "#FFFFE6";			 
			}
%>
			<%if(!(tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			|| tmpAcc_Code.equals("990814")/*因淨值降低*/
			)){%> 
			<tr bgcolor='<%=bgcolor%>' class="sbody">			
			<%}%>
			<%//102.01.21 未上線 begin if(tmpAcc_Code.equals("990810")){/*信用部固定資產淨額*/%>
			<!-- td bgcolor="<%=bgcolor%>" rowspan="3" --!> 
			<%//}else	if(tmpAcc_Code.equals("990811")/*信用部固定資產淨值，無超過其淨值*/
			//|| tmpAcc_Code.equals("990812")){/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/%>			
			<!-- td bgcolor="<%=bgcolor%>" colspan="2" --!>	
			<%//}else if(tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			//|| tmpAcc_Code.equals("990814")){/*因淨值降低*/%>					
			<%//}else{%>  
			<td bgcolor="<%=bgcolor%>">	
			<%//102.01.21 未上線 eng}%>
			<%if(fontbold){%><b><%}%>	
			<%if(!(tmpAcc_Code.equals("990811")/*信用部固定資產淨值，無超過其淨值*/
			|| tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			|| tmpAcc_Code.equals("990814")/*因淨值降低*/
			)){%> 					
			<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
			<%}%>
			<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
			<input type=hidden name=acc_div value="01">
			<%if(tmpAcc_Code.equals("990421") || tmpAcc_Code.equals("990621") || tmpAcc_Code.equals("990422") || tmpAcc_Code.equals("990622")){ %>
			</td>			
			<td colspan="2">	
			<%}else{%>	
			<%if(!(tmpAcc_Code.equals("990811")/*信用部固定資產淨值，無超過其淨值*/
			|| tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			|| tmpAcc_Code.equals("990814")/*因淨值降低*/
			)){%> 	
			</td>
			
			<td>	
			<%}
			}%>
			<%if(fontbold){%><b><%}%>	

			<div align=left>
			<%if((((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
		    || (((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
		    
			<%if(tmpAcc_Code.equals("990811")/*信用部固定資產淨值，無超過其淨值*/
			|| tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			|| tmpAcc_Code.equals("990814")/*因淨值降低*/
			){%>
			   <%if(tmpAcc_Code.equals("990812")){/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/ %>
				<font color='red'>信用部固定資產淨額，不得超過其淨值但下列情形之一，不在此限(可複選)：<br></font>
			   <%}%>	
				<input type="checkbox" name="chk<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="" <%if((tmpAcc_Code.equals("990811") && act.equals("new")) || tmpAcc_Code_Amt.equals("1")) out.print("checked");%>>
				<input type="hidden" name='amt' value="<%=tmpAcc_Code_Amt%>">
			<%}%>
			
			<%if(tmpAcc_Code.equals("990421")){/*已申請符合逾放比率低於百分之一、其資本適足率高於百分之十且備抵保帳覆蓋率高於百分之一百經主管機關同意*/ %>				
			<input type="checkbox" name="chk<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="" <%if(tmpAcc_Code.equals("990421") && tmpAcc_Code_Amt.equals("1")) out.print("checked");%>>
			<input type="hidden" name='amt' value="<%=tmpAcc_Code_Amt%>">
			<%}%>	
			<%if(tmpAcc_Code.equals("990621")){/*申請符合逾放比率低於百分之二且資本適足率高於百分之八經主管機關同意,農金局回文文號*/ %>
			<input type="checkbox" name="chk<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="" <%if(tmpAcc_Code.equals("990621") && tmpAcc_Code_Amt.equals("1")) out.print("checked");%>>
			<input type="hidden" name='amt' value="<%=tmpAcc_Code_Amt%>">
			<%}%>	
			<%if(tmpAcc_Code.equals("990422")){/*已申請符合逾放比率低於百分之一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意*/ %>				
			<input type="checkbox" name="chk<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="" <%if(tmpAcc_Code.equals("990422") && tmpAcc_Code_Amt.equals("1")) out.print("checked");%>>
			<input type="hidden" name='amt' value="<%=tmpAcc_Code_Amt%>">
			<%}%>	
			<%if(tmpAcc_Code.equals("990622")){/*已申請符合逾放比率低於百分之 一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意*/ %>
			<input type="checkbox" name="chk<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="" <%if(tmpAcc_Code.equals("990622") && tmpAcc_Code_Amt.equals("1")) out.print("checked");%>>
			<input type="hidden" name='amt' value="<%=tmpAcc_Code_Amt%>">
			<%}%>		  
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>	
			<%if(tmpAcc_Code.equals("990421") || tmpAcc_Code.equals("990621") || tmpAcc_Code.equals("990422") || tmpAcc_Code.equals("990622")){/*農金局回文文號*/ %>	
			<br>農委會回文函號				
			<input type="text" name="txt<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>" value="<%if( ((DataObject)data_div01.get(i)).getValue("amt_name") != null  ) out.print((String)((DataObject)data_div01.get(i)).getValue("amt_name"));%>" >
			<%}%>				  	
			</div>
			<%if(fontbold){%></b><%}%>		
			<%if(!(tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/			
			)){%> 			
			</td>
			<%}%>
			<%if(!(tmpAcc_Code.equals("990811")/*信用部固定資產淨值，無超過其淨值*/
			|| tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/
			|| tmpAcc_Code.equals("990814")/*因淨值降低*/
			|| tmpAcc_Code.equals("990421")
			|| tmpAcc_Code.equals("990621")
			|| tmpAcc_Code.equals("990422")
			|| tmpAcc_Code.equals("990622")
			)){%> 			
			<td><a name="<%=tmpAcc_Code%>">
			<%if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){%>
				<input type='text' name='amt' value="" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			<%}else{%>
				<input type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right;'>
			<%}%>
				</a>
			</td>		
			<%}%>
			<%if(!(tmpAcc_Code.equals("990812")/*因購置或汰換安全維護或營業相關設備，經中央主管機關核准*/
			|| tmpAcc_Code.equals("990813")/*因固定資產重估增值*/			
			)){%> 						
		   </tr>
	    	<%}%>
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
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A02','','','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A02','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A02','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
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
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="626" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="600"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A02")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                          <li><font color='#0000BB'>【990620】非會員存款總額及【990310】申報時,皆需包含公庫存款。</font></li>	
                          <li><font color='#0000BB'>【990610】非會員授信總額申報時,不含內部融資。</font></li>	
                          <font color='red'>
                          <li>【990120】信用部月平均淨值<br>信用部淨值＝信用部權益及公積－統一農(漁)貸公積。</li>
                          <li>【990130】信用部月平均固定資產淨額<br>信用部固定資產淨額＝信用部固定資產－備抵折舊。</li>	
                          <li>【990210】內部融資餘額<br>1.依農授金字第0955070048號函應包含「農業發展基金放款」科目下之「內部融資」。<br>2.應大(等)於信用部資產負債表「內部融資」金額。</li>
                          <li>【990220】內部融資中、長期<br>【990220】小(等)於【990210】。</li>
                          <li>【990230】上年度信用部決算淨值<br>1.	信用部決算淨值為上會計年度經會員代表大會承認之盈餘分配後之淨值。<br>2.	若因淨值尚未完成分配，則請填列前年底信用部決算淨值，並註明年度【99141Y】。</li>	
                          <li>【990320】上年度農(漁)會決算淨值<br>1.	全體決算淨值為上會計年度經會員代表大會承認之盈餘分配後之淨值。<br>2.	若因淨值尚未完成分配，則請填列前年底信用部決算淨值，並註明年度【99141Y】。</li>
                          <li>【990510】非會員無擔保消費性貸款<br>消費性貸款係指對於房屋修繕、耐久性消費品（包括汽車）、支付學費及其他個人之小額貸款。（不含內部融資）</li>                           
                          <font color='#FF0000'>
                          <li>【990610】</font><font color='#0000BB'>非會員授信總額<br>1.需
							包含【990611】對直轄市、縣（市）政府、</font><font color='red'><font color="#0000BB">鄉(鎮、市)公所辦理之授信總額<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
							及【990612】</font></font><font color="#FF0000">非會員</font><font color='red'><font color="#0000BB">政策性農業專案貸款
							<br>2.【990611】對直轄市、縣（市）政府、鄉(鎮、市)公所辦理授信總額 ＝ <br>&nbsp;&nbsp; 【996111】對直轄市、縣（市）政府辦理授信總額 ＋【996112】對鄉(鎮、市)公所辦理授信總額。
							<br>3.【996112】對鄉(鎮、市)公所辦理授信總額 ＝ <br>&nbsp;&nbsp; 【996113】對鄉(鎮、市)公所授信經所隸屬之縣政府保證 ＋【996114】對鄉(鎮、市)公所授信
							<br>&nbsp;&nbsp; &nbsp;&nbsp; 未經所隸屬之縣政府保證。
							<br>4.基準日信用部資產負債表之「放款淨額」＋「備抵呆帳_放款」＋「備抵呆帳_催收款<br>&nbsp;&nbsp; 
							項」 ＝ 科目A99.【992140】+A02.【990410】+A02.</font>【990610】<font color="#0000BB">。</font><br>&nbsp; </li>
                          <li>【990620】非會員存款總額與【990630】月底日公庫存款總額<br>基準日信用部資產負債表之「存款總額」＝科目A99.【992130】+A02.【990420】+A02.【990310】。</li>	                         
                          <li>【990810】信用部固定資產淨額<br>應與信用部資產負債表「固定資產淨額」金額相符。</li>                       
                          <li>【990930】上年度信用部淨值<br>應與上年度資產負債表之淨值相符。</li>                       
                          <li>【99141Y】最近決算年度<br>需與【990230】及【990320】之年度配合。</li>                       
                          </font>
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
