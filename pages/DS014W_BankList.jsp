<%
//95.11.13 create by 2295
//95.11.15 add 區分BOAF.MIS配色 by 2295
//         add 有農/漁會的menu時,才可顯示農/漁會;登入者為地方主管機關時,只能看到自己的主管機關
//         add 可選機構代號權限設定 by 2295
//		   add 登入者為A111111111 or 農金局時,才可選全部 by 2295
//95.12.06 add 增加年季區間 by 2295
//         add 起始日期不可大於結束日期 by 2295
//95.12.07 add 金融機構代號若本來選全部->各信用部 or 各信用部->全部,清空已選報表欄位/排序欄位 by 2295
//99.05.14 fix 配合縣市合併調整 程式 by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	//查詢條件值 
	Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS014W";	
	String act = Utility.getTrimString(dataMap.get("act"));					
	//營運中/已裁撤===================================================================================
	String cancel_no = ( session.getAttribute("CANCEL_NO")==null ) ? "N" : (String)session.getAttribute("CANCEL_NO");				
	//========================================================================================================
	String bank_type = (session.getAttribute("nowbank_type")==null)?"6":(String)session.getAttribute("nowbank_type");
	//String title=(bank_type.equals("6"))?"農會":"漁會";	
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");		
	System.out.println("DS014W_BankList.szExcelAction="+szExcelAction);
	String hsien_id = ( session.getAttribute("HSIEN_ID")==null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");				
	System.out.println("DS014W_BankList.hsien_id="+hsien_id);
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");
	String E_YEAR = (session.getAttribute("E_YEAR")==null)?YEAR:(String)session.getAttribute("E_YEAR");
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");            
	String E_MONTH = (session.getAttribute("E_MONTH")==null)?MONTH:(String)session.getAttribute("E_MONTH");                              
   	String Unit = (session.getAttribute("Unit")==null)?"":(String)session.getAttribute("Unit");   	
   	//95.11.15 取得登入者資訊.DS_bank_type=================================================================================================
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
    String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
    String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
    //==============================================================================================================    	      	  		
	System.out.println("S_MONTH="+S_MONTH);
	System.out.println("bank_type="+bank_type);
	System.out.println("DS_bank_type="+DS_bank_type);			
%>
<%@include file="./include/DS_bank_no_hsien_id.include" %>


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
   if(cnd == 'createRpt'){      
      if(this.document.forms[0].BankListDst.length == 0){      	 
      	 alert('金融機構代碼必須選擇');
      	 return;
      }
      if(!chInput(this.document.forms[0],'BankList')) return;//95.12.06 add 起始日期不可大於結束日期
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }
   MoveSelectToBtn(this.document.forms[0].BankList, this.document.forms[0].BankListDst);	
   fn_ShowPanel('<%=report_no%>',cnd);    
}



function ResetAllData(){
    if(confirm("確定要清除已選定的資料嗎？")){  	
        this.document.forms[0].BankListDst.length = 0;
        this.document.forms[0].HSIEN_ID[0].selected=true;	   
        changeOption(this.document.forms[0],'');
        //clearBankList();95.12.07
	}
	return;	
}
function hideOption(){//機構類別為地方主管機關時.不可選縣市別
    if(this.document.forms[0].bank_type.value == 'B'){       
       this.document.forms[0].HSIEN_ID.disabled = true;        
    }else{
       this.document.forms[0].HSIEN_ID.disabled = false;        
    }
}

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='BankListfrm'>
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
                <td width="230"><img src="images/banner_bg1.gif" width="230" height="17"></td>
                <td width="*" class="title_font">金融機構警示帳戶調查資料 </td>
                <td width="230"><img src="images/banner_bg1.gif" width="230" height="17"></td>
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
                      	  <a href="javascript:doSubmit('createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image41" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image51" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image61','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image61" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">1.金融機構</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('RptColumn')"><font color='black'>2.報表欄位</font></a></td>                            
                            <td width="100"><a href="javascript:doSubmit('RptOrder')"><font color='black'>3.排序欄位</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>                    
                     
                    <tr> 
                      <td class="body_bgcolor"> <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                        <!--查詢年月-->
                         <tr class="sbody">
                           <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
                              <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' onchange="javascript:changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);changeOption(document.forms[0],'change');"><font color='#000000'>年                             
                         		<select id="hide1" name=S_MONTH>        						
                         		<%
                         			for (int j = 1; j <= 12; j++) {
                         			if(("DS013W".equals(report_no) || "DS014W".equals(report_no) || "DS015W".equals(report_no)) && (j > 4)) break;
                         			if (j < 10){%>        	
                         			<option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
                         			<%}else{%>
                         			<option value=<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
                         			<%}%>
                         		<%}%>
                         		</select><font color='#000000'><%if("DS013W".equals(report_no) || "DS014W".equals(report_no) || "DS015W".equals(report_no)) out.print("季"); else out.print("月");%></font>~
                         	<input type='text' name='E_YEAR' value="<%=E_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'><font color='#000000'>年	
                         	<select id="hide1" name=E_MONTH>        						
                         		<%
                         			for (int j = 1; j <= 12; j++) {
                         			if(("DS013W".equals(report_no) || "DS014W".equals(report_no) || "DS015W".equals(report_no)) && (j > 4)) break;
                         			if (j < 10){%>        	
                         			<option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
                         			<%}else{%>
                         			<option value=<%=j%> <%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
                         			<%}%>
                         		<%}%>
                         		</select><font color='#000000'><%if("DS013W".equals(report_no) || "DS014W".equals(report_no) || "DS015W".equals(report_no)) out.print("季"); else out.print("月");%></font>
                         		<input type=hidden name=S_DATE value=''>
                         		<input type=hidden name=E_DATE value=''>
                             </td>
                         </tr>  
                    	 
                    	 <tr class="sbody">
                         <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">機構類別 :</span>
                            <select size="1" name="bank_type" onchange="javascript:hideOption();changeOption(document.forms[0],'change');">                            
                              <%if(DS_bank_type.equals("6")){//95.11.10 有農會的menu時,才可顯示農會%>
                              <option value ='6' <%if((!bank_type.equals("")) && bank_type.equals("6")) out.print("selected");%>>農會</option>                                                            
                              <%}%>
                              <%if(DS_bank_type.equals("7")){//95.11.10 有漁會的menu時,才可顯示漁會%>
                              <option value ='7' <%if((!bank_type.equals("")) && bank_type.equals("7")) out.print("selected");%>>漁會</option>                              
                              <%}%>
                              <%if(DS_bank_type.equals("B")){//95.11.13 有地方主管機關的menu時,才可顯示地方主管機關%>                              
                              <option value ='B' <%if((!bank_type.equals("")) && bank_type.equals("B")) out.print("selected");%>>地方主管機關</option>                                                            
                              <%}%> 
                            </select>
		                  </td>
                         </tr>  
                    
                         <%@include file="./include/DS_Unit.include" %><!-- 金額單位-->
                         <%@include file="./include/DS_Cancel_No_Hsien_ID.include" %><!-- 1.營運中/裁撤別 2.縣市別-->
                        </table></td>
                    </tr>
                    <tr> 
                      <td class="body_bgcolor"> <table width="750" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
                            <tr> 
                            <td width="215">
                            <table>
                            <tr><td align="center" class="chooseitem_bgcolor">可選擇項目</td></tr>
                            <tr><td>  
                            <select multiple  size=10  name="BankListSrc" ondblclick="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);fn_changeBankListSrc(document.forms[0]);" style="width: 292; height: 190">							
							</select>
							</td></tr>
							</table>
                            </td>
                            <td width="52"><table width="40" border="0" align="center" cellpadding="3" cellspacing="3">
                                <tr> 
                                  <td>
                                  <div align="center">                                 
                                  <a href="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);fn_changeBankListSrc(document.forms[0]);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr> 
                                  <td>
                                  <div align="center">                                  
                                  <a href="javascript:moveallsel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);"><img src="images/arrow_rightall.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr> 
                                  <td>
                                  <div align="center">                                  
                                  <a href="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);fn_changeBankListSrc(document.forms[0]);""><img src="images/arrow_left.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr> 
                                  <td height="22">
                                  <div align="center">                                  
                                  <a href="javascript:moveallsel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);"><img src="images/arrow_leftall.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                              </table></td>
                            <td width="340">
                            <table>
                            <tr><td align="center" class="chooseitem_bgcolor">已選擇項目</td></tr> 
                            <tr><td>
                            <select multiple size=10  name="BankListDst" ondblclick="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);fn_changeBankListSrc(document.forms[0]);"" style="width: 292; height: 190">							
							</select>
							</td></tr>
							</table>
                          </tr>
                        </table></td>
                    </tr>
                     <tr> 
                 	  <td class="body_bgcolor">
                  	    <table width="750" border="0" cellpadding="1" cellspacing="1">
                    	<tr>                           
                          <td width="750" align=left><font color="red" size=2>註：『已選擇項目』清單選取之項目說明</font></td>                              
                    	</tr>                              
                        <!--tr>                           
                          <td width="750" align=left><font color="red" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)選取[全部]：係以「各縣市彙總表」之報表格式列印 </font></td>                              
                        </tr-->                              
                        <tr>                           
                          <td width="750" align=left><font color="red" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;選取[各信用部/地方主管機關]：按「金融機構警示帳戶調查明細資料」的報表內容列印</font></td>                              
                        </tr>                                                  
                        </table>
                      </td>
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
   System.out.println("DS014W_BankList.BankList="+(String)session.getAttribute("BankList"));
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
hideOption();
//changeOption(this.document.forms[0],'');
changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);
/*95.12.07
function clearBankList(){
 <%
	//session.setAttribute("BankList",null);//清除已勾選的BankList
 %>
}*/
-->
</script>

</body>
</html>
