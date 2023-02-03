<%
// 95.04.27 fix 漁會A01放款合計120700=120700+120900,其餘顯示原amt by 2295
// 95.05.03 fix 加入A01.990000逾期放款金額 by 2295
//          add 若上月無資料,則不顯示上月資料匯入 by 2295
// 95.05.16 add A01.990000/A01逾放合計.合計 by 2295
// 95.05.16 fix 若為120601/120602/120603/120604不累加 by 2295		   	
// 95.05.16 add 逾放合計金額不可大於A01放款合計 by 2295
// 95.08.10 fix 拿掉960500.合計位置16改成15 by 2295
// 95.08.16 fix 6個月~未滿1年,開放輸入 by 2295
// 95.10.18 fix 拿掉顯示逾放比率 by 2295
// 95.11.09 fix 拿掉計算逾放比率 by 2295
// 99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295 
//103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
    
    //fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	//農金局F01的起始年月
	String F01_APPLY_INI = ( request.getParameter("F01_APPLY_INI")==null ) ? "false" : (String)request.getParameter("F01_APPLY_INI");		
	
	System.out.println("WMFileEdit_A01.act="+act);
	System.out.println("S_YEAR="+S_YEAR);	
	System.out.println("S_MONTH="+S_MONTH);	
	System.out.println("bank_type="+bank_type);	
	System.out.println("F01_APPLY_INI="+F01_APPLY_INI);			
	List data_div01 = null;
	List data_div02 = null;
	String alertMsg = "";
	StringBuffer sqlCmd = new StringBuffer();
	StringBuffer sqlCmd_sum = new StringBuffer();
	String ncacno = "ncacno";
	List datahasLast = (List)request.getAttribute("datahasLast");//上月資料	
	List paramList = new ArrayList();
	
	ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
	if(bank_type.equals("7")){   
	   //103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
	   if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
          ncacno = "ncacno_7_rule";
       }
    }  		
	if(act.equals("new")){
	   data_div01 = Utility.getAcc_Code(ncacno,"A06","08");  	  	   
	}else{
	   data_div01 = (List)request.getAttribute("data_div01");		
	   if(data_div01.size() == 0){
	      data_div01 = Utility.getAcc_Code(ncacno,"A06","08");  		      
	      alertMsg = "無上月資料可供匯入";   
	   }
	}	
	if(bank_type.equals("6")){
	   sqlCmd.append(" select A01.amt,A01.acc_code ");
 	   sqlCmd.append(" from A01 LEFT JOIN ncacno ON A01.acc_code = ncacno.acc_code  ");
 	   sqlCmd.append(" where A01.m_year = ?");
 	   sqlCmd.append(" and A01.m_month = ?");
 	   sqlCmd.append(" and bank_code=?");
 	   sqlCmd.append(" and (ncacno.acc_div='08' or ncacno.acc_code='990000')");
 	   sqlCmd.append(" order by acc_range "); 
 	   paramList.add(S_YEAR);
 	   paramList.add(S_MONTH);
 	   paramList.add(bank_code);
 	   
 	   sqlCmd_sum.append(" select sum(amt) as a01970000 ");
 	   sqlCmd_sum.append(" from (select A01.amt,A01.acc_code ");
 	   sqlCmd_sum.append(" from A01 LEFT JOIN ncacno ON A01.acc_code = ncacno.acc_code  ");
 	   sqlCmd_sum.append(" where A01.m_year = ?");
 	   sqlCmd_sum.append(" and A01.m_month = ?");
 	   sqlCmd_sum.append(" and bank_code=?");
 	   sqlCmd_sum.append(" and (ncacno.acc_div='08' and (ncacno.acc_code <> '990000' and ncacno.acc_code <> '120600'))");
 	   sqlCmd_sum.append(" order by acc_range )");
 	   sqlCmd_sum.append(" where acc_code not in ('120700','120900') ");      
 		      
	}else{
	   //103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
	   if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){	
	   	  sqlCmd.append(" select A01.amt,A01.acc_code ");
 	      sqlCmd.append(" from A01 LEFT JOIN ncacno_7_rule ON A01.acc_code = ncacno_7_rule.acc_code  ");
 	      sqlCmd.append(" where A01.m_year = ?");
 	      sqlCmd.append(" and A01.m_month = ?");
 	      sqlCmd.append(" and bank_code=?");
 	      sqlCmd.append(" and (ncacno_7_rule.acc_div='08' or ncacno_7_rule.acc_code='990000')");
 	      sqlCmd.append(" order by acc_range "); 
 	      paramList.add(S_YEAR);
 	      paramList.add(S_MONTH);
 	      paramList.add(bank_code);
 	      
 	      sqlCmd_sum.append(" select sum(amt) as a01970000 ");
 	      sqlCmd_sum.append(" from (select A01.amt,A01.acc_code ");
 	      sqlCmd_sum.append(" from A01 LEFT JOIN ncacno_7_rule ON A01.acc_code = ncacno_7_rule.acc_code  ");
 	      sqlCmd_sum.append(" where A01.m_year = ?");
 	      sqlCmd_sum.append(" and A01.m_month = ?");
 	      sqlCmd_sum.append(" and bank_code=?");
 	      sqlCmd_sum.append(" and (ncacno_7_rule.acc_div='08' and (ncacno_7_rule.acc_code <> '990000' and ncacno_7_rule.acc_code <> '120600'))");
 	      sqlCmd_sum.append(" order by acc_range )");
 	      sqlCmd_sum.append(" where acc_code not in ('120700','120900') ");  
	   }else{	   		
	      /*把120700+120900金額加總.其餘顯示原amt
	      select a01.acc_code,ncacno_7.acc_range,ncacno_7.acc_name,a01.amt from 
	      (select acc_code,amt  --not in ('120700','120900')
	      from a01
	      where acc_code not in ('120700','120900')
   	      and m_year=94 and m_month=6 and bank_code='5030019'
	      union
	      select '120700',sum(amt) amt from a01  --in('120700','120900') 
	      where acc_code in('120700','120900') 
	      and m_year=94 and m_month=6 and bank_code='5030019') a01 
	      ,ncacno_7 where a01.acc_code=ncacno_7.acc_code and (ncacno_7.acc_div='08' or ncacno_7.acc_code='990000')
	      order by ncacno_7.acc_range
	      */
	      sqlCmd.append(" select a01.acc_code,ncacno_7.acc_range,a01.amt ");
	      sqlCmd.append(" from  ( select acc_code,amt ");
	      sqlCmd.append("		  from a01 ");
	      sqlCmd.append("		  where acc_code not in ('120700','120900') ");
	      sqlCmd.append("		   and  m_year =?");
 	      sqlCmd.append(" 		   and  m_month = ?");
 	      sqlCmd.append(" 		   and  bank_code=?");
	      sqlCmd.append("         union ");
	      sqlCmd.append("		  select '120700',sum(amt) amt from a01 ");
	      sqlCmd.append("		  where acc_code in('120700','120900') ");
	      sqlCmd.append("		   and  m_year = ?");
 	      sqlCmd.append(" 		   and  m_month = ?");
 	      sqlCmd.append(" 		   and  bank_code=?");
	      sqlCmd.append("		) a01,ncacno_7 ");
	      sqlCmd.append(" where a01.acc_code=ncacno_7.acc_code and (ncacno_7.acc_div='08' or ncacno_7.acc_code='990000')");
	      sqlCmd.append(" order by ncacno_7.acc_range ");
	      paramList.add(S_YEAR);
 	      paramList.add(S_MONTH);
 	      paramList.add(bank_code);
 	      paramList.add(S_YEAR);
 	      paramList.add(S_MONTH);
 	      paramList.add(bank_code);
 	      
	      
	      sqlCmd_sum.append(" select sum(amt) as a01970000 ");
	      sqlCmd_sum.append(" from (select a01.acc_code,ncacno_7.acc_range,a01.amt ");
	      sqlCmd_sum.append(" from  ( select acc_code,amt ");
	      sqlCmd_sum.append("		  from a01 ");
	      sqlCmd_sum.append("		  where acc_code not in ('120700','120900') ");
	      sqlCmd_sum.append("		   and  m_year = ?");
 	      sqlCmd_sum.append(" 		   and  m_month = ?");
 	      sqlCmd_sum.append(" 		   and  bank_code=?");
	      sqlCmd_sum.append("         union ");
	      sqlCmd_sum.append("		  select '120700',sum(amt) amt from a01 ");
	      sqlCmd_sum.append("		  where acc_code in('120700','120900') ");
	      sqlCmd_sum.append("		   and  m_year = ?");
 	      sqlCmd_sum.append(" 		   and  m_month = ?");
 	      sqlCmd_sum.append(" 		   and  bank_code=?");
	      sqlCmd_sum.append("		) a01,ncacno_7 ");
	      sqlCmd_sum.append(" where a01.acc_code=ncacno_7.acc_code and (ncacno_7.acc_div='08' and ( ncacno_7.acc_code <> '990000' and ncacno_7.acc_code <> '120600'))");
	      sqlCmd_sum.append(" order by ncacno_7.acc_range ) ");	
	   }	  
	}
	data_div02 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");		  
	List data_sum = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),paramList,"a01970000");		  
	System.out.println("data_div01.size="+data_div01.size());	
	if(data_div02 != null){
	   System.out.println("data_div02.size="+data_div02.size());	
	}
	//取得WMFileEdit的權限
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A06.permission == null");
    }else{
       System.out.println("WMFileEdit_A06.permission.size ="+permission.size());
               
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

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0" <%if(!alertMsg.equals("")) out.print("onload=alert('"+alertMsg+"') ");%> >
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
<input type="hidden" name="act" value="">  
<table width="1010" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td width="1010"><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF" width="1010">
			<table width="1010" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1010"><table width="1010" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="380"><img src="images/banner_bg1.gif" width="380" height="17"></td>
                      <td width="250"><b> 
                        <center>
                          <b> 
                          <center>
                          
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A06")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b> </td>
                      <td width="380"><img src="images/banner_bg1.gif" width="380" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td width="1010"><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td width="1010"><table width="1010" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp?width=1009" flush="true" /></div> 
                    </tr>
                      <td width="1010"><table width=1009 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                          
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="70"> <div align=left>申報年月</div></td>
                            <td colspan=2 width="725">                            
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
        						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="button" value="載入上月資料"-->	
        						<%if(datahasLast != null && datahasLast.size() != 0){ //95.05.03若上月無資料,則不顯示上月資料匯入%>
        						     <input type="button" name="LastMonthDataBtn" value="上月資料匯入" onclick="javascript:doSubmit(this.document.forms[0],'getLastMonthData','A06','','');">        						
        						<%}%>
                            </td>
                          </tr>
                          </table>
                          
                          <table width=1009 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                                                                              
                          <tr bgcolor='e7e7e7'>
                           <td width="50" rowspan="2" bgcolor="#B1DEDC" class="sbody">項目
                           </td>
                           <td width="350" rowspan="2" bgcolor="#B1DEDC" class="sbody">項目名稱
                           </td>
                           <td width="474" bgcolor="#B1DEDC" class="sbody" colspan="5"><div align="center">逾放期數</div></td>
                           <td width="88" bgcolor="#B1DEDC" class="sbody" rowspan="2"><div align="center">逾放期數合計(A)+(B)+(C)+(D)+(E)</div></td>
                           <td width="100" bgcolor="#B1DEDC" class="sbody" rowspan="2">A01放款合計</td>
                           <!--95.10.18拿掉顯示逾放比率td width="59" bgcolor="#B1DEDC" class="sbody" rowspan="2">逾放比率%</td-->
                          </tr>
                          <tr bgcolor='e7e7e7'>
                           <td width="90" bgcolor="#B1DEDC" class="sbody" align="center">未滿3個月<br>(A)</td>
                           <td width="116" bgcolor="#B1DEDC" class="sbody" align="center"><div align="center">3個月~未滿6個月<br>(B)</div></td>
                           <td width="96" bgcolor="#B1DEDC" class="sbody" align="center"><div align="center">6個月~未滿1年<br>(C)</div></td>
                           <td width="90" bgcolor="#B1DEDC" class="sbody" align="center"><div align="center">1年~未滿2年<br>(D)</div></td>
                           <td width="91" bgcolor="#B1DEDC" class="sbody" align="center">2年以上<br>(E)</td>
                          </tr>                            
                          
                    
                          
                          <% int i = 0 ;							
							 String bgcolor="#F2F2F2";	
							 String tmpAcc_Code = "";	
							 
							 while( i < data_div01.size()){										       
		    					   tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  			
		    					   System.out.println("i="+i);
		    					   System.out.println(":"+tmpAcc_Code);		    					   		    					    
						  %>							  
						 
						  <tr bgcolor="<%if(tmpAcc_Code.equals("120601") || tmpAcc_Code.equals("120602") || tmpAcc_Code.equals("120603") || tmpAcc_Code.equals("120604")){
                                            out.print("#FFFFE6");//淺黃色
                                         }else if(tmpAcc_Code.equals("120600") || tmpAcc_Code.equals("970000")){
                                            out.print("#B1DEDC");
                                         }else{
                                            out.print("#D2F0FF");
                                         }%>" class="sbody">                          
                          <td width="50">
                             <div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
							 <input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
							 <input type=hidden name=acc_div value="08">			
						  </td>  
                          <td width="350"><div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>
                            <%if(tmpAcc_Code.equals("120700") && bank_type.equals("7")){
                            	if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) < 10301){ 
                                out.print("(內含120900)");     
                                }                        
                             }%>
                          </div></td>                          
                          
                          <td width="90">
                            <% //若金額為0時,顯示空白                            
						        if( ((DataObject)data_div01.get(i)).getValue("amt_3month") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_3month") != null && ((((DataObject)data_div01.get(i)).getValue("amt_3month")).toString()).equals("0")) ){%>
			                     <input align='right' type='text' name='amt_3month' value="" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_3month");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>   			   
			                 <% }else{%>    			                 
			       			     <input align='right' type='text' name='amt_3month' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_3month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_3month"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this);' onchange='' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_3month");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   
			       			 <% }%>   
                          </td>
                          
                          <td width="116">
                           <%if(((DataObject)data_div01.get(i)).getValue("amt_6month") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_6month") != null && ((((DataObject)data_div01.get(i)).getValue("amt_6month")).toString()).equals("0")) ){%>                             
			                     <input align='right' type='text' name='amt_6month' value="" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_6month");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>   			   
			                 <% }else{%>    
			       			     <input align='right' type='text' name='amt_6month' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_6month")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_6month"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onchange='' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_6month");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   
			       			 <% }
			       			%>   	
                          </td>
                          
                          <td width="96">
                          <%if(((DataObject)data_div01.get(i)).getValue("amt_1year") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_1year") != null && ((((DataObject)data_div01.get(i)).getValue("amt_1year")).toString()).equals("0")) ){%>
			                     <input align='right' type='text' name='amt_1year' value="" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_1year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>   			   
			                 <% }else{%>    
			       			     <input align='right' type='text' name='amt_1year' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_1year")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_1year"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onchange='' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_1year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   
			       			 <% }%>   
                          </td>
                          
                          <td width="90">
                          <%if(((DataObject)data_div01.get(i)).getValue("amt_2year") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_2year") != null && ((((DataObject)data_div01.get(i)).getValue("amt_2year")).toString()).equals("0")) ){%>
			                     <input align='right' type='text' name='amt_2year' value="<%if(!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500") || tmpAcc_Code.equals("970000"))) out.print("0");%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_2year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000") || (!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500")))) out.print(" color:#808080; background-color:#FFFFE6");%>'
			                     <%if(tmpAcc_Code.equals("970000") || (!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500")))) out.print("readonly");%>
			                     >   			   
			                 <% }else{%>    
			       			     <input align='right' type='text' name='amt_2year' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_2year")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_2year"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onchange='' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_2year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   
			       			 <% }%>  
                          </td>
                          
                          <td width="91">
                          <%if(((DataObject)data_div01.get(i)).getValue("amt_over2year") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_over2year") != null && ((((DataObject)data_div01.get(i)).getValue("amt_over2year")).toString()).equals("0")) ){%>
			                     <input align='right' type='text' name='amt_over2year' value="<%if(!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500") || tmpAcc_Code.equals("970000"))) out.print("0");%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_over2year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000") || (!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500")))) out.print(" color:#808080; background-color:#FFFFE6");%>'
			                     <%if(tmpAcc_Code.equals("970000") || (!(tmpAcc_Code.equals("150200") || tmpAcc_Code.equals("960500")))) out.print("readonly");%>
			                     >   			   
			                 <% }else{%>    
			       			     <input align='right' type='text' name='amt_over2year' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_over2year")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_over2year"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onchange=' ' onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_over2year");' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   
			       			 <% }%>   
                          </td>
                                                    
                          <td width="90">
                          <%if(((DataObject)data_div01.get(i)).getValue("amt_total") == null ||  (((DataObject)data_div01.get(i)).getValue("amt_total") != null && ((((DataObject)data_div01.get(i)).getValue("amt_total")).toString()).equals("0")) ){%>
			                     <input align='right' type='text' name='amt_total' value="" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onChange="A06_checkA01('<%=tmpAcc_Code%>','<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>',this.form,<%=i%>);" onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_total");A06_per(this.form,<%=i%>);' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>   			   
			                 <% }else{%>    
			       			     <input align='right' type='text' name='amt_total' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt_total")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt_total"))).toString())%>" size=14 maxlength=14 onFocus='this.value=changeVal(this)' onChange="A06_checkA01('<%=tmpAcc_Code%>','<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>',this.form,<%=i%>);" onBlur='checkPoint_focus(this);this.value=changeStr(this);A06_TOTAL(this.form,"amt_total");A06_per(this.form,<%=i%>);' style='text-align: right; <%if(tmpAcc_Code.equals("970000")) out.print(" color:#808080; background-color:#FFFFE6");%>' <%if(tmpAcc_Code.equals("970000")) out.print("readonly ");%>>			   			       			     
			       			 <% }%>   
                          </td>
                                                    
                          <td width="100" align="right">     
                                                                         
                          <%
                          if(i==16)/*970000*/{//95.05.16 add A01逾放合計.合計
                              System.out.println("i===16");
                              if(data_sum != null && data_sum.size() > 1){ 
                                 out.print(Utility.setCommaFormat(((((DataObject)data_sum.get(0)).getValue("a01970000")) == null ? "":(((DataObject)data_sum.get(0)).getValue("a01970000"))).toString()));			       			     
                                 out.print("<input type='hidden' name='A01' value="+(((DataObject)data_sum.get(0)).getValue("a01970000")).toString()+">");
                               }else{
			       		             out.print("&nbsp;<input type='hidden' name='A01' value='0'>");			       		       
			       		       }
                          }else if(i<16){
                              if(i == 15){/*i=15-->960500*/
                                 out.print("&nbsp;<input type='hidden' name='A01' value='0'>");			       		       
                              }else{//i!=15   
                                 System.out.println("i<16 && i!= 15");
                                 if( data_div02 != null && (data_div02.size() > 1  ) && (!(((DataObject)data_div02.get(i)).getValue("amt") == null ||  (((DataObject)data_div02.get(i)).getValue("amt") != null && ((((DataObject)data_div02.get(i)).getValue("amt")).toString()).equals("0")) ) ) ){
                                     out.print(Utility.setCommaFormat(((((DataObject)data_div02.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div02.get(i)).getValue("amt"))).toString()));			       			     
                                     out.print("<input type='hidden' name='A01' value="+(((DataObject)data_div02.get(i)).getValue("amt")).toString()+">");                                     
                                 }else{
			       		             out.print("&nbsp;<input type='hidden' name='A01' value='0'>");			       		       
			       		         } 
			       		      }//end of i!=15   
			       		  }//end of i<16%>                    
                          </td>
                                                    
                          <!--95.10.18拿掉顯示逾放比率td width="59">
                          <div id='div_A01per'></div>		
						  <input type='hidden' name='A01per' value="">
                          </td-->
                                                    
                        </tr>
						  
						<%	     i++;	
						     }
					    %>  
					    <tr bgcolor="#FFFFE6" class="sbody">                          
                          <td width="1132" colspan="10">
                             <p align="right">&nbsp;&nbsp;&nbsp;&nbsp; 
                             A01.990000金額為
                             <%//95.05.16 add a01.990000
                             if(data_div02 != null && data_div02.size() > 1 ){
                                System.out.println("a01.990000.data_div02 != null");
                               out.print(Utility.setCommaFormat(((((DataObject)data_div02.get(15)).getValue("amt")) == null ? "":(((DataObject)data_div02.get(15)).getValue("amt"))).toString()));			       			                                  
			       		     }else{
			       		       out.print("0");			       		       
			       		     }%> 
                             元&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						  </td>  
                                                    
                        </tr>

                        </table></td>
                    </tr>
                    <tr> 
                      <td width="892">&nbsp;</td>
                    </tr>
                  </table>
                       </div>
              </tr>
               <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp?width=1010" flush="true" /></div></td>              
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="1010" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
                       <td width="100">&nbsp;</td>
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A06','','','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A06','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A06','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
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
                        <td width="300">&nbsp;</td>
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
          <td bgcolor="#FFFFFF"><table width="1010" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="1000" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="1010" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A06")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【確定】</font>或<font color="#666666">【修改】</font>時,會一併執行線上檢核,需耗時5-7秒。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>             
            </table></td>
        </tr>      
      </table></td>
  </tr>
</table>
</form>
<script language="JavaScript" >
<!--
//A06_per_onload(this.document.forms[0]);//95.11.09 fix 拿掉計算逾放比率
-->
</script>
</body>

</html>
