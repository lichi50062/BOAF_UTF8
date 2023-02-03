<%
//94/11/02 first design by 4180
//95.05.24 fix 獨立成支票存款維護 by 2295
//95.05.26 fix 畫面格式 by 2295
//95.06.13 add FX007WA權限 by 2295
//96.01.15 fix 預設的申報年月 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
    List WLX07_M_CHECKBANK= (request.getAttribute("WLX07_M_CHECKBANK")==null)?null:(List)request.getAttribute("WLX07_M_CHECKBANK");
    List inidate= (request.getAttribute("WLX07_INI")==null)?null:(List)request.getAttribute("WLX07_INI");  //取得鎖定年月
    List lockdate= (request.getAttribute("WLX07_LOCK")==null)?null:(List)request.getAttribute("WLX07_LOCK");  //取得鎖定年季
   
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX007WA")==null ) ? new Properties() : (Properties)session.getAttribute("FX007WA");
    if (permission == null) {
        System.out.println("FX007WA_List.permission == null");
    }else {
        System.out.println("FX007WA_List.permission.size ="+permission.size());
    }
    int iniyear=0, inimonth=0;
    String lockyear="", lockmonth="";//鎖定年月
		
    if(inidate!=null){
      iniyear = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_year").toString());
      inimonth = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_month").toString());
    }

	Calendar now = Calendar.getInstance();
	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
    if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
   	   YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
   	   MONTH = "12";
    }else{    
   	   MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
    }
    
	String m_year="";
	String m_month="" ;
	
	String checkbank_cnt="";
	String checkbank_bal="";
	String checkbank_cnt_s="";
	String checkbank_bal_s="";
	String checkbank_cnt_n="";
	String checkbank_bal_n="";	
	String maintain_id="" ;
	String maintain_name="" ;
	String maintain_date="" ;
%>
<script language="javascript" src="js/FX007WA.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>支票存款辦理情形維護</title>
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
</head>
<body topmargin="0" leftmargin="15">

<%
if(lockdate!=null){
%>
	<script language="javascript" type="text/JavaScript">
	<%
		for(int k=0;k<lockdate.size();k++)	
		{
	 		lockyear=String.valueOf(((DataObject)lockdate.get(k)).getValue("m_year"));
    	    lockmonth=String.valueOf(((DataObject)lockdate.get(k)).getValue("m_quarter"));
	%>
		pushArray('<%=lockyear%>');
		pushArray('<%=lockmonth%>');
	<%}%>
	</script>
<%}%>
<form name="date" method="post">
<table width="770" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" height="400">
        <tr>
         	<td width="770" height="16"><img src="images/space_1.gif" width="12" height="12"></td>
        </tr>       
        <tr>        
          <td bgcolor="#FFFFFF" width="770" height="310">        
          <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" height="310">
              <tr>
                <td width="770" height="18">
                  <table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="275"><img src="images/banner_bg1.gif" width="275" height="17" align="left"></td>
                      <td width="240" align="center"><b><center><font color="#000000" size="4">支票存款辦理情形</font></center></b> </td>
                      <td width="275" align="right"><img src="images/banner_bg1.gif" width="275" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
                <td width="770 height="300">
                <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" height="10">                    
					<tr><div align="right"><jsp:include page="getLoginUser.jsp?width=770" flush="true" /></div></tr>                        
          			<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                   	<tr><td height="42" width="770">
          				<table border="1" cellspacing="1" bordercolor="#3A9D99" width="770" height="30" class="sbody" cellpadding="0">
            			<tr>            
                 		<td class="sbody" bgcolor="#D8EFEE" width="131" height="30"><p align="center">申報年月</p></td>
                 		<td class="sbody" bgcolor="#E7E7E7" width="635" height="30"  nowrap>                               
                     		<input type="text" name="S_YEAR" value="<%=YEAR%>"      
                     		size="3" maxlength="3" onblur="CheckYear(this)">年 第
                       		<select name="S_MONTH" size="1">
                    		<%                   
                    			for(int i=1;i<=12;i++){
                        			if(i == Integer.parseInt(MONTH)){
                           				out.print("<option value="+i+" selected>"+i+"</option>");
                        			}else{
                           				out.print("<option value="+i+">"+i+"</option>");
                           			}	
                    			}
                    		%>
                    		</select> 月
                    		<a href="javascript:newSubmit(this.document.forms[0],'new','<%=iniyear*12+inimonth%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                    		<img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                		</td>  		
            		</tr>
         			</table>
            		</td>
             		</tr>
                  	<%}%>
                    <tr class="sbody">
                      <td  class="sbody"   height="100" width="770">
                      <table class="sbody" width="770" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="20" >
                          <tr>
                            <td width="82" rowspan="2" bgcolor="#9AD3D0" align="center" >申報年月</td>
                            <td width="506" colspan="6" bgcolor="#9AD3D0" align="center">支票存款</td>
                            <td width="68" rowspan="2" bgcolor="#9AD3D0" align="center">
                              <p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">異動者</p>
                              <p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">帳號/姓名</td>
                            <td width="94" rowspan="2" bgcolor="#9AD3D0" align="center">異動日期</td>
                          </tr>
                          <tr>
                            <td width="166" colspan="2" bgcolor="#9AD3D0" align="center">正會員</td>
                            <td width="166" colspan="2" bgcolor="#9AD3D0" align="center">贊助會員</td>
                            <td width="166" colspan="2" bgcolor="#9AD3D0" align="center">非會員</td>
                          </tr>
                          <%if(WLX07_M_CHECKBANK.size()==0){%>
                          <tr>
                            <td width="762" align="center" colspan="9">尚無資料</td>
                          </tr>
                          <%}else{                          	  
                              for(int i=0;i<WLX07_M_CHECKBANK.size();i++){
                                  m_year = String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("m_year"));
                                  m_month =String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("m_month"));                                       
								  checkbank_cnt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_cnt")));
								  checkbank_bal=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_bal")));
								  checkbank_cnt_s=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_cnt_s")));
								  checkbank_bal_s=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_bal_s")));
								  checkbank_cnt_n=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_cnt_n")));
								  checkbank_bal_n=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("checkbank_bal_n")));
								  maintain_id=String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("user_id"));
								  maintain_name=String.valueOf(((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("user_name"));
								  maintain_date=Utility.getCHTdate((((DataObject)WLX07_M_CHECKBANK.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   
                          %>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>">
                          <td class="sbody" width="82" align="center">                            
                          <%boolean locked=false;
                            if(inidate!=null ){//有初始申報日期限制
                              if(iniyear*12+inimonth <= Integer.parseInt(m_year)*12+Integer.parseInt(m_month)){//在合法申報日期內                                 
                                 if(lockdate!=null ){
                                    for(int c=0;c<lockdate.size();c++){                                 	 	
                                 	 lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                 	 lockmonth=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                  	 if(m_year.equals(lockyear) && m_month.equals(lockmonth))locked=true;
                                    }
                                 }  
                                 if(locked==false){
                                   	out.print( "<u><a href=\"FX007WA.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
                                 }
                                 out.print(m_year+"/"+ m_month+"</a>");
                                 locked=false;                              
                              }else{
                                out.print(m_year+"/"+ m_month);
                              }
                            }else{
                              if(lockdate!=null ){
                             	  for(int c=0;c<lockdate.size();c++){     	 	
                                  	  lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                 	  lockmonth=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                  	  if(m_year.equals(lockyear) && m_month.equals(lockmonth))locked=true; 
                                  }
                              }         
                              if(locked==false){
                                 out.print( "<u><a href=\"FX007WA.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
                              }
                              out.print(m_year+"/"+ m_month+"</a>");
                              locked=false;
                            }
                          %></u></td>             
                          
                            <td class="sbody" width="81" align="right"><%=checkbank_cnt%></td>
                            <td class="sbody" width="81" align="right"><%=checkbank_bal%></td>
                            <td class="sbody" width="81" align="right"><%=checkbank_cnt_s%></td>
                            <td class="sbody" width="81" align="right"><%=checkbank_bal_s%></td>
                            <td class="sbody" width="81" align="right"><%=checkbank_cnt_n%></td>
                            <td class="sbody" width="81" align="right"><%=checkbank_bal_n%></td>
                            <td class="sbody" width="68" align="left"><%=maintain_id%>/<%=maintain_name%></td>
                            <td class="sbody" width="94" align="center"><%=maintain_date%></td>                                                                     
                           </tr>                 
                          <% }//end for
                            }//end else %>
                         </table>                     
                      </td>
                    </tr>


                    <td width="770" height="103">
						<div align="right"><jsp:include page="getMaintainUser.jsp?width=770" flush="true" /></div>                    		
                    </td>
      			</table>
      			</td>
  			  </tr>
  			  <tr>
                <td width="770" height="81">
                <table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody">
                </td>
                    <tr>
                      <td colspan="2" width="583"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明: </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td class="sbody" width="561">
                        <ul>
                          <li>輸入申報之年月，再點選<font color="#666666">【新增】按鈕</font>可新增該年月「支票存款」之資料。
                          <li>點選所列之[申報年月]可變更該申報之資料。
                          <li><font color="#ff0000"></font><font color="red">「支票存款」係以申報月底日的資料來申報</font></li>
                        </ul>
                      </td>
                    </tr>
                 </table>
                 </td>
              </tr>
          </table>
          </td>
       </tr>   
</table>
</form>
</html>
