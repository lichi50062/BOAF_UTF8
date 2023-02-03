<%
// 94.11.02 first design by 4180
// 95.05.24 fix 獨立成統一農(漁)貸資料辦理情形維護 by 2495
// 95.06.13 add FX007WB權限 by 2295
// 96.01.15 fix 預設的申報年月,調整html code by 2295
//102.12.30 add 本月平均利率/本年平均利率 by 2295
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
    List WLX07_M_Credit= (request.getAttribute("WLX07_M_Credit")==null)?null:(List)request.getAttribute("WLX07_M_Credit");
    List inidate= (request.getAttribute("WLX07_INI")==null)?null:(List)request.getAttribute("WLX07_INI");  //取得鎖定年月
    List lockdate= (request.getAttribute("WLX07_LOCK")==null)?null:(List)request.getAttribute("WLX07_LOCK");  //取得鎖定年季
   
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX007WB")==null ) ? new Properties() : (Properties)session.getAttribute("FX007WB");
    if (permission == null) {
        System.out.println("FX007WB_List.permission == null");
    }else {
        System.out.println("FX007WB_List.permission.size ="+permission.size());
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
    //清單資料
	String m_year="";
	String m_month="" ;
	String creditmonth_cnt="";
	String creditmonth_amt="";
	String credityear_cnt_acc="";
	String credityear_amt_acc="";
	String credit_cnt="";
	String credit_bal="";
	String overcreditmonth_cnt="";
	String overcreditmonth_amt="";
	String overcredit_cnt="";
	String overcredit_bal="";
	String maintain_id="" ;
	String maintain_name="" ;
	String maintain_date="" ;
	String creditmonth_avgrate="";
	String credityear_avgrate="";
%>
<script language="javascript" src="js/FX007WB.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>統一農(漁)貸資料辦理情形維護</title>
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
                      <td width="270"><img src="images/banner_bg1.gif" width="260" height="17" align="left"></td>
                      <td width="240" align="center"><b><center><font color="#000000" size="4">統一農(漁)貸資料辦理情形</font></center></b> </td>
                      <td width="270" align="right"><img src="images/banner_bg1.gif" width="260" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>              
              <tr>
                <td width="770 height="300">
                <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" height="10">                    
					<tr><div align="right"><jsp:include page="getLoginUser.jsp?width=770" flush="true" /></div></tr>                        
          			<%      if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
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
                      <td  class="sbody"   height="100" width="770"><div align="center">
                      <table class="sbody" width="770" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="20" >
                          <tr  bgcolor="#9AD3D0">
                          <td  width="31" rowspan="4" align="center" height="44">
                          <p style="margin-top: 0; margin-bottom: 0">申報年月</p></td>
                          <td  colspan="12"  align="center" height="8" width="522"><p style="margin-top: 0; margin-bottom: 0">統一農(漁)貸資料</p></td>
                          <td rowspan="4" width="71" align="center" height="44">
                          <p style="margin-top: 0; margin-bottom: 0">異動者</p>
                          <p style="margin-top: 0; margin-bottom: 0">帳號/</p>
                          <p style="margin-top: 0; margin-bottom: 0">姓名</p>
                          </td>
                          <td rowspan="4" width="52" align="center" height="44">異動日期</td>
                          </tr>

                          <tr  bgcolor="#9AD3D0">
                          <td  colspan="8"  align="center" height="9" width="302"><p style="margin-top: 0; margin-bottom: 0">貸放資料</p>
                          </td>
                          <td  colspan="4"  align="center" height="9" width="220"><p style="margin-top: 0; margin-bottom: 0">逾放資料</p>
                          </td>
                          </tr>
                          <tr class="sbody" bgcolor="#9AD3D0">
                          <td  align="center" height="19" width="161"  colspan="3"><p style="margin-top: -2; margin-bottom: -2">本月新增</p></td>
                          <td  align="center" height="19" width="152"  colspan="3"><p style="margin-top: -2; margin-bottom: -2">本年累計</p></td>
                          <td  align="center" height="10" width="111" colspan="2"><p style="margin-top: -2; margin-bottom: -2">貸放餘額</p></td>
                          <td  align="center" height="19" width="70" colspan="2"><p style="margin-top: -2; margin-bottom: -2">本月新增</p></td>
                          <td  align="center" height="19" width="80" colspan="2"><p style="margin-top: -2; margin-bottom: -2">逾放餘額</p></td>
                          </tr>
						  <tr class="sbody" bgcolor="#9AD3D0">
						  <td align="center" height="9" width="47">戶數</td>
						  <td align="center" height="9" width="51">金額</td>
						  <td align="center" height="9" width="58">平均利率</td>
						  <td align="center" height="9" width="48">戶數</td>
					      <td align="center" height="9" width="45">金額</td>
					      <td align="center" height="9" width="58">平均利率</td>
                          <td align="center" height="9" width="57">戶數</td>
                          <td align="center" height="9" width="54">餘額</td>
                          <td align="center" height="9" width="34">戶數</td>
						  <td align="center" height="9" width="36">金額</td>
					      <td align="center" height="9" width="35">戶數</td>
						  <td align="center" height="9" width="45">餘額</td>
						  </tr>
                          
                          <%if(WLX07_M_Credit.size()==0){%>
                          <tr class="sbody" bgcolor="#D8EFEE">
                          	 <td colspan="15" align="center"  height="16" width="900" colspan="15" >尚無資料</td>
                          </tr>
                          <%}else{                          	  
                              for(int i=0;i<WLX07_M_Credit.size();i++){
                                  m_year = String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("m_year"));                                  
                                  System.out.println("m_year="+m_year);
                                  m_month =String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("m_month"));                                       
								  System.out.println("m_month="+m_month);
								  creditmonth_cnt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("creditmonth_cnt")));
								  System.out.println("creditmonth_cnt="+creditmonth_cnt);
								  creditmonth_amt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("creditmonth_amt")));
								  System.out.println("creditmonth_amt="+creditmonth_amt);
								  credityear_cnt_acc=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("credityear_cnt_acc")));
								  System.out.println("credityear_cnt_acc="+credityear_cnt_acc);
								  credityear_amt_acc=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("credityear_amt_acc")));
								  System.out.println("credityear_amt_acc="+credityear_amt_acc);
								  credit_cnt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("credit_cnt")));
								  System.out.println("credit_cnt="+credit_cnt);
								  credit_bal=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("credit_bal")));
								  System.out.println("credit_bal="+credit_bal);
								  overcreditmonth_cnt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("overcreditmonth_cnt")));
								  System.out.println("overcreditmonth_cnt="+overcreditmonth_cnt);
								  overcreditmonth_amt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("overcreditmonth_amt")));
								  System.out.println("overcreditmonth_amt="+overcreditmonth_amt);
								  overcredit_cnt=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("overcredit_cnt")));
								  System.out.println("overcredit_cnt="+overcredit_cnt);
								  overcredit_bal=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("overcredit_bal")));
								  System.out.println("overcredit_bal="+overcredit_bal);
								  creditmonth_avgrate=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("creditmonth_avgrate")));
								  System.out.println("creditmonth_avgrate="+creditmonth_avgrate);
								  credityear_avgrate=Utility.setCommaFormat(String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("credityear_avgrate")));
								  System.out.println("credityear_avgrate="+credityear_avgrate);
								  maintain_id=String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("user_id"));
								  maintain_name=String.valueOf(((DataObject)WLX07_M_Credit.get(i)).getValue("user_name"));
								  maintain_date=Utility.getCHTdate((((DataObject)WLX07_M_Credit.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   
                          %>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>">
                          <td class="sbody"  align=right height="16" width="31">                            
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
                                   	out.print( "<u><a href=\"FX007WB.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
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
                                 out.print( "<u><a href=\"FX007WB.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
                              }
                              out.print(m_year+"/"+ m_month+"</a>");
                              locked=false;
                            }
                          %></u></td>                          
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=creditmonth_cnt%></p></td>
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=creditmonth_amt%></p></td> 
                            <td class="sbody"  align=right height="16" width="58"><p align="right"><%=creditmonth_avgrate%></p></td>
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=credityear_cnt_acc%></p></td>     
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=credityear_amt_acc%></p></td>  
                            <td class="sbody"  align=right height="16" width="58"><p align="right"><%=credityear_avgrate%></p></td>    
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=credit_cnt%></p></td>
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=credit_bal%></p></td>
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=overcreditmonth_cnt%></p></td>
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=overcreditmonth_amt%></p></td>      
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=overcredit_cnt%></p></td>     
                            <td class="sbody"  align=right height="16" width="50"><p align="right"><%=overcredit_bal%></p></td>
				 										<td class="sbody"  align=left height="16" width="75"><%=maintain_id%>/<%=maintain_name%></td>
                            <td class="sbody"  align=center height="16" width="52"><%=maintain_date%></td>                            
                           </tr>                 
                          <% }//end for
                            }//end else %>
                         </table>
                      </div>
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
                          <li>輸入申報之年月，再點選<font color="#666666">【新增】按鈕</font>可新增該年月「統一農(漁)貸」之資料。
                          <li>點選所列之[申報年月]可變更該申報之資料。
                          <li><font color="#ff0000">[</font><font color="red">統一農(漁)貸」係以申報月底日的資料來申報</font></li>
                          <li><font color="red">「本年累計貸放戶數」、「本年累計貸放金額」<br>
                            (1)「本年累計貸放戶數」、「本年累計貸放金額」係以每年作一累計；故每年元月時，「本月新增  
                            </font>
                            </font>
                          </font>
                            </font><font color="red">
                            貸放戶數」與「本年累計貸放戶數」、「本月新增貸放金額」與「本年累計貸放金額」的數字應相同<br> 
                            (2) 其他2-12月份的「本年累計貸放戶數」應為「本月新增貸放戶數+上月的本年累計貸放戶數」「本年累計貸放金額」應為「本月新增貸放金額+上月的本年累計貸放金額」<br>
                            (3) 起始申報年月時，由農(漁)會信用部負責輸入，後續則由系統提供自動累加功能。<br> 
                            (4) 本月平均利率=本月貸放案件核定利率之合計數/本月貸放戶數。<br> 
                            (5) 平均利率=本年貸放案件核定利率之合計數/本年貸放戶數。
                            </li>   
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
