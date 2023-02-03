<%
//94.10.17 first design by 4180
//96.11.28 add 金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元) by 2295
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
<%
    List WLX05_M_ATM= (request.getAttribute("WLX05_M_ATM")==null)?null:(List)request.getAttribute("WLX05_M_ATM");
    List inidate= (request.getAttribute("WLX05_INI")==null)?null:(List)request.getAttribute("WLX05_INI");  //取得鎖定年月
    List lockdate= (request.getAttribute("WLX05_LOCK")==null)?null:(List)request.getAttribute("WLX05_LOCK");  //取得鎖定年季
   
    
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX005W")==null ) ? new Properties() : (Properties)session.getAttribute("FX005W");
    if (permission == null) {
        System.out.println("FX005W_List.permission == null");
    }
    else {
        System.out.println("FX005W_List.permission.size ="+permission.size());
    }
      String lockyear="", lockmonth="";//鎖定年月
      int iniyear=0, inimonth=0;
      
      if(inidate!=null){
      iniyear = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_year").toString());
      inimonth = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_month").toString());
    }

    String m_year="";
    String m_month="" ;
    String pdebitcard="" ;
    String udebitcard="";
    String cancdebitcard="";
    String pbincard="";
    String ubincard="" ;
    String cancbincard="";
    String atmcnt="";
    String mtrancnt="";
    String ytrancnt="";
    String monthamt="";
    String yearamt="";                        
    String maintain_id="" ;
    String maintain_name="" ;
    String maintain_date="" ;
    String debitcard_mtrancnt="";      
	String debitcard_ytrancnt="";      
	String debitcard_monthamt="";      
	String debitcard_yearamt="";       

%>
<script language="javascript" src="js/FX005W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>金融卡發卡及ATM裝設情形維護</title>
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
<table width="830" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
         <td width="830"><img src="images/space_1.gif" width="12" height="12"></td>
         </tr>              
       <tr>
          <td bgcolor="#FFFFFF" width="830">
        
              <tr>
                <td width="830" height="18">
                 <table width="830" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17" align="left"></td>
                      <td width="400" align="center"><b><font size="4">金融卡發卡及ATM裝設情形</font></b></td>
                      <td width="200" align="right"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
                  <tr>
              <td width="830" height="50">
                <div align="left">
                <table width="830" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="725" >               
                      </td>
                    </tr>
                    <tr>
                       
                       <div>
                       <jsp:include page="getLoginUser.jsp?width=830" flush="true" />             			
                       </div>			 							
			 		                
                    </tr>
                    
        <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr ><td height="1" width="831">

          <table border="1" cellspacing="1" bordercolor="#3A9D99" width="100%" height="1" class="sbody" cellpadding="0">
            <tr>
            <form name="date" method="post">
                <td class="sbody" bgcolor="#D8EFEE" width="131" height="1"><p align="center">申報年月</p></td>
                 <td class="sbody" bgcolor="#E7E7E7" width="635" height="1" valign="middle">
                    
                     <input type="text" name="S_YEAR" value="<% 
                      Date today = new Date();
                     out.print(today.getYear()-11);
                     %>"      
                     size="3" maxlength="3" onblur="CheckYear(this)">年 第
                       <select name="S_MONTH" size="1">
                    <%
                   
                    for(int i=1;i<=12;i++)
                    {
                   if(i == today.getMonth())
                        out.print("<option value="+i+" selected>"+i+"</option>");
                   else
                       out.print("<option value="+i+">"+i+"</option>");


                    }
                    %>

                        </select> 月
                    <a href="javascript:newSubmit(this.document.forms[0],'new','<%=iniyear*12+inimonth%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                        <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>

               
                </td>
  </form>

            </tr>

         </table>
                      </td>
                    </tr>
                  <%}%>

                    <tr class="sbody">
                      <td  class="sbody"   height="126" width="831">
                      <div align="right">
                      <table class="sbody" width="100%" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="1" >
                          <tr bgcolor="#9AD3D0">
                            <td width="31" rowspan="3" align="center" height="1">
                              <p style="margin-top: -2; margin-bottom: -2">申報年月</p>
                            </td>
                            <td colspan="7"  align="center" height="1" width="306"> 
                              <p style="margin-top: -2; margin-bottom: -2">金融卡</p>
                            </td>
                            <td colspan="5"  align="center" height="1" width="269">
                              <p style="margin-top: -2; margin-bottom: -2">ATM自動化機器</p>
                            </td>
                            <td rowspan="3" align="center" height="1" width="70">
                              <p style="margin-top: -2; margin-bottom: -2" align="left">異動者帳號/姓名</p>
                            </td>
                            <td rowspan="3" align="center" height="1" width="53">
                              <p style="margin-top: -2; margin-bottom: -2">異動</p>
                              <p style="margin-top: -2; margin-bottom: -2">日期</p>
                            </td>
                          </tr>

                          <tr class="sbody" bgcolor="#9AD3D0">
                            <td width="123" align="center" height="1" colspan="3">歷年發卡情形</td>
                            <td width="82"   align="center" height="1" colspan="2">交易次數</td>
                            <td width="77"  align="center" height="1" colspan="2">交易金額</td>
                            <td align="center" height="2" width="50" rowspan="2">
                            	<p style="margin-top: -2; margin-bottom: -2">裝設</p>
                            	<p style="margin-top: -2; margin-bottom: -2">台數</p>
                            </td>
                            <td width="82" align="center" height="1" colspan="2">交易次數</td>
                            <td width="77"  align="center" height="1" colspan="2">交易金額</td>                          
                          </tr>
                          
                          <tr class="sbody" bgcolor="#9AD3D0">
                            <td  width="41" align="center" height="1">
                              <p style="margin-top: -2; margin-bottom: -2">發行</p>
                              <p style="margin-top: -2; margin-bottom: -2">張數</p>
                            </td>
                            <td width="41" class="sbody"   align="center" height="1">
                              <p style="margin-top: -2; margin-bottom: -2" >停卡</p>
                              <p style="margin-top: -2; margin-bottom: -2">張數</p>                         
                            </td>
                            <td width="41"   align="center" height="1">
                              <p style="margin-top: -2; margin-bottom: -2">流通</p>
                              <p style="margin-top: -2; margin-bottom: -2">張數</p>
                            </td>
                            <td align="center" height="1" width="50">
                              <p style="margin-top: -2; margin-bottom: -2">本月</p>
                            </td>
                            <td align="center" height="1" width="51">
                              <p style="margin-top: -2; margin-bottom: -2">本年</p>
                              <p style="margin-top: -2; margin-bottom: -2">累計</p>
                            </td>
                            <td align="center" height="1" width="76">
                              <p style="word-spacing: 0; margin-top: -2; margin-bottom: -2">本月</p>
                            </td>
                            <td align="center" height="1" width="81">
                              <p style="word-spacing: 0; margin-top: -2; margin-bottom: -2">本年累計</p>
                            </td>                          
                            <td align="center" height="1" width="50">
                              <p style="margin-top: -2; margin-bottom: -2">本月</p>
                            </td>
                            <td align="center" height="1" width="51">
                              <p style="margin-top: -2; margin-bottom: -2">本年</p>
                              <p style="margin-top: -2; margin-bottom: -2">累計</p>
                            </td>
                            <td align="center" height="1" width="76">
                              <p style="word-spacing: 0; margin-top: -2; margin-bottom: -2">本月</p>
                            </td>
                            <td align="center" height="1" width="81">
                              <p style="word-spacing: 0; margin-top: -2; margin-bottom: -2">本年累計</p>
                            </td>                          
                          </tr>

                          <%
                          if(WLX05_M_ATM.size()==0)
                          {
                          %>
                            <tr class="sbody" bgcolor="#D8EFEE"><td colspan="15" align="center"  height="19" width="768" >
                              <font   class="sbody">尚無資料&nbsp;</font></td></tr>
                           <%
                           }else{
							     boolean locked=false;
                                 for(int i=0;i<WLX05_M_ATM.size();i++){
                                     m_year = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("m_year"));
                                     m_month =String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("m_month"));
                                     pdebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("push_debitcard_cnt"));
                                     cancdebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("canc_debitcard_cnt"));
                                     udebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("use_debitcard_cnt"));                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
                                     pbincard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("push_bincard_cnt"));
                                     cancbincard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("canc_bincard_cnt"));
                                     ubincard = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("use_bincard_cnt"));
                                     atmcnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("atm_cnt")));
                                     mtrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("month_tran_cnt")));
                                     ytrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("year_acctran_cnt")));
                                     monthamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("month_tran_amt")));
                                     yearamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("year_acctran_amt")));
                                     //96.11.23 add 金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元)
                                     debitcard_mtrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("debitcard_month_tran_cnt")));
                                     debitcard_ytrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("debitcard_year_acctran_cnt")));
                                     debitcard_monthamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("debitcard_month_tran_amt")));
                                     debitcard_yearamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("debitcard_year_acctran_amt")));
 					 				 maintain_id = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("user_id"));
	      				 			 maintain_name = String.valueOf(((DataObject)WLX05_M_ATM.get(i)).getValue("user_name"));
	     				 			 maintain_date = Utility.getCHTdate((((DataObject)WLX05_M_ATM.get(i)).getValue("update_date")).toString().substring(0, 10), 0);
                              %>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>">
                          <td class="sbody"  align=right height="19" width="31">
                             <%
                            	     locked=false;
                            		 if(inidate!=null ){//有初始申報日期限制
                              		    if(iniyear*12+inimonth <= Integer.parseInt(m_year)*12+Integer.parseInt(m_month)){//在合法申報日期內                                 
                                			if(lockdate!=null ){
                                			   for(int c=0;c<lockdate.size();c++){                                 	 	
                                 	               lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                                   lockmonth=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                				   if(m_year.equals(lockyear) && m_month.equals(lockmonth))locked=true;
                                			   }//end of lockdate
                                			}//end of lockdate!=null  
                                			if(locked==false){
                                			   out.print( "<u><a href=\"FX005W.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
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
                                		   }//end of lockdate
                            			}//end of  lockdate!=null        
                                		if(locked==false){
                                			out.print( "<u><a href=\"FX005W.jsp?act=Edit&myear="+ m_year+"&mmonth="+m_month+"\">");
                                		}
                                		out.print(m_year+"/"+ m_month+"</a>");
                                		locked=false;
                            		 }
                          %>
                            </u></td>
                            <td class="sbody"  align=right height="19" width="41">
                              <%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(pdebitcard)+Integer.parseInt(pbincard)))%>
                            </td>
                            <td class="sbody"  align=right height="19" width="41">
                              <%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(cancdebitcard)+Integer.parseInt(cancbincard)))%>
                            </td>
                            <td class="sbody" align=right height="19" width="41">
                              <%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(udebitcard)+Integer.parseInt(ubincard)))%>
                            </td>
                            <td class="sbody"  align=right height="19" width="50">
                              <%=debitcard_mtrancnt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="51">
                              <%=debitcard_ytrancnt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="78">
                              <%=debitcard_monthamt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="80">
                              <%=debitcard_yearamt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="50">
                              <%=atmcnt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="50">
                              <%=mtrancnt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="51">
                              <%=ytrancnt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="78">
                              <%=monthamt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="80">
                              <%=yearamt%>
                            </td>
                            <td class="sbody"  align=right height="19" width="70">
                			  <p align="left"><%=maintain_id%>/<%=maintain_name%></p>
                            </td>
                            <td class="sbody"  align=right height="19" width="53">
                    		  <p align="center"><%=maintain_date%></p>
                            </td>
                          </tr>
                          <%     }//end for
                           }//end of WLX05_M_ATM have data%>
                         </table>
                      </div>
                      </td>
                    </tr>
<tr>
                    <td width="831" height="56">
<div align="right"><jsp:include page="getMaintainUser.jsp?width=830" flush="true" /></div>           
    </td>
    </tr>
      </table></td>
  </tr>
  <tr>
                <td width="825" height="123">
                <table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr>
                      <td colspan="2" width="583"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明                   
                        : </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td class="sbody" width="561">
                        <ul>
                          <li class="sbody">輸入申報之年月，再點選<font color="#666666">【新增】</font>按鈕可新增該年月「金融卡發卡及ATM裝設情形」之資料。</li>
                          <li class="sbody">點選所列之[申報年月]可變更該申報之資料。</li>
                          <li class="sbody">本表係按最近的[申報年月]先排序,依此類推。</li>
                          <li class="sbody">如果在[申報年月]欄位沒有出現底線,表示巳辦理[鎖定],僅提供查詢不可再異動。</li>
                           <li class="sbody"><font color="red">歷年磁條金融卡與晶片卡之發行張數及停卡張數與流通張數係表示截至申報月底日之歷年
                           	 總張數。</font></li>
                        	 <li class="sbody" ><font color="red">流通張數欄位定義為發行張數與停卡張數相減。</font></li>
                        	 <li class="sbody"><font color="red">交易次數與金額係以每年作一累計；故每年元月時，本月交易與本年累計申報的數字應相同，
                        	                        其他2-12月份的本年累計數應為「本月交易數+上月的本年累計交易數」</font></li>
                        </ul>
                          </font>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</table>
</html>
