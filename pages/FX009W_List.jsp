<%
//94.11.07 first design by 4180
//95.11.06 ADD 列印 by 2495
//96.01.15 fix 預設的申報年季,調整html code by 2295
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
    List WLX09_S_WARNING= (request.getAttribute("WLX09_S_WARNING")==null)?null:(List)request.getAttribute("WLX09_S_WARNING");
    List inidate= (request.getAttribute("WLX09_INI")==null)?null:(List)request.getAttribute("WLX09_INI");  //取得初始年季
    List lockdate= (request.getAttribute("WLX09_LOCK")==null)?null:(List)request.getAttribute("WLX09_LOCK");  //取得鎖定年季
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");   
    Properties permission = ( session.getAttribute("FX009W")==null ) ? new Properties() : (Properties)session.getAttribute("FX009W");
    if (permission == null) {
        System.out.println("FX009W_List.permission == null");
    }else {
        System.out.println("FX009W_List.permission.size ="+permission.size());
    }
    int iniyear=0, inimonth=0, iniquarter=0;//初始年季
	String lockyear="", lockquarter="";//鎖定年季		
    if(inidate!=null){//初始年月
       iniyear = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_year").toString());//初始年
       iniquarter= Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_month").toString());//初始季
    }
    //96.01.15 fix 預設的申報年季===================================================================================
    Calendar now = Calendar.getInstance();
    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;   	
   	String select1 ="";//選擇第一季
    String select2 ="";//選擇第一季
    String select3 ="";//選擇第一季
    String select4 ="";//選擇第一季
    if(Integer.parseInt(MONTH) <= 3){
      select4 ="selected";
      YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
    }else if(Integer.parseInt(MONTH) <= 6){
      select1 ="selected";
    }else if(Integer.parseInt(MONTH) <= 9){
      select2 ="selected";
    }else if(Integer.parseInt(MONTH) <= 12){
      select3 ="selected";	
    }                   		
    //===============================================================================================================
    //清單所使用的欄位==============================================================================================
	String m_year="";
	String m_quarter="";
	String warnaccount_tcnt="";
	String warnaccount_tbal="";
	String warnaccount_remit_tcnt="";
	String warnaccount_refund_apply_cnt="";
	String warnaccount_refund_apply_amt="";
	String warnaccount_refund_cnt="";
	String warnaccount_refund_amt="";
	String maintain_id="";
	String maintain_name="";
	String maintain_date="";
	

//out.print("iniyear"+iniyear+"iniquarter"+iniquarter);

%>
<script language="javascript" src="js/FX009W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>金融機構警示帳戶調查資料維護</title>
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
		for(int k=0;k<lockdate.size();k++){
	 		lockyear=String.valueOf(((DataObject)lockdate.get(k)).getValue("m_year"));
    		lockquarter=String.valueOf(((DataObject)lockdate.get(k)).getValue("m_quarter"));
	%>
		pushArray('<%=lockyear%>');
		pushArray('<%=lockquarter%>');
	<%}%>
	</script>
<%}%>
<table width="725" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
           <td width="725"><img src="images/space_1.gif" width="12" height="12"></td>
          </tr>       
        <tr>
        
          <td bgcolor="#FFFFFF" width="725">
        
          <table width="725" border="0" align="center" cellpadding="0" cellspacing="0" height="309">
              <tr>
                <td width="725" height="18">
                  <table width="725" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17" align="left"></td>
                      <td width="365" align="center"><b><center><font color="#000000" size="4">金融機構警示帳戶調查資料維護</font></center></b> </td>
                      <td width="200" align="right"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
                <td width="725" height="281">
                <table width="725" border="0" align="center" cellpadding="0" cellspacing="0" height="164">
                    <tr><td height="18" width="725" >  </td></tr>
                    <tr>
                <td width="725 height="300">
                <table width="725" border="0" align="center" cellpadding="0" cellspacing="0" height="10">                    
					<tr><div align="right"><jsp:include page="getLoginUser.jsp?width=725" flush="true" /></div></tr>                        
          			<% if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
             		<tr><td height="42" width="725">
          				<table border="1" cellspacing="1" bordercolor="#3A9D99" width="725" height="30" class="sbody" cellpadding="0">
            			<tr>
            <form name="date" method="post">
                <td class="sbody" bgcolor="#D8EFEE" width="131" height="1"><p align="center">申報年季</p></td>
                 <td class="sbody" bgcolor="#E7E7E7" width="635" height="1" valign="middle">                   
                     <%
                       
                     %>
                     <input type="text" name="S_YEAR" size="3" value="<%=YEAR%>">年 第                      
               		  <select type="text" name='S_QUARTER' size="1"> 
                		<option value="1" <%= select1 %> >01</option>
                		<option value="2" <%= select2 %> >02</option>
                		<option value="3" <%= select3 %> >03</option>
                		<option value="4" <%= select4 %> >04</option>
                	 </select>季  
                    <a href="javascript:newSubmit(this.document.forms[0],'new','<%=iniyear*12+iniquarter%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1);">
                    <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                   
					<%
                    	String session_bank_type=(String)session.getAttribute("bank_type");
                    	System.out.println("session_bank_type ="+session_bank_type);
                    	String request_bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
						System.out.println("request_bank_type ="+request_bank_type);
                    if(/*!muser_id.equals("A111111111")&&(bank_type.equals("6")||bank_type.equals("7"))*/true){%>
                     <a href="javascript:printSubmit(this.document.forms[0],'print','<%=session.getAttribute("bank_type")%>',<%=bank_no%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    <img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                		<%}%>                		
                </td>
  		</form>
        </tr>
         </table>
            </td>
             </tr>
                  <%}else{%>
                  	             <tr><td height="42" width="725">
          				<table border="1" cellspacing="1" bordercolor="#3A9D99" width="725" height="30" class="sbody" cellpadding="0">
            			<tr>
            <form name="date" method="post">
                <td class="sbody" bgcolor="#D8EFEE" width="131" height="1"><p align="center">申報年季</p></td>
                 <td class="sbody" bgcolor="#E7E7E7" width="635" height="1" valign="middle">                                        
                     <input type="text" name="S_YEAR" size="3" value="<%=YEAR%>">年 第                     
               		 <select type="text" name='S_QUARTER' size="1"> 
                		<option value="1" <%= select1 %> >01</option>
                		<option value="2" <%= select2 %> >02</option>
                		<option value="3" <%= select3 %> >03</option>
                		<option value="4" <%= select4 %> >04</option>
                	 </select>季  
                    <%
                    //String bank_type=(String)session.getAttribute("bank_type");
                    String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
										
                     if(/*!muser_id.equals("A111111111")&&(bank_type.equals("6")||bank_type.equals("7"))*/true){%>
                     <a href="javascript:printSubmit(this.document.forms[0],'print','<%=bank_type%>',<%=bank_no%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    <img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                		<%}%>
                </td>
  		</form>
        </tr>
         </table>
            </td>
             </tr>
									<%}%>
                    <tr class="sbody">
                      <td  class="sbody"   height="126" width="725">
                      <div align="right">
                      <table class="sbody" width="725" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="88" >
                          <tr  bgcolor="#9AD3D0">
                          <td  width="32" rowspan="3" align="center" height="44"><p style="margin-top: 0; margin-bottom: 0">申報年季</p></td>
                          <td colspan="3"  align="center" height="8" width="242"><p style="margin-top: 0; margin-bottom: 0">警示帳戶(91.11.1~該季末日)</p></td>
                          <td  colspan="4"  align="center" height="8" width="326"><p style="margin-top: 0; margin-bottom: 0">警示帳戶內剩餘款項之返還情形(91.11.1~該季末日)</p></td>
                          <td rowspan="3" width="70" align="center" height="44">
                          <p style="margin-top: 0; margin-bottom: 0">異動者</p>
                          <p style="margin-top: 0; margin-bottom: 0">帳號/</p>
                          <p style="margin-top: 0; margin-bottom: 0">姓名</p>
                          </td>
                          <td rowspan="3" width="55" align="center" height="44">異動日期</td>
                          </tr>

                          <tr  bgcolor="#9AD3D0">
                          <td  rowspan="2" align="center" height="9" width="79"><p style="margin-top: 0; margin-bottom: 0">總戶數</p></td>
                          <td  rowspan="2" align="center" height="9" width="84"><p style="margin-top: 0; margin-bottom: 0">總餘額</p></td>
                          <td rowspan="2" align="left" height="9" width="79"><p style="margin-top: -1; margin-bottom: -3">警示帳戶內所匯（轉）入總筆數</p></td>
                          <td  colspan="2"  align="center" height="9" width="163"><p style="margin-top: 0; margin-bottom: 0">申請退還</p></td>
                          <td  colspan="2"  align="center" height="9" width="163"><p style="margin-top: 0; margin-bottom: 0">巳辦理退還</p></td>
                          </tr>
                          <tr class="sbody" bgcolor="#9AD3D0">
                          <td  width="79" align="center" height="19"><p style="margin-top: 0; margin-bottom: 0">戶數</p></td>
                          <td  width="84" align="center" height="19"><p style="margin-top: 0; margin-bottom: 0">金額</p></td>
                          <td width="79" class="sbody"   align="center" height="19"><p style="margin-top: 0; margin-bottom: 0">戶數</p></td>
                          <td width="84" class="sbody"   align="center" height="19"><p style="margin-top: 0; margin-bottom: 0">金額</p></td>
                    
                          </tr>

                          <%if(WLX09_S_WARNING.size()==0){%>
                            <tr class="sbody" bgcolor="#D8EFEE"><td colspan="13" align="center"  height="16" width="725" >
                              <font   class="sbody">尚無資料</font></td></tr>
                           <%}else{
							for(int i=0;i<WLX09_S_WARNING.size();i++){
								m_year = String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("m_year"));
								m_quarter = String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("m_quarter"));
								warnaccount_tcnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_tcnt")));
								warnaccount_tbal =Utility.setCommaFormat( String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_tbal")));
								warnaccount_remit_tcnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_remit_tcnt")));
								warnaccount_refund_apply_cnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_refund_apply_cnt")));
								warnaccount_refund_apply_amt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_refund_apply_amt")));
								warnaccount_refund_cnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_refund_cnt")));
								warnaccount_refund_amt =Utility.setCommaFormat( String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("warnaccount_refund_amt")));
								maintain_id = String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("user_id"));
								maintain_name = String.valueOf(((DataObject)WLX09_S_WARNING.get(i)).getValue("user_name"));
								maintain_date=Utility.getCHTdate((((DataObject)WLX09_S_WARNING.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   
                         %>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>">
                          <td class="sbody"  align=right height="16" width="30">
                          <%
                            boolean locked=false;
                            if(inidate!=null ){//有初始申報日期限制
                              if(iniyear*12+iniquarter <= Integer.parseInt(m_year)*12+Integer.parseInt(m_quarter)){//在合法申報日期內
                                 if(lockdate!=null ){
                                 	for(int c=0;c<lockdate.size();c++){                                	 	
                                 		lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                 	  	lockquarter=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                  		if(m_year.equals(lockyear) && m_quarter.equals(lockquarter))locked=true;
                                  	}
                                  }                                     
                                  if(locked==false){
                                  	   out.print( "<u><a href=\"FX009W.jsp?act=Edit&myear="+ m_year+"&mquarter="+m_quarter+"\">");
                                  }
                                  out.print(m_year+"/"+ m_quarter+"</a>");
                                  locked=false;                                  
                              }else{
                                out.print(m_year+"/"+ m_quarter);
                              }
                           }else{
                             if(lockdate!=null ){
                            	for(int c=0;c<lockdate.size();c++){   	 	
                                	lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                 	lockquarter=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                  	if(m_year.equals(lockyear) && m_quarter.equals(lockquarter))locked=true; 
                                }
                             }
                                     
                             if(locked==false){
                                 out.print( "<u><a href=\"FX009W.jsp?act=Edit&myear="+ m_year+"&mquarter="+m_quarter+"\">");
                             }
                             out.print(m_year+"/"+ m_quarter+"</a>");
                             locked=false;
                           }
                          %>

                            </u></td>
                            <td class="sbody"  align=right height="16" width="79"><%=warnaccount_tcnt%></td>
                            <td class="sbody"  align=right height="16" width="84"><%=warnaccount_tbal%></td>      
                            <td class="sbody"  align=right height="16" width="79"><%=warnaccount_remit_tcnt%></td>
                            <td class="sbody"  align=right height="16" width="79"><%=warnaccount_refund_apply_cnt%></td>
                            <td class="sbody"  align=right height="16" width="84"><%=warnaccount_refund_apply_amt%></td>
                            <td class="sbody"  align=right height="16" width="79"><%=warnaccount_refund_cnt%></td>
                            <td class="sbody" align=right height="16" width="84"><%=warnaccount_refund_amt%></td>
				 			<td class="sbody"  align=left height="16" width="70"><%=maintain_id%> / <%=maintain_name%></td>
                            <td class="sbody" align=center height="16" width="55"><%=maintain_date%></td>
                            </tr>
                          <% }//end for
                            }//end else %>
                         </table>
                      </div>
                      </td>
                    </tr>


                    <td width="725" height="56"><div align="right"><div align="right"><div align="right"><jsp:include page="getMaintainUser.jsp?width=725" flush="true" /></div>
                    </div>
                     <p align="left">　</div></td>
      </table></td>
  </tr>
  <tr>
                <td width="825" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr>
                      <td colspan="2" width="583"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明                        
                        : </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td class="sbody" width="561">
                        <ul>
                          <li>輸入申報之年季，再點選<font color="#666666">【新增】按鈕</font>可新增該年季「警示帳戶調查」之資料。
                          <li>點選所列之[申報年季]可變更該申報之資料。
                          <li>本表係按最近的[申報年季]先排序,依此類推。
                          <li><font color="#ff0000">如果在[申報年季]欄位沒有出現底線,表示巳辦理[鎖定],僅提供查詢不可再異動</font></li>
                        </ul>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>
           
</table>
</form>
</table>
</html>
