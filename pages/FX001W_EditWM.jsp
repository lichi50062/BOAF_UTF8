<%
//93.12.20 add 權限
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
	
	Calendar now = Calendar.getInstance();
   	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
   	if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
       YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
       MONTH = "12";
    }else{    
      MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
    }
	List WLX01_WM = (List)request.getAttribute("WLX01_WM");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");					
	
	
	//取得FX001W的權限
	Properties permission = ( session.getAttribute("FX001W")==null ) ? new Properties() : (Properties)session.getAttribute("FX001W"); 
	if(permission == null){
       System.out.println("FX001W_Edit.permission == null");
    }else{
       System.out.println("FX001W_Edit.permission.size ="+permission.size());
               
    }				     
	
	if(WLX01_WM != null){
		System.out.println("WLX01_WM.size="+WLX01_WM.size());
	}else{
	   System.out.println("WLX01_WM == null");
	}	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX001W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>總機構每月填報資訊</title>
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
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post  action='#'>
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
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                      <td width="244"><font color='#000000' size=4><b> 
                        <center>
                          總機構每月填報資訊 
                        </center>
                        </b></font> </td>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
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
                          <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>金融機構代號</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>&nbsp;
                            <%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print(((String)((DataObject)WLX01_WM.get(0)).getValue("bank_no") == null)?"":(String)((DataObject)WLX01_WM.get(0)).getValue("bank_no"));%>
                          </td>
                          </tr>
                          
                          <tr class="sbody">
						  <td width='35%' align='left' bgcolor='#D8EFEE'>基準日</td>
						  <td width='65%' colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH>
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
        						<input type='button' name='getBeforeMonth' value="帶入" onClick="javascript:loadData(form,'loadWM');">                            	
        						-->將上個月的申報資料代入
                            </td>
                          </tr>
                       
                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>發行金融卡數(含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>
                            	<input type='text' name='PUSH_DebitCard_CNT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("push_debitcard_cnt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("push_debitcard_cnt")).toString());%>" size='8' maxlength='8' >    
                        	</td>
                          </tr>   
                        
                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>流通金融卡數(含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='TRAN_DebitCard_CNT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("tran_debitcard_cnt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("tran_debitcard_cnt")).toString());%>" size='8' maxlength='8' >                            	
                        	</td>
                          </tr>   
							
                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>ATM裝設台數(含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='ATM_CNT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("atm_cnt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("atm_cnt")).toString());%>" size='8' maxlength='8' >                            	
                        	</td>
                          </tr>   

                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>交易次數(本月) (含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='TRAN_CNT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("tran_cnt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("tran_cnt")).toString());%>" size='14' maxlength='14' >                            	
                        	</td>
                          </tr>   

                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>交易金額(本年累計) (含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='TRAN_AMT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("tran_amt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("tran_amt")).toString());%>" size='14' maxlength='14' >                            	
                        	</td>
                          </tr>   

                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>支票存款戶(含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='CHECK_DEPOSIT_CNT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("check_deposit_cnt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("check_deposit_cnt")).toString());%>" size='8' maxlength='8' >                            	
                        	</td>
                          </tr>   

                       	  <tr class="sbody">
							<td width='35%' align='left' bgcolor='#D8EFEE'>支票存款餘額(含分支機構)</td>
							<td width='65%' colspan=2 bgcolor='e7e7e7'>							
                            	<input type='text' name='CHECK_DEPOSIT_AMT' value="<%if(WLX01_WM != null && WLX01_WM.size() != 0) out.print( ((DataObject)WLX01_WM.get(0)).getValue("check_deposit_amt") == null ?"":(((DataObject)WLX01_WM.get(0)).getValue("check_deposit_amt")).toString());%>" size='14' maxlength='14' >                            	
                        	</td>
                          </tr>   
	                        
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
				        <%if(act.equals("newWM")){%>
				           <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'InsertWM','WM');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				           <%}%>
				      	<%}else if(act.equals("EditWM")){%> 
				      	   <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>                   	        	                                   		      				        
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'UpdateWM','WM');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>                        				        
				           <%}%>
				           <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){//Delete %>                   	        	                                   		     
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'DeleteWM','WM');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>                        				        				        
				           <%}%>
				      	<%}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>                        
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      
                      </tr>
                    </table>
                  </div></td>
              </tr>                       
      </table></td>
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
                          <li>本網頁提供總機構每月填報申報資訊功能。</li>                        
                          <li>新增時,可直接於空格內更改資料，資料更改完畢後，按<font color="#666666">【確定】</font>即將本表上的資料於資料庫中建檔。</li>                          
                          <li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>                         
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>                         
                          <li>【<font color="red" size=4>*</font>】為必填欄位。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</body>
</html>
