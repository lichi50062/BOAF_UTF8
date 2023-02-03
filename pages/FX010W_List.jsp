<%
// 94.11.09 first design by 4180
// 98.03.31 add 基準利率-指標利率(月調)/基準利率(月調)/指數型房貸指標利率(月調) by 2295
//100.01.27 fix 無資料查詢 by 2295
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
    List WLX_S_RATE= (request.getAttribute("WLX_S_RATE")==null)?null:(List)request.getAttribute("WLX_S_RATE");
    List inidate= (request.getAttribute("WLX10_INI")==null)?null:(List)request.getAttribute("WLX10_INI");  //取得鎖定年季
    List lockdate= (request.getAttribute("WLX10_LOCK")==null)?null:(List)request.getAttribute("WLX10_LOCK");  //取得鎖定年季
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX010W")==null ) ? new Properties() : (Properties)session.getAttribute("FX010W");
    if (permission == null) {
        System.out.println("FX010W_List.permission == null");
    }else {
        System.out.println("FX010W_List.permission.size ="+permission.size());
    }
    int iniyear=0,  iniquarter=0;//初始年季
	String lockyear="", lockquarter="";//鎖定年季
		
    if(inidate!=null){
      iniyear = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_year").toString());
      iniquarter = Integer.parseInt(((DataObject)inidate.get(0)).getValue("m_month").toString());
    }

	String m_year="";
    String m_quarter="";
    Double period_1_fix_rate;
    Double period_1_var_rate;
    Double period_3_fix_rate;
    Double period_3_var_rate;
    Double period_6_fix_rate;
    Double period_6_var_rate;
    Double period_9_fix_rate;
    Double period_9_var_rate;
    Double period_12_fix_rate;
	Double period_12_var_rate;
    Double basic_pay_var_rate;
    Double period_house_var_rate;
    Double base_mark_rate;
    Double base_fix_rate;
    Double base_base_rate;
    String maintain_id="";
    String maintain_name="";
    String maintain_date="";
    
   	//96.10.01增加的利率==================================================================================
    Double period_24_fix_rate=null;//定期存款-二年-固定
    Double period_24_var_rate=null;//定期存款-二年-變動
    Double period_36_fix_rate=null;//定期存款-三年-固定
    Double period_36_var_rate=null;//定期存款-三年-變動
    Double deposit_12_fix_rate=null;//定期儲蓄存款-一年-固定
    Double deposit_12_var_rate=null;//定期儲蓄存款-一年-變動
    Double deposit_24_fix_rate=null;//定期儲蓄存款-二年-固定
    Double deposit_24_var_rate=null;//定期儲蓄存款-二年-變動
    Double deposit_36_fix_rate=null;//定期儲蓄存款-三年-固定
    Double deposit_36_var_rate=null;//定期儲蓄存款-三年-變動
    Double deposit_var_rate=null;//活期存款機動利率
    Double save_var_rate=null;//活期儲蓄存款機動利率 
    //98.03.31增加的利率==================================================================================
    Double base_mark_rate_month=null;//基準利率-指標利率(月調)
    Double base_base_rate_month=null;//基準利率(月調)
    Double period_house_var_rate_month=null;//指數型房貸指標利率(月調)
    //=======================================================================================================

%>
<script language="javascript" src="js/FX010W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>農會信用部牌告利率彙總表</title>
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
    	    lockquarter=String.valueOf(((DataObject)lockdate.get(k)).getValue("m_quarter"));
	%>
		pushArray('<%=lockyear%>');
		pushArray('<%=lockquarter%>');
	<%}%>
	</script>
<%}%>
<table width="832" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
           <td width="832"><img src="images/space_1.gif" width="12" height="12"></td>
          </tr>       
        <tr>
        
          <td bgcolor="#FFFFFF" width="832">
        
          <table width="1060" border="0" align="center" cellpadding="0" cellspacing="0" height="310">
              <tr>
                <td width="1060" height="18">
                  <table width="1060" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="344"><img src="images/banner_bg1.gif" width="330" height="17" align="left"></td>
                      <td width="372" align="center"><b>
                        <center>
                        <font color="#000000" size="4">農會信用部牌告利率彙總表</font>
                        </center>
                        </b> </td>
                      <td width="344" align="right"><img src="images/banner_bg1.gif" width="330" height="17"></td>
                    </tr>
                  </table>

               </td>
              </tr>
              <tr>
                <td width="1060" height="281">
                <table width="1060" border="0" align="center" cellpadding="0" cellspacing="0" height="16">
                    <tr>
                      <td height="16" width="16" >  </td>
                    </tr>

                    <tr>
                       
                       <div align="right"><jsp:include page="getLoginUser.jsp?width=1060" flush="true" /></div>
                    </tr>
          <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr ><td height="1" width="1060">

          <table border="1" cellspacing="1" bordercolor="#3A9D99" width="1060" height="30" class="sbody" cellpadding="0">
            <tr>
            <form name="date" method="post">
                <td class="sbody" bgcolor="#D8EFEE" width="131" height="30"><p align="center">申報年季</p></td>
                 <td class="sbody" bgcolor="#E7E7E7" width="635" height="30"  nowrap>                   
                     <input type="text" name="S_YEAR" size="3" value="<% 
                      Date today = new Date();
                     out.print(today.getYear()-11);
                     %>">年 第      
                    <%
                    String select1 ="";
                    String select2 ="";
                    String select3 ="";
                    String select4 ="";

                   	if(today.getMonth() < 3)
                   		select4 ="selected";
                   	else if(today.getMonth() < 6)
                   		select1 ="selected";
                   	else if(today.getMonth() < 9)
                   		select2 ="selected";
                   	else if(today.getMonth() < 12)
                   		select3 ="selected";	
                   		
                %>
               		<select type="text" name='S_QUARTER' size="1">      
                		<option value="1" <%= select1 %> >01</option>
                		<option value="2" <%= select2 %> >02</option>
                		<option value="3" <%= select3 %> >03</option>
                		<option value="4" <%= select4 %> >04</option>
                	 </select>
                     季       
           
                    <a href="javascript:newSubmit(this.document.forms[0],'new','<%=iniyear*12+iniquarter%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1);">
                    <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                    <font size="2" color="red" face="標楷體">單位:年利率x.xxx%</font>
            </td>
  		</form>
        </tr>
         </table>
            </td>
             </tr>
                  <%}%>

                    <tr class="sbody">
                      <td  class="sbody"   height="126" width="1060">
                      <div align="right">	
                      <table class="sbody" width="1060" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="34" >
                          <tr  bgcolor="#9AD3D0">
                          	<td  width="32" rowspan="3" align="center" height="1">
                          	   <p style="margin-top: 0; margin-bottom: 0">申報年月(季)</p>
                          	</td>
                          	<td  width="32" rowspan="3" align="center" height="1">
                          	   <p style="margin-top: 0; margin-bottom: 0">活期存款機動利率</p>
                          	</td>
                          	<td  width="32" rowspan="3" align="center" height="1">
                          	   <p style="margin-top: 0; margin-bottom: 0">活期儲蓄存款機動利率</p>
                          	</td>
                          	<td  width="450" colspan="14" align="center" height="1">
                          	   <p style="margin-top: 0; margin-bottom: 0">定期存款</p>
                          	</td>
                          	<td  width="250" colspan="6" align="center" height="1">
                          	   <p style="margin-top: 0; margin-bottom: 0">定期儲蓄存款</p>
                          	</td>
                          	<td rowspan="3"  align="center" height="1" width="45"> 
                          		<p style="margin-top: 0; margin-bottom: 0"> 基本放款利率(機動)</p>
                          	</td>
                          	<td  rowspan="2"  align="center" height="1" width="40" colspan="2">
                          		<p style="margin-top: 0; margin-bottom: 0">指數型房貸指標利率</p>
                          	</td>
                          	<td colspan="5" width="89" align="center" height="1">
                          		<p style="margin-top: 0; margin-bottom: 0">基準利率</p> 
                          	</td>                          	                         
                          	<!--td rowspan="3" width="70" align="center" height="1">
                          		<p style="margin-top: 0; margin-bottom: 0">異動者</p>
                          		<p style="margin-top: 0; margin-bottom: 0">帳號/</p>
                          		<p style="margin-top: 0; margin-bottom: 0">姓名</p>
                          	</td-->
                          	<td rowspan="3" width="55" align="center" height="1">異動日期</td>              
                            </tr>

                          	<tr  bgcolor="#9AD3D0">
                          	<td  colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: 0; margin-bottom: 0">一個月</p>
                          	</td>
                          	<td  colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: 0; margin-bottom: 0">三個月</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">六個月</p>
                          	</td>
                         	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">九個月</p>
                          	</td>
                         	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">一年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">二年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">三年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">一年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">二年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="90"> 
                          		<p style="margin-top: -1; margin-bottom: -3">三年</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="45"> 
                          		<p style="margin-top: -1; margin-bottom: -3">指標</p>
                          		<p style="margin-top: -1; margin-bottom: -3">利率</p>
                          	</td>
                          	<td rowspan="2" align="center" height="1" width="45"> 
                          		<p style="margin-top: -1; margin-bottom: -3">一定</p>
                          		<p style="margin-top: -1; margin-bottom: -3">比率</p>
                          	</td>
                          	<td colspan="2" align="center" height="1" width="45"> 
                          		<p style="margin-top: -1; margin-bottom: -3">基準</p>
                          		<p style="margin-top: -1; margin-bottom: -3">利率</p>
                          	</td>
                          </tr>
                          <tr class="sbody" bgcolor="#9AD3D0">
                          <td  width="45" class="sbody" align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td  width="45" class="sbody" align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>
                    	  <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>
                           <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>    
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>    
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>     
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>    
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>    
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">固定</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">                 
                          	<p style="margin-top: 0; margin-bottom: 0">機動</p>
                          </td>  
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">季調</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1">
                          	<p style="margin-top: 0; margin-bottom: 0">月調</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1"> 
                          	<p style="margin-top: 0; margin-bottom: 0">季調</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1"> 
                          	<p style="margin-top: 0; margin-bottom: 0">月調</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1"> 
                          	<p style="margin-top: 0; margin-bottom: 0">季調</p>
                          </td>
                          <td width="45" class="sbody"   align="center" height="1"> 
                          	<p style="margin-top: 0; margin-bottom: 0">月調</p>
                          </td>
                 
                          </tr>

                          <%
                          if(WLX_S_RATE.size()==0)
                          {
                          %>
                            <tr class="sbody" bgcolor="#D8EFEE"><td colspan="32" align="center"  height="17" width="832" >
                              <font   class="sbody">尚無資料</font></td></tr>
                           <%
                           }else{
							for(int i=0;i<WLX_S_RATE.size();i++){
							    period_24_fix_rate=null;//定期存款-二年-固定
   							    period_24_var_rate=null;//定期存款-二年-變動
    							period_36_fix_rate=null;//定期存款-三年-固定
    							period_36_var_rate=null;//定期存款-三年-變動
    							deposit_12_fix_rate=null;//定期儲蓄存款-一年-固定
    							deposit_12_var_rate=null;//定期儲蓄存款-一年-變動
    							deposit_24_fix_rate=null;//定期儲蓄存款-二年-固定
    							deposit_24_var_rate=null;//定期儲蓄存款-二年-變動
    							deposit_36_fix_rate=null;//定期儲蓄存款-三年-固定
    							deposit_36_var_rate=null;//定期儲蓄存款-三年-變動
    							deposit_var_rate=null;//活期存款機動利率
    							save_var_rate=null;//活期儲蓄存款機動利率
    							//98.03.31增加的利率==================================================================================
    							base_mark_rate_month=null;//基準利率-指標利率(月調)
    							base_base_rate_month=null;//基準利率(月調)
    							period_house_var_rate_month=null;//指數型房貸指標利率(月調)
    							
								m_year =String.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("m_year"));
								m_quarter = String.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("m_quarter"));
								period_1_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_1_fix_rate").toString());
								period_1_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_1_var_rate").toString());
								period_3_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_3_fix_rate").toString());
								period_3_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_3_var_rate").toString());
								period_6_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_6_fix_rate").toString());
								period_6_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_6_var_rate").toString());
								period_9_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_9_fix_rate").toString());
								period_9_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_9_var_rate").toString());
								period_12_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_12_fix_rate").toString());
								period_12_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_12_var_rate").toString());
								basic_pay_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("basic_pay_var_rate").toString());
								period_house_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_house_var_rate").toString());
								base_mark_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("base_mark_rate").toString());
								base_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("base_fix_rate").toString());
								base_base_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("base_base_rate").toString());
								maintain_id = String.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("user_id"));
								maintain_name = String.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("user_name"));								
								maintain_date=Utility.getCHTdate((((DataObject)WLX_S_RATE.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   
                         	    //96.10.01增加的利率==================================================================================
								period_24_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_24_fix_rate").toString());
								period_24_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_24_var_rate").toString());
								period_36_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_36_fix_rate").toString());
								period_36_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_36_var_rate").toString());
								deposit_12_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_12_fix_rate").toString());
								deposit_12_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_12_var_rate").toString());
								deposit_24_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_24_fix_rate").toString());
								deposit_24_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_24_var_rate").toString());
								deposit_36_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_36_fix_rate").toString());
								deposit_36_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_36_var_rate").toString());
								deposit_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("deposit_var_rate").toString());
								save_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("save_var_rate").toString());
		                 	    //98.03.31增加的利率==================================================================================
		                 	    base_mark_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("base_mark_rate_month").toString());//基準利率-指標利率(月調)
								base_base_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("base_base_rate_month").toString());//基準利率(月調)
								period_house_var_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(i)).getValue("period_house_var_rate_month").toString());//指數型房貸指標利率(月調)
							    //=======================================================================================================
        						
                         
                         %>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>">
                          <td class="sbody"  align=right height="17" width="30">
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
                                   	if(locked==false)
                                   	{
                                   	out.print( "<u><a href=\"FX010W.jsp?act=Edit&myear="+ m_year+"&mquarter="+m_quarter+"\">");
                                   	}
                                   	out.print(m_year+"/"+ m_quarter+"</a>");
                                   	locked=false;
                              
                              }else
                              {
                                out.print(m_year+"/"+ m_quarter);
                              }
                           }
                           else
                            {
                            if(lockdate!=null ){
                             for(int c=0;c<lockdate.size();c++){     	 	
                                 	 	lockyear=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_year"));
                                 	  lockquarter=String.valueOf(((DataObject)lockdate.get(c)).getValue("m_quarter"));
                                  	if(m_year.equals(lockyear) && m_quarter.equals(lockquarter))locked=true; 
                                  }
                            }         
                                   	if(locked==false)
                                   	{
                                   	out.print( "<u><a href=\"FX010W.jsp?act=Edit&myear="+ m_year+"&mquarter="+m_quarter+"\">");
                                   	}
                                   	out.print(m_year+"/"+ m_quarter+"</a>");
                                   	locked=false;
                            }
                          %>

                            </u></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=save_var_rate%></td>  
                            <td class="sbody"  align=right height="17" width="45"><%=period_1_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_1_var_rate%></td>      
                            <td class="sbody"  align=right height="17" width="45"><%=period_3_fix_rate%></td>     
                            <td class="sbody"  align=right height="17" width="45"><%=period_3_var_rate%></td>      
                            <td class="sbody"  align=right height="17" width="45"><%=period_6_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_6_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_9_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_9_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_12_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_12_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_24_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_24_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_36_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_36_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_12_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_12_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_24_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_24_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_36_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=deposit_36_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=basic_pay_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_house_var_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=period_house_var_rate_month%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=base_mark_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=base_mark_rate_month%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=base_fix_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=base_base_rate%></td>
                            <td class="sbody"  align=right height="17" width="45"><%=base_base_rate_month%></td>                            
				 			<!--td class="sbody"  align=left height="17" width="70"><%=maintain_id%> / <%=maintain_name%></td-->
                            <td class="sbody" align=center height="17" width="55"><%=maintain_date%></td>
                            
                            </tr>
                          <% }//end for
                                   }//end else %>
                         </table>
                      </div>
                      </td>
                    </tr>


                    <td width="1060" height="56"><div align="right"><div align="right">
						<div align="right"><jsp:include page="getMaintainUser.jsp?width=1060" flush="true" /></div>
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
                          <li><font color="red">註:96年10月以前為季報,96年10月以後.開始為月報。</font></li>
                          <li>輸入申報之年季，再點選<font color="#666666">【新增】按鈕</font>可新增該年/月(季)「牌告利率」之資料。
                          <li>點選所列之[申報年/月(季)]可變更該申報之資料。
                          <li>本表係按最近的[申報年/月(季)]先排序,依此類推。
                          <li><font color="#ff0000">如果在[申報年/月(季)]欄位沒有出現底線,表示巳辦理[鎖定],僅提供查詢不可再異動</font></li>
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
