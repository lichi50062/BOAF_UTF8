<%
//96.10.02 add 增加牌告利率欄位 by 2295
//96.10.11 fix 讀取FX010W的權限(之前是讀取FX009W) by 2295
//98.03.31 add 基準利率-指標利率(月調)/基準利率(月調)/指數型房貸指標利率(月調) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	List WLX_S_RATE= (request.getAttribute("WLX_S_RATE")==null)?null:(List)request.getAttribute("WLX_S_RATE");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
    String loaddata = ( request.getParameter("loaddata")==null ) ? "" : (String)request.getParameter("loaddata");
    String syear = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
    String squarter = ( request.getParameter("S_QUARTER")==null ) ? "" : (String)request.getParameter("S_QUARTER");
    if(WLX_S_RATE!=null && !loaddata.equals("loaded"))//若不是要載入上月資料並且是修改資料時
    {
    syear =String.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("m_year"));
    squarter = String.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("m_quarter"));
    }

	Properties permission = ( session.getAttribute("FX010W")==null ) ? new Properties() : (Properties)session.getAttribute("FX010W");
	if(permission == null){
    System.out.println("FX010W_List.permission == null");
    }else{
    System.out.println("FX010W_List.permission.size ="+permission.size());
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

	String m_year="";
	String m_quarter="";
	Double period_1_fix_rate = null;
	Double period_1_var_rate = null;
	Double period_3_fix_rate = null;
	Double period_3_var_rate = null;
	Double period_6_fix_rate = null;
	Double period_6_var_rate = null;
	Double period_9_fix_rate = null;
	Double period_9_var_rate = null;
	Double period_12_fix_rate = null;
	Double period_12_var_rate = null;
	Double basic_pay_var_rate = null;
	Double period_house_var_rate = null;
	Double base_mark_rate = null;
	Double base_fix_rate = null;
	Double base_base_rate = null;
	
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
	
	
	if(WLX_S_RATE!=null)
	{
		period_1_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_1_fix_rate").toString());
		period_1_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_1_var_rate").toString());
		period_3_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_3_fix_rate").toString());
		period_3_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_3_var_rate").toString());
		period_6_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_6_fix_rate").toString());
		period_6_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_6_var_rate").toString());
		period_9_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_9_fix_rate").toString());
		period_9_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_9_var_rate").toString());
		period_12_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_12_fix_rate").toString());
		period_12_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_12_var_rate").toString());
		basic_pay_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("basic_pay_var_rate").toString());
		period_house_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_house_var_rate").toString());
		base_mark_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("base_mark_rate").toString());
		base_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("base_fix_rate").toString());
		base_base_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("base_base_rate").toString());
		
		//96.10.01增加的利率==================================================================================
		period_24_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_24_fix_rate").toString());
		period_24_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_24_var_rate").toString());
		period_36_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_36_fix_rate").toString());
		period_36_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_36_var_rate").toString());
		deposit_12_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_12_fix_rate").toString());
		deposit_12_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_12_var_rate").toString());
		deposit_24_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_24_fix_rate").toString());
		deposit_24_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_24_var_rate").toString());
		deposit_36_fix_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_36_fix_rate").toString());
		deposit_36_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_36_var_rate").toString());
		deposit_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("deposit_var_rate").toString());
		save_var_rate = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("save_var_rate").toString());
		//98.03.31增加的利率==================================================================================
		base_mark_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("base_mark_rate_month").toString());
		base_base_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("base_base_rate_month").toString());
		period_house_var_rate_month = Double.valueOf(((DataObject)WLX_S_RATE.get(0)).getValue("period_house_var_rate_month").toString());	
		//=======================================================================================================
	}
%>

<script language="javascript" src="js/FX010W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>金融機構警示帳戶調查資料維護</title>
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
<body leftmargin="15" topmargin="0">
<%if(act.equals("modify")){%>
<script language="javascript" type="text/JavaScript">
alert("已有該年季資料，將進行修改動作!");
</script>
<%}%>
<form method="post" name="atmdata">
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
     <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
  </tr>
  <tr>
     <td bgcolor="#FFFFFF" width="618">
      <table width="569" border="0" align="center" cellpadding="0" cellspacing="0" height="328">
      
              <tr>
                <td width="660" height="18"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="154" align="center"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="288" align="center"><font color="#000000" size="4"><b>農會信用部牌告利率彙總表</b></font> </td>
                      <td width="154" valign="middle"><img src="images/banner_bg1.gif" width="150" height="17" align="right"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="660" height="300">
                <table width="600" border="0" align="center" cellpadding="0" cellspacing="0" height="8">
                    <tr>
                    <td width="600" height="10">
                    </tr>
                    <tr>
                      
     				<div align="right"><jsp:include page="getLoginUser.jsp?" flush="true" /></div>
                    </tr>
                    <tr>
            <td width="600" height="82">
              <table width="600" border="1" align="center" cellpadding="1" cellspacing="1" bordercolor="#3A9D99" height="343">
                <tr>
                  <td class="sbody" width="190" bgcolor="#D8EFEE" align="left" height="27" colspan="5">金融機構代號</td>
                  <td class="sbody" width="410" bgcolor="e7e7e7" height="27">
                  <%=bank_no%>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="190" bgcolor="#D8EFEE" align="left" height="27" colspan="5">申報年/月(季)<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="27">
                    <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
                    <input type="text" name="nyear" value="<%=syear%>" size="3" maxlength="3" onblur="CheckYear(this)"
                     <%
                             if(!syear.equals("")){
                              out.print("disabled");
                              }
                            %>> 
                      <font color="#000000">年        
                      <select name="nquarter" size="1"
                   		 <%
                    		if(!squarter.equals("")){
                    		out.print("disabled");
                    		}
                    		%>>
                      <option><%=squarter%></option>
                    
                    </select>
                     
         		     <%//96.11.05 96/10月以前為季報.96/10月以後為月報
         		       if((Integer.parseInt(syear) * 100 + Integer.parseInt(squarter)) < 9610){ 
         		          out.print("季");
         		       }else{
         		          out.print("月");
         		       }         
                    %>
                    &nbsp;        
                    <%if(!squarter.equals("") && !syear.equals("")){
                    %>
                    <input type="hidden" name="hyear" value="<%=syear%>">                
                    <input type="hidden" name="hquarter" value="<%=squarter%>"><%}%>       
                    <%if(loaddata.equals("ok")){%>             
                    <input type="button" name="load" value="載入上月(季)申報資料" onClick="javascript:doSubmit(this.document.forms[0],'load');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        
                  
                    </p><%}%>            
                    </font></td>
                </tr>
                <tr class="sbody">
                  <td width="190" align="center" bgcolor="#D8EFEE" height="33" colspan="5">
                    活期存款機動利率<font color="red">*</font>
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='deposit_var_rate'
                    value="<% out.print((deposit_var_rate == null)? "":deposit_var_rate.toString()); %>" 
                   size='20' maxlength='20' style='text-align: right;'>                  
                   　                  
                  </td>
                </tr>
                
                <tr class="sbody">
                  <td width="190" align="center" bgcolor="#D8EFEE" height="33" colspan="5">
                    活期儲蓄存款機動利率<font color="red">*</font>
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='save_var_rate'
                    value="<% out.print((save_var_rate == null)? "":save_var_rate.toString()); %>" 
                   size='20' maxlength='20' style='text-align: right;'>                  
                   　                  
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" height="188" rowspan="14" align="center">
                    <p align="center"><font color="#000000">定期存款</font></p>
                  </td>
                  <td width="70" bgcolor="#D8EFEE" align="left" rowspan="2" colspan="2">一個月</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="32" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="32">     
                  <input type='text' name='period_1_fix_rate' 
                  value="<% out.print((period_1_fix_rate == null)? "":period_1_fix_rate.toString()); %>"
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="32" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="32">
                   <input type='text' name='period_1_var_rate' 
                    value="<% out.print((period_1_var_rate == null)? "":period_1_var_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>                  　
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left" rowspan="2" colspan="2">三個月</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='period_3_fix_rate'
                    value="<% out.print((period_3_fix_rate == null)? "":period_3_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>　
            		
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_3_var_rate' 
                   value="<% out.print((period_3_var_rate == null)? "":period_3_var_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
            	 </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">六個月</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="#E7E7E7" height="33">
                  <input type='text' name='period_6_fix_rate' 
                   value="<% out.print((period_6_fix_rate == null)? "":period_6_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_6_var_rate' 
                   value="<% out.print((period_6_var_rate == null)? "":period_6_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">九個月</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_9_fix_rate' 
                   value="<% out.print((period_9_fix_rate == null)? "":period_9_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>　
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="#E7E7E7" height="33">
                  <input type='text' name='period_9_var_rate' 
                   value="<% out.print((period_9_var_rate == null)? "":period_9_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>　
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">一年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_12_fix_rate' 
                   value="<% out.print((period_12_fix_rate == null)? "":period_12_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_12_var_rate'
                   value="<% out.print((period_12_var_rate == null)? "":period_12_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">二年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_24_fix_rate' 
                   value="<% out.print((period_24_fix_rate == null)? "":period_24_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_24_var_rate'
                   value="<% out.print((period_24_var_rate == null)? "":period_24_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">三年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_36_fix_rate' 
                   value="<% out.print((period_36_fix_rate == null)? "":period_36_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='period_36_var_rate'
                   value="<% out.print((period_36_var_rate == null)? "":period_36_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                
                
                <tr class="sbody">
                  <td width="100" bgcolor="#D8EFEE" height="188" rowspan="6" align="center">
                    <p align="center"><font color="#000000">定期儲蓄存款</font></p>
                  </td>
                  <td width="70" bgcolor="#D8EFEE" align="left" rowspan="2" colspan="2" >一年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="32" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="32">     
                  <input type='text' name='deposit_12_fix_rate' 
                  value="<% out.print((deposit_12_fix_rate == null)? "":deposit_12_fix_rate.toString()); %>"
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="32" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="32">
                   <input type='text' name='deposit_12_var_rate' 
                    value="<% out.print((deposit_12_var_rate == null)? "":deposit_12_var_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>                  　
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left" rowspan="2" colspan="2">二年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='deposit_24_fix_rate'
                    value="<% out.print((deposit_24_fix_rate == null)? "":deposit_24_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>　
            		
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='deposit_24_var_rate' 
                   value="<% out.print((deposit_24_var_rate == null)? "":deposit_24_var_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
            	 </td>
                </tr>
                <tr class="sbody">
                  <td width="70" bgcolor="#D8EFEE" align="left"  rowspan="2" colspan="2">三年</td>
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">固定<font color="red">*</font></td>
                  <td width="410" bgcolor="#E7E7E7" height="33">
                  <input type='text' name='deposit_36_fix_rate' 
                   value="<% out.print((deposit_36_fix_rate == null)? "":deposit_36_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="60" bgcolor="#D8EFEE" align="left" height="33" colspan="2">機動<font color="red">*</font></td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                  <input type='text' name='deposit_36_var_rate' 
                   value="<% out.print((deposit_36_var_rate == null)? "":deposit_36_var_rate.toString()); %>" 
                  size='20' maxlength='20' style='text-align: right;'>
                  </td>
                </tr>
               
                
                
                <tr class="sbody">
                  <td width="190" align="center" bgcolor="#D8EFEE" height="33" colspan="5">
                    基本放款利率(機動)<font color="red">*</font>
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='basic_pay_var_rate'
                    value="<% out.print((basic_pay_var_rate == null)? "":basic_pay_var_rate.toString()); %>" 
                   size='20' maxlength='20' style='text-align: right;'>                  
                   　                  
                  </td>
                </tr>
                
                <tr class="sbody">
                  <td width="190" align="center" bgcolor="#D8EFEE" height="33" colspan="2" rowspan="2">
                    指數型房貸指標利率<font color="red">*</font>
                  </td>
                  <td width="119" align="center" bgcolor="#D8EFEE" height="17" colspan="3">
                    季調
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                    <input type='text' name='period_house_var_rate'  
                     value="<% out.print((period_house_var_rate == null)? "":period_house_var_rate.toString()); %>"
                    size='20' maxlength='20' style='text-align: right;'>                　                  
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="119" align="center" bgcolor="#D8EFEE" height="16" colspan="3">
                    月調
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                    <input type='text' name='period_house_var_rate_month'  
                     value="<% out.print((period_house_var_rate_month == null)? "":period_house_var_rate_month.toString()); %>"
                    size='20' maxlength='20' style='text-align: right;'>                　                  
                  </td>
                </tr>

                <tr class="sbody">
                  <td width="110" align="center" bgcolor="#D8EFEE" height="71" colspan="3" rowspan="5">
                    基準利率
                  </td>
                  <td width="80" align="center" bgcolor="#D8EFEE" height="33" rowspan="2">
                    指標利率<font color="red">*</font>
                  </td>
                  <td width="44" align="center" bgcolor="#D8EFEE" height="17" >
                    季調
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='base_mark_rate'
                    value="<% out.print((base_mark_rate == null)? "":base_mark_rate.toString()); %>" 
                   size='20' maxlength='20' style='text-align: right;'>                 　                  
                  </td>
                </tr>
                 <tr class="sbody">
                  <td width="44" align="center" bgcolor="#D8EFEE" height="17" >
                    月調</td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='base_mark_rate_month'
                    value="<% out.print((base_mark_rate_month == null)? "":base_mark_rate_month.toString()); %>" 
                    size='20' maxlength='20' style='text-align: right;'>    
                  </td>
                </tr>

                <tr class="sbody">
                  <td width="80" align="center" bgcolor="#D8EFEE" height="33" colspan="2">
                    一定比率<font color="red">*</font>
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='base_fix_rate'
                    value="<% out.print((base_fix_rate == null)? "":base_fix_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>                 　                  
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="80" align="center" bgcolor="#D8EFEE" height="33" rowspan="2">
                    基準利率<font color="red">*</font>
                  </td>
                  <td width="44" align="center" bgcolor="#D8EFEE" height="17" >
                    季調
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='base_base_rate' 
                    value="<% out.print((base_base_rate == null)? "":base_base_rate.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>                 　                  
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="44" align="center" bgcolor="#D8EFEE" height="17" >
                    月調
                  </td>
                  <td width="410" bgcolor="e7e7e7" height="33">
                   <input type='text' name='base_base_rate_month' 
                    value="<% out.print((base_base_rate_month == null)? "":base_base_rate_month.toString()); %>"
                   size='20' maxlength='20' style='text-align: right;'>                 　                  
                  </td>
                </tr>
              </table>
              <p>　
   </td>
 </tr>


                    <td width="600" height="21"><div align="right"><div align="right">
						<div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>
                    </div>
                      </div></td>
      			</table>
      			</td>
  			  </tr>
 			  <tr>
                <td width="660" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody" height="176">
                    <tr>
                      <td colspan="2" width="583" height="41">
                      <div align="center">
                      <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
               <%if(act.equals("new")){
                    if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add
                %>    
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
               <% } }else{
               if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>            
                <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
               <% }}%>         
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                      </table>
                      </div>
                     </td>
                    </tr>
                    <tr>
                      <td colspan="2" width="583" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明           
                        :</font></font> </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
 						<ul>
 						  <li class="sbody" ><font color="red">註:96年10月以前為季報,96年10月以後.開始為月報。</font></li>
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】</font>即將本表上的資料, 於資料庫中建檔。</li>         
                      	  <li class="sbody" >按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】</font>即將本表上的資料清空</li>         
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>         
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</li>
                        </ul>                            
                      </td>
                    </tr>
                  </table></td>
              </tr>
         
</table>
</form>
</table>
</body>
