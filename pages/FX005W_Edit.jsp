<%
//96.11.28 add 金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.lang.Integer" %>
<%
	List WLX05_M_ATM= (request.getAttribute("WLX05_M_ATM")==null)?null:(List)request.getAttribute("WLX05_M_ATM");
	List inidate= (request.getAttribute("WLX05_INI")==null)?null:(List)request.getAttribute("WLX05_INI");  //取得初始年月
	List hisdata= (request.getAttribute("WLX05_HIS")==null)?null:(List)request.getAttribute("WLX05_HIS");  //取得已申報資料
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String loaddata = ( request.getParameter("loaddata")==null ) ? "" : (String)request.getParameter("loaddata");
    String syear = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
    String smonth = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
    String editable = ( request.getParameter("editable")==null ) ? "" : (String)request.getParameter("editable");
    
    if(WLX05_M_ATM!=null && !loaddata.equals("loaded")){//若不是要載入上月資料並且是修改資料時        
    	syear =String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("m_year"));
    	smonth = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("m_month"));
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
    String yearamt_atm="";                        
    String maintain_id="" ;
    String maintain_name="" ;
    String maintain_date="" ;
    String debitcard_mtrancnt="";      
	String debitcard_ytrancnt="";      
	String debitcard_monthamt="";      
	String debitcard_yearamt="";       
    if(WLX05_M_ATM!=null){
   	   m_year = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("m_year"));
       m_month =String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("m_month"));
       pdebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("push_debitcard_cnt"));
       cancdebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("canc_debitcard_cnt"));
       udebitcard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("use_debitcard_cnt"));
       pbincard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("push_bincard_cnt"));
       cancbincard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("canc_bincard_cnt"));
       ubincard = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("use_bincard_cnt"));
       atmcnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("atm_cnt")));
       mtrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("month_tran_cnt")));
       ytrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("year_acctran_cnt")));
       monthamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("month_tran_amt")));
       yearamt_atm =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("year_acctran_amt")));
       //96.11.23 add 金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元)
       debitcard_mtrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("debitcard_month_tran_cnt")));
       debitcard_ytrancnt = Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("debitcard_year_acctran_cnt")));
       debitcard_monthamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("debitcard_month_tran_amt")));
       debitcard_yearamt =Utility.setCommaFormat(String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("debitcard_year_acctran_amt")));
 	   maintain_id = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("user_id"));
	   maintain_name = String.valueOf(((DataObject)WLX05_M_ATM.get(0)).getValue("user_name"));
	   maintain_date = Utility.getCHTdate((((DataObject)WLX05_M_ATM.get(0)).getValue("update_date")).toString().substring(0, 10), 0);
   	}
	//初始年月
	String iniyear="", inimonth="";
    if(inidate!=null){
    	iniyear = String.valueOf(((DataObject)inidate.get(0)).getValue("m_year"));
    	inimonth =String.valueOf(((DataObject)inidate.get(0)).getValue("m_month"));
	}
	
	Properties permission = ( session.getAttribute("FX005W")==null ) ? new Properties() : (Properties)session.getAttribute("FX005W");
	if(permission == null){
    	System.out.println("FX005W_List.permission == null");
    }else{
    	System.out.println("FX005W_List.permission.size ="+permission.size());
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

	//取得上個月的資料 本月 及 上月累計
	String lastyear="";
	String lastmonth="";	
	String yearcnt ="0"; 
	String yearamt ="0";
	String sqlTemp="";
	List WLX05_M_ATM_LAST=null;
	List paramList = new ArrayList();
	if(!smonth.equals("1")){ 	    
		lastyear=syear; 
	    lastmonth =String.valueOf(Integer.parseInt(smonth)-1);     
    	sqlTemp = "select * from WLX05_M_ATM where bank_no=? and m_year=? and m_month=?";
    	paramList.add(bank_no);
    	paramList.add(lastyear);
    	paramList.add(lastmonth);
		WLX05_M_ATM_LAST = DBManager.QueryDB_SQLParam(sqlTemp,paramList,"year_acctran_cnt,year_acctran_amt");
		if(WLX05_M_ATM_LAST.size() ==1){
 			yearcnt = String.valueOf(((DataObject)WLX05_M_ATM_LAST.get(0)).getValue("year_acctran_cnt"));
 			yearamt = String.valueOf(((DataObject)WLX05_M_ATM_LAST.get(0)).getValue("year_acctran_amt"));
		}
    } 
 	
%>

<script language="javascript" src="js/FX005W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>金融卡發卡及ATM裝設情形維護</title>
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
alert("已有該年月資料，將進行修改動作!");
</script>
<%}%>
<!--95/01/09 fix by 4180 若上月份無資料，則無法新增資料-->
<%if(loaddata.equals("false") && hisdata.size()!=0){%>
<script language="javascript" type="text/JavaScript">
alert("請完成前一個月的資料申報後，再辦理本月份資料申報");
history.go(-1);
</script>
<%}%>
<!--95/01/09 fix by 4180 若下月份已申報資料，則無法修改資料-->
<%if(editable.equals("false")){%>
<script language="javascript" type="text/JavaScript">
alert("次一個月的申報資料巳建檔，僅提供查詢功能，該月份不可再辦理修改");
history.go(-1);
</script>
<%}%>
<!--95/01/09 fix by 4180 若不為初始申報年月，則提醒使用者-->
<%if((!syear.equals(iniyear) || !smonth.equals(inimonth)) && inidate!=null 
	  && hisdata.size()==0 ){%>
<script language="javascript" type="text/JavaScript">  
 if(confirm("1.農金局規定本案起始的申報年月為<%=iniyear%>年<%=inimonth%>月，\n"+
            "   與你第一次將申報年月(畫面.申報年月) 不一致!\n"+
      		"2.如果貴單位是<%=iniyear%>年<%=inimonth%>月前成立的單位，\n"+
            "   請務必從<%=iniyear%>年<%=inimonth%>月的資料開始申報。\n"+
            "3.如果<%=iniyear%>年<%=inimonth%>月後才成立的單位，則從成立之\n"+
            "   當月起之資料申報。\n"+
            "4.資料自[起始申報]月份起，不管有無異動均須逐月申報。")
            )
	 history.go(1);
  else
    history.go(-1);     

</script>
<%}%>

<script language="javascript" type="text/JavaScript">
	pushData('<%=yearcnt%>','<%=yearamt%>','<%=loaddata%>');
</script>
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr>
   		 <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr>
          <td bgcolor="#FFFFFF" width="618">
		  <table width="569" border="0" align="center" cellpadding="0" cellspacing="0" height="328">
		  <form method="post" name="atmdata">
              <tr>
                <td width="660" height="18"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="154" align="center"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="288" align="center"><b>
                        <center><font color="#000000" size="4">金融卡發卡及ATM裝設情形維護</font>
                        </center>
                        </b> </td>
                      <td width="154" valign="middle"><img src="images/banner_bg1.gif" width="150" height="17" align="right"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="660" height="300"><table width="579" border="0" align="center" cellpadding="0" cellspacing="0" height="8">

                    <tr>
                      <td width="600" height="10">
                    </tr>

                    <tr>
                      
     <div><jsp:include page="getLoginUser.jsp?" flush="true" /></div>
                    </tr>


                    <tr>
                      <td width="600" height="82"><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99" height="365">
                        <tr>
						<td class="sbody" width='146' bgcolor='#D8EFEE' align='left' height="24" colspan="3">金融機構代號</td>
						<td class="sbody" width='438' bgcolor='e7e7e7' height="24" >

                            <%=bank_no%>
						</td>
                        </tr>
        			    <tr class="sbody">
						<td width='146' bgcolor='#D8EFEE' align='left' height="17" colspan="3">申報年月<font color="red" size="3">*</font></td>
						<td width='438' bgcolor='e7e7e7' height="17" >
                            <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
                            <input type='text' name='nyear' value="<%=syear%>" size='3' maxlength='3' onblur='CheckYear(this)'
                            <%
                             if(!syear.equals("")){
                              out.print("disabled");
                              }
                            %>
                            >
        						<font color='#000000'>年 

        						<select name=nmonth
        						<%
        						if(!smonth.equals("")){
        						out.print("disabled");
        						}
        						%> size="1"
        						>

        						<option></option>
        						 <%
        						 for(int i=0;i<=12;i++)
        						 {
        						  if(!smonth.equals(String.valueOf(i)))
        						   out.print("<option value="+i+" >"+i+"</option>");
        						  else
        						    out.print("<option value="+i+" selected >"+i+"</option>");
        						 }
        						 %>


        						</select> 月 
        						<%
        						if(!smonth.equals("") && !syear.equals("")){
        						%>
        						<input type="hidden" name="hyear" value="<%=syear%>"> 
                    			<input type="hidden" name="hmonth" value="<%=smonth%>"> 
        						<%}%> 

        						 <%if(loaddata.equals("ok")){%> 


        						<input type="button" name="load" value="載入上月申報資料" onClick="javascript:doSubmit(this.document.forms[0],'load');"> 
        					
                            </p><%}%>
                            </font>
						      </td>
        			    </tr>



        			    <tr class="sbody">
						    <td width='27' bgcolor='#D8EFEE' rowspan="7">
                            	<p align="center" style="margin-top: 2">金</p>
								<p align="center" style="margin-top: 2">融</p>
								<p align="center" style="margin-top: 2">卡&nbsp;</p>                            
                            </td>
							<td width='20' bgcolor='#D8EFEE' height="45" rowspan="3" align="center"><font color="red">(歷年)</font>發卡情形</td>
							<td width='92' bgcolor='#D8EFEE' align='left' height="29">發行張數(A)<font color="red" size="3">*</font></td>
							<td width='438' bgcolor='e7e7e7' height="29">
                            	<input type='text' name='debitcard'
                            	 	   value="<%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(pdebitcard)+Integer.parseInt(pbincard)))%>"
                            		   size='20' maxlength='20'
                            		   onFocus='this.value=changeVal(this)'
                            		   onBlur='checkPoint_focus(this);this.value=changeStr(this);'
                            		   style='text-align: right;'>
							</td>
        			    </tr>

        			    <tr class="sbody">
							<td width='92' bgcolor='#D8EFEE' align='left' height="15">停卡張數(B)<font color="red" size="3">*</font></td>
							<td width='438' bgcolor='e7e7e7' height="15">
						  		<input type='text' name='cancdebitcard' 
						  			   value="<%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(cancdebitcard)+Integer.parseInt(cancbincard)))%>" 
						  			   size='20' maxlength='25' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
							</td>
        			    </tr>

        			    <tr class="sbody">
							<td width='92' bgcolor='#D8EFEE' align='left' height="14">流通張數(A-B)<font color="red" size="3">*</font></td>
							<td width='438' bgcolor='e7e7e7' height="14">
						    	<input type='text' name='usedebitcard' 
						    		   value="<%=Utility.setCommaFormat(String.valueOf(Integer.parseInt(udebitcard)+Integer.parseInt(ubincard)))%>" 
						    		   size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
							</td>
        			    </tr>
						
						<tr class="sbody">
						  	<td width='24' align='center' bgcolor='#D8EFEE' rowspan="2">交易次數</td>
						  	<td width='89' align='left' bgcolor='#D8EFEE' height="39">本月<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="39">
                            	<input  type='text' name='debitcard_mtrancnt' 
                            			value="<%=debitcard_mtrancnt%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                        	次</td>
	  					</tr>

                        <tr class="sbody">
						  	<td width='89' align='left' bgcolor='#D8EFEE' height="40">本年累計<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="40">
                            	<input  type='text' name='debitcard_ytrancnt' value="<%=debitcard_ytrancnt%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;' >
                        	次</td>
	  					</tr>
						
                        <tr class="sbody">
						  	<td width='24' align='center' bgcolor='#D8EFEE' height="80" rowspan="2">交易金額</td>                         
						  	<td width='92' align='left' bgcolor='#D8EFEE' height="32">本月<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="32">
                            	<input type='text' name='debitcard_monthamt' 
                            		   value="<%=debitcard_monthamt%>" size='25' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                        	元</td> 
	  					</tr>

                        <tr class="sbody">
						  	<td width='89' align='left' bgcolor='#D8EFEE' height="43">本年累計<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="43">
                            	<input type='text' name='debitcard_yearamt' 
                            		   value="<%=debitcard_yearamt%>"  size='25' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;' >
                        	元</td>
	  					</tr>

                       

                        <tr class="sbody">
						  	<td width='140' align='left' bgcolor='#D8EFEE' colspan="3">ATM裝設台數<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="15">
                            	<input type='text' name='setup_atm' value="<%=atmcnt%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                        		<font color="red" size="2">(係表示截至申報月底日運作中的台數)</font></td>
	  					</tr>

                        <tr class="sbody">
						  <td width='27' align='left' bgcolor='#D8EFEE' rowspan="4">
                            <p align="center">ATM自</p>
							<p align="center">動</p>
							<p align="center">化</p>
							<p align="center">機</p>
							<p align="center">器</td>
							<td width='24' align='center' bgcolor='#D8EFEE' rowspan="2">
							交易次數</td>                         
							<td width='89' align='left' bgcolor='#D8EFEE' height="32">本月<font color="red" size="3">*</font></td>
							<td width='438' bgcolor='e7e7e7' height="32">
                            	<input  type='text' name='mtrancnt' value="<%=mtrancnt%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                        	次</td> 
	  					</tr>

                        <tr class="sbody">
						  	<td width='89' align='left' bgcolor='#D8EFEE' height="38">本年累計<font color="red" size="3">*</font></td>
						  	<td width='438' bgcolor='e7e7e7' height="38">                            	
                            	<input  type='text' name='ytrancnt' value="<%=ytrancnt%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;' >
                        	次</td>
	  					</tr>

                        <tr class="sbody">
						  <td width='24' align='center' bgcolor='#D8EFEE' rowspan="2">交易金額<p></td>
						  <td width='89' align='left' bgcolor='#D8EFEE' height="23">本月<font color="red" size="3">*</font></td>
						  <td width='438' bgcolor='e7e7e7' height="23">
                            <input type='text' name='monthamt' value="<%=monthamt%>" size='25' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                        	元</td>
	  					</tr>

                        <tr class="sbody">
						  <td width='89' align='left' bgcolor='#D8EFEE' height="38">本年累計<font color="red" size="3">*</font></td>
						  <td width='438' bgcolor='e7e7e7' height="38">
							<input type='text' name='yearamt_atm' value="<%=yearamt_atm%>"  size='25' maxlength='20' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;' >                            
                        	元</td>
                        </tr>
                        	</form>
                        </Table>
                      </td>
                    </tr>


                    <td width="600" height="21"><div align="right"><div align="right">

					<div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>

                    </div>
                      </div></td>
      </table></td>
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
				      <% }}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
                        <td width="93"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'returnList');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                    </table>
                  </div>





                      </td>
                    </tr>
                    <tr>
                      <td colspan="2" width="583" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明
                        : </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
 <ul>
                      	  
                      	  <li class="sbody" ><font color="#666666">確認輸入資料無誤後, 按【確定】即將本表上的資料, 於資料庫中建檔。</li>
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</li>
                          <li class="sbody" >交易金額欄位定義為存,提,轉出與轉入的總計。</li>	
                          <li class="sbody" >歷年磁條金融卡與晶片卡之發行張數及停卡張數與流通張數係表示截至申報月底日之歷年總張數。</li>	
                          <li class="sbody" >流通張數欄位定義為發行張數與停卡張數相減。</li>
                        </ul>
                      </td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</table>
</body>
