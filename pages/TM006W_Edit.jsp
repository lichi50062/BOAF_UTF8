<%
//105.10.03 create by2968
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
	List dbData = (request.getAttribute("dbData")==null)?null:(List)request.getAttribute("dbData");
	List EditData = (request.getAttribute("EditData")==null)?null:(List)request.getAttribute("EditData");
	List dbData_Pre = (request.getAttribute("dbData_Pre")==null)?null:(List)request.getAttribute("dbData_Pre");	//申報基準日以前的累計資料
	List accDivList01= (request.getAttribute("accDivList01")==null)?null:(List)request.getAttribute("accDivList01");
	List accDivList02= (request.getAttribute("accDivList02")==null)?null:(List)request.getAttribute("accDivList02");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	String acc_tr_type = "",acc_tr_name = "";
	if(accDivList01!=null && accDivList01.size()>0){
		acc_tr_type = String.valueOf(((DataObject)accDivList01.get(0)).getValue("acc_tr_type"));
		acc_tr_name = String.valueOf(((DataObject)accDivList01.get(0)).getValue("acc_tr_name"));
	}else if(accDivList02!=null && accDivList02.size()>0){
		acc_tr_type = String.valueOf(((DataObject)accDivList02.get(0)).getValue("acc_tr_type"));
		acc_tr_name = String.valueOf(((DataObject)accDivList02.get(0)).getValue("acc_tr_name"));
	}
	String applyType="",applyDate = "",applyDate_B = "",applyDate_E = "",cnt_Name = "",cnt_Tel ="";
	if(EditData!=null && EditData.size()>0){
		if("2".equals(String.valueOf(((DataObject)EditData.get(0)).getValue("applytype")))){
			applyType="2週";
		}else if("1".equals(String.valueOf(((DataObject)EditData.get(0)).getValue("applytype")))){
			applyType="週";
		}else{
			applyType="月";
		}
		applyDate = ((DataObject)EditData.get(0)).getValue("applydate")==null?"":String.valueOf(((DataObject)EditData.get(0)).getValue("applydate"));
		applyDate_B = ((DataObject)EditData.get(0)).getValue("applydate_b")==null?"":String.valueOf(((DataObject)EditData.get(0)).getValue("applydate_b"));
		applyDate_E = ((DataObject)EditData.get(0)).getValue("applydate_e")==null?"":String.valueOf(((DataObject)EditData.get(0)).getValue("applydate_e"));
		cnt_Name = ((DataObject)EditData.get(0)).getValue("cnt_name")==null?"":String.valueOf(((DataObject)EditData.get(0)).getValue("cnt_name"));
		cnt_Tel = ((DataObject)EditData.get(0)).getValue("cnt_tel")==null?"":String.valueOf(((DataObject)EditData.get(0)).getValue("cnt_tel"));
	}
	
	
	String apply_cnt = "0",apply_amt = "0",apply_bal = "0",apply_cnt_sum = "0",apply_amt_sum = "0",apply_bal_sum = "0";
	String appr_cnt = "0",appr_amt = "0",appr_bal = "0",appr_cnt_sum = "0",appr_amt_sum = "0",appr_bal_sum = "0";
	String nonappr_cnt = "0",nonappr_reason = "";
	boolean isDif = false;
	if(dbData_Pre.size()>0 && dbData.size()>0){
		for(int p=0;p<dbData_Pre.size();p++){
			DataObject pb = (DataObject)dbData_Pre.get(p);
			String acc_div_P = String.valueOf(pb.getValue("acc_div"));
			String acc_code_P = String.valueOf(pb.getValue("acc_code"));
			String apply_cnt_sum_P = String.valueOf(pb.getValue("apply_cnt_sum"));
			String apply_amt_sum_P = String.valueOf(pb.getValue("apply_amt_sum"));
			String apply_bal_sum_P = String.valueOf(pb.getValue("apply_bal_sum"));
			String appr_cnt_sum_P = String.valueOf(pb.getValue("appr_cnt_sum"));
			String appr_amt_sum_P = String.valueOf(pb.getValue("appr_amt_sum"));
			String appr_bal_sum_P = String.valueOf(pb.getValue("appr_bal_sum"));
			for(int s=0;s<dbData.size();s++){
				DataObject b = (DataObject)dbData.get(s);
				if(acc_div_P.equals(String.valueOf(b.getValue("acc_div")))){
					if(acc_code_P.equals(String.valueOf(b.getValue("acc_code")))){
						if( !apply_cnt_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("apply_cnt_sum").toString())-Integer.parseInt(b.getValue("apply_cnt").toString()))) ||
							!apply_amt_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("apply_amt_sum").toString())-Integer.parseInt(b.getValue("apply_amt").toString()))) ||
							!apply_bal_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("apply_bal_sum").toString())-Integer.parseInt(b.getValue("apply_bal").toString()))) ||
							!appr_cnt_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("appr_cnt_sum").toString())-Integer.parseInt(b.getValue("appr_cnt").toString()))) ||
							!appr_amt_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("appr_amt_sum").toString())-Integer.parseInt(b.getValue("appr_amt").toString()))) ||
							!appr_bal_sum_P.equals(String.valueOf(Integer.parseInt(b.getValue("appr_bal_sum").toString())-Integer.parseInt(b.getValue("appr_bal").toString())))){
								isDif = true;
						}
					}
				}
			}
		}
	}
	//取得TM006W的權限
  	Properties permission = ( session.getAttribute("TM006W")==null ) ? new Properties() : (Properties)session.getAttribute("TM006W"); 
  	if(permission == null){
         System.out.println("TM006W_Edit.permission == null");
      }else{
         System.out.println("TM006W_Edit.permission.size ="+permission.size());
                 
      }
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/TM006W.js"></script>

<script language="javascript" event="onresize" for="window"></script>
<head>
<title>適用協助措施之經辦機構申報作業</title>
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

function autoCalcul(form) {
	
		for(var i=0; i<document.getElementsByName("acc_Code01").length; i++){
			if(document.getElementsByName("apply_Cnt01")[i].value=="")document.getElementsByName("apply_Cnt01")[i].value="0";
			if(document.getElementsByName("apply_Amt01")[i].value=="")document.getElementsByName("apply_Amt01")[i].value="0";
			if(document.getElementsByName("apply_Bal01")[i].value=="")document.getElementsByName("apply_Bal01")[i].value="0";
			if(document.getElementsByName("apply_Cnt_Sum01_P")[i].value=="")document.getElementsByName("apply_Cnt_Sum01_P")[i].value="0";
			if(document.getElementsByName("apply_Amt_Sum01_P")[i].value=="")document.getElementsByName("apply_Amt_Sum01_P")[i].value="0";
			if(document.getElementsByName("apply_Bal_Sum01_P")[i].value=="")document.getElementsByName("apply_Bal_Sum01_P")[i].value="0";
			if(document.getElementsByName("appr_Cnt01")[i].value=="")document.getElementsByName("appr_Cnt01")[i].value="0";
			if(document.getElementsByName("appr_Amt01")[i].value=="")document.getElementsByName("appr_Amt01")[i].value="0";
			if(document.getElementsByName("appr_Bal01")[i].value=="")document.getElementsByName("appr_Bal01")[i].value="0";
			if(document.getElementsByName("appr_Cnt_Sum01_P")[i].value=="")document.getElementsByName("appr_Cnt_Sum01_P")[i].value="0";
			if(document.getElementsByName("appr_Amt_Sum01_P")[i].value=="")document.getElementsByName("appr_Amt_Sum01_P")[i].value="0";
			if(document.getElementsByName("appr_Bal_Sum01_P")[i].value=="")document.getElementsByName("appr_Bal_Sum01_P")[i].value="0";
			
			document.getElementsByName("apply_Cnt_Sum01_Show")[i].value = (parseInt(document.getElementsByName("apply_Cnt_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Cnt01")[i].value));
			document.getElementsByName("apply_Amt_Sum01_Show")[i].value = (parseInt(document.getElementsByName("apply_Amt_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Amt01")[i].value));
			document.getElementsByName("apply_Bal_Sum01_Show")[i].value = (parseInt(document.getElementsByName("apply_Bal_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Bal01")[i].value));
			document.getElementsByName("appr_Cnt_Sum01_Show")[i].value = (parseInt(document.getElementsByName("appr_Cnt_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Cnt01")[i].value));
			document.getElementsByName("appr_Amt_Sum01_Show")[i].value = (parseInt(document.getElementsByName("appr_Amt_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Amt01")[i].value));
			document.getElementsByName("appr_Bal_Sum01_Show")[i].value = (parseInt(document.getElementsByName("appr_Bal_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Bal01")[i].value));
		}  
		
	
		for(var i=0; i<document.getElementsByName("acc_Code02").length; i++){
			if(document.getElementsByName("apply_Cnt02")[i].value=="")document.getElementsByName("apply_Cnt02")[i].value="0";
			if(document.getElementsByName("apply_Amt02")[i].value=="")document.getElementsByName("apply_Amt02")[i].value="0";
			if(document.getElementsByName("apply_Cnt_Sum02_P")[i].value=="")document.getElementsByName("apply_Cnt_Sum02_P")[i].value="0";
			if(document.getElementsByName("apply_Amt_Sum02_P")[i].value=="")document.getElementsByName("apply_Amt_Sum02_P")[i].value="0";
			if(document.getElementsByName("appr_Cnt02")[i].value=="")document.getElementsByName("appr_Cnt02")[i].value="0";
			if(document.getElementsByName("appr_Amt02")[i].value=="")document.getElementsByName("appr_Amt02")[i].value="0";
			if(document.getElementsByName("appr_Bal02")[i].value=="")document.getElementsByName("appr_Bal02")[i].value="0";
			if(document.getElementsByName("appr_Cnt_Sum02_P")[i].value=="")document.getElementsByName("appr_Cnt_Sum02_P")[i].value="0";
			if(document.getElementsByName("appr_Amt_Sum02_P")[i].value=="")document.getElementsByName("appr_Amt_Sum02_P")[i].value="0";
			if(document.getElementsByName("appr_Bal_Sum02_P")[i].value=="")document.getElementsByName("appr_Bal_Sum02_P")[i].value="0";
			
			document.getElementsByName("apply_Cnt_Sum02_Show")[i].value = (parseInt(document.getElementsByName("apply_Cnt_Sum02_P")[i].value)+parseInt(document.getElementsByName("apply_Cnt02")[i].value));
			document.getElementsByName("apply_Amt_Sum02_Show")[i].value = (parseInt(document.getElementsByName("apply_Amt_Sum02_P")[i].value)+parseInt(document.getElementsByName("apply_Amt02")[i].value));
			document.getElementsByName("appr_Cnt_Sum02_Show")[i].value = (parseInt(document.getElementsByName("appr_Cnt_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Cnt02")[i].value));
			document.getElementsByName("appr_Amt_Sum02_Show")[i].value = (parseInt(document.getElementsByName("appr_Amt_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Amt02")[i].value));
			document.getElementsByName("appr_Bal_Sum02_Show")[i].value = (parseInt(document.getElementsByName("appr_Bal_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Bal02")[i].value));
            
		}
	
}

//-->
</script>
<style>  
	input{background:expression((this.disabled && this.disabled==true)?"#FFFFFF":"#FFFFFF") }  
</style> 

</head>
<body leftmargin="15" topmargin="0">
<Form name='form' method=post action='/pages/TM006W.jsp' >
<input type='hidden' name='loan_Sbm'>
<input type='hidden' name='acc_Tr_Type' value='<%=acc_tr_type%>'>
<table width="99%" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="618"><img src="images/space_1.gif" width="12" ></td>
      </tr>
      <tr>
          <td bgcolor="#FFFFFF" width="1200">
      	  <table width="1200" border="0" align="center" cellpadding="0" cellspacing="0" >      	  
              <tr>
                <td width="1200" >
                <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="342" align="center"><img src="images/banner_bg1.gif" width="317" height="17" align="left"></td>
                      <td width="443" align="center"><b>
                        <center><font color="#000000" size="4">適用協助措施之經辦機構申報作業</font>
                        </center>
                        </b> </td>
                      <td width="404" valign="middle"><img src="images/banner_bg1.gif" width="375" height="17" align="right"></td>
                    </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td width="1200" >
                <table width="1193" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr><td width="1193" ></tr>
                    <tr> 
                      <div align="cenger"><jsp:include page="getLoginUser.jsp?width=1191" flush="true" /></div> 
                    </tr>   
                    <tr>
            		<td width="1193" >
					<table width="1192"  align="center" cellpadding="1" cellspacing="1" border="1" bordercolor="#3A9D99" >
            		<tr>
            			<td class="sbody" width="380" bgcolor='#D8EFEE' align="left" colspan="3">
							協助措施名稱</td>
            			<td class="sbody" width="662" bgcolor="e7e7e7" colspan="14" >
							<%=acc_tr_name %></td>
            		</tr>
            	    <tr class="sbody">
            			<td width="380" bgcolor='#D8EFEE' align="left" colspan="3">
							申報基準日</td>
            			<td width="810" bgcolor="e7e7e7" colspan="14">
                            <%=applyDate.length()>=10?Utility.getCHTdate(applyDate.substring(0,10),0):applyDate%> 
                            <input type='hidden' name='applyDate' id='applyDate' value='<%=applyDate.substring(0,10) %>' >
                           <font color="red">&nbsp;&nbsp;&nbsp; 
							申報資料範圍:<%=applyDate_B.length()>=10?Utility.getCHTdate(applyDate_B.substring(0,10),0):applyDate_B %>~
									 <%=applyDate_E.length()>=10?Utility.getCHTdate(applyDate_E.substring(0,10),0):applyDate_E %>&nbsp;&nbsp;&nbsp; 
							<%if(isDif){ %>＊該申報基準日前有異動資料,請點選[修改]按鈕,協助更新申請累計、核准累計資料<%} %></font>
            		</tr>
            		<tr class="sbody">
            			<td width='380' bgcolor="#D8EFEE" align="left" colspan="3">
							聯絡人員</td>
            			<td width='810' bgcolor='e7e7e7' colspan="14">
                            <input type='text' name='cnt_Name' id='cnt_Name' value='<%=cnt_Name %>' size='36' maxlength='130' >
                        </td>
            		</tr>
            		<tr class="sbody">
            			<td width='380' bgcolor='#D8EFEE' align='left' colspan="3">
							聯絡電話</td>
            			<td width='810' bgcolor='e7e7e7' colspan="14">
             	 			<input type='text' name='cnt_Tel' id='cnt_Tel' value='<%=cnt_Tel %>' size='37' maxlength='130' >&nbsp;&nbsp;&nbsp; 
             	 		</td>
            		</tr>
            		
					<%if(accDivList01!=null && accDivList01.size()>0){ %>
					<tr class="sbody">
            			<td width='330' bgcolor='#D8EFEE' align='left' colspan="2">
							<p align="center">舊貸展延需求</td>
            			<td width='170' bgcolor='#D8EFEE' colspan="3" align="center">
				         	本<%=applyType %>申請</td>
				        <td width='170' bgcolor='#D8EFEE' colspan="3" align="center">
				                                    申請累計</td>
				        <td width='170' bgcolor='#D8EFEE' colspan="3" align="center">
				                                    本<%=applyType %>核准</td>
				        <td width='170' bgcolor='#D8EFEE' colspan="3" align="center">
				                                    核准累計</td>
				        <td width='170' bgcolor='#D8EFEE' colspan="2" align="center">
				                                    不予核貸</td>
            	    </tr>
            	    <tr class="sbody">
            			<td width='50' bgcolor="#D8EFEE" align='left' rowspan="<%=accDivList01.size()+1%>">
						舊貸展延需求</td>
            			<td width='280' bgcolor='#D8EFEE' align='left' >貸款種類</td>
            			<td width='50' bgcolor='e7e7e7'>件數</td>
					    <td width='60' bgcolor='e7e7e7'>貸款金額</td>
					    <td width='60' bgcolor='e7e7e7'>貸款餘額</td>
					    <td width='50' bgcolor='e7e7e7'>件數</td>
					    <td width='60' bgcolor='e7e7e7'>貸款金額</td>
					    <td width='60' bgcolor='e7e7e7'>貸款餘額</td>
					    <td width='50' bgcolor='e7e7e7'>件數</td>
					    <td width='60' bgcolor='e7e7e7'>貸款金額</td>
					    <td width='60' bgcolor='e7e7e7'>貸款餘額</td>
					    <td width='50' bgcolor='e7e7e7'>件數</td>
					    <td width='60' bgcolor='e7e7e7'>貸款金額</td>
					    <td width='60' bgcolor='e7e7e7'>貸款餘額</td>
					    <td width='50' bgcolor='e7e7e7'>件數</td>
					    <td width='120' bgcolor='e7e7e7'>原因</td>
            	    </tr>
						<%for(int i=0;i<accDivList01.size();i++){ 
							String acc_code01 = String.valueOf(((DataObject)accDivList01.get(i)).getValue("acc_code"));
							%>
							<tr class="sbody">
		            			<td width='280' bgcolor='#D8EFEE' align='left' >
		            				<input type='hidden' name='acc_Code01' id='acc_Code01' value='<%=acc_code01 %>'>
									<%=String.valueOf(((DataObject)accDivList01.get(i)).getValue("acc_name")) %></td>
									<%
									String apply_cnt_sum_P="0",apply_amt_sum_P="0",apply_bal_sum_P="0";
									String appr_cnt_sum_P="0",appr_amt_sum_P="0",appr_bal_sum_P="0";
									apply_cnt = "0";apply_amt = "0";apply_bal = "0";
									apply_cnt_sum = "0";apply_amt_sum = "0";apply_bal_sum = "0";
									appr_cnt = "0";appr_amt = "0";appr_bal = "0";
									appr_cnt_sum = "0";appr_amt_sum = "0";appr_bal_sum = "0";
									nonappr_cnt = "0";nonappr_reason = "";
									
									if(dbData_Pre.size()>0){
										for(int p=0;p<dbData_Pre.size();p++){
											DataObject pb = ((DataObject)dbData_Pre.get(p));
											if("01".equals(String.valueOf(pb.getValue("acc_div")))){
												String acc_code_P = String.valueOf(pb.getValue("acc_code"));
												if(acc_code01.equals(acc_code_P)){
													apply_cnt_sum_P = String.valueOf(pb.getValue("apply_cnt_sum"));
													apply_amt_sum_P = String.valueOf(pb.getValue("apply_amt_sum"));
													apply_bal_sum_P = String.valueOf(pb.getValue("apply_bal_sum"));
													appr_cnt_sum_P = String.valueOf(pb.getValue("appr_cnt_sum"));
													appr_amt_sum_P = String.valueOf(pb.getValue("appr_amt_sum"));
													appr_bal_sum_P = String.valueOf(pb.getValue("appr_bal_sum"));
													apply_cnt_sum = apply_cnt_sum_P;
													apply_amt_sum = apply_amt_sum_P;
													apply_bal_sum = apply_bal_sum_P;
													appr_cnt_sum = appr_cnt_sum_P;
													appr_amt_sum = appr_amt_sum_P;
													appr_bal_sum = appr_bal_sum_P;
												}
											}
										}
									}
									if(dbData!=null && dbData.size()>0){
										for(int s=0;s<dbData.size();s++){
											DataObject b = (DataObject)dbData.get(s);
											if("01".equals(String.valueOf(b.getValue("acc_div")))){
												if(acc_code01.equals(String.valueOf(b.getValue("acc_code")))){
													apply_cnt = b.getValue("apply_cnt")==null?"0":String.valueOf(b.getValue("apply_cnt"));
													apply_amt = b.getValue("apply_amt")==null?"0":String.valueOf(b.getValue("apply_amt"));
													apply_bal = b.getValue("apply_bal")==null?"0":String.valueOf(b.getValue("apply_bal"));
													apply_cnt_sum = b.getValue("apply_cnt_sum")==null?"0":String.valueOf(b.getValue("apply_cnt_sum"));
													apply_amt_sum = b.getValue("apply_amt_sum")==null?"0":String.valueOf(b.getValue("apply_amt_sum"));
													apply_bal_sum = b.getValue("apply_bal_sum")==null?"0":String.valueOf(b.getValue("apply_bal_sum"));
													appr_cnt = b.getValue("appr_cnt")==null?"0":String.valueOf(b.getValue("appr_cnt"));
													appr_amt = b.getValue("appr_amt")==null?"0":String.valueOf(b.getValue("appr_amt"));
													appr_bal = b.getValue("appr_bal")==null?"0":String.valueOf(b.getValue("appr_bal"));
													appr_cnt_sum = b.getValue("appr_cnt_sum")==null?"0":String.valueOf(b.getValue("appr_cnt_sum"));
													appr_amt_sum = b.getValue("appr_amt_sum")==null?"0":String.valueOf(b.getValue("appr_amt_sum"));
													appr_bal_sum = b.getValue("appr_bal_sum")==null?"0":String.valueOf(b.getValue("appr_bal_sum"));
													nonappr_cnt = b.getValue("nonappr_cnt")==null?"0":String.valueOf(b.getValue("nonappr_cnt"));
													nonappr_reason = b.getValue("nonappr_reason")==null?"":String.valueOf(b.getValue("nonappr_reason"));
												}
											}
										}
									}
									
									
									%>
		            			<td width='50' bgcolor='e7e7e7' >
             					 	<input type='text' name='apply_Cnt01' id='apply_Cnt01' value="<%=apply_cnt %>" size='3' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='text' name='apply_Amt01' id='apply_Amt01' value="<%=apply_amt %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='text' name='apply_Bal01' id='apply_Bal01' value="<%=apply_bal%>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='50' bgcolor='e7e7e7' >
                            		<input type='hidden' name='apply_Cnt_Sum01_P' id='apply_Cnt_Sum01_P' value='<%=apply_cnt_sum_P %>'>
             	 					<input type='text' name='apply_Cnt_Sum01_Show' value="<%=apply_cnt_sum %>" size='3' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='hidden' name='apply_Amt_Sum01_P' id='apply_Amt_Sum01_P' value='<%=apply_amt_sum_P %>'>
             	 					<input type='text' name='apply_Amt_Sum01_Show' value="<%=apply_amt_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='hidden' name='apply_Bal_Sum01_P' id='apply_Bal_Sum01_P' value='<%=apply_bal_sum_P %>'>
             	 					<input type='text' name='apply_Bal_Sum01_Show' value="<%=apply_bal_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='50' bgcolor='e7e7e7' >
             	 					<input type='text' name='appr_Cnt01' id='appr_Cnt01' value="<%=appr_cnt %>" size='3' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='text' name='appr_Amt01' id='appr_Amt01' value="<%=appr_amt %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='text' name='appr_Bal01' id='appr_Bal01' value="<%=appr_bal %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='50' bgcolor='e7e7e7' >
             	 					<input type='hidden' name='appr_Cnt_Sum01_P' id='appr_Cnt_Sum01_P' value='<%=appr_cnt_sum_P %>'>
             	 					<input type='text' name='appr_Cnt_Sum01_Show' value="<%=appr_cnt_sum %>" size='3' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='hidden' name='appr_Amt_Sum01_P' id='appr_Amt_Sum01_P' value='<%=appr_amt_sum_P %>'>
             	 					<input type='text' name='appr_Amt_Sum01_Show' value="<%=appr_amt_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='e7e7e7' >
             	 					<input type='hidden' name='appr_Bal_Sum01_P' id='appr_Bal_Sum01_P' value='<%=appr_bal_sum_P %>'>
             	 					<input type='text' name='appr_Bal_Sum01_Show' value="<%=appr_bal_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='50' bgcolor='e7e7e7' >
             	 					<input type='text' name='nonappr_Cnt01' id='nonappr_Cnt01' value="<%=nonappr_cnt %>" size='2'  style='text-align: right;'></td>
            					<td width='120' bgcolor='e7e7e7' >
									<input type='text' name='nonappr_Reason01' id='nonappr_Reason01' value='<%=nonappr_reason %>' size='12' maxlength='300' ></td>
		            	    </tr>
	            	    <%} %>
            	    <%} %>
            	   

					<%if(accDivList02!=null && accDivList02.size()>0){ %>
					<tr class="sbody">
						<td width='330' bgcolor='#BDDE9C' align='left' colspan="2">
							<p align="center">新貸需求</td>
            			<td width='170' bgcolor='#BDDE9C' colspan="3" align="center">
                            	本<%=applyType %>申請</td>
            			<td width='170' bgcolor='#BDDE9C' colspan="3" align="center">
                            	申請累計</td>
            			<td width='170' bgcolor='#BDDE9C' colspan="3" align="center">
                            	本<%=applyType %>核准</td>
            			<td width='170' bgcolor='#BDDE9C' colspan="3" align="center">
                            	核准累計</td>
            			<td width='170' bgcolor='#BDDE9C' colspan="2" align="center">
                           	 	不予核貸</td>
            	    </tr>
            	    <tr class="sbody">
            			<td width='50' bgcolor="#BDDE9C" rowspan="<%=accDivList02.size()+1%>" >
						新貸需求</td>
            			<td width='280' bgcolor="#BDDE9C">貸款種類</td>
            			<td width='50' bgcolor='#EBF4E1'>件數</td>
            			<td width='120' bgcolor='#EBF4E1' colspan="2">貸款金額</td>
            			<td width='50' bgcolor='#EBF4E1'>件數</td>
            			<td width='120' bgcolor='#EBF4E1' colspan="2">貸款金額</td>
            			<td width='50' bgcolor='#EBF4E1'>件數</td>
            			<td width='60' bgcolor='#EBF4E1'>貸款金額</td>
            			<td width='60' bgcolor='#EBF4E1'>貸款餘額</td>
            			<td width='50' bgcolor='#EBF4E1'>件數</td>
            			<td width='60' bgcolor='#EBF4E1'>貸款金額</td>
            			<td width='60' bgcolor='#EBF4E1'>貸款餘額</td>
            			<td width='49' bgcolor='#EBF4E1'>件數</td>
            			<td width='119' bgcolor='#EBF4E1'>原因</td>
            	    </tr>
						<%for(int i=0;i<accDivList02.size();i++){ 
							String acc_code02 = String.valueOf(((DataObject)accDivList02.get(i)).getValue("acc_code"));
							%>
							<tr class="sbody">
		            			<td width='280' bgcolor='#BDDE9C' align='left' >
		            				<input type='hidden' name='acc_Code02' id='acc_Code02' value='<%=acc_code02 %>'>
									<%=String.valueOf(((DataObject)accDivList02.get(i)).getValue("acc_name")) %></td>
									<%
									String apply_cnt_sum_P="0",apply_amt_sum_P="0",apply_bal_sum_P="0";
									String appr_cnt_sum_P="0",appr_amt_sum_P="0",appr_bal_sum_P="0";
									apply_cnt = "0";apply_amt = "0";apply_bal = "0";
									apply_cnt_sum = "0";apply_amt_sum = "0";apply_bal_sum = "0";
									appr_cnt = "0";appr_amt = "0";appr_bal = "0";
									appr_cnt_sum = "0";appr_amt_sum = "0";appr_bal_sum = "0";
									nonappr_cnt = "0";nonappr_reason = "";
									if(dbData_Pre!=null && dbData_Pre.size()>0){
										for(int p=0;p<dbData_Pre.size();p++){
											DataObject pb = ((DataObject)dbData_Pre.get(p));
											if("02".equals(String.valueOf(pb.getValue("acc_div")))){
												String acc_code_P = String.valueOf(pb.getValue("acc_code"));
												if(acc_code02.equals(acc_code_P)){
													apply_cnt_sum_P = String.valueOf(pb.getValue("apply_cnt_sum"));
													apply_amt_sum_P = String.valueOf(pb.getValue("apply_amt_sum"));
													apply_bal_sum_P = String.valueOf(pb.getValue("apply_bal_sum"));
													appr_cnt_sum_P = String.valueOf(pb.getValue("appr_cnt_sum"));
													appr_amt_sum_P = String.valueOf(pb.getValue("appr_amt_sum"));
													appr_bal_sum_P = String.valueOf(pb.getValue("appr_bal_sum"));
													apply_cnt_sum = apply_cnt_sum_P;
													apply_amt_sum = apply_amt_sum_P;
													apply_bal_sum = apply_bal_sum_P;
													appr_cnt_sum = appr_cnt_sum_P;
													appr_amt_sum = appr_amt_sum_P;
													appr_bal_sum = appr_bal_sum_P;
												}
											}
										}
									}
									if(dbData!=null && dbData.size()>0){
										for(int s=0;s<dbData.size();s++){
											DataObject b =(DataObject)dbData.get(s);
											if("02".equals(String.valueOf(b.getValue("acc_div")))){
												if(acc_code02.equals(String.valueOf(b.getValue("acc_code")))){
													apply_cnt = b.getValue("apply_cnt")==null?"0":String.valueOf(b.getValue("apply_cnt"));
													apply_amt = b.getValue("apply_amt")==null?"0":String.valueOf(b.getValue("apply_amt"));
													apply_bal = b.getValue("apply_bal")==null?"0":String.valueOf(b.getValue("apply_bal"));
													apply_cnt_sum = b.getValue("apply_cnt_sum")==null?"0":String.valueOf(b.getValue("apply_cnt_sum"));
													apply_amt_sum = b.getValue("apply_amt_sum")==null?"0":String.valueOf(b.getValue("apply_amt_sum"));
													apply_bal_sum = b.getValue("apply_bal_sum")==null?"0":String.valueOf(b.getValue("apply_bal_sum"));
													appr_cnt = b.getValue("appr_cnt")==null?"0":String.valueOf(b.getValue("appr_cnt"));
													appr_amt = b.getValue("appr_amt")==null?"0":String.valueOf(b.getValue("appr_amt"));
													appr_bal = b.getValue("appr_bal")==null?"0":String.valueOf(b.getValue("appr_bal"));
													appr_cnt_sum = b.getValue("appr_cnt_sum")==null?"0":String.valueOf(b.getValue("appr_cnt_sum"));
													appr_amt_sum = b.getValue("appr_amt_sum")==null?"0":String.valueOf(b.getValue("appr_amt_sum"));
													appr_bal_sum = b.getValue("appr_bal_sum")==null?"0":String.valueOf(b.getValue("appr_bal_sum"));
													nonappr_cnt = b.getValue("nonappr_cnt")==null?"0":String.valueOf(b.getValue("nonappr_cnt"));
													nonappr_reason = b.getValue("nonappr_reason")==null?"":String.valueOf(b.getValue("nonappr_reason"));
												}
											}
										}
									}
									
									%>
		            			<td width='50' bgcolor='#EBF4E1' >
             					 	<input type='text' name='apply_Cnt02' id='apply_Cnt02' value="<%=apply_cnt %>" size='3' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='120' bgcolor='#EBF4E1' colspan="2">
             	 					<input type='text' name='apply_Amt02' id='apply_Amt02' value="<%=apply_amt %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);">
             	 					<input type='hidden' name='apply_Bal02' id='apply_Bal02' value="<%=apply_bal %>" size='6' maxlength='14' style='text-align: right;'>
             	 					</td>
            					<td width='50' bgcolor='#EBF4E1' >
             	 					<input type='text' name='apply_Cnt_Sum02_Show' value="<%=apply_cnt_sum %>" size='3' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;">
             	 					<input type='hidden' name='apply_Cnt_Sum02_P' id='apply_Cnt_Sum02_P' value='<%=apply_cnt_sum_P %>'></td>
            					<td width='120' bgcolor='#EBF4E1' colspan="2">
            						<input type='text' name='apply_Amt_Sum02_Show' value="<%=apply_amt_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;">
             	 					<input type='hidden' name='apply_Amt_Sum02_P' id='apply_Amt_Sum02_P' value='<%=apply_amt_sum_P %>'>
             	 					<input type='hidden' name='apply_Bal_Sum02_P' id='apply_Bal_Sum02_P' value='<%=apply_bal_sum_P %>'>
             	 					</td>
            					<td width='50' bgcolor='#EBF4E1' >
             	 					<input type='text' name='appr_Cnt02' id='appr_Cnt02' value="<%=appr_cnt %>" size='3' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='#EBF4E1' >
             	 					<input type='text' name='appr_Amt02' id='appr_Amt02' value="<%=appr_amt %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='60' bgcolor='#EBF4E1' >
             	 					<input type='text' name='appr_Bal02' id='appr_Bal02' value="<%=appr_bal %>" size='6' maxlength='14' style='text-align: right;' onBlur="autoCalcul(form);"></td>
            					<td width='50' bgcolor='#EBF4E1' >
             	 					<input type='hidden' name='appr_Cnt_Sum02_P' id='appr_Cnt_Sum02_P' value='<%=appr_cnt_sum_P %>'>
             	 					<input type='text' name='appr_Cnt_Sum02_Show' value="<%=appr_cnt_sum %>" size='3' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='#EBF4E1' >
             	 					<input type='hidden' name='appr_Amt_Sum02_P' id='appr_Amt_Sum02_P' value='<%=appr_amt_sum_P %>'>
             	 					<input type='text' name='appr_Amt_Sum02_Show' value="<%=appr_amt_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='60' bgcolor='#EBF4E1' >
             	 					<input type='hidden' name='appr_Bal_Sum02_P' id='appr_Bal_Sum02_P' value='<%=appr_bal_sum_P %>'>
             	 					<input type='text' name='appr_Bal_Sum02_Show' value="<%=appr_bal_sum %>" size='6' style='text-align: right;' readOnly onKeydown="javascript:if(window.event.keyCode==8) return false;" style="color:#A4A4A4;"></td>
            					<td width='49' bgcolor='#EBF4E1' >
             	 					<input type='text' name='nonappr_Cnt02' id='nonappr_Cnt02' value="<%=nonappr_cnt %>" size='2'  style='text-align: right;'></td>
            					<td width='119' bgcolor='#EBF4E1' >
									<input type='text' name='nonappr_Reason02' id='nonappr_Reason02' value='<%=nonappr_reason %>' size='13' maxlength='300' ></td>
		            	    </tr>
	            	    <%} %>
            	    <%} %>
            	        	
            	</Table>
   			</td>
 		</tr>

      </table></td>
  </tr>
   <tr>
                <td width="1196" >
				<table width="734" border="0" cellpadding="1" cellspacing="1" class="sbody" >
                    <tr>
                      <td colspan="2" width="684" >
                      <div align="center">
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr align=center>
			<td >
			<%if(act.equals("new")){%>
									<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ %> 
										<a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
										<a href="javascript:AskReset(form);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
									<%}%>
								<%}%>
								<%if(act.equals("Edit")){%>
									<%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>
										 <a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
											        		<!-- <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td> -->
									<%}%>
									<%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>
										 <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
									<%}%>
								<%}%>
							    <a href="javascript:doSubmit(form,'List');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
	    		</td> 
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
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>              
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>              
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>              
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</font></font></font><font color="red"> </li>   
                        </ul>
                            </font>
                            </font>
                          </font>
 </font>
                      </td>
                    </tr>
                  </table></td>
      </tr>          
	</table>
</table>
</form>
</body>
