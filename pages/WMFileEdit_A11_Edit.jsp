<%
//102.01.08 add by 2968
//104.05.12 add 增加A111111111可調整案件編號及分項項數 by 2295
//104.05.12 fix 補舊案件資料時,新增下一筆參貸項目時,該案件編號會是空值 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>
<%
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
	   System.out.println("WMFileEdit_A11.permission == null");
	}else{
	   System.out.println("WMFileEdit_A11.permission.size ="+permission.size());
	}
	//接收資料
	List A11_S_Edit = (List)request.getAttribute("A11_S_Edit");
	List A11_NoApply_Edit = (List)request.getAttribute("A11_NoApply_Edit");
	String A11_Lock = (String)request.getAttribute("A11_Lock");    //檢查有無鎖定
	String act = (String)request.getAttribute("act");
	String isLastMonthData = (String)request.getAttribute("isLastMonthData"); //上個月是否有申報資料
	String isHaveApplyData = (String)request.getAttribute("isHaveApplyData"); //本月是否已有申報資料
	String isHaveNoApplyData = (String)request.getAttribute("isHaveNoApplyData"); //本月是否已有申報過無資料
	String isHaveNoApplyLoanData = (String)request.getAttribute("isHaveNoApplyLoanData"); //本月是否已有申報過無資料但有實際授信餘額
	String firstLoan_amt_sum = (String)request.getAttribute("firstLoan_amt_sum"); //同一筆但不同分項的授信案總金額
	String maxCase_No= (String)request.getAttribute("maxCase_No"); //該件聯貸之最大項
	List bankNoMaxList = (List)request.getAttribute("bankNoMaxList"); //主辦行選單
	List loanKindList = (List)request.getAttribute("loanKindList"); //參貸型式選單
	List loanTypeList = (List)request.getAttribute("loanTypeList"); //信用部參貸部分之授信用途選單
	List payStateList = (List)request.getAttribute("payStateList"); //目前放款繳息情形選單
	List violateTypeList = (List)request.getAttribute("violateTypeList"); //有無違反契約承諾條款選單
	String title = ("New".equals(act))?"新增":"維護";
	String bank_no =  ( request.getParameter("bank_no")==null ) ? " " : (String)request.getParameter("bank_no");
	String case_no =  ( request.getParameter("case_no")==null ) ? " " : (String)request.getParameter("case_no");
	String case_cnt =  ( request.getParameter("case_cnt")==null ) ? " " : (String)request.getParameter("case_cnt");
	String s_year = (  request.getParameter("s_year")==null ) ? " " : (String) request.getParameter("s_year");
	String s_month = (  request.getParameter("s_month")==null ) ? " " : (String) request.getParameter("s_month");
	String hyear="";
	String hmonth="";
	String seq_no = "";
	DataObject Bean = new DataObject(); //使用dataobject type 來做資料傳遞
	DataObject NoApplyBean = new DataObject();
	Date today = new Date();
	if(A11_S_Edit == null){
	   System.out.println("A11_S_Edit == null");
	}else{
	    if(A11_S_Edit.size() != 0){
		   //使用dataobject type 來做資料傳遞
		   Bean = (DataObject)A11_S_Edit.get(0);//把選取的資料取出來
		   /*if(!"".equals(Bean.getValue("loan_amt_sum"))){
		       firstLoan_amt_sum = String.valueOf(Bean.getValue("loan_amt_sum"));
			}*/
	    }
	   System.out.println("A11_S_Edit.size()="+A11_S_Edit.size());
	}
	if(A11_NoApply_Edit == null){
		   System.out.println("A11_NoApply_Edit == null");
	}else{
		 if(A11_NoApply_Edit.size() != 0){
			 //使用dataobject type 來做資料傳遞
			 NoApplyBean = (DataObject)A11_NoApply_Edit.get(0);//把選取的資料取出來
		 }
		System.out.println("A11_NoApply_Edit.size()="+A11_NoApply_Edit.size());
	}
    //判斷申報年月若已鎖定，將返回到list畫面
 	if("Y".equals(A11_Lock) && "New".equals(act)){//該年月申報資料已鎖定,無法再異動該申報資料
 		out.print("<script>alert(\"該年月申報資料已鎖定,無法再異動該申報資料\");history.go(-1);</script>");
 	}
 	
%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
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
-->
</script>
<script language="javascript" src="js/FileEdit_A11.js"> </script>
<script language="javascript" src="js/Common.js"></script>

<head>
<title> 聯合貸款案件資料表 申請情形<%=title%> </title>
</head>


<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method="post" >

<table border="0" cellspacing="1" width="600">
  
  <tr>
    <td width="600"></td>
  </tr>
  
  <tr>
  <center>
    <td width="600">
      <table border="0" cellspacing="1">
        <tr>
          <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
          <td width="380"><p align="center"><font color='#000000' size=4><b>聯合貸款案件資料表 申請情形<%=title%></b></font></p></td>
          <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
        </tr>
      </table> 
    </td>
  </center>
  </tr>
  
  <tr> 
     <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
  </tr>

  <!-- ---------------------------------------------------------------------------------------------------------------------------- -->
  <tr>
    <td width="600">
    <center>    
    <table class="sbody"  border="1" cellspacing="1" bordercolor="#3A9D99" width="600" >
    	<tr>
    		<td width='239' bgcolor="#D8EFEE" colspan="3">金融機構代號</td>
			<% 
				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
					bank_no = (String)Bean.getValue("bank_no");
				}else if(A11_NoApply_Edit != null && A11_NoApply_Edit.size() != 0){
					bank_no = (String)NoApplyBean.getValue("bank_no");
				}
			%>
			<td width='398' bgcolor='e7e7e7' height="27"><%=bank_no%></td>
        </tr>
        <tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3">申報年月</td>
			<td width='398' bgcolor='e7e7e7' height="10">
               	<input type='text' name='m_year'
               			value="<%
               			 	if(A11_S_Edit != null && A11_S_Edit.size() != 0){
               			 		out.print(Bean.getValue("m_year"));
               			 		hyear=Bean.getValue("m_year").toString();
               			 	}else if(A11_NoApply_Edit != null && A11_NoApply_Edit.size() != 0){
	               			 	out.print(NoApplyBean.getValue("m_year"));
	           			 		hyear=NoApplyBean.getValue("m_year").toString();
               			 	}else{
                           		out.print(s_year);
               			 		hyear = s_year;
               			 	}
               			 	%>"
               			size='3' maxlength='3' onblur='CheckYear(this)' disabled >
               	<font color='#000000'>年
             	<select id="hide2" name='m_month' disabled >
             		<%
		            	if(A11_S_Edit != null && A11_S_Edit.size() != 0){
		            		out.print("<option value=\""+Bean.getValue("m_month")+"\">"+Bean.getValue("m_month")+"</option>");
		            		hmonth=Bean.getValue("m_month").toString();
		            	}else if(A11_NoApply_Edit != null && A11_NoApply_Edit.size() != 0){
		            		out.print("<option value=\""+NoApplyBean.getValue("m_month")+"\">"+NoApplyBean.getValue("m_month")+"</option>");
		            		hmonth=NoApplyBean.getValue("m_month").toString();
		            	}else{
		            	 	out.print("<option value=\""+s_month+"\">"+s_month+"</option>");
	               			hmonth = s_month;
		            	}
		            %>
        			
        		</select>月&nbsp;&nbsp;</font>

<!-- -----------------------------hide type---------------------------- -->
<input type='hidden' name='isNoApplyData'>
<input type='hidden' name='hyear' value="<%=hyear%>" >
<input type='hidden' name='hmonth' value="<%=hmonth%>">
<input type='hidden' name='hbank_no' value="<%=bank_no%>">
<%
		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
			seq_no = Bean.getValue("seq_no").toString();
			case_cnt=(Bean.getValue("case_cnt") != null)?Bean.getValue("case_cnt").toString():"";
			
%>
<%
	 }
%>
<input type='hidden' name='hseq_no' value="<%=seq_no%>" >
<input type='hidden' name='case_cnt' value="<%=case_cnt%>" >
<input type='hidden' name='firstLoan_amt_sum' value="<%=firstLoan_amt_sum%>" >

<!-- -----------------------------end of hide type---------------------------- -->
			</td>
        </tr>
        
        <%if(((String)session.getAttribute("muser_id")).equals("A111111111")){%>     
        <tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3">更改申報編號</td>
			<td width='398' bgcolor='e7e7e7' height="10">			    
               	案件編號:<input type='text' name='upd_case_no' value="<%=case_no%>" size=9>
               	分項:<input type='text' name='upd_case_cnt' value="<%=case_cnt%>" size=2>
               	<input type="button" value="更新案件編號" onclick="javascript:updateCase_No(this.document.forms[0],'<%=hyear%>','<%=hmonth%>','<%=bank_no%>','<%=seq_no%>');">               	
            </td>  	
        </tr>			
	    <%}%>
        <tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3">申報編號</td>
			<td width='398' bgcolor='e7e7e7' height="10">			    
               	<input type='hidden' name='case_no' value="<%
             				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
             				    case_no = (Bean.getValue("case_no")==null)?"":Bean.getValue("case_no").toString();
             				}
               			 	out.print(case_no);
               		       	%>" >
               		       	<% 
               		       	if("Edit".equals(act)){
	               		       	if(A11_S_Edit != null && A11_S_Edit.size() != 0 ){
	               		       		out.print(case_no);
	    		            	}else{
	    		            	    out.print("<font color='red'>本月無新增聯貸案件申報資料</font>");
	    		            	}
    		            	}else{
    		            	    out.print(case_no);
    		            	}%>
    		     &nbsp&nbsp&nbsp
    		     
    		     <% if(A11_S_Edit != null && A11_S_Edit.size() != 0){ 
	    		     	if(Bean.getValue("case_no") != null && (case_no).equals(maxCase_No) && (Integer.parseInt(maxCase_No.substring(7, 9))) != Integer.parseInt(case_cnt) ){ %>
	    		     		<input type="button" value="新增下一筆參貸項目" onclick="javascript:continueEditCnt(this.document.forms[0],'<%=hyear%>','<%=hmonth%>','<%=bank_no%>','<%=case_no%>');">
						<%}
    		     	}%>
			</td>
        </tr>
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">借款人統一編號</td>
			<td width='398' bgcolor='e7e7e7' >
				<input type='text' name='loan_idn'
             			value="<%
             				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
             				    String loan_idn = (Bean.getValue("loan_idn")==null)?"":Bean.getValue("loan_idn").toString();
               					out.print(loan_idn);
               			 	}
               		       	%>" maxlength='10' style='text-align: left;'> 
            	<font color='red' size=2 >*</font>
			</td>
        </tr>
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">借款人名稱</td>
			<td width='398' bgcolor='e7e7e7' >
				<input type='text' name='loan_name'
             			value="<%
             				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
             				    String loan_name = (Bean.getValue("loan_name")==null)?"":Bean.getValue("loan_name").toString();
               					out.print(loan_name);
               			 	}
               		       	%>" size='20' maxlength='100' style='text-align: left;'> 
				<font color='red' size=2 >*</font>
			</td>
        </tr>
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">授信案總金額(元)</td>
			<td width='398' bgcolor='e7e7e7' >
				<%if(Bean.getValue("loan_amt_sum")==null||Bean.getValue("loan_amt_sum")!=null && "0".equals(Bean.getValue("loan_amt_sum"))){ %>
             		<input align='right' type='text' name='loan_amt_sum' value="" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>   
               	<%	
               	}else{ %>
               		<input align='right' type='text' name='loan_amt_sum' value="<%=Utility.setCommaFormat((Bean.getValue("loan_amt_sum")) == null ? "":(Bean.getValue("loan_amt_sum")).toString())%>" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>
               	<%  if("Edit".equals(act)){firstLoan_amt_sum=(Bean.getValue("loan_amt_sum")).toString();}
               	}%>
				<font color='red' size=2 >*</font>
			</td>
        </tr>
        <tr>
			<td bgcolor='#D8EFEE' rowspan ="2" >授信案期間</td>
			<td bgcolor='#D8EFEE' colspan ="2" >起始年月</td>
			<td width='398' bgcolor='e7e7e7'>
				<input type='text' name='case_begin_year' 
						value="<%
								if(A11_S_Edit != null && A11_S_Edit.size() != 0 && Bean.getValue("case_begin_year") != null){
									String case_begin_year = (Bean.getValue("case_begin_year")==null)?"":Bean.getValue("case_begin_year").toString();
								    out.print(case_begin_year);
								}
								%>" 
						size='3' maxlength='3'>&nbsp;年
						<%	
								if(A11_S_Edit != null && A11_S_Edit.size() != 0 && Bean.getValue("case_begin_year") != null){
									String case_begin_year = (Bean.getValue("case_begin_year")==null)?"":Bean.getValue("case_begin_year").toString();
								    out.print("<input type='hidden' name='case_begin_yearOri' value='"+case_begin_year+"' >");
								}else{
								    out.print("<input type='hidden' name='case_begin_yearOri' value='' >");
								}
								%>
		            <select id="text" name='case_begin_month'>
						<%
						String case_begin_month="";
						if(A11_S_Edit != null && A11_S_Edit.size() != 0){
		            	    case_begin_month= (Bean.getValue("case_begin_month")==null)?"":Bean.getValue("case_begin_month").toString();
		            	}
						for(int i=1;i<=12;i++){
						    if(!"".equals(case_begin_month)){
		                        if(i == Integer.parseInt(case_begin_month))
		                           out.print("<option value="+i+" selected>"+i+"</option>");
		                        else
		                           out.print("<option value="+i+">"+i+"</option>");
						    }else{
						        if(i==1){
						        	out.print("<option value=''></option>");
						        }
						        out.print("<option value="+i+">"+i+"</option>");
						    }
	                    }
                    %>
					</select>月</font>                           
			</td>
		</tr>
        <tr>
			<td width='110' bgcolor='#D8EFEE' colspan ="2" >到期年月</td>
			<td width='398' bgcolor='e7e7e7'>
				<input type='text' name='case_end_year' 
						value="<%
								if(A11_S_Edit != null && A11_S_Edit.size() != 0 && Bean.getValue("case_end_year") != null){
								    String case_end_year = (Bean.getValue("case_end_year")==null)?"":Bean.getValue("case_end_year").toString();
									out.print(case_end_year);
								}
								%>" 
						size='3' maxlength='3'>&nbsp;年
		            <select id="text" name='case_end_month'>
		            <%
						String case_end_month="";
						if(A11_S_Edit != null && A11_S_Edit.size() != 0){
						    case_end_month= (Bean.getValue("case_end_month")==null)?"":Bean.getValue("case_end_month").toString();
		            	}
						for(int i=1;i<=12;i++){
						    if(!"".equals(case_end_month)){
		                        if(i == Integer.parseInt(case_end_month))
		                           out.print("<option value="+i+" selected>"+i+"</option>");
		                        else
		                           out.print("<option value="+i+">"+i+"</option>");
						    }else{
						        if(i==1){
						        	out.print("<option value=''></option>");
						        }
						        out.print("<option value="+i+">"+i+"</option>");
						    }
	                    }
                    %>
					</select>月</font> 
			</td>
		</tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >主辦行</td>
			<td width='398' bgcolor='e7e7e7' >
              	<select name='bank_no_max'> 
	              	<option value=""></option>
	              	<%for(int i=0;i<bankNoMaxList.size();i++){%>
			            <option value="<%=(String)((DataObject)bankNoMaxList.get(i)).getValue("cmuse_id")%>"
			            <% 
				                String bank_no_max = " ";
				        		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
				        		    bank_no_max = (Bean.getValue("bank_no_max")==null)?"":(String)Bean.getValue("bank_no_max");
				        		}
				                if(bank_no_max.equals(((DataObject)bankNoMaxList.get(i)).getValue("cmuse_id"))) out.print("selected");
			            %>
			            ><%=(String)((DataObject)bankNoMaxList.get(i)).getValue("cmuse_name")%></option>                            
	                <%}%>
                </select> 
				<font color='red' size=2 >*</font>
			</td>
		</tr>
		
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">管理行</td>
			<td width='398' bgcolor='e7e7e7' >
				<input type='text' name='manabank_name' value="<%
             				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
             				    String manabank_name = (Bean.getValue("manabank_name")==null)?"":Bean.getValue("manabank_name").toString();
               					out.print(manabank_name);
               			 	}
               		       	%>" size='20' maxlength='40' style='text-align: left;'>
				<font color='red' size=2 >*</font>
			</td>
        </tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >參貸形式</td>
			<td width='398' bgcolor='e7e7e7' >
              	<select name='loan_kind'> 
	              	<option value=""></option>                                                       
	                <%for(int i=0;i<loanKindList.size();i++){%>
		                <option value="<%=(String)((DataObject)loanKindList.get(i)).getValue("cmuse_id")%>"
		                <% 
			                String loan_kind = " ";
			        		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
			        		    loan_kind = (Bean.getValue("loan_kind")==null)?"":(String)Bean.getValue("loan_kind");
			        		}
			                if(loan_kind.equals(((DataObject)loanKindList.get(i)).getValue("cmuse_id"))) out.print("selected");
		                %>
		                ><%=(String)((DataObject)loanKindList.get(i)).getValue("cmuse_name")%></option>                            
	                <%}%>
	                </select> 
				<font color='red' size=2 >*</font>
			</td>
		</tr>
		
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">參貸額度(元)</td>
			<td width='398' bgcolor='e7e7e7' >
				<%if(Bean.getValue("loan_amt")==null||Bean.getValue("loan_amt")!=null && "0".equals(Bean.getValue("loan_amt"))){ %>
             		<input align='right' type='text' name='loan_amt' value="" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>  
             		<input type='hidden' name='loan_amtOri' value='' > 
               	<%}else{ %>
               		<input align='right' type='text' name='loan_amt' value="<%=Utility.setCommaFormat((Bean.getValue("loan_amt")) == null ? "":(Bean.getValue("loan_amt")).toString())%>" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>
               		<input type='hidden' name='loan_amtOri' value='<%=Utility.setCommaFormat((Bean.getValue("loan_amt")) == null ? "":(Bean.getValue("loan_amt")).toString())%>' > 
               	<%}%>
				<font color='red' size=2 >*</font>
			</td>
        </tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">實際授信餘額(元)</td>
			<td width='398' bgcolor='e7e7e7' >
				<%if(Bean.getValue("loan_bal_amt")==null||Bean.getValue("loan_bal_amt")!=null && "0".equals(Bean.getValue("loan_bal_amt"))){ %>
             		<input align='right' type='text' name='loan_bal_amt' value="" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>
             		<input type='hidden' name='loan_bal_amtOri' value='' >    
               	<%}else{ %>
               		<input align='right' type='text' name='loan_bal_amt' value="<%=Utility.setCommaFormat((Bean.getValue("loan_bal_amt")) == null ? "":(Bean.getValue("loan_bal_amt")).toString())%>" size='20' maxlength='14' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);' style='text-align: right;'>
               		<input type='hidden' name='loan_bal_amtOri' value='<%=Utility.setCommaFormat((Bean.getValue("loan_bal_amt")) == null ? "":(Bean.getValue("loan_bal_amt")).toString())%>' >
               	<%}%>
			</td>
        </tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >信用部參貸部分之授信用途</td>
			<td width='398' bgcolor='e7e7e7' >
              	<select name='loan_type'>
              		<option value=""></option>                                                        
                	<%for(int i=0;i<loanTypeList.size();i++){%>
	                <option value="<%=(String)((DataObject)loanTypeList.get(i)).getValue("cmuse_id")%>" 
	                <% 
		                String loan_type = " ";
		        		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
		        		    loan_type = (Bean.getValue("loan_type")==null)?"":(String)Bean.getValue("loan_type");
		        		}
		                if(loan_type.equalsIgnoreCase((String)((DataObject)loanTypeList.get(i)).getValue("cmuse_id"))) out.print("selected");
	                %>
	                ><%=(String)((DataObject)loanTypeList.get(i)).getValue("cmuse_name")%></option>                            
                	<%}%>
                </select> 
	        	<font color='red' size=2 >*</font>
			</td>
		</tr>
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >目前放款繳息情形</td>
			<td width='398' bgcolor='e7e7e7' >
              	<select name='pay_state'>
              		<option value=""></option>                                                        
                	<%for(int i=0;i<payStateList.size();i++){%>
	            	<option value="<%=(String)((DataObject)payStateList.get(i)).getValue("cmuse_id")%>"                                                        
		        	<%      String pay_state = " ";
		        			if(A11_S_Edit != null && A11_S_Edit.size() != 0){
		        			    pay_state = (Bean.getValue("pay_state")==null)?"":(String)Bean.getValue("pay_state");
		        			}
		           		    if(pay_state.equalsIgnoreCase((String)((DataObject)payStateList.get(i)).getValue("cmuse_id"))) out.print("selected");
	                %>
	                ><%=(String)((DataObject)payStateList.get(i)).getValue("cmuse_name")%></option>                            
                	<%}%>
                </select> 
			</td>
		</tr>
		
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >有無違反契約承諾條款</td>
			<td width='398' bgcolor='e7e7e7' >
              	<select name='violate_type'>
	              	<option value=""></option>                                                        
	                <%for(int i=0;i< violateTypeList.size();i++){%>
	                <option value="<%=(String)((DataObject) violateTypeList.get(i)).getValue("cmuse_id")%>"                                                        
		                <%
			                String violate_type = " ";
			        		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
			        		    violate_type = (Bean.getValue("violate_type")==null)?"":(String)Bean.getValue("violate_type");
			        		}
			                if(violate_type.equalsIgnoreCase((String)((DataObject) violateTypeList.get(i)).getValue("cmuse_id"))) out.print("selected");
			            %>
		                ><%=(String)((DataObject) violateTypeList.get(i)).getValue("cmuse_name")%></option>                            
	                <%}%>
                </select> 
			</td>
		</tr>
		
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">目前放款利率%</td>
			<td width='398' bgcolor='e7e7e7' >
				<input type='text' name='loan_rate' value="<%
             				if(A11_S_Edit != null && A11_S_Edit.size() != 0){
             				    String loan_rate = (Bean.getValue("loan_rate")==null)?"":String.valueOf(Double.parseDouble(Bean.getValue("loan_rate").toString()));
               					out.print(loan_rate);
               			 	}
               		       	%>"
               			size='10' maxlength='8' style='text-align: right;'> %
               	<font color='red' size=2 >(填至小數點第４位)</font>
			</td>
        </tr>
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">是否本月新增案件</td>
			<td width='380' bgcolor='e7e7e7'>
			    <select id="text" name='new_case'>
				    <option value=""></option> 
				    <%for(int i=1;i<3;i++){%>
				    <option value="<%=i%>"                                                        
			                <%
				                String new_case = " ";
				        		if(A11_S_Edit != null && A11_S_Edit.size() != 0){
				        		    new_case = (Bean.getValue("new_case")==null)?"":(String)Bean.getValue("new_case");
				        		}
				                if(new_case.equalsIgnoreCase(String.valueOf(i))) out.print("selected");
				            %>
			                ><%= (i==1)?"是":"否"%></option>                            
		             <%}%>
					</select>
				<font color='red' size=2 >*</font>                          
			</td>
		</tr>
      </table>
      </center>
    </td>
  </tr>
  <!-- ---------------------------------------------------------------------------------------------------------------------------- -->
  <tr>
    <td width="600">
       <center>
          <jsp:include page="getMaintainUser.jsp" flush="true" /><!--載入維護者資訊 -->
       </center>
    </td>
  </tr>

  <tr>
    <td width="600">
      <center>
        <table border="0" cellspacing="1">
          <tr>
          	<%
          		if(act.equals("New") || act.equals("nextPg") || act.equals("continueEditCnt")){
          		   if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>
		          	<td width="66">
		          		<a href="javascript:doSubmit(this.document.forms[0],'Insert','A11','<%=A11_Lock%>','<%=isLastMonthData%>','<%=isHaveApplyData%>','<%=isHaveNoApplyData%>','<%=isHaveNoApplyLoanData%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)">
		          		<img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
		          	</td>
          		<%	}
          		}else {
          		   if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%> 
          		   		<%if(A11_S_Edit != null && A11_S_Edit.size() != 0){ %>
							<td width="66">
								<a href="javascript:doSubmit(this.document.forms[0],'Update','A11','<%=A11_Lock%>','<%=isLastMonthData%>','<%=isHaveApplyData%>','<%=isHaveNoApplyData%>','<%=isHaveNoApplyLoanData%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_updateb.gif',1)">
								<img src="images/bt_update.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
							</td>
				   		<%} %>
				   <%} %>
				   <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>
						<td width="66">
							<a href="javascript:doSubmit(this.document.forms[0],'Delete','A11','<%=A11_Lock%>','<%=isLastMonthData%>','<%=isHaveApplyData%>','<%=isHaveNoApplyData%>','<%=isHaveNoApplyLoanData%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)">
							<img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
						</td>
					<%}
				 }%>
           	  <td width="93">
            		<a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)">
            		<img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
            	</td>
          </tr>
        </table>
      </center>
   </td>
  </tr>

  <tr>
    <td width="600">
    
      <table class="sbody" border="0" cellspacing="1">
        <tr>
        	<td colspan="2">
				<font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle">
				<font color="#007D7D">使用說明 : </font></font>
          	</td>
        </tr>
        
        <tr>
        	<td width="16">&nbsp;</td>
          	<td width="577"> 
          		<ul>
          			<li><b>「授信案總金額」請填列聯貸案之總金額。若聯貸案有授信細項（甲、乙…項），請填甲、乙、丙…等之總額。</b></li>
          			<li>確認輸入資料無誤後, 按【確定】即將本表上的資料, 於資料庫中建檔,按[回查詢主畫面]回至查詢清單畫面。</li>
            		<li><b>每月均已將上月資料全數載入申報當月，請確認修正<font color="red">實際授信餘額</font>等申報相關資料。</b></li>
            		<li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>
        			<li>如放棄修改或無修改之資料需輸入, 按【回上一頁】即離開本程式。</li>
        			<li><font color="red">* </font>為必填欄位。</li>
            	</ul>
          	</td>
        </tr>
      </table>
      
    </td>
  </tr>

</table>

</body>

</html>
