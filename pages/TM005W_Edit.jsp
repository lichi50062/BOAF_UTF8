<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/PopupCal.js"></script>
<script language="javascript" src="js/TM005W.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<%
	String acc_Tr_Type = ( request.getAttribute("acc_Tr_Type")==null ) ? "" : (String)request.getAttribute("acc_Tr_Type");		
	int loopSize = 8;//申報基準日loop
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	List AllAccTr = (List)request.getAttribute("AllAccTr");
	List AllLoanItem = (List)request.getAttribute("AllLoanItem");//貸款種類別
	List AllLoanSubItem = (List)request.getAttribute("AllLoanSubItem");//貸款子項別
	List AllBankList = (List)request.getAttribute("AllBankList");//貸款經辦機構
	List AllCityList = (List)request.getAttribute("AllCityList");
	List AllSelDataList = (List)request.getAttribute("AllSelDataList");//已挑選的舊貸展延需求、新貸需求
	List AllSelItemList = (List)request.getAttribute("AllSelItemList");//預設舊貸展延需求、新貸需求
	List SelPreBankList = (List)request.getAttribute("SelPreBankList");//預設貸款經辦經機名稱
	List SelBankList = (List)request.getAttribute("SelBankList");//已挑選的貸款經辦經機名稱
	List EditInfo = (List)request.getAttribute("EditInfo");
	int AllLoanItemSize=AllLoanItem.size(), AllLoanSubItemSize=AllLoanSubItem.size();
	//預設
	String acc_Tr_Name="";
	String loanKind1 = (String)((DataObject)AllLoanItem.get(0)).getValue("loan_item");
	String loanKind2 = loanKind1;
	String bank_Type = (String)((DataObject)AllCityList.get(0)).getValue("hsien_id");
	String startDate_year="",startDate_month="",startDate_day="";
	String endDate_year="",endDate_month="",endDate_day="";
	String beginDate_year="",beginDate_month="",beginDate_day="";
	String applyDate_year="",applyDate_month="",applyDate_day="";
	String applyType="",applyDateS="";
	if(SelBankList != null && SelBankList.size() !=0 ){
		bank_Type = (String)((DataObject)SelBankList.get(0)).getValue("hsien_id");
	}
	if(EditInfo != null && EditInfo.size() > 0 ){
		for(int i=0;i<EditInfo.size();i++){
			acc_Tr_Name = (String)((DataObject)EditInfo.get(i)).getValue("acc_tr_name");
			applyType = (String)((DataObject)EditInfo.get(i)).getValue("applytype");
			String startDate = (String)((DataObject)EditInfo.get(i)).getValue("startdate");
			if(!"".equals(startDate)){
				startDate_year=startDate.substring(0, 3);
				startDate_month=startDate.substring(3, 5);
				startDate_day=startDate.substring(5, 7);
			}
			String endDate = (String)((DataObject)EditInfo.get(i)).getValue("enddate");
			if(!"".equals(endDate)){
				endDate_year=endDate.substring(0, 3);
				endDate_month=endDate.substring(3, 5);
				endDate_day=endDate.substring(5, 7);
			}
			String beginDate = (String)((DataObject)EditInfo.get(i)).getValue("begindate");
			if(!"".equals(beginDate)){
				beginDate_year=beginDate.substring(0, 3);
				beginDate_month=beginDate.substring(3, 5);
				beginDate_day=beginDate.substring(5, 7);
			}
			if(!"".equals(applyDateS)) applyDateS+=";";
			applyDateS += (String)((DataObject)EditInfo.get(i)).getValue("applydate");
		}
	}
	//取得TM005W的權限
  	Properties permission = ( session.getAttribute("TM005W")==null ) ? new Properties() : (Properties)session.getAttribute("TM005W"); 
  	if(permission == null){
         System.out.println("TM005W_Edit.permission == null");
      }else{
         System.out.println("TM005W_Edit.permission.size ="+permission.size());
                 
      }
	
%>
<script type="text/javascript">
function doValid(form){
	form.startDate.value = '';
	form.endDate.value = '';
	form.beginDate.value = '';
	form.applyDateS.value = '';
	if(form.startDate_year.value!=""){
	    if(!mergeCheckedDate("startDate_year;startDate_month;startDate_day","startDate")){
	        form.startDate_year.focus();
	        return false;
	    }
    }
	if(form.endDate_year.value!=""){
	    if(!mergeCheckedDate("endDate_year;endDate_month;endDate_day","endDate")){
	        form.endDate_year.focus();
	        return false;
	    }
    }
	if(form.beginDate_year.value!=""){
	    if(!mergeCheckedDate("beginDate_year;beginDate_month;beginDate_day","beginDate")){
	        form.beginDate_year.focus();
	        return false;
	    }
    }
	if(form.startDate.value=='') {
		alert('請輸入申報期間(起)!');
		return false ;
    }
    if(form.endDate.value=='') {
		alert('請輸入申報期間(迄)!');
		return false ;
    }
    if(form.startDate.value > form.endDate.value){
    	alert('查詢申報期間  起日不得大於迄日!');
		return false ;
    }
    if(form.beginDate.value=='') {
		alert('請輸入發布日期!');
		return false ;
    }
    size = <%=loopSize%>;
    <%for(int t=0;t<loopSize;t++){%>
    	var obj = document.getElementById("applyDate_"+<%=t%>);
		if(obj.style.display == "block"){
			if(!mergeCheckedDate("applyDate_year<%=t%>;applyDate_month<%=t%>;applyDate_day<%=t%>","applyDate<%=t%>")){
		        form.applyDate_year<%=t%>.focus();
		        return false;
		    }else{
		    	if(form.applyDate<%=t%>.value!=''){
		    		if(form.applyDateS.value!='')form.applyDateS.value+=";";
		    		form.applyDateS.value+=form.applyDate<%=t%>.value;
		    	}
		    }
		}
    <%}%>
    if(form.applyDateS.value=='') {
		alert('請輸入申報基準日!');
		return false ;
    }
    if(form.acc_Tr_Name.value==''){
		alert("請輸入協助措施名稱");
		return false ;
	}
	getSelListDst(this.document.forms[0].selSubList1, this.document.forms[0].itemListDst1);
	getSelListDst(this.document.forms[0].selSubList2, this.document.forms[0].itemListDst2);
	getSelListDst(this.document.forms[0].selBankList, this.document.forms[0].BankListDst);
	if(form.selSubList1.value=='' && form.selSubList2.value==''){
		alert("請選擇舊貸展延需求或新貸需求");
		return false ;
	}
	if(form.selBankList.value==''){
		alert("請選擇貸款經辦機構名稱");
		return false ;
	}
    return true;
}
function fn_changeAccTrName(){
	var element = form.AccTr;
	form.acc_Tr_Name.value = element.options[element.selectedIndex].text;
	changePreList('01',element.value,form.itemListDst1);
	changePreList('02',element.value,form.itemListDst2);
	changePreList('03',element.value,form.BankListDst);
}
function changePreList(ss,newType,dst){
	dst.options.length = 0;
	var n=0;
	if(ss=='03'){
		<%if(SelPreBankList != null && SelPreBankList.size() !=0 ){
			for(int s=0;s<SelPreBankList.size();s++){%>
				if(newType == '<%=(String)((DataObject)SelPreBankList.get(s)).getValue("acc_tr_type")%>'){
						dst.options[n] = new Option('<%=(String)((DataObject)SelPreBankList.get(s)).getValue("bank_name")%>', '<%=(String)((DataObject)SelPreBankList.get(s)).getValue("bank_no")%>');
						n++;
				}
			 <%}
		}%>	
	}else{
		<%if(AllSelItemList != null && AllSelItemList.size() !=0 ){
			for(int s=0;s<AllSelItemList.size();s++){%>
				if(ss=='<%=(String)((DataObject)AllSelItemList.get(s)).getValue("acc_div")%>'
					&& newType == '<%=(String)((DataObject)AllSelItemList.get(s)).getValue("acc_tr_type")%>'){
						dst.options[n] = new Option('<%=(String)((DataObject)AllSelItemList.get(s)).getValue("acc_name")%>', '<%=(String)((DataObject)AllSelItemList.get(s)).getValue("acc_code")%>');
						n++;
				}
			 <%}
		}%>
	}
	fn_changeListSrc(ss);
}
function doOnload(){
	form.loanKind1.value="<%=loanKind1%>";
	form.loanKind2.value="<%=loanKind2%>";
	form.bank_Type.value="<%=bank_Type%>";
	
	if("<%=applyType%>"=='1'){form.applyType[1].checked=true;}
	if("<%=applyType%>"=='4'){form.applyType[2].checked=true;}
}
</script>
<HTML>
<HEAD>
<TITLE>協助措施申報項目維護作業</TITLE>
</HEAD>
<BODY bgColor=white onLoad='doOnload();'>
<font color='#000000' size=4><b><center>協助措施申報項目維護作業</center></b></font>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<Form name='form' method=post action='/pages/TM005W.jsp' >
<input type='hidden' name='sAcc_Tr_Name' value='<%=( request.getAttribute("sAcc_Tr_Name")==null ) ? "" : (String)request.getAttribute("sAcc_Tr_Name")%>'>
<input type='hidden' name='selSubList1'>
<input type='hidden' name='selSubList2'>
<input type='hidden' name='selBankList'>
<table width="835"  border=1 align='center' cellpadding="1" cellspacing="1">
	<div align="right"><jsp:include page="getLoginUser.jsp?width=835" flush="true" /></div>
</table>
		<%
          String nameColor="nameColor_sbody";
          String textColor="#e7e7e7";
          String bordercolor="#3A9D99";
        %>
<Table width="835" border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
	<tr> 
    	<td align="left" width="99" colspan='2' class="<%=nameColor%>">協助措施名稱</td>
		<td width="709" height="2" bgColor="<%=textColor%>" colspan='5'>
		<%if(act.equals("new")){%>
			<select name="AccTr" onChange="fn_changeAccTrName();">
			   <%if(AllAccTr!=null && AllAccTr.size()>0){
					for(int i=0; i<AllAccTr.size(); i++){ 
						out.print("<option value='"+(String)((DataObject)AllAccTr.get(i)).getValue("acc_tr_type")+"'>"+(String)((DataObject)AllAccTr.get(i)).getValue("acc_tr_name")+"</option>");
					}
				}%>
			</select>
			<input type='text' size=80 name='acc_Tr_Name' maxlength='270' value='<%=acc_Tr_Name%>'>
			
		<%}else{ %>
			<%=acc_Tr_Name%>
			<input type='hidden'  name='acc_Tr_Name' value='<%=acc_Tr_Name%>'>
		<%} %>
		<input type='hidden' name='acc_Tr_Type' value='<%=acc_Tr_Type%>'>
		</td>			    
	</tr>
  	<tr> 
    	<td align="left" width="99" colspan='2' class="<%=nameColor%>">申報頻率</td>
		<td width="709" height="2" bgColor="<%=textColor%>" colspan='5'>
		<input type='radio' name='applyType' value='2' checked>雙週報
		<input type='radio' name='applyType' value='1'>週報
		<input type='radio' name='applyType' value='4'>月報
		</td>			    
	</tr>
	<tr> 
    	<td align="left" width="99" colspan='2' class="<%=nameColor%>">申報區間</td>
		<td width="709" height="2" bgColor="<%=textColor%>" colspan='5'>
		<input type='hidden' name='startDate'><!-- //申報區間(起) -->
		<input type='hidden' name='endDate'><!-- //申報區間(迄) -->
		<input type='text' id='startDate_year' name='startDate_year' value='<%=startDate_year %>' size='2' maxlength='3' >年
		<select id='startDate_month' name='startDate_month'>
			<option></option>
			<%for (int j = 1; j <= 12; j++) {
				if (j < 10){%>
					<option value=0<%=j%> <%if(!"".equals(startDate_month) && Integer.parseInt(startDate_month)==j) out.print("selected");%> onblur='CheckYear(this)' >0<%=j%></option>
				<%}else{%>
					<option value=<%=j%>  <%if(!"".equals(startDate_month) && Integer.parseInt(startDate_month)==j) out.print("selected");%> onblur='CheckYear(this)' ><%=j%></option>
				<%}
			}%>
		</select>月
				            			
		<select id='startDate_day' name='startDate_day' >
			<option></option>
			<%for (int j = 1; j <= 31; j++) {
				if (j < 10){ %>
					<option value=0<%=j%>  <%if(!"".equals(startDate_day) && Integer.parseInt(startDate_day)==j) out.print("selected");%> >0<%=j%></option>
				<%}else{ %>
					<option value=<%=j%> <%if(!"".equals(startDate_day) && Integer.parseInt(startDate_day)==j) out.print("selected");%> ><%=j%></option>
				<%}
			}%>
		</select>日
		<button name='button_startDate' onClick="popupCal('form','startDate_year,startDate_month,startDate_day','BTN_date_startDate',event)">
			<img align="absmiddle" border='0' name='BTN_date_startDate' src='images/clander.gif'>
			</button>
		~
		<input type='text' id='endDate_year' name='endDate_year' value='<%=endDate_year%>' size='2' maxlength='3' >年
		<select id='endDate_month' name='endDate_month' >
			<option></option>
			<%for (int j = 1; j <= 12; j++) {
				if (j < 10){%>
					<option value=0<%=j%> <%if(!"".equals(endDate_month) && Integer.parseInt(endDate_month)==j) out.print("selected");%> >0<%=j%></option>
				<%}else{%>
					<option value=<%=j%>  <%if(!"".equals(endDate_month) && Integer.parseInt(endDate_month)==j) out.print("selected");%>  ><%=j%></option>
				<%}
			}%>
		</select>月
				            			
		<select id='endDate_day' name='endDate_day' >
			<option></option>
			<%for (int j = 1; j <= 31; j++) {
				if (j < 10){ %>
					<option value=0<%=j%>  <%if(!"".equals(endDate_day) && Integer.parseInt(endDate_day)==j) out.print("selected");%> >0<%=j%></option>
				<%}else{ %>
					<option value=<%=j%> <%if(!"".equals(endDate_day) && Integer.parseInt(endDate_day)==j) out.print("selected");%> ><%=j%></option>
				<%}
			}%>
		</select>日
		<button name='button_endDate' onClick="popupCal('form','endDate_year,endDate_month,endDate_day','BTN_date_endDate',event)">
			<img align="absmiddle" border='0' name='BTN_date_endDate' src='images/clander.gif'>
			</button>
		</td>			    
	</tr>
	<tr> 
    	<td align="left" width="99" colspan='2' class="<%=nameColor%>">發布日期</td>
		<td width="709" height="2" bgColor="<%=textColor%>" colspan='5'>
			<input type='hidden' name='beginDate'><!-- //發布日期 -->
			<input type='text' id='beginDate_year' name='beginDate_year' value='<%=beginDate_year%>' size='2' maxlength='3' >年
			<select id='beginDate_month' name='beginDate_month' >
				<option></option>
				<%for (int j = 1; j <= 12; j++) {
					if (j < 10){%>
						<option value=0<%=j%> <%if(!"".equals(beginDate_month) && Integer.parseInt(beginDate_month)==j) out.print("selected");%> >0<%=j%></option>
					<%}else{%>
						<option value=<%=j%>  <%if(!"".equals(beginDate_month) && Integer.parseInt(beginDate_month)==j) out.print("selected");%>  ><%=j%></option>
					<%}
				}%>
			</select>月
					            			
			<select id='beginDate_day' name='beginDate_day' >
				<option></option>
				<%for (int j = 1; j <= 31; j++) {
					if (j < 10){ %>
						<option value=0<%=j%>  <%if(!"".equals(beginDate_day) && Integer.parseInt(beginDate_day)==j) out.print("selected");%> >0<%=j%></option>
					<%}else{ %>
						<option value=<%=j%> <%if(!"".equals(beginDate_day) && Integer.parseInt(beginDate_day)==j) out.print("selected");%> ><%=j%></option>
					<%}
				}%>
			</select>日
			<button name='button_beginDate' onClick="popupCal('form','beginDate_year,beginDate_month,beginDate_day','BTN_date_beginDate',event)">
			<img align="absmiddle" border='0' name='BTN_date_beginDate' src='images/clander.gif'>
			</button>
		</td>			    
	</tr>
	<tr> 
    	<td align="left" colspan='2' class="<%=nameColor%>">申報基準日</td>
		<td width="709" bgColor="<%=textColor%>" colspan='5'>
		<input type='hidden' name='applyDateS'><!-- //APPLYDATE(多筆) 要日曆小視窗、+、- -->
		<%for(int t=0; t<loopSize; t++){
			String showTag = "none";
			if(t==0)showTag = "block";
			if(!"".equals(applyDateS)){
				String[] dates = applyDateS.split(";");
				for(int s=0;s<dates.length;s++){
					if(t==s){
						showTag = "block";
						applyDate_year=dates[s].substring(0, 3);
						applyDate_month=dates[s].substring(3, 5);
						applyDate_day=dates[s].substring(5, 7);
					}
				}
			}
			out.println("<div id='applyDate_"+t+"' style='display:"+showTag+"'> ");
			out.println("<table>");
    		%>
    		<%
    		out.println("<input type='hidden' name='applyDate"+t+"'>");
			out.println("<input type='text' id='applyDate_year"+t+"' name='applyDate_year"+t+"' value='"+applyDate_year+"' size='2' maxlength='3' >年");
		   	out.println("<select id='applyDate_month"+t+"' name='applyDate_month"+t+"' >");
	 		out.println("<option></option>");
	  		for (int j = 1; j <= 12; j++) {
	   			if (j < 10){%>
		    		<option value=0<%=j%> <%if(!"".equals(applyDate_month) && Integer.parseInt(applyDate_month)==j) out.print("selected");%> >0<%=j%></option>
    	    	<%}else{%>
    	    		<option value=<%=j%>  <%if(!"".equals(applyDate_month) && Integer.parseInt(applyDate_month)==j) out.print("selected");%> ><%=j%></option>
    	    	<%}
	   		}
			out.println("</select>月");
			
			out.println("<select id='applyDate_day"+t+"' name='applyDate_day"+t+"' >");
	 		out.println("<option></option>");
	  		for (int j = 1; j <= 31; j++) {
	   			if (j < 10){ %>
		    		<option value=0<%=j%>  <%if(!"".equals(applyDate_day) && Integer.parseInt(applyDate_day)==j) out.print("selected");%> >0<%=j%></option>
    	    	<%}else{ %>
    	    		<option value=<%=j%> <%if(!"".equals(applyDate_day) && Integer.parseInt(applyDate_day)==j) out.print("selected");%> ><%=j%></option>
    	    	<%}
	  		}
			out.println("</select>日");%>
			<button name='button<%=t %>' onClick="popupCal('form','applyDate_year<%=t %>,applyDate_month<%=t %>,applyDate_day<%=t %>','BTN_date_<%=t %>',event)">
			<img align="absmiddle" border='0' name='BTN_date_<%=t %>' src='images/clander.gif'>
			</button>
			<%if(t<loopSize-1){ %>
				<input type='button' class='button' value='＋' id='add_applyDate_<%=t+1 %>' onclick="showDiv('applyDate_<%=t+1 %>',<%=t+1 %>)" >
			<%} %>
			<%if(t>0){ %>
				<input type='button' class='button' value='－' id='rmv_applyDate_<%=t %>' onclick="rmvDiv('applyDate_<%=t %>',<%=t %>)" >
			<%} %>
			<%
			out.println("</table>");
			out.println("</div>");
			
			
			showTag = "none";
		}%>
		<font color="#FF0000">點選[+]可新增一筆申報日期,[-]可刪除該申報日期</font>
		</td>			    
	</tr>
	<tr class="sbody">
		<td align="left" class="<%=nameColor%>" colspan="2">舊貸展延需求</td>
		<td width="709" bgColor="<%=textColor%>">
		<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
        	<tr class="sbody"> 
        		<td width="307" >貸款種類別
					<select size="1" name="loanKind1" onChange="fn_changeListSrc('01');">
						<%
						if(AllLoanItem!=null && AllLoanItem.size()>0){
							for(int i=0; i<AllLoanItem.size(); i++){ 
								out.print("<option value='"+(String)((DataObject)AllLoanItem.get(i)).getValue("loan_item")+"'>"+(String)((DataObject)AllLoanItem.get(i)).getValue("loan_item_name")+"</option>");
							}
						}%>
					</select><br>貸款子項別
					<select multiple size=10  name="itemListSrc1" id="itemListSrc1" onDblclick="javascript:movesel(form.itemListSrc1,form.itemListDst1);" style="width: 265; height: 91">							
						<%
						if(AllLoanSubItem!=null && AllLoanSubItem.size()>0){
							boolean flg = true;
							for(int i=0; i<AllLoanSubItem.size(); i++){ 
								if((loanKind1).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("loan_item"))){
									flg = true;
									if(AllSelDataList != null && AllSelDataList.size() !=0 ){
										for(int s=0;s<AllSelDataList.size();s++){
											if("01".equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_div"))
													&& (acc_Tr_Type).equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_tr_type"))
													&& ((String)((DataObject)AllSelDataList.get(s)).getValue("acc_code")).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem"))){
												flg= false;
											}
										}
									}
									if(flg){
										out.print("<option value='"+(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem")+"'>"+(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem_name")+"</option>");
									}
								}
							}
						}%>
					</select>
				</td>
        		<td width="52">
        			<table width="40" border="0" align="center" cellpadding="3" cellspacing="3">
			            <tr> 
			              	<td>
			              	<div align="center">                                  
			              	&nbsp;<a href="javascript:movesel(form.itemListSrc1,form.itemListDst1);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a></div>
			              	</td>
			            </tr>
            		</table>
            	</td>
        		<td width="337">
			        <table>
			        	<tr><td>
			              	<select multiple size=10  name="itemListDst1" id="itemListDst1" onDblclick="javascript:movesel(form.itemListDst1,form.itemListSrc1);" style="width: 292; height: 91">							
								<%if(AllSelDataList != null && AllSelDataList.size() !=0 ){
									for(int s=0;s<AllSelDataList.size();s++){
										if("01".equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_div"))
												&& (acc_Tr_Type).equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_tr_type"))){
											out.print("<option value='"+(String)((DataObject)AllSelDataList.get(s)).getValue("acc_code")+"'>"+(String)((DataObject)AllSelDataList.get(s)).getValue("acc_name")+"</option>");
										}
									}
								}%>
							</select>
						</td></tr>
					</table>
				</td>
      		</tr>
    	</table>
		</td>   
	</tr>
	

    <tr class="sbody">
		<td align="left" class="<%=nameColor%>" colspan="2">新貸需求</td>
        <td width="709" bgColor="<%=textColor%>" >
			<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
        		<tr class="sbody"> 
        			<td width="307">貸款種類別
						<select size="1" name="loanKind2" onChange="fn_changeListSrc('02');">
							<%
							if(AllLoanItem!=null && AllLoanItem.size()>0){
								for(int i=0; i<AllLoanItem.size(); i++){ 
									out.print("<option value='"+(String)((DataObject)AllLoanItem.get(i)).getValue("loan_item")+"'>"+(String)((DataObject)AllLoanItem.get(i)).getValue("loan_item_name")+"</option>");
								}
							}%>
						</select><br>貸款子項別
						<select multiple size=10  name="itemListSrc2" id="itemListSrc2" onDblclick="javascript:movesel(form.itemListSrc2,form.itemListDst2);" style="width: 265; height: 91">							
						<%
						if(AllLoanSubItem!=null && AllLoanSubItem.size()>0){
							boolean flg = true;
							for(int i=0; i<AllLoanSubItem.size(); i++){ 
								if((loanKind2).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("loan_item"))){
									flg = true;
									if(AllSelDataList != null && AllSelDataList.size() !=0 ){
										for(int s=0;s<AllSelDataList.size();s++){
											if("02".equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_div"))
													&& (acc_Tr_Type).equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_tr_type"))
													&& ((String)((DataObject)AllSelDataList.get(s)).getValue("acc_code")).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem"))){
												flg= false;
											}
										}
									}
									if(flg){
										out.print("<option value='"+(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem")+"'>"+(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem_name")+"</option>");
									}
								}
							}
						}%>
						</select></td>
        			<td width="52">
        				<table width="40" border="0" align="center" cellpadding="3" cellspacing="3">
				            <tr> 
				              	<td>
				              	<div align="center">                                  
				              	&nbsp;<a href="javascript:movesel(form.itemListSrc2,form.itemListDst2);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a></div>
				              	</td>
				            </tr>
            			</table>
            		</td>
        			<td width="337">
				        <table>
				        <tr><td>
				              	<select multiple size=10  name="itemListDst2" ondblclick="javascript:movesel(form.itemListDst2,form.itemListSrc2);" style="width: 292; height: 91">							
									<%if(AllSelDataList != null && AllSelDataList.size() !=0 ){
										for(int s=0;s<AllSelDataList.size();s++){
											if("02".equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_div"))
													&& (acc_Tr_Type).equals((String)((DataObject)AllSelDataList.get(s)).getValue("acc_tr_type"))){
												out.print("<option value='"+(String)((DataObject)AllSelDataList.get(s)).getValue("acc_code")+"'>"+(String)((DataObject)AllSelDataList.get(s)).getValue("acc_name")+"</option>");
											}
										}
									}%>
								</select>
						</td></tr>
						</table>
					</td>
      			</tr>
    		</table>
		</td>   
	</tr>
	
                        
	<tr class="sbody">
		<td align="left" colspan="2" class="<%=nameColor%>">貸款經辦機構名稱 </td>
		<td width="709" bgColor="<%=textColor%>" height="30">
			<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
        		<tr class="sbody"> 
        			<td width="215">
        			<span class="mtext">&#32291;&#24066;&#21029; :</span>                                                            
     					<select name="bank_Type" onchange="javascript:fn_changeListSrc('03');" size="1"> 
						    <%
							if(AllCityList!=null && AllCityList.size()>0){
								for(int i=0; i<AllCityList.size(); i++){
									if("100".equals(((DataObject)AllCityList.get(i)).getValue("m_year").toString())){
										out.print("<option value='"+(String)((DataObject)AllCityList.get(i)).getValue("hsien_id")+"'>"+(String)((DataObject)AllCityList.get(i)).getValue("hsien_name")+"</option>");
									}
								}
							}%>                                                            
    					</select>
        				<table>
					        <tr ><td align="center" bgcolor=#9AD3D0>可選擇項目</td></tr>
					        <tr><td>  
					        <select multiple  size=10  name="BankListSrc" id="BankListSrc" ondblclick="javascript:movesel(form.BankListSrc,form.BankListDst);" style="width: 292; height: 160">							
								<%
								if(AllBankList!=null && AllBankList.size()>0){
									boolean flg = true;
									for(int i=0; i<AllBankList.size(); i++){ 
										flg = true;
										if((bank_Type).equals((String)((DataObject)AllBankList.get(i)).getValue("hsien_id"))){
											if(SelBankList != null && SelBankList.size() !=0 ){
												for(int s=0;s<SelBankList.size();s++){
													//將已選擇的項目排除
													if(((String)((DataObject)SelBankList.get(s)).getValue("bank_no")).equals((String)((DataObject)AllBankList.get(i)).getValue("bank_no"))){
														flg = false;
													}
												}
											}
										}else{
											flg = false;
										}
										if(flg){
											out.print("<option value='"+(String)((DataObject)AllBankList.get(i)).getValue("bank_no")+"'>"+(String)((DataObject)AllBankList.get(i)).getValue("bank_no")+(String)((DataObject)AllBankList.get(i)).getValue("bank_name")+"</option>");
										}
									}
								}%>
							</select>
							</td></tr>
						</table>
        			</td>
        			<td width="52">
        			<table width="40" border="0" align="center" cellpadding="3" cellspacing="3">
			            <tr> 
			              	<td>
			              	<div align="center">                                 
			              	<a href="javascript:movesel(form.BankListSrc,form.BankListDst);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a>
			              	</div>
			              	</td>
			            </tr>
			            <tr> 
			              	<td>
			              	<div align="center">                                  
			              	<a href="javascript:moveallsel(form.BankListSrc,form.BankListDst);"><img src="images/arrow_rightall.gif" width="24" height="22" border="0"></a>
			              	</div>
			              	</td>
			            </tr>
			            <tr> 
			              	<td>
			              	<div align="center">                                  
			              	<a href="javascript:movesel(form.BankListDst,form.BankListSrc);"><img src="images/arrow_left.gif" width="24" height="22" border="0"></a>
			              	</div>
			              	</td>
			            </tr>
			            <tr> 
			              	<td height="22">
			              	<div align="center">                                  
			              	<a href="javascript:moveallsel(form.BankListDst,form.BankListSrc);"><img src="images/arrow_leftall.gif" width="24" height="22" border="0"></a>
			              	</div>
			              	</td>
			            </tr>
          		</table></td>
        		<td width="324">
			        　<table>
			        <tr ><td align="center" bgcolor=#9AD3D0>已選擇項目</td></tr> 
			        <tr><td>
				        <select multiple size=10  name="BankListDst" id="BankListDst" ondblclick="javascript:movesel(form.BankListDst,form.BankListSrc);" style="width: 292; height: 160">		
				        	<%if(SelBankList != null && SelBankList.size() !=0 ){
								for(int i=0;i<SelBankList.size();i++){
									out.print("<option value='"+(String)((DataObject)SelBankList.get(i)).getValue("bank_no")+"'>"+(String)((DataObject)SelBankList.get(i)).getValue("bank_no")+(String)((DataObject)SelBankList.get(i)).getValue("bank_name")+"</option>");
								}
							}%>					
						</select>
					</td></tr>
					</table>
			    	</tr>
			    </table>
			</td>   
		</tr>
</Table>
	     <!-- <div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>  -->   
	<table border=0 align=center width="835">
		<tr>                  
			<td><div align="right"><jsp:include page="getMaintainUser.jsp?width=835" flush="true" /></div></td>                                              
    	</tr>
    </table>
	<table border=0 align=center width="835">
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
	</Table>
	<!-- 
<hr>
<p><font color='#FF0000'><strong>使用說明:</strong></font></p>
<ul>
<li>確認輸入資料無誤後,按【確定】即將本表上的資料,於資料庫中建檔。</li>
<li>按【修改】即將修改的資料，寫入資料庫中。</li>
<li>欲重新輸入資料,按【取消】即將本表上的資料清空。</li>
<li>如放棄修改或無修改之資料需輸入，按【回上一頁】即離開本程式。</li>
</ul><hr>
 -->
</Form>

<script language="JavaScript" >


function fn_changeListSrc(ss){
	if(ss=='01'){
		form.itemListSrc1.options.length = 0;
		n=0;
		<%for (int i =0; i < AllLoanSubItemSize; i ++){ %>
			var loanItem = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("loan_item")%>';
			if(loanItem==form.loanKind1.value){
				flg = true;
				var subItem = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem")%>';
				var subItemName = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem_name")%>';
				//將已選擇的項目排除
				for (var j =0; j < form.itemListDst1.options.length; j++){
					if (subItem == form.itemListDst1.options[j].value){
						flg = false;
					}
				}	
				if(flg){
					form.itemListSrc1.options[n] = new Option(subItemName, subItem);
					n++;
				}
			}
		<%}%>
	}else if(ss=='02'){
		form.itemListSrc2.options.length = 0;
		n = 0;
		<%for (int i =0; i < AllLoanSubItemSize; i ++){ %>
			var loanItem = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("loan_item")%>';
			if(loanItem==form.loanKind2.value){
				//將已選擇的項目排除
				flg = true;
				var subItem = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem")%>';
				var subItemName = '<%=(String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem_name")%>';
				//將已選擇的項目排除
				for (var j =0; j < form.itemListDst2.options.length; j++){
					if (subItem == form.itemListDst2.options[j].value){
						flg = false;
					}
				}
				if(flg){
					form.itemListSrc2.options[n] = new Option(subItemName, subItem);
					n++;
				}
			}
		<%}%>
	}else if(ss=='03'){
		form.BankListSrc.options.length = 0;
		n=0;
		<%
		for (int i =0; i < AllBankList.size(); i ++){ %>
			if('<%=(String)((DataObject)AllBankList.get(i)).getValue("hsien_id")%>'==form.bank_Type.value){
				//將已選擇的項目排除
				flg = true;
				var bank_no = '<%=(String)((DataObject)AllBankList.get(i)).getValue("bank_no")%>';
				var bank_name = '<%=(String)((DataObject)AllBankList.get(i)).getValue("bank_name")%>';
				//將已選擇的項目排除
				for (var j =0; j < form.BankListDst.options.length; j++){
					if (bank_no == form.BankListDst.options[j].value){
						flg = false;
					}
				}	
				if(flg){
					form.BankListSrc.options[n] = new Option(bank_name, bank_no);
					n++;
				}
				
			}
		<%}%>
	}
}



function rmvDiv(id,t){
	var obj = document.getElementById(id);
	obj.style.display = "none";
	document.getElementById("applyDate_year"+t).value='';
	document.getElementById("applyDate_month"+t).value='';
	document.getElementById("applyDate_day"+t).value='';
}
function showDiv(id,t){
	var obj = document.getElementById(id);
	if(obj.style.display == "none"){
		obj.style.display = "block";
		document.getElementById("applyDate_year"+t).value='';
		document.getElementById("applyDate_month"+t).value='';
		document.getElementById("applyDate_day"+t).value='';
	}
}

-->
</script>

</BODY>
</HTML>

