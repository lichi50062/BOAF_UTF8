<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/TM002W.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<%
	String acc_Tr_Type = ( request.getAttribute("acc_Tr_Type")==null ) ? "" : (String)request.getAttribute("acc_Tr_Type");		
	System.out.println(" sAcc_Tr_Type="+acc_Tr_Type);
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	List AllLoanItem = (List)request.getAttribute("AllLoanItem");//貸款種類別
	List AllLoanSubItem = (List)request.getAttribute("AllLoanSubItem");//貸款子項別
	List AllBankList = (List)request.getAttribute("AllBankList");//貸款經辦機構
	List AllCityList = (List)request.getAttribute("AllCityList");
	//List dataList = (List)request.getAttribute("EditInfo");
	List SelDataList = (List)request.getAttribute("SelDataList");//已挑選的舊貸展延需求、新貸需求
	List SelBankList = (List)request.getAttribute("SelBankList");//已挑選的貸款經辦經機名稱
	int AllLoanItemSize=AllLoanItem.size(), AllLoanSubItemSize=AllLoanSubItem.size();
	//預設
	String acc_Tr_Name="",rate_Period1="0",rate_Period2="0";
	String loanKind1 = (String)((DataObject)AllLoanItem.get(0)).getValue("loan_item");
	String loanKind2 = loanKind1;
	String bank_Type = (String)((DataObject)AllCityList.get(0)).getValue("hsien_id");
	if(SelDataList != null && SelDataList.size() > 0 ){
		for(int i=0;i<SelDataList.size();i++){
			//acc_div,acc_code,acc_name
			acc_Tr_Name = (String)((DataObject)SelDataList.get(0)).getValue("acc_tr_name");
			if("01".equals((String)((DataObject)SelDataList.get(i)).getValue("acc_div"))){
		    	String acc_code = (String)((DataObject)SelDataList.get(i)).getValue("acc_code");
		    	//if(!"".equals(acc_code)){
		    		//loanKind1 = acc_code.substring(0,2);
		    	//}
		   		rate_Period1 = (String)((DataObject)SelDataList.get(i)).getValue("rate_period");
		    }
			if("02".equals((String)((DataObject)SelDataList.get(i)).getValue("acc_div"))){
				String acc_code = (String)((DataObject)SelDataList.get(i)).getValue("acc_code");
				//if(!"".equals(acc_code)){
					//loanKind2 = acc_code.substring(0,2);
				//}
		    	rate_Period2 = (String)((DataObject)SelDataList.get(i)).getValue("rate_period");
		    }
		}
	}
	if(SelBankList != null && SelBankList.size() !=0 ){
		bank_Type = (String)((DataObject)SelBankList.get(0)).getValue("hsien_id");
	}
	//取得TM002W的權限
  	Properties permission = ( session.getAttribute("TM002W")==null ) ? new Properties() : (Properties)session.getAttribute("TM002W"); 
  	if(permission == null){
         System.out.println("TM002W_Edit.permission == null");
      }else{
         System.out.println("TM002W_Edit.permission.size ="+permission.size());
                 
      }
%>
<script type="text/javascript">
function doOnload(){
	form.rate_Period1.value="<%=rate_Period1%>";
	form.rate_Period2.value="<%=rate_Period2%>";
	form.loanKind1.value="<%=loanKind1%>";
	form.loanKind2.value="<%=loanKind2%>";
	form.bank_Type.value="<%=bank_Type%>";
}
</script>
<HTML>
<HEAD>
<TITLE>協助措施規劃內容基本資料維護作業</TITLE>
</HEAD>
<BODY bgColor=white onLoad='doOnload();'>
<font color='#000000' size=4><b><center>協助措施規劃內容基本資料維護作業</center></b></font>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<Form name='form' method=post action='/pages/TM002W.jsp' >
<input type='hidden' name='selSubList1'>
<input type='hidden' name='selSubList2'>
<input type='hidden' name='selBankList'>
<table width="835"  border=0 align='center' cellpadding="0" cellspacing="0">
	<tr> 
		<div align="right"><jsp:include page="getLoginUser.jsp?width=835" flush="true" /></div> 
    </tr>
</table>
		<%
          String nameColor="nameColor_sbody";
          String textColor="textColor_sbody";
          String bordercolor="#3A9D99";
        %>
<Table width="835"  border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
		  
	<tr> 
    	<td align="left" width="99" colspan='2' class="<%=nameColor%>">協助措施名稱</td>
		<td width="709" height="2" class="<%=textColor%>" colspan='5'>
			<input type='text' size=80 name='acc_Tr_Name' maxlength='270' value='<%=acc_Tr_Name%>'>
			<input type='hidden' name='acc_Tr_Type' value='<%=acc_Tr_Type%>'>
		</td>			    
	</tr>
  
	<tr class="sbody">
		<td align="left" width="50"  height="4" rowspan="2" class="<%=nameColor%>">舊貸展延需求</td>
		<td align="left" width="32"  height="2" class="<%=nameColor%>">貸款種類</td>
		<td width="709"  class="<%=textColor%>" height="2">
		<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" >
        	<tr class="sbody"> 
        		<td width="307">貸款種類別
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
									if(SelDataList != null && SelDataList.size() !=0 ){
										for(int s=0;s<SelDataList.size();s++){
											if("01".equals((String)((DataObject)SelDataList.get(s)).getValue("acc_div"))
													&& ((String)((DataObject)SelDataList.get(s)).getValue("acc_code")).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem"))){
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
								<%if(SelDataList != null && SelDataList.size() !=0 ){
									for(int s=0;s<SelDataList.size();s++){
										if("01".equals((String)((DataObject)SelDataList.get(s)).getValue("acc_div"))){
											out.print("<option value='"+(String)((DataObject)SelDataList.get(s)).getValue("acc_code")+"'>"+(String)((DataObject)SelDataList.get(s)).getValue("acc_name")+"</option>");
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
		<td align="left" width="32"  height="2" class="<%=nameColor%>">免息期間</td>
		<td width="709"  class="<%=textColor%>" height="2">
			<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
        		<tr class="sbody"> 
        			<td width="307">
                    	<select size="1" name="rate_Period1">
							<option value='0'>半年</option>
							<option value='1'>1年</option>
							<option value='2'>2年</option>
						</select></td>
        			<td width="52">　</td>
        			<td width="337">
        　			</tr>
    		</table>
		</td>   
	</tr>

    <tr class="sbody">
		<td align="left" width="50"  height="4" rowspan="2" class="<%=nameColor%>">新貸需求</td>
        <td align="left" width="32"  height="2" class="<%=nameColor%>">貸款種類</td>
        <td width="709"  class="<%=textColor%>" height="2">
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
									if(SelDataList != null && SelDataList.size() !=0 ){
										for(int s=0;s<SelDataList.size();s++){
											if("02".equals((String)((DataObject)SelDataList.get(s)).getValue("acc_div"))
													&& ((String)((DataObject)SelDataList.get(s)).getValue("acc_code")).equals((String)((DataObject)AllLoanSubItem.get(i)).getValue("subitem"))){
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
									<%if(SelDataList != null && SelDataList.size() !=0 ){
										for(int s=0;s<SelDataList.size();s++){
											if("02".equals((String)((DataObject)SelDataList.get(s)).getValue("acc_div"))){
												out.print("<option value='"+(String)((DataObject)SelDataList.get(s)).getValue("acc_code")+"'>"+(String)((DataObject)SelDataList.get(s)).getValue("acc_name")+"</option>");
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
		<td align="left" width="32"  height="2" class="<%=nameColor%>">免息期間</td>
		<td width="709"  class="<%=textColor%>" height="2">
			<table width="709" border="0" align="center" cellpadding="1" cellspacing="1" class="body_bgcolor">
        		<tr class="sbody"> 
        			<td width="307">
             			<select size="1" name="rate_Period2">
							<option value='0'>半年</option>
							<option value='1'>1年</option>
							<option value='2'>2年</option>
						</select></td>
        			<td width="52">　</td>
        			<td width="337">
        　			</tr>
    		</table>
		</td>   
	</tr>

                        
	<tr class="sbody">
		<td align="left" colspan="2" class="<%=nameColor%>">貸款經辦機構名稱 </td>
		<td width="709"  class="<%=textColor%>" >
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
			        <tr ><td align="center" bgcolor=#9AD3D0 >已選擇項目</td></tr> 
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
	
	<table border=0 align=center width="824">
		<tr >
			<td colspan="13"><div align="center">
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
							    
	    		</div></td> 
	    </tr>
	    <tr>
                      <td colspan="2" width="684" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明        
                        :</font></font> </td>
                    </tr>
	    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
 						<ul>
 							<li class="sbody">確認輸入資料無誤後,按【確定】即將本表上的資料,於資料庫中建檔。</li>
							<li class="sbody">按【修改】即將修改的資料，寫入資料庫中。</li>
							<li class="sbody">欲重新輸入資料,按【取消】即將本表上的資料清空。</li>
							<li class="sbody">如放棄修改或無修改之資料需輸入，按【回上一頁】即離開本程式。</li>
                        </ul>                            
                      </td>
                    </tr>
	</Table>

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


-->
</script>

</BODY>
</HTML>

