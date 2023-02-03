<%
// 95.08.15 create by 2295
// 95.08.23 add 縣市別/金融機構代號名稱預設放在已選的sort報表欄位 by 2295
// 95.09.04 add 若已選的機構代號 != 全部時，才可根據縣市別．金融機構代號sort	   
//                               == 全部時，根據縣市別小計sort by 2295
// 95.11.10 add 區分BOAF.MIS配色 by 2295
// 95.11.16 add 拿掉年 sort by 2295
// 95.12.01 add 各別機構可根據縣市別.金融機構代號.年月sort
//              全部可選根據縣市別小計.年月sort by 2295
// 95.12.04 add 起始日期不可大於結束日期
// 99.04.29 add 排序欄位.可挑選項目套用共用include by 2295
//100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295
//102.11.11 add 漁會科目代號新舊無法同時列印 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS001W";	
	String act = Utility.getTrimString(dataMap.get("act"));		
	
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");	
	String szSortBy = ( session.getAttribute("SortBy")==null ) ? "asc" : (String)session.getAttribute("SortBy");					
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");		
	
	System.out.print(report_no+"_RptOrder.szExcelAction="+szExcelAction);
	System.out.println(":szSortBy="+szSortBy);
%>



<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
//102.11.11 add 漁會科目代號新舊無法同時列印
function doSubmit(report_no,cnd){
    if(cnd == 'createRpt'){
      if(this.document.forms[0].BankList.value == ''){        
         alert('金融機構代碼必須選擇');
         return;
      }
      if(!chInput(this.document.forms[0],'')) return;//95.12.04 add 起始日期不可大於結束日期
      //漁會科目代號新舊無法同時列印    
      if(this.document.forms[0].DS_bank_type.value == '7'){
        if(parseInt(this.document.forms[0].S_YEAR.value)<=102 && parseInt(this.document.forms[0].E_YEAR.value)>=103){
          alert('漁會於103年起改用新的科目代號，新舊漁會科目代號無法同時輸出，請重新輸入結束日期');
          return;
        } 	 
      }
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }   
   MoveSelectToBtn(this.document.forms[0].SortList, this.document.forms[0].SortListDst);	
   fn_ShowPanel(report_no,cnd);      
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='RptOrderfrm'>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="750" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td><table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="240"><img src="images/banner_bg1.gif" width="240" height="17"></td>
                <td width="*" class="title_font">A01營運明細資料表</td>
                <td width="240"><img src="images/banner_bg1.gif" width="240" height="17"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><img src="images/space_1.gif" width="8" height="8"></td>
        </tr>        
        <tr> 
          <td><table width="750" border="1" align="center" cellpadding="0" cellspacing="0" class="bordercolor">
              <tr> 
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3"><table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> <div align="right">
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                          <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %> 
                      	  <a href="javascript:doSubmit('<%=report_no%>','createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData('RptOrder');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','BankList')"><font color='black'>1.金融機構</font></a></td>                            
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptColumn')"><font color='black'>2.報表欄位</font></a></td>
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">3.排序欄位</font></a></td>                                                       
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <%@include file="./include/DS_RptOrder.include" %><!-- 排序欄位.可挑選項目-->
                     <tr> 
                     <td class="body_bgcolor">
                      <table width="750" border="0" cellpadding="1" cellspacing="1">
                       <tr>                           
                        <td width="750" align=left><font color="red" size=2>註：「各縣市彙總表」本排序欄位若未選取，則按「縣市別小計」及「年月」等欄位排序印出。</font></td>                              
                       </tr>                              
                       <tr>                           
                        <td width="750" align=left><font color="red" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;「營運明細資料」本排序欄位若未選取，則按「縣市別」、「金融機構代號名稱」及「年月」等欄位排序印出。</font></td>                              
                       </tr>                              
                      </table>
                     </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=SortList><!--//btnSortList儲存已勾選的報表排序欄位名稱-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構-->
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.04儲存已勾選的結束日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.04儲存已勾選的結束日期-月-->
<INPUT type="hidden" name=DS_bank_type value='<%if(session.getAttribute("DS_bank_type") != null) out.print((String)session.getAttribute("DS_bank_type"));%>'><!--//102.11.11儲存已勾選的結束日期-月-->
</form>
<script language="JavaScript" >
<!--  

<%//從session裡把勾選的欲sort的報表欄位讀出來.放在SortListDst 	
if(session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")){ 
%>
	var bnlist;
	bnlist = "<%=(String)session.getAttribute("SortList")%>";//95.12.06 add
	var a = bnlist.split(',');
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		this.document.forms[0].SortListDst.options[i] = new Option(j[1], j[0]);
	}
<%}%>

<%//從session裡把已勾選的報表欄位讀出來.放在FieldListDst 	
if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){ 
   System.out.println("DS001W_RptOrder.btnFieldList="+(String)session.getAttribute("btnFieldList"));
%>
<%}%>
getbtnFieldList();
function getbtnFieldList(){
    //95.08.23 add 縣市別/金融機構代號名稱放在已選的sort報表欄位===================================================    
    //this.document.forms[0].SortListDst.options[0] = new Option("縣市別", "hsien_id");
    //this.document.forms[0].SortListDst.options[1] = new Option("金融機構代號名稱", "bank_no");        
    //=============================================================================================================
    //根據所選加入欲sort的報表欄位跟已選sort的報表欄位
	var bnlist;
	var checkAdd = false;
	var addCount = 0;		
	bnlist = "<%=(String)session.getAttribute("btnFieldList")%>";//95.12.06 add	
	<%
	session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
	%>
	var a = bnlist.split(',');
	//95.08.23 add 若已選排序欄位無"縣市別"."金融機構代號名稱"時,才加入到可選排序欄位 by 2295 ==============================================
	//95.08.30 add 若已選排序欄位無"年"時,才加入到可選排序欄位 by 2295 		
	//alert(this.document.forms[0].BankList.value);
	//95.09.04 add 若已選的機構代號不為全部時，才可根據縣市別．金融機構代號sort	   
	//                               =全部時，根據年，縣市別小計sort
	//95.12.01 add 各別機構可根據縣市別.金融機構代號.年月sort
	//             全部可選根據縣市別小計.年月sort
	if(this.document.forms[0].BankList.value.indexOf("ALL") == -1){	   
	   //縣市別======================================================================================================
	   checkAdd = false;
       for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
	       if(this.document.forms[0].SortListDst.options[k].text == "縣市別"){		    
		      checkAdd = true;			       
	       }  
	   }  
	   if(checkAdd == false){	  
	      this.document.forms[0].SortListSrc.options[addCount] = new Option("縣市別", "wlx01.hsien_id");
	      addCount++;
	   }
	   //金融機構代號名稱============================================================================================
	   checkAdd = false;
	   for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
	       if(this.document.forms[0].SortListDst.options[k].text == "金融機構代號名稱"){		    
		    checkAdd = true;			       
	       }  
	   }    
       if(checkAdd == false){	  
          this.document.forms[0].SortListSrc.options[addCount] = new Option("金融機構代號名稱", "a01.bank_code");
          addCount++;
       }
       //年.月===================================================================================================
       checkAdd = false;
       for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
	       if(this.document.forms[0].SortListDst.options[k].text == "年月"){		    
		      checkAdd = true;			       
	       }  
	   } 	    
	   if(checkAdd == false){	  
	      this.document.forms[0].SortListSrc.options[addCount] = new Option("年月", "order_YYMM");
	      addCount++;
	   }	   
    }else{//已選的機構代號=ALL全部
       //縣市別小計============================================================================================
       checkAdd = false;
	   for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
	       if(this.document.forms[0].SortListDst.options[k].text == "縣市別小計"){		    
		      checkAdd = true;			       
	       }  
	   }  
	   if(checkAdd == false){	  
	      this.document.forms[0].SortListSrc.options[addCount] = new Option("縣市別小計", "cd01.FR001W_OUTPUT_ORDER");
	      addCount++;
	   }
	   //年.月===================================================================================================
	   checkAdd = false;
       for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
	       if(this.document.forms[0].SortListDst.options[k].text == "年月"){		    
		      checkAdd = true;			       
	       }  
	   } 	    
	   if(checkAdd == false){	  
	      this.document.forms[0].SortListSrc.options[addCount] = new Option("年月", "order_YYMM");
	      addCount++;
	   }	   	   
    }//end of 已選的機構代號=ALL全部
    
    //======================================================================================================
    //已選的報表欄位.95.12.01 不可根據已選欄位做sort
	/*for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		checkAdd=false;		
		for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
			if(this.document.forms[0].SortListDst.options[k].text == j[1]){		    			    
		   		checkAdd = true;			       
	    	}   
    	}    	
    	if(checkAdd == false){	       	       
    	    //if(j[1] == '單位數'){//95.09.05 add 單位數不可做為sort欄位
    	    //   continue;
    	    //}
       		this.document.forms[0].SortListSrc.options[addCount] = new Option(j[1], j[0]);
       		addCount++;
    	}	
	}*/	
}
-->
</script>
</body>
</html>
