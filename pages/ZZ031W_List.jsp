<%
// 93.12.29 create by 2295
// 94.01.05 fix 不是農金局的,只能鎖定檢核成功的 by 2295
// 94.01.13 fix 除了bank_type=2/6/7/8/4可以查詢之外..其他的不行查 by 2295
// 94.02.04 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
// 94.02.14 fix A111111111可以看到全部 by 2295
//          fix 農漁會共用中心..可以看到其所屬的機構 by 2295
//          fix 下拉式選單使用xml by 2295
// 94.02.18 fix 按查詢時,傳輸項目會跳回原則初始值 by 2295
// 94.04.20 fix 檢核為0時.顯示內容值皆為零 by 2295
// 94.09.09 add 增加全國農業金庫,查詢A01-A05 by 2295
// 94.11.15 add 農.漁會可查F01在台無住所之外國人新台幣存款表 by 2295
// 95.02.09 add 全國農業金庫比照農金局進行鎖定.解除鎖定 by 2295
//          add 檢核結果增加"未申報"."申報內容均為0"."A01_逾放金額為0" by 2295
// 95.02.20 add 檢核結果增加"A04_申報逾放金額為0" by 2295
// 95.02.21 add 報表為A01時,檢核結果才"A01_逾放金額為0" by 2295
//          add 報表為A04時,檢核結果才"A04_逾放金額為0" 
// 95.11.07 add 增加使用者姓名.電話.機構電話 by 2295
// 98.10.13 add 產生申報者通訊錄 by 2295
//101.07.20 fix 傳輸項目.排列順序 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
    
	String bank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");				
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");				
	String sztrans_type = ( request.getParameter("TRANS_TYPE")==null ) ? "" : (String)request.getParameter("TRANS_TYPE");				
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String szreport_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String szupd_code = ( request.getParameter("upd_code")==null ) ? "" : (String)request.getParameter("upd_code");				
	String szlock_status = ( request.getParameter("lock_status")==null ) ? "" : (String)request.getParameter("lock_status");				
    String szbank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");				
    String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	System.out.println("act="+act);
	
	System.out.println("bank_type="+bank_type);
	System.out.println("sztrans_type="+sztrans_type);
	System.out.println("szreport_no="+szreport_no);
	System.out.println("S_YEAR="+S_YEAR);
	System.out.println("S_MONTH="+S_MONTH);
	System.out.println("szupd_code="+szupd_code);
	System.out.println("szlock_status="+szlock_status);
	

	String sqlCmd = "";
	List trans_typeList = null;
	
	if(bank_type.equals("2") || bank_type.equals("1")/*全國農業金庫*/ || lguser_id.equals("A111111111")){//農金局->可做A01-A05,M01-M08,B01-B03
	   sqlCmd = "select * from cdshareno where cmuse_div='011' order by input_order";
	}else if(bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("8") ){//農會.漁會.農漁會共用中心
	   sqlCmd = "select * from cdshareno where cmuse_div='011' and ";
	    if(bank_type.equals("6") || bank_type.equals("7")){//94.11.15 add農.漁會可查F01在台無住所之外國人新台幣存款表
	       sqlCmd += " (cmuse_id='1' or cmuse_id='4')" ;
	    }else{
	       sqlCmd += " cmuse_id='1' " ;
	    }
	    sqlCmd += " order by input_order";
	}else if(bank_type.equals("4")){//農業信用保証基金
	   sqlCmd = "select * from cdshareno where cmuse_div='011' and cmuse_id='2' order by input_order";
	}
	if(!sqlCmd.equals("")){
	    trans_typeList = DBManager.QueryDB_SQLParam(sqlCmd,null,"");
	}
		
	List lockList = (List)request.getAttribute("lockList");		
		
	if(lockList == null){
	   System.out.println("lockList == null");
	}else{
	   System.out.println("lockList.size()="+lockList.size());
	}
	
	//取得ZZ031W的權限
	Properties permission = ( session.getAttribute("ZZ031W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ031W"); 
	if(permission == null){
       System.out.println("ZZ031W_List.permission == null");
    }else{
       System.out.println("ZZ031W_List.permission.size ="+permission.size());               
    }	
    
    sqlCmd = " select * from cdshareno  where cmuse_div='012' or cmuse_div='013' or cmuse_div='014' or cmuse_div='030'"
           + " order by cmuse_div,to_number(input_order)";//101.07.20 fix 排列順序    		       
    List ReportList = DBManager.QueryDB_SQLParam(sqlCmd,null,"");                
    
    // XML Ducument for 總機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"ReportListXML\">");
    out.println("<datalist>");
    for(int i=0;i< ReportList.size(); i++) {
        DataObject bean =(DataObject)ReportList.get(i);
        out.println("<data>");
        if(((String)bean.getValue("cmuse_div")).equals("012")){
           out.println("<transType>1</transType>");        
        }else if(((String)bean.getValue("cmuse_div")).equals("013")){
           out.println("<transType>2</transType>");        
        }else if(((String)bean.getValue("cmuse_div")).equals("014")){
           out.println("<transType>3</transType>");        
        }else if(((String)bean.getValue("cmuse_div")).equals("030")){
           out.println("<transType>4</transType>"); //94.11.15 add '030'-->F01_在台無住所之外國人新台幣存款表     
        }
        if(((String)bean.getValue("cmuse_name")).indexOf("_") != -1){
           out.println("<reportValue>"+((String)bean.getValue("cmuse_name")).substring(0,((String)bean.getValue("cmuse_name")).indexOf("_"))+"</reportValue>");                              
        } 
        out.println("<reportName>"+bean.getValue("cmuse_name")+"</reportName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 總機構代碼 end 	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ031W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「申報資料追蹤管理」</title>
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
<form method=post action='#'>
<input type="hidden" name="act" value="">   
<%if(lockList != null && lockList.size() != 0){%>
<input type="hidden" name="row" value="<%=lockList.size()+1%>">   
<%}%>
<table width="600" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="235"><img src="images/banner_bg1.gif" width="235" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          「申報資料追蹤管理」 
                        </center>
                        </b></font> </td>
                      <td width="235"><img src="images/banner_bg1.gif" width="235" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=773" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                      <td><table width=773 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                                                    
                          <tr class="sbody">						  
						  <td width='15%' align='left' bgcolor='#D8EFEE'>傳輸類別</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <select name='TRANS_TYPE' onchange="javascript:changeOption(document.forms[0]);changeResult(this.document.forms[0]);">                                                        
                            <!--select name='TRANS_TYPE' onchange="javascript:getData(this.document.forms[0],'<%=act%>');"-->                                                        
                            <%
                            if(trans_typeList != null){
                             for(int i=0;i<trans_typeList.size();i++){%>
                            <option value="<%=(String)((DataObject)trans_typeList.get(i)).getValue("cmuse_id")%>"                                                        
                            <%if(sztrans_type.equals((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)trans_typeList.get(i)).getValue("cmuse_name")%></option>                            
                            <%}
                            }//end of if
                            %>
                            </select>    
                            
                            &nbsp;
                            <%if(bank_type.equals("1") || bank_type.equals("2") || bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("8") || bank_type.equals("4")){%>
                            <a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                                                   
                            <%}%>
                            <%//if(act.equals("Qry") && !bank_type.equals("1")/*全國農業金庫*/){
                            if(act.equals("Qry")){//95.02.09 add 全國農業金庫可進行鎖定/解除鎖定
                            %>
                            <a href="javascript:doSubmit(this.document.forms[0],'Lock');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_lockb.gif',1)"><img src="images/bt_lock.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>                       
 							<a href="javascript:doSubmit(this.document.forms[0],'unLock');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_nolockb.gif',1)"><img src="images/bt_nolock.gif" name="Image103" width="80" height="25" border="0" id="Image103"></a>                                                     						
 							<%}%>
 							<input type="button" name="print" value="列印申報通訊錄" onClick="javascript:doSubmit(this.document.forms[0],'printUserData');"><!-- 98.10.13 add 產生申報者通訊錄-->
                          </td>             
                          </tr>       
                                                         
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>傳輸項目</td>
                          <td width='85%' bgcolor='e7e7e7'>
                            <select name='REPORT_NO' onchange="javascript:changeResult(this.document.forms[0]);">                                                                                    
                            </select >                                 
                          </td>         
                          </tr>
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' bgcolor='#D8EFEE'>總機構代號</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <input type="text" name="TBANK_NO" value="<% if(!szbank_code.equals("")){ out.print(szbank_code);}else if((bank_type.equals("6") || bank_type.equals("7")) && (!lguser_id.equals("A111111111")))  {out.print(tbank_no);}%>"  <%if((bank_type.equals("6") || bank_type.equals("7")) && (!lguser_id.equals("A111111111"))){out.print("disabled");}%> "size="7" maxlength="7" >                            
                            <input type="hidden" name="TBANK_NO" value="<%if(bank_type.equals("6") || bank_type.equals("7")) {out.print(tbank_no);}%>" >                            
                          </td>                                   
                          </tr>
                          
                          
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' bgcolor='#D8EFEE'>申報基準年月</td>
                          <td width='85%' bgcolor='e7e7e7'>	
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
                          </td>                                   
                          </tr>
                          
                          
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>檢核結果</td>
                          <td width='85%' bgcolor='e7e7e7'>                          
                            <select name='UPD_CODE'>                                                                                                                
                            <option value="ALL" <%if(szupd_code.equals("ALL")) out.print("selected");%>>全部</option>                            
                            <option value="U" <%if(szupd_code.equals("U")) out.print("selected");%>>成功</option>                            
                            <option value="E" <%if(szupd_code.equals("E")) out.print("selected");%>>錯誤</option>         
                            <option value="N" <%if(szupd_code.equals("N")) out.print("selected");%>>未申報</option>                                               
                            <option value="Z" <%if(szupd_code.equals("Z")) out.print("selected");%>>申報內容均為0</option>                            
                            <%if(szreport_no.equals("A01")){%>
                            <option value="A01_990000_0" <%if(szupd_code.equals("A01_990000_0")) out.print("selected");%>>A01_申報逾放金額為0</option>                            
                            <%}%>
                            <%if(szreport_no.equals("A04")){%>
                            <option value="A04_840740_0" <%if(szupd_code.equals("A04_840740_0")) out.print("selected");%>>A04_申報逾放金額為0</option>                                 
                            <%}%>
                          </td>         
                          </tr>
                          
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>鎖定與否</td>
                          <td width='85%' bgcolor='e7e7e7'>                          
                            <select name='LOCK_STATUS'>                                                                                    
                            <option value="ALL" <%if(szlock_status.equals("ALL")) out.print("selected");%>>全部</option>                            
                            <option value="Y" <%if(szlock_status.equals("Y")) out.print("selected");%>>鎖定</option>                            
                            <option value="N" <%if(szlock_status.equals("N")) out.print("selected");%>>未鎖定</option>                            
                            </select>                                 
                            
                          </td>         
                          </tr>                                                    		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=773 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                      
                         <%//95.02.09if(!bank_type.equals("1")/*全國農業金庫*/){%>			
                        <tr class="sbody">
                          <td width='100%' colspan=13 bgcolor='D2F0FF'>		                         	 						  
 							<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						  
                          </td>
                        </tr>     		
                       <%//}%>   
                      <%
                      String tmpbank_type="";                      
                      int i = 0; 
                      int lockListidx=0;     
                      String bgcolor="#D3EBE0";     
                      String upd_code = "";
                      String input_method ="";
                      String upd_method = "";
                      if(lockList != null){ %>
                       	  <tr class="sbody" bgcolor="#9AD3D0">                       	  
                            <td width="30">序號</td>                                  
                            <td width="30">選項</td>                            
                            <td width="120">總機構代碼</td>
            				<td width="60">申報年月</td>            				
            				<td width="60">檢核結果</td>
            				<td width="60">申報方式</td>            				
            				<td width="30">鎖定</td>            				            				
            				<td width="30">代傳</td>            				            				
            				<td width="60">更新方式</td>            				            				            				
            				<td width="60">局內鎖定</td>   
            				<td width="86">使用者姓名</td>   
					        <td width="79">電話</td>   
					        <td width="68">機構電話</td>           				            				
					      </tr>   
                   		    <% if(lockList.size() == 0){%>
                   			   <tr class="sbody" bgcolor="<%=bgcolor%>">
                   			   <td colspan=13 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                   			   String upd_code_tmp = "";
                    		   while(lockListidx < lockList.size()){ 
                    		      upd_code_tmp = "";
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
                    		      if (((String)((DataObject)lockList.get(lockListidx)).getValue("input_method")) == null){
										input_method = "&nbsp;";
					  			  }else if (((String)((DataObject)lockList.get(lockListidx)).getValue("input_method")).equals("F")){
										input_method = "檔案上傳";
					  			  }else if (((String)((DataObject)lockList.get(lockListidx)).getValue("input_method")).equals("W")){
										input_method = "線上編輯";
					  			  }else{
										input_method = "";
					  			  }		
					  			  upd_code = (String)((DataObject)lockList.get(lockListidx)).getValue("upd_code");
					  			  upd_code_tmp = upd_code;
					  			  if (upd_code == null){
					      			  upd_code = "待檢核";
					  			  }else if (upd_code.equals("E")){
					      			  upd_code = "檢核有誤";
					  			  }else if (upd_code.equals("U")){
						  			  upd_code = "檢核成功";
					  			  }else if (upd_code.equals("F")){
					      			  upd_code = "上傳檔案不存在";
					  			  }else if (upd_code.equals("Z")){
					      		      upd_code = "內容值皆為零";//94.04.20檢核為0時.顯示內容值皆為零   					  
					              }
					              
					  			  if(((String)((DataObject)lockList.get(lockListidx)).getValue("havingdata")).equals("N")){					  			    
					  			     upd_code = "未申報";
					  			  }
					  			  if(szupd_code.equals("N")){//95.02.09 add 檢核結果為未申報,先抓全部,只顯示結果為未申報的
					  			     //System.out.println("bank_no="+(String)((DataObject)lockList.get(lockListidx)).getValue("bank_no"));
					  			     //System.out.println("upd_code="+upd_code);
					  			     //System.out.println("havingdata="+((String)((DataObject)lockList.get(lockListidx)).getValue("havingdata")));					  			  
					  			     if(!upd_code.equals("未申報")){
					  			         lockListidx++;
					  			         continue;
					  			     }   
					  			  }		
                    		      upd_method = (String)((DataObject)lockList.get(lockListidx)).getValue("upd_method");
								  if(upd_method == null){
								     upd_method = "&nbsp;";
								  }else if(upd_method.equals("A")){
					      			  upd_method = "自動";
					  			  }else if (upd_method.equals("M")){
						  			  upd_method = "人工";
                    		      }
                    		      
                    		                          		      
                      %>               

					  		  	
					  	
          	  
                          <tr class="sbody" bgcolor="<%=bgcolor%>">
                            <td width="30"><%=i+1%></td>                       				            				
            				<td width="30">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%if( ((DataObject)lockList.get(lockListidx)).getValue("bank_no") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("bank_no"));%>"
            				<%if(bank_type.equals("8") || bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("4")){
            				     if( ((DataObject)lockList.get(lockListidx)).getValue("wml01_lock_status") != null  && ((String)((DataObject)lockList.get(lockListidx)).getValue("wml01_lock_status")).equals("Y")) out.print("disabled");            				     
            				  } 
            				  if(!(bank_type.equals("2") || bank_type.equals("1"))){/*不是農金局/95.02.09 add 全國農業金庫的,只能鎖定檢核成功的*/              				      
            				      if(upd_code_tmp == null || ( upd_code_tmp != null && (!upd_code_tmp.equals("U")))){
            				         out.print("disabled");;            				        
            				      }               				      
            				  }   
            				%>
            				>
            				</td>            				
            				<td width="120">
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("bank_no") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("bank_no")); else out.print("&nbsp;");%>
            				<br>
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("bank_name") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("bank_name")); else out.print("&nbsp;");%>
            				</td>
            				<td width="60">&nbsp;
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("m_year") != null ) out.print((((DataObject)lockList.get(lockListidx)).getValue("m_year")).toString()+"/");%>
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("m_month") != null ) out.print((((DataObject)lockList.get(lockListidx)).getValue("m_month")).toString());%>
            				</td>
            				<td width="60"><%=upd_code%></td>
            				<td width="60"><%=input_method%></td>            				
            				<td width="30"><%if( ((DataObject)lockList.get(lockListidx)).getValue("lock_status") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("lock_status")); else out.print("&nbsp;");%></td>
            				<td width="30"><%if( ((DataObject)lockList.get(lockListidx)).getValue("common_center") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("common_center")); else out.print("&nbsp;");%></td>            				
            				<td width="60"><%=upd_method%></td>            				
            				<td width="60"><%if( ((DataObject)lockList.get(lockListidx)).getValue("wml01_lock_status") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("wml01_lock_status")); else out.print("&nbsp;");%></td>            				
            				<td width="86">
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("muser_name") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("muser_name")); else out.print("&nbsp;");%>            				
            				</td>
            				<td width="79">
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("m_telno") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("m_telno")); else out.print("&nbsp;");%>            				
            				</td>
            				<td width="69">
            				<%if( ((DataObject)lockList.get(lockListidx)).getValue("telno") != null ) out.print((String)((DataObject)lockList.get(lockListidx)).getValue("telno")); else out.print("&nbsp;");%>            				
            				</td>   
					      </tr> 					      
					      <%
                  			   i++;
                  			   lockListidx++;
	                  		   }//end of while
	                  		}//end of if
    			          %>  
                  		  
					      </table>      
                      </td>    
                      </tr>
                                  
      </table></td>
  </tr> 
</table>
</form>
</body>
<script language="JavaScript" >
<!--
setSelect(this.document.forms[0].TRANS_TYPE,"<%=sztrans_type%>");
changeOption(this.document.forms[0]);
setSelect(this.document.forms[0].REPORT_NO,"<%=szreport_no%>");
-->
</script>

</html>
