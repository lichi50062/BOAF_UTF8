<%
// 97.08.20 create 縣市政府變現性資產查核 by 2295
// 98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295
//              查核項目加至200中文字/查核結果加至500中文字 by 2295
// 99.12.07 fix sqlInjection by 2808
//100.01.27 fix 日期顯示格式 by 2295
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
<%
	List BOAF_ASSETCHECK= (request.getAttribute("BOAF_ASSETCHECK")==null)?null:(List)request.getAttribute("BOAF_ASSETCHECK");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	 																										
	Properties permission = ( session.getAttribute("FX011W")==null ) ? new Properties() : (Properties)session.getAttribute("FX011W");
	if(permission == null){
       System.out.println("FX011W_Edit.permission == null");
    }else{
       System.out.println("FX011W_Edit.permission.size ="+permission.size());
    }
    
    String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
    String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");			
    String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
    String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
    System.out.println("muser_type="+muser_type);
    System.out.println("muser_tbank_no="+muser_tbank_no);
    System.out.println("muser_bank_type="+muser_bank_type);
    
    
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	
    String CHECK_DATE_Y="";
    String CHECK_DATE_M="";
    String CHECK_DATE_D="";
    String m_year="",bank_name="",examine="",name="",item="",check_date="",result="",content="",remark="";
    String sqlCmd = "";
    List paramList = new ArrayList() ;
	DataObject bean = null; 
	List date_tmp = null;//100.01.27    
    if(BOAF_ASSETCHECK!=null){
       bean = (DataObject)BOAF_ASSETCHECK.get(0);
       examine = bean.getValue("examine") == null?"":(String)bean.getValue("examine");
       bank_name = bean.getValue("bank_name") == null?"":(String)bean.getValue("bank_name");
       item = bean.getValue("item") == null?"":(String)bean.getValue("item");
       name = bean.getValue("name") == null?"":(String)bean.getValue("name");
       result = bean.getValue("result") == null?"":(String)bean.getValue("result");
       content = bean.getValue("content") == null?"":(String)bean.getValue("content");
       remark = bean.getValue("remark") == null?"":(String)bean.getValue("remark");
       m_year = bean.getValue("m_year") == null?"":bean.getValue("m_year").toString();
       check_date = bean.getValue("check_date") == null?"":bean.getValue("check_date").toString();
       check_date = (check_date.equals(""))?"-":Utility.getCHTdate(check_date.substring(0, 10),0);     				 	
		
	   //100.01.27 fix 日期顯示格式   				
       if(check_date.length() > 1){
    	   date_tmp = Utility.getStringTokenizerData(check_date,"/");//100.01.27 add
    	   CHECK_DATE_Y = (String)date_tmp.get(0);
    	   CHECK_DATE_M = (String)date_tmp.get(1);
    	   CHECK_DATE_D = (String)date_tmp.get(2);
       }	 
	   	 	
   }    
    /***
     * 99.12.08 fix 縣市合併年度區分
     **/
     String s_year = "99" ;
     if(Integer.parseInt(Utility.getYear()) > 99) {
     	s_year = "100" ;
     }
   sqlCmd = " select bn01.bank_no,bn01.bank_name"
		  + " from (select * from bn01 where m_year=?)bn01 left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no=wlx01.bank_no "
		  + " where bn01.bn_type <> '2'";
   paramList.add(s_year) ;
   paramList.add(s_year) ;
   if(muser_bank_type.equals("B")){//登入者為地方主管機關		  
	  sqlCmd += " and wlx01.m2_name=? ";   
      paramList.add(muser_tbank_no) ;
   }else{
      sqlCmd += " and wlx01.m2_name=? ";   
      paramList.add(bank_no) ;
   }		  
   List tbankNoList = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");  
   paramList.clear() ;
   
   sqlCmd = " select tbank_no,bank_no,bank_name"
		  + " from "
		  + " ("
		  + " select bank_no as tbank_no,bank_no,bank_name,'1' as ordertype "
		  + " from (select * from bn01 where m_year=?)bn01 where bn_type <> '2'"		  
		  + " union "
		  + " select tbank_no,bank_no,bank_name,'2' as ordertype "
		  + " from (select * from bn02 where m_year=? )bn02 where bn_type <> '2' "
		  + " ) "		  
		  + " order by tbank_no,ordertype ";  
   paramList.add(s_year) ;
   paramList.add(s_year) ;
   List bankNoList = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");   
   
    // XML Ducument for 分支機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"BankNoXML\">");
    out.println("<datalist>");
    for(int i=0;i< bankNoList.size(); i++) {
        bean =(DataObject)bankNoList.get(i);
        out.println("<data>");
        out.println("<bankType>"+bean.getValue("tbank_no")+"</bankType>");
        out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
        out.println("<bankName>"+bean.getValue("bank_no")+"  "+bean.getValue("bank_name")+"</bankName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 分支機構代碼 end 
	
%>

<script language="javascript" src="js/FX011W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>辦理農漁會信用部變現性資產查核</title>
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
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr>
   		 <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
      <tr>
        <td bgcolor="#FFFFFF" width="600">  
           <tr>
                <td width="600" height="18">
                 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="300" align="center"><b><font size="4">辦理農漁會信用部變現性資產查核</font></b></td>
                      <td width="150" align="right"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="600" height="50">
                <div align="left">
                <table width="600" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="600" >               
                      </td>
                    </tr>
                    <tr>
                      
                       <div>
                       <jsp:include page="getLoginUser.jsp?width=600" flush="true" />             			
                       </div>			 							
			 		   
                    </tr>
                <tr>
                <td width="600" height="300">
                <form method="post" name="outpushdata">
                <table width="600" border="0" align="center" cellpadding="0" cellspacing="0" height="2">
                    <tr>
                      <td  class="sbody" width="600" height="76">
                        <table height="300"  bordercolor="#3A9D99" border="1" width="600">                          
                            <tr class="sbody">
							<td align="left" width="171" bgColor="#d8efee">查核年度</td>
							<td width="416" bgcolor="#e7e7e7" height="1">
							<%if(act.equals("Edit")){%>
							 <%=m_year%>
							  <input type="hidden" name="m_year" value="<%=m_year%>">
							<%}else{%>
  								<input type="text" name="m_year" value='<%=m_year%>' size='3'>
  							<%}%>	
  							年   							 
  							</td>  								  							
							</tr>
							<%if(act.equals("new")){%>
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee">總機構單位</td>
							<td width="416" bgcolor="#e7e7e7" height="1">
								<select name='tbank_no_examine' onchange="javascript:changeOption(document.forms[0]);">                                                        
                            	 <%for(int i=0;i<tbankNoList.size();i++){%>
                           		 <option value="<%=(String)((DataObject)tbankNoList.get(i)).getValue("bank_no")%>"                            
                           			 <%if(bank_no.equals((String)((DataObject)tbankNoList.get(i)).getValue("bank_no"))) out.print("selected");%>
                           			 ><%=(String)((DataObject)tbankNoList.get(i)).getValue("bank_no")%><%=(String)((DataObject)tbankNoList.get(i)).getValue("bank_name")%></option>                            
                           		 <%}%>
                            	 </select>     															
  							</td>
							</tr>
							<%}%>
							<tr class="sbody">
							<td align="left" width="171" bgColor="#d8efee">受檢單位名稱</td>
							<td width="416" bgcolor="#e7e7e7" height="1">
							<%if(act.equals("Edit")){%>
							 <%=bank_name%>
							 <input type="hidden" name="examine" value="<%=examine%>">
							<%}else{%> 
  								<select size="1" name="examine">    							 
  								</select> 
  							<%}%>	
  							</td>  								  							
							</tr>							
							<tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="30" >查核日期
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="30">
                            <input type='hidden' name="CHECK_DATE" value="">
                            <input type='text' name="CHECK_DATE_Y" value="<%=CHECK_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年 
        						<select  name="CHECK_DATE_M">
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CHECK_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CHECK_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>月 
        						<select id="hide1" name="CHECK_DATE_D">
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CHECK_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CHECK_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>日</font> 
                            </td>   
                            </tr>
                            
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >查核人員
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">
                            <!--98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295-->
                            <textarea rows="4" cols="50" name='name'><%=name%></textarea>                           
                            </td>                               
                            </tr>
                            
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >查核項目
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">                            
                            <textarea rows="4" cols="50" name='item'><%=item%></textarea>
                            </td>   
                            </tr>                          
                            
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >查核結果
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">                            
                            <textarea rows="3" cols="50" name='result'><%=result%></textarea>
                            </td>   
                            </tr>     
                            
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >處理情形
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">                            
                            <textarea rows="3" cols="50" name='content'><%=content%></textarea>
                            </td>   
                            </tr>    
                            
                            <tr class="sbody">
                            <td align="left" width="171" bgColor="#d8efee" height="1" >備註
                            </td>
                            <td width="411" bgColor="#e7e7e7" height="2">                            
                            <textarea rows="3" cols="50" name='remark'><%=remark%></textarea>
                            </td>   
                            </tr>    
                        </table>
                      </td>
                    </tr>
                    <td width="600" height="21"><div align="right"><div align="right">
                    <div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>
                 </td>
             </table>
             </td></tr>
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
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				        <% } }else{
				        if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
				         <% }}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
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
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>      
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>      
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>                                
                          
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
</form>
</table>
<%if(act.equals("new")){%>
<script language="JavaScript" >
<!--
changeOption(this.document.forms[0]);
-->
</script>
<%}%>

</body>
