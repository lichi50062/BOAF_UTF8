<%
// 96.11.12 add 增加委外項目.委外範圍 by 2295
// 97.07.09 fix 結束日期.可不輸入 by 2295
//100.01.26 fix 日期顯示格式 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	List WLX06_M_OUTPUSH= (request.getAttribute("WLX06_M_OUTPUSH")==null)?null:(List)request.getAttribute("WLX06_M_OUTPUSH");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	Properties permission = ( session.getAttribute("FX006W")==null ) ? new Properties() : (Properties)session.getAttribute("FX006W");
	if(permission == null){
       System.out.println("FX006W_List.permission == null");
    }else{
       System.out.println("FX006W_List.permission.size ="+permission.size());
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

    String seq_no="";
    String companyname="";
    String contractname="";
    String contracttel="";
    String complainname="";
    String complaintel="";
    String beg_date="";
    String end_date="";
    String comment="";
    String BEG_DATE_Y="";
    String BEG_DATE_M="";
    String BEG_DATE_D="";
    String END_DATE_Y="";
    String END_DATE_M="";
    String END_DATE_D="";
    String out_item="";//委外項目
    String out_range="";//委外範圍
	List date_tmp = null;
	if(WLX06_M_OUTPUSH!=null){
       seq_no = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("seq_no")); 
 	   companyname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("outcompanyname"));
	   contractname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("outcontractname"));
	   contracttel = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("outcontracttel"));														
	   complainname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("bankcomplainname"));
	   complaintel = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("bankcomplaintel"));
	   comment = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("outcomment"));		
	   beg_date = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_begin_date"));	
	   end_date = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_end_date"));	
	   
	   beg_date = (beg_date.equals("null"))?"-":Utility.getCHTdate(beg_date.substring(0, 10),0);     				
	   end_date = (end_date.equals("null"))?"-":Utility.getCHTdate(end_date.substring(0, 10),0);  
	   //100.01.26 fix 日期顯示格式   				
	   if(beg_date.length() > 1){
	      date_tmp = Utility.getStringTokenizerData(beg_date,"/");//100.01.26 add
	      BEG_DATE_Y = (String)date_tmp.get(0);
	      BEG_DATE_M = (String)date_tmp.get(1);
	      BEG_DATE_D = (String)date_tmp.get(2);
	   }
	  		
	   if(end_date.length() > 1){
	      date_tmp = Utility.getStringTokenizerData(end_date,"/");//100.01.26 add
	      END_DATE_Y = (String)date_tmp.get(0);
	      END_DATE_M = (String)date_tmp.get(1);
	      END_DATE_D = (String)date_tmp.get(2);
	   }
	     
	   
	   //96.11.09 add 委外項外.委外範圍
	   out_item = (((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_item")==null)?"": (String)((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_item");	
	   out_range = (((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_range")==null)?"": (String)((DataObject)WLX06_M_OUTPUSH.get(0)).getValue("out_range");	
    }//end of WLX06_M_OUTPUSH!=null

	comment=(comment.equals("null"))?"":comment;
	
	
	String sqlCmd = " select cmuse_id,cmuse_name"
			    + " from cdshareno "
			    + " where cmuse_div='031' "
			    + " order by cmuse_id ";
    	
  	List out_item_desc_List = DBManager.QueryDB_SQLParam(sqlCmd,null,"cmuse_id,cmuse_name");  
  	DataObject bean = null;
    String out_item_desc = "";
%>

<script language="javascript" src="js/FX006W.js"></script>
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
function do_submit(){
 var newwin = window.open("FX006W_out_item.jsp", "QueryOUT_ITEM", "height=640,width=500,resizable=yes,scrollbars=yes");
        if ( newwin!= null)
            newwin.focus(); // 讓window浮到上層
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
                      <td width="300" align="center"><b><font size="4">各農漁會委外內部作業資料申報</font></b></td>
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
                        <table height="290"  bordercolor="#3A9D99" border="1">
                        
                         <%
        				  if(!seq_no.equals("")){
        				 %>
        					<input type="hidden" name="editseq_no" value="<%=seq_no%>">
        				 <%}%>
                        
                            <tr class="sbody" >
                            <td align="left" width="172" bgColor="#d8efee" colSpan="2" height="21">金融機構代號</td>
                            <td width="412" bgColor="#e7e7e7" height="21"> <%=bank_no%></td>
                        	</tr>
                        	<tr class="sbody" >
                            <td align="left" width="172" bgColor="#d8efee" colSpan="2" height="21">委外項目</td>
                            <td width="412" bgColor="#e7e7e7" height="21">                            
                                <select  name="out_item">
                                <% for(int i=0;i<out_item_desc_List.size();i++){
        							   bean = (DataObject)out_item_desc_List.get(i);
        						       out_item_desc = (String)bean.getValue("cmuse_name");
  							    %>
                                    <option value="<%=(String)bean.getValue("cmuse_id")%>" <%if(out_item.equals((String)bean.getValue("cmuse_id"))) out.print("selected");%>><%=out_item_desc.substring(0,out_item_desc.indexOf("、"))%></option>        
        						<%}%>    
        						</select>
        						<a href="javascript:do_submit();">
        						  <img src="images/ico.gif" alt="委外項目說明" align='absmiddle' border=0>
        						</a>
                            </td>
                        	</tr>
                            <tr class="sbody">
                            <td align="left" width="70" bgColor="#d8efee" height="123" rowspan="5">受委託機構
                            </td>
                            <td align="left" width="104" bgColor="#d8efee" height="23">機構名稱</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                            <input maxLength="100" size="32" name="companyname" value="<%=companyname%>">              
                            <font color="red">*</font></td>     
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="104" bgColor="#d8efee" height="23">聯絡人</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                            <input maxLength="20" name="contractname" size="20" value="<%=contractname%>">               
                            <font color="red">*</font></td>     
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="104" bgColor="#d8efee" height="23">聯絡電話</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                            <input maxLength="60" name="contracttel" size="20" value= "<%=contracttel%>" >               
                            <font color="red">*</font></td>     
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="104" bgColor="#d8efee" height="23">起始日期</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                                <input type='hidden' name="BEG_DATE" value="">
                            <input type='text' name="BEG_DATE_Y" value="<%=BEG_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年 
        						<select  name="BEG_DATE_M">
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(BEG_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(BEG_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>月 
        						<select id="hide1" name="BEG_DATE_D">
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(BEG_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(BEG_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>日</font>             
                            <font color="red">*</font></td>    
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="104" bgColor="#d8efee" height="23">結束日期</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                           <input type='hidden' name="END_DATE" value="">
                            <input type='text' name="END_DATE_Y" value="<%=END_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年 
        						<select  name="END_DATE_M">
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(END_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(END_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>月 
        						<select id="hide1" name="END_DATE_D">
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(END_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(END_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%> 
        						</select></font><font color='#000000'>日</font>             
                            <!--font color="red">*</font--></td>     
                            </tr>
                           
                           
                            <tr class="sbody">
                               <td align="left" width="172" bgColor="#d8efee" colSpan="2" height="47">委外事項範圍</td>
                               <td width="420" bgColor="#e7e7e7" height="47"><textarea rows="3" name="out_range" cols="54"><%=out_range%></textarea></td>                          	
                            </tr>
                           
                           
                            <tr class="sbody">
                            <td align="left" width="58" bgColor="#d8efee" height="61" rowSpan="2">信用部
                            <p>申訴窗口</p>
                            </td>
                            <td align="left" width="104" bgColor="#d8efee" height="23">聯絡人</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                            <input maxLength="60" name="complainname" size="20" value="<%=complainname%>">
                            <font color="red">*</font></td>         
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="104" bgColor="#d8efee" height="36">專線電話</td>
                            <td width="420" bgColor="#e7e7e7" height="36">
                            <input maxLength="60" size="35" name="complaintel" value="<%=complaintel%>"> 
                            
                            <font color="red">*</font></td>     
                            </tr>
                            <tr class="sbody">
                            <td align="left" width="172" bgColor="#d8efee" colSpan="2" height="23">   
                             備註</td>
                            <td width="420" bgColor="#e7e7e7" height="23">
                            <input maxLength="100" size="47" name="comment" value="<%=comment%>"></td>
                          	
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
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				        <% } }else{
				        if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
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
                      <font color="#FF0000" >
                      <ul>
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>      
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>      
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>      
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報的數字欄位仍請填「0」</li>
                        </ul>
                            </font>
                            </font>
                          </font>
                       </font>
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
