<%
//104.10.12 create by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	
   	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();   	
	String act = ( request.getParameter("act")==null ) ? "new" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	System.out.println("WMFileEdit_A13.act="+act);
	System.out.println("S_YEAR="+S_YEAR);	
	System.out.println("S_MONTH="+S_MONTH);	
	System.out.println("bank_type="+bank_type);		
	List data_div01 = null;
	List data_div02 = null;
	String ncacno = "ncacno";
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String Report_no="A13";	
	String acc_div = "";
	
	if(act.equals("new")){	  
		acc_div = "12";
	    data_div01 = Utility.getAcc_Code(ncacno,Report_no,acc_div); //建築貸款
	      
	    acc_div = "13";
	    data_div02 = Utility.getAcc_Code(ncacno,Report_no,acc_div); //農舍農地 
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
		data_div02 = (List)request.getAttribute("data_div02");
	}
	
	System.out.println("data_div01.size="+data_div01.size());
	System.out.println("data_div02.size="+data_div02.size());
	//取得WMFileEdit的權限
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A13.permission == null");
    }else{
       System.out.println("WMFileEdit_A13.permission.size ="+permission.size());
               
    }
%>
<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
</style>
<%if(act.equals("Query")){%>
<title>申報資料查詢</title>
<%}else{%>
<title>線上編輯申報資料</title>
<%}%>

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
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<script language="JavaScript" src="js/menu.js"></script>
<script language="JavaScript">
showToolbar();
</script>
<script language="JavaScript">
function UpdateIt(){
if (document.all){
document.all["MainTable"].style.top = document.body.scrollTop;
setTimeout("UpdateIt()", 200);
}
}
UpdateIt();
</script>
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
<input type="hidden" name="act" value="">  
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                      <td width="400"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A13")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                          
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="112"> <div align=left>基準日</div></td>
                            <td colspan=2>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' <%if(act.equals("Edit")) out.print("disabled");%>>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH <%if(act.equals("Edit")) out.print("disabled");%>>
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
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody" bgcolor='#D8EFEE'> 
                            <td width="112"> <div align=left>申報資料</div></td>
                            <td colspan=2>A13&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A13")%></td>
                          </tr>  
                          <%}%>    
                          <tr bgcolor='e7e7e7'>		
							<td colspan=3><b><div align=center><a name="A13_div01">建築貸款</a></div></b></td>		
						  </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td width=111 bgcolor="#D8EFEE"> <div align=left>項目代碼</div></td>
                            <td width="294"> <div align=left>項目名稱</div></td>
                            <td width=177 bgcolor="#B1DEDC"> <div align=left>項目數值</div></td>
                          </tr>
 <% 	int i = 0 ;
		boolean fontbold=false;
		String bgcolor="#F2F2F2";
		//String fontsize="2";		
		while( i < data_div01.size()){
			//bgcolor = (i % 2 == 0)?"#F2F2F2":"#FFFFE6";			 
			bgcolor="#F2F2F2";
			fontbold=false;
			//fontsize="2";
		    String tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
			
			//建築貸款
			if(tmpAcc_Code.equals("995200") || tmpAcc_Code.equals("995210")			
			|| tmpAcc_Code.equals("995361") || tmpAcc_Code.equals("995300")		
			|| tmpAcc_Code.equals("995400")
			){ 			
				//fontbold=true;
				//fontsize="4";
				bgcolor = "#FFFFE6";			 
			}
			
%>			
			<tr bgcolor='<%=bgcolor%>' class="sbody">
	
			<td bgcolor="<%=bgcolor%>">			
			<%if(fontbold){%><b><%}%>				
			<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
			<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
			<input type=hidden name=acc_name value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>">		
			<input type=hidden name=acc_div value="01">
			
			</td>
			<td>
			
			<%if(fontbold){%><b><%}%>	

			<div align=left>	
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>		
			</div>
			<%if(fontbold){%></b><%}%>		
			
			</td>
			
			<td><a name="<%=tmpAcc_Code%>">
			<%
			     //fix 94.02.21 若金額為0時,顯示空白 by 2295
			     if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){%>
			       <input align='right' type='text' name='amt' value="" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this);' style='text-align: right;'>   
			<%   }else{ 
				    if(tmpAcc_Code.equals("995400")){%>
				     <input type='text' name='amt' value="<%=Utility.getPercentNumber(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this);' style='text-align: right;'>
			<%  	}else{	%> 
			         <input align='right' type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this);' style='text-align: right;'>
			<%      }
			     }//end of amt ? 0 )			  
			%>
			</a>
			
			</td>		
		
		</tr>
	
		
	<%	    i++;	
		}
	%>
	
			<tr bgcolor='e7e7e7' class="sbody">	
				<td colspan=3><b><div align=center><font size="4"><a name="A13_div02">農地農舍</a></font></div></b></td>		
			</tr>
			<tr bgcolor='e7e7e7' class="sbody"> 
                <td width=111 bgcolor="#D8EFEE"> <div align=left>項目代碼</div></td>
                <td width="294"> <div align=left>項目名稱</div></td>
                <td width=177 bgcolor="#B1DEDC"> <div align=left>項目數值</div></td>
            </tr>
            <%  i = 0 ;
		fontbold=false;
		//fontsize="2";
		
		while( i < data_div02.size()){
			fontbold=false;
			//fontsize="2";
		    //bgcolor = (i % 2 == 0)?"#F2F2F2":"#FFFFE6";			 
		    bgcolor="#F2F2F2";
		    String tmpAcc_Code = ((String)((DataObject)data_div02.get(i)).getValue("acc_code")).trim();
		    
		    
		    //農地農舍
			if(tmpAcc_Code.equals("993210") || tmpAcc_Code.equals("993810")			
			){			
				bgcolor = "#FFFFE6";	
			}			
	%>	
	<tr bgcolor='<%=bgcolor%>' class="sbody">
	
		<td bgcolor="<%=bgcolor%>">			
		<%if(fontbold){%><b><%}%>		

		<div align=left><%=(String)((DataObject)data_div02.get(i)).getValue("acc_code")%></div>		
		<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div02.get(i)).getValue("acc_code")%>">        
        <input type=hidden name=acc_name value="<%=(String)((DataObject)data_div02.get(i)).getValue("acc_name")%>">
		<input type=hidden name=acc_div value="01">
		
		</td>
		<td>
		
		<%if(fontbold){%><b><%}%>	
		
		<div align=left>
		<%=(String)((DataObject)data_div02.get(i)).getValue("acc_name")%>		
		</div>
		<%if(fontbold){%></b><%}%>		
		
		</td>
		
		<td><a name="<%=tmpAcc_Code%>">	
		<%
		     //fix 94.02.21 若金額為0時,顯示空白 by 2295
			 if( ((DataObject)data_div02.get(i)).getValue("amt") == null ||  (((DataObject)data_div02.get(i)).getValue("amt") != null && ((((DataObject)data_div02.get(i)).getValue("amt")).toString()).equals("0")) ){%>
			     <input align='right' type='text' name='amt' value="" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'>
	    <%   }else{
	    		if(tmpAcc_Code.equals("993810")){%>
				 <input type='text' name='amt' value="<%=Utility.getPercentNumber(((((DataObject)data_div02.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div02.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this);' style='text-align: right;'>
		<% 	    }else{%> 		   
		         <input align='right' type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div02.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div02.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'>
		<%      } 
		     }//end of amt ? 0		 
		%>

		</a>	
		
		</td>		
	</font>	
	</tr>
	
		
	<%	    i++;	
		}
	%>
	
	          </Table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>              
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A13','','','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A13','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A13','<%=S_YEAR%>','<%=S_MONTH%>','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
                <%}%>        
                        <td width="80"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a></div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A13")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【確定】</font>或<font color="#666666">【修改】</font>時,會一併執行線上檢核,需耗時5-7秒。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
            </table></td>
        </tr>        
      </table></td>
  </tr>
</table>
</form>
</body>

</html>
