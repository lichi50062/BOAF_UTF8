<%
//95.02.16 create by 2295
//95.02.20 add 使用說明 by 2295
//         add A01、A04及A05_BIS 如果有申報金額0仍要印出0，NULL 才印出空白 by 2295
//95.05.05 add A01與A04總放款金額不符/A01與A04總放款金額均為0/A01或A04總放款金額尚未申報 by 2295
//95.08.16 add 增加使用者姓名.電話.機構電話.拿掉裁撤日期 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
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
    String szbank_type_list = ( request.getParameter("BANK_TYPE_List")==null ) ? "" : (String)request.getParameter("BANK_TYPE_List");				
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String szupd_code = ( request.getParameter("upd_code")==null ) ? "" : (String)request.getParameter("upd_code");				
	System.out.println("act="+act);
	
	System.out.println("ZZ033W_List.bank_type_list="+szbank_type_list);	
	System.out.println("ZZ033W_List.S_YEAR="+S_YEAR);
	System.out.println("ZZ033W_List.S_MONTH="+S_MONTH);
	System.out.println("ZZ033W_List.szupd_code="+szupd_code);
	
	String sqlCmd = "";
		
	List zz033wList = (List)request.getAttribute("zz033wList");		
		
	if(zz033wList == null){
	   System.out.println("zz033wList == null");
	}else{
	   System.out.println("zz033wList.size()="+zz033wList.size());
	}
	
	//取得ZZ033W的權限
	Properties permission = ( session.getAttribute("ZZ033W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ033W"); 
	if(permission == null){
       System.out.println("ZZ033W_List.permission == null");
    }else{
       System.out.println("ZZ033W_List.permission.size ="+permission.size());               
    }	
    
   
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「逾放金額及BIS申報管理」</title>
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
function doSubmit(form,cnd){          
	     if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		      form.S_YEAR.focus();
		      return;
	     }  	    		     
	     form.action="/pages/ZZ033W.jsp?act="+cnd+"&test=nothing";	    	    
	     if( cnd == "List") form.submit();	    		    
}	
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">   
<%if(zz033wList != null && zz033wList.size() != 0){%>
<input type="hidden" name="row" value="<%=zz033wList.size()+1%>">   
<%}%>
<table width="885" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="885" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="885" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="292"><img src="images/banner_bg1.gif" width="292" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          「逾放金額及BIS申報管理」 
                        </center>
                        </b></font> </td>
                      <td width="292"><img src="images/banner_bg1.gif" width="292" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="885" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=885" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                       <td ><table width=885 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                          <tr class="sbody">						  
						  <td width='15%' align='left' bgcolor='#D8EFEE'>農(漁)會別</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <select name='BANK_TYPE_List' >                                                        
                            <option value="6" <%if(szbank_type_list.equals("6")) out.print("selected");%>>農會</option>
                            <option value="7" <%if(szbank_type_list.equals("7")) out.print("selected");%>>漁會</option>
                            <option value="ALL" <%if(szbank_type_list.equals("ALL")) out.print("selected");%>>農漁會</option>
                            </select>     
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:doSubmit(this.document.forms[0],'List');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                                                  
                          </td>             
                          </tr>       
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' bgcolor='#D8EFEE'>查詢年月</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%><%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%><%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
                          </td>                                   
                          </tr>
                          
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>檢核結果</td>
                          <td width='85%' bgcolor='e7e7e7'>                          
                            <select name='UPD_CODE' >                                                                                                                
                            <option value="ALL" <%if(szupd_code.equals("ALL")) out.print("selected");%>>全部</option>                            
                            <option value="0" <%if(szupd_code.equals("0")) out.print("selected");%>>有任一不符</option>                            
                            <option value="1" <%if(szupd_code.equals("1")) out.print("selected");%>>A01與A04申報逾放金額不符</option>                                
                            <option value="1a" <%if(szupd_code.equals("1a")) out.print("selected");%>>A01與A04放款總額不符</option>                                
                            <option value="2" <%if(szupd_code.equals("2")) out.print("selected");%>>A01與A04申報逾放金額均為0</option>                                                                                                                
                            <option value="2a" <%if(szupd_code.equals("2a")) out.print("selected");%>>A01與A04放款總額均為0</option>                                                                                                                
                            <option value="3" <%if(szupd_code.equals("3")) out.print("selected");%>>A01或A04逾放金額尚未申報</option>                                                                                                                
                            <option value="3a" <%if(szupd_code.equals("3a")) out.print("selected");%>>A01或A04放款總額尚未申報</option>                                                                                                                
                            <option value="4" <%if(szupd_code.equals("4")) out.print("selected");%>>A05_BIS未申報</option>                                                                                                                
                            <option value="5" <%if(szupd_code.equals("5")) out.print("selected");%>>A05_BIS申報值為0</option>                                                                                                                
                            </select>                                 
                          </td>         
                          </tr>                                		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=885 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                         <tr class="sbody" bgcolor="#9AD3D0">    
                            <td width="20">序號</td>                                                              
                            <td width="220">總機構代號名稱</td>            				       				            				
        				    <!--td width="71">裁撤日期</td-->            				            				            				
        				    <td width="90">A01_逾放金額</td>
        				    <td width="90">A04_逾放金額</td>
        				    <td width="90">A01_放款總額</td>            				            				            				
        				    <td width="90">A04_放款總額</td>            				            				            				
					        <td width="52">A05_BIS</td>   
					        <td width="86">使用者姓名</td>   
					        <td width="79">電話</td>   
					        <td width="68">機構電話</td>  
					      </tr> 

					       <%
		                      int i = 0;      
        		              String bgcolor="#D3EBE0";                     		      
                   		      if((zz033wList == null) || (zz033wList.size() == 0)){%>
                   			   <tr class="sbody" bgcolor="<%=bgcolor%>">
                   			   <td colspan=22 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<%}else{                   			                      			   
                   			   String a01_amt_990000="";
                   			   String a01_amt_total="";
                   			   String a04_amt_840740="";
                   			   String a04_amt_840750="";
                   			   String a05_amt="";
                   			   while(i < zz033wList.size()){                     		      
                    		      a01_amt_990000="";
                    		      a01_amt_total="";
                   			      a04_amt_840740="";
                   			      a04_amt_840750="";
                   			      a05_amt="";
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	                    		                          		      
                    		      if( ((DataObject)zz033wList.get(i)).getValue("a01_amt_990000") != null ){
                    		        a01_amt_990000 = (((DataObject)zz033wList.get(i)).getValue("a01_amt_990000")).toString();
                    		      }
                    		      if( ((DataObject)zz033wList.get(i)).getValue("a01_amt_total") != null ){
                    		        a01_amt_total = (((DataObject)zz033wList.get(i)).getValue("a01_amt_total")).toString();
                    		      }
                    		      if( ((DataObject)zz033wList.get(i)).getValue("a04_amt_840740") != null ){
                    		        a04_amt_840740 = (((DataObject)zz033wList.get(i)).getValue("a04_amt_840740")).toString();
                    		      }
                    		      if( ((DataObject)zz033wList.get(i)).getValue("a04_amt_840750") != null ){
                    		        a04_amt_840750 = (((DataObject)zz033wList.get(i)).getValue("a04_amt_840750")).toString();
                    		      }
                    		      if( ((DataObject)zz033wList.get(i)).getValue("a05_amt") != null ){
                    		        a05_amt = (((DataObject)zz033wList.get(i)).getValue("a05_amt")).toString();
                    		      }
                    		   
                      %>                         	  
                          <tr class="sbody" bgcolor="<%=bgcolor%>">
                            <td width="20"><%=i+1%></td>                       				            				            				            				
            				<td width="220">
            				<%if( ((DataObject)zz033wList.get(i)).getValue("s_report_name") != null ) out.print((String)((DataObject)zz033wList.get(i)).getValue("s_report_name")); else out.print("&nbsp;");%>            				
            				</td>
            				<!--td width="71">    
            				<%//95.08.16拿掉裁撤日期 by 2295            				
            				//if(((DataObject)zz033wList.get(i)).getValue("cancel_date") != null && !(((DataObject)zz033wList.get(i)).getValue("cancel_date")).equals("")){
            				//   out.print(Utility.getCHTdate((((DataObject)zz033wList.get(i)).getValue("cancel_date")).toString().substring(0, 10), 0));        				
            				//}
            				%> 
            				&nbsp;              				
            				</td-->
            				
            				<td width="90" align='right'>
            				<%if(!a01_amt_990000.equals("")) out.print(Utility.setCommaFormat(a01_amt_990000)); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="90" align='right'>
            				<%if(!a04_amt_840740.equals("")) out.print(Utility.setCommaFormat(a04_amt_840740)); else out.print("&nbsp;");%>            				            				            				
            				</td>
            				<td width="90" align='right'>            				
            				<%if(!a01_amt_total.equals("")) out.print(Utility.setCommaFormat(a01_amt_total)); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="90" align='right'>            				
            				<%if(!a04_amt_840750.equals("")) out.print(Utility.setCommaFormat(a04_amt_840750)); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="52" align='right'>            				
            				<%if(!a05_amt.equals("")) out.print(Utility.setCommaFormat(a05_amt)); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="86">
            				<%if( ((DataObject)zz033wList.get(i)).getValue("muser_name") != null ) out.print((String)((DataObject)zz033wList.get(i)).getValue("muser_name")); else out.print("&nbsp;");%>            				
            				</td>
            				<td width="79">
            				<%if( ((DataObject)zz033wList.get(i)).getValue("m_telno") != null ) out.print((String)((DataObject)zz033wList.get(i)).getValue("m_telno")); else out.print("&nbsp;");%>            				
            				</td>
            				<td width="69">
            				<%if( ((DataObject)zz033wList.get(i)).getValue("telno") != null ) out.print((String)((DataObject)zz033wList.get(i)).getValue("telno")); else out.print("&nbsp;");%>            				
            				</td>            				
            			  </tr> 					      
					      <%
                  			   i++;
	                  		   }//end of while
	                  		}//end of if
    			          %>  
                  		  
					      </table>      
                      </td>    
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
                          <li>◆查詢(農漁會信用部)權限說明：</li>
                          <li>&nbsp;&nbsp;▲「農金局及全國農業金庫」：可查詢全部機構</li>
                          <li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>
                          <li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>
                          <li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>
                          <li>◆「A01_逾放金額」、「A04_逾放金額」、「A01_放款總額」、「A04_放款總額」及「A05_BIS」欄位，如果顯示內容是空白，表示未申報</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
            </table></td>
        </tr>        
</table>
</form>
</body>
</html>
