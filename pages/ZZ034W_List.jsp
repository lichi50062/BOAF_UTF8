<%
//95.06.06 create by 2295
//95.08.16 fix 各逾放期數合計之科目【960500】欄位的金額，
//             應等於A04.【840731】+ A04.【840732】+A04.【840733】+A04.【840734】+ A04.【840735】的合計申報金額
//             不跨表檢核 by 2295
//98.09.22 fix 欲比對的申報檔案/檢核結果代碼化 by 2295
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
    String szbank_type_list = ( request.getParameter("BANK_TYPE_List")==null ) ? "" : (String)request.getParameter("BANK_TYPE_List");				
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? Utility.getYear() : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? Utility.getMonth() : (String)request.getParameter("S_MONTH");				
	String szUnit = ( request.getParameter("unit")==null ) ? "" : (String)request.getParameter("unit");				
	String szupd_code = ( request.getParameter("upd_code")==null ) ? "" : (String)request.getParameter("upd_code");				
	String szCheckOption = ( request.getParameter("checkoption")==null ) ? "1" : (String)request.getParameter("checkoption");					
	
	System.out.println("ZZ034W_List.act="+act);	
	System.out.println("ZZ034W_List.bank_type_list="+szbank_type_list);	
	System.out.println("ZZ034W_List.S_YEAR="+S_YEAR);
	System.out.println("ZZ034W_List.S_MONTH="+S_MONTH);
	System.out.println("ZZ034W_List.szupd_code="+szupd_code);
	System.out.println("ZZ034W_List.szCheckOption="+szupd_code);
	
	String sqlCmd = "";
	DataObject bean = null;	
	List zz034wList = (List)request.getAttribute("zz034wList");			
	List zz033wList = (List)request.getAttribute("zz033wList");		
	List file_list = (List)request.getAttribute("file_list");		
	List upd_codeList = (List)request.getAttribute("upd_codeList");		
	List A01List = null;
	List A05List = null;
	if(zz034wList == null){
	   System.out.println("zz034wList == null");
	}else{
	   System.out.println("zz034wList.size()="+zz034wList.size());
	   zz034wList = null;
	   A01List = (List)zz034wList.get(0);
	   A05List = (List)zz034wList.get(1);
	}
	
	//取得ZZ034W的權限
	Properties permission = ( session.getAttribute("ZZ034W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ034W"); 
	if(permission == null){
       System.out.println("ZZ034W_List.permission == null");
    }else{
       System.out.println("ZZ034W_List.permission.size ="+permission.size());               
    }	
    
    // XML Ducument for 檢核結果 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"ReportListXML\">");
    out.println("<datalist>");       
    
    for(int i=0;i<upd_codeList.size();i++){
        bean = (DataObject)upd_codeList.get(i);
        out.println("<data>");
        out.println("<fileType>"+(String)bean.getValue("identify_no")+"</fileType>");        
        out.println("<optionValue>"+(String)bean.getValue("input_order")+"</optionValue>");   
        out.println("<optionName>"+(String)bean.getValue("cmuse_name")+"</optionName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 檢核結果 end 	
   
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ034W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「申報資料跨表檢核」</title>
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
	     form.action="/pages/ZZ034W.jsp?act="+cnd+"&test=nothing";	    	    
	     
	     if(confirm("本項查詢會報行10-15秒，是否確定執行？")){
	         form.submit();	    		    
	     }    
}	
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">   
<%if(zz034wList != null && zz034wList.size() != 0){%>
<input type="hidden" name="row" value="<%=zz034wList.size()+1%>">   
<%}%>
<table width="713" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="713" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="713" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="206"><img src="images/banner_bg1.gif" width="206" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「申報資料跨表檢核」 
                        </center>
                        </b></font> </td>
                      <td width="206"><img src="images/banner_bg1.gif" width="206" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="713" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=713" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                       <td ><table width=713 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                          <tr class="sbody">						  
						  <td width='18%' align='left' bgcolor='#D8EFEE'>農(漁)會別</td>
                          <td width='82%' bgcolor='e7e7e7'>	
                            <select name='BANK_TYPE_List' >                                                        
                            <option value="6" <%if(szbank_type_list.equals("6")) out.print("selected");%>>農會</option>
                            <option value="7" <%if(szbank_type_list.equals("7")) out.print("selected");%>>漁會</option>
                            <%if(!szCheckOption.equals("2")){%> 
                            <option value="ALL" <%if(szbank_type_list.equals("ALL")) out.print("selected");%>>農漁會</option>
                            <%}%>
                            </select>     
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:doSubmit(this.document.forms[0],'List');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                                                                              
                          </td>             
                          </tr>       
                          
                          <tr class="sbody">                          
						  <td width='18%' align='left' bgcolor='#D8EFEE'>查詢年月</td>
                          <td width='82%' bgcolor='e7e7e7'>	
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
                          <td width='110' align='left' bgcolor='#D8EFEE'>金額單位</td>
                          <td width='*' bgcolor='e7e7e7'>                                                    
                        	<select id="hide1" name="Unit" >
        			            <option value="1" <%if(szUnit.equals("1")) out.print("selected");%>>元</option>        			            
                              	<option value="1000" <%if(szUnit.equals("1000")) out.print("selected");%>>千元</option>
                              	<option value="10000" <%if(szUnit.equals("10000")) out.print("selected");%>>萬元</option>
                              	<option value="1000000" <%if(szUnit.equals("1000000")) out.print("selected");%>>百萬元</option>
                              	<option value="10000000" <%if(szUnit.equals("10000000")) out.print("selected");%>>仟萬元</option>
                              	<option value="100000000" <%if(szUnit.equals("100000000")) out.print("selected");%>>億元</option>
                           	</select>
                          </td>
                          </tr>

                          <tr class="sbody">                          
                          <td width='18%' align='left' bgcolor='#D8EFEE'>欲比對的申報檔案</td>
                          <td width='82%' bgcolor='e7e7e7'>                          
                            <select name='CheckOption' onchange="javascript:changeOption(document.forms[0]);">  
                            <%
                            for(int i=0;i<file_list.size();i++){
                                bean = (DataObject)file_list.get(i);
                            %>
                               <option value="<%=(String)bean.getValue("cmuse_id")%>" <%if(szCheckOption.equals((String)bean.getValue("cmuse_id"))) out.print("selected");%>><%=(String)bean.getValue("cmuse_name")%></option>                            
                            <%    
                            }
                            %>                                                                                                                                         
                          </td>         
                          </tr>         
                          
                          <tr class="sbody">                          
                          <td width='18%' align='left' bgcolor='#D8EFEE'>檢核結果</td>
                          <td width='82%' bgcolor='e7e7e7'>                          
                            <select name='UPD_CODE' >                                                                                                                
                            <%for(int i=0;i<upd_codeList.size();i++){
                                bean = (DataObject)upd_codeList.get(i);
                                if(((String)bean.getValue("identify_no")).equals("1")){//A01/A05相對科目金額之比對
                            %>
                               <option value="<%=(String)bean.getValue("input_order")%>" <%if(szupd_code.equals((String)bean.getValue("input_order"))) out.print("selected");%>><%=(String)bean.getValue("cmuse_name")%></option>                            
                            <%  }
                              }
                            %>       
                            </select>                                 
                          </td>         
                          </tr>                                		      					      
                          </table>      
                      </td>    
                      </tr>
      </table></td>
  </tr> 
  <tr>
          <td bgcolor="#FFFFFF"><table width="720" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="710" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="720" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>                      
                      <td width="720">
                          <ul>                      
                          <div id='div_memo'>
                          <li>◆查詢(農漁會信用部)權限說明：</li>
                          <li>&nbsp;&nbsp;▲「農金局/中央存保及全國農業金庫」：可查詢全部機構</li>
                          <li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>
                          <li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>
                          <li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>
                          <li>◆各欄位金額，如果顯示內容是空白，表示未申報</li>                     
                          </div>
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
