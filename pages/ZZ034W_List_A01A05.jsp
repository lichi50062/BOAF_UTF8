<%
// 95.06.06-12 create by 2295
// 96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) by 2295
// 98.09.22 fix 欲比對的申報檔案/檢核結果代碼化 by 2295
//102.08.22 fix 103.01以後漁會沒有310800統一漁貸公積,310800為未實現長期投資-跌價損失 by 2295
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
	
	System.out.println("ZZ034W_List_A01A05.act="+act);	
	System.out.println("ZZ034W_List_A01A05.bank_type_list="+szbank_type_list);	
	System.out.println("ZZ034W_List_A01A05.S_YEAR="+S_YEAR);
	System.out.println("ZZ034W_List_A01A05.S_MONTH="+S_MONTH);
	System.out.println("ZZ034W_List_A01A05.szupd_code="+szupd_code);
	System.out.println("ZZ034W_List_A01A05.szCheckOption="+szCheckOption);
	
	String sqlCmd = "";
	DataObject bean = null;
	String[] title_name= (String[])request.getAttribute("titleLength");			                   	            		        
    String[] title_width = {"127","137","137","140","145","135","155","193","159"};
    int table_width=0;
    int i=0;	
	List zz034wList = (List)request.getAttribute("zz034wList");				
	List file_list = (List)request.getAttribute("file_list");				
	List upd_codeList = (List)request.getAttribute("upd_codeList");	
	if(zz034wList == null){
	   System.out.println("zz034wList == null");
	}else{	   
	   System.out.println("zz034wList.size()="+zz034wList.size());	  	     
	   
       for(i=0;i<title_width.length;i++){
           if(title_name[i].equals("1")){//累加需要顯示的title寬度
              table_width += Integer.parseInt(title_width[i]);
           }
       }       
	}
	
	table_width = (table_width < 713) ? 713:table_width;
    request.setAttribute("table_width",String.valueOf(table_width));
    
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
    for(i=0;i<upd_codeList.size();i++){
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
<table width="<%=table_width%>" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="<%=table_width%>" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="<%=table_width%>" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="206"><img src="images/banner_bg1.gif" width="206" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「申報資料跨表檢核」 
                        </center>
                        </b></font> </td>
                      <td width="*"><img src="images/banner_bg1.gif" width="<%if(zz034wList != null && zz034wList.size() != 0) out.print(String.valueOf(table_width-206-300)); else out.print("206"); %>  " height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="<%=table_width%>" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right">
                      <%if(zz034wList != null && zz034wList.size() != 0){%>                      
                      <jsp:include page="getLoginUser.jsp?width=713" flush="true" />
                      <%}else{%>
                      <jsp:include page="getLoginUser.jsp?width=713" flush="true" />
                      <%}%>
                      </div> 
                    </tr>                    
                    <tr> 
                       <td ><table width=<%=table_width%> border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                          <tr class="sbody">						  
						  <td width='110' align='left' bgcolor='#D8EFEE'>農(漁)會別</td>
                          <td width='*' bgcolor='e7e7e7'>	
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
						  <td width='110' align='left' bgcolor='#D8EFEE'>查詢年月</td>
                          <td width='*' bgcolor='e7e7e7'>	
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
                          <td width='110' align='left' bgcolor='#D8EFEE'>欲比對的申報檔案</td>
                          <td width='*' bgcolor='e7e7e7'>                          
                            <select name='CheckOption' onchange="javascript:changeOption(document.forms[0]);">                                                                                                                
                            <%for(i=0;i<file_list.size();i++){
                                bean = (DataObject)file_list.get(i);
                            %>
                               <option value="<%=(String)bean.getValue("cmuse_id")%>" <%if(szCheckOption.equals((String)bean.getValue("cmuse_id"))) out.print("selected");%>><%=(String)bean.getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>                                 
                          </td>         
                          </tr>         
                          
                          <tr class="sbody">                          
                          <td width='110' align='left' bgcolor='#D8EFEE'>檢核結果</td>
                          <td width='*' bgcolor='e7e7e7'>                          
                            <select name='UPD_CODE' >  
                            <%for(i=0;i<upd_codeList.size();i++){
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
                      
                      
                      <tr> 
                      <%
                       String bgcolor="#D3EBE0";  
                       if(zz034wList == null || zz034wList.size() == 0){
                      %>                   		
                      <td width="713"><table width=713 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                      <div id='div_title'>
                         <tr class="sbody" bgcolor="#9AD3D0">    
                            <td width="22" rowspan="2">序號</td>                                                              
                            <td width="200" rowspan="2">總機構代號名稱</td>            				       				            				
        				    <td width="127" colspan="2" align="center">事業資金</td>            				            				            				
        				    <td width="137" colspan="2" align="center">事業公積</td>
        				    <td width="137" colspan="2" align="center">法定公積</td>
        				    <td width="140" colspan="2" align="center">特別公積</td>            				            				            				
        				    <td width="145" colspan="2" align="center">捐贈公積</td>            				            				            									        					        
					      </tr> 
					      <tr class="sbody" bgcolor="#9AD3D0">    
        				    <td width="44">A01<br>(310100)</td>            				            				            				
        				    <td width="70">A05<br>(910101)</td>            				            				            				
        				    <td width="59">A01<br>(310200)</td>            				            				            				
        				    <td width="60">A05<br>(910102)</td>            				            				            				
        				    <td width="64">A01<br>(310300)</td>            				            				            				
        				    <td width="59">A05<br>(910103)</td>            				            				            				
        				    <td width="61">A01<br>(310400)</td>            				            				            				
        				    <td width="60">A05<br>(910104)</td>            				            				            				
        				    <td width="72">A01<br>(310500)</td>            				            				            				
        				    <td width="64">A05<br>(910105)</td>            				            				            				        				    
					      </tr> 

					      <tr class="sbody" bgcolor="#D3EBE0">
                   		   <td colspan=20 align=center>無資料可供查詢</td>
                   		  <tr>
                   	 	</div>	  
                   	 <%}else{%>                   			                         
                      <td><table width=<%=table_width%> border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                         <tr class="sbody" bgcolor="#9AD3D0">    
                            <td width="20" rowspan="2">序號</td>                                                              
                            <td width="220" rowspan="2">總機構代號名稱</td>    
                            <% //顯示檢核有誤的title name                               
                               String[] title = {"事業資金","事業公積","法定公積","特別公積","捐贈公積","資產公積","統一農(漁)貸公積","累積盈虧(A05加計「上期損益」","聯營出資(投資)"};
                               //96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) 
							   if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 9701){
    							  title[8] = "長期投資";
    						   }
                               for(i=0;i<title_name.length;i++){
                                   if(title_name[i].equals("1")){System.out.println("show"+title[i]);
                                   	  if(szbank_type_list.equals("7") && i == 6 && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){
                                   		//103.01.漁會沒有310800統一漁貸公積
                                   		System.out.println("no show");
                                      }else{                                   	
                                   	%>
                                    <td width="<%=title_width%>" colspan="2" align="center"><%=title[i]%></td>             
                           <%         }
                                   }
                               }
                           %>			
					      </tr> 
					      <tr class="sbody" bgcolor="#9AD3D0">    
					      <% //顯示檢核有誤的amt name					           
        		               String[][] amt_width = {{"44","70"},{"59","60"},{"64","59"},{"61","60"},{"72","64"},{"82","76"},{"78","75"},{"87","100"},{"67","90"}};
        		               String[] acc_code_A01 = {"310100","310200","310300","310400","310500","310600","310800","320100+320200","130200"};
        		               String[] acc_code_A05 = {"910101","910102","910103","910104","910105","910106","910107","910108","910401+910402+910403+910404"};
                               for(i=0;i<title_name.length;i++){                               	   
                                   if(title_name[i].equals("1")){
                                   	 if(szbank_type_list.equals("7") && acc_code_A01[i].equals("310800") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){
                                   		//103.01.漁會沒有310800統一漁貸公積
                                     }else{
                                   	%>
                                    <td width="<%=amt_width[i][0]%>">A01<br>(<%=acc_code_A01[i]%>)</td>            				            				            				
        				            <td width="<%=amt_width[i][1]%>">A05<br>(<%if(i==8) out.print("910401+910402<br>+910403+910404"); else out.print(acc_code_A05[i]);%>)</td>            	             
                           <%        }
                                   }
                               }
                           %>			        				    	            				            				
					      </tr> 
					      <%
					        int rowidx=0; 
					        //int idx=0;
					        String s_report_name="";
					        String[][] Amt=null;
					        String tmpBank_no="";						    
					        String tmpBank_name="";						    
					        while(rowidx < zz034wList.size()){ 
					             bgcolor = (rowidx % 2 == 0)?"#e7e7e7":"#D3EBE0";	   
						         s_report_name = (String)((List)zz034wList.get(rowidx)).get(0);//s_report_name			           			 
						         Amt = (String[][])((List)zz034wList.get(rowidx)).get(1);//Amt
			           		%>			           		  
					           <tr class="sbody" bgcolor="<%=bgcolor%>">
                               <td width="20"><%=rowidx+1%></td>                       				            				            				            				
            				   <td width="220">
            				   <%if(!s_report_name.equals("")) out.print(s_report_name); else out.print("&nbsp;");%>            				
            				   </td>             	
            				   <%for(i=0;i<title_name.length;i++){
                                   if(title_name[i].equals("1")){
                                   	  if(szbank_type_list.equals("7") && i == 6 && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){
                                       //103.01.漁會沒有310800統一漁貸公積
                                      }else{
                                   	%>            				   
                               <td width="<%=amt_width[i][0]%>" align="right"><%if(!szupd_code.equals("3") && Amt[i][2] != null && Amt[i][2].equals("1")) out.print("<u>");%><%if(Amt[i][0] != null && !Amt[i][0].equals("") && !Amt[i][0].equals("null")) out.print(Utility.setCommaFormat(Amt[i][0])); else out.print("&nbsp;");%><%if(!szupd_code.equals("3") && Amt[i][2] != null && Amt[i][2].equals("1")) out.print("</u>");%></td>            				            				            				
        				       <td width="<%=amt_width[i][1]%>" align="right"><%if(!szupd_code.equals("3") && Amt[i][2] != null && Amt[i][2].equals("1")) out.print("<u>");%><%if(Amt[i][1] != null && !Amt[i][1].equals("") && !Amt[i][1].equals("null")) out.print(Utility.setCommaFormat(Amt[i][1])); else out.print("&nbsp;");%><%if(!szupd_code.equals("3") && Amt[i][2] != null && Amt[i][2].equals("1")) out.print("</u>");%></td>            	             
                            <%        }
                                    }
                                 }
                               %>
            			       </tr> 					      
					         <%
                  			   rowidx++;
	                  	    }//end of while	                  		
	                   }//end of zz034wList != null%>	  
					      </table>      
                      </td>   
                       
                      </tr>
                                  
      </table></td>
  </tr> 
  <tr>
          <td bgcolor="#FFFFFF"><table width="<%=table_width%>" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="<%=table_width%>" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="<%=table_width%>" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>                      
                      <td width="<%=table_width%>">
                          <ul>                      
                          <div id='div_memo'>
                          <li>◆查詢(農漁會信用部)權限說明：</li>
                          <li>&nbsp;&nbsp;▲「農金局及全國農業金庫」：可查詢全部機構</li>
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
