<%
//105.03.31 create by 2968
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
    String[] title_width = {"127","137","137","140","145","180","155","193","159",
    		                "127","137","137","140","145","135","155","193","159"};
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
                            <%if(!szCheckOption.equals("2")){
                            //103.01以後.漁會套用新科目代號.取消農漁會共同顯示%> 
                            <!--<option value="ALL" <%if(szbank_type_list.equals("ALL")) out.print("selected");%>>農漁會</option>  -->
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
                                if(((String)bean.getValue("identify_no")).equals("5")){
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
        				    <td width="127" colspan="2" align="center">A01資產負債表之資產合計=負債合計+淨值合計</td>            				            				            				
        				    <td width="137" colspan="2" align="center">A01資產負債表中之存款總額=法定比率分析統計表中(19)正會員存款總額+(20)贊助會員存款總額=(21)非會員存款總額(含公庫存款)</td>
        				    <td width="137" colspan="2" align="center">A01資產負債表中之放款淨額+備抵呆帳-放款+備抵呆帳-催收項項-內部融資=法定比率分析統計表中(22)正會員放款總額+(23)贊助會員放款總額+(24)非會員放款總額(不含內部融資)</td>
        				    <td width="140" colspan="2" align="center">A01資產負債表中之1/2公庫存款=法定比率分析統計表中(25)1/2公庫存款</td>            				            				            				
        				    <td width="145" colspan="2" align="center">A01資產負債表中之固定資產淨額=法定比率分析統計表中(27)信用部固定資產淨額</td>            				            				            									        					        
					      </tr> 
					      <tr class="sbody" bgcolor="#9AD3D0">    
        				    <td width="44">A01<br>(190000)</td>            				            				            				
        				    <td width="70">A01<br>(400000)</td>            				            				            				
        				    <td width="59">A01<br>(220000)</td>            				            				            				
        				    <td width="60">A99.992130+<br>A02.990420+<br>A02.990620</td>            				            				            				
        				    <td width="64">A01.120000+<br>A01.120800+<br>A01.150300-<br>A01.120700</td>            				            				            				
        				    <td width="59">A99.992140+<br>A02.990410+<br>A02.990610</td>            				            				            				
        				    <td width="61">A01<br>(220900)<br>/2</td>            				            				            				
        				    <td width="60">A02<br>(990630)<br>/2</td>            				            				            				
        				    <td width="72">A01<br>(140000)</td>            				            				            				
        				    <td width="64">A02<br>(990810)</td>            				            				            				        				    
					      </tr> 

					      <tr class="sbody" bgcolor="#D3EBE0">
                   		   <td colspan=20 align=center>無資料可供查詢</td>
                   		  <tr>
                   	 	</div>	  
                   	 <%}else{%>                   			                         
                      <td><table width=<%=table_width%> border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                         <tr class="sbody" bgcolor="#9AD3D0">    
                            <td width="30" rowspan="2">序號</td>                                                              
                            <td width="200" rowspan="2">總機構代號名稱</td>    
                            <% //顯示檢核有誤的title name                               
                               String[] title = {"A01資產負債表之資產合計=負債合計+淨值合計",
                            		"A01資產負債表中之存款總額=法定比率分析統計表中(19)正會員存款總額+(20)贊助會員存款總額=(21)非會員存款總額(含公庫存款)",
                            		"A01資產負債表中之放款淨額+備抵呆帳-放款+備抵呆帳-催收項項-內部融資=法定比率分析統計表中(22)正會員放款總額+(23)贊助會員放款總額+(24)非會員放款總額(不含內部融資)",
                            		"A01資產負債表中之1/2公庫存款=法定比率分析統計表中(25)1/2公庫存款",
                            		"A01資產負債表中之固定資產淨額=法定比率分析統計表中(27)信用部固定資產淨額",
                            		"A01資產負債表中之逾期放款總額=A04應予觀察放款總計表中之逾期放款總額=法定比率分析統計表中(38)正會員放款中之逾期放款+(39)贊助會員放款中之逾期放款+(40)非會員放款中之逾期放款+(41)內部融資中之逾期放款",
                            		"A04應予觀察放款統計表中之應予觀察放款總額=法定比率分析統計表中(42)正會員放款中之應予觀察放款+(43)贊助會員放款中之應予觀察放款+(44)非會員放款中之應予觀察放款+(45)內部融資中之應予觀察放款",
                            		"法定比率分析統計表中(50)非會員無擔保消費性貸款=(51)非會員無擔保消費性政策貸款+(52)非會員無擔保消費性非政策貸款",
                            		"法定比率分析統計表中(53)無擔保消費性貸款 > (50)非會員無擔保消費性貸款",
                            		"法定比率分析統計表中(53)無擔保消費性貸款 > (54)無擔保消費性貸款中之逾期放款",
                            		"法定比率分析統計表中(53)無擔保消費性貸款 > (55)無擔保消費性貸款中之應予觀察放款",
                            		"法定比率分析統計表中(24)非會員放款總額(不含內部融資) > (56)非會員放款中之政策性農業專案貸款",
                            		"A01資產負債表中之備抵呆帳-放款+ 備抵呆帳-催收款項=A10應予評估資產彙總表中之放款帳列備抵呆帳合計",
                            		"A10應予評估資產彙總表中之放款帳列備抵呆帳一類+放款帳列備抵呆帳二類+放款帳列備抵呆帳三類+放款帳列備抵呆帳四類=放款帳列備抵呆帳合計",
                            		"A10應予評估資產彙總表中之應予評估放款二類+應予評估放款三類+應予評估放款四類 >=A04應予觀察放款總計表中之應予觀察放款總額+A01資產負債表中之逾期放款總額",
                            		"A10應予評估資產彙總表中之應予評估放款一類+應予評估放款二類+應予評估放款三類+應予評估放款四類=應予評估放款合計",
                            		"A10應予評估資產彙總表中之應予評估投資一類+應予評估投資二類+應予評估投資三類+應予評估投資四類=應予評估投資合計",
                            		"A10應予評估資產彙總表中之應予評估放款合計=A01資產負債表中之放款總額-A02(990611)對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額 "};
                               for(i=0;i<title_name.length;i++){
                                   if(title_name[i].equals("1")){
                                   	  if(i==5){%>
                                   	  <td width="<%=title_width%>" colspan="3" align="center"><%=title[i]%></td>
                                   <% }else{     %>                              	
                                      <td width="<%=title_width%>" colspan="2" align="center"><%=title[i]%></td>             
                           <%         }
                                   }
                               }
                           %>			
					      </tr> 
					      <tr class="sbody" bgcolor="#9AD3D0">    
					      <% //顯示檢核有誤的amt name	
        		               String[][] amt_width = {{"44","70"},{"59","60"},{"64","59"},{"61","60"},{"72","64"},{"82","76"},{"78","75"},{"87","100"},{"67","90"},
        		               						   {"44","70"},{"59","60"},{"64","59"},{"61","60"},{"72","64"},{"82","76"},{"78","75"},{"87","100"},{"67","90"}};
					      		String[] acc_code_L = {"A01<br>(190000)","A01(220000)","A01.120000+<br>A01.120800+<br>A01.150300-<br>A01.120700","A01<br>(220900)<br>/2","A01<br>(140000)",
									      "A01<br>(990000)","A04<br>(840760)","A02<br>(990510)","A99<br>(992150)","A99<br>(992150)",
									      "A99<br>(992150)","A02<br>(990610)","A01.120800+<br>A01.150300","A10.990021+<br>A10.990022+<br>A10.990023+<br>A10.990024","A10.990002+<br>A10.990003+<br>A10.990004",
									      "A10.990001+<br>A10.990002+<br>A10.990003+<br>A10.990004","A10.990006+<br>A10.990007+<br>A10.990008+<br>A10.990009","A10<br>(990005)"};
      							String[] acc_code_R = {"A01<br>(400000)","A99.992130+<br>A02.990420+<br>A02.990620","A99.992140+<br>A02.990410+<br>A02.990610","A02<br>(990630)<br>/2","A02<br>(990810)",
      									"A99.992510+<br>A99.992520+<br>A99.992530+<br>A99.992540","A99.992610+<br>A99.992620+<br>A99.992630+<br>A99.992640","A02.990511+<br>A02.990512","A02<br>(990510)","A99<br>(992550)",
      									"A99<br>(992650)","A02<br>(990612)","A10<br>(990025)","A10<br>(990025)","A04.840760+<br>A01.990000",
      									"A10<br>(990005)","A10<br>(990010)","A01.120000+<br>A01.120800+<br>A01.150300-<br>A02.990611"};
                               for(i=0;i<title_name.length;i++){  
                            	   
                                   if(title_name[i].equals("1")){
                                   		%>
                                    <td width="<%=amt_width[i][0]%>"><%=acc_code_L[i]%></td>   
                                    <%if(i==5){ %>
        				            <td width="<%=amt_width[i][0]%>">A04<br>(840740)</td>
        				            <%} %>         				            				            				
        				            <td width="<%=amt_width[i][1]%>"><%=acc_code_R[i]%></td> 
        				                       	             
                           <%        }
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
					           <tr class="sbody" bgcolor="<%=bgcolor%>" >
                               <td width="20"><%=rowidx+1%></td>                       				            				            				            				
            				   <td width="220" align="left">
            				   <%if(!s_report_name.equals("")) out.print(s_report_name); else out.print("&nbsp;");%>            				
            				   </td>             	
            				   <%for(i=0;i<title_name.length;i++){
                                   if(title_name[i].equals("1")){
                                   	%>            				   
	                               <td width="<%=amt_width[i][0]%>" align="right" ><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("<u>");%><%if(Amt[i][0] != null && !Amt[i][0].equals("") && !Amt[i][0].equals("null")) out.print(Utility.setCommaFormat(Amt[i][0])); else out.print("&nbsp;");%><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("</u>");%></td>            				            				            				
	        				       <%if(i==5){ %>
	        				       <td width="<%=amt_width[i][1]%>" align="right" ><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("<u>");%><%if(Amt[title_name.length][1] != null && !Amt[title_name.length][1].equals("") && !Amt[title_name.length][1].equals("null")) out.print(Utility.setCommaFormat(Amt[title_name.length][1])); else out.print("&nbsp;");%><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("</u>");%></td>
                            	   <%} %>
                            	   <td width="<%=amt_width[i][1]%>" align="right" ><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("<u>");%><%if(Amt[i][1] != null && !Amt[i][1].equals("") && !Amt[i][1].equals("null")) out.print(Utility.setCommaFormat(Amt[i][1])); else out.print("&nbsp;");%><%if(!szupd_code.equals("3") && Amt[i][2].equals("1")) out.print("</u>");%></td>            
                            <%        
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
