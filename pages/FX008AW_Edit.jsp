<%
// 94.11.24 additional function (audit applydelay blabla)
// 94.10.23 standard edit and insert ok by 4183
// 94.10.20 first design by 4183
// 98.01.07 fix 不試算核准申請延長期限(A)+(B) by 2295 
//100.01.27 fix title寬度 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>
<script language="javascript" src="js/FX008AW.js"> </script>
<script language="javascript" src="js/Common.js"></script>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link href="css/b51.css" rel="stylesheet" type="text/css">

<!-- ----------------- Javascript ---------------------------------------------------------- -->
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
-->
</script>
<!-- ----------------- end of Javascript ---------------------------------------------------------- -->

<%	//--------------------取得資料------------------------
	String str ="登入成功";
	String title = "審核";
	String bank_no =  ( request.getParameter("bank_no")==null ) ? " " : (String)request.getParameter("bank_no");
	String quarter = " ";
	List WLX08_S_Edit = (List)session.getAttribute("WLX08_AS_Edit");     //所選取那筆年季資料
	List WLX08_R_Data = (List)session.getAttribute("WLX08_RptBOAF_Data");//把資料轉成xml然後讓javascript使用
	
	//使用dataobject type 來做資料傳遞
	DataObject Bean = (DataObject)WLX08_S_Edit.get(0);//把選取的資料取出來
	
	
	//判斷list合格否有取得資料
	if(WLX08_S_Edit == null){
	   System.out.println("WLX08_S_Edit == null");
	}else{
	   System.out.println("WLX08_S_Edit.size()="+WLX08_S_Edit.size());
	}
	
	//取得FX008W的權限
	Properties permission = ( session.getAttribute("FX008AW")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ005W");
	if(permission == null){
		System.out.println("FX008AW_List.permission == null");
	}
	else{
		System.out.println("FX008AW_List.permission.size ="+permission.size());
	}
	
	// XML Ducument for 記載最近一期呈報出去的文號以及日期 for 抄錄功能使用 begin
	Date tempDate;
    int temp_year   = 0;                                                     
	int temp_month  = 0;                                                   
	int temp_day    = 0;
	
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"Rpt_BOAF\">");
    out.println("<datalist>");
    out.println("<data>");
    
    if((WLX08_R_Data!=null) && WLX08_R_Data.size() > 0){
    	
        DataObject temp_bean =(DataObject)WLX08_R_Data.get(0);
       	out.println("<flag>Y</flag>");
        out.println("<docno>"+temp_bean.getValue("applyok_docno")+"</docno>");
        out.println("<boafdocno>"+temp_bean.getValue("report_boaf_docno")+"</boafdocno>");
        
		//把年月日濾出來
		tempDate = (Date)temp_bean.getValue("applyok_date");
		temp_year   = tempDate.getYear()-11;//轉成民國制
		temp_month  = tempDate.getMonth()+1;  
		temp_day    = tempDate.getDate();    
		
		out.println("<apoky>"+temp_year+"</apoky>");
		out.println("<apokm>"+temp_month+"</apokm>");
		out.println("<apokd>"+temp_day+"</apokd>"); 
		
		tempDate = (Date)temp_bean.getValue("report_boaf_date");
		temp_year   = tempDate.getYear()-11;//轉成民國制
		temp_month  = tempDate.getMonth()+1;  
		temp_day    = tempDate.getDate();    
		
		out.println("<boafy>"+temp_year+"</boafy>");
		out.println("<boafm>"+temp_month+"</boafm>");
		out.println("<boafd>"+temp_day+"</boafd>"); 
    }   
    out.println("</data>");
    out.println("</datalist>\n</xml>");
    // XML Ducument for 最近一期呈報出去的文號以及日期 end
	
%>

<title> 各農漁會承受擔保品延長處分申報情形<%=title%> </title>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method="post" >
<table border="0" cellspacing="1" width="600">
	<tr>
    	<td width="600"></td>
	</tr>
  	
  	<tr>
  		<td width="600">
        <center>
        	<table border="0" cellspacing="1">
          		<tr>
          			<td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                   	<td width="400"><p align="center"><font color='#000000' size=4><b>各農漁會承受擔保品延長處分申報情形<%=title%></b></font></p></td>
                   	<td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
				</tr>
			</table>
       </center>
    	</td>
	</tr>
  
	<tr> 
       <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
    </tr>
  
<!-- ---------------------------------------------------------------------------------------------------------------------------- -->
	<tr>
    	<td width="600">
      	<center>
        <table class="sbody"  border="1" cellspacing="1" bordercolor="#3A9D99"  bordercolor="#3A9D99" width="599" height="346">
        	<tr>
        		<td width='191' bgcolor="#D8EFEE" colspan="3">金融機構代號</td>
				<% 
					if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
						bank_no = Bean.getValue("bank_no").toString();
					}
				%>
				<td width='385' bgcolor='e7e7e7' height="27"><%=bank_no%></td>
        	</tr>
        	
        	<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3">申報年季</td>
				<td width='385' bgcolor='e7e7e7' height="10">
            		<input type='text' name='m_year'
            				value="<%
               					 	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
               			 				out.print(Bean.getValue("m_year"));
               			 			}
               			 			%>" size='3' maxlength='3' onblur='CheckYear(this)' disabled >
               			 	
					<font color='#000000'>年第                              
					<select id="hide2" name='m_quarter' disabled>
					<%
        				if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
        			 		quarter = Bean.getValue("m_quarter").toString();
        			 		out.print(quarter);
        			 	}
               			
               			String sel1=" ";
               			String sel2=" ";
               			String sel3=" ";
               			String sel4=" ";

               			if(quarter.equals("1"))
               				sel1="selected";
            			else if(quarter.equals("2"))
               				sel2="selected";
                		else if(quarter.equals("3"))
               				sel3="selected";
               			else if(quarter.equals("4"))
               				sel4="selected";

        	 		%>
					<option value="1" <%=sel1%> >01</option>																																																			
					<option value="2" <%=sel2%> >02</option>																																																			
					<option value="3" <%=sel3%> >03</option>																																																			
					<option value="4" <%=sel4%> >04</option>																																																			
					</select>季&nbsp;&nbsp;</font>
				</td>
        	</tr>
		
			<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3">承受擔保品編號</td>
				<td width='385' bgcolor='e7e7e7' >
					<input type='text' name='dure_no'
							value="<%
	             					if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
	               						out.print(Bean.getValue("dureassure_no"));
	               			 		}
	               		    		%>" size='6' maxlength='6' onBlur='this.value=changeStr(this);'
	               		    			style='text-align: right;' disabled >
	            	<font color='red' size=4 > *</font>                              
				</td>
	       	</tr>
       
			<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3" >機構名稱(借款人)</td>
				<td width='385' bgcolor='e7e7e7' >
					<input type='text' name='debtname'
							value="<%
	                			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
	               			 		out.print(Bean.getValue("debtname"));         			 		
	               			 	}
	               			 	%>"size='20' maxlength='20' disabled>
	            	<font color='red' size=4>*</font><input type="button" name="Check_Borrowe_Summary" value="機構名稱(借款人)申辦之歷史明細" onClick="javascript:doSubmit(this.document.forms[0],'load','');">
	            </td>
			</tr>
		
			<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3">承受日期 <font color='blue' size=2>(A)</font></td>
				<td width='385' bgcolor='e7e7e7' >
					<% //----------------get duredate and format it---------------------------------
						Date duDate = new Date();
						int month = 1;
		      			int day = 1;
		      
		        		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		        			
		        			duDate = (Date)Bean.getValue("duredate");	               			 	 
		        			month = duDate.getMonth()+1;
		              		day = duDate.getDate();
		        		}
		   			%>
		            <input type='text' name='accept_year'
		            		value="<%
		            				if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && duDate != null){
		               			 		out.print(duDate.getYear()-11);
		               				}
		               				%>"size='4' maxlength='4' onblur='CheckYear(this)' disabled >
		            <font color='#000000'>年                              
		            <select  name='accept_month' disabled>
						<option value="<%=month%>"><%=month%></option>
						<option value="01" >01</option>
						<option value="02" >02</option>
						<option value="03" >03</option>
						<option value="04" >04</option>
						<option value="05" >05</option>
						<option value="06" >06</option>
						<option value="07" >07</option>
						<option value="08" >08</option>
						<option value="09" >09</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
					</select>
		
					月                              
					<select id="hide4" name='accept_day' disabled >
						<option value="<%=day%>" ><%=day%></option>
						<option value="01" >01</option>
						<option value="02" >02</option>
						<option value="03" >03</option>
						<option value="04" >04</option>
						<option value="05" >05</option>
						<option value="06" >06</option>
						<option value="07" >07</option>
						<option value="08" >08</option>
						<option value="09" >09</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
						<option value="13" >13</option>
						<option value="14" >14</option>
						<option value="15" >15</option>
						<option value="16" >16</option>
						<option value="17" >17</option>
						<option value="18" >18</option>
						<option value="19" >19</option>
						<option value="20" >20</option>
						<option value="21" >21</option>
						<option value="22" >22</option>
						<option value="23" >23</option>
						<option value="24" >24</option>
						<option value="25" >25</option>
						<option value="26" >26</option>
						<option value="27" >27</option>
						<option value="28" >28</option>
						<option value="29" >29</option>
						<option value="30" >30</option>
						<option value="31" >31</option>
					</select>日</font><font color='red' size=4>*</font>
				</td>
			</tr>
		
			<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3" height="17">承受擔保品座落</td>
				<td width='385' bgcolor='e7e7e7' height="1">
				<input type='text' name='duresite'
						value="<%
								if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
	               			 		out.print( Bean.getValue("dureassuresite") );
	               			 	}
	               			 	%>" size='50' maxlength='50' disabled >
	      		<font color='red' size=4>*</font><br>                              
	      		<font color="#FF0000" size="2">(土地:段、地號、面積、持分；建物:建號)</font>
				</td>
			</tr>
        
	        <tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3" height="17">帳列金額</td>
				<td width='385' bgcolor='e7e7e7' height="36">
					<input type='text' name='account'
							value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
										out.print( Utility.setCommaFormat( Bean.getValue("accountamt").toString())  );
									}
		          					%>"size='20' maxlength='20' onFocus='this.value=changeVal(this)'
		          					   onBlur='checkPoint_focus(this);this.value=changeStr(this);'
		          					   style='text-align: right;' disabled >
					<font color='red' size=4>*&nbsp;&nbsp;</font><font size="2" color="#FF0000">(單位:新台幣元)</font>                              
				</td>
			</tr>
			
			<tr>
				<td width='191' bgcolor='#D8EFEE' colspan="3" height="17">申請延長期間</td>
				<td width='385' bgcolor='e7e7e7' height="24">
					<input type='text' name='apply_year'
							value="<%
		                 			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		                 				out.print( Bean.getValue("applydelayyear") );
		                 			}
		                 			%>" size='5' maxlength='3' onblur='CheckYear(this)' disabled >
		            年                              
					<input type='text' name='apply_month'
							value="<%
		                 			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		                 				out.print( Bean.getValue("applydelaymonth") );
		                 			}
		                 			%>" size='2' maxlength='2' onblur='CheckYear(this)' disabled >
					</select>個月</font>                           
		            <font color='red' size=4>*</font>
				</td>
			</tr>
		
			<tr>
				<td width='190' bgcolor='#D8EFEE' colspan="3" height="17">申請延長理由</td>
				<td width='400' bgcolor='e7e7e7' height="39">
					<textarea rows="4" cols="50" name='apply_reason' disabled ><%
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
							out.print( Bean.getValue("applydelayreason") );
						}
		            	%></textarea>
		            <font color='red' size=4>*</font>                              
				</td>
			</tr>
			
			<tr>
				<td width='70'  bgcolor='#D8EFEE' rowspan ="2" > <font color="#0000FF">縣市政府</font></td>
				<td width='150' bgcolor='#D8EFEE' colspan ="2" height="17">核准申請延長期間<font color='blue' size=2>(B)</font></td>
				<td width='380' bgcolor='e7e7e7'>
					<input type='text' name='Audit_Delay_y' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("audit_applydelayyear")!= null){
										out.print(Bean.getValue("audit_applydelayyear"));
									}
								%>" 
						size='2' maxlength='4' onblur='CheckYear(this)'>&nbsp;年
						      
					<input type='text' name='Audit_Delay_m' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("audit_applydelayyear")!= null){
										out.print(Bean.getValue("audit_applydelaymonth"));
									}
								%>" 
						size='2' maxlength='4' onblur='CheckYear(this),checkMonth(this)'>
					</select>個月</font>                           
		            <font color='red' size=4>*</font>    
                    <!-- 按了之後 抄錄上面的申請延長日期 -->
                    <input type="button" name="Load_Upper_Dure" value="同申請延長日期" onClick="javascript:loadUpperDate(this.document.forms[0]);">             
				</td>
			</tr>
			
			<tr>
				<td width='150' bgcolor='#D8EFEE' colspan ="2" height="17">核准申請延長期限<font color='blue' size=2>(A)+(B)</font></td>
					<%
						Date get_Cnt_date;
						
						String Cnt_year  = " ";
						String Cnt_month = " ";
						String Cnt_date  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("audit_duredate")!= null){	
							get_Cnt_date = (Date)Bean.getValue("audit_duredate");
							Cnt_year  = String.valueOf(get_Cnt_date.getYear()-11);
							Cnt_month = String.valueOf(get_Cnt_date.getMonth()+1);
							Cnt_date  = String.valueOf(get_Cnt_date.getDate());
						}
					%>
				<td width='380' bgcolor='#e7e7e7'>
                    <input type='text' name='Cnt_year'  value="<%=Cnt_year %>" size='3' maxlength='3' onblur='CheckYear(this)'>&nbsp;年      
					<input type='text' name='Cnt_month' value="<%=Cnt_month %>" size='2' maxlength='2' onblur='CheckYear(this)'>&nbsp;月      
					<input type='text' name='Cnt_date'  value="<%=Cnt_date %>" size='2' maxlength='2' onblur='CheckYear(this)'>&nbsp;日                  
		            <font color='red' size=4>*&nbsp;</font> 
                    <!-- 按了之後 試算出延長之後之日期 --> 
                    <!--98.01.07 fix 不試算核准申請延長期限(A)+(B)input type="button" name="calculus" value="試算" onClick="javascript:calculusDate(this.document.forms[0]);"-->        
				</td>
			</tr>
			
			<tr>
	          
				<td width='70' bgcolor='#D8EFEE' rowspan="4" height="124"><font color="#FF0000">審查總結</font><font color="#0000FF">(縣/市政府評註)</font></td>
				<td width='60' bgcolor='#D8EFEE' rowspan="4" height="124"><font color="#0000FF">是否符合行政院農業委員會93.10.15農授金字第0935080230號令規定</font></td>
				<td width='90' bgcolor='#D8EFEE' >是否提足備抵趺價損失</td>
				<td width='380' bgcolor='#e7e7e7'>
					<select name='damage_yn' size="1">
					<%
						String selected1 = " ";
						String selected2 = " ";
						String selected3 = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
							
							String YN1 = String.valueOf(Bean.getValue("damage_yn"));
		           			
		           			if(YN1 != null){
		           				
		            			if( YN1.equals("Y"))
		              				selected1 = "selected";
		            			else if(YN1.equals("N"))
		            				selected2 = "selected";
		            			else 
		            				selected3 = "selected";
		            				
		            		}
						}//end of if (WLX08_S_Edit...
					%>
		      			<option value=" " <%=selected3 %> >  </option>
		                <option value="Y" <%=selected1 %> >是</option>
		                <option value="N" <%=selected2 %> >否</option>
					</select>
					<font color='red' size=4>*</font>                              
				</td>
			</tr>
		
			<tr>
				<td width='90' align='left' bgcolor='#D8EFEE' height="24">是否有積極處分之事實</td>
				<td width='380' bgcolor='e7e7e7' >
					<select name='disposal_fact_yn' size="1">
					<%
						selected1 = " ";
						selected2 = " ";
						selected3 = " ";
		
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		
							String YN2 = String.valueOf(Bean.getValue("disposal_fact_yn"));
		           			
		           			if(YN2 != null){
		
		            			if( YN2.equals("Y") )
		              				selected1 = "selected";
		            			else if(YN2.equals("N"))
		            				selected2 = "selected";
		            			else 
		            				selected3 = "selected";
		            		}
		        		}//end of if (WLX08_S_Edit...
		     		%>
		      			<option value=" " <%=selected3 %> >  </option>
		                <option value="Y" <%=selected1 %> >是</option>
		                <option value="N" <%=selected2 %> >否</option>
					</select>
					<font color='red' size=4>*</font>                              
	           	</td>
			</tr>
       
			<tr>
				<td width='90' align='left' bgcolor='#D8EFEE'>未來處分計劃是否合理可行</td>
				<td width='380' bgcolor='#e7e7e7' >
					<select  name='disposal_plan_yn' size="1">
					<%
						selected1 = " ";
						selected2 = " ";
						selected3 = " ";
		
		        		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		
							String YN3 = String.valueOf(Bean.getValue("disposal_plan_yn"));
							
		           			if(YN3 != null){
		           			
		           				if( YN3.equals("Y") )
		              				selected1 = "selected";
		            			else if(YN3.equals("N"))
		            				selected2 = "selected";
		            			else 
		            				selected3 = "selected";
		            		}
		         		}//end of if (WLX08_S_Edit...
		         	%>
		      			<option value=" " <%=selected3 %> >  </option>
		                <option value="Y" <%=selected1 %> >是</option>
		                <option value="N" <%=selected2 %> >否</option>
					</select>
					<font color='red' size=4>*</font>                              
				</td>
			</tr>
		
			<tr>
				<td width='90' align='left' bgcolor='#D8EFEE' >審核結果</td>
				<td width='380' bgcolor='#e7e7e7' >
					<select name='auditresult_yn' size="1" >
					<%
					    selected1 = " ";
					    selected2 = " ";
					    selected3 = " ";
		
		         		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		
							String YN4 = String.valueOf(Bean.getValue("auditresult_yn"));
		           			
		           			if(YN4 != null){
		
		            			if( YN4.equals("Y") )
		              				selected1 = "selected";
		            			else if(YN4.equals("N"))
		            				selected2 = "selected";
		            			else 
		            				selected3 = "selected";
		            		}
		          		}//end of if (WLX08_S_Edit...
		
		      		%>
		      			<option value=" " <%=selected3 %> >  </option>
		                <option value="Y" <%=selected1 %> >合格</option>
		                <option value="N" <%=selected2 %> >不合格</option>
		                
		            </select>
		            <font color='red' size=4>*</font>                              
				</td>
			</tr>
			
			<tr>
				<td width="600" bgcolor='#D8EFEE' colspan="4" align= 'Center' >
					<!-- 按了之後 抄錄最近一筆異動 -->
					<input type="button" name="Load_change_Data" value="抄錄本年本季最近一筆文號及日期" onClick="javascript:LoadRecentData(this.document.forms[0],'Rpt_BOAF');">    
					<font color='red' size=2>(請注意文號及日期都將抄入)</font>
				</td>
			</tr>
			
			<tr>
				<td width="70" bgcolor='#D8EFEE' rowspan="4"> <font color="#0000FF">縣市政府</font></td>
				<td width="60" bgcolor='#D8EFEE' rowspan="2">審核核准</td>
				<td width="90" bgcolor='#D8EFEE' >核准文號</td>
				<td width="380" bgcolor='#e7e7e7'>
					<input type='text' name='DocNo' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("applyok_docno")!= null){
										out.print(Bean.getValue("applyok_docno"));
									}
								%>"  
						size='30' maxlength='30' >
					<font color='red' size=4>*</font>         
				</td>
			</tr>
	
			<tr>
				<td width="90" bgcolor='#D8EFEE'>核准日期</td>
				<td width="380" bgcolor='#e7e7e7'>
					<%
						Date get_ApplyOK_date;
						
						String ApplyOK_y  = " ";
						String ApplyOK_m = " ";
						String ApplyOK_d  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("applyok_date") != null){	
							get_ApplyOK_date = (Date)Bean.getValue("applyok_date");
							ApplyOK_y = String.valueOf(get_ApplyOK_date.getYear()-11);       
							ApplyOK_m = String.valueOf(get_ApplyOK_date.getMonth()+1);
							ApplyOK_d = String.valueOf(get_ApplyOK_date.getDate());
						}
					%>
					<input type='text' name='ApplyOK_y' value="<%=ApplyOK_y%>" size='3' maxlength='3' onblur='CheckYear(this)'>&nbsp;年      
					
					<select id="text" name='ApplyOK_m'>
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+ApplyOK_m+"\">"+ApplyOK_m+"</option>");
		            	 }
		            	 else{
		            	 	out.print("<option value=\" \"> </option>");
		            	 }
		            %>
						<option value="1" >1</option>
						<option value="2" >2</option>
						<option value="3" >3</option>
						<option value="4" >4</option>
						<option value="5" >5</option>
						<option value="6" >6</option>
						<option value="7" >7</option>
						<option value="8" >8</option>
						<option value="9" >9</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
					</select>月</font>                           
		            <font color='red' size=4>*</font>          
					
					<select id="text" name='ApplyOK_d'>
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+ApplyOK_d+"\">"+ApplyOK_d+"</option>");
		            	 }
		            	 else{
		            	 	out.print("<option value=\" \"> </option>");
		            	 }
		            %>
						<option value="01" >1</option>
						<option value="02" >2</option>
						<option value="03" >3</option>
						<option value="04" >4</option>
						<option value="05" >5</option>
						<option value="06" >6</option>
						<option value="07" >7</option>
						<option value="08" >8</option>
						<option value="09" >9</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
						<option value="13" >13</option>
						<option value="14" >14</option>
						<option value="15" >15</option>
						<option value="16" >16</option>
						<option value="17" >17</option>
						<option value="18" >18</option>
						<option value="19" >19</option>
						<option value="20" >20</option>
						<option value="21" >21</option>
						<option value="22" >22</option>
						<option value="23" >23</option>
						<option value="24" >24</option>
						<option value="25" >25</option>
						<option value="26" >26</option>
						<option value="27" >27</option>
						<option value="28" >28</option>
						<option value="29" >29</option>
						<option value="30" >30</option>
						<option value="31" >31</option>
					</select>日</font>                           
		            <font color='red' size=4>*</font>        			
				</td>
			</tr>
			
			<tr>
				<td width="70" bgcolor='#D8EFEE' rowspan="2" >函報農金局</td>
				<td width="90" bgcolor='#D8EFEE'>函報文號</td>
				<td width="380" bgcolor='#e7e7e7'>
					<input type='text' name='BOAF_DocNo' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("report_boaf_docno")!= null){
										out.print(Bean.getValue("report_boaf_docno"));
									}
								%>" 
						size='30' maxlength='30'>
					     
				</td>          
			</tr>              
			                   
			<tr>
				<td width="75" bgcolor='#D8EFEE'>函報日期</td>
				<td width="380" bgcolor='#e7e7e7'>
					<%
						Date get_BOAF_date;
						
						String BOAF_y  = " ";
						String BOAF_m = " ";
						String BOAF_d  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("report_boaf_date")!= null){	
							get_BOAF_date = (Date)Bean.getValue("report_boaf_date");
							BOAF_y = String.valueOf(get_BOAF_date.getYear()-11);       
							BOAF_m = String.valueOf(get_BOAF_date.getMonth()+1);
							BOAF_d = String.valueOf(get_BOAF_date.getDate());
						}
					%>
					<input type='text' name='BOAF_y' value="<%=BOAF_y%>" size='3' maxlength='3' onblur='CheckYear(this)'>&nbsp;年      
					<select id="text" name='BOAF_m'>
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+BOAF_m+"\">"+BOAF_m+"</option>");
		            	 }
		            	 else{
		            	 	out.print("<option value=\" \"> </option>");
		            	 }
		            %>
						<option value="1" >1</option>
						<option value="2" >2</option>
						<option value="3" >3</option>
						<option value="4" >4</option>
						<option value="5" >5</option>
						<option value="6" >6</option>
						<option value="7" >7</option>
						<option value="8" >8</option>
						<option value="9" >9</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
					</select>月</font>                           
		            <font color='red' size=4>*</font>        
		            	
					<select id="text" name='BOAF_d'>
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+BOAF_d+"\">"+BOAF_d+"</option>");
		            	 }
		            	 else{
		            	 	out.print("<option value=\" \"> </option>");
		            	 }
		            %>
						<option value="1" >01</option>
						<option value="2" >02</option>
						<option value="3" >03</option>
						<option value="4" >04</option>
						<option value="5" >05</option>
						<option value="6" >06</option>
						<option value="7" >07</option>
						<option value="8" >08</option>
						<option value="9" >09</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
						<option value="13" >13</option>
						<option value="14" >14</option>
						<option value="15" >15</option>
						<option value="16" >16</option>
						<option value="17" >17</option>
						<option value="18" >18</option>
						<option value="19" >19</option>
						<option value="20" >20</option>
						<option value="21" >21</option>
						<option value="22" >22</option>
						<option value="23" >23</option>
						<option value="24" >24</option>
						<option value="25" >25</option>
						<option value="26" >26</option>
						<option value="27" >27</option>
						<option value="28" >28</option>
						<option value="29" >29</option>
						<option value="30" >30</option>
						<option value="31" >31</option>
					</select>日</font>                           
		            <font color='red' size=4>*</font>          
				</td>
			</tr>
		
		</table>
		</center>
		</td>
	</tr>
  <!-- ---------------------------------------------------------------------------------------------------------------------------- -->
	
	<tr>
    	<td width="600">
       		<center>
          		<jsp:include page="getMaintainUser.jsp" flush="true" /><!--載入維護者資訊 -->
       		</center>
    	</td>
	</tr>
	
	<tr>
    	<td width="600">
      		<center>
        		<table border="0" cellspacing="1">
          			<tr>
						<td width="66">
  							<a href="javascript:doSubmit(this.document.forms[0],'Update','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_updateb.gif',1)">
  							<img src="images/bt_update.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
						</td>
						<td width="93">
  							<a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)">
  							<img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
						</td>
          			</tr>
        		</table>
      		</center>
      </td>
	</tr>

	<tr>
		<td width="600">
      		<table class="sbody" border="0" cellspacing="1">
				<tr>
					<td colspan="2">
            			<font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle">
            			<font color="#007D7D">使用說明 : </font></font>     
          			</td>
        		</tr>
   				<tr>
					<td width="16">&nbsp;</td>
          			<td width="577"> 
						<ul>
						<li>確認輸入資料無誤後, 按【確定】即將本表上的資料, 於資料庫中建檔,建檔完成按[繼續新增下一筆]可在本畫面繼續新增建檔,或按[回查詢主畫面]回至查詢清單畫面。</li>    
						<li>如果本季無須申報資料,仍須選確定(Y)並按[本季無承受擔保品延長處分申報資料]按鈕,以示確定</li>
						<li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>
						<li>欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</font></li>    
						<li>如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>    
						<li>【<font color="red" size=4>*</font>】為必填欄位。</li>
						<li>如果審查總結中,有乙項填「否」,而審核結果填「合格」會先出現「審查結果欄填註與細項欄內有不一致,請確定?」警訊,若你確定仍會依你所建檔結果存檔</li>
						</ul>
          			</td>
        		</tr>
			</table>
    	</td>
	</tr>
</table>

</body>
</html>

