<%
// 94.12.19 create by 2295
// 95.04.21 add 全部報表發佈 by 2295
// 98.07.28 fix 只抓FR開頭的報表 by 2295
// 99.12.10 fix sqlInjection by 2808
//100.01.03 fix 日期顯示 by 2295
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
    //DataObject rptData = null;
	String bank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");				
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");					
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String szrpt_code = ( request.getParameter("RPT_CODE")==null ) ? "" : (String)request.getParameter("RPT_CODE");					
	String szrpt_output_type = ( request.getParameter("szRPT_OUTPUT_TYPE")==null ) ? "" : (String)request.getParameter("szRPT_OUTPUT_TYPE");					
	String szUnit = ( request.getParameter("Unit")==null ) ? "1000" : (String)request.getParameter("Unit");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String szbank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");				
	String szhsien_id = ( request.getParameter("HSIEN_ID")==null ) ? "ALL" : (String)request.getParameter("HSIEN_ID");
    String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		    
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	System.out.print("ZZ041W_List.act="+act);	
	System.out.print(":bank_type="+bank_type);		    
	System.out.print(":szrpt_code="+szrpt_code);		    
	System.out.print(":szrpt_output_type="+szrpt_output_type);		    
	System.out.print(":szhsien_id="+szhsien_id);		    	
	System.out.print(":S_YEAR="+S_YEAR);		    
	System.out.print(":S_MONTH="+S_MONTH);		    
	System.out.println(":firstStatus="+firstStatus);		    
    List reportList = (List)request.getAttribute("reportList");	
	String sqlCmd = "";
	List paramList =new ArrayList() ;
	//代出所有報表
	sqlCmd = " select  distinct  rpt_nof.rpt_code,rpt_nof.rpt_name,rpt_nof.rpt_output_type,rpt_nof.rpt_include "
		   + " from rpt_nof "
		   + " where rpt_nof.rpt_code like 'FR%'" //98.07.28 fix 只抓FR開頭的報表 
	       + " order by  rpt_nof.rpt_code";
    List rptname_list = DBManager.QueryDB_SQLParam(sqlCmd,null,"");//報表名稱    
    if(szrpt_code.equals("")){
       szrpt_code = (String)((DataObject)rptname_list.get(0)).getValue("rpt_code");
    }
    sqlCmd = " select  distinct  rpt_nof.rpt_code,rpt_nof.rpt_name,rpt_nof.rpt_output_type,rpt_nof.rpt_include "
		   + " from rpt_nof "
		   + " where rpt_nof.rpt_code = ?"
		   + " order by  rpt_nof.rpt_code";	
    paramList.add(szrpt_code);   
    List rptData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");//報表data  		          
   
    System.out.println("rptData.size()="+rptData.size());
	    
    // XML Ducument for 報表名稱 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"ReportListXML\">");
    out.println("<datalist>");
    for(int i=0;i< rptname_list.size(); i++) {
        DataObject bean =(DataObject)rptname_list.get(i);
        out.println("<data>");        
        out.println("<rptCode>"+bean.getValue("rpt_code")+"</rptCode>");        
        out.println("<rptName>"+bean.getValue("rpt_name")+"</rptName>");                
        out.println("<rptOutputType>"+bean.getValue("rpt_output_type")+"</rptOutputType>");                
        out.println("<rptInclude>"+bean.getValue("rpt_include")+"</rptInclude>");                
		out.println("</data>");        
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 報表名稱 end 	
    //縣市別
	List hsien_id_data = DBManager.QueryDB_SQLParam("select hsien_id,hsien_name from cd01 order by fr001w_output_order",null,"");     
    // XML Ducument for 縣市別 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"hsienListXML\">");
    out.println("<datalist>");
    for(int i=0;i< hsien_id_data.size(); i++) {
        DataObject bean =(DataObject)hsien_id_data.get(i);
        out.println("<data>");        
        out.println("<hsien_id>"+bean.getValue("hsien_id")+"</hsien_id>");        
        out.println("<hsien_name>"+bean.getValue("hsien_name")+"</hsien_name>");                        
		out.println("</data>");        
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 縣市別 end 	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ041W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「管理報表產生及發佈」</title>
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
<%if(act.equals("List")){%>
<input type="hidden" name="RPT_CODE" value="<%=szrpt_code%>">  
<input type="hidden" name="szRPT_OUTPUT_TYPE" value="<%=szrpt_output_type%>">  
<input type="hidden" name="S_MONTH" value="<%=S_MONTH%>">  
<input type="hidden" name="Unit" value="<%=szUnit%>">  
<%}%>
<%if(reportList != null && reportList.size() != 0){%>
<input type="hidden" name="row" value="<%=reportList.size()+1%>">   
<%}%>
<table width="660" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <%
          String nameColor="nameColor_sbody";
          String textColor="textColor_sbody";
          String bordercolor="#3A9D99";
        %>  
        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="210"><img src="images/banner_bg1.gif" width="210" height="17"></td>
                      <td width="240"><font color='#000000' size=4><b> 
                        <center>
                          「管理報表產生及發佈」 
                        </center>
                        </b></font> </td>
                      <td width="210"><img src="images/banner_bg1.gif" width="210" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=660" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                      <td><table width=660 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                    
                          <tr class="sbody">						  
						  <td width='15%' align='left' class="<%=nameColor%>">報表名稱</td>
                          <td width='85%' class="<%=textColor%>">	
                            <select name='RPT_CODE' onchange="javascript:changeOption(this.document.forms[0]);" <%if(act.equals("List")) out.print("disabled");%>>                                                                                    
                            <%
                            if(rptname_list != null){
                             for(int i=0;i<rptname_list.size();i++){%>
                            <option value="<%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_code")%>"                                                        
                            <%if(szrpt_code.equals((String)((DataObject)rptname_list.get(i)).getValue("rpt_code"))){ 
                                 out.print("selected");                                
                              }   
                            %>
                            ><%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_code")%>&nbsp;<%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_name")%></option>                            
                            <%}
                            }//end of if
                            %>
                            </select>    
                           </td>             
                          </tr>       
                                                         
                          <tr class="sbody">                          
                          <td width='15%' align='left' class="<%=nameColor%>">農(漁)會別</td>
                          <td width='85%' class="<%=textColor%>">
                            <select name='szRPT_OUTPUT_TYPE' <%if(act.equals("List")) out.print("disabled");%>>     
                              <%          
                              if(rptData != null){                    
                                if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("X")){//顯示農會/漁會/農漁會  				  				 
                                   if(szrpt_output_type.equals("6")){
                                      out.print("<option value='6' selected>農會</option>");
                                   }else{ 
                                      out.print("<option value='6' >農會</option>");
                                   }
                                   if(szrpt_output_type.equals("7")){
  				                      out.print("<option value='7' selected>漁會</option>");
  				                   }else{
  				                      out.print("<option value='7'>漁會</option>");
  				                   }
  				                   if(szrpt_output_type.equals("ALL")){
  				                      out.print("<option value='ALL' selected>農/漁會</option>");  			       
  				                   }else{
  				                      out.print("<option value='ALL'>農/漁會</option>");  			       
  				                   }
  		    				    }else if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("T")){//顯示農漁會  				  		    	 
  		    				       if(szrpt_output_type.equals("ALL")){
  		    				          out.print("<option value='ALL' selected>農/漁會</option>");  			         		    				       
  		    				       }else{
  		    				          out.print("<option value='ALL'>農/漁會</option>");  			       
  		    				       }
  		    				    }   
  		    				 }
                            %>
                            </select>
                            
                            <%if(act.equals("Qry")){%>                            
                            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
                            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
                            <input type="button" value="查詢清單" onclick="javascript:doSubmit(this.document.forms[0],'List');"> &nbsp; &nbsp;                       
                            <input type="button" value="全部報表產生發佈" onclick="javascript:doSubmit(this.document.forms[0],'GenerateRptALL');">  &nbsp;                                                                                                                                                            
                            <%}else{%>                            
                            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;                               
                            <input type="button" value="回查詢條件區" onclick="javascript:doSubmit(this.document.forms[0],'goQry');">  &nbsp;                              
                            <input type="button" value="報表產生發佈" onclick="javascript:doSubmit(this.document.forms[0],'GenerateRpt');">  &nbsp;                                                                                                                                                                                        
                            <%}%>   
                          </td>         
                          </tr>
                          
                          
                          <tr class="sbody"> 
                          
                          <td width='15%' align='left' class="<%=nameColor%>">縣市別</td> 
                          <td width='85%' class="<%=textColor%>">                                                      
                               <select name="HSIEN_ID" <%if(act.equals("List")) out.print("disabled");%>>                                            
                                <%if(rptData != null){
                                   System.out.println("rpt_output_type="+(String)((DataObject)rptData.get(0)).getValue("rpt_output_type"));
                                   System.out.println("rpt_include="+(String)((DataObject)rptData.get(0)).getValue("rpt_include"));
                                   if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("T")){//顯示全部縣市%>                                                     
                                         <option value="ALL">全部縣市</option>
                                <% }else if (!((String)((DataObject)rptData.get(0)).getValue("rpt_include")).equals("X")){//顯示全部縣市%>                                                               
                                         <option value="ALL">全部縣市</option>
                                <% }else{%>
                                         <option value="ALL">全部縣市</option>
                                         <%for(int i=0;i<hsien_id_data.size();i++){%>                                
                                		   <option value="<%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")%>"
                                           <%if(((String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")).equals(szhsien_id)) out.print("selected");%>
                                           ><%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_name")%></option>                            
                                        <%}%>
                                <% }
                                  }//end of rptData != null
                                %>
                                
                              </select>
                            </td>
                          </tr>
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' class="<%=nameColor%>">基準年月</td>
                          <td width='85%' class="<%=textColor%>">	
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' <%if(act.equals("List")) out.print("readonly");%>>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH <%if(act.equals("List")) out.print("disabled");%>>
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
                          <td width='15%' align='left' class="<%=nameColor%>">金額單位</td>
                          <td width='85%' class="<%=textColor%>">                                                      
                            <select size="1" name="Unit" <%if(act.equals("List")) out.print("disabled");%>>
                            <option value ='1' <%if(szUnit.equals("1")) out.print("selected");%>>元</option>
                            <option value ='1000' <%if(szUnit.equals("1000")) out.print("selected");%>>千元</option>
                            <option value ='10000' <%if(szUnit.equals("10000")) out.print("selected");%>>萬元</option>
                            <option value ='1000000' <%if(szUnit.equals("1000000")) out.print("selected");%>>百萬元</option>
                            <option value ='10000000' <%if(szUnit.equals("10000000")) out.print("selected");%>>千萬元</option>
                            <option value ='100000000' <%if(szUnit.equals("100000000")) out.print("selected");%>>億元</option>
                          </select>                          
                            
                          </td>         
                          </tr>                                                    		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=660 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                              	
                        <tr class="sbody">
                          <td width='100%' colspan=10 class="<%=nameColor%>">		                         	 						  
 							<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						  
                          </td>
                        </tr>     		                      
                      <%                     
                      int i = 0;      
                      String bgcolor="#D3EBE0";     
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";                          
                      if(reportList != null){ %>
                       	  <tr >                       	  
                            <td class="<%=listTitle%>" width="30">序號</td>                                  
                            <td class="<%=listTitle%>" width="30">選項</td>                            
                            <td class="<%=listTitle%>" width="240">總機構代碼名稱(或檔名代碼與報表名稱)</td>
            				<td class="<%=listTitle%>" width="60">農(漁)會別</td>            				
            				<td class="<%=listTitle%>" width="40">應申報</td>
            				<td class="<%=listTitle%>" width="60">完成申報</td>            				
            				<td class="<%=listTitle%>" width="80">檔案名</td>            				            				
            				<td class="<%=listTitle%>" width="100">檔案產生日</td>            				            				
					      </tr>   
                   		    <% if(reportList.size() == 0){%>
                   			   <tr class="<%=list1Color%>">                   			   
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }else{ 
                   			    String update_date="";
                   			    List update_date_tmp = new LinkedList();//100.01.03 add
                   			    String bank_type_name="";
                   			    String hsien_id="";
                   			    while(i < reportList.size()){  
                   			      listColor = (i % 2 == 0)?list2Color:list1Color;	                    		                         		      
                   			      update_date="";
                   			      bank_type_name="";
                   			      if( ((DataObject)reportList.get(i)).getValue("bank_type_name") != null ){
                   			         bank_type_name = ((String)((DataObject)reportList.get(i)).getValue("bank_type_name"));
                   			         if(bank_type_name.equals("農會")){
                   			            bank_type_name="6";
                   			         }else if(bank_type_name.equals("漁會")){
                   			            bank_type_name="7";
                   			         }else if(bank_type_name.equals("農漁會")){
                   			            bank_type_name="T";
                   			         }   
                   			      } 
                   			      hsien_id=((String)((DataObject)reportList.get(i)).getValue("hsien_id"));
                   			      System.out.println("'"+(String)((DataObject)reportList.get(i)).getValue("rpt_fname")+"'");
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                    		      if(((String)((DataObject)reportList.get(i)).getValue("s_updatedate") != null) && (!(((String)((DataObject)reportList.get(i)).getValue("s_updatedate")).trim()).equals(""))){
                    		         update_date = (String)((DataObject)reportList.get(i)).getValue("s_updatedate");                    		                             		         
                    		         //100.01.03 fix 日期顯示 by 2295
                    		         if(update_date.startsWith("1")){
                                		update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                                	 }else{
                                		update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                                	 }        
                                	 System.out.println("update_date='"+update_date+"'");
                    		      }
                    		%>                         	       
                    		<tr class="<%=listColor%>">                       	  
                               <td width="30"><%=i+1%></td>                                  
                               <td width="30"><input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=bank_type_name%>:<%=hsien_id%>:<%if( ((DataObject)reportList.get(i)).getValue("rpt_fname") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("rpt_fname"));%>:<%if( ((DataObject)reportList.get(i)).getValue("s_report_name") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("s_report_name"));%>"</td>                            
                               <td width="240">
                               <%if( ((DataObject)reportList.get(i)).getValue("s_report_name") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("s_report_name")); else out.print("&nbsp;");%>
                               </td>
            				   <td width="60">
            				   <%if( ((DataObject)reportList.get(i)).getValue("bank_type_name") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("bank_type_name")); else out.print("&nbsp;");%>            				   
            				   </td>            				
            				   <td width="40">
            				   <%if( ((DataObject)reportList.get(i)).getValue("t_cnt") != null ) out.print((((DataObject)reportList.get(i)).getValue("t_cnt")).toString()); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="60">
            				   <%if( ((DataObject)reportList.get(i)).getValue("ok_cnt") != null ) out.print((((DataObject)reportList.get(i)).getValue("ok_cnt")).toString()); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="80">            				  
            				   <%if(( ((DataObject)reportList.get(i)).getValue("rpt_fname") != null ) && !(((String)((DataObject)reportList.get(i)).getValue("rpt_fname")).trim()).equals("")) out.print((String)((DataObject)reportList.get(i)).getValue("rpt_fname")); else out.print("&nbsp;");%>
            				   </td>            				
            				   <td width="100">            				  
            				   <%if(!update_date.equals("")) out.print(update_date); else out.print("&nbsp;");%>
            				   </td>            				            				
					        </tr>   
                   			<%  i++; 
                   			    }//end of while
                   			   }//end of if%>
                   	  <%}%>		
					      </table>      
                      </td>    
                      </tr>
                                  
      </table></td>
  </tr> 
</table>
</form>
</body>
<%if(firstStatus.equals("true")){%>
<script language="JavaScript" >
<!--
changeOption(this.document.forms[0]);
-->
</script>
<%}%>
</html>
