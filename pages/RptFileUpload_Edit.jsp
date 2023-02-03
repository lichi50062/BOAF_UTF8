<%
// 98.07.30 add 農業金庫各類報表上傳作業 by 2295
// 98.10.09 fix 在MIS系統才可上傳至檢查局 by 2295
// 99.10.13 add 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
//100.06.01 fix 日期顯示 by 2295
//100.11.30 fix 上傳至檢查局日期顯示 by 2295
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
    
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String szrpt_code = ( request.getParameter("RPT_CODE")==null ) ? "" : (String)request.getParameter("RPT_CODE");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");	
	//登入者資訊			
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");	
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
	    
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	System.out.println("RptFileUpload.act="+act);
	System.out.println("RptFileUpload.szrpt_code="+szrpt_code);
	System.out.println("RptFileUpload.S_YEAR="+S_YEAR);		    
	System.out.println("RptFileUpload.S_MONTH="+S_MONTH);
	System.out.println("RptFileUpload.firstStatus="+firstStatus);		    
    List reportList = (List)request.getAttribute("reportList");	
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();		
	
	//可查詢到的報表名稱
	//代出所有報表
	sqlCmd.append(" select  distinct  rpt_nof.rpt_code,rpt_nof.rpt_name,rpt_nof.rpt_output_type,rpt_nof.rpt_include ");
	sqlCmd.append(" from rpt_nof ");
	sqlCmd.append(" where rpt_nof.rpt_code like 'Agri%'");
	sqlCmd.append(" order by  rpt_nof.rpt_code");
    List rptname_list = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//報表名稱    
    if(szrpt_code.equals("")){
       szrpt_code = (String)((DataObject)rptname_list.get(0)).getValue("rpt_code");
    }
    
    
	//取得RptFileUpload的權限
	Properties permission = ( session.getAttribute("RptFileUpload")==null ) ? new Properties() : (Properties)session.getAttribute("RptFileUpload"); 
	if(permission == null){
       System.out.println("RptFileUpload_Edit.permission == null");
    }else{
       System.out.println("RptFileUpload_Edit.permission.size ="+permission.size());               
    }	    
    
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
    
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/RptFileUpload.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「農業金庫財務資料上傳作業」</title>
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
<form method=post ENCTYPE="multipart/form-data" action='/pages/RptFileUpload.jsp' onload='javascript:message();'>
<input type="hidden" name="act" value="Status">   
<%if(reportList != null && reportList.size() != 0){%>
<input type="hidden" name="row" value="<%=reportList.size()+1%>">   
<%}%>
<table width="695" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
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
		  <table width="695" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="695" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          「農業金庫各類報表上傳作業」 
                        </center>
                        </b></font> </td>
                      <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="695" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=695" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                      <td><table width=695 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                    
                          <tr>						  
						  <td width='20%' class="<%=nameColor%>" align='left'>報表名稱</td>
                          <td width='80%' class="<%=textColor%>">	
                            <select name='RPT_CODE'>                                                                                    
                            <%
                            if(rptname_list != null){
                             for(int i=0;i<rptname_list.size();i++){%>
                            <option value="<%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_code")%>"                                                        
                            <%if(szrpt_code.equals((String)((DataObject)rptname_list.get(i)).getValue("rpt_code"))){ 
                                 out.print("selected");                                
                              }   
                            %>
                            ><%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_name")%></option>                            
                            <%}
                            }//end of if
                            %>
                            </select>    
                            &nbsp; 
                            <a href="javascript:doSubmit(this.document.forms[0],'List');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                                                       
                          </td>             
                          </tr>                          
                          
                          
                          <tr>                          
						  <td width='20%' class="<%=nameColor%>" align='left'>申報年月</td>
                          <td width='80%' class="<%=textColor%>">	
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
        						 <input type=hidden name=S_DATE value=''>
                          </td>                                   
                          </tr>
                          
                          
                          
                          <%if(bank_type.equals("1") || lguser_id.equals("A111111111")){%>
                          <tr class="sbody">                          
                          <td width='20%' class="<%=nameColor%>" align='left'>上傳檔案位置</td>
                          <td width='80%' class="<%=textColor%>">                          
                            <input type=file size=40 name=UpFileName >&nbsp;&nbsp;&nbsp;&nbsp; 
                            <%if(permission != null && permission.get("up") != null && permission.get("up").equals("Y")){//Query %>                   	        	                                   		     
				         		<a href="javascript:doSubmit(this.document.forms[0],'Upload');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_uploadb.gif',1)"><img src="images/bt_upload.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                    		
                      		<%}%>
                          </td>         
                          </tr>    
                          <%}%>                                                		      					      
                          </table>      
                      </td>    
                      </tr>
                      <!--tr> 
                		<td><div align="center"> 
                    		<table width="243" border="0" cellpadding="1" cellspacing="1">
                      		<tr>     				        
                      		<%if(permission != null && permission.get("up") != null && permission.get("up").equals("Y")){//Query %>                   	        	                                   		     
				         		<td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Upload');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_uploadb.gif',1)"><img src="images/bt_upload.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                         		<td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0],'Reset');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>                        
                      		<%}%>
                      		</tr>
                    		</table>
                  		  </div></td>
              		  </tr-->
                      
                      <tr> 
                      <td><table width=695 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                              	
                        <%if(Utility.getProperties("APServer")!=null && Utility.getProperties("APServer").equals("MIS")){//98.10.09 在MIS系統才可上傳至檢查局 by 2295
                             if(bank_type.equals("2") || lguser_id.equals("A111111111")){//農金局.A111111111可上傳至檢查局
                        %>
                        <tr class="sbody">
                          <td width='100%' colspan=8 class="<%=nameColor%>">
 						    <a href="javascript:doSubmit(this.document.forms[0],'UploadFeb');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_uploadb.gif',1)"><img src="images/bt_upload.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                    		
                          </td>
                        </tr>
                        <%   }
                          }%>     	                      
                      <%                     
                      int i = 0;      
                      String bgcolor="#D3EBE0";  
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";                           
                      if(reportList != null){ %>
                       	  <tr>                       	  
                            <td class="<%=listTitle%>" width="30">序號</td>  
                            <td class="<%=listTitle%>" width="30">選項</td>  
                            <td class="<%=listTitle%>" width="170">報表名稱</td>
            				<td class="<%=listTitle%>" width="54">報表年月</td>
            				<td class="<%=listTitle%>" width="141">檔名</td>            				
            				<td class="<%=listTitle%>" width="90">檔案上傳日期</td> 
            				<td class="<%=listTitle%>" width="52">版本別</td> 
            				<td class="<%=listTitle%>" width="87">傳至檢查局</td>				            				
					      </tr>   
                   		    <% if(reportList.size() == 0){%>
                   			   <tr class="<%=list1Color%>">
                   			   <td colspan=8 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }else{ 
                   			    String update_date="";
                   			    String upload_date="";                   			    
                   			    String yyymm="";
                   			    DataObject bean = null;
                   			    while(i < reportList.size()){    
                   			      listColor = (i % 2 == 0)?list2Color:list1Color;	                 		      
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                    		      bean = (DataObject)reportList.get(i);
                    		      if((String)bean.getValue("s_updatedate") != null){
                    		         update_date = (String)bean.getValue("s_updatedate");
                    		         //100.06.01 fix 日期顯示 by 2295
                    		         if(update_date.startsWith("1")){
                                		update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                                	 }else{
                                		update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                                	 }                            		         
                    		         System.out.println("update_date"+update_date);
                    		      }
                    		      if((String)bean.getValue("s_uploaddate") != null){
                    		         upload_date = (String)bean.getValue("s_uploaddate");
                    		         //100.11.30 fix 日期顯示 by 2295
                    		         if(upload_date.startsWith("1")){
                                		upload_date = upload_date.substring(0,3)+"/"+upload_date.substring(3,5)+"/"+upload_date.substring(5,7)+ " " + upload_date.substring(8,upload_date.length());	                    		
                                	 }else{
                                		upload_date = upload_date.substring(0,2)+"/"+upload_date.substring(2,4)+"/"+upload_date.substring(4,6)+ " " + upload_date.substring(7,upload_date.length());	                    		
                                	 }                            		
                    		         System.out.println("upload_date"+upload_date);
                    		      }
                    		      if( bean.getValue("s_yymm") != null){
                    		         yyymm = (String)bean.getValue("s_yymm");
                    		         if(yyymm.length() == 4){
		            					 yyymm = "0"+yyymm;
		          					 } 
                    		      }
                    		%>                         	       
                    		<tr class="<%=listColor%>">                       	  
                               <td width="30"><%=i+1%></td> 
                               <td width="30"><input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=S_YEAR%>:<%=S_MONTH%>:<%=szrpt_code%>:<%=(String)bean.getValue("rpt_version")%>"</td>                                                                                                                         
                               <td width="170">
                               <%if( bean.getValue("rpt_name") != null ) out.print("<a href='RptFileUpload.jsp?act=Download&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&RPT_CODE="+szrpt_code+"&rpt_version="+(String)bean.getValue("rpt_version")+"'>"+(String)bean.getValue("rpt_name")+"</a>"); else out.print("&nbsp;");%>
                               </td>
            				   <td width="54">
            				   <%if( bean.getValue("s_yymm") != null ) out.print((String)bean.getValue("s_yymm")); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="141">            				  
            				   <%if( bean.getValue("rpt_fname") != null ) out.print((String)bean.getValue("rpt_fname")); else out.print("&nbsp;");%>
            				   </td>            				
            				   <td width="90">            				  
            				   <%if(!update_date.equals("")) out.print(update_date); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="52">
            				   <%if( bean.getValue("rpt_version") != null ) out.print((String)bean.getValue("rpt_version")); else out.print("&nbsp;");%>
            				   </td> 
            				   <td width="87">
            				   <%if( bean.getValue("feb_flag") != null ){
            				        if(((String)bean.getValue("feb_flag")).equals("Y")){//已上傳至檢查局
            				           out.print(upload_date);    
            				        }else{
            				          out.print((String)bean.getValue("feb_flag"));    
            				        }  
            				     }else out.print("&nbsp;");  
            				   %>
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
</html>
