<%
//94.02.02 fix 區分網際網路申報跟MIS管理系統的配色 by 2295
//99.02.09 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.02.10 fix 取得該功能項目權限移至Utility.getPermission by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "AS001W";
	String act = Utility.getTrimString(dataMap.get("act"));
	String szbank_type = Utility.getTrimString(dataMap.get("bank_type"));
	String sztbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	
	System.out.print(report_no+"_List.sztbank_no="+sztbank_no);	
	System.out.print(":szbank_type="+szbank_type);
	System.out.println(":szbank_no="+sztbank_no);
	List BN01List = (List)request.getAttribute("BN01List");		
	DataObject BN01bean = null;
	if(BN01List == null){
	   System.out.println("BN01List == null");
	}else{
	   System.out.println("BN01List.size()="+BN01List.size());
	}
	
	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/AS001W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「總機構代號」建檔管理 </title>
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
<table width="600" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「總機構代號」建檔管理 
                        </center>
                        </b></font> </td>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
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
                     <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">

                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構類別</td>
						  <% List bank_type = Utility.getBank_Kind("001","null");%>						  
						  <td width='85%' colspan=2 class="<%=textColor%>"'>
                            <select name='BANK_TYPE'>                                                        
                            <%
                            DataObject bank_type_bean = null;
                            for(int i=0;i<bank_type.size();i++){
                                bank_type_bean = (DataObject)bank_type.get(i);
                            %>
                            <option value="<%=(String)bank_type_bean.getValue("cmuse_id")%>"                            
                            <%if(szbank_type.equals((String)bank_type_bean.getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)bank_type_bean.getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <%if(Utility.getPermission(request,report_no,"Q")){ %>                   	        	                                   		     
                            <a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                            <%}%>
                            <%if(Utility.getPermission(request,report_no,"A")){ %>                   	        	                                   		     
                            <a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>                       
                            <%}%>
                           </td>
                          </td>
                           
                          </tr>     
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">總機構代碼</td>						  
						  <td width='85%' colspan=2 class="<%=textColor%>">
						    <input type='text' name='TBANK_NO' value="<%=sztbank_no%>" size='7' maxlength='7' >                                                      
                            </select>
                          </td>
                          </tr>   
                          </table>      
                      </td>    
                      </tr>
                      <tr><td><table><tr><br></tr></table></td></tr>  
                      
                      <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>"">                      
                      <%
                      String tmpbank_type="";                      
                      int i = 0;                           
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";
                      if(BN01List != null){ %>
                       	  <tr>
                            <td class="<%=listTitle%>" width="25">&nbsp;</td>           
                            <td class="<%=listTitle%>" width="120">機構類別</td>
            				<td class="<%=listTitle%>" width="100">總機構代碼</td>
            				<td class="<%=listTitle%>" width="200">機構名稱</td>
            				<td class="<%=listTitle%>" width="100">異動日期</td>            				
					      </tr>   
                   		    <% if(BN01List.size() == 0){%>
                   			   <tr class="<%=list1Color%>">
                   			   <td colspan=11 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < BN01List.size()){ 
                    		      BN01bean = (DataObject)BN01List.get(i);
                    		      listColor = (i % 2 == 0)?list2Color:list1Color;	                      		                        		      
                      %>                         	  
                          <tr>
                            <td class="<%=listColor%>" width="25"><%=i+1%></td>           
            				<td class="<%=listColor%>" width="120">
            				<%if( BN01bean.getValue("cmuse_id") != null ){                      		
                       			  if(tmpbank_type.equals((String)BN01bean.getValue("cmuse_id"))){   
                       			     out.print("&nbsp;");
                    		      }else{
                    		        out.print((String)BN01bean.getValue("cmuse_name"));                    		        
                    		        tmpbank_type = (String)BN01bean.getValue("cmuse_id");
                    		      }  
                    		  }    
                    		%>    
                    		</td>
            				<td class="<%=listColor%>" width="100"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)BN01List.get(i)).getValue("bank_no")%>');"><%if( ((DataObject)BN01List.get(i)).getValue("bank_no") != null ) out.print((String)((DataObject)BN01List.get(i)).getValue("bank_no")); else out.print("&nbsp;");%></a></td>
            				<td class="<%=listColor%>" width="200"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)BN01List.get(i)).getValue("bank_no")%>');"><%if( ((DataObject)BN01List.get(i)).getValue("bank_name") != null ) out.print((String)((DataObject)BN01List.get(i)).getValue("bank_name")); else out.print("&nbsp;");%></a></td>
            				<td class="<%=listColor%>" width="100">
            				<%if( BN01bean.getValue("update_date") != null ){%>            				
            				<%=Utility.getCHTdate((BN01bean.getValue("update_date")).toString().substring(0, 10), 0)%>            				
            				<%}else{%>
            				 &nbsp;
            				<%}%>
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
</table>
</form>
</body>
</html>
