<%
// 94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//102.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "AS002W";
	
	String act = Utility.getTrimString(dataMap.get("act"));
	String szbank_type = Utility.getTrimString(dataMap.get("bank_type"));
	String sztbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	String szbank_no = Utility.getTrimString(dataMap.get("bank_no"));
	
	List BN02List = (List)request.getAttribute("BN02List");		
	if(BN02List == null){
	   System.out.println("BN02List == null");
	}else{
	   System.out.println("BN02List.size()="+BN02List.size());
	}
	
    List tbankList = (List)request.getAttribute("tbank_no");
	System.out.println("test="+tbankList.size());
    
    // XML Ducument for 總機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
    out.println("<datalist>");
    for(int i=0;i< tbankList.size(); i++) {
        DataObject bean =(DataObject)tbankList.get(i);
        out.println("<data>");
        out.println("<bankType>"+bean.getValue("bank_type")+"</bankType>");
        out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
        out.println("<bankName>"+bean.getValue("bank_no")+"  "+bean.getValue("bank_name")+"</bankName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 總機構代碼 end 	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/AS002W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「分支機構代號」建檔管理 </title>
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
                          「分支機構代號」建檔管理 
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
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>                 
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構類別</td>
						  <% List bank_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' order by input_order",null,"");%>
						  <td width='85%' colspan=2 class="<%=textColor%>"'>
                            <select name='BANK_TYPE' onchange="javascript:changeOption(document.forms[0]);">                                                        
                            <%for(int i=0;i<bank_type.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_id")%>"><%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_name")%></option>                            
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
						  
						  <td width='85%' colspan=2 class="<%=textColor%>"'>
                            <select name='TBANK_NO'>
                            </select>
                          </td>
                          </tr>   
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">分支機構代碼</td>						  
						  <td width='85%' colspan=2 class="<%=textColor%>"'>
						    <input type='text' name='BANK_NO' value="<%=szbank_no%>" size='7' maxlength='7' >                                                      
                            </select>
                          </td>
                          </tr>   
                          </table>      
                      </td>    
                      </tr>
                      <tr><td><table><tr><br></tr></table></td></tr>  
                      
                      <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                      <%
                      String tmpbank_type="";                      
                      String tmptbank_no="";                      
                      int i = 0;      
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";      
                      if(BN02List != null){ %>
                       	  <tr>
                            <td class="<%=listTitle%>" width="25">&nbsp;</td>           
                            <td class="<%=listTitle%>" width="120">機構類別</td>
            				<td class="<%=listTitle%>" width="120">總機構名稱</td>
            				<td class="<%=listTitle%>" width="80">分支機構代號</td>
            				<td class="<%=listTitle%>" width="180">分支機構名稱</td>
            				<td class="<%=listTitle%>" width="80">異動日期</td>            				
					      </tr>   
                   		    <% if(BN02List.size() == 0){%>
                   			   <tr class="<%=list1Color%>" >
                   			   <td colspan=11 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < BN02List.size()){ 
                    		      listColor = (i % 2 == 0)?list2Color:list1Color;	                    		      
                      %>                         	  
                          <tr >
                            <td class="<%=listColor%>" width="25"><%=i+1%></td>           
            				<td class="<%=listColor%>" width="120">
            				<%if( ((DataObject)BN02List.get(i)).getValue("cmuse_id") != null ){                      		
                       			  if(tmpbank_type.equals((String)((DataObject)BN02List.get(i)).getValue("cmuse_id"))){   
                       			     out.print("&nbsp;");
                    		      }else{
                    		        out.print((String)((DataObject)BN02List.get(i)).getValue("cmuse_name"));                    		        
                    		        tmpbank_type = (String)((DataObject)BN02List.get(i)).getValue("cmuse_id");
                    		      }  
                    		  }    
                    		%>                        		
                    		</td>
                    		<td class="<%=listColor%>" width="120">
            				<%if( ((DataObject)BN02List.get(i)).getValue("tbank_no") != null ){                      		
                       			  if(tmptbank_no.equals((String)((DataObject)BN02List.get(i)).getValue("tbank_no"))){   
                       			     out.print("&nbsp;");
                    		      }else{
                    		        out.print((String)((DataObject)BN02List.get(i)).getValue("tbank_name"));                    		        
                    		        tmptbank_no = (String)((DataObject)BN02List.get(i)).getValue("tbank_no");
                    		      }  
                    		  }    
                    		%>    
                    		</td>
            				<td class="<%=listColor%>" width="80"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)BN02List.get(i)).getValue("tbank_no")%>','<%=(String)((DataObject)BN02List.get(i)).getValue("bank_no")%>');"><%if( ((DataObject)BN02List.get(i)).getValue("bank_no") != null ) out.print((String)((DataObject)BN02List.get(i)).getValue("bank_no")); else out.print("&nbsp;");%></a></td>
            				<td class="<%=listColor%>" width="180"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)BN02List.get(i)).getValue("tbank_no")%>','<%=(String)((DataObject)BN02List.get(i)).getValue("bank_no")%>');"><%if( ((DataObject)BN02List.get(i)).getValue("bank_name") != null ) out.print((String)((DataObject)BN02List.get(i)).getValue("bank_name")); else out.print("&nbsp;");%></a></td>
            				<td class="<%=listColor%>" width="80"><%if( ((DataObject)BN02List.get(i)).getValue("update_date") != null ) out.print("&nbsp;"+Utility.getCHTdate((((DataObject)BN02List.get(i)).getValue("update_date")).toString().substring(0, 10), 0)); else out.print("&nbsp;");%></td>            				
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
<script language="JavaScript" >
<!--
setSelect(this.document.forms[0].BANK_TYPE,"<%=szbank_type%>");
changeOption(this.document.forms[0]);
setSelect(this.document.forms[0].TBANK_NO,"<%=sztbank_no%>");
-->
</script>

</html>
