<%
//93.12.22 fix 目前只檢核農會/漁會/農漁共用中心/農業金融局/農業信用保證基金/全部類別 by 2295  
//96.12.27 add 全國農業金庫增加檢核 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%
	List bank_type = (List)request.getAttribute("bank_type");
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	String lgbank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
%>


<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileCheck.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>人工檢核轉檔作業</title>
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
<form name=frmWMFileEdit method=post>
<input type="hidden" name="act" value="Status">  
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
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                      <td width="244"><font color='#000000' size=4><b> 
                        <center>
                          人工檢核轉檔作業 
                        </center>
                        </b></font> </td>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
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
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr align=left bgcolor='#D2F0FF'>                           
                          <td align='center' width=100% class="sbody">金融機構類別</td>						  
                          </tr>
                        </Table></td>
                    </tr>
                  
                    <tr> 
                      <%if(lguser_type.equals("S")){%>		
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                   
                      <%}else{%>
                      <td><table width='600' border=1 align='center' cellpadding="0" cellspacing="0" bordercolor="#3A9D99">                   
                      <%}%>
                      <%
                          int i = 0;
                          int j = 0;
                          String bgcolor="#FFFFE6";
                          DataObject bean = null;
                          while(i<bank_type.size()){
                          bean = (DataObject)bank_type.get(i);
                          System.out.println((String)((DataObject)bank_type.get(i)).getValue("cmuse_name"));
                          bgcolor = (j % 2 == 0)?"#FFFFE6":"#F2F2F2";
                          if( (((String)bean.getValue("cmuse_id")).equals("6")) //農會
                          ||  (((String)bean.getValue("cmuse_id")).equals("7")) //漁會
                          ||  (((String)bean.getValue("cmuse_id")).equals("8")) //農漁共用中心
                          ||  (((String)bean.getValue("cmuse_id")).equals("1")) //全國農業金庫96.12.27 add
                          ||  (((String)bean.getValue("cmuse_id")).equals("2")) //農業金融局
                          ||  (((String)bean.getValue("cmuse_id")).equals("4")) //農業信用保證基金 
                          ||  (((String)bean.getValue("cmuse_id")).equals("Z")) //全部類別
                          ){
                          	  j++;
                          %>
						   <tr bgcolor='<%=bgcolor%>' class="sbody"> 	
						   <%if(lguser_type.equals("S")){%>						   					
						     <td align='center' width=100% bordercolor="#3A9D99"><a href="javascript:doSubmit(this.document.forms[0],'Query','<%=(String)bean.getValue("cmuse_id")%>');"><%=(String)bean.getValue("cmuse_name")%></a></td>	   							     
						   <%}else if(lgbank_type.equals((String)bean.getValue("cmuse_id"))){%>						   
							 <td align='center' width=100% bordercolor="#3A9D99"><a href="javascript:doSubmit(this.document.forms[0],'Query','<%=(String)bean.getValue("cmuse_id")%>');"><%=(String)bean.getValue("cmuse_name")%></a></td>	   	
						   <%}%>	
    					  </tr>        					  
    				  <%   
    				  	   }//end of if 
    				  	   i ++;
    					  }
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
