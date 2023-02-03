<%
//97.08.26 create 縣市政府變現性資產查核 by 2295
//98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	List BOAF_ASSETCHECK= (request.getAttribute("BOAF_ASSETCHECK")==null)?null:(List)request.getAttribute("BOAF_ASSETCHECK");
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX011W")==null ) ? new Properties() : (Properties)session.getAttribute("FX011W");
    if (permission == null) {
        System.out.println("FX011W_List.permission == null");
    }else {
        System.out.println("FX011W_List.permission.size ="+permission.size());
    }
    String m_year="",examine="",bank_name="",chinese_name="",name="",item="",check_date="",result="",content="",remark="";    
	DataObject bean = null; 

%>
<script language="javascript" src="js/FX011W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>辦理農漁會信用部變現性資產查核</title>
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
<body topmargin="0" leftmargin="15">
<table width="815" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
            <td width="815"><img src="images/space_1.gif" width="12" height="12"></td>
            </tr>
        <tr>
          <td bgcolor="#FFFFFF" width="815">
        
              <tr>
                <td width="815" height="18">
                  <table width="815" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17" align="left"></td>
                      <td width="415" align="center"><b><font size="4">辦理農漁會信用部變現性資產查核</font></b></td>
                      <td width="200" align="right"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="815" height="50">
                <div align="left">
                <table width="815" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="815" >               
                      </td>
                    </tr>
                    <tr>                       
                    <div><jsp:include page="getLoginUser.jsp?width=815" flush="true" /></div>			 							
			 		               
                    </tr>
                    <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr>
                    <td width="815" align="right" valign="bottom">
                    <table border="1" cellspacing="1" bordercolor="#3A9D99" width="815"  class="sbody" cellpadding="0" height="20">
                      <tr>
                      <form name="date" method="post">                          				  		
                           <tr>
                           <td class="sbody" bgcolor="#E7E7E7" width="815"  valign="middle" height="20">
						  	 執行情形報告表
						  	 <input type="text" name="m_year" size="3" value="">年度						  	 
                     		 <a href="javascript:doSubmit(this.document.forms[0],'Print','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    			<img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>  
                    			&nbsp;&nbsp;
                    	      <a href="/pages/FX011W.jsp?act=new" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                            <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>		
						   </td>
						   </tr>
                      </form>
                      </tr>
                    </table>
                    </td>
                    </tr>
                    <%}%>
					
					 <tr>
                      <td  align="center" width="815" valign="top">
					  <table class="sbody" width="815" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="1" >
                         <tr bgcolor="#9AD3D0">                          
                          <td width="32" align="center" height="1">查核年度</td>
                          <td width="131" align="center" height="2">受檢單位</td>
                          <td width="70" align="center" height="2">查核日期</td>
                          <td width="125" align="center" height="2">查核人員</td>
                          <td width="120" align="center" height="2">查核項目</td>
                          <td width="152" align="center" height="2">查核結果</td>
                          <td width="181"  align="center" height="44">處理情形</td>
                          <td width="181"  align="center" height="44">備註</td>
                         </tr>
                          
                         <% if(BOAF_ASSETCHECK.size()==0){%>
                         <tr class="sbody" bgcolor="#D8EFEE"><td colspan="8" align="center"  height="19" width="835" >
                           <font class="sbody">尚無資料</font></td></tr>
						 <%}else{
						      List name_List = null;
						      List item_List = null;
						      //DBManager.QueryDB("select muser_id,muser_name from wtt01 where tbank_no='"+bank_no+"'","");
						      for(int i=0;i<BOAF_ASSETCHECK.size();i++){
						          chinese_name="";examine="";bank_name="";item="";name="";result="";content="";remark="";check_date="";
						          bean = (DataObject)BOAF_ASSETCHECK.get(i);
						          examine = bean.getValue("examine") == null?"":(String)bean.getValue("examine");
						          bank_name = bean.getValue("bank_name") == null?"":(String)bean.getValue("bank_name");
                                  item = bean.getValue("item") == null?"":(String)bean.getValue("item");
                                  name = bean.getValue("name") == null?"":(String)bean.getValue("name");
                                  result = bean.getValue("result") == null?"":(String)bean.getValue("result");
                                  content = bean.getValue("content") == null?"":(String)bean.getValue("content");
                                  remark = bean.getValue("remark") == null?"":(String)bean.getValue("remark");
                                  m_year = bean.getValue("m_year") == null?"":bean.getValue("m_year").toString();                                 
                                  check_date = bean.getValue("check_date") == null?"":bean.getValue("check_date").toString();
                                  check_date = (check_date.equals(""))?"-":Utility.getCHTdate(check_date.substring(0, 10),0);     										                                  
                                  System.out.println("name="+name);
                                  
                                  if(name.indexOf("\n") != -1){
                                     System.out.println("test1");
                                     name_List = Utility.getStringTokenizerData(name,"\n");
                                     System.out.println("name_List.size()="+name_List.size());
                                     name = "";
                                     for(int j=0;j<name_List.size();j++){
                                         name += (String)name_List.get(j)+"<br>";
                                     }                                     
                                  }
                                  if(item.indexOf("\n") != -1){                                     
                                     item_List = Utility.getStringTokenizerData(item,"\n");
                                     System.out.println("item_List.size()="+item_List.size());
                                     item = "";
                                     for(int j=0;j<item_List.size();j++){
                                         item += (String)item_List.get(j)+"<br>";
                                     }                                     
                                  }
                                  /*98.01.09 fix 查核人員.直接key in
                                  for(int j=0;j<name_List.size();j++){                            
                                     if(name.indexOf((String)((DataObject)name_List.get(j)).getValue("muser_id")) != -1){
                                        chinese_name += (String)((DataObject)name_List.get(j)).getValue("muser_name");
                                        if(j < name_List.size() -1) chinese_name +="<br>";
                                     }     
                                  }
                                  */
						  %>					  
						  
						  
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>" width="835">                           
                             <td class="sbody"  align=center height="38" width="32"><%=m_year%>&nbsp;</td>
                             <td class="sbody"  align=center height="19" width="131"><a href=FX011W.jsp?act=Edit&m_year=<%=m_year%>&tbank_no=<%=bank_no%>&examine=<%=examine%>><%=examine%><br><%=bank_name%></a>&nbsp;</td>
                             <td class="sbody"  align=center height="19" width="70"><%=check_date%>&nbsp;</td>
                             <td class="sbody"  align=center height="19" width="125"><%=name%></td>
                             <td class="sbody"  align=center height="19" width="120"><%=item%>&nbsp;</td>
                             <td class="sbody"  align=center height="19" width="152"><%=result%>&nbsp;</td>
                             <td class="sbody"  align=center height="16" width="181"><%=content%>&nbsp;</td>
                             <td class="sbody"  align=center height="16" width="181"><%=remark%>&nbsp;</td>                          
                          </tr>
                          <%  }//end of for 
                           }%>
					</table>                 
                      </td>
                    </tr>




                    <td width="815" height="56"><div align="right"><div align="right">
                    <div align="center"><jsp:include page="getMaintainUser.jsp?width=815" flush="true" /></div>
                    </div>
                    <p align="left">　</div></td>
      </table>
                </div>
              </td>
  </tr>
  <tr>
                <td width="825" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr>
                      <td colspan="2" width="583"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明     
                        : </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td class="sbody" width="561">
                        <ul>
                          <li class="sbody">點選<FONT color=#666666>【</FONT>新增】按鈕可新增「辦理基層金融機構變現性資產查核」之資料。</li>
                          <li class="sbody">點選所列之[受檢單位]可變更該申報之資料。 </li>                         
                        </ul>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>

</table>
</form>
