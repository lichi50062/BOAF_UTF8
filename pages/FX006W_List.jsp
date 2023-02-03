<%
//94.10.17 first design by 4180
//94.12.13 add out_begin_date,out_end_date
//96.11.12 add 增加委外項目.委外範圍 by 2295
//97.07.09 fix 結束日期.可不輸入 by 2295
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
	List WLX06_M_OUTPUSH= (request.getAttribute("WLX06_M_OUTPUSH")==null)?null:(List)request.getAttribute("WLX06_M_OUTPUSH");
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX006W")==null ) ? new Properties() : (Properties)session.getAttribute("FX006W");
    if (permission == null) {
        System.out.println("FX006W_List.permission == null");
    } else {
        System.out.println("FX006W_List.permission.size ="+permission.size());
    }
    String seq_no="";
    String companyname="";
    String contractname="";
    String contracttel="";
    String complainname="";
    String complaintel="";
    String comment="";
    String maintain_id="" ;
    String maintain_name="" ;
    String maintain_date="" ;
    String out_beg_date="";
    String out_end_date="";
    String out_item="";
    String out_range="";
    String countdata="";
%>
<script language="javascript" src="js/FX006W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>各農漁會委外催收委外之對象維護</title>
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
<table width="835" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
            <td width="835"><img src="images/space_1.gif" width="12" height="12"></td>
            </tr>
        <tr>
          <td bgcolor="#FFFFFF" width="835">
        
              <tr>
                <td width="835" height="18">
                  <table width="835" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17" align="left"></td>
                      <td width="365" align="center"><b><font size="4">各農漁會委外內部作業資料申報</font></b></td>
                      <td width="200" align="right"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="835" height="50">
                <div align="left">
                <table width="835" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="835" >               
                      </td>
                    </tr>
                    <tr>                       
                    <div><jsp:include page="getLoginUser.jsp?width=835" flush="true" /></div>			 							
			 		               
                    </tr>
                    
        <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr ><td width="835" align="right" valign="bottom">
          <table border="1" cellspacing="1" bordercolor="#3A9D99" width="835"  class="sbody" cellpadding="0" height="20">
            <tr>
            <form name="date" method="post">
                 <td class="sbody" bgcolor="#E7E7E7" width="835"  valign="middle" height="20">
                  <a href="/pages/FX006W.jsp?act=new" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                  <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                  </td>
            </form>
            </tr>
         </table>
         </td>
         </tr>
         <%}%>

                    <tr>
                      <td  align="center" width="835" valign="top">
					  <table class="sbody" width="835" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="1" >
                         <tr bgcolor="#9AD3D0">                          
                          <td width="30" rowspan="2" align="center" height="1">委外項目</td>
                          <td width="439" colspan="5"  align="center" height="1">受委託機構</td>
                          <td width="131" rowspan="2" align="center" height="2">&nbsp;委外事項範圍</td>
                          <td width="116" colspan="2"  align="center" height="1">信用部申訴窗口</td>
                          <td width="43" rowspan="2"  align="center" height="44">備註</td>
                         </tr>
                         <tr class="sbody" bgcolor="#9AD3D0">
                          <td width="131" class="sbody" align="center" height="1">機構名稱</td>
                          <td width="49" class="sbody"   align="center" height="1">聯絡人</td>
                          <td width="93" class="sbody"   align="center" height="1">聯絡電話</td>
                          <td width="71" align="center" height="1">起始日期</td>
                          <td width="60" align="center" height="1">結束日期</td>                       
                          <td width="43" class="sbody"    align="center" height="1">聯絡人</td>
                          <td width="68" class="sbody"    align="center" height="1">專線電話</td>                      
                        </tr>
                        
                        
                        
                        
						<% if(WLX06_M_OUTPUSH.size()==0){%>
                           <tr class="sbody" bgcolor="#D8EFEE"><td colspan="12" align="center"  height="19" width="835" >
                           <font class="sbody">尚無資料</font></td></tr>
						   <%}else{
						           String pre_out_item="";
						           String[] out_item_name = {"","一","二","三","四","五","六","七","八","九","十",
						                       			    "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十"};
								   for(int i=0;i<WLX06_M_OUTPUSH.size();i++){
								       seq_no="";companyname="";contractname="";contracttel="";complainname="";complaintel="";
    								   comment="";maintain_id="";maintain_name="";maintain_date="";out_beg_date="";out_end_date="";
    								   out_item="";out_range="";													     
									   seq_no = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("seq_no")); 
           							   companyname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("outcompanyname"));
									   contractname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("outcontractname"));
									   contracttel = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("outcontracttel"));														
									   complainname = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("bankcomplainname"));
									   complaintel = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("bankcomplaintel"));
									   comment = String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("outcomment"));													
									   maintain_id=String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("user_id"));
									   maintain_name=String.valueOf(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("user_name"));
									   maintain_date=Utility.getCHTdate((((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   													
        							   out_beg_date = Utility.getCHTdate((((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_begin_date")).toString().substring(0, 10), 0);   
        							   if(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_end_date") != null){
        							      out_end_date = "";
        							      out_end_date = Utility.getCHTdate((((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_end_date")).toString().substring(0, 10), 0);   
        							   }
        							   //96.11.09 add 委外項外.委外範圍
	   								   out_item = (((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_item")==null)?"": (String)((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_item");	
	   							       out_range = (((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_range")==null)?"": (String)((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("out_range");	
	   							       countdata = ((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("countdata") == null ?"":(((DataObject)WLX06_M_OUTPUSH.get(i)).getValue("countdata")).toString();
	   							       System.out.println("i="+i+":pre_out_item="+pre_out_item+":out_item="+out_item);
	   							       
						  %>					  
						  
						  
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>" width="835">
                            <!--td class="sbody"  align=center height="38" width="44" rowspan="2"><%out.print(String.valueOf(i+1));%></td-->
                            <%if(out_item.equals("")){%>
                            <td class="sbody"  align=center height="38" width="30" >&nbsp;</td>
                            <%}else{%>
                            <%if(i==0 || (!pre_out_item.equals(out_item))){%>
                            <td class="sbody"  align=center height="38" width="30" rowspan="<%=countdata%>"><%=out_item_name[Integer.parseInt(out_item)]%>&nbsp;</td>
                            <%}else{%>
                            <%}
                              }%>
                            <td class="sbody"  align=center height="19" width="131">
                                <u><a href=FX006W.jsp?act=Edit&seq_no=<%=seq_no%>><%=companyname%></a></u>
                            </td>
                            <td class="sbody"  align=center height="19" width="49"><%=contractname%></td>
                            <td class="sbody"  align=center height="19" width="93"><%=contracttel%></td>
                            <td class="sbody"  align=center height="19" width="55"><%=out_beg_date%></td>
                            <td class="sbody"  align=center height="19" width="55"><%=out_end_date%>&nbsp;</td>
                            <td class="sbody"  align=center height="19" width="131" ><%=out_range%>&nbsp;</td>
                            <td class="sbody"  align=center height="19" width="43" ><%=complainname%></td>
                            <td class="sbody"  align=center height="19" width="68"><%=complaintel%></td>
                            <td class="sbody"  align=center height="16" width="43">
                               <%out.print((comment.equals("null"))?"-":comment);%>
                            </td>
                          
                          </tr>
                          <%    if(!pre_out_item.equals(out_item))pre_out_item = out_item;
                                }//end of for 
                           }%>
                         </table>                 
                      </td>
                    </tr>


                    <td width="835" height="56"><div align="right"><div align="right">
                    <div align="center"><jsp:include page="getMaintainUser.jsp?width=835" flush="true" /></div>
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
                          <li class="sbody">點選<FONT color=#666666>【</FONT>新增】按鈕可新增「委外催收委外之對象」之資料。</li>
                          <li class="sbody">點選所列之[機構名稱]可變更該申報之資料。 </li>
                        </ul>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>

</table>
</form>
