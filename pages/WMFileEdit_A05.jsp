<%
//93.12.17 add 權限檢核 by 2295
//94.02.14 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份 by 2295
//94.03.03 fix 金額用BigDecimal來處理,再執行相乘 by 2295
//94.03.11 fix 910500=920101+920201+920301+920401+920501+920601+920710+920720+920730+920740+920750+920801 by 2295
//94.03.17 fix 91060P=910400/910500取到小數第二位(第三位四捨五入) by 2295
//94.03.17 fix 累加風險性資產額至910500 by 2295
//94.03.24 fix text field靠右 by 2295
//94.07.15 add 920710/720/730/740/750改變底色 by 2295		
//94.07.20 add 920710.920720.920730.920740.920750 在編輯時,先將風險權數一併算好 by 2295
//95.05.11 add 910500=920101+920201+920301+920401+920501+920601+920710+920720+920730+920740+920750+920801-910203
//			   減910203(2_1)特定損失所提列之備抵呆帳、損失準備及營業準備		 by 2295
//95.05.15 fix 修改新增申報項目sql by 2295 
//96.11.30 add 910204的計算公式 by 2295
//99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>

<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.lang.*" %>
<%@ page import="java.math.BigDecimal" %>

<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	System.out.println("S_MONTH="+S_MONTH);
	//String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	//System.out.println("Report="+Report);
	List data_div01 = null;
	List paramList = new ArrayList();
	StringBuffer sql = new StringBuffer();
	String ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
	if(act.equals("new")){
	    sql.append("select "+ncacno+".acc_code,"+ncacno+".acc_name,a05_assumed.assumed from "+ncacno+" left join a05_assumed on "+ncacno+".acc_code=a05_assumed.acc_code where acc_tr_type = ? and acc_div=? order by acc_range");
	    paramList.add("A05");
	    paramList.add("07");
	    data_div01 = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"assumed,update_date");  	    	    
	}else{
		data_div01 = (List)request.getAttribute("data_div01");
	}
	System.out.println("data_div01.size="+data_div01.size());
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A05.permission == null");
    }else{
       System.out.println("WMFileEdit_A05.permission.size ="+permission.size());
               
    }
%>
<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
</style>
<%if(act.equals("Query")){%>
<title>申報資料查詢</title>
<%}else{%>
<title>線上編輯申報資料</title>
<%}%>

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
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<script language="JavaScript" src="js/menu.js"></script>
<script language="JavaScript" src="js/menucontext_A05.js"></script>
<script language="JavaScript">
showToolbar();
</script>
<script language="JavaScript">
function UpdateIt(){
if (document.all){
document.all["MainTable"].style.top = document.body.scrollTop;
setTimeout("UpdateIt()", 200);
}
}
UpdateIt();
</script>
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
 <input type="hidden" name="act" value="">  
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
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A05")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
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
                      <td><Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>A05&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A05")%></td>
                          </tr>  
                          <%}%>
                          <tr class="sbody" bgcolor='#D2F0FF'> 
                            <td width="112"> <div align=left>基準日</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" <%if(act.equals("Edit")) out.print("disabled");%> size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH <%if(act.equals("Edit")) out.print("disabled");%>>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
                            </td>
                          </tr>
                          </table>
                          <Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr bgcolor='e7e7e7'>		
							<td colspan=5><b><div align=center><a name="A05_div01">信用部淨值占風險性資產比率計算表</a></div></b></td>		
						  </tr>
						  <tr bgcolor='e7e7e7' class="sbody"> 
                            <td width=55 bgcolor="#B1DEDC"> <div align=left>項目代碼</div></td>
                            <td width=* bgcolor="#B1DEDC"> <div align=left>項目名稱</div></td>
                            <td width=55 bgcolor="#B1DEDC"> <div align=left>風險權數</div></td>
                            <td width=120 bgcolor="#B1DEDC"> <div align=left>帳面金額</div></td>
                            <td width=80 bgcolor="#B1DEDC"> <div align=left>風險性資產額</div></td>
                          </tr>
                          
 <% 	int i = 0 ;
		boolean fontbold=false;
		String bgcolor="#F2F2F2";
		String amtName="";
		//String fontsize="2";		
		String tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
		System.out.println("tmpAcc_Code="+tmpAcc_Code+"tmpAcc_Code.substring(2,2)="+tmpAcc_Code.substring(1,2));
		while(  i < data_div01.size() && tmpAcc_Code.substring(1,2).equals("2")){  //信用部淨值占風險性資產比率計算表
			bgcolor="#F2F2F2";
			fontbold=false;
			//fontsize="2";
			System.out.println("tmpAcc_Code="+tmpAcc_Code+"tmpAcc_Code.substring(2,2)="+tmpAcc_Code.substring(1,2));
			tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
		    String tmpamt = ((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString();
		    //System.out.println("amt="+tmpamt);
	    	String tmpassumed = ((((DataObject)data_div01.get(i)).getValue("assumed")) == null ? "":(((DataObject)data_div01.get(i)).getValue("assumed"))).toString();
	    	//System.out.println("(((DataObject)data_div01.get(i)).getValue("assumed"))="+(((DataObject)data_div01.get(i)).getValue("assumed")));
	    	
	    	//94.07.20 add 920710.920720.920730.920740.920750 在編輯時,先將風險權數一併算好==========================
	    	if(tmpAcc_Code.equals("920710") || tmpAcc_Code.equals("920720") || tmpAcc_Code.equals("920730") || tmpAcc_Code.equals("920740") || tmpAcc_Code.equals("920750") ){
		       tmpamt=((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString();
		       tmpassumed = ((((DataObject)data_div01.get(i+2)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i+2)).getValue("amt"))).toString();
		       System.out.println(tmpAcc_Code+"="+tmpamt+"*"+tmpassumed);
		    } 
		    //=======================================================================================================   
	    	System.out.println("tmpassumed="+tmpassumed);
		    String tmpassumed_amt="";		    
		    
		    if(tmpamt.length()==0){
		       tmpassumed_amt="";
		    }else if(tmpassumed.length()==0){
		       tmpassumed_amt="0";	
		    }else{
		       //94.03.03金額用BigDecimal來處理,再執行相乘 	    
		       BigDecimal a1 = new BigDecimal(Double.parseDouble(tmpamt)); 		       
		       //94.07.20 add 920710.920720.920730.920740.920750 在編輯時,先將風險權數一併算好======================
		       if(tmpAcc_Code.equals("920710") || tmpAcc_Code.equals("920720") || tmpAcc_Code.equals("920730") || tmpAcc_Code.equals("920740") || tmpAcc_Code.equals("920750") ){
		          tmpassumed = String.valueOf((Double.parseDouble(tmpassumed)/1000));
		          System.out.println(tmpAcc_Code+"="+tmpamt+"*"+tmpassumed);
		       }		       
		       //==============================================================================================
		       BigDecimal a2 = new BigDecimal(tmpassumed); 		       
		       System.out.println("a2="+a2);
		       System.out.println("nowBigDecimal="+a2.multiply(a1));
		       
		       //tmpassumed_amt = String.valueOf(Long.parseLong(tmpamt)*Float.parseFloat(tmpassumed));
		       //double tmpassumed_amt_d = a2.multiply(a1).toString();
		       //System.out.println("tmpassumed_amt_f="+a2.multiply(a1).toString());
		       //System.out.println("*1000="+ Math.round(a2.multiply(a1).doubleValue() * 1000));
		       //double tmpassumed_amt_d = Math.round(a2.multiply(a1).doubleValue() * 1000)/1000;	
		       //System.out.println("tmpassumed_amt_d="+String.valueOf(tmpassumed_amt_d));
		       //tmpassumed_amt = (new BigDecimal(tmpassumed_amt_d)).toString();		       
		       //System.out.println("tmpassumed_amt="+tmpassumed_amt);
		       
		       tmpassumed_amt = a2.multiply(a1).toString();
		       if(tmpAcc_Code.equals("920710") || tmpAcc_Code.equals("920720") || tmpAcc_Code.equals("920730") || tmpAcc_Code.equals("920740") || tmpAcc_Code.equals("920750") ){
		          tmpassumed = "";
		       }
		       if(tmpassumed_amt.indexOf(".") != -1){
		          try{		              
		              tmpassumed_amt = tmpassumed_amt.substring(0,tmpassumed_amt.indexOf(".")+4);     
		          }catch(Exception e){}
		       }
		    }   
		    
			
			//信用部淨值占風險性資產比率計算表
			//94.07.15 add 920710/720/730/740/750改變底色		
			if(tmpAcc_Code.equals("920710")/*資產(一)之帳面金額 */   
			|| tmpAcc_Code.equals("920720")/*資產(二)之帳面金額 */   
			|| tmpAcc_Code.equals("920730")/*資產(三)之帳面金額 */   
			|| tmpAcc_Code.equals("920740")/*資產(四)之帳面金額 */   
			|| tmpAcc_Code.equals("920750")/*資產(五)之帳面金額 */   
			|| tmpAcc_Code.equals("929901")/*信用部淨值占風險性資產比率計算表中合計的帳面金額  */   
			){ 			
				fontbold=true;
				bgcolor = "#FFFFE6";			 
				//fontsize="4";
			}
	%>	
			<tr bgcolor='<%=bgcolor%>' class="sbody">
	
			<td bgcolor="<%=bgcolor%>">			
			<%if(fontbold){%><b><%}%>				
			<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
			<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">		
			<input type=hidden name=acc_div value="01">
			
			</td>
			<td>
			
			<%if(fontbold){%><b><%}%>	

			<div align=left><%if((((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
							  || (((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
			<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>		
			</div>
			<%if(fontbold){%></b><%}%>		
			
			</td>
			
			<td><a name="<%=tmpAcc_Code%>"><%=tmpassumed%>
			<input type='hidden' name='assumed' value="<%=tmpassumed%>">
			</a>
			
			</td>		

			<td>
		
			<%if(tmpAcc_Code.indexOf("N")!=-1){%>
			<input type='text' name='amt' value="<%=(String)((DataObject)data_div01.get(i)).getValue("amt_name")==null ? "" : (String)((DataObject)data_div01.get(i)).getValue("amt_name")%>" size=16 maxlength=16 style='text-align: right;'>
			<%}else{ 
				if (tmpAcc_Code.indexOf("P")!=-1){				   
			       if( tmpamt == null || tmpamt != null && (tmpamt.equals("0")) ){%>
			           <input type='text' name='amt' value="" <%if(tmpAcc_Code.equals("91060P")) out.print(" readonly");%> size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
			       <% }else{ %>
			           <input type='text' name='amt' value="<%=Utility.getPercentNumber(tmpamt)%>" <%if(tmpAcc_Code.equals("91060P")) out.print(" readonly");%> size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
			       <% } %>
			<%	//modify by 2354 12.23
				}else{
					if(tmpAcc_Code.equals("910500")){
					   if( tmpamt == null || tmpamt != null && (tmpamt.equals("0")) ){%>
					       <input type='text' name='amt' value="" size=16 maxlength=16 readonly onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
					   <% }else{ %>
			               <input type='text' name='amt' value="<%=Utility.setCommaFormat(tmpamt)%>" size=16 maxlength=16 readonly onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
			           <% } %>
			<%		}else{ 
			           System.out.println("tmpAcc_Code="+tmpAcc_Code); 			           
			           if( tmpamt == null || tmpamt != null && (tmpamt.equals("0")) ){%>			              
	                       <input type='text' name='amt' value="" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>					               
			           <% }else{%>     
			              <input type='text' name='amt' value="<%=Utility.setCommaFormat(tmpamt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>			          
			           <%}//end of tmpamt != "0"%>			
			<%  	}
				}//end of !='910500'
			  }//end of "P"
			%>
			</a>
				
			</td>		
			<td><div id='div_assumed_amt'><%=Utility.setCommaFormat(tmpassumed_amt)%></div>
		
			<input type='hidden' name='assumed_amt' value="<%=Utility.setCommaFormat(tmpassumed_amt)%>">
			</a>
			
			</td>		
		</tr>
	
		
	<%
		    i++;	
		    tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();		  
		}
	%>
	
			</table>
			<Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
			<tr bgcolor='e7e7e7' class="sbody">	
				<td colspan=3><b><div align=center><font size="4"><a name="A05_div02">信用部淨值占風險性資產比率</a></font></div></b></td>		
			</tr>
			<tr bgcolor='e7e7e7' class="sbody"> 
                <td width=111 bgcolor="#B1DEDC"> <div align=left>項目代碼</div></td>
                <td width=294 bgcolor="#B1DEDC"> <div align=left>項目名稱</div></td>
                <td width=177 bgcolor="#B1DEDC"> <div align=left>項目數值</div></td>
            </tr>
<%  	fontbold=false;
		//fontsize="2";
		bgcolor="#F2F2F2";		
		while( i < data_div01.size()){
			fontbold=false;
			//fontsize="2";
			bgcolor="#F2F2F2";
		    tmpAcc_Code = ((String)((DataObject)data_div01.get(i)).getValue("acc_code")).trim();
		    
			//信用部淨值占風險性資產比率計算表
			if(tmpAcc_Code.equals("910199")/*第一類資本合計(A)*/      
			|| tmpAcc_Code.equals("910299")/*第二類資本合計(B)*/          
			|| tmpAcc_Code.equals("910300")/*淨值總額(C)=(A)+(B)*/      
			|| tmpAcc_Code.equals("910400")/*合格淨值(G)=(C) -  [ (D)+(E)+(F) ]*/      
			|| tmpAcc_Code.equals("910500")/*風險性資產總額(H)*/      
			|| tmpAcc_Code.equals("91060P")/*信用部淨值占風險性資產比率（資本適足率）*/   			
			){ 			
				fontbold=true;
				bgcolor = "#FFFFE6";			 
				//fontsize="4";
			}
	%>	
	<tr bgcolor='<%=bgcolor%>' class="sbody">
	
		<td bgcolor="<%=bgcolor%>">			
		<%if(fontbold){%><b><%}%>		

		<div align=left><%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%></div>
		<input type=hidden name=acc_code value="<%=(String)((DataObject)data_div01.get(i)).getValue("acc_code")%>">
		<input type=hidden name=acc_div value="01">
		
		</td>
		<td>
		
		<%if(fontbold){%><b><%}%>	
		
		<div align=left><%if((((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--合計") != -1) 
						  || (((String)((DataObject)data_div01.get(i)).getValue("acc_name")).indexOf("--小計") != -1)){%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%}%>		
		<%=(String)((DataObject)data_div01.get(i)).getValue("acc_name")%>		
		</div>
		<%if(fontbold){%></b><%}%>		
		
		</td>
		
		<td><a name="<%=tmpAcc_Code%>">	
		<%
		//modify by 2354 12.23
		
		
		%>
	    <%if (tmpAcc_Code.indexOf("P")==-1){	        	        
			if(tmpAcc_Code.equals("910500")){
			   if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){%>
			       <input type='text' name='amt' readonly value="" size=16 maxlength=16 style='text-align: right;'>
			   <% }else{ %>	
		           <input type='text' name='amt' readonly value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'>
		       <% } %>
		<%	}else{//end of Acc_Code =='910500'
		      if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){		           		      	  
		      %>
			      <%if(tmpAcc_Code.equals("920101") || tmpAcc_Code.equals("920201") || tmpAcc_Code.equals("920301") || tmpAcc_Code.equals("920401") || tmpAcc_Code.equals("920501") || tmpAcc_Code.equals("920601")
				  || tmpAcc_Code.equals("920710") || tmpAcc_Code.equals("920720") || tmpAcc_Code.equals("920730") || tmpAcc_Code.equals("920740") || tmpAcc_Code.equals("920750") || tmpAcc_Code.equals("920801") || tmpAcc_Code.equals("910400") || tmpAcc_Code.equals("910203")){				  
				  
				  %>	                       
				   <input type='text' name='amt' value="" size=16 maxlength=16  onchange='changeStr_A05(this,<%=i%>,this.form)' onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'>
	               <%}else{	              
	               	%>			           		               
		           <input type='text' name='amt' value="" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='this.value=changeStr(this)' style='text-align: right;'>
		           <%}//end of addText%> 
		      <% }else{ 		      	
		      	%>
		          <input type='text' name='amt' value="<%=Utility.setCommaFormat(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
		      <% } %>
		<%	}//end of Acc_Code !='910500'
		}else{			
		     if( ((DataObject)data_div01.get(i)).getValue("amt") == null ||  (((DataObject)data_div01.get(i)).getValue("amt") != null && ((((DataObject)data_div01.get(i)).getValue("amt")).toString()).equals("0")) ){%>
		          <input type='text' name='amt' value="" <%if(tmpAcc_Code.equals("91060P")) out.print(" readonly");%> size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
		     <% }else{ %>
		          <input type='text' name='amt' value="<%=Utility.getPercentNumber(((((DataObject)data_div01.get(i)).getValue("amt")) == null ? "":(((DataObject)data_div01.get(i)).getValue("amt"))).toString())%>" <%if(tmpAcc_Code.equals("91060P")) out.print(" readonly");%> size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='changeStr_A05(this,<%=i%>,this.form)' style='text-align: right;'>
		     <% } %>
		<%}%>
		

		<a>	
		
		</td>		
	</font>	
	</tr>
	
		
	<%	    i++;	
		}
	%>
	
	          </Table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>              
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A05','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A05','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A05','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
                <%}%>        
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image81" width="80" height="25" border="0" id="Image81"></a></div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A05")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                          <li><font color='red'>【910500】風險性資產總額(H),已扣除【910203】(2_1)特定損失所提列之備抵呆帳、損失準備及營業準備之金額。</font></li>
                          <li><font color='red'>【929901】係信用部淨值佔風險性資產比率計算表中合計的帳面金額,須等於下列科目帳面金額的總計(【920101】+【920201】+【920301】+【920401】+【920501】+【920601】+【920710】+【920720】+【920730】+【920740】+【920750】+【920801】)。</font></li>
                          <li><font color='red'>資本適足率第二類資本中的「910204調整後備抵呆帳、損失準備及營業準備」不得超過「910500修正後之風險性資產總額(H) 」×1.25%
												，因此「 910204調整後備抵呆帳、損失準備及營業準備」最高為「910500修正後之風險性資產總額(H) 」×1.25%。</font></li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
</body>

</html>
