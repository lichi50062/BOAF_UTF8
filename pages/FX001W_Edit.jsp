<%
// 93.12.20 add 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
//          add 權限設定 by 2295
// 94.01.12 fix SETUP_NO(原始核准設立文號) 40->120
// 			   CHG_LICENSE_NO(最近換照文號)	40->120
// 			   ADDR(地址) 40 ->120
// 94.01.14 fix 地方主管機關代號.是否參加電腦共用中心.參加的電腦共用中心機構代碼.
// 			   本機構員工總人數不可為空白 by 2295
// 94.02.14 fix 把每月申報資料先disable by 2295
// 94.04.01 add 同一職務不可有一人以上擔任 by 2295
// 94.04.06 fix 高階主管卸任基本資料.簡易清單.加上卸任日期 by 2295
// 95.05.26 add 從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數 by 2295
// 95.08.21 fix 顯示日期用getStringTokenizerData來拆日期字串 by 2295
// 99.12.08 fix cd01縣市別.區分100年度/99年度 by 2808                                          
//100.01.26 fix cd02區域別.區分100年度/99年度 by 2295
//              地方主管機關代號/參加的電腦共用中心機構代碼.區分100年度/99年度 by 2295
//101.06.13 fix 增加機構英文欄位、郵遞區號 by2968
//102.11.22 fix 修正當機構無資料時,無法顯示網頁 by 2295
//103.01.06 fix 原稽核人員拿掉,移至可新增多筆稽核人員
//104.06.24 add 人員配置情形 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>

<%	
    
	//String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
	//bank_no = ( request.getParameter("bank_code")==null ) ? bank_no : (String)request.getParameter("bank_code");			
	//fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================	
	//取得FX001W的權限
	Properties permission = ( session.getAttribute("FX001W")==null ) ? new Properties() : (Properties)session.getAttribute("FX001W"); 
	if(permission == null){
       System.out.println("FX001W_Edit.permission == null");
    }else{
       System.out.println("FX001W_Edit.permission.size ="+permission.size());
               
    }		    
	List WLX01 = (List)request.getAttribute("WLX01");
	List WLX01_M = (List)request.getAttribute("WLX01_M");	
	List WLX01_WM = (List)request.getAttribute("WLX01_WM");	
	List WLX01_Count = (List)request.getAttribute("WLX01_Count");	
	List WLX01_Audit = (List)request.getAttribute("WLX01_Audit");
	//94.04.06 
	String haveCancel="N";	
	if(WLX01 != null && WLX01.size() != 0 &&  ((DataObject)WLX01.get(0)).getValue("cancel_no") != null){
	   haveCancel = (String)((DataObject)WLX01.get(0)).getValue("cancel_no");
	}   

	if(WLX01 != null){
	   System.out.println("WLX01.size()="+WLX01.size());	
	}else{
	   System.out.println("WLX01 == null");
	}
	if(WLX01_M != null){
		System.out.println("WLX01_M.size()="+WLX01_M.size());	
	}else{
	   System.out.println("WLX01_M == null");
	}
	if(WLX01_WM != null){
		System.out.println("WLX01_WM.size()="+WLX01_WM.size());	
	}else{
	   System.out.println("WLX01_WM == null");
	}
	if(WLX01_Count != null){
		System.out.println("WLX01_Count.size()="+WLX01_Count.size());	
	}else{
	   System.out.println("WLX01_Count == null");
	}
	if(WLX01_Audit != null){
		System.out.println("WLX01_Audit.size()="+WLX01_Audit.size());	
	}else{
	   System.out.println("WLX01_Audit == null");
	}
	String HSIEN_ID_AREA_ID = "";
	if(WLX01 != null && WLX01.size() != 0){
	     HSIEN_ID_AREA_ID =((((DataObject)WLX01.get(0)).getValue("hsien_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("hsien_id"))	   
						  +"/"+ ((((DataObject)WLX01.get(0)).getValue("area_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("area_id"));	   
	}	
	
	
	String IT_HSIEN_ID_AREA_ID = "";
	if(WLX01 != null && WLX01.size() != 0){
	     IT_HSIEN_ID_AREA_ID =((((DataObject)WLX01.get(0)).getValue("it_hsien_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("it_hsien_id"))	   
						  +"/"+ ((((DataObject)WLX01.get(0)).getValue("it_area_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("it_area_id"));	   
	}
	
	String AUDIT_HSIEN_ID_AREA_ID = "";
	if(WLX01 != null && WLX01.size() != 0){
	     AUDIT_HSIEN_ID_AREA_ID =((((DataObject)WLX01.get(0)).getValue("audit_hsien_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("audit_hsien_id"))	   
						  +"/"+ ((((DataObject)WLX01.get(0)).getValue("audit_area_id") == null) ?"":(String)((DataObject)WLX01.get(0)).getValue("audit_area_id"));	   
	}
		
	String SETUP_DATE_Y="";
	String SETUP_DATE_M="";
	String SETUP_DATE_D="";
	String SETUP_DATE="";
	
	String CHG_LICENSE_DATE_Y="";
	String CHG_LICENSE_DATE_M="";
	String CHG_LICENSE_DATE_D="";
	String CHG_LICENSE_DATE="";
	
	String START_DATE_Y="";
	String START_DATE_M="";
	String START_DATE_D="";
	String START_DATE="";
	
	String OPEN_DATE_Y="";
	String OPEN_DATE_M="";
	String OPEN_DATE_D="";
	String OPEN_DATE="";
	
	String CANCEL_DATE_Y="";
	String CANCEL_DATE_M="";
	String CANCEL_DATE_D="";
	String CANCEL_DATE="";
	String area_id = "";
	String it_area_id = "";
	String audit_area_id = "";
	List tmpDate=null;
	if(WLX01.size() != 0){
   	    int i = 0;
		if(((DataObject)WLX01.get(0)).getValue("setup_date") != null){		   
		   SETUP_DATE = Utility.getCHTdate((((DataObject)WLX01.get(0)).getValue("setup_date")).toString().substring(0, 10), 0);
		   System.out.println("SETUP_DATE="+SETUP_DATE);		  
		   /*
		   i = 0;
		   if(SETUP_DATE.length() == 9) i = 1; 
		   SETUP_DATE_Y = SETUP_DATE.substring(0,2+i);		   
		   SETUP_DATE_M = SETUP_DATE.substring(3+i,5+i);		   
		   SETUP_DATE_D = SETUP_DATE.substring(6+i,SETUP_DATE.length());		 
		   */		  
		   //95.08.21 顯示日期用getStringTokenizerData來拆日期字串
		   tmpDate = Utility.getStringTokenizerData(SETUP_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		      SETUP_DATE_Y = (String)tmpDate.get(0);
		      SETUP_DATE_M = (String)tmpDate.get(1);
		      SETUP_DATE_D = (String)tmpDate.get(2);
		      System.out.println("SETUP_DATE_Y="+SETUP_DATE_Y);
		      System.out.println("SETUP_DATE_M="+SETUP_DATE_M);
		      System.out.println("SETUP_DATE_D="+SETUP_DATE_D);
		   }
		}
		if(((DataObject)WLX01.get(0)).getValue("chg_license_date") != null){
		   CHG_LICENSE_DATE = Utility.getCHTdate((((DataObject)WLX01.get(0)).getValue("chg_license_date")).toString().substring(0, 10), 0);
		   i = 0;
		   System.out.println(CHG_LICENSE_DATE);		   
		   tmpDate = Utility.getStringTokenizerData(CHG_LICENSE_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		      CHG_LICENSE_DATE_Y = (String)tmpDate.get(0);
		      CHG_LICENSE_DATE_M = (String)tmpDate.get(1);
		      CHG_LICENSE_DATE_D = (String)tmpDate.get(2);		      
		   }
		   /*
		   if(CHG_LICENSE_DATE.length() == 9) i = 1; 
		   CHG_LICENSE_DATE_Y = CHG_LICENSE_DATE.substring(0,2+i);
		   CHG_LICENSE_DATE_M = CHG_LICENSE_DATE.substring(3+i,5+i);
		   CHG_LICENSE_DATE_D = CHG_LICENSE_DATE.substring(6+i,CHG_LICENSE_DATE.length());*/		 
		}
		if(((DataObject)WLX01.get(0)).getValue("start_date") != null){
		   START_DATE = Utility.getCHTdate((((DataObject)WLX01.get(0)).getValue("start_date")).toString().substring(0, 10), 0);
		   tmpDate = Utility.getStringTokenizerData(START_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		      START_DATE_Y = (String)tmpDate.get(0);
		      START_DATE_M = (String)tmpDate.get(1);
		      START_DATE_D = (String)tmpDate.get(2);		      
		   }
		   /*i = 0;
		   if(START_DATE.length() == 9) i = 1; 
		   START_DATE_Y = START_DATE.substring(0,2+i);
		   START_DATE_M = START_DATE.substring(3+i,5+i);
		   START_DATE_D = START_DATE.substring(6+i,START_DATE.length());*/		 
		}
		if(((DataObject)WLX01.get(0)).getValue("open_date") != null){
		   OPEN_DATE = Utility.getCHTdate((((DataObject)WLX01.get(0)).getValue("open_date")).toString().substring(0, 10), 0);
		   tmpDate = Utility.getStringTokenizerData(OPEN_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		      OPEN_DATE_Y = (String)tmpDate.get(0);
		      OPEN_DATE_M = (String)tmpDate.get(1);
		      OPEN_DATE_D = (String)tmpDate.get(2);		      
		   }
		   /*i = 0;
		   if(OPEN_DATE.length() == 9) i = 1; 
		   OPEN_DATE_Y = OPEN_DATE.substring(0,2+i);
		   OPEN_DATE_M = OPEN_DATE.substring(3+i,5+i);
		   OPEN_DATE_D = OPEN_DATE.substring(6+i,OPEN_DATE.length());*/		 
		}
		
		if(((DataObject)WLX01.get(0)).getValue("cancel_date") != null){
		   CANCEL_DATE = Utility.getCHTdate((((DataObject)WLX01.get(0)).getValue("cancel_date")).toString().substring(0, 10), 0);
		   System.out.println("CANCEL_DATE="+CANCEL_DATE);
		   tmpDate = Utility.getStringTokenizerData(CANCEL_DATE,"//");
		   if(tmpDate != null && tmpDate.size() != 0){
		      CANCEL_DATE_Y = (String)tmpDate.get(0);
		      CANCEL_DATE_M = (String)tmpDate.get(1);
		      CANCEL_DATE_D = (String)tmpDate.get(2);		      
		   }
		   /*
		   i = 0;
		   if(CANCEL_DATE.length() == 9) i = 1;
		   CANCEL_DATE_Y = CANCEL_DATE.substring(0,2+i);
		   CANCEL_DATE_M = CANCEL_DATE.substring(3+i,5+i);
		   CANCEL_DATE_D = CANCEL_DATE.substring(6+i,CANCEL_DATE.length());		 
		  */
		}
	    area_id = (((DataObject)WLX01.get(0)).getValue("area_id")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("area_id");
	    it_area_id = (((DataObject)WLX01.get(0)).getValue("it_area_id")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("it_area_id");
	    audit_area_id = (((DataObject)WLX01.get(0)).getValue("audit_area_id")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("audit_area_id");
	}
	String cd01Table = Integer.parseInt(Utility.getYear())> 99 ?"cd01" : "cd01_99" ;
	String cd02Table = Integer.parseInt(Utility.getYear())> 99 ?"cd02" : "cd02_99" ;
	String sqlcmd = "Select CD01.HSIEN_ID,CD01.HSIEN_NAME,CD02.AREA_ID,CD02.AREA_NAME From "+cd01Table+" CD01, "+cd02Table+" CD02 "+
					 " Where  CD01.HSIEN_ID=CD02.HSIEN_ID Order by CD01.HSIEN_ID, CD02.AREA_ID ";
	List hsien_id_area_id = DBManager.QueryDB_SQLParam(sqlcmd,null,"");
	String wlx01_m_year = (Integer.parseInt(Utility.getYear()) < 100)?"99":"100"; 
	List<String> paramList = new ArrayList<String>();
	
	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX001W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>總機構基本資料維護</title>
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
<input type="hidden" name="wlx02date" value="<%if(request.getAttribute("wlx02date") != null) out.print((String)request.getAttribute("wlx02date"));%>">    

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
                          總機構基本資料維護 
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
                          <tr class="sbody">
							 <td width='30%' align='left' bgcolor='#D8EFEE'>英文名稱</td>
							 <td width='70%' colspan=2 bgcolor='e7e7e7'><%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("english")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("english"));%>
	                        	 <input type='hidden' name='ENGLISH' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("english")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("english"));%>" >                            
	                         </td>
                          </tr>
                          <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>原始核准設立日期</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='hidden' name='SETUP_DATE' value="">
                            <input type='text' name='SETUP_DATE_Y' value="<%=SETUP_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=SETUP_DATE_M>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(SETUP_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(SETUP_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=SETUP_DATE_D>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(SETUP_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(SETUP_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日</font>
        						<font color='red' size=4>*</font>    
                            </td>
	  					</tr>
						
						<tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>原始核准該單位設立<br>之機構代號</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>						    
                            <% List setup_approval_unt = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='019' order by input_order",null,"");%>
							<select name='SETUP_APPROVAL_UNT'>
						 	 <%for(int i=0;i<setup_approval_unt.size();i++){%>
                             <option value="<%=(String)((DataObject)setup_approval_unt.get(i)).getValue("cmuse_id")%>"
                             <%if((WLX01 != null && WLX01.size() != 0) && ( ((DataObject)WLX01.get(0)).getValue("setup_approval_unt") != null && ((String)((DataObject)WLX01.get(0)).getValue("setup_approval_unt")).equals((String)((DataObject)setup_approval_unt.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                             ><%=(String)((DataObject)setup_approval_unt.get(i)).getValue("cmuse_name")%></option>                            
	                         <%}%>						
							</select>
                            </td>
	  					</tr>

							  					
						<tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>狀態</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>&nbsp;</td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>原始核准設立文號</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='SETUP_NO' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("setup_no")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("setup_no"));%>" size='50' maxlength='120' ><br>
                            (含字號，另號碼以阿拉伯數字列載)
                        </td>
                        </tr>                       
                        
                        <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>最近換照日期</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='hidden' name='CHG_LICENSE_DATE' value="">
                            <input type='text' name='CHG_LICENSE_DATE_Y' value="<%=CHG_LICENSE_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=CHG_LICENSE_DATE_M>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CHG_LICENSE_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CHG_LICENSE_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=CHG_LICENSE_DATE_D>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CHG_LICENSE_DATE_D.equals("0"+String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CHG_LICENSE_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日</font>
                            </td>
	  					</tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>最近換照文號</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='CHG_LICENSE_NO' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("chg_license_no")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("chg_license_no"));%>" size='50' maxlength='120' ><br>
                            (含字號，另號碼以阿拉伯數字列載)
                        </td>
                        </tr>
						
						<tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>最近換照事由</td>
						<% List chg_license_reason = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='004' order by input_order",null,"");%>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
                            <select name='CHG_LICENSE_REASON'>                            
                            <%for(int i=0;i<chg_license_reason.size();i++){%>
                            <option value="<%=(String)((DataObject)chg_license_reason.get(i)).getValue("cmuse_id")%>"
                            <%if((WLX01 != null && WLX01.size() != 0) && (((DataObject)WLX01.get(0)).getValue("chg_license_reason") != null && ((String)((DataObject)WLX01.get(0)).getValue("chg_license_reason")).equals((String)((DataObject)chg_license_reason.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                            ><%=(String)((DataObject)chg_license_reason.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>
                        </tr>
						
						<tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>原始開業日期</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='hidden' name='OPEN_DATE' value="">
                            <input type='text' name='OPEN_DATE_Y' value="<%=OPEN_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=OPEN_DATE_M>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(OPEN_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(OPEN_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=OPEN_DATE_D>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(OPEN_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(OPEN_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日</font>
                            </td>
	  					</tr>
                        
						                        
                        <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>開始營業日</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='hidden' name='START_DATE' value="">
                            <input type='text' name='START_DATE_Y' value="<%=START_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=START_DATE_M>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(START_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(START_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=START_DATE_D>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(START_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(START_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日(遷移後)</font>
                            </td>
	  					</tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>統一編號</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='BUSINESS_ID' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("business_id")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("business_id"));%>" size='10' maxlength='10' >                            
                        </td>
                        </tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>地方主管機關代號</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						<% 
						paramList.add(wlx01_m_year);
						List m2_name = DBManager.QueryDB_SQLParam("select bank_no,bank_name from ba01 where bank_type='B' and bank_kind='0' and m_year=?",paramList,"");%>
						<select name='M2_NAME'>
						<option value=""></option>
						 <%for(int i=0;i<m2_name.size();i++){%>
                            <option value="<%=(String)((DataObject)m2_name.get(i)).getValue("bank_no")%>"
                            <%if((WLX01 != null && WLX01.size() != 0) && (((DataObject)WLX01.get(0)).getValue("m2_name") != null && ((String)((DataObject)WLX01.get(0)).getValue("m2_name")).equals((String)((DataObject)m2_name.get(i)).getValue("bank_no")))) out.print("selected");%>
                            ><%=(String)((DataObject)m2_name.get(i)).getValue("bank_no")%>&nbsp;<%=(String)((DataObject)m2_name.get(i)).getValue("bank_name")%></option>                            
                         <%}%>	
                        </select>      
                        <font color='red' size=4>*</font>                      
                        </td>
                        </tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>是否參加電腦共用中心</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>							
						    <select name='CENTER_FLAG' onchange="javascript:setCenterNO(this.document.forms[0]);">
                            <option value="Y" <%if((WLX01 == null) || (WLX01 != null && WLX01.size() != 0) &&  (((DataObject)WLX01.get(0)).getValue("center_flag") != null && ((String)((DataObject)WLX01.get(0)).getValue("center_flag")).equals("Y"))  ) out.print("selected");%>>Y</option>
                            <option value="N" <%if( ((WLX01 != null && WLX01.size() != 0) &&  ((((DataObject)WLX01.get(0)).getValue("center_flag") == null) || ((String)((DataObject)WLX01.get(0)).getValue("center_flag")).equals("N")))   ) out.print("selected");%>>N</option>                            
							</select>      							
							<font color='red' size=4>*</font>    
                        </td>
                        </tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>參加的電腦共用中心機構代碼</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>							
						    <% 
						    List center_no = DBManager.QueryDB_SQLParam("select bank_no,bank_name from ba01 where bank_type='8' and bank_kind='0' and m_year=?",paramList,"");%>
							<select name='CENTER_NO' <%if( ((WLX01 != null && WLX01.size() != 0) &&  ((((DataObject)WLX01.get(0)).getValue("center_flag") == null) || ((String)((DataObject)WLX01.get(0)).getValue("center_flag")).equals("N")))   ) out.print("disabled");%>>
								<option value=""></option>
						 		<%for(int i=0;i<center_no.size();i++){%>
                            	<option value="<%=(String)((DataObject)center_no.get(i)).getValue("bank_no")%>"
                            	<%if((WLX01 != null && WLX01.size() != 0) && (((DataObject)WLX01.get(0)).getValue("center_no") != null && ((String)((DataObject)WLX01.get(0)).getValue("center_no")).equals((String)((DataObject)center_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                            	><%=(String)((DataObject)center_no.get(i)).getValue("bank_no")%>&nbsp;<%=(String)((DataObject)center_no.get(i)).getValue("bank_name")%></option>                            
                         		<%}%>							                            
							</select>                        
							<font color='red' size=4>*</font>    
                        </td>
                        </tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>本機構員工總人數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						     <input type='text' name='STAFF_NUM' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("staff_num") == null ?"":(((DataObject)WLX01.get(0)).getValue("staff_num")).toString());%>" size='5' maxlength='5' >
						     (不含分支機構人數)  <font color='red' size=4>*                               
                        </td>
                        </tr>
                        
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						     <input type='text' name='CREDIT_STAFF_NUM' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("credit_staff_num") == null ?"":(((DataObject)WLX01.get(0)).getValue("credit_staff_num")).toString());%>" size='5' maxlength='5' >
						     (含總機構及其分支機構之合計總人數)<font color='red' size=4>*</font><br>(依農金局95.2.7&nbsp;&nbsp;農授金字第&nbsp;&nbsp;0955010504&nbsp;&nbsp;號函之定義)                  
                        </td>
                        </tr>
                        <tr class="sbody">
						<td width='30%' align='left' bgcolor='#D8EFEE'>人員配置情形(人數)</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
							     正式職員<input type='text' name='CREDIT_STAFF' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("credit_staff") == null ?"":(((DataObject)WLX01.get(0)).getValue("credit_staff")).toString());%>" size='3'>
							     技工<input type='text' name='SKILL_STAFF' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("skill_staff") == null ?"":(((DataObject)WLX01.get(0)).getValue("skill_staff")).toString());%>" size='3'>
							     工友<input type='text' name='MANUAL_STAFF' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("manual_staff") == null ?"":(((DataObject)WLX01.get(0)).getValue("manual_staff")).toString());%>" size='3'>
							     特約人員<input type='text' name='TEMP_STAFF' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( ((DataObject)WLX01.get(0)).getValue("temp_staff") == null ?"":(((DataObject)WLX01.get(0)).getValue("temp_staff")).toString());%>" size='3'>                     
                        </td>
                        </tr>
                        <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>總機構地址</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
						<select name='HSIEN_ID_AREA_ID'>
						<%for(int i=0;i<hsien_id_area_id.size();i++){%>
						<option value="<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")%>"
						<%if((HSIEN_ID_AREA_ID).equals(((String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id"))+"/"+((String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")))) out.print("selected");%>
						><%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_name")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_name")%></option>						
						<%}%>
						</select>
						地區別
						<% List hsien_div = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='009' order by input_order",null,"");%>
						<select name='HSIEN_DIV'>
						 <%for(int i=0;i<hsien_div.size();i++){%>
                            <option value="<%=(String)((DataObject)hsien_div.get(i)).getValue("cmuse_id")%>"
                            <%if((WLX01 != null && WLX01.size() != 0) && ( ((DataObject)WLX01.get(0)).getValue("hsien_div_1") != null && ((String)((DataObject)WLX01.get(0)).getValue("hsien_div_1")).equals((String)((DataObject)hsien_div.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                            ><%=(String)((DataObject)hsien_div.get(i)).getValue("cmuse_name")%></option>                            
                         <%}%>						
						</select>
						<br><%=area_id%>
        				 <input type='text' name='ADDR' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("addr")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("addr"));%>" size='57' maxlength='120' >                            
        			    </td>
        			    </tr>	
                        
                        <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>電話</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='TELNO' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("telno")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("telno"));%>" size='20' maxlength='20' >(含區域號碼，並以"-"區隔)                            
        			    </td>
        			    </tr>	
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>傳真號碼</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='FAX' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("fax")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("fax"));%>" size='20' maxlength='20' >(含區域號碼，並以"-"區隔)                            
        			    </td>
        			    </tr>	
                        
                        <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>電子郵件帳號</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='EMAIL' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("email")==null)?"":  (String)((DataObject)WLX01.get(0)).getValue("email"));%>" size='60' maxlength='60' >
        			    </td>
        			    </tr>	
                        
                        <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>網址</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='WEB_SITE' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("web_site")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("web_site"));%>" size='60' maxlength='60' >
        			    </td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>國內營業分支機構家數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						<%if(WLX01_Count != null && WLX01_Count.size() != 0) out.print( ((DataObject)WLX01_Count.get(0)).getValue("bn02count") == null ?"":(((DataObject)WLX01_Count.get(0)).getValue("bn02count")).toString());%>						
						</td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>分支機構員工總人數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'><input type='hidden' name='wlx02staff_num' value='<%if(WLX01_Count != null && WLX01_Count.size() != 0) out.print( ((DataObject)WLX01_Count.get(0)).getValue("wlx02staff_num") == null ?"0":(((DataObject)WLX01_Count.get(0)).getValue("wlx02staff_num")).toString());%>'>
						<%if(WLX01_Count != null && WLX01_Count.size() != 0) out.print( ((DataObject)WLX01_Count.get(0)).getValue("wlx02staff_num") == null ?"":(((DataObject)WLX01_Count.get(0)).getValue("wlx02staff_num")).toString());%>
						</td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>理事人數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>	
						<%if(WLX01_Count != null && WLX01_Count.size() != 0) out.print( ((DataObject)WLX01_Count.get(0)).getValue("wlx04_1count") == null ?"":(((DataObject)WLX01_Count.get(0)).getValue("wlx04_1count")).toString());%>						 
						</td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>監事人數</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						<%if(WLX01_Count != null && WLX01_Count.size() != 0) out.print( ((DataObject)WLX01_Count.get(0)).getValue("wlx04_2count") == null ?"":(((DataObject)WLX01_Count.get(0)).getValue("wlx04_2count")).toString());%>						
						</td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>資訊單位地址</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
						<select name='IT_HSIEN_ID_AREA_ID'>
						<%for(int i=0;i<hsien_id_area_id.size();i++){%>
						<option value="<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")%>"
						<%if((IT_HSIEN_ID_AREA_ID).equals(((String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id"))+"/"+((String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")))) out.print("selected");%>
						><%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_name")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_name")%></option>
						<%}%>						
						</select>
	       				<input type='button' name='ToSameAddr' value="同地址" onClick="javascript:setAddr(form,'IT_ADDR');">                            
						<br><%=it_area_id%>
        				 <input type='text' name='IT_ADDR' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("it_addr")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("it_addr"));%>" size='57' maxlength='80' >                            
        			    </td>
        			    </tr>	
                        
                        <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>資訊人員姓名</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>
        				 <input type='text' name='IT_NAME' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("it_name")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("it_name"));%>" size='20' maxlength='20' >
        			    </td>
        			    </tr>
        			    
        			    <tr class="sbody">
						<td width='30%' bgcolor='#D8EFEE' align='left'>資訊人員電話</td>
						<td width='70%' colspan=2 bgcolor='e7e7e7'>						
        				 <input type='text' name='IT_TELNO' value="<%if(WLX01 != null && WLX01.size() != 0) out.print( (((DataObject)WLX01.get(0)).getValue("it_telno")==null)?"": (String)((DataObject)WLX01.get(0)).getValue("it_telno"));%>" size='20' maxlength='20' >
        				 (含區域號碼，並以"-"區隔) 
        			    </td>
        			    </tr>
                        
                        <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>裁撤生效日期</td>
						  <td width='35%' bgcolor='e7e7e7'>
						    <input type='hidden' name='CANCEL_DATE' value="">
                            <input type='text' name='CANCEL_DATE_Y' value="<%=CANCEL_DATE_Y%>" size='3' maxlength='3' onblur='CheckYear(this)' <%if((WLX01 == null) || ((WLX01 != null && WLX01.size() != 0) && ( (((DataObject)WLX01.get(0)).getValue("cancel_no") == null ) || ((String)((DataObject)WLX01.get(0)).getValue("cancel_no")).equals("N"))) ) out.print("disabled");%>>
        						<font color='#000000'>年
        						<select id="hide1" name=CANCEL_DATE_M <%if((WLX01 == null) || ((WLX01 != null && WLX01.size() != 0) && ( (((DataObject)WLX01.get(0)).getValue("cancel_no") == null ) || ((String)((DataObject)WLX01.get(0)).getValue("cancel_no")).equals("N"))) ) out.print("disabled");%>>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CANCEL_DATE_M.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CANCEL_DATE_M.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>月
        						<select id="hide1" name=CANCEL_DATE_D <%if((WLX01 == null) || ((WLX01 != null && WLX01.size() != 0) && ( (((DataObject)WLX01.get(0)).getValue("cancel_no") == null ) || ((String)((DataObject)WLX01.get(0)).getValue("cancel_no")).equals("N"))) ) out.print("disabled");%>>
        						<option></option>
        						<%
        							for (int j = 1; j < 32; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(CANCEL_DATE_D.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(CANCEL_DATE_D.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select></font><font color='#000000'>日</font>
                            </td>
                            <td width='35%'bgcolor='e7e7e7'>
                            確定裁撤
                            <select name=CANCEL_NO onchange="javascript:setCancelDate(this.document.forms[0]);">
                            <option value="Y" <%if((WLX01 != null && WLX01.size() != 0) &&  (((DataObject)WLX01.get(0)).getValue("cancel_no") != null && ((String)((DataObject)WLX01.get(0)).getValue("cancel_no")).equals("Y"))  ) out.print("selected");%>>Y</option>
                            <option value="N" <%if((WLX01 == null) || (WLX01.size() == 0) || ((WLX01 != null && WLX01.size() != 0) &&  ((((DataObject)WLX01.get(0)).getValue("cancel_no") == null) || ((String)((DataObject)WLX01.get(0)).getValue("cancel_no")).equals("N")))   ) out.print("selected");%>>N</option>   
                            </select>                         
                            </td>
	  					</tr>
                        
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>              
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>   
                      <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>                   	        	                                   		     
                        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','Master');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>                        				        				         				        
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Revoke','Master');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_revokeb.gif',1)"><img src="images/bt_revoke.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>                        				      
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>                        
                      <%}%>   
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
              		  </tr>
                    </table>
                  </div></td>
              </tr>   
              
              <!--tr>
                <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                    <tr bgcolor="#9AD3D0" class="sbody"> 
                        <td width=112 bgcolor=#9AD3D0><font face=細明體 color=#000000>每月申報資料</font></td>
                    	<td colspan=3>&nbsp;
                    	<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
                    	<input type='button' value="每月申報資料新增" onClick="javascript:AddWM(form,'<%=WLX01.size()%>');">
                    	->新增一筆申報資料
                    	<%}%>                    	
                    	</td>	
                    </tr>
                    <tr class="sbody">
						<td colspan=4 bgcolor='#D8EFEE' bgcolor='#D8EFEE' align='center'>基準日</td>						
        			</tr>
        			
        			<%if(WLX01_WM.size() == 0){%>
        			    <tr class="sbody">        			    
        				<td colspan=4 bgcolor='e7e7e7' bgcolor='#D8EFEE' align='center'>無每月申報資料</td>
        				</tr>
        			<%}else{
        				for(int i=0;i<WLX01_WM.size();i++){							
        			%>
						<tr class="sbody">
						<td bgcolor='e7e7e7' colspan=4 bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditWM&S_YEAR=<%=(((DataObject)WLX01_WM.get(i)).getValue("m_year")).toString()%>&S_MONTH=<%=(((DataObject)WLX01_WM.get(i)).getValue("m_month")).toString()%>'><%=(((DataObject)WLX01_WM.get(i)).getValue("m_year")).toString()%>/<%=(((DataObject)WLX01_WM.get(i)).getValue("m_month")).toString()%></a></td>					  
						</td>						
						</tr>						
					<%     
						}
					}%>	
        			
                </table></td>
              </tr-->  
                         
              <tr>
                <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                    <tr bgcolor="#9AD3D0" class="sbody"> 
                        <td width=112 bgcolor=#9AD3D0><font face=細明體 color=#000000>高階主管基本資料</font></td>
                    	<td colspan=3>&nbsp;
                    	<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
                    	<input type='button' value="高階主管資料新增" onClick="javascript:AddManager(form,'<%=WLX01.size()%>');">
                    	-->新增一筆高階主管資料
                    	<%}%>
                    	</td>	
                    </tr>
                    <tr class="sbody">
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>職稱</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>姓名</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>電話</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>就任日期</td>						
        			</tr>
        			
        			<%if(WLX01_M.size() == 0){%>
        			    <tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無高階主管相關資料</td>
        				</tr>
        			<%}else{
        				for(int i=0;i<WLX01_M.size();i++){
							if(((String)((DataObject)WLX01_M.get(i)).getValue("abdicate_code") == null) || ((String)((DataObject)WLX01_M.get(i)).getValue("abdicate_code")).equals("N")){        				  
        			%>
						<tr class="sbody">
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditM&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX01_M.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX01_M.get(i)).getValue("seq_no")).toString()%>&haveCancel=<%=haveCancel%>'><%=(((DataObject)WLX01_M.get(i)).getValue("cmuse_name") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("cmuse_name")%></a></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditM&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX01_M.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX01_M.get(i)).getValue("seq_no")).toString()%>&haveCancel=<%=haveCancel%>'><%=(((DataObject)WLX01_M.get(i)).getValue("name") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("name")%></a></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=(((DataObject)WLX01_M.get(i)).getValue("telno") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("telno")%></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=Utility.getCHTdate((((DataObject)WLX01_M.get(i)).getValue("induct_date")).toString().substring(0, 10), 0)%></td>						        				
						</tr>						
					<%      }//end if
						}//end of for
					  }//end of else
					%>	
        			
                </table></td>
              </tr>  
              <tr>
                <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                    <tr bgcolor="#9AD3D0" class="sbody"> 
                        <td colspan=4 bgcolor=#9AD3D0><font face=細明體 color=#000000>高階主管卸任基本資料</font></td>                    	
                    </tr>
                    <tr class="sbody">
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>職稱</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>姓名</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>就任日期</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>卸任日期</td>						
        			</tr>
        			
        			<%if(WLX01_M.size() == 0){%>
        			    <tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無卸任之高階主管相關資料</td>
        				</tr>
        			<%}else{
        			    boolean haveabdicate=false;
        				for(int i=0;i<WLX01_M.size();i++){
							if(((String)((DataObject)WLX01_M.get(i)).getValue("abdicate_code") != null) && ((String)((DataObject)WLX01_M.get(i)).getValue("abdicate_code")).equals("Y")){        				  
							    haveabdicate=true;
        			%>
        			    <tr class="sbody">
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditM&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX01_M.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX01_M.get(i)).getValue("seq_no")).toString()%>&haveCancel=<%=haveCancel%>'><%=(((DataObject)WLX01_M.get(i)).getValue("cmuse_name") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("cmuse_name")%></a></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditM&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX01_M.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX01_M.get(i)).getValue("seq_no")).toString()%>&haveCancel=<%=haveCancel%>'><%=(((DataObject)WLX01_M.get(i)).getValue("name") == null) ?"&nbsp;":(String)((DataObject)WLX01_M.get(i)).getValue("name")%></a></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=Utility.getCHTdate((((DataObject)WLX01_M.get(i)).getValue("induct_date")).toString().substring(0, 10), 0)%></td>						        										
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=Utility.getCHTdate((((DataObject)WLX01_M.get(i)).getValue("abdicate_date")).toString().substring(0, 10), 0)%></td>						        																
						</tr>
						
					<%      }//end if
						}//end for%>
					   	<%if(!haveabdicate){%>        				
        				<tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無卸任之高階主管相關資料</td>
        				</tr>
        			    <%}//haveabdicate=false
					  }//end of else
					%>	        			
                </table></td>
              </tr>
              <tr>
                <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                    <tr bgcolor="#9AD3D0" class="sbody"> 
                        <td width=112 bgcolor=#9AD3D0><font face=細明體 color=#000000>稽核人員基本資料</font></td>
                    	<td colspan=3>&nbsp;
                    	<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
                    	<input type='button' value="稽核人員資料新增" onClick="javascript:AddAudit(form,'<%=WLX01.size()%>');">
                    	-->新增一筆稽核人員資料
                    	<%}%>
                    	</td>	
                    </tr>
                    <tr class="sbody">
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>姓名</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>電話</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>隸屬部門</td>
        			</tr>
        			
        			<%if(WLX01_Audit.size() == 0){%>
        			    <tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無稽核人員相關資料</td>
        				</tr>
        			<%}else{
        				for(int i=0;i<WLX01_Audit.size();i++){
        			%>
						<tr class="sbody">
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX001W.jsp?act=EditAudit&bank_no=<%=bank_no%>&seq_no=<%=(((DataObject)WLX01_Audit.get(i)).getValue("seq_no")).toString()%>&haveCancel=<%=haveCancel%>'><%=(((DataObject)WLX01_Audit.get(i)).getValue("name") == null) ?"&nbsp;":(String)((DataObject)WLX01_Audit.get(i)).getValue("name")%></a></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=(((DataObject)WLX01_Audit.get(i)).getValue("telno") == null) ?"&nbsp;":(String)((DataObject)WLX01_Audit.get(i)).getValue("telno")%></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=(((DataObject)WLX01_Audit.get(i)).getValue("department") == null) ?"&nbsp;":(String)((DataObject)WLX01_Audit.get(i)).getValue("department")%></td>
						</tr>						
					<%     
						}//end of for
					  }//end of else
					%>	
        			
                </table></td>
              </tr>   
      </table></td>
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
                          <li>本網頁提供總機構基本資料維護功能。</li>                          
                          <li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</li>                         
                          <li>新增時,可直接於空格內更改資料，資料更改完畢後，按"確定"即將本表上的資料於資料庫中建檔。</li>
                          <li>按<font color="#666666">【高階主管資料新增】</font>新增一筆高階主管基本資料。</li>
                          <li>按<font color="#666666">【稽核人員資料新增】</font>新增一筆稽核人員基本資料。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>  
                          <li><font color='red'>國內營業分支機構家數、分支機構員工總人數若需修正，請至國內營業分支構機基本資料維護(FX002W)進行資料修改。</li>                       
                          <li>【<font color="red" size=4>*</font>】為必填欄位。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
 <%System.out.println("FX001W_Edit.ContentLength="+request.getContentLength());%>
</body>
</html>
