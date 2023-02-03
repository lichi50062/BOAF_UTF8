<% 
//94.01.11 fix 更改密碼至少為 6碼 by 2295
//94.01.13 fix 更改密碼不可以帳號相同 by 2295
//94.10.28 fix 新增公告區 by 2495
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.dao.DAOFactory" %>
<%@ page import="com.tradevan.util.dao.RdbCommonDao" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.math.BigDecimal" %>
<%@ page import="java.io.*" %>

<%!
    //取得該標題之公告資料
    private List Get_WLX_Notify(){    	
    	//查詢條件    
    	String sqlCmd = "select seq_no,headmark,to_char(notify_date,'yyyy/mm/dd hh:mi') as notify_date,to_char(notify_end_date,'yyyy/mm/dd') as notify_end_date,append_file,user_id,user_name,to_char(update_date,'mm/dd/yyyy hh:mi:ss') as update_date,appfile_link,notify_url from WLX_Notify ";  		
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"seq_no");            
        return dbData;
    }
    
    //取得該序號之最大值
    private List TakeMaxSeqno(){    	
    	//查詢條件    
    	String sqlCmd = "select max(seq_no) as maxseqno from WLX_Notify";  		
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"maxseqno");            
        return dbData;
    }

    //取得資料筆數
    private List Count_WLX_Notify(){    	
    	//查詢條件    
    	String sqlCmd = "select count(seq_no) as countseqno from WLX_Notify";  		
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"countseqno");            
        return dbData;
    }
    
    
    //取得該user_id之user_name
    private List TakeUserName(String muser_id){
            List paramList = new ArrayList();
            //查詢條件    
    		String sqlCmd = "select *  from WTT01 where  muser_id =?";  	
    		paramList.add(muser_id);
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
  
%>






<html>
<head>
<title>網際網路申報系統</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/Common.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
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
function doSubmit(form){
    if(trimString(form.muser_id.value) =="" ){
       alert("用戶帳號不可為空白");
       form.muser_id.focus();
       return;
    }
    if(trimString(form.muser_password.value) =="" ){
       alert("用戶密碼不可為空白");
       form.muser_password.focus();
       return;
    }else{            
       if((trimString(form.ChangePwd.value) !="" ) && (form.muser_password.value == form.ChangePwd.value)){
          alert("卻更改之密碼不可與舊密碼相同");
          form.ChangePwd.focus();
          return;
       }       
    }
    
    if((trimString(form.ChangePwd.value) !="" ) || (trimString(form.ConfirmPwd.value) !="" )){
       if(form.ChangePwd.value != form.ConfirmPwd.value){
          alert("欲更改之密碼與確認密碼不符合");
          form.ConfirmPwd.focus();
          return;
       }
    }
    
    if(trimString(form.ChangePwd.value) !=""  && form.ChangePwd.value.length < 6 ){
       alert("欲更改之密碼至少為6碼");
       form.ChangePwd.focus();
       return;
    }
    
    if(trimString(form.ChangePwd.value) !="" && (trimString(form.ChangePwd.value) == trimString(form.muser_id.value))){
       alert("欲更改之密碼不可與帳號相同");     
       form.ChangePwd.focus();
       return;   
    }
    form.submit();
}

function doLink(url){
	
	window.open(url); 
}

var keyPressed;

function chkKey(e){  
    keyPressed = String.fromCharCode(window.event.keyCode);
    if (keyPressed == "\x0D") {
      doSubmit(window.document.loginfrm);
    }  
}

if (window.document.captureEvents!=null) 
  window.document.captureEvents(Event.KEYPRESS)
window.document.onkeypress = chkKey;


//-->

</script>
</head>

<body background="images/bg_1.gif" leftmargin="0" topmargin="0" onLoad="">
<form name="loginfrm" method=post action='/pages/Login.jsp'>
<table width="764" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#006666">
  <tr>
    <td height="20" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="764" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#70BEBB">
          <tr> 
          <td width="68"><img src="images/Login_Image_01.gif" width="68" height="23"></td>
          <td width="261" bgcolor="#4B9C99">&nbsp;</td>
          <td width="34" bgcolor="#4B9C99">&nbsp;</td>
          <td width="396" bgcolor="#4B9C99">&nbsp;</td>
          <td width="21"><img src="images/Login_Image_02.gif" width="21" height="23"></td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_03.gif" width="68" height="233"></td>
          <td><img src="images/Login_Image_04.gif" width="261" height="233"></td>
          <td><img src="images/Login_Image_05.gif" width="34" height="233"></td>
          <td><img src="images/Login_Image_06.gif" width="380" height="233"></td>
          <td><img src="images/Login_Image_07.gif" width="21" height="233"></td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_08.gif" width="68" height="69"></td>
          <td><img src="images/Login_Image_09.gif" width="261" height="69"></td>
          <td><img src="images/Login_Image_10.gif" width="34" height="69"></td>
          <td rowspan="2">
            <table width="330" border="0" align="center" cellpadding="0" cellspacing="0">													              
              <tr>
                <td><table width="330" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#006666">
                    <tr>
                    	
        <td bordercolor="#63B1AE"><table width="330" border="0" align="center" cellpadding="1" cellspacing="0" bordercolor="#FFFFFF">                                           
<%
         String headtext[] = new String[40];
		 String headtext_cut[] = new String[40];
         String notifydate[] = new String[40];
         int fileamount[] = new int[40];
		 String urlamount[] = new String[40];
		 int seqindex[] = new int[40];
         String headmark="",seq_no="0",notify_date="",notify_end_date="";
		 List WLX_SEQNO = TakeMaxSeqno();	
		 int j=0;
		 int seqamount=0;					 
         int seqcount = Integer.parseInt( (((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno")==null )? "-1" : ((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno").toString());        		  					
         int notify_number=0;
         System.out.println("最大公告數目 ="+seqcount);
		 List WLX_Notify = Get_WLX_Notify();
		 List Count_Notify = Count_WLX_Notify();
		 String countseqno = (Count_Notify == null ? "0":(((DataObject)Count_Notify.get(0)).getValue("countseqno")).toString());
		 System.out.println("countseqno ="+countseqno);
		 int count = Integer.parseInt(countseqno);

		 for(j=0; j<count;j++)
         { 
			//公告結束日期是否到期			
			notify_end_date = (String)((DataObject)WLX_Notify.get(j)).getValue("notify_end_date");
			String sub_notify = notify_end_date.substring(0,4);
			int end_year=Integer.parseInt(sub_notify);
			//System.out.println("end_year:"+end_year);
			sub_notify= notify_end_date.substring(5,7);
			int end_month=Integer.parseInt(sub_notify);				
			//System.out.println("end_month:"+end_month);
			sub_notify= notify_end_date.substring(8);				
			int end_day=Integer.parseInt(sub_notify);
			//System.out.println("end_day:"+end_day);
			
			Calendar today=new GregorianCalendar();
			end_month--;
			Calendar end_date=new GregorianCalendar(end_year,end_month,end_day);
			long days=(end_date.getTimeInMillis()-today.getTimeInMillis())/1000/60/60/24;
			//System.out.println("days:"+days);
			if(days>=0)  
			{
				
				if(notify_number<40)
				{	
				seq_no = (((DataObject)WLX_Notify.get(j)).getValue("seq_no")).toString();
				seqindex[notify_number]=Integer.parseInt(seq_no);
				headmark = (String)((DataObject)WLX_Notify.get(j)).getValue("headmark");
				headtext[notify_number] = headmark;	 
				System.out.println("公告序號有哪些 ="+seqindex[notify_number]);
				//處理西元年轉民國年
				notify_date = (String)((DataObject)WLX_Notify.get(j)).getValue("notify_date");
				String  sub_str1 = notify_date.substring(0,4);								
				int east_year=Integer.parseInt(sub_str1);				
				east_year = east_year-1911;
				sub_str1 = Integer.toString(east_year);
				String sub_str2 = notify_date.substring(4);
				notify_date = east_year + sub_str2;

				notifydate[notify_number]=notify_date;
    			String append_file = (String)((DataObject)WLX_Notify.get(j)).getValue("append_file");
				String notify_url = (String)((DataObject)WLX_Notify.get(j)).getValue("notify_url");
				//System.out.println("有網址的公告notify_url:"+notify_url);
    			if(append_file.equals(" "))
    		  		fileamount[notify_number]=0;
    		  	else
    		  		fileamount[notify_number]=1;
				
				if(notify_url=="null")
    		  		urlamount[notify_number]="null";
    		  	else
    		  		urlamount[notify_number]=notify_url;
        		//System.out.println("該公告標題是否有檔案 ="+fileamount[notify_number]);
				notify_number++;
		     	}
		     	else
					break;
		  }
             continue;
		 }	
		 System.out.println("本次一共有幾則公告="+notify_number);     
          
	   if(notify_number!=0 && seqcount!=(-1))
           {
        %>
	<tr colspan="0" bordercolor="#63B1AE" bgcolor="#408C87"><img src="images/hotnews.gif" width="70" height="18"></tr>	
        <%
		 System.out.println("======="+notify_number); 
			
					               
            %>	
            <script language=JavaScript> 
						var index = 40
						seqno = new Array(40); 
						linkfile = new Array(0);
						linkpage = new Array(0); 
						text = new Array(39);						
						amount = new Array(39);
						urlboolem = new Array(39);
						startday = new Array(39);
            index = '<%=notify_number%>'
						linkfile[0] ='/pages/DomloadNotify.jsp'
						linkpage[0] ='/pages/LinkNotify.jsp'
				
						seqno[0] = '<%=seqindex[0] %>' 
						seqno[1] = '<%=seqindex[1] %>'
						seqno[2] = '<%=seqindex[2] %>' 
						seqno[3] = '<%=seqindex[3] %>'
						seqno[4] = '<%=seqindex[4] %>'
						seqno[5] = '<%=seqindex[5] %>' 
						seqno[6] = '<%=seqindex[6] %>'
						seqno[7] = '<%=seqindex[7] %>' 
						seqno[8] = '<%=seqindex[8] %>'
						seqno[9] = '<%=seqindex[9] %>' 
						seqno[10] = '<%=seqindex[10] %>' 
						seqno[11] = '<%=seqindex[11] %>'
						seqno[12] = '<%=seqindex[12] %>' 
						seqno[13] = '<%=seqindex[13] %>'
						seqno[14] = '<%=seqindex[14] %>'
						seqno[15] = '<%=seqindex[15] %>' 
						seqno[16] = '<%=seqindex[16] %>'
						seqno[17] = '<%=seqindex[17] %>' 
						seqno[18] = '<%=seqindex[18] %>'
						seqno[19] = '<%=seqindex[19] %>' 
						seqno[20] = '<%=seqindex[20] %>' 
						seqno[21] = '<%=seqindex[21] %>'
						seqno[22] = '<%=seqindex[22] %>' 
						seqno[23] = '<%=seqindex[23] %>'
						seqno[24] = '<%=seqindex[24] %>'
						seqno[25] = '<%=seqindex[25] %>' 
						seqno[26] = '<%=seqindex[26] %>'
						seqno[27] = '<%=seqindex[27] %>' 
						seqno[28] = '<%=seqindex[28] %>'
						seqno[29] = '<%=seqindex[29] %>' 
						seqno[30] = '<%=seqindex[30] %>' 
						seqno[31] = '<%=seqindex[31] %>'
						seqno[32] = '<%=seqindex[32] %>' 
						seqno[33] = '<%=seqindex[33] %>'
						seqno[34] = '<%=seqindex[34] %>'
						seqno[35] = '<%=seqindex[35] %>' 
						seqno[36] = '<%=seqindex[36] %>'
						seqno[37] = '<%=seqindex[37] %>' 
						seqno[38] = '<%=seqindex[38] %>'
						seqno[39] = '<%=seqindex[39] %>'
																							
						text[0] = '<%=headtext[0] %>' 
						text[1] = '<%=headtext[1] %>'
						text[2] = '<%=headtext[2] %>' 
						text[3] = '<%=headtext[3] %>'
						text[4] = '<%=headtext[4] %>'
						text[5] = '<%=headtext[5] %>' 
						text[6] = '<%=headtext[6] %>'
						text[7] = '<%=headtext[7] %>' 
						text[8] = '<%=headtext[8] %>'
						text[9] = '<%=headtext[9] %>' 
						text[10] = '<%=headtext[10] %>' 
						text[11] = '<%=headtext[11] %>'
						text[12] = '<%=headtext[12] %>' 
						text[13] = '<%=headtext[13] %>'
						text[14] = '<%=headtext[14] %>'
						text[15] = '<%=headtext[15] %>' 
						text[16] = '<%=headtext[16] %>'
						text[17] = '<%=headtext[17] %>' 
						text[18] = '<%=headtext[18] %>'
						text[19] = '<%=headtext[19] %>'
						text[20] = '<%=headtext[20] %>' 
						text[21] = '<%=headtext[21] %>'
						text[22] = '<%=headtext[22] %>' 
						text[23] = '<%=headtext[23] %>'
						text[24] = '<%=headtext[24] %>'
						text[25] = '<%=headtext[25] %>' 
						text[26] = '<%=headtext[26] %>'
						text[27] = '<%=headtext[27] %>' 
						text[28] = '<%=headtext[8] %>'
						text[29] = '<%=headtext[29] %>' 
						text[30] = '<%=headtext[30] %>' 
						text[31] = '<%=headtext[31] %>'
						text[32] = '<%=headtext[32] %>' 
						text[33] = '<%=headtext[33] %>'
						text[34] = '<%=headtext[34] %>'
						text[35] = '<%=headtext[35] %>' 
						text[36] = '<%=headtext[36] %>'
						text[37] = '<%=headtext[37] %>' 
						text[38] = '<%=headtext[38] %>'
						text[39] = '<%=headtext[39] %>'
						
						
						amount[0] = '<%=fileamount[0] %>' 
						amount[1] = '<%=fileamount[1] %>'
						amount[2] = '<%=fileamount[2] %>' 
						amount[3] = '<%=fileamount[3] %>'
						amount[4] = '<%=fileamount[4] %>'
						amount[5] = '<%=fileamount[5] %>' 
						amount[6] = '<%=fileamount[6] %>'
						amount[7] = '<%=fileamount[7] %>' 
						amount[8] = '<%=fileamount[8] %>'
						amount[9] = '<%=fileamount[9] %>' 
						amount[10] = '<%=fileamount[10] %>' 
						amount[11] = '<%=fileamount[11] %>'
						amount[12] = '<%=fileamount[12] %>' 
						amount[13] = '<%=fileamount[13] %>'
						amount[14] = '<%=fileamount[14] %>'
						amount[15] = '<%=fileamount[15] %>' 
						amount[16] = '<%=fileamount[16] %>'
						amount[17] = '<%=fileamount[17] %>' 
						amount[18] = '<%=fileamount[18] %>'
						amount[19] = '<%=fileamount[19] %>'
						amount[20] = '<%=fileamount[20] %>' 
						amount[21] = '<%=fileamount[21] %>'
						amount[22] = '<%=fileamount[22] %>' 
						amount[23] = '<%=fileamount[23] %>'
						amount[24] = '<%=fileamount[24] %>'
						amount[25] = '<%=fileamount[25] %>' 
						amount[26] = '<%=fileamount[26] %>'
						amount[27] = '<%=fileamount[27] %>' 
						amount[28] = '<%=fileamount[28] %>'
						amount[29] = '<%=fileamount[29] %>' 
						amount[30] = '<%=fileamount[30] %>' 
						amount[31] = '<%=fileamount[31] %>'
						amount[32] = '<%=fileamount[32] %>' 
						amount[33] = '<%=fileamount[33] %>'
						amount[34] = '<%=fileamount[34] %>'
						amount[35] = '<%=fileamount[35] %>' 
						amount[36] = '<%=fileamount[36] %>'
						amount[37] = '<%=fileamount[37] %>' 
						amount[38] = '<%=fileamount[38] %>'
						amount[39] = '<%=fileamount[39] %>'

						urlboolem[0] = '<%=urlamount[0] %>' 
						urlboolem[1] = '<%=urlamount[1] %>'
						urlboolem[2] = '<%=urlamount[2] %>' 
						urlboolem[3] = '<%=urlamount[3] %>'
						urlboolem[4] = '<%=urlamount[4] %>'
						urlboolem[5] = '<%=urlamount[5] %>' 
						urlboolem[6] = '<%=urlamount[6] %>'
						urlboolem[7] = '<%=urlamount[7] %>' 
						urlboolem[8] = '<%=urlamount[8] %>'
						urlboolem[9] = '<%=urlamount[9] %>' 
						urlboolem[10] = '<%=urlamount[10] %>' 
						urlboolem[11] = '<%=urlamount[11] %>'
						urlboolem[12] = '<%=urlamount[12] %>' 
						urlboolem[13] = '<%=urlamount[13] %>'
						urlboolem[14] = '<%=urlamount[14] %>'
						urlboolem[15] = '<%=urlamount[15] %>' 
						urlboolem[16] = '<%=urlamount[16] %>'
						urlboolem[17] = '<%=urlamount[17] %>' 
						urlboolem[18] = '<%=urlamount[18] %>'
						urlboolem[19] = '<%=urlamount[19] %>'
						urlboolem[20] = '<%=urlamount[20] %>' 
						urlboolem[21] = '<%=urlamount[21] %>'
						urlboolem[22] = '<%=urlamount[22] %>' 
						urlboolem[23] = '<%=urlamount[23] %>'
						urlboolem[24] = '<%=urlamount[24] %>'
						urlboolem[25] = '<%=urlamount[25] %>' 
						urlboolem[26] = '<%=urlamount[26] %>'
						urlboolem[27] = '<%=urlamount[27] %>' 
						urlboolem[28] = '<%=urlamount[28] %>'
						urlboolem[29] = '<%=urlamount[29] %>' 
						urlboolem[30] = '<%=urlamount[30] %>' 
						urlboolem[31] = '<%=urlamount[31] %>'
						urlboolem[32] = '<%=urlamount[32] %>' 
						urlboolem[33] = '<%=urlamount[33] %>'
						urlboolem[34] = '<%=urlamount[34] %>'
						urlboolem[35] = '<%=urlamount[35] %>' 
						urlboolem[36] = '<%=urlamount[36] %>'
						urlboolem[37] = '<%=urlamount[37] %>' 
						urlboolem[38] = '<%=urlamount[38] %>'
						urlboolem[39] = '<%=urlamount[39] %>'
						
						startday[0] = '<%=notifydate[0] %>' 
						startday[1] = '<%=notifydate[1] %>'
						startday[2] = '<%=notifydate[2] %>' 
						startday[3] = '<%=notifydate[3] %>'
						startday[4] = '<%=notifydate[4] %>'
						startday[5] = '<%=notifydate[5] %>' 
						startday[6] = '<%=notifydate[6] %>'
						startday[7] = '<%=notifydate[7] %>' 
						startday[8] = '<%=notifydate[8] %>'
						startday[9] = '<%=notifydate[9] %>' 
						startday[10] = '<%=notifydate[10] %>' 
						startday[11] = '<%=notifydate[11] %>'
						startday[12] = '<%=notifydate[12] %>' 
						startday[13] = '<%=notifydate[13] %>'
						startday[14] = '<%=notifydate[14] %>'
						startday[15] = '<%=notifydate[15] %>' 
						startday[16] = '<%=notifydate[16] %>'
						startday[17] = '<%=notifydate[17] %>' 
						startday[18] = '<%=notifydate[18] %>'
						startday[19] = '<%=notifydate[19] %>'
						startday[20] = '<%=notifydate[20] %>' 
						startday[21] = '<%=notifydate[21] %>'
						startday[22] = '<%=notifydate[22] %>' 
						startday[23] = '<%=notifydate[23] %>'
						startday[24] = '<%=notifydate[24] %>'
						startday[25] = '<%=notifydate[25] %>' 
						startday[26] = '<%=notifydate[26] %>'
						startday[27] = '<%=notifydate[27] %>' 
						startday[28] = '<%=notifydate[28] %>'
						startday[29] = '<%=notifydate[29] %>' 
						startday[30] = '<%=notifydate[30] %>' 
						startday[31] = '<%=notifydate[31] %>'
						startday[32] = '<%=notifydate[32] %>' 
						startday[33] = '<%=notifydate[33] %>'
						startday[34] = '<%=notifydate[34] %>'
						startday[35] = '<%=notifydate[35] %>' 
						startday[36] = '<%=notifydate[36] %>'
						startday[37] = '<%=notifydate[37] %>' 
						startday[38] = '<%=notifydate[38] %>'
						startday[39] = '<%=notifydate[39] %>'
						index--;
						document.write ("<marquee  bgcolor=white scrollamount='2' scrolldelay='6' direction= 'up' width='380' id=xiaoqing height='110' onmouseover=xiaoqing.stop() onmouseout=xiaoqing.start()>"); 				  				
						for (i=index;i>-1;i--)
						{ 										  																														
							if(amount[i]!=0)
							{	 							  		
							    document.write ("&nbsp;<a href=\"javascript:doLink('"+linkfile[0]+"?seq_no="+seqno[i]+"');\"><img src=images/download.gif>"); 		 		
							    //if(!(urlboolem[i]==" "))
							    if(!(urlboolem[i]=="null"))
							    { 
							    	if(i==index)
							        	document.write("</A>&nbsp;&nbsp;<B><font size='2' color=BLUE>"+startday[i]+"&nbsp;<a href=\"javascript:doLink('"+urlboolem[i]+"');\">"+text[i]+"</A></font></B><br>"); 
					  		    	else
							        	document.write("</A>&nbsp;&nbsp;<B><font size='2' color=BLACK>"+startday[i]+"&nbsp;<a href=\"javascript:doLink('"+urlboolem[i]+"');\">"+text[i]+"</A></font></B><br>"); 
							    }
							    else
							    {
								if(i==index)
									document.write ("</A>&nbsp;&nbsp;<B><font size='2' color=BLUE>"+startday[i]+"&nbsp;"+text[i]+"</font><br>");							 
							    	else
									document.write ("</A>&nbsp;&nbsp;<B><font size='2' color=BLACK>"+startday[i]+"&nbsp;"+text[i]+"</font><br>");							 									
							    }
							}
					  		else 
					  		{
							    if(!(urlboolem[i]=="null"))
							    { 
							    	if(i==index)
							        	document.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><font size='2' color=BLUE>"+startday[i]+"&nbsp;<a href=\"javascript:doLink('"+urlboolem[i]+"');\">"+text[i]+"</A></font</B>><br>"); 
					  		    	else
							        	document.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><font size='2' color=BLACK>"+startday[i]+"&nbsp;<a href=\"javascript:doLink('"+urlboolem[i]+"');\">"+text[i]+"</A></font></B><br>"); 
							    }
							    else
							    {
								if(i==index)
									document.write ("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><font size='2' color=BLUE>"+startday[i]+"&nbsp;"+text[i]+"</font></B><br>");							 
							    	else
									document.write ("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><font size='2' color=BLACK>"+startday[i]+"&nbsp;"+text[i]+"</font></B><br>");							 									
							    }
							}									  			  
						} 
						
						document.write ("</marquee>") 					
			</script>
		<%}else{%>
		<tr> 
                <td width="12"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td width="355" class="sbody">網際網路申報系統線上申請作業說明。</td>
              </tr>
              <tr> 
                <td valign="top"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td class="sbody">本系統全年全天候開放，惟須暫停連線服務時，將事先於本網站首頁公佈。</td>
              </tr>
		<%	}%>
                          </tr>	
			                    </table>
                     </td>
                    </tr>
                    
                  </table></td>
              </tr>
              
            </table>
	  </td>
          <td bgcolor="#70BEBB">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/Login_Image_11.gif"><img src="images/Login_Image_11.gif" width="68" height="100"></td>
          <td rowspan="2" valign="top" bgcolor="#78B5B3">
			<table width="250" border="0" align="center" cellpadding="0" cellspacing="1" class="sbody">
              <tr> 
                <td>用戶帳號:</td>
                <td><input type="text" maxlength=12 name=muser_id size=12></td>                
              </tr>
              <tr> 
                <td width="65">用戶密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=muser_password size=20></td>
              </tr>              
              <tr> 
                <td width="65">更改密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=ChangePwd size=20></td>
              </tr>
              <tr> 
                <td width="65">確認密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=ConfirmPwd size=20></td>
              </tr>
              
              <tr>
                <td>&nbsp;</td>
                <td><a href="javascript:doSubmit(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image23','','images/Login_bt_3b.gif',1)"><img src="images/Login_bt_3.gif" name="Image23" width="58" height="22" border="0"></a> 
                  <a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image20','','images/Login_bt_1b.gif',1)"><img src="images/Login_bt_1.gif" name="Image20" width="58" height="22" border="0"></a></td>
              </tr>
            </table>
          </td>
          <td background="images/Login_Image_12.gif"><img src="images/Login_Image_12.gif" width="34" height="100"></td>
          <td bgcolor="#70BEBB">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/Login_Image_11.gif"><img src="images/Login_Image_11.gif" width="68" height="42"></td>
          <td background="images/Login_Image_11b.gif"><img src="images/Login_Image_11b.gif" width="34" height="42"></td>
          <td bgcolor="#4B9C99"><table width="366" border="0" align="center" cellpadding="0" cellspacing="0" class="sbody">
              <tr> 
                <td>建議使用IE 5.0以上版本之瀏覽器螢幕解析度800X600以上瀏覽</td>
              </tr>
              <tr>
                <td>版權所有 翻版必究 Copyringht 2004<a href="#"> BOAF </a>All Rights 
                  Reserved .</td>
              </tr>
            </table></td>
          <td bgcolor="#4B9C99">&nbsp;</td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_13.gif" width="68" height="18"></td>
          <td><img src="images/Login_Image_14.gif" width="261" height="18"></td>
          <td><img src="images/Login_Image_15.gif" width="34" height="18"></td>
          <td bgcolor="#4B9C99"><img src="images/space_1.gif" width="5" height="12"></td>
          <td><img src="images/Login_Image_16.gif" width="21" height="18"></td>
        </tr><a name="start">        
      </table></td>
  </tr>
</table>
</form>
</body>
<script language="javascript">
  window.scrollTo(window.pageXOffset,1200);
</script>

</html>
