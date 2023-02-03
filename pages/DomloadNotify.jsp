//102.12.03 add pdf 下載 by 2295
<%@  page  import="com.jspsmart.upload.*"  %>
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


<%
String seq_no=" ",headmark=" ",notify_date=" ",append_file=" ",user_id=" ",user_name=" ",user_telon=" ",append_file2=" ",append_file3=" ",bank_no=" ",connect_peopole=" ";

String[] seq_number = request.getParameterValues("seq_no");
seq_no = seq_number[0];  
List WLX_Notify = Get_WLX_Notify(seq_no);

headmark = (((DataObject)WLX_Notify.get(0)).getValue("headmark")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("headmark");	
user_id = (String)((DataObject)WLX_Notify.get(0)).getValue("user_id");
user_name = (((DataObject)WLX_Notify.get(0)).getValue("user_name")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("user_name");	
append_file = (((DataObject)WLX_Notify.get(0)).getValue("append_file")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("append_file");	


%>
<%!
    //取得該標題之公告資料
    private List Get_WLX_Notify(String seq_no){
    		//查詢條件    
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select seq_no,headmark,to_char(notify_date,'yyyy/mm/dd hh:mi:ss') as notify_date,to_char(notify_end_date,'mm/dd/yyyy hh:mi:ss') as notify_end_date,append_file,user_id,user_name,to_char(update_date,'mm/dd/yyyy hh:mi:ss') as update_date,appfile_link from WLX_Notify where  seq_no =? ";
    		paramList.add(seq_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"seq_no");            
        return dbData;
    }


%>
<HTML>
<head>
<title>download.jsp</title>
</head>
</HTML>


<jsp:useBean  id="mySmartUpload"  scope="page"  class="com.jspsmart.upload.SmartUpload"  />
<%  
       String head_mark = headmark;       
			 String file_name = append_file;
			 //String  downloadlink = "C:\\Sun\\WebServer6.1\\BOAF\\exp\\wlx\\WLX_Notify\\"+file_name;
			 String downloadlink = Utility.getProperties("notifyDir")+file_name;
			 System.out.println("downloadlink:"+downloadlink);
       mySmartUpload.initialize(pageContext);
       
       
       int index = file_name.indexOf('.');      
       String file_type = file_name.substring(index+1,index+4);
       System.out.println("file_type:"+file_type);
       if(file_type.equals("doc"))
       {    
       		System.out.println("doc_file_name:"+file_name);
       		mySmartUpload.setContentDisposition(null);
       		mySmartUpload.downloadFile(downloadlink,"application/msword","download.doc"); 
       } 
       if(file_type.equals("xls"))
       {
     		  System.out.println("file_name:"+file_name);
     		  mySmartUpload.setContentDisposition(null);       
     		  mySmartUpload.downloadFile(downloadlink,"application/vnd.ms-excel","download.xls");       
       }
       if(file_type.equals("pps"))
       {
     		  System.out.println("file_name:"+file_name);
     		  mySmartUpload.setContentDisposition(null);       
     		  mySmartUpload.downloadFile(downloadlink,"application/vnd.ms-powerpoint","download.pps");
       }
       if(file_type.equals("zip"))
       {
     		  System.out.println("file_name:"+file_name);
     		  mySmartUpload.setContentDisposition("inline;");       
     		  mySmartUpload.downloadFile(downloadlink,"application/x-zip-compressed","download.zip");      
       }
       if(file_type.equals("rar"))
       {
     		  System.out.println("file_name:"+file_name);
     		  mySmartUpload.setContentDisposition("inline;");       
     		  mySmartUpload.downloadFile(downloadlink,"application/x-zip-compressed","download.rar");      
       }
       //102.12.03 add pdf 下載 by 2295
       if(file_type.equals("pdf"))
       {
     		  System.out.println("file_name:"+file_name);
     		  mySmartUpload.setContentDisposition("inline;");       
     		  mySmartUpload.downloadFile(downloadlink,"application/pdf","download.pdf");      
       }
       
      
%>
 



