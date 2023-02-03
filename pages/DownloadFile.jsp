<%
// 98.08.06 add 金庫報表檔案下載 by 2295
//102.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.jspsmart.upload.*" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="java.io.*" %>
<%@ page import="java.net.URLEncoder" %>

<HTML>
<head>
<title>下載檔案</title>
</head>
</HTML>
<jsp:useBean  id="mySmartUpload"  scope="page"  class="com.jspsmart.upload.SmartUpload" />

<%  
  
    String append_file="";
    String rpt_code = ( request.getParameter("rpt_code")==null ) ? "" : (String)request.getParameter("rpt_code");
    String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
    String m_month = ( request.getParameter("m_month")==null ) ? "" : (String)request.getParameter("m_month");
    String rpt_version = ( request.getParameter("rpt_version")==null ) ? "" : (String)request.getParameter("rpt_version");
    
    List uploadfile = getUploadfile(rpt_code,m_year,m_month,rpt_version);
   
    DataObject bean = null;
    if(uploadfile != null && uploadfile.size() != 0){
        bean = (DataObject)uploadfile.get(0);                    
		append_file =(String)bean.getValue("rpt_fname");    	
		
		System.out.println("append_file="+append_file);  
		 	        		
	    String downloadlink = Utility.getProperties("ClientRptDir")+System.getProperty("file.separator")+append_file;	 
	    int index = append_file.lastIndexOf('.');      
        String file_type = append_file.substring(index+1,index+4);
        System.out.println("downloadlink:"+downloadlink);
        System.out.println("file_type:"+file_type);
        System.out.println("file_name:"+append_file);
        System.out.println("encode:"+URLEncoder.encode(append_file, "utf-8"));
        append_file = URLEncoder.encode(append_file, "utf-8");
	   
		mySmartUpload.initialize(pageContext);            
       
        if(file_type.equals("doc")){
           mySmartUpload.setContentDisposition(null);
           mySmartUpload.downloadFile(downloadlink,"application/msword; charset=UTF-8",append_file); 
        }else if(file_type.equals("xls")){          
           mySmartUpload.setContentDisposition(null);       
           mySmartUpload.downloadFile(downloadlink,"application/vnd.ms-excel; charset=UTF-8",append_file);       
        }else if(file_type.equals("pps")){           
           mySmartUpload.setContentDisposition(null);       
           mySmartUpload.downloadFile(downloadlink,"application/vnd.ms-powerpoint; charset=UTF-8",append_file);
        }else if(file_type.equals("zip")){           
           mySmartUpload.setContentDisposition("inline;");       
           mySmartUpload.downloadFile(downloadlink,"application/x-zip-compressed; charset=UTF-8",append_file);      
        }else if(file_type.equals("rar")){           
           mySmartUpload.setContentDisposition("inline;");       
           mySmartUpload.downloadFile(downloadlink,"application/x-zip-compressed; charset=UTF-8",append_file);      
        }else if(file_type.equals("pdf")){           
           mySmartUpload.setContentDisposition("inline;");       
           mySmartUpload.downloadFile(downloadlink,"application/pdf; charset=UTF-8",append_file);      
        }else if(file_type.equals("rtf")){           
           mySmartUpload.setContentDisposition("inline;");       
           mySmartUpload.downloadFile(downloadlink,"application/rtf; charset=UTF-8",append_file);      
        }   
        java.io.File tmpFile = new java.io.File(Utility.getProperties("ClientRptDir")+System.getProperty("file.separator")+append_file);
        System.out.println("tmpFile.exists()??"+tmpFile.exists());
		if(tmpFile.exists()) tmpFile.delete();
    }//end of have data          
%>

<%!
    
    private List getUploadfile(String rpt_code,String m_year,String m_month,String rpt_version){       
        List paramList = new ArrayList();
    	String sqlCmd = " select rpt_fname"
    				  + " from agrirpt_dirf"
					  + " where rpt_code=?"
					  + " and m_year=?"
					  + " and m_month=?"
					  + " and rpt_version=?";
    	paramList.add(rpt_code);
    	paramList.add(m_year);
    	paramList.add(m_month);
    	paramList.add(rpt_version);
    	
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
%>