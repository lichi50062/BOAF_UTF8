<%
//97.09.22 fix 不砍掉.zip檔.留做備份(將.zip檔一併上傳至主機) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.io.BufferedOutputStream" %>
<%@ page import="java.io.File" %>
<%@ page import="java.io.FileInputStream" %>
<%@ page import="java.io.FileNotFoundException" %>
<%@ page import="java.io.FileOutputStream" %>
<%@ page import="java.io.IOException" %>
<%@ page import="java.io.PrintStream" %>
<%@ page import="com.oreilly.servlet.MultipartRequest" %>
<%@ page import="UploadHtml.Unzip"%>
<%@ page import="UploadHtml.UploadHtml"%>

<%
File logfile;
FileOutputStream logos=null;    	
BufferedOutputStream logbos = null;
PrintStream logps = null;
Date nowlog = new Date();
SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	
SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
SimpleDateFormat pwdformat = new SimpleDateFormat("MMdd");
Calendar logcalendar;
RequestDispatcher rd = null;
String actMsg = "";	
String UserID = ( request.getParameter("UserID")==null ) ? "" : (String)request.getParameter("UserID");		
String UserPWD = ( request.getParameter("UserPWD")==null ) ? "" : (String)request.getParameter("UserPWD");		
	
try{
    //正式的logfile = new File("/APSCGF/JAVA/uploadHtml.log");	 	   
	//正式的logos = new FileOutputStream("/APSCGF/JAVA/uploadHtml.log",false);  
	System.out.println("c:"+System.getProperty("file.separator")+"test2"+System.getProperty("file.separator")+"uploadHtml.log");		   
	logfile = new File("c:"+System.getProperty("file.separator")+"test2"+System.getProperty("file.separator")+"uploadHtml.log");	 	   
	logos = new FileOutputStream("c:"+System.getProperty("file.separator")+"test2"+System.getProperty("file.separator")+"uploadHtml.log",false);       	   
	logbos = new BufferedOutputStream(logos);
	logps = new PrintStream(logbos);		   
	logcalendar = Calendar.getInstance();
	nowlog = logcalendar.getTime();
	String pwd = pwdformat.format(nowlog);				    	
	logps.println(logformat.format(nowlog)+"============file upload begin===================");		
	logps.flush();	  
	//System.out.println("pwd="+pwd);
	//System.out.println("UserPWD="+UserPWD);
	//System.out.println("UserID="+UserID);
	if(UserID.equals("smeg") && UserPWD.equals(pwd)){  	  	      
       File tmpFile = new File(getProperties("homeDir"));
       if(tmpFile.exists()){
          System.out.println(getProperties("homeDir")+" is exists");
          boolean deleteOK = Unzip.deleteUploadDir();
          System.out.println("delete uploadDir="+deleteOK);
          nowlog = logcalendar.getTime();			    	
	      logps.println(logformat.format(nowlog)+"delete uploadDir && subdir ??"+deleteOK);		    					    
	      logps.flush();	
       }else{
          System.out.println(getProperties("homeDir")+"is not exists");
       }   
       
	   int UploadSize = Integer.parseInt(getProperties("UploadSize"));
	   MultipartRequest multi = new MultipartRequest(request,getProperties("homeDir"), UploadSize * 1024); 
	   nowlog = logcalendar.getTime();			    	 
	   logps.println(logformat.format(nowlog)+multi.getFilesystemName("UpFileName")+" upload success");		    					     
	   logps.flush();	
	   System.out.println("jsp.filename="+multi.getFilesystemName("UpFileName")); 
	   actMsg = Unzip.exeUnZIP(multi.getFilesystemName("UpFileName"));   
	   System.out.println("actMsg="+actMsg); 
	   nowlog = logcalendar.getTime();			    	
	   logps.println(logformat.format(nowlog)+"unzip file ??"+actMsg);		    					    
	   logps.flush();
	   if(actMsg.equals("unzip complete")){
	      actMsg = multi.getFilesystemName("UpFileName")+"檔案上傳完成";
	   }	
	   //96.05.18 解壓縮完成,先砍掉.zip檔,再執行上傳程式
	   //97.09.22 不砍掉.zip檔.留做備份 =========================================================================================
	   //tmpFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+multi.getFilesystemName("UpFileName"));
	   //if(tmpFile.exists()) tmpFile.delete();  
	   //======================================================================================================================== 
	   //96.05.18 產生flag.txt
	   File updateSMBCGFFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+"flag.txt");
	   System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+"flag.txt??create "+updateSMBCGFFile.createNewFile());	   
	  
	   UploadHtml uploadHtml = new UploadHtml();//移至.sct
	   actMsg = uploadHtml.exeUploadHtml(multi.getFilesystemName("UpFileName"));	
	  
	}else{
	   nowlog = logcalendar.getTime();				    	
	   logps.println(logformat.format(nowlog)+"ID/PWD is not correct!!");		
	   logps.flush();
	   actMsg="帳號/密碼不正確";	 
	}
	//rd = application.getRequestDispatcher("/FileUpload_Edit.jsp?test=nothing");//正式的
	rd = application.getRequestDispatcher("/pages/FileUpload_Edit.jsp?test=nothing");//測試的
	request.setAttribute("actMsg",actMsg);
	try {
 	    //forward to next present jsp
  	    rd.forward(request, response);
	} catch (NullPointerException npe) {
	   out.println(npe.getMessage());
	}
}catch(Exception e){
  out.println("FileUpload.Error:"+e+e.getMessage());
}finally{
		try{
			   if (logos  != null) logos.close();
 	           if (logbos != null) logbos.close();
 	           if (logps  != null) logps.close(); 	           
		}catch(Exception ioe){
				System.out.println(ioe.getMessage());
		}
}	
%>
<%!
private String getProperties(String key) throws Exception{
    //正式的String dirpath = "/APSCGF/JAVA/htmlDir.properties";
    String dirpath = "D:\\workProject\\BOAF\\WEB-INF\\classes\\UploadHtml\\htmlDir.properties";
  	String value	= "";
  	try {
  		Properties p = new Properties();
  		p.load(new FileInputStream(dirpath));
  		value = (String) p.get(key);
  	}catch (FileNotFoundException e) {
  		throw new Exception("FileUpload.getProperties:FileNotFoundException["+e.getMessage()+"]");
  	}catch (IOException e) {
  		throw new Exception("FileUpload.getProperties:IOException["+e.getMessage()+"]");
  	}
  	
  	return value;
}
/******************************************************************8
* 檢查檔案是否存在該目錄下
*/
private boolean CheckFileExist(String UploadDir,String UploadFileName){//檢查檔案是否存在該目錄下===
		String tmpFileName = "";
		if(UploadFileName.lastIndexOf("\\") != -1){
		   tmpFileName = UploadFileName.substring(UploadFileName.lastIndexOf("\\")+1,UploadFileName.length());
		}
		System.out.println("tmpFileName="+tmpFileName);
		File tmpFile = new File(UploadDir+System.getProperty("file.separator")+tmpFileName);
		System.out.println(UploadDir+System.getProperty("file.separator")+tmpFileName);
		if(tmpFile.exists()){
			return true;
		}
		
    	return false;		
}
%>
