<%
// 94.03.10	增加機構類別判斷及更改檔名FR006W為FR006W處理 add by egg
//100.06.24 add 套用saveSearchParameter取得參數 by 2295
//105.03.24 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>								          
<%@ page import="com.tradevan.util.report.FR006W_Excel" %>				
<%@include file="./include/Header.include" %>				          
<%
    String bank_type = Utility.getTrimString(dataMap.get("bank_type"));        
    String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));     
    String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));     
    String filename="信用部違反法定比率規定分析表-";
   	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
    
   	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
    	rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
	 	 //set next jsp
   	     if(act.equals("view")){
   	        //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
   	        //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
   	        response.setHeader("Content-disposition","inline; filename=view.xls");
   	     }else if (act.equals("download")){   
   	        response.setHeader("Content-Disposition","attachment; filename=download.xls");
   	     }
         if(act.equals("view") || act.equals("download")){    

	        try{
	            actMsg = FR006W_Excel.createRpt(S_YEAR,S_MONTH,bank_type);
	            System.out.println("createRpt="+actMsg);	            
	            System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename+(bank_type.equals("6")?"農會":bank_type.startsWith("B")?"農漁會":"漁會")+".xls");
	 			FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename+(bank_type.equals("6")?"農會":bank_type.startsWith("B")?"農漁會":"漁會")+".xls");	            
	        	ServletOutputStream out1 = response.getOutputStream();           
	        	byte[] line = new byte[8196];
	        	int getBytes=0;
	        	while( ((getBytes=fin.read(line,0,8196)))!=-1 ){		    		
	        		out1.write(line,0,getBytes);
	        		out1.flush();
	            }
	        	//關閉檔案
	        	fin.close();
	        	out1.close();            		      
	        	
	        }catch(Exception e){
	           System.out.println(e.getMessage());
	        }
         }
    }
   	request.setAttribute("actMsg",actMsg);		
%>
<%@include file="./include/Tail.include" %>
<%!
    private final static String report_no = "FR006W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
%>
