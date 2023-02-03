<%
// 95.08.21 add 金額單位 by 2295
// 99.04.28 fix request改以dataMap存取 by 2808
//105.03.22 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.FR008W_Excel" %>
<%
   Map dataMap =Utility.saveSearchParameter(request);
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
   String act = Utility.getTrimString(dataMap.get("act")) ;
   String bank_type = Utility.getTrimString(dataMap.get("bank_type")) ;
   String BANK_DATA = Utility.getTrimString(dataMap.get("BANK_NO")) ;
   String BANK_NO = BANK_DATA.substring(0,BANK_DATA.indexOf("/"));
   String BANK_NAME = BANK_DATA.substring(BANK_DATA.indexOf("/")+1,BANK_DATA.length());
   String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR")) ; 
   String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH")) ;
   String Unit = dataMap.get("Unit")==null?"1" : (String)dataMap.get("Unit") ;			  
   String excelAction =  Utility.getTrimString(dataMap.get("excelaction")) ; 
   if(excelAction.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
      response.setHeader("Content-disposition","inline; filename=view.xls");
   }else if (excelAction.equals("download")){
      response.setHeader("Content-Disposition","attachment; filename=download.xls");
   }

%>
<%
	try{
	    String actMsg = FR008W_Excel.createRpt(S_YEAR,S_MONTH,BANK_NO,BANK_NAME,bank_type,Unit);
	    System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部淨值占風險性資產比率.xls");
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部淨值占風險性資產比率.xls");
		ServletOutputStream out1 = response.getOutputStream();
		byte[] line = new byte[8192];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8192)))!=-1 ){
			out1.write(line,0,getBytes);
			out1.flush();
	    }

		fin.close();
		out1.close();

	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
%>