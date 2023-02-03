<%
//103.05.27 add 上傳A01/A02/A03/A04/A05/WLX01/WLX01_M/WLX02/WLX02_M至金庫  by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ftp.AgriBankFTP" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="com.tradevan.util.sftp.*" %>
<%@ page import="com.tradevan.util.ftp.*" %>
<%
 
  //全國農業金庫主機
  //String FTP_SERVER = "www.agribank.com.tw";
  String FTP_SERVER = "59.124.54.12";	
  //String FTP_SERVER = "59.124.54.11";
  String FTP_USER = "TNBOAF";
  String FTP_PASSWORD = "TNBOAF123";
  //String FTP_DIRECTORY = "/TNBOAF/";
  String FTP_DIRECTORY = "";
  String LOCAL_DIRECTORY = "C:\\Sun\\WebServer6.1\\BOAF\\AgriBankData";
  
  /*
  //公司測試主機
  String FTP_SERVER = "10.89.8.170";
  String FTP_USER = "pdntmgr";
  String FTP_PASSWORD = "AAllen6812";
  String FTP_DIRECTORY = "APBOAF";	  
  String LOCAL_DIRECTORY = "D:\\workProject\\BOAF\\AgriBankData";
  */
  String alertMsg = "";	
  
  String sqlCmd = "";
  List dbData = new LinkedList();
 
  FileOutputStream logos=null;    	
  BufferedOutputStream logbos = null;
  PrintStream logps = null;    
  File logDir = null;   
  File logfile; 
  Date nowlog = new Date();
  SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
  SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
  SimpleDateFormat zipfileformat = new SimpleDateFormat("yyyyMMdd");
  Calendar logcalendar;
  
  FileOutputStream wmos=null;    	
  BufferedOutputStream wmbos = null;
  PrintStream wmps = null;   
  File AgriBankDir = null;      
  File wmfile; 
  String filename="";
  List filename_List = new LinkedList();
  String[] wm_idx = {"A01","A02","A03","A04","A05","WLX01","WLX01_M","WLX02","WLX02_M","WLX04"};
 
  List paramList = new ArrayList();
  String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");		
  String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");		
  //申報上個月份的報表
  Calendar now = Calendar.getInstance();
  String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
  String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
  if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
      YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
      MONTH = "12";	
  }else{    
      MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
  }
    
  if(!S_YEAR.equals("") && !S_MONTH.equals("")){
     YEAR = S_YEAR;
     MONTH = S_MONTH;
  }
  
  logDir  = new File(Utility.getProperties("logDir"));
  if(!logDir.exists()){
     if(!Utility.mkdirs(Utility.getProperties("logDir"))){
   	  System.out.println("目錄新增失敗");
     }    
  }
   
  logfile = new File(logDir + System.getProperty("file.separator") +"AgriBankWM."+ logfileformat.format(nowlog));						 
  System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + "AgriBankWM."+ logfileformat.format(nowlog));
  logos = new FileOutputStream(logfile,true);  		        	   
  logbos = new BufferedOutputStream(logos);
  logps = new PrintStream(logbos);			          
  logcalendar = Calendar.getInstance(); 
  nowlog = logcalendar.getTime();			    	
  logps.println(logformat.format(nowlog)+" "+"============執行上傳至金庫轉檔程式開始============");		    					    
  logps.flush();
 
  for(int i=0;i<wm_idx.length;i++){
     System.out.println(wm_idx[i]+" generate begin");
     filename = wm_idx[i]+".txt";
     wmfile = new File(Utility.getProperties("AgriBankDir") + System.getProperty("file.separator")+filename);		
     filename_List.add(filename);				 
     if(wmfile.exists()){
        System.out.println(logfile.getAbsoluteFile()+" is exists");
 	    wmfile.delete(); 	      
 	 }	
  	 dbData = getData(wm_idx[i],YEAR,MONTH);
  	 if(dbData != null && dbData.size() > 0){
  	 	logcalendar = Calendar.getInstance(); 
 		nowlog = logcalendar.getTime();		
        logps.println(logformat.format(nowlog)+" "+"取得"+wm_idx[i]+"資料完成,筆數="+dbData.size());		    					    
        logps.flush();
  	 }
  	  	 
     alertMsg = createfile(wm_idx[i],filename,dbData);
     logcalendar = Calendar.getInstance(); 
     nowlog = logcalendar.getTime();		
     if(alertMsg.indexOf("Error") == -1){
        logps.println(logformat.format(nowlog)+" "+"寫入"+filename+"檔案完成");		    					           
     }else{
        logps.println(logformat.format(nowlog)+" "+"產生檔案失敗"+alertMsg);	
     }	
     logps.flush();	
     System.out.println(wm_idx[i]+" generate end");
  }
  boolean uploadSuccess = false;
  filename = "BOAF"+zipfileformat.format(nowlog)+".ZIP";
  uploadSuccess = Utility.createZipFile(Utility.getProperties("AgriBankDir"),filename_List,filename);
  logcalendar = Calendar.getInstance(); 
  nowlog = logcalendar.getTime();		
  logps.println(logformat.format(nowlog)+" "+filename+(uploadSuccess==true?"檔案zip完成":"檔案zip失敗"));		    					    
  logps.flush();
  
  MySFTPClient msftp = new MySFTPClient(FTP_SERVER, FTP_USER,FTP_PASSWORD); 
  uploadSuccess = false;
  uploadSuccess = msftp.sendMyFiles(FTP_DIRECTORY, LOCAL_DIRECTORY, filename);
  
  logcalendar = Calendar.getInstance(); 
  nowlog = logcalendar.getTime();		
  logps.println(logformat.format(nowlog)+" "+filename+(uploadSuccess==true?"檔案上傳完成":"檔案上傳失敗"));		    					    
  logps.flush();
 
  logcalendar = Calendar.getInstance(); 
  nowlog = logcalendar.getTime();		
  logps.println(logformat.format(nowlog)+" "+"============執行上傳至金庫轉檔程式結束============");		    					    
  logps.flush();
%>
<%!
private List getData(String report_no,String m_year,String m_month){    
    StringBuffer sqlCmd = new StringBuffer();
    List paramList = new ArrayList();//傳內的參數List   
    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
    if("A01".equals(report_no)|| "A02".equals(report_no) || "A03".equals(report_no) || "A04".equals(report_no) || "A05".equals(report_no)){
       sqlCmd.append(" select "+report_no+".m_year||'|'||m_month||'|'||bank_code||'|'||acc_code||'|'||amt");
       if("A02".equals(report_no)|| "A05".equals(report_no)){
       	sqlCmd.append("||'|'||amt_name");
       }	
       sqlCmd.append(" as data");
       sqlCmd.append(" from "+report_no+" left join (select * from bn01 where m_year=?)bn01 on "+report_no+".bank_code=bn01.bank_no");
       sqlCmd.append(" where "+report_no+".m_year = ?");
       sqlCmd.append(" and  "+report_no+".m_month = ?");
       paramList.add(wlx01_m_year);       
       paramList.add(m_year); 
       paramList.add(m_month); 
    }else if("WLX01".equals(report_no)){    	
       sqlCmd.append(" select BANK_NO||");
       sqlCmd.append("'|'||ENGLISH||");
       sqlCmd.append("'|'||SETUP_APPROVAL_UNT||");
       sqlCmd.append("'|'||to_char(SETUP_DATE,'MM/DD/YYYY HH:MI:SS AM')||");
       sqlCmd.append("'|'||SETUP_NO||");
       sqlCmd.append("'|'||to_char(CHG_LICENSE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||CHG_LICENSE_NO||");
       sqlCmd.append("'|'||CHG_LICENSE_REASON||");
       sqlCmd.append("'|'||to_char(START_DATE,'MM/DD/YYYY HH:MI:SS AM')||");
       sqlCmd.append("'|'||BUSINESS_ID||");
       sqlCmd.append("'|'||HSIEN_ID||");
       sqlCmd.append("'|'||AREA_ID||");
       sqlCmd.append("'|'||ADDR||");
       sqlCmd.append("'|'||TELNO||");
       sqlCmd.append("'|'||FAX||");
       sqlCmd.append("'|'||EMAIL||");
       sqlCmd.append("'|'||WEB_SITE||");
       sqlCmd.append("'|'||CENTER_FLAG||");
       sqlCmd.append("'|'||CENTER_NO||");
       sqlCmd.append("'|'||STAFF_NUM||");
       sqlCmd.append("'|'||IT_HSIEN_ID||");
       sqlCmd.append("'|'||IT_AREA_ID||");
       sqlCmd.append("'|'||IT_ADDR||");
       sqlCmd.append("'|'||IT_NAME||");
       sqlCmd.append("'|'||IT_TELNO||");
       sqlCmd.append("'|'||AUDIT_HSIEN_ID||");
       sqlCmd.append("'|'||AUDIT_AREA_ID||");
       sqlCmd.append("'|'||AUDIT_ADDR||");
       sqlCmd.append("'|'||AUDIT_NAME||");
       sqlCmd.append("'|'||AUDIT_TELNO||");
       sqlCmd.append("'|'||FLAG||");
       sqlCmd.append("'|'||to_char(OPEN_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||M2_NAME||");
       sqlCmd.append("'|'||Hsien_div_1||");
       sqlCmd.append("'|'||CANCEL_NO||");
       sqlCmd.append("'|'||to_char(CANCEL_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||USER_ID||");
       sqlCmd.append("'|'||USER_NAME||");
       sqlCmd.append("'|'||to_char(update_date,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||credit_staff_num ||");
       sqlCmd.append("'|'||m_year");
       sqlCmd.append(" as data");
       sqlCmd.append(" from WLX01");  
       sqlCmd.append(" where m_year=?");
       sqlCmd.append(" order by bank_no");
       paramList.add(wlx01_m_year);      
     }else if("WLX01_M".equals(report_no)){    	
       sqlCmd.append(" select BANK_NO||");
       sqlCmd.append("'|'||POSITION_CODE||");
       sqlCmd.append("'|'||ID||");
       sqlCmd.append("'|'||NAME||");
       sqlCmd.append("'|'||to_char(BIRTH_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||DEGREE||");
       sqlCmd.append("'|'||SEX||");
       sqlCmd.append("'|'||TELNO||");
       sqlCmd.append("'|'||FAX||");
       sqlCmd.append("'|'||to_char(INDUCT_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||BACKGROUND||");
       sqlCmd.append("'|'||CHOOSE_ITEM||");
       sqlCmd.append("'|'||RANK||");
       sqlCmd.append("'|'||SPECIALITY||");
       sqlCmd.append("'|'||INCHARGE||");
       sqlCmd.append("'|'||ABDICATE_CODE||");
       sqlCmd.append("'|'||to_char(ABDICATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||EMAIL||");
       sqlCmd.append("'|'||USER_ID||");
       sqlCmd.append("'|'||USER_NAME||");
       sqlCmd.append("'|'||to_char(UPDATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||SEQ_NO||");
       sqlCmd.append("'|'||ID_CODE");
       sqlCmd.append(" as data");
       sqlCmd.append(" from WLX01_M");
       sqlCmd.append(" order by bank_no,position_code");
    }else if("WLX02".equals(report_no)){      
       sqlCmd.append(" select TBANK_NO||");
       sqlCmd.append("'|'||BANK_NO||");
       sqlCmd.append("'|'||CONST_TYPE||");
       sqlCmd.append("'|'||SETUP_APPROVAL_UNT||");
       sqlCmd.append("'|'||to_char(SETUP_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||SETUP_NO||");
       sqlCmd.append("'|'||to_char(SETUP_NO_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||to_char(CHG_LICENSE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||CHG_LICENSE_NO||");
       sqlCmd.append("'|'||CHG_LICENSE_REASON||");
       sqlCmd.append("'|'||to_char(START_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||HSIEN_ID||");
       sqlCmd.append("'|'||AREA_ID||");
       sqlCmd.append("'|'||ADDR||");
       sqlCmd.append("'|'||TELNO||");
       sqlCmd.append("'|'||FAX||");
       sqlCmd.append("'|'||EMAIL||");
       sqlCmd.append("'|'||WEB_SITE||");
       sqlCmd.append("'|'||FLAG||");
       sqlCmd.append("'|'||to_char(OPEN_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||STAFF_NUM||");
       sqlCmd.append("'|'||HSIEN_DIV_1||");
       sqlCmd.append("'|'||CANCEL_NO||");
       sqlCmd.append("'|'||to_char(CANCEL_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||USER_ID||");
       sqlCmd.append("'|'||USER_NAME||");
       sqlCmd.append("'|'||to_char(UPDATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||m_year");
       sqlCmd.append(" as data");
       sqlCmd.append(" from WLX02");
       sqlCmd.append(" where m_year=?");
       sqlCmd.append(" order by tbank_no,bank_no");
       paramList.add(wlx01_m_year);    
    }else if("WLX02_M".equals(report_no)){  
       sqlCmd.append(" select BANK_NO||");                                             
       sqlCmd.append("'|'||POSITION_CODE||");                                    
       sqlCmd.append("'|'||ID||");                                               
       sqlCmd.append("'|'||NAME||");                                             
       sqlCmd.append("'|'||to_char(BIRTH_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");    
       sqlCmd.append("'|'||DEGREE||");                                           
       sqlCmd.append("'|'||SEX||");                                              
       sqlCmd.append("'|'||TELNO||");                                            
       sqlCmd.append("'|'||to_char(INDUCT_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");  
       sqlCmd.append("'|'||BACKGROUND||");                                      
       sqlCmd.append("'|'||CHOOSE_ITEM||");                                     
       sqlCmd.append("'|'||RANK||");                                            
       sqlCmd.append("'|'||ABDICATE_CODE||");                                   
       sqlCmd.append("'|'||to_char(ABDICATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");
       sqlCmd.append("'|'||EMAIL||");                                           
       sqlCmd.append("'|'||USER_ID||");                                         
       sqlCmd.append("'|'||USER_NAME||");                                       
       sqlCmd.append("'|'||to_char(UPDATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");  
       sqlCmd.append("'|'||SEQ_NO||");                                          
       sqlCmd.append("'|'||ID_CODE"); 
       sqlCmd.append(" as data");
       sqlCmd.append(" from WLX02_M");
       sqlCmd.append(" order by bank_no,position_code");
    }else if("WLX04".equals(report_no)){    
       sqlCmd.append(" select BANK_NO||");                                       
       sqlCmd.append("'|'||POSITION_CODE||");                                    
       sqlCmd.append("'|'||ID||");                                               
       sqlCmd.append("'|'||NAME||");                                             
       sqlCmd.append("'|'||to_char(BIRTH_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");    
       sqlCmd.append("'|'||RANK||");                                             
       sqlCmd.append("'|'||PASSPORT_AREA||");                                    
       sqlCmd.append("'|'||PASSPORT_NO||");                                      
       sqlCmd.append("'|'||to_char(INDUCT_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");   
       sqlCmd.append("'|'||ABDICATE_CODE||");                                    
       sqlCmd.append("'|'||to_char(ABDICATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||"); 
       sqlCmd.append("'|'||APPOINTED_NUM||");                                    
       sqlCmd.append("'|'||PERIOD_START||");                                     
       sqlCmd.append("'|'||PERIOD_END||");                                       
       sqlCmd.append("'|'||SEX||");                                              
       sqlCmd.append("'|'||DEGREE||");                                           
       sqlCmd.append("'|'||BACKGROUND||");                                       
       sqlCmd.append("'|'||PROFESSIONAL||");                                     
       sqlCmd.append("'|'||TELNO||");                                            
       sqlCmd.append("'|'||FINANCE_EXP||");                                      
       sqlCmd.append("'|'||EMAIL||");                                            
       sqlCmd.append("'|'||USER_ID||");                                          
       sqlCmd.append("'|'||USER_NAME||");                                        
       sqlCmd.append("'|'||to_char(UPDATE_DATE,'MM/DD/YYYY HH:MI:SS AM') ||");   
       sqlCmd.append("'|'||SEQ_NO||");                                           
       sqlCmd.append("'|'||ID_CODE");                                            
       sqlCmd.append(" as data");                                 
       sqlCmd.append(" from WLX04");                            
       sqlCmd.append(" order by bank_no,position_code");    
    }	    
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");   
    return dbData;
}      
private String createfile(String report_no,String filename,List dbData){  
	FileOutputStream logos=null;    	
    BufferedOutputStream logbos = null;
    PrintStream logps = null;    
    File AgriBankDir = null;   
    File logfile; 
    String errMsg="";
	try{
       
       AgriBankDir  = new File(Utility.getProperties("AgriBankDir"));
 	   if(!AgriBankDir.exists()){
           if(!Utility.mkdirs(Utility.getProperties("AgriBankDir"))){
         	   System.out.println("目錄新增失敗");
           }    
       }
       logfile = new File(AgriBankDir + System.getProperty("file.separator") +filename);
 	   logos = new FileOutputStream(logfile,false);  		        	   
 	   logbos = new BufferedOutputStream(logos);
 	   logps = new PrintStream(logbos);	
 	   for(int i=0;i<dbData.size();i++){	 	   	     	   	  
           if(i==dbData.size()-1){
            logps.print((String)((DataObject)dbData.get(i)).getValue("data"));		
           }else{
             logps.println((String)((DataObject)dbData.get(i)).getValue("data"));		
           }
	   }    					    
 	   logps.flush(); 	 	   
 	   
 	}catch (Exception e){
   	   System.out.println(e+":"+e.getMessage());  
   	   errMsg = report_no+" createfile Error "+e+":"+e.getMessage();
   	} 
   	return errMsg;
}	
  
%>















