/*
 * 105.01.08 create 首頁公告區資料發佈 by 2295
 */
package com.tradevan.util;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;



public class AutoGenNotify{
    static final String driverName = "oracle.jdbc.driver.OracleDriver";
    static String dbURL = ""; 
    static String userName = ""; 
    static String userPwd = "";     
    static Connection dbConn = null;
    static ResultSet rs = null;
    static PreparedStatement pstmt = null;
    static File logfile,notifyfile;
    static FileOutputStream logos,notifyos=null;      
    static BufferedOutputStream logbos,notifybos = null;
    static PrintStream logps,notifyps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir,notifyDir = null;
	public static void main(String[] args){
		System.out.println("AutoGenNotify begin");
		AutoGenNotify a = new AutoGenNotify();
		a.generate();
		System.out.println("AutoGenNotify end");
	}


	public static void generate()
	{	 
	 String sqlCmd="";	   
	 
	 try{	    
	    logDir  = new File(Utility.getProperties("logDir"));
        if(!logDir.exists()){
            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
               System.out.println("目錄新增失敗");
            }    
        }
        logfile = new File(logDir + System.getProperty("file.separator") + "AutoGenNotify."+ logfileformat.format(nowlog));                       
        System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"AutoGenNotify."+ logfileformat.format(nowlog));
        logos = new FileOutputStream(logfile,true);                         
        logbos = new BufferedOutputStream(logos);
        logps = new PrintStream(logbos);   
       
	    
        notifyDir  = new File(Utility.getProperties("notifyPage"));       
        notifyfile = new File(notifyDir + System.getProperty("file.separator") + "Notify.html");  
        System.out.println("notifyfile="+notifyfile);
        
        notifyos = new FileOutputStream(notifyfile,false); 
        notifybos = new BufferedOutputStream(notifyos);        
        notifyps = new PrintStream(notifybos,true,"UTF-8");
        
        dbURL = Utility.getProperties("BOAFDBURL");
        userName = Utility.getProperties("rptID");
        userPwd = Utility.getProperties("rptPwd");
        Class.forName(driverName); 
        dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
        System.out.println("Connection Successful!");
        
        sqlCmd = "select seq_no,headmark,to_char(notify_date,'yyyy/mm/dd hh:mi') as notify_date,"
                + " to_char(notify_end_date,'yyyy/mm/dd') as notify_end_date,append_file,"
                + " user_id,user_name,to_char(update_date,'mm/dd/yyyy hh:mi:ss') as update_date,"
                + " appfile_link,notify_url from WLX_Notify "
                + " where to_char(notify_date,'yyyy/mm/dd hh') <= to_char(sysdate,'yyyy/mm/dd hh') "
                + " and to_char(notify_end_date,'yyyy/mm/dd hh') >= to_char(sysdate,'yyyy/mm/dd hh') "
                + " order by notify_date desc";  
	    pstmt = dbConn.prepareStatement(sqlCmd);
	    ResultSet rs = pstmt.executeQuery();
	    System.out.println("notify.size()="+rs.getFetchSize());
	    List list = new LinkedList();    
	    Map values = new HashMap();    
	    
	    while (rs.next()) {
	        values = new HashMap();    
	        System.out.println("append_file="+rs.getObject("append_file"));
	        System.out.println("notify_url="+rs.getObject("notify_url"));
	        System.out.println("notify_date="+rs.getObject("notify_date"));
	        System.out.println("headmark="+rs.getObject("headmark"));
	        System.out.println("seq_no="+rs.getObject("seq_no"));
	        values.put("append_file", rs.getObject("append_file"));
	        values.put("notify_url", rs.getObject("notify_url"));
	        values.put("notify_date", rs.getObject("notify_date"));
	        values.put("headmark", rs.getObject("headmark"));
	        values.put("seq_no", rs.getObject("seq_no"));
	        list.add(values);
        }
	    
	    if(list.size() > 0){	       
	       printLog(logps,"取得最新公告資料,共"+list.size()+"筆");
	       //printNotify(notifyps,"<%@ page language=\"java\" contentType=\"text/html; charset=UTF-8\" pageEncoding=\"UTF-8\"%>");
	       printNotify(notifyps,"<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/>");
	       printNotify(notifyps,"<tr colspan=0 bordercolor=#63B1AE bgcolor='#408C87'><img src=images/hotnews.gif width=70 height=18></tr>");  
	       printNotify(notifyps,"<script language=JavaScript>"); 
	       printNotify(notifyps,"document.write(\"<marquee  bgcolor=#70BEBB scrollamount='1' scrolldelay='6' direction= 'up' width='380' id=xiaoqing height='110' onmouseover=xiaoqing.stop() onmouseout=xiaoqing.start()>\"); ");
	       
            String notify_date = "";
            String fontColor="black";
            String headmark="";
            String seq_no="";
            String notify_url = ""; 
                   
            for(int i=0;i<list.size();i++){
                values = (HashMap)list.get(i);
                
                notify_date = (String)values.get("notify_date");
                notify_url = (String)values.get("notify_url");
                headmark = (String)values.get("headmark");
                seq_no = values.get("seq_no").toString();                
                
                notify_date = Utility.getCHTdate(notify_date.substring(0,10),0)+notify_date.substring(10,13)+":00";
                fontColor=((i == 0)?"white":"black");//最新一筆.以藍色顯示
                if(headmark.indexOf("異常") != -1) fontColor="yellow";//系統異常訊息.以黃色顯示
                
                if(values.get("append_file") != null && !((String)values.get("append_file")).trim().equals("")){//有附加檔案%>                    
                    printNotify(notifyps,"document.write(\"&nbsp;<a href=\\\"javascript:doLink('/pages/DomloadNotify.jsp?seq_no="+seq_no+"');\\\"><img src=images/download.gif border=0></a>&nbsp;<font size='2' \");");                    
                    if(notify_url != null){ //有URL%>
                        printNotify(notifyps,"document.write(\"color="+fontColor+">"+notify_date+"&nbsp;<a href=\\\"javascript:doLink('"+notify_url+"');\\\">"+headmark+"</a></font><br>\");");     
                    }else{//無URL%>
                        printNotify(notifyps,"document.write(\"color="+fontColor+">"+notify_date+"&nbsp;"+headmark+"</font><br>\");");   
                    }
                }else{//無附加檔案%>
                    printNotify(notifyps,"document.write(\"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size='2'\");");
                    if(notify_url != null){ //有URL%>
                       printNotify(notifyps,"document.write(\"color="+fontColor+">"+notify_date+"&nbsp;<a href=\\\"javascript:doLink('"+notify_url+"');\\\">"+headmark+"</a></font><br>\");"); 
                    }else{//無URL%>
                       printNotify(notifyps,"document.write(\"color="+fontColor+">"+notify_date+"&nbsp;"+headmark+"</font><br>\");");                                                             
                    }
                }//end of 無附加檔案 
            }//end of list            
              
            printNotify(notifyps,"document.write(\"</marquee>\");");               
            printNotify(notifyps,"</script>");
            printLog(logps,"完成產生公告顯示區");
	    }else{
           printLog(logps,"無最新公告資料");
           printNotify(notifyps,"<tr>"); 
           printNotify(notifyps,"<td width=12><img src=images/arrow_01.gif width=9 height=9 align=absmiddle></td>");
           printNotify(notifyps,"<td width=355 class=sbody>網際網路申報系統線上申請作業說明。</td>");
           printNotify(notifyps,"</tr>");
           printNotify(notifyps,"<tr>"); 
           printNotify(notifyps,"<td valign=top><img src=images/arrow_01.gif width=9 height=9 align=absmiddle></td>");
           printNotify(notifyps,"<td class=sbody>本系統全年全天候開放，惟須暫停連線服務時，將事先於本網站首頁公佈。</td>");
           printNotify(notifyps,"</tr>");
        }
       
        if(pstmt != null){
            pstmt.close();
            pstmt = null;             
        }
        if(rs != null){
            rs.close();
            rs = null;             
        }
        if(!dbConn.isClosed()){
            dbConn.close();
            dbConn = null;
        }
     }catch (Exception e){
	    System.out.println(e+e.getMessage());
     }
    }//end of generate
	public static void printLog(PrintStream logps,String errRptMsg){
	       if(!errRptMsg.equals("")){
	          logcalendar = Calendar.getInstance(); 
	          nowlog = logcalendar.getTime();
	          logps.println(logformat.format(nowlog)+errRptMsg);
	          logps.flush();
	       }
	}
	public static void printNotify(PrintStream notifyps,String notifyMsg){
        if(!notifyMsg.equals("")){
           notifyps.println(notifyMsg);
           notifyps.flush();
        }
    }
}



