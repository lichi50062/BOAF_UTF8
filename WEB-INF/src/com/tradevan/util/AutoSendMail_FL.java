/*
 * 105.10.20 create 未結案案件逾期25日未回文者e-mail通知農金局承辦人 by 2295
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



public class AutoSendMail_FL{
    static final String driverName = "oracle.jdbc.driver.OracleDriver";
    static String dbURL = ""; 
    static String userName = ""; 
    static String userPwd = "";     
    static Connection dbConn = null;
    static ResultSet rs = null;
    static PreparedStatement pstmt = null;
    static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    static String sendMsg = "";
	public static void main(String[] args){
		System.out.println("AutoSendMail_FL begin");
		AutoSendMail_FL a = new AutoSendMail_FL();
		a.sendMail();
		System.out.println("AutoSendMail_FL end");
	}


	public static void sendMail()
	{	 
	 String sqlCmd="";	   
	 
	 try{	    
	    logDir  = new File(Utility.getProperties("logDir"));
        if(!logDir.exists()){
            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
               System.out.println("目錄新增失敗");
            }    
        }
        logfile = new File(logDir + System.getProperty("file.separator") + "AutoSendMail_FL."+ logfileformat.format(nowlog));                       
        System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"AutoSendMail_FL."+ logfileformat.format(nowlog));
        logos = new FileOutputStream(logfile,true);                         
        logbos = new BufferedOutputStream(logos);
        logps = new PrintStream(logbos);   
	    
	    dbURL = Utility.getProperties("BOAFDBURL");
        userName = Utility.getProperties("rptID");
        userPwd = Utility.getProperties("rptPwd");
        Class.forName(driverName); 
        dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
        System.out.println("Connection Successful!");
         
	    sqlCmd = " select frm_snrtdoc.bank_no,"//--農漁會別.機構代號
	           + " bank_name,"//--農漁會別.機構名稱
	           + " frm_exmaster.ex_type,"
	           + " decode(frm_exmaster.ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,"//--查核類別
	           + " frm_snrtdoc.ex_no, "//--查核報告編號
	           + " decode(ex_type,'FEB',frm_snrtdoc.ex_no,'AGRI',substr(frm_snrtdoc.ex_no,0,3)||'年第'||substr(frm_snrtdoc.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_snrtdoc.ex_no,'yyyymmdd')),'') as ex_no_list, "//--報表需顯示的檢查報告編號或查核季別或訪查日期
	           + " F_TRANSCHINESEDATE(doc_date) as doc_date,"//--農金局發文.日期
	           + " docno,"//--農金局發文.文號 
	           + " doc_type,"//--發文性質代碼 A:陳述意見/B:核處/C:結案
	           + " cmuse_name as doc_type_name, "//--農金局發文.發文性質名稱
	           + " F_TRANSCHINESEDATE(limitdate) as limitdate,"//--應回覆日期
	           + " wlx01.telno,"//--電話
	           + " F_COMBWLX01_MNAME(frm_snrtdoc.bank_no,'4') as mname, "//--信用部主任
	           + " F_COMBWLX01_AuditNAME(frm_snrtdoc.bank_no) as audit_name, "//--稽核
	           + " m_email,"//--承辦人email
	           + " limitdate+25 as over25day "//--限期函報日+25天   
	           + " from frm_snrtdoc "
	           + " left join (select * from bn01 where m_year=100)bn01 on frm_snrtdoc.bank_no=bn01.bank_no"
	           + " left join (select * from cdshareno where cmuse_div='048')cdshareno on frm_snrtdoc.doc_type=cdshareno.cmuse_id "
	           + " left join frm_exmaster on frm_snrtdoc.ex_no=frm_exmaster.ex_no "
	           + " left join (select * from wlx01 where m_year=100)wlx01 on frm_snrtdoc.bank_no=wlx01.bank_no "
	           + " left join muser_data on frm_snrtdoc.user_id = muser_data.muser_id "
	           + " where doc_type='B'"//--核處
	           + " and   frm_exmaster.case_status !='0'"// --未結案
	           + " and audit_id in ('A1','B1') "//--限期改善具報
	           + " and bank_rt_docno is null "
	           + " and sysdate  > limitdate+25 "//--超過限期函報日(25天)
	           + " and m_email is not null "
	           + " group by frm_snrtdoc.bank_no,bank_name,frm_exmaster.ex_type,frm_snrtdoc.ex_no,frm_snrtdoc.ex_no,doc_date,docno,doc_type,cmuse_name,limitdate,wlx01.telno,frm_snrtdoc.bank_no,m_email,limitdate "
	           + " order by m_email,frm_snrtdoc.bank_no asc,ex_no asc,doc_date desc ";
	     
	    pstmt = dbConn.prepareStatement(sqlCmd);
	   
        ResultSet rs = pstmt.executeQuery();
        System.out.println("sendmail.size()="+rs.getFetchSize());
        
        String m_email = "";
         
        while (rs.next()) {          
            /*
            System.out.println("bank_no="+rs.getString("bank_no"));
            System.out.println("bank_name="+rs.getString("bank_name"));
            System.out.println("ex_type="+rs.getString("ex_type"));
            System.out.println("ex_type_name="+rs.getString("ex_type_name"));
            System.out.println("ex_no="+rs.getString("ex_no"));
            System.out.println("ex_no_list="+rs.getString("ex_no_list"));
            System.out.println("doc_date="+rs.getString("doc_date"));
            System.out.println("docno="+rs.getString("docno"));
            System.out.println("limitdate="+rs.getString("limitdate"));
            System.out.println("telno="+rs.getString("telno"));
            System.out.println("m_email="+rs.getString("m_email"));
            */
            if(!m_email.equals("") && !rs.getString("m_email").equals(m_email)){
                sendMsg = "農漁會別                                 查核類別          查核報告編號    農金局發文日期   農金局發文文號  應回覆日期       電話\n"
                        + "==================================================================================================\n"
                        + sendMsg;   
                Utility.sendMail(m_email,"專案農貸檢查追蹤管理系統-未結案且逾期25日未回文資料",sendMsg);
                printLog(logps,"mail to "+m_email);
                printLog(logps,sendMsg);
                sendMsg="";
                sendMsg += rs.getString("bank_no")+rs.getString("bank_name")+" "
                        + rs.getString("ex_type_name")+" "
                        +("AGRI".equals(rs.getString("ex_type"))?"    ":"")
                        +("BOAF".equals(rs.getString("ex_type"))?"        ":"")
                        + rs.getString("ex_no_list")+"  "
                        +("AGRI".equals(rs.getString("ex_type"))?"    ":"")
                        +("FEB".equals(rs.getString("ex_type"))?"           ":"")
                        + rs.getString("doc_date")+"    "
                        + rs.getString("docno")+"          "
                        + rs.getString("limitdate")+"  "
                        + rs.getString("telno")+"\n";            
                m_email= rs.getString("m_email");
            }else{
                if(m_email.equals("")){
                    m_email= rs.getString("m_email");
                }
                sendMsg += rs.getString("bank_no")+rs.getString("bank_name")+" "
                        + rs.getString("ex_type_name")+" "
                        +("AGRI".equals(rs.getString("ex_type"))?"    ":"")
                        +("BOAF".equals(rs.getString("ex_type"))?"        ":"")
                        + rs.getString("ex_no_list")+"  "
                        +("AGRI".equals(rs.getString("ex_type"))?"    ":"")
                        +("FEB".equals(rs.getString("ex_type"))?"           ":"")
                        + rs.getString("doc_date")+"    "
                        + rs.getString("docno")+"          "
                        + rs.getString("limitdate")+"  "
                        + rs.getString("telno")+"\n";                
            }
        }
        
        //寄送email
        sendMsg = "農漁會別                                 查核類別          查核報告編號    農金局發文日期   農金局發文文號  應回覆日期       電話\n"
                + "==================================================================================================\n"
                + sendMsg;   
        Utility.sendMail(m_email,"專案農貸檢查追蹤管理系統-未結案且逾期25日未回文資料",sendMsg);
        printLog(logps,"mail to "+m_email);
        printLog(logps,sendMsg);
       
        if(pstmt != null){
            pstmt.close();
            pstmt = null;             
        }
        if(!dbConn.isClosed()){
            dbConn.close();
            dbConn = null;
        }
     }catch (Exception e){
	    System.out.println(e.getMessage());
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
}



