/* 
 * 106.09.06-22 create 配賦代號轉檔作業 by 2295 
 * 106.11.28 add 輸入的轉換月份=系統月份時.才轉檔 by 2295  
 * 107.01.22 add 增加A01_operation_month by 2295
 * 107.01.31 add 每月28日凌晨1:00執行排程 by 2295
 * 107.04.03 add 增加單一機構轉檔 by 2295
 * 107.04.09 fix 調整為每月25日凌1:00執行排程
 */
package com.tradevan.util.transfer;


import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.*;
import com.tradevan.util.dao.DataObject;

public class transBankNo{
    static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    static final String driverName = "oracle.jdbc.driver.OracleDriver";
    static String dbURL = "";//jdbc:oracle:thin:@172.20.5.22:1521:TBOAF"; 
    static String userName = ""; 
    static String userPwd = "";
    static Connection dbConn = null;
    static ResultSet rs = null;
    static Statement stat = null;
    /*{"更新的tablename","rule類別","功能類別"}*/
    static String[][] TBANK = {
        {"bn02","5",""},{"wlx02","5",""},{"wlx02_log","5",""},{"wlx02_m","6",""},{"wlx02_m_log","6",""},
        {"bn01","2",""},{"wlx01","2",""},{"wlx01_log","2",""},{"wlx01_m","2",""},{"wlx01_m_log","2",""},
        {"wlx01_audit","2",""},{"bn04_log","5",""},{"ba01","3",""},{"bn01_reset","2",""},{"wlxoperate_log","3",""},
        {"rpt_month","4",""},{"wlx04","2",""},{"wlx04_log","2",""}
        };
    static String[][] WTT = {
        {"wtt01","1",""},{"wtt02","8",""},{"wtt04","8",""},{"wtt04_1d","8",""},{"muser_data","8",""},{"muser_data_log","8",""},
        {"wtt01_log","1",""},{"wtt06","8",""},{"wlx_apply_lock","4",""},{"wlx_apply_lock_log","4",""}        
    };
    static String[][] OTHER_WML={
        {"wlx05_atm_setup","2","FX005AW"},{"wlx05_atm_setup_log","2","FX005AW"},                
        {"wlx06_m_outpush","2","FX006W"},{"wlx06_m_outpush_log","2","FX006W"},                
        {"wlx05_m_atm","2","FX005W"},{"wlx05_m_atm_log","2","FX005W"},                    
        {"wlx07_m_checkbank","2","FX007WA"},{"wlx07_m_checkbank_log","2","FX007WA"},              
        {"wlx07_m_credit","2","FX007WB"},{"wlx07_m_credit_log","2","FX007WB"},                 
        {"wlx08_s_gage","2","FX008W"},{"wlx08_s_gage_apply","2","FX008W"},{"wlx08_s_gage_log","2","FX008W"},{"wlx08_s_gage_apply_log","2","FX008W"},
        {"wlx09_s_warning","2","FX009W"},{"wlx09_s_warning_log","2","FX009W"},                
        {"wlx_s_rate","2","FX010W"},{"wlx_s_rate_log","2","FX010W"},      
        {"boaf_assetcheck","7","FX011W"},{"boaf_assetcheck_log","7","FX011W"},                
        {"boaf_account","2","FX012W"},{"boaf_account_log","2","FX012W"},               
        {"wlx_trainning","5","FX013W"}
    };

    static String[][] WML={
        {"wml01","4","WML01"},{"wml01_log","4","WML01"},{"wml02","4","WML01"},{"wml02_log","4","WML01"},{"wml03","4","WML01"},{"wml03_log","4","WML01"},
        {"wml01_lock","4","WML01"},{"wml01_m_upload","4","WML01"},//{"wml01_upload","4"},
        {"A01","4","A01"},{"A02","4","A02"},{"A03","4","A03"},{"A04","4","A04"},{"A05","4","A05"},{"A06","4","A06"},{"A08","4","A08"},{"A09","4","A09"},{"A10","4","A10"},{"A12","4","A12"},{"A99","4","A99"},{"F01","4","F01"},
        {"A01_log","4","A01"},{"A02_log","4","A02"},{"A03_log","4","A03"},{"A04_log","4","A04"},{"A05_log","4","A05"},{"A06_log","4","A06"},{"A08_log","4","A08"},
        {"A09_log","4","A09"},{"A10_log","4","A10"},{"A12_log","4","A12"},{"A99_log","4","A99"},{"F01_log","4","F01"},
        {"wlx10_m_loan","2","A11"},{"wlx10_m_loan_apply","2","A11"},{"wlx10_m_loan_log","2","A11"},
        {"A01_operation","4","ZZ043W"},{"A01_operation_month","4","ZZ043W"},{"A02_operation","4","ZZ043W"},{"A03_operation","4","ZZ043W"},{"A04_operation","4","ZZ043W"},
        {"wr_operation","4","ZZ045W"},{"wr_operation_tmp","4","ZZ045W"},
        {"agri_loan","4","ZZ092W"}
    };

    static String[][] TC={
        {"exreportf","5",""},{"exreportf_log","5",""},{"exwarningf","2",""},{"exwarningf_log","2",""},{"exdistripf","5",""},{"exdistripf_log","5",""}
    };

    static String[][] MC={
        {"mis_violatelaw","2","MC001W"},{"mis_violatelaw_log","2","MC001W"},
        {"mis_moneylaunder","2","MC002W"},{"mis_moneylaunder_log","2","MC002W"},  
        {"mis_rptyear","2","MC003W"},{"mis_rptyear_log","2","MC003W"},
        {"mis_loal","2","MC010W"},
        {"mis_dffe","2","MC011W"},
        {"mis_em","2","MC012W"},
        {"mis_ap","2","MC013W"},
        {"mis_ta","2","MC014W"}
    };

    static String[][] FL={
        {"frm_exMaster","2",""},{"frm_exdef","2",""},{"frm_snrtdoc","2",""},{"frm_agsncorrdoc","2",""},{"frm_refunddoc","2",""}
    };

    static String[][] TM={
        {"loan_bn01","4",""},{"loan_rpt","4",""},{"loanapply_bn01","4",""},{"loanapply_rpt","4",""},
        {"loanapply_wml01","4",""},{"loanapply_wml01_log","4",""}
    };
    
    
    static Map rule1 = new HashMap(); 
    static Map rule2 = new HashMap(); 
    static Map rule3 = new HashMap(); 
    static Map rule4 = new HashMap(); 
    static Map rule5 = new HashMap(); 
    static Map rule6 = new HashMap(); 
    static Map rule7 = new HashMap(); 
    static Map rule8 = new HashMap(); 
    static String pbank_no="",bank_kind="",src_bank_no="",bank_no="",pbank_no_new="";
    //static int tbank_cnt=0,wtt_cnt=0,other_wml_cnt=0,wml_cnt=0,tc_cnt=0,mc_cnt=0,fl_cnt=0,tm_cnt=0;
    
    static String[][] funCnt={
        {"FX005AW","0","OTHER_WML"},{"FX006W" ,"0","OTHER_WML"},{"FX005W" ,"0","OTHER_WML"} ,{"FX007WA","0","OTHER_WML"},{"FX007WB","0","OTHER_WML"},
        {"FX008W" ,"0","OTHER_WML"},{"FX009W" ,"0","OTHER_WML"},{"FX010W" ,"0","OTHER_WML"},{"FX011W" ,"0","OTHER_WML"},{"FX012W" ,"0","OTHER_WML"},
        {"FX013W" ,"0","OTHER_WML"},
        {"WML01"  ,"0","WML"},{"A01"    ,"0","WML"},{"A02"    ,"0","WML"},{"A03"    ,"0","WML"},
        {"A04"    ,"0","WML"},{"A05"    ,"0","WML"},{"A06"    ,"0","WML"},{"A08"    ,"0","WML"},{"A09"    ,"0","WML"},
        {"A10"    ,"0","WML"},{"A11"    ,"0","WML"},{"A12"    ,"0","WML"},{"A99"    ,"0","WML"},{"F01"    ,"0","WML"},
        {"ZZ043W" ,"0","WML"},{"ZZ045W" ,"0","WML"},{"ZZ092W" ,"0","WML"},
        {"MC001W" ,"0","MC"},{"MC002W" ,"0","MC"},{"MC003W" ,"0","MC"},{"MC010W" ,"0","MC"},{"MC011W" ,"0","MC"},
        {"MC012W" ,"0","MC"},{"MC013W" ,"0","MC"},{"MC014W" ,"0","MC"}
    };
    
    static String[][] bakCnt={
        {"BAK","0"},{"TBANK","0"},{"WTT","0"},{"OTHER_WML","0"},{"WML","0"},{"TC","0"},{"MC","0"},{"FL","0"},
        {"TM","0"}
    };
    
    static String[] UpdUserid={
        "bn01","bn02","boaf_account","boaf_assetcheck","exdistripf","exreportf","exwarningf","loan_bn01","loan_rpt",
        "loanapply_bn01","loanapply_rpt","loanapply_wml01","loanapply_wml01_log","muser_data","muser_data_log",
        "wlx_apply_lock","wlx_apply_lock_log","wlx_trainning","wlx01","wlx01_audit","wlx01_log","wlx01_m","wlx01_m_log",
        "wlx02","wlx02_log","wlx02_m","wlx02_m_log","wlx04","wlx04_log",
        "wlx05_atm_setup","wlx05_atm_setup_log","wlx05_m_atm","wlx05_m_atm_log","wlx06_m_outpush","wlx06_m_outpush_log",
        "wlx07_m_checkbank","wlx07_checkbank_log","wlx07_m_credit","wlx07_m_credit_log",
        "wlx08_s_gage","wlx08_s_gage_apply","wlx08_s_gage_apply_log","wlx09_s_warning","wlx09_s_warning_log",
        "wlx10_m_loan","wlx10_m_loan_apply","wlx10_m_loan_log",
        "wml01","wml01_lock","wml01_log","wml02","wml02_log","wml03","wml03_log",
        "wtt01","wtt01_log","wtt02","wtt04","wtt04_1d"
    };
    
    static List insertSQL = new LinkedList();
    static List updateSQL = new LinkedList();
    static List newSQL = new LinkedList();
    static String sqlCmd="";

	 public static void main(String args[]) {
	     if(args.length == 1){     
	        System.out.println("bank_no="+args[0]);
	        transBankNo.doTransfer(args[0]);  
	     }else{
	         transBankNo.doTransfer("");    
	     }	     
	 }
	 
   public static String doTransfer(String trans_bank_no){       
       String errMsg = "";
       String  txtline  = null; 
       List paramList = new ArrayList();
       DataObject bean = null;
       
       try {
           
           logDir  = new File(Utility.getProperties("logDir"));
           if(!logDir.exists()){
               if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                  System.out.println("目錄新增失敗");
               }    
           }
           //nowlog = logcalendar.getTime();
           logfile = new File(logDir + System.getProperty("file.separator") + "transBankNo."+ logfileformat.format(nowlog));                       
           System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"transBankNo."+ logfileformat.format(nowlog));
           logos = new FileOutputStream(logfile,true);                         
           logbos = new BufferedOutputStream(logos);
           logps = new PrintStream(logbos);
           
           //取得機構資料
           dbURL = Utility.getProperties("BOAFDBURL");
           userName = Utility.getProperties("rptID");
           userPwd = Utility.getProperties("rptPwd");
           Class.forName(driverName); 
           dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
           System.out.println("Connection Successful!");
           stat = dbConn.createStatement(); 
           
           String pre_pbank_no="";
           List pbank_Map = new LinkedList();
           List pbankList = new LinkedList();
           Map h = new HashMap();          
           
           //須轉換的機構代碼
           //107.04.03 add若有機構代碼,則依單一機構代碼轉換
           if(!"".equals(trans_bank_no)){
             rs = stat.executeQuery("select pbank_no from ba01_trans where trans_date is null and pbank_no='"+trans_bank_no+"' group by pbank_no order by pbank_no"); 
           }else{
             //106.11.28 add輸入的欲轉換月份=系統月份,才執行轉換
             rs = stat.executeQuery("select pbank_no from ba01_trans where trans_date is null and TO_CHAR(sysdate, 'MM') = TO_CHAR(online_date, 'MM') group by pbank_no order by pbank_no");
           //for 測試用
           //rs = stat.executeQuery("select pbank_no from ba01_trans where trans_date is null group by pbank_no order by pbank_no");
           }   
           while (rs.next()) {   
               pbank_no = rs.getString("pbank_no");
               pbankList.add(pbank_no);
           };
           
           System.out.println("pbankList.size="+pbankList.size());
           if(pbankList.size() != 0){
           //設定轉檔table
           setRuleMap(TBANK);
           setRuleMap(WTT);
           setRuleMap(OTHER_WML);
           setRuleMap(WML);
           setRuleMap(TC);
           setRuleMap(MC);
           setRuleMap(FL);
           setRuleMap(TM); 
           
           List bank_no_list = new LinkedList();
           List pbank_no_list = new LinkedList();
           String pbank_no_new = "";
           for(int i=0;i<pbankList.size();i++){
               pbank_no=(String)pbankList.get(i);
               System.out.println("pbank_no="+pbank_no);
               rs = stat.executeQuery("select ba01_trans.*,bank_no as pbank_no_new from ba01_trans where pbank_no='"+pbank_no+"' and trans_date is null order by pbank_no,bank_kind,bank_no");
               System.out.println("select ba01_trans.*,bank_no as pbank_no_new from ba01_trans where pbank_no='"+pbank_no+"' and trans_date is null order by pbank_no,bank_kind,bank_no");
               while (rs.next()) {
                   if("0".equals(rs.getString("bank_kind"))){
                       pbank_no_new = rs.getString("pbank_no_new");
                   }
                   paramList = new LinkedList();
                   paramList.add(rs.getString("bank_kind"));//機構類別 0:總機構 1:分支機構
                   paramList.add(rs.getString("src_bank_no"));//原始機構代碼
                   paramList.add(rs.getString("bank_no"));//變更後機構代碼
                   paramList.add(pbank_no_new);//變更後總機構代碼
                   bank_no_list.add(paramList);
               }
               if(bank_no_list.size() !=0){
                   h.put(pbank_no, bank_no_list);    
                   System.out.println("bank_no_list.size()="+bank_no_list.size());
                   bank_no_list = new LinkedList();
               }
           }
           int insert_cnt=0;
           String tbl_name = "";
           
           System.out.println("pbankList.size="+pbankList.size());
           
           for(int i=0;i<pbankList.size();i++){
               System.out.println("i="+i);
               pbank_no=(String)pbankList.get(i);
               pbank_Map = (LinkedList)h.get(pbank_no);
               System.out.println("pbank_no="+pbank_no);
               printLog(logps,"原總機構代碼:pbank_no["+pbank_no+"]");
               for(int bak_i=0;bak_i<bakCnt.length;bak_i++){
                   bakCnt[bak_i][1]="0";
               }                  
               for(int fun_i=0;fun_i<funCnt.length;fun_i++){
                   funCnt[fun_i][1]="0";                    
               }
               insertSQL = new LinkedList(); 
               updateSQL = new LinkedList(); 
               //tbank_cnt=0;wtt_cnt=0;other_wml_cnt=0;wml_cnt=0;tc_cnt=0;mc_cnt=0;fl_cnt=0;tm_cnt=0;               
               for(int j=0;j<pbank_Map.size();j++){
                   bank_kind=(String)((List)pbank_Map.get(j)).get(0);//機構類別 0:總機構 1:分支機構
                   src_bank_no=(String)((List)pbank_Map.get(j)).get(1);//原始機構代碼
                   bank_no=(String)((List)pbank_Map.get(j)).get(2);//變更後機構代碼
                   pbank_no_new=(String)((List)pbank_Map.get(j)).get(3);//變更後機構代碼
                   printLog(logps,"機構類別["+bank_kind+"]:原始機構代碼["+src_bank_no+"]:變更後機構代碼["+bank_no+"]:變更後總機構代碼["+pbank_no_new+"]");
                   System.out.println("bank_kind="+bank_kind+":src_bank_no="+src_bank_no+":bank_no="+bank_no+":pbank_no_new="+pbank_no_new);
                   
                   getBAKCount("TBANK",TBANK,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("WTT",WTT,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("OTHER_WML",OTHER_WML,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("WML",WML,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("TC",TC,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("MC",MC,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("FL",FL,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
                   getBAKCount("TM",TM,bank_kind,src_bank_no,bank_no,pbank_no,pbank_no_new);
               }//end of pbank_Map
                
               
               printLog(logps,"開始批次備份資料");     
               stat.clearBatch();
               if(insertSQL.size() >= 1){
                   dbConn.setAutoCommit(false);        
                   for(int idx = 0;idx < insertSQL.size();idx++){
                       System.out.println((String)insertSQL.get(idx));
                       stat.addBatch((String)insertSQL.get(idx));                       
                       //printLog(logps,(String)insertSQL.get(idx));
                   }
                   int[] rowCount  = stat.executeBatch();
                   System.out.println("insertSQL.rowCount="+rowCount.length);
                   
                   int idx=0;
                   insert_cnt=0;
                   boolean updateOK=true;
                   while(idx < rowCount.length){
                       insert_cnt += rowCount[idx];                      
                       System.out.println("idx="+idx+":"+rowCount[idx]+":sql="+(String)insertSQL.get(idx));              
                       printLog(logps,"idx="+idx+":updateOK="+rowCount[idx]+":sql="+(String)insertSQL.get(idx));
                       idx++;
                   }
                   dbConn.commit();  
                   
                   System.out.println("已備份資料"+insert_cnt+"筆");
                   printLog(logps,"已備份資料"+insert_cnt+"筆");
               }//end of insertSQL
               int bakidx=0;
               insertSQL = new LinkedList(); 
               bakCnt[0][1] = String.valueOf(insert_cnt);//總備份筆數
               for(bakidx=0;bakidx<bakCnt.length;bakidx++){
                   System.out.println("bakType["+bakCnt[bakidx][0]+"]="+bakCnt[bakidx][1]);
                   printLog(logps,"備份類別["+bakCnt[bakidx][0]+"]備份筆數:"+bakCnt[bakidx][1]);
                   sqlCmd = "insert into ba01_trans_master values ('"+pbank_no_new+"','"+pbank_no+"','"+bakCnt[bakidx][0]+"',"+bakCnt[bakidx][1]+",sysdate)";                
                   insertSQL.add(sqlCmd); 
               }
                              
               printLog(logps,"開始批次更新資料");
               stat.clearBatch();
               if(updateSQL.size() >= 1){
                   dbConn.setAutoCommit(false);        
                   for(int idx = 0;idx < updateSQL.size();idx++){
                       stat.addBatch((String)updateSQL.get(idx));
                       System.out.println((String)updateSQL.get(idx)+";");
                       //printLog(logps,(String)updateSQL.get(idx));
                   }
                   int[] rowCount  = stat.executeBatch();
                   System.out.println("updateSQL.rowCount="+rowCount.length);
                   
                   int idx=0;
                   int update_cnt=0;
                   boolean updateOK=true;
                   while(idx < rowCount.length){
                       update_cnt += rowCount[idx];                       
                       System.out.println("idx="+idx+":"+rowCount[idx]+":sql="+(String)updateSQL.get(idx));              
                       printLog(logps,"idx="+idx+":updateOK="+rowCount[idx]+":sql="+(String)updateSQL.get(idx));
                       idx++;
                   }
                   dbConn.commit();
                   
                   System.out.println("已更新資料"+update_cnt+"筆");
                   printLog(logps,"已更新資料"+update_cnt+"筆");
               }//end of updateSQL
               
                  
               for(bakidx=0;bakidx<funCnt.length;bakidx++){
                   System.out.println("funName["+funCnt[bakidx][0]+"]="+funCnt[bakidx][1]);
                   printLog(logps,"["+funCnt[bakidx][0]+"]功能筆數:"+funCnt[bakidx][1]);
                   sqlCmd = "insert into ba01_trans_detail values ('"+pbank_no_new+"','"+funCnt[bakidx][2]+"','"+funCnt[bakidx][0]+"',"+funCnt[bakidx][1]+",sysdate)";                
                   insertSQL.add(sqlCmd); 
               }
                               
               sqlCmd = "update ba01_trans set trans_date=sysdate where pbank_no='"+pbank_no+"'";
               insertSQL.add(sqlCmd); 
                           
               printLog(logps,"更新轉檔筆數資料");
               stat.clearBatch();
               if(insertSQL.size() >= 1){
                    dbConn.setAutoCommit(false);        
                    for(int idx = 0;idx < insertSQL.size();idx++){
                        stat.addBatch((String)insertSQL.get(idx));
                        System.out.println((String)insertSQL.get(idx));
                        //printLog(logps,(String)insertSQL.get(idx));
                    }
                    int[] rowCount  = stat.executeBatch();
                    System.out.println("insertSQL.rowCount="+rowCount.length);
                    
                    int idx=0;
                    insert_cnt=0;
                    boolean updateOK=true;
                    while(idx < rowCount.length){
                        insert_cnt += rowCount[idx];
                        System.out.println("idx="+idx+":"+rowCount[idx]+":sql="+(String)insertSQL.get(idx));              
                        printLog(logps,"idx="+idx+":updateOK="+rowCount[idx]+":sql="+(String)insertSQL.get(idx));
                        idx++;
                    }
                    dbConn.commit();                    
               }//end of insertSQL
           }//end of pbankList
           
           }
       }catch(Exception e){
           System.out.println("doTransfer Error:"+e+e.getMessage());
           printLog(logps,"doTransfer Error:"+e.getMessage());     
       }
       
       return errMsg;
   }
   
   public static void printLog(PrintStream logps,String errRptMsg){
       if(!errRptMsg.equals("")){
          logcalendar = Calendar.getInstance(); 
          nowlog = logcalendar.getTime();
          logps.println(logformat.format(nowlog)+errRptMsg);
          logps.flush();
       }
   }
   private static void setRuleMap(String[][] tbl){
       for(int tbl_idx=0;tbl_idx<tbl.length;tbl_idx++){
           if("1".equals(tbl[tbl_idx][1])){
               rule1.put(tbl[tbl_idx][0],"1");
           }
           if("2".equals(tbl[tbl_idx][1])){
               rule2.put(tbl[tbl_idx][0],"1");
           }
           if("3".equals(tbl[tbl_idx][1])){
               rule3.put(tbl[tbl_idx][0],"1");
           }
           if("4".equals(tbl[tbl_idx][1])){
               rule4.put(tbl[tbl_idx][0],"1");
           }
           if("5".equals(tbl[tbl_idx][1])){
               rule5.put(tbl[tbl_idx][0],"1");
           }
           if("6".equals(tbl[tbl_idx][1])){
               rule6.put(tbl[tbl_idx][0],"1");
           }
           if("7".equals(tbl[tbl_idx][1])){
               rule7.put(tbl[tbl_idx][0],"1");
           }
           if("8".equals(tbl[tbl_idx][1])){
               rule8.put(tbl[tbl_idx][0],"1");
           }
       }
   }
   
   private static void getBAKCount(String bakType,String[][] tbl,String bank_kind,String src_bank_no,String bank_no,String pbank_no,String pbank_no_new){
       String tbl_name = "";
       String fun_name = "";
      
       try{
           if("0".equals(bank_kind)){//總機構代碼              
               for(int tbank_idx=0;tbank_idx<tbl.length;tbank_idx++){
                   tbl_name = tbl[tbank_idx][0];
                   fun_name = tbl[tbank_idx][2];
                   System.out.println("tbank_idx="+tbank_idx+":tbl_name="+tbl_name);
                   
                   if("1".equals((String)rule1.get(tbl_name))){                       
                       //System.out.println("SQL=select count(*) as data from "+tbl_name+" where tbank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where tbank_no='"+src_bank_no+"'");//總機構代碼
           
                       while (rs.next()) {
                           System.out.println("SQL=select count(*) as data from "+tbl_name+" where tbank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL=select count(*) as data from "+tbl_name+" where tbank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));                          
                       }     
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where tbank_no='"+src_bank_no+"'";//原總機構代碼
                       insertSQL.add(sqlCmd);   
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set tbank_no='"+bank_no+"',muser_id='"+bank_no+"'||substr(muser_id,8,10) where tbank_no ='"+src_bank_no+"'";
                       updateSQL.add(sqlCmd);     
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set add_user='"+bank_no+"'||substr(add_user,8,10) where add_user like '"+src_bank_no+"%'";//原總機構代碼 106.09.21 fix //106.10.02 OK
                       updateSQL.add(sqlCmd);   
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   
                   if("1".equals((String)rule2.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"'");//原總機構代碼
                       
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }  
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where bank_no='"+src_bank_no+"'";//原總機構代碼
                       insertSQL.add(sqlCmd);  
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set bank_no='"+bank_no+"'";
                       if("bn01".equals(tbl_name)){
                           sqlCmd += ",exchange_no='"+bank_no+"'"
                                  + ",trans_date=sysdate"
                                  + ",ori_bank_no='"+ src_bank_no+"'";
                       }
                       sqlCmd += " where bank_no ='"+src_bank_no+"'";//原總機構代碼
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   
                   if("1".equals((String)rule3.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where pbank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where pbank_no='"+src_bank_no+"'");//總機構代碼
                       
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where pbank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where pbank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }       
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where pbank_no='"+src_bank_no+"'"; //原總機構代碼
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       if("ba01".equals(tbl_name)){
                           sqlCmd = " update "+tbl_name+" set pbank_no='"+bank_no+"',bank_no='"+bank_no+"' where pbank_no ='"+src_bank_no+"' and bank_no='"+src_bank_no+"'";//--原總機構代碼 
                       }else{
                           sqlCmd = " update "+tbl_name+" set pbank_no='"+bank_no+"' where pbank_no ='"+src_bank_no+"' and bank_no is null";//--總機構
                       }
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   if("1".equals((String)rule4.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_code='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where bank_code='"+src_bank_no+"'");//總機構代碼
                       
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_code='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where bank_code='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }    
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where bank_code='"+src_bank_no+"'";//原總機構代碼
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set  bank_code='"+bank_no+"' where bank_code='"+src_bank_no+"'";//原總機構代碼
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       
                       if("wlx_apply_lock".equals(tbl_name) || "wlx_apply_lock_log".equals(tbl_name)){ 
                           sqlCmd = " update "+tbl_name+" set add_user='"+bank_no+"'||substr(add_user,8,10) where add_user like '"+src_bank_no+"%'";//原總機構代碼 106.09.21 fix //106.10.02 OK
                           updateSQL.add(sqlCmd);   
                           printLog(logps,sqlCmd);
                           System.out.println(sqlCmd);
                       }
                   }
                   
                   if("1".equals((String)rule5.get(tbl_name))){//106.09.21 add 
                       if("exreportf".equals(tbl_name) || "exreportf_log".equals(tbl_name) || "exdistripf".equals(tbl_name) || "exdistripf_log".equals(tbl_name) || "wlx_trainning".equals(tbl_name)){                        
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"'");//總機構代碼+總機構                           
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }    
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where tbank_no='"+pbank_no+"' and bank_no='"+src_bank_no+"'";//總機構代碼+總機構
                       insertSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set  tbank_no='"+pbank_no_new+"',bank_no='"+bank_no+"'"
                              + " where tbank_no ='"+pbank_no+"' and bank_no='"+src_bank_no+"'";//總機構代碼+分支機構
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       }//end of exreportf/exreportf_log/exdistripf/exdistripf_log
                   }
                   
                   if("1".equals((String)rule7.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"'");//總機構                          
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }  
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where examine='"+src_bank_no+"'";//總機構    
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set  examine='"+bank_no+"' where examine='"+src_bank_no+"'";//總機構   
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   
                   if("1".equals((String)rule8.get(tbl_name))){
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where  muser_id like '"+src_bank_no+"%'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where  muser_id like '"+src_bank_no+"%'");//總機構代碼
                       
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where  muser_id like '"+src_bank_no+"%':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where  muser_id like '"+src_bank_no+"%':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }  
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where muser_id like '"+src_bank_no+"%'";//總機構代碼
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set muser_id='"+bank_no+"'||substr(muser_id,8,10) where muser_id like '"+src_bank_no+"%'";//總機構代碼
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       if(!"wtt06".equals(tbl_name)){       
                           sqlCmd = " update "+tbl_name+" set user_id='"+bank_no+"'||substr(user_id,8,10) where user_id like '"+src_bank_no+"%'";//總機構代碼
                           updateSQL.add(sqlCmd); 
                           printLog(logps,sqlCmd);
                           System.out.println(sqlCmd);
                       }
                      
                   }
                   
                   if("wlxoperate_log".equals(tbl_name)){//106.09.21 add
                       sqlCmd = " update "+tbl_name+" set muser_id='"+bank_no+"'||substr(muser_id,8,10) where muser_id like '"+src_bank_no+"%'";//總機構代碼
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }else if(tbl_name.endsWith("log")){//106.09.21 add       
                       sqlCmd = " update "+tbl_name+" set user_id_c='"+bank_no+"'||substr(user_id_c,8,10) where user_id_c like '"+src_bank_no+"%'";//總機構代碼
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   
                   for(int uid_idx=0;uid_idx<UpdUserid.length;uid_idx++){//106.09.21 add
                       if(UpdUserid[uid_idx].equals(tbl_name)){                          
                           sqlCmd = " update "+tbl_name+" set user_id='"+bank_no+"'||substr(user_id,8,10) where user_id like '"+src_bank_no+"%'";//總機構代碼
                           updateSQL.add(sqlCmd); 
                           printLog(logps,sqlCmd);
                           System.out.println(sqlCmd);
                       }
                   }
                
               }
           }else if("1".equals(bank_kind)){//分支機構代碼
               for(int tbank_idx=0;tbank_idx<tbl.length;tbank_idx++){
                   tbl_name = tbl[tbank_idx][0];
                   
                   if("1".equals((String)rule3.get(tbl_name))){
                       sqlCmd = " update "+tbl_name+" set pbank_no='"+pbank_no_new+"',bank_no='"+bank_no+"' where pbank_no ='"+pbank_no+"' and bank_no='"+src_bank_no+"'";//--原總/分支機構
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   
                   if("1".equals((String)rule5.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"'");//總機構代碼+分支機構                           
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where tbank_no = '"+pbank_no+"' and bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }    
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where tbank_no='"+pbank_no+"' and bank_no='"+src_bank_no+"'";//總機構代碼+分支機構
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set  tbank_no='"+pbank_no_new+"',bank_no='"+bank_no+"'";
                       
                       if("bn02".equals(tbl_name)){
                           sqlCmd +=",exchange_no='"+bank_no+"'"
                                  + ",trans_date=sysdate"
                                  + ",ori_tbank_no='"+pbank_no+"'"
                                  + ",ori_bank_no='"+src_bank_no+"'";
                       }
                       sqlCmd += " where tbank_no ='"+pbank_no+"' and bank_no='"+src_bank_no+"'";//總機構代碼+分支機構
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   if("1".equals((String)rule6.get(tbl_name))){ 
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"'");//分支機構                           
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where bank_no='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           //tbl_cnt += Integer.parseInt(rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       }
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where bank_no='"+src_bank_no+"'";
                       insertSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set bank_no='"+bank_no+"' where bank_no='"+src_bank_no+"'";
                       updateSQL.add(sqlCmd); 
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
                   if("1".equals((String)rule7.get(tbl_name))){   
                       //System.out.println("SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"'");
                       rs = stat.executeQuery("select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"'");//分支機構                          
                       while (rs.next()) {
                           System.out.println("SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           printLog(logps,"SQL:select count(*) as data from "+tbl_name+" where examine='"+src_bank_no+"':"+tbl[tbank_idx][0]+".count="+rs.getString("data"));
                           addCnt(bakType,Integer.parseInt(rs.getString("data")));
                           addFunCnt(fun_name,Integer.parseInt(rs.getString("data")));      
                       } 
                       
                       sqlCmd = " insert into bak_"+tbl_name
                              + " select "+tbl_name+".*,sysdate from "+tbl_name+" where examine='"+src_bank_no+"'";//分支機構
                       insertSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                       sqlCmd = " update "+tbl_name+" set examine='"+bank_no+"' where examine='"+src_bank_no+"'";//分支機構
                       updateSQL.add(sqlCmd);
                       printLog(logps,sqlCmd);
                       System.out.println(sqlCmd);
                   }
               }//end of tbl
           }
      
       }catch(Exception e){
           printLog(logps,"getBAKCount Error:"+e+e.getMessage());
           System.out.println("getBAKCount Error:"+e+e.getMessage());
       }   
   }   
   
  
   private static void addCnt(String bakType,int cnt){/*計算各功能大類更新筆數*/
       for(int i=0;i<bakCnt.length;i++){
           if(bakCnt[i][0].equals(bakType)){
               //System.out.println("bakType["+bakType+"]="+bakCnt[i][1]);
               bakCnt[i][1] = String.valueOf(Integer.parseInt(bakCnt[i][1])+cnt);
               //System.out.println("bakType["+bakType+"]="+bakCnt[i][1]);
           }
       }
   }
   
   private static void addFunCnt(String funName,int cnt){/*計算各功能更新筆數*/
       if("".equals(funName)) return ;
       for(int i=0;i<funCnt.length;i++){
           if(funCnt[i][0].equals(funName)){
               //System.out.println("funName["+funName+"]="+funCnt[i][1]);
               funCnt[i][1] = String.valueOf(Integer.parseInt(funCnt[i][1])+cnt);
               //System.out.println("funName["+funName+"]="+funCnt[i][1]);
           }
       }
   }
}

