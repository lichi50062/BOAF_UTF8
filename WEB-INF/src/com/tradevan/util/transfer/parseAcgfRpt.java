/* 
 * 101.08.10 create 農信保資料轉檔(M106/M201/M206) by 2295
 * 101.12.17 fix M106使用者上傳格式調整少一列 by 2295  
 */
package com.tradevan.util.transfer;


import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.*;
import com.tradevan.util.dao.DataObject;

public class parseAcgfRpt{
    static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    
	 public static void main(String args[]) {   
	    //parseAcgfRpt.getM106FileData("M106","M10610012.csv","101","2");	    
	    //parseAcgfRpt.getM201FileData("M201","M20110012.csv","101","1");
	    //parseAcgfRpt.getM206FileData("M206","M20610012.csv","101","1");
	 }
	 
   public static String doParserRpt(String report_no,String filename,String m_year,String m_month){
       String errMsg = "";
       if("M106".equals(report_no)){
           errMsg = getM106FileData(report_no,filename,m_year,m_month);
       }else if("M201".equals(report_no)){
           errMsg = getM201FileData(report_no,filename,m_year,m_month);
       }else if("M206".equals(report_no)){
           errMsg = getM206FileData(report_no,filename,m_year,m_month);
       }
       return errMsg;
   }
   //M106匯入CSV檔案  
   public static String getM106FileData(String report_no,String filename,String m_year,String m_month){
           String errMsg = "";
           String  txtline  = null;  
           List allList = new LinkedList();
           List dbData = null;  
           List paramList = new ArrayList();
           StringBuffer sqlCmd = new StringBuffer();  
           List updateDBList = new LinkedList();//0:sql 1:data     
           List updateDBSqlList = new LinkedList();
           List updateDBDataList = new LinkedList();//儲存參數的List
           List dataList = new LinkedList();//儲存參數的data
           try {
               System.out.println("M106 begin");
               
               logDir  = new File(Utility.getProperties("logDir"));
               if(!logDir.exists()){
                   if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                      System.out.println("目錄新增失敗");
                   }    
               }
               logfile = new File(logDir + System.getProperty("file.separator") + "parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));                       
               System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));
               logos = new FileOutputStream(logfile,true);                         
               logbos = new BufferedOutputStream(logos);
               logps = new PrintStream(logbos);                  
               
               sqlCmd.append("select * from M106 where m_year=? and m_month=?");
               paramList.add(m_year);
               paramList.add(m_month);               
               dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
               System.out.println("M106.dbData.size()="+dbData.size());
               printLog(logps,"m_year="+m_year+":m_month="+m_month+":M106.dbData.size()="+dbData.size());
               if(dbData != null && dbData.size() > 0){
                  sqlCmd.setLength(0) ;
                  sqlCmd.append(" delete M106 where m_year=? and m_month=?");
                  updateDBDataList.add(paramList);
                  updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                  updateDBList.add(updateDBSqlList);
               }
               
               sqlCmd.setLength(0) ;
               updateDBDataList = new LinkedList();             
               updateDBSqlList = new LinkedList(); 
               
               String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
               File tmpFile = new File(WMdataDir+filename);        
               Date filedate = new Date(tmpFile.lastModified());
                   
               FileReader  f       = new FileReader(tmpFile);
               LineNumberReader in = new LineNumberReader(f);
               printLog(logps,filename+"======begin==============================");
               int i = 0;
               doLoop://將txt檔案儲存至LinkedList
               while ((txtline = in.readLine()) != null) {
                       i++;
                       if ( i >= 4){
                           dataList = new LinkedList();
                           dataList.add(m_year);
                           dataList.add(m_month);
                           allList = Utility.getStringTokenizerData(txtline,",");
                           //System.out.println("allList.size()="+allList.size());
                           printLog(logps,txtline);
                           dataList.add((String)allList.get(0));//貸款用途別代碼
                           dataList.add((String)allList.get(1));//當年度件數
                           dataList.add((String)allList.get(2));//當年度貸款金額
                           dataList.add((String)allList.get(3));//當年度保證金額
                           dataList.add((String)allList.get(4));//累計件數
                           dataList.add((String)allList.get(5));//累計貸款金額
                           dataList.add((String)allList.get(6));//累計保證金額
                           updateDBDataList.add(dataList);//1:傳內的參數List
                       }
               }   
               in.close();
               f.close();
               printLog(logps,filename+"=======end===============================");
               sqlCmd.append(" insert into M106 values (?,?,?,?,?,?,?,?,?)");
               updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
               updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
               updateDBList.add(updateDBSqlList);
               //寫入合計列
               sqlCmd.setLength(0) ;
               sqlCmd.append(" insert into M106");
               sqlCmd.append(" select m_year,m_month,?,sum(guarantee_cnt_year),sum(loan_amt_year),");
               sqlCmd.append(" sum(guarantee_amt_year),sum(guarantee_cnt_sum),sum(loan_amt_sum),sum(guarantee_amt_sum)");
               sqlCmd.append(" from M106 ");
               sqlCmd.append(" where m_year = ? and m_month=?");
               sqlCmd.append(" group by m_year,m_month");
               updateDBDataList = new LinkedList();             
               updateDBSqlList = new LinkedList(); 
               dataList = new LinkedList();
               dataList.add("0");//合計
               dataList.add(m_year);
               dataList.add(m_month);
               updateDBDataList.add(dataList);
               updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
               updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
               updateDBList.add(updateDBSqlList);
               
               if(DBManager.updateDB_ps(updateDBList)){
                  errMsg = "<br>M106匯入資料庫成功";
                  System.out.println("M106 Insert ok");  
                  printLog(logps,"M106匯入資料庫成功");
               }
               
           }catch(Exception e){
               System.out.println("parseAcgfRpt.getM106FileData Error:"+e+e.getMessage());
               printLog(logps,"parseAcgfRpt.getM106FileData Error:"+e.getMessage());     
           }
           updateDBSqlList = null;
           return errMsg;
   }
   
   //M201匯入CSV檔案  
   public static String getM201FileData(String report_no,String filename,String m_year,String m_month){
           String errMsg = "";
           String  txtline  = null;  
           List allList = new LinkedList();
           List dbData = null;  
           List paramList = new ArrayList();
           StringBuffer sqlCmd = new StringBuffer();  
           List updateDBList = new LinkedList();//0:sql 1:data     
           List updateDBSqlList = new LinkedList();
           List updateDBDataList = new LinkedList();//儲存參數的List
           List dataList = new LinkedList();//儲存參數的data
           try {
               System.out.println("M201 begin");
               
               logDir  = new File(Utility.getProperties("logDir"));
               if(!logDir.exists()){
                   if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                      System.out.println("目錄新增失敗");
                   }    
               }
               logfile = new File(logDir + System.getProperty("file.separator") + "parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));                       
               System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));
               logos = new FileOutputStream(logfile,true);                         
               logbos = new BufferedOutputStream(logos);
               logps = new PrintStream(logbos);   
               
               sqlCmd.append("select * from M201 where m_year=? and m_month=?");
               paramList.add(m_year);
               paramList.add(m_month);               
               dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
               System.out.println("M201.dbData.size()="+dbData.size());
               printLog(logps,"m_year="+m_year+":m_month="+m_month+":M201.dbData.size()="+dbData.size());
               if(dbData != null && dbData.size() > 0){
                  sqlCmd.setLength(0) ;
                  sqlCmd.append(" delete M201 where m_year=? and m_month=?");
                  updateDBDataList.add(paramList);
                  updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                  updateDBList.add(updateDBSqlList);
               }
               
               sqlCmd.setLength(0) ;
               updateDBDataList = new LinkedList();             
               updateDBSqlList = new LinkedList(); 
               
               String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
               File tmpFile = new File(WMdataDir+filename);        
               Date filedate = new Date(tmpFile.lastModified());
                   
               FileReader  f       = new FileReader(tmpFile);
               LineNumberReader in = new LineNumberReader(f);
               printLog(logps,filename+"======begin==============================");
               int i = 0;
               doLoop://將txt檔案儲存至LinkedList
               while ((txtline = in.readLine()) != null) {
                       i++;
                       if ( i >= 4){
                           dataList = new LinkedList();
                           dataList.add(m_year);
                           dataList.add(m_month);
                           allList = Utility.getStringTokenizerData(txtline,",");
                           //System.out.println("allList.size()="+allList.size());
                           printLog(logps,txtline);
                           dataList.add((String)allList.get(0));//保證項目別代碼
                           dataList.add((String)allList.get(1));//本月份保證件數
                           dataList.add((String)allList.get(2));//本月份貸款金額
                           dataList.add((String)allList.get(3));//本月份保證金額
                           dataList.add((String)allList.get(4));//本月份保證餘額
                           dataList.add((String)allList.get(5));//本年度保證件數
                           dataList.add((String)allList.get(6));//本年度貸款金額
                           dataList.add((String)allList.get(7));//本年度保證金額
                           dataList.add((String)allList.get(8));//本年度保證餘額
                           dataList.add((String)allList.get(9));//累計保證件數
                           dataList.add((String)allList.get(10));//累計貸款金額
                           dataList.add((String)allList.get(11));//累計保證金額
                           dataList.add((String)allList.get(12));//累計保證餘額
                           updateDBDataList.add(dataList);//1:傳內的參數List
                       }
               }   
               in.close();
               f.close();
               printLog(logps,filename+"=======end===============================");
               sqlCmd.append(" insert into M201 values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
               updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
               updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
               updateDBList.add(updateDBSqlList);
               //寫入合計列
               sqlCmd.setLength(0) ;
               sqlCmd.append(" insert into M201");
               sqlCmd.append(" select m_year,m_month,?,sum(guarantee_cnt_month),");
               sqlCmd.append(" sum(loan_amt_month),sum(guarantee_amt_month),sum(guarantee_bal_month),");
               sqlCmd.append(" sum(guarantee_cnt_year),sum(loan_amt_year),sum(guarantee_amt_year),");
               sqlCmd.append(" sum(guarantee_bal_year),sum(guarantee_cnt_sum),sum(loan_amt_sum),");
               sqlCmd.append(" sum(guarantee_amt_sum),sum(guarantee_bal_sum)");
               sqlCmd.append(" from M201 ");
               sqlCmd.append(" where m_year = ? and m_month=?");
               sqlCmd.append(" group by m_year,m_month");
              
               updateDBDataList = new LinkedList();             
               updateDBSqlList = new LinkedList(); 
               dataList = new LinkedList();
               dataList.add("0");//合計
               dataList.add(m_year);
               dataList.add(m_month);
               updateDBDataList.add(dataList);
               updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
               updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
               updateDBList.add(updateDBSqlList);
               
               if(DBManager.updateDB_ps(updateDBList)){
                  errMsg = "<br>M201匯入資料庫成功";
                  System.out.println("M201 Insert ok");     
                  printLog(logps,"M201寫入資料庫成功");
               }
               
           }catch(Exception e){
               System.out.println("parseAcgfRpt.getM201FileData Error:"+e+e.getMessage());
               printLog(logps,"parseAcgfRpt.getM201FileData Error:"+e.getMessage());     
           }
           updateDBSqlList = null;  
           return errMsg;
   }
   
   //M206匯入CSV檔案  
   public static String getM206FileData(String report_no,String filename,String m_year,String m_month){
       String errMsg = "";
       String  txtline  = null;  
       List allList = new LinkedList();
       List dbData = null;  
       List paramList = new ArrayList();
       StringBuffer sqlCmd = new StringBuffer();  
       List updateDBList = new LinkedList();//0:sql 1:data     
       List updateDBSqlList = new LinkedList();
       List updateDBDataList = new LinkedList();//儲存參數的List
       List dataList = new LinkedList();//儲存參數的data
       try {
           System.out.println("M206 begin");
           logDir  = new File(Utility.getProperties("logDir"));
           if(!logDir.exists()){
               if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                  System.out.println("目錄新增失敗");
               }    
           }
           logfile = new File(logDir + System.getProperty("file.separator") + "parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));                       
           System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseAcgfRpt_"+report_no+"."+ logfileformat.format(nowlog));
           logos = new FileOutputStream(logfile,true);                         
           logbos = new BufferedOutputStream(logos);
           logps = new PrintStream(logbos);   
        
           sqlCmd.append("select * from M206 where m_year=? and m_month=?");
           paramList.add(m_year);
           paramList.add(m_month);               
           dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
           System.out.println("M206.dbData.size()="+dbData.size());
           printLog(logps,"m_year="+m_year+":m_month="+m_month+":M206.dbData.size()="+dbData.size());
           if(dbData != null && dbData.size() > 0){
              sqlCmd.setLength(0) ;
              sqlCmd.append(" delete M206 where m_year=? and m_month=?");              
              updateDBDataList.add(paramList);
              updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
              updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
              updateDBList.add(updateDBSqlList);
           }
           
           sqlCmd.setLength(0) ;
           updateDBDataList = new LinkedList();             
           updateDBSqlList = new LinkedList(); 
           
           String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
           File tmpFile = new File(WMdataDir+filename);        
           Date filedate = new Date(tmpFile.lastModified());
               
           FileReader  f       = new FileReader(tmpFile);
           LineNumberReader in = new LineNumberReader(f);
           
           printLog(logps,filename+"======begin==============================");
           
           int i = 0;
           doLoop://將txt檔案儲存至LinkedList
           while ((txtline = in.readLine()) != null) {
                   i++;
                   if ( i >= 4){
                       dataList = new LinkedList();
                       dataList.add(m_year);
                       dataList.add(m_month);
                       allList = Utility.getStringTokenizerData(txtline,",");
                       printLog(logps,txtline);
                       //System.out.println("allList.size()="+allList.size());
                       dataList.add((String)allList.get(0));//貸款機構別代碼
                       dataList.add((String)allList.get(1));//本年度保證案件件數
                       dataList.add((String)allList.get(2));//本年度保證案件保證金額
                       dataList.add((String)allList.get(3));//本年度保證案件融資金額
                       dataList.add((String)allList.get(4));//累計件數
                       dataList.add((String)allList.get(5));//累計保證金額
                       dataList.add((String)allList.get(6));//累計融資金額
                       dataList.add((String)allList.get(7));//保證餘額                      
                       updateDBDataList.add(dataList);//1:傳內的參數List
                   }
           }   
           in.close();
           f.close();
           printLog(logps,filename+"=======end===============================");
           sqlCmd.append(" insert into M206 values (?,?,?,?,?,?,?,?,?,?)");
           updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
           updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
           updateDBList.add(updateDBSqlList);
           //寫入合計列
           sqlCmd.setLength(0) ;
           sqlCmd.append(" insert into M206");
           sqlCmd.append(" select m_year,m_month,?,sum(guarantee_cnt_year),sum(guarantee_amt_year),");
           sqlCmd.append(" sum(loan_amt_year),sum(guarantee_cnt_sum),sum(guarantee_amt_sum),sum(loan_amt_sum),sum(guarantee_bal_year)");
           sqlCmd.append(" from M206");
           sqlCmd.append(" where m_year = ? and m_month=?");
           sqlCmd.append(" group by m_year,m_month");
          
           updateDBDataList = new LinkedList();             
           updateDBSqlList = new LinkedList(); 
           dataList = new LinkedList();
           dataList.add("0");//合計
           dataList.add(m_year);
           dataList.add(m_month);
           updateDBDataList.add(dataList);
           updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
           updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
           updateDBList.add(updateDBSqlList);
           
           if(DBManager.updateDB_ps(updateDBList)){
              errMsg = "<br>M206匯入資料庫成功";
              System.out.println("M206 Insert ok");        
              printLog(logps,"M206寫入資料庫成功");
           }
           
       }catch(Exception e){
           System.out.println("parseAcgfRpt.getM206FileData Error:"+e+e.getMessage());
           printLog(logps,"parseAcgfRpt.getM206FileData Error:"+e.getMessage());           
       }
       updateDBSqlList = null;    
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
}

