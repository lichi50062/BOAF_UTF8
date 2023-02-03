/*
 * Created on 2006/11/03 by ABYSS Brenda
 * 稽核清檔備份--農業發展基金及農業天然災害基金貸款報行情形表
 * 107.01.19 fix DB讀取設定檔 by 2295
 */

package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;

public class RptCG002W {
  private static boolean debug = false;     //列印debug 訊息
  private static boolean isDelete = false;  //是否刪除LOG檔的資料
  private static boolean inAbyss = false;  //資料庫連結設定

  public static void main(String args[]){
    String enCode = "Big5";
    String DB_Driver = "oracle.jdbc.driver.OracleDriver";
    String DB_Url = "";    
    String DB_Url2 = "";
    String DB_Id = "";
    String DB_Pwd = "";
    String DB_Id2 = "";
    String DB_Pwd2 = "";
    Connection conn1 = null;
    Connection conn2 = null;

    

    try{
        
        if(args.length > 0){
            DB_Url = "jdbc:oracle:thin:@"+args[0]+":1521:"+args[1];
            DB_Id = args[2];
            DB_Pwd = args[3];
            DB_Url2 = "jdbc:oracle:thin:@"+args[4]+":1521:"+args[5];
            DB_Id2 = args[6];
            DB_Pwd2 = args[7];
          }

          if(debug){
            System.out.println("DB_Url=" + DB_Url);
            System.out.println("DB_Id=" + DB_Id);
            System.out.println("DB_Pwd=" + DB_Pwd);
            System.out.println("DB_Url2=" + DB_Url2);
            System.out.println("DB_Id2=" + DB_Id2);
            System.out.println("DB_Pwd2=" + DB_Pwd2);
          }  
           
          DB_Url = Utility.getProperties("BOAFDBURL");
          DB_Url2 = Utility.getProperties_conf("JDBC_URL2");
          DB_Id =  Utility.getProperties_conf("JDBC_USER");
          DB_Pwd =  Utility.getProperties_conf("JDBC_PASSWORD");
          DB_Id2 = Utility.getProperties_conf("JDBC_USER2");
          DB_Pwd2 = Utility.getProperties_conf("JDBC_PASSWORD2");
          
        
      //取得時間
      Calendar calendar = Calendar.getInstance();
      calendar.add(2, -3);
      calendar.set(calendar.get(1), calendar.get(2), 1);
      String eYear = Integer.toString(calendar.get(Calendar.YEAR));
      String eMonth = Integer.toString(calendar.get(Calendar.MONTH) + 1);
      calendar.add(2, -2);
      calendar.set(calendar.get(1), calendar.get(2), 1);
      String sYear = Integer.toString(calendar.get(Calendar.YEAR));
      String sMonth = Integer.toString(calendar.get(Calendar.MONTH) + 1);
      if(sMonth.length() < 2) sMonth = "0" + sMonth;
      if(eMonth.length() < 2) eMonth = "0" + eMonth;
      if(debug) System.out.println("sDate="+sYear+"/"+sMonth);
      if(debug) System.out.println("eDate="+eYear+"/"+eMonth);
      conn1 = new RptCG002W().getConnection(enCode,DB_Driver,DB_Url,DB_Id,DB_Pwd);
      conn2 = new RptCG002W().getConnection(enCode,DB_Driver,DB_Url2,DB_Id2,DB_Pwd2);

      String isBakDB_A = new RptCG002W_A().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_B = new RptCG002W_B().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_E = new RptCG002W_E().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_M = new RptCG002W_M().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_W = new RptCG002W_W().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
    }catch (Exception e) {
      System.out.println("//Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
        conn1.close();
        conn2.close();
      }catch (Exception sqlEx) {
        conn1 = null;
        conn2 = null;
      }
    }
  }


  /**
   * 建立資料庫連線
   * @param DB_Url String
   * @return Connection
   */
  public Connection getConnection(String enCode, String DB_Driver,
                                  String DB_Url, String DB_Id, String DB_Pwd) {
    Connection sConn = null;
    try{
      if(sConn==null||sConn.isClosed()){
        Properties pt = new Properties();
        pt.put("characterEncoding", enCode);
        pt.put("useUnicode", "TRUE");
        if(debug) System.out.println(DB_Id+":"+DB_Pwd);
        pt.put("user", DB_Id);
        pt.put("password", DB_Pwd);
        if(debug) System.out.println("新建一個資料庫連線 " + DB_Url);
        Class.forName(DB_Driver);
        sConn = DriverManager.getConnection(DB_Url, pt);
      }
      return sConn;
    } catch(Exception e){
      System.out.println("=======DBConnection getConnection() Error =========");
      System.out.println(e);
      System.out.println("==============================================");
      return null;
    }
  }

  /**
   * 手動備份資料庫
   * @param sYear String
   * @param eYear String
   * @param sMonth String
   * @param eMonth String
   * @return String 產生錯誤的TB_NAME
   */
  public String goBakDB(String sYear,String eYear,String sMonth,String eMonth){
    Connection conn1 = null;
    Connection conn2 = null;
    String isBakDB = "";
    try {
      conn1 = (new RdbCommonDao("")).newConnection();
      conn2 = (new RdbCommonDao("BBOAFPool")).newConnection();
      sYear = Integer.toString(Integer.parseInt(sYear) + 1911);
      eYear = Integer.toString(Integer.parseInt(eYear) + 1911);
      if(sMonth.length() < 2) sMonth = "0" + sMonth;
      if(sMonth.length() < 2) eMonth = "0" + eMonth;
      //isOK = this.bakDB(sYear, eYear, sMonth, eMonth, conn1, conn2);
      String isBakDB_A = new RptCG002W_A().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_B = new RptCG002W_B().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_E = new RptCG002W_E().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_M = new RptCG002W_M().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);
      String isBakDB_W = new RptCG002W_W().bakDB(debug,isDelete,sYear,eYear,sMonth,eMonth,conn1,conn2);

      isBakDB = isBakDB_A;
      if(isBakDB_B.length() > 0){
        if (isBakDB.length() > 0){
          isBakDB += ",";
        }
        isBakDB += isBakDB_B;
      }
      if (isBakDB_E.length() > 0) {
        if (isBakDB.length() > 0) {
          isBakDB += ",";
        }
        isBakDB += isBakDB_E;
      }
      if (isBakDB_M.length() > 0) {
        if (isBakDB.length() > 0) {
          isBakDB += ",";
        }
        isBakDB += isBakDB_M;
      }
      if (isBakDB_W.length() > 0) {
        if (isBakDB.length() > 0) {
          isBakDB += ",";
        }
        isBakDB += isBakDB_W;
      }
    }catch (Exception e) {
      System.out.println("//RptCG002W() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
        conn1.close();
        conn2.close();
      }catch (Exception sqlEx) {
        conn1 = null;
        conn2 = null;
      }
    }
    return isBakDB;
  }

  /**
   * 將物件轉為字串格式，若為NULL 傳回""
   * @param str Object
   * @return String
   */
  public static String getTrimString(Object str) {
    return str != null ? ( (String) str).trim() : "";
  }

  /**
   * 將物件轉為字串格式，若為NULL 傳回0
   * @param str Object
   * @return String
   */
  public static int getTrimInt(Object str) {
    return str != null ? Integer.parseInt( ( (String) str).trim()) : 0;
  }

  /**
   * 將物件轉為Long, 如果為null 傳回0
   * @param str Object
   * @return long
   */
  public static long getTrimLong(Object str){
    return str != null ? Long.parseLong(( (String) str).trim()) : 0;
  }

  /**
   * 將物件轉為字串格式，若為NULL 傳回tmp
   * @param str Object
   * @param tmp String
   * @return String
   */
  public static String getTrimString(Object str,String tmp) {
    return str != null ? ( (String) str).trim() : tmp;
  }

}
