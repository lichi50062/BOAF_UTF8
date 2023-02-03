package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

public class RptCG002W_W {

  /**
   * 備份DB--W組
   * @param debug boolean 列印debug 訊息
   * @param isDelete boolean 是否刪除LOG檔的資料
   * @param sYear String 起始年份
   * @param eYear String 結束年份
   * @param sMonth String 起始月份
   * @param eMonth String 結束月份
   * @param conn1 Connection 原資料庫
   * @param conn2 Connection 備份資料庫
   * @return String	產生錯誤的TB_NAME
   */
  public String bakDB(boolean debug, boolean isDelete,
                            String sYear, String eYear,
                            String sMonth, String eMonth,
                            Connection conn1, Connection conn2) {
      Statement st = null;
      Statement st2 = null;
      Statement st3 = null;
      Statement st4 = null;
      ResultSet rs = null;
      String sqlCmd = null;
      String sqlCmd2 = null;
      String tbNameNow = "";

      try {
        conn1.setAutoCommit(false);     //原DB(TBOAF)
        conn2.setAutoCommit(false);     //備份DB(BBOAF)
        st = conn1.createStatement();   //查詢xxx_LOG
        st2 = conn2.createStatement();  //維護xxx_BAK
        st3 = conn1.createStatement();  //維護STATISTICS_BAK
        st4 = conn1.createStatement();  //刪除xxx_LOG

        //取得機構名稱及類別
        HashMap bankMap = new HashMap();
        sqlCmd = "SELECT BANK_NO,BANK_NAME FROM BN01";
        rs = st.executeQuery(sqlCmd);
        while (rs.next()) {
          bankMap.put(rs.getString("BANK_NO"),rs.getString("BANK_NAME"));
        }

        //WLX_APPLY_LOCK_log================================================
        String tbName30[] = {"WLX_APPLY_LOCK_log"};
        for(int i=0;i<tbName30.length;i++){
          tbNameNow = tbName30[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_QUARTER,A.BANK_CODE,A.REPORT_NO,"
              + "A.ADD_USER,A.ADD_NAME,A.ADD_DATE,A.LOCK_OWN,A.LOCK_MGR,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName30[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName30[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName30[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMQuarter = RptCG002W.getTrimString(rs.getString("M_QUARTER"));
            String dbBankCode = RptCG002W.getTrimString(rs.getString("BANK_CODE"));
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORT_NO"));
            String dbAddUser = RptCG002W.getTrimString(rs.getString("ADD_USER"));
            String dbAddName = RptCG002W.getTrimString(rs.getString("ADD_NAME"));
            Timestamp dbAddDateTimestamp = rs.getTimestamp("ADD_DATE");
            String dbAddDate = "";
            if(dbAddDateTimestamp != null){
              dbAddDate = dbAddDateTimestamp.toString();
              dbAddDate = dbAddDate.substring(0, dbAddDate.indexOf("."));
            }
            String dbLockOwn = RptCG002W.getTrimString(rs.getString("LOCK_OWN"));
            String dbLockMgr = RptCG002W.getTrimString(rs.getString("LOCK_MGR"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName30[i]+"_BAK("
                + "M_YEAR,M_QUARTER,BANK_CODE,REPORT_NO,ADD_USER,ADD_NAME,"
                + "ADD_DATE,LOCK_OWN,LOCK_MGR,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMQuarter + ","
                + "'" + dbBankCode + "','" + dbReportNo + "',"
                + "'" + dbAddUser + "','" + dbAddName + "',"
                + "TO_DATE('" + dbAddDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbLockOwn + "','" + dbLockMgr + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName30[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName30[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName30[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX_Notify_log================================================
        String tbName31[] = {"WLX_Notify_log"};
        for(int i=0;i<tbName31.length;i++){
          tbNameNow = tbName31[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.SEQ_NO,A.HEADMARK,A.NOTIFY_DATE,A.NOTIFY_END_DATE,"
              + "A.APPEND_FILE,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,A.NOTIFY_URL,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName31[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName31[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName31[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbHeadmark = RptCG002W.getTrimString(rs.getString("HEADMARK"));
            Timestamp dbNotifyDateTimestamp = rs.getTimestamp("NOTIFY_DATE");
            String dbNotifyDate = "";
            if(dbNotifyDateTimestamp != null){
              dbNotifyDate = dbNotifyDateTimestamp.toString();
              dbNotifyDate = dbNotifyDate.substring(0, dbNotifyDate.indexOf("."));
            }
            Timestamp dbNotifyEndDateTimestamp = rs.getTimestamp("NOTIFY_END_DATE");
            String dbNotifyEndDate = "";
            if(dbNotifyEndDateTimestamp != null){
              dbNotifyEndDate = dbNotifyEndDateTimestamp.toString();
              dbNotifyEndDate = dbNotifyEndDate.substring(0, dbNotifyEndDate.indexOf("."));
            }
            String dbAppendFile = RptCG002W.getTrimString(rs.getString("APPEND_FILE"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbNotifyUrl = RptCG002W.getTrimString(rs.getString("NOTIFY_URL"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName31[i]+"_BAK("
                + "SEQ_NO,HEADMARK,NOTIFY_DATE,NOTIFY_END_DATE,"
                + "APPEND_FILE,USER_ID,USER_NAME,UPDATE_DATE,NOTIFY_URL,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbSeqNo + ",'" + dbHeadmark + "',"
                + "TO_DATE('" + dbNotifyDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbNotifyEndDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbAppendFile + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbNotifyUrl + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName31[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName31[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName31[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX_S_RATE_log================================================
        String tbName32[] = {"WLX_S_RATE_log"};
        for(int i=0;i<tbName32.length;i++){
          tbNameNow = tbName32[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_QUARTER,A.BANK_NO,A.PERIOD_1_FIX_RATE,"
              + "A.PERIOD_1_VAR_RATE,A.PERIOD_3_FIX_RATE,A.PERIOD_3_VAR_RATE,"
              + "A.PERIOD_6_FIX_RATE,A.PERIOD_6_VAR_RATE,A.PERIOD_9_FIX_RATE,"
              + "A.PERIOD_9_VAR_RATE,A.PERIOD_12_FIX_RATE,A.PERIOD_12_VAR_RATE,"
              + "A.BASIC_PAY_VAR_RATE,A.PERIOD_HOUSE_VAR_RATE,A.BASE_MARK_RATE,"
              + "A.BASE_FIX_RATE,A.BASE_BASE_RATE,A.USER_ID,"
              + "A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName32[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName32[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName32[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMQuarter = RptCG002W.getTrimString(rs.getString("M_QUARTER"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbPeriod1FixRate = RptCG002W.getTrimString(rs.getString("PERIOD_1_FIX_RATE"));
            String dbPeriod1VarRate = RptCG002W.getTrimString(rs.getString("PERIOD_1_VAR_RATE"));
            String dbPeriod3FixRate = RptCG002W.getTrimString(rs.getString("PERIOD_3_FIX_RATE"));
            String dbPeriod3VarRate = RptCG002W.getTrimString(rs.getString("PERIOD_3_VAR_RATE"));
            String dbPeriod6FixRate = RptCG002W.getTrimString(rs.getString("PERIOD_6_FIX_RATE"));
            String dbPeriod6VarRate = RptCG002W.getTrimString(rs.getString("PERIOD_6_VAR_RATE"));
            String dbPeriod9FixRate = RptCG002W.getTrimString(rs.getString("PERIOD_9_FIX_RATE"));
            String dbPeriod9VarRate = RptCG002W.getTrimString(rs.getString("PERIOD_9_VAR_RATE"));
            String dbPeriod12FixRate = RptCG002W.getTrimString(rs.getString("PERIOD_12_FIX_RATE"));
            String dbPeriod12VarRate = RptCG002W.getTrimString(rs.getString("PERIOD_12_VAR_RATE"));
            String dbBasicPayVarRate = RptCG002W.getTrimString(rs.getString("BASIC_PAY_VAR_RATE"));
            String dbPeriodHouseVarRate = RptCG002W.getTrimString(rs.getString("PERIOD_HOUSE_VAR_RATE"));
            String dbBaseMarkRate = RptCG002W.getTrimString(rs.getString("BASE_MARK_RATE"));
            String dbBaseFixRate = RptCG002W.getTrimString(rs.getString("BASE_FIX_RATE"));
            String dbBaseBaseRate = RptCG002W.getTrimString(rs.getString("BASE_BASE_RATE"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName32[i]+"_BAK("
                + "M_YEAR,M_QUARTER,BANK_NO,PERIOD_1_FIX_RATE,"
                + "PERIOD_1_VAR_RATE,PERIOD_3_FIX_RATE,PERIOD_3_VAR_RATE,"
                + "PERIOD_6_FIX_RATE,PERIOD_6_VAR_RATE,PERIOD_9_FIX_RATE,"
                + "PERIOD_9_VAR_RATE,PERIOD_12_FIX_RATE,PERIOD_12_VAR_RATE,"
                + "BASIC_PAY_VAR_RATE,PERIOD_HOUSE_VAR_RATE,BASE_MARK_RATE,"
                + "BASE_FIX_RATE,BASE_BASE_RATE,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMQuarter + "," + "'" + dbBankNo + "',"
                + dbPeriod1FixRate + "," + dbPeriod1VarRate + ","
                + dbPeriod3FixRate + "," + dbPeriod3VarRate + ","
                + dbPeriod6FixRate + "," + dbPeriod6VarRate + ","
                + dbPeriod9FixRate + "," + dbPeriod9VarRate + ","
                + dbPeriod12FixRate + "," + dbPeriod12VarRate + ","
                + dbBasicPayVarRate + "," + dbPeriodHouseVarRate + ","
                + dbBaseMarkRate + "," + dbBaseFixRate + ","
                + dbBaseBaseRate + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName32[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName32[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName32[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX01_log================================================
        String tbName33[] = {"WLX01_log"};
        for(int i=0;i<tbName33.length;i++){
          tbNameNow = tbName33[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.ENGLISH,A.SETUP_APPROVAL_UNT,A.SETUP_DATE,"
              + "A.SETUP_NO,A.CHG_LICENSE_DATE,A.CHG_LICENSE_NO,"
              + "A.CHG_LICENSE_REASON,A.START_DATE,A.BUSINESS_ID,A.HSIEN_ID,"
              + "A.AREA_ID,A.ADDR,A.TELNO,A.FAX,A.EMAIL,A.WEB_SITE,"
              + "A.CENTER_FLAG,A.CENTER_NO,A.STAFF_NUM,A.IT_HSIEN_ID,"
              + "A.IT_AREA_ID,A.IT_ADDR,A.IT_NAME,A.IT_TELNO,A.AUDIT_HSIEN_ID,"
              + "A.AUDIT_AREA_ID,A.AUDIT_ADDR,A.AUDIT_NAME,A.AUDIT_TELNO,"
              + "A.FLAG,A.OPEN_DATE,A.M2_NAME,A.HSIEN_DIV_1,A.CANCEL_NO,"
              + "A.CANCEL_DATE,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,A.CREDIT_STAFF_NUM,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName33[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName33[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName33[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbEnglish = RptCG002W.getTrimString(rs.getString("ENGLISH"));
            String dbSetupApprovalUnt = RptCG002W.getTrimString(rs.getString("SETUP_Approval_Unt"));
            Timestamp tmTimestamp = rs.getTimestamp("SETUP_DATE");
            String dbSetupDate = "";
            if(tmTimestamp != null){
              dbSetupDate = tmTimestamp.toString();
              dbSetupDate = dbSetupDate.substring(0, dbSetupDate.indexOf("."));
            }
            String dbSetupNo = RptCG002W.getTrimString(rs.getString("SETUP_NO"));
            tmTimestamp = rs.getTimestamp("CHG_LICENSE_DATE");
            String dbChgLicenseDate = "";
            if(tmTimestamp != null){
              dbChgLicenseDate = tmTimestamp.toString();
              dbChgLicenseDate = dbChgLicenseDate.substring(0, dbChgLicenseDate.indexOf("."));
            }
            String dbChgLicenseNo = RptCG002W.getTrimString(rs.getString("CHG_LICENSE_NO"));
            String dbChgLicenseReason = RptCG002W.getTrimString(rs.getString("CHG_LICENSE_REASON"));
            tmTimestamp = rs.getTimestamp("START_DATE");
            String dbStartDate = "";
            if(tmTimestamp != null){
              dbStartDate = tmTimestamp.toString();
              dbStartDate = dbStartDate.substring(0, dbStartDate.indexOf("."));
            }
            String dbBusinessId = RptCG002W.getTrimString(rs.getString("BUSINESS_ID"));
            String dbHsienId = RptCG002W.getTrimString(rs.getString("HSIEN_ID"));
            String dbAreaId = RptCG002W.getTrimString(rs.getString("AREA_ID"));
            String dbAddr = RptCG002W.getTrimString(rs.getString("ADDR"));
            String dbTelno = RptCG002W.getTrimString(rs.getString("TELNO"));
            String dbFax = RptCG002W.getTrimString(rs.getString("FAX"));
            String dbEmail = RptCG002W.getTrimString(rs.getString("EMAIL"));
            String dbWebSite = RptCG002W.getTrimString(rs.getString("WEB_SITE"));
            String dbCenterFlag = RptCG002W.getTrimString(rs.getString("CENTER_FLAG"));
            String dbCenterNo = RptCG002W.getTrimString(rs.getString("CENTER_NO"));
            String dbStaffNum = RptCG002W.getTrimString(rs.getString("STAFF_NUM"),"0");
            String dbItHsienId = RptCG002W.getTrimString(rs.getString("IT_HSIEN_ID"));
            String dbItAreaId = RptCG002W.getTrimString(rs.getString("IT_AREA_ID"));
            String dbItAddr = RptCG002W.getTrimString(rs.getString("IT_ADDR"));
            String dbItName = RptCG002W.getTrimString(rs.getString("IT_NAME"));
            String dbItTelno = RptCG002W.getTrimString(rs.getString("IT_TELNO"));
            String dbAuditHsienId = RptCG002W.getTrimString(rs.getString("AUDIT_HSIEN_ID"));
            String dbAuditAreaId = RptCG002W.getTrimString(rs.getString("AUDIT_AREA_ID"));
            String dbAuditAddr = RptCG002W.getTrimString(rs.getString("AUDIT_ADDR"));
            String dbAuditName = RptCG002W.getTrimString(rs.getString("AUDIT_NAME"));
            String dbAuditTelno = RptCG002W.getTrimString(rs.getString("AUDIT_TELNO"));
            String dbFlag = RptCG002W.getTrimString(rs.getString("FLAG"));
            tmTimestamp = rs.getTimestamp("OPEN_DATE");
            String dbOpenDate = "";
            if(tmTimestamp != null){
              dbOpenDate = tmTimestamp.toString();
              dbOpenDate = dbOpenDate.substring(0, dbOpenDate.indexOf("."));
            }
            String dbM2Name = RptCG002W.getTrimString(rs.getString("M2_NAME"));
            String dbHsienDiv1 = RptCG002W.getTrimString(rs.getString("Hsien_div_1"));
            String dbCancelNo = RptCG002W.getTrimString(rs.getString("CANCEL_NO"));
            tmTimestamp = rs.getTimestamp("CANCEL_DATE");
            String dbCancelDate = "";
            if(tmTimestamp != null){
              dbCancelDate = tmTimestamp.toString();
              dbCancelDate = dbCancelDate.substring(0, dbCancelDate.indexOf("."));
            }
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbCreditStaffNum = RptCG002W.getTrimString(rs.getString("CREDIT_STAFF_NUM"),"0");


            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName33[i]+"_BAK("
                + "BANK_NO,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,"
                + "SETUP_NO,CHG_LICENSE_DATE,CHG_LICENSE_NO,"
                + "CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,"
                + "AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,"
                + "CENTER_FLAG,CENTER_NO,STAFF_NUM,IT_HSIEN_ID,"
                + "IT_AREA_ID,IT_ADDR,IT_NAME,IT_TELNO,AUDIT_HSIEN_ID,"
                + "AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,"
                + "FLAG,OPEN_DATE,M2_NAME,HSIEN_DIV_1,CANCEL_NO,"
                + "CANCEL_DATE,USER_ID,USER_NAME,UPDATE_DATE,CREDIT_STAFF_NUM,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbEnglish + "',"
                + "'" + dbSetupApprovalUnt + "',"
                + "TO_DATE('" + dbSetupDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbSetupNo + "',"
                + "TO_DATE('" + dbChgLicenseDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbChgLicenseNo + "','" + dbChgLicenseReason + "',"
                + "TO_DATE('" + dbStartDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbBusinessId + "','" + dbHsienId + "',"
                + "'" + dbAreaId + "','" + dbAddr + "','" + dbTelno + "',"
                + "'" + dbFax + "','" + dbEmail + "',"
                + "'" + dbWebSite + "','" + dbCenterFlag + "',"
                + "'" + dbCenterNo + "','" + dbStaffNum + "',"
                + "'" + dbItHsienId + "','" + dbItAreaId + "',"
                + "'" + dbItAddr + "','" + dbItName + "',"
                + "'" + dbItTelno + "','" + dbAuditHsienId + "',"
                + "'" + dbAuditAreaId + "','" + dbAuditAddr + "',"
                + "'" + dbAuditName + "','" + dbAuditTelno + "',"
                + "'" + dbFlag + "',"
                + "TO_DATE('" + dbOpenDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbM2Name + "','" + dbHsienDiv1 + "',"
                + "'" + dbCancelNo + "',"
                + "TO_DATE('" + dbCancelDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbCreditStaffNum + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName33[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName33[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName33[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX01_M_log================================================
        String tbName34[] = {"WLX01_M_log"};
        for(int i=0;i<tbName34.length;i++){
          tbNameNow = tbName34[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.SEQ_NO,A.POSITION_CODE,A.ID,A.ID_CODE,"
              + "A.NAME,A.BIRTH_DATE,A.DEGREE,A.SEX,A.TELNO,A.FAX,"
              + "A.INDUCT_DATE,A.BACKGROUND,A.CHOOSE_ITEM,A.RANK,"
              + "A.SPECIALITY,A.INCHARGE,A.ABDICATE_CODE,A.ABDICATE_DATE,"
              + "A.EMAIL,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName34[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName34[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName34[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbPositionCode = RptCG002W.getTrimString(rs.getString("POSITION_CODE"));
            String dbId = RptCG002W.getTrimString(rs.getString("ID"));
            String dbIdCode = RptCG002W.getTrimString(rs.getString("ID_CODE"));
            String dbName = RptCG002W.getTrimString(rs.getString("NAME"));
            Timestamp tmTimestamp = rs.getTimestamp("BIRTH_DATE");
            String dbBirthDate = "";
            if(tmTimestamp != null){
              dbBirthDate = tmTimestamp.toString();
              dbBirthDate = dbBirthDate.substring(0, dbBirthDate.indexOf("."));
            }
            String dbDegree = RptCG002W.getTrimString(rs.getString("DEGREE"));
            String dbSex = RptCG002W.getTrimString(rs.getString("SEX"));
            String dbTelno = RptCG002W.getTrimString(rs.getString("TELNO"));
            String dbFax = RptCG002W.getTrimString(rs.getString("FAX"));
            tmTimestamp = rs.getTimestamp("INDUCT_DATE");
            String dbInductDate = "";
            if(tmTimestamp != null){
              dbInductDate = tmTimestamp.toString();
              dbInductDate = dbInductDate.substring(0, dbInductDate.indexOf("."));
            }
            String dbBackground = RptCG002W.getTrimString(rs.getString("BACKGROUND"));
            String dbChooseItem = RptCG002W.getTrimString(rs.getString("CHOOSE_ITEM"));
            String dbRank = RptCG002W.getTrimString(rs.getString("RANK"),"0");
            String dbSpeciality = RptCG002W.getTrimString(rs.getString("SPECIALITY"));
            String dbIncharge = RptCG002W.getTrimString(rs.getString("INCHARGE"));
            String dbAbdicateCode = RptCG002W.getTrimString(rs.getString("ABDICATE_CODE"));
            tmTimestamp = rs.getTimestamp("ABDICATE_DATE");
            String dbAbdicateDate = "";
            if(tmTimestamp != null){
              dbAbdicateDate = tmTimestamp.toString();
              dbAbdicateDate = dbAbdicateDate.substring(0, dbAbdicateDate.indexOf("."));
            }
            String dbEmail = RptCG002W.getTrimString(rs.getString("EMAIL"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName34[i]+"_BAK("
                + "BANK_NO,SEQ_NO,POSITION_CODE,ID,ID_CODE,NAME,BIRTH_DATE,"
                + "DEGREE,SEX,TELNO,FAX,INDUCT_DATE,BACKGROUND,CHOOSE_ITEM,"
                + "RANK,SPECIALITY,INCHARGE,ABDICATE_CODE,ABDICATE_DATE,"
                + "EMAIL,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbSeqNo + "',"
                + "'" + dbPositionCode + "','" + dbId + "',"
                + "'" + dbIdCode + "','" + dbName + "',"
                + "TO_DATE('" + dbBirthDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbDegree + "','" + dbSex + "',"
                + "'" + dbTelno + "','" + dbFax + "',"
                + "TO_DATE('" + dbInductDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbBackground + "','" + dbChooseItem + "',"
                + dbRank + ",'" + dbSpeciality + "',"
                + "'" + dbIncharge + "','" + dbAbdicateCode + "',"
                + "TO_DATE('" + dbAbdicateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbEmail + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName34[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName34[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName34[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX02_log================================================
        String tbName35[] = {"WLX02_log"};
        for(int i=0;i<tbName35.length;i++){
          tbNameNow = tbName35[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.TBANK_NO,A.BANK_NO,A.CONST_TYPE,A.SETUP_APPROVAL_UNT,"
              + "A.SETUP_DATE,A.SETUP_NO,A.SETUP_NO_DATE,A.CHG_LICENSE_DATE,"
              + "A.CHG_LICENSE_NO,A.CHG_LICENSE_REASON,A.START_DATE,"
              + "A.HSIEN_ID,A.AREA_ID,A.ADDR,A.TELNO,A.FAX,A.EMAIL,"
              + "A.WEB_SITE,A.FLAG,A.OPEN_DATE,A.STAFF_NUM,A.HSIEN_DIV_1,"
              + "A.CANCEL_NO,A.CANCEL_DATE,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName35[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName35[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName35[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbTbankNo = RptCG002W.getTrimString(rs.getString("TBANK_NO"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbConstType = RptCG002W.getTrimString(rs.getString("CONST_TYPE"));
            String dbSetupApprovalUnt = RptCG002W.getTrimString(rs.getString("SETUP_APPROVAL_UNT"));
            Timestamp tmTimestamp = rs.getTimestamp("SETUP_DATE");
            String dbSetupDate = "";
            if(tmTimestamp != null){
              dbSetupDate = tmTimestamp.toString();
              dbSetupDate = dbSetupDate.substring(0, dbSetupDate.indexOf("."));
            }
            String dbSetupNo = RptCG002W.getTrimString(rs.getString("SETUP_NO"));
            tmTimestamp = rs.getTimestamp("SETUP_NO_DATE");
            String dbSetupNoDate = "";
            if(tmTimestamp != null){
              dbSetupNoDate = tmTimestamp.toString();
              dbSetupNoDate = dbSetupNoDate.substring(0, dbSetupNoDate.indexOf("."));
            }
            tmTimestamp = rs.getTimestamp("CHG_LICENSE_DATE");
            String dbChgLicenseDate = "";
            if(tmTimestamp != null){
              dbChgLicenseDate = tmTimestamp.toString();
              dbChgLicenseDate = dbChgLicenseDate.substring(0, dbChgLicenseDate.indexOf("."));
            }
            String dbChgLicenseNo = RptCG002W.getTrimString(rs.getString("CHG_LICENSE_NO"));
            String dbChgLicenseReason = RptCG002W.getTrimString(rs.getString("CHG_LICENSE_REASON"));
            tmTimestamp = rs.getTimestamp("START_DATE");
            String dbStartDate = "";
            if(tmTimestamp != null){
              dbStartDate = tmTimestamp.toString();
              dbStartDate = dbStartDate.substring(0, dbStartDate.indexOf("."));
            }
            String dbHsienId = RptCG002W.getTrimString(rs.getString("HSIEN_ID"));
            String dbAreaId = RptCG002W.getTrimString(rs.getString("AREA_ID"));
            String dbAddr = RptCG002W.getTrimString(rs.getString("ADDR"));
            String dbTelno = RptCG002W.getTrimString(rs.getString("TELNO"));
            String dbFax = RptCG002W.getTrimString(rs.getString("FAX"));
            String dbEmail = RptCG002W.getTrimString(rs.getString("EMAIL"));
            String dbWebSite = RptCG002W.getTrimString(rs.getString("WEB_SITE"));
            String dbFlag = RptCG002W.getTrimString(rs.getString("FLAG"));
            tmTimestamp = rs.getTimestamp("OPEN_DATE");
            String dbOpenDate = "";
            if(tmTimestamp != null){
              dbOpenDate = tmTimestamp.toString();
              dbOpenDate = dbOpenDate.substring(0, dbOpenDate.indexOf("."));
            }
            String dbStaffNum = RptCG002W.getTrimString(rs.getString("STAFF_NUM"),"0");
            String dbHsienDiv1 = RptCG002W.getTrimString(rs.getString("HSIEN_DIV_1"));
            String dbCancelNo = RptCG002W.getTrimString(rs.getString("CANCEL_NO"));
            tmTimestamp = rs.getTimestamp("CANCEL_DATE");
            String dbCancelDate = "";
            if(tmTimestamp != null){
              dbCancelDate = tmTimestamp.toString();
              dbCancelDate = dbCancelDate.substring(0, dbCancelDate.indexOf("."));
            }
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName35[i]+"_BAK("
                + "TBANK_NO,BANK_NO,CONST_TYPE,SETUP_APPROVAL_UNT,"
                + "SETUP_DATE,SETUP_NO,SETUP_NO_DATE,CHG_LICENSE_DATE,"
                + "CHG_LICENSE_NO,CHG_LICENSE_REASON,START_DATE,HSIEN_ID,"
                + "AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,FLAG,OPEN_DATE,"
                + "STAFF_NUM,HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,USER_ID,"
                + "USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbTbankNo + "','" + dbBankNo + "',"
                + "'" + dbConstType + "','" + dbSetupApprovalUnt + "',"
                + "TO_DATE('" + dbSetupDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbSetupNo + "',"
                + "TO_DATE('" + dbSetupNoDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbChgLicenseDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbChgLicenseNo + "','" + dbChgLicenseReason + "',"
                + "TO_DATE('" + dbStartDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbHsienId + "','" + dbAreaId + "',"
                + "'" + dbAddr + "','" + dbTelno + "',"
                + "'" + dbFax + "','" + dbEmail + "',"
                + "'" + dbWebSite + "','" + dbFlag + "',"
                + "TO_DATE('" + dbOpenDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + dbStaffNum + ",'" + dbHsienDiv1 + "',"
                + "'" + dbCancelNo + "',"
                + "TO_DATE('" + dbCancelDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName35[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName35[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName35[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX02_M_log================================================
        String tbName36[] = {"WLX02_M_log"};
        for(int i=0;i<tbName36.length;i++){
          tbNameNow = tbName36[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.SEQ_NO,A.POSITION_CODE,A.ID,A.ID_CODE,"
              + "A.NAME,A.BIRTH_DATE,A.DEGREE,A.SEX,A.TELNO,A.INDUCT_DATE,"
              + "A.BACKGROUND,A.CHOOSE_ITEM,A.RANK,A.ABDICATE_CODE,"
              + "A.ABDICATE_DATE,A.EMAIL,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName36[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName36[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName36[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbPositionCode = RptCG002W.getTrimString(rs.getString("POSITION_CODE"));
            String dbId = RptCG002W.getTrimString(rs.getString("ID"));
            String dbIdCode = RptCG002W.getTrimString(rs.getString("ID_CODE"));
            String dbName = RptCG002W.getTrimString(rs.getString("NAME"));
            Timestamp tmTimestamp = rs.getTimestamp("BIRTH_DATE");
            String dbBirthDate = "";
            if(tmTimestamp != null){
              dbBirthDate = tmTimestamp.toString();
              dbBirthDate = dbBirthDate.substring(0, dbBirthDate.indexOf("."));
            }
            String dbDegree = RptCG002W.getTrimString(rs.getString("DEGREE"));
            String dbSex = RptCG002W.getTrimString(rs.getString("SEX"));
            String dbTelno = RptCG002W.getTrimString(rs.getString("TELNO"));
            tmTimestamp = rs.getTimestamp("INDUCT_DATE");
            String dbInductDate = "";
            if(tmTimestamp != null){
              dbInductDate = tmTimestamp.toString();
              dbInductDate = dbInductDate.substring(0, dbInductDate.indexOf("."));
            }
            String dbBackground = RptCG002W.getTrimString(rs.getString("BACKGROUND"));
            String dbChooseItem = RptCG002W.getTrimString(rs.getString("CHOOSE_ITEM"));
            String dbRank = RptCG002W.getTrimString(rs.getString("RANK"),"0");
            String dbAbdicateCode = RptCG002W.getTrimString(rs.getString("ABDICATE_CODE"));
            tmTimestamp = rs.getTimestamp("ABDICATE_DATE");
            String dbAbdicateDate = "";
            if(tmTimestamp != null){
              dbAbdicateDate = tmTimestamp.toString();
              dbAbdicateDate = dbAbdicateDate.substring(0, dbAbdicateDate.indexOf("."));
            }
            String dbEmail = RptCG002W.getTrimString(rs.getString("EMAIL"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName36[i]+"_BAK("
                + "BANK_NO,SEQ_NO,POSITION_CODE,ID,ID_CODE,NAME,BIRTH_DATE,"
                + "DEGREE,SEX,TELNO,INDUCT_DATE,BACKGROUND,CHOOSE_ITEM,"
                + "RANK,ABDICATE_CODE,ABDICATE_DATE,EMAIL,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbSeqNo + "',"
                + "'" + dbPositionCode + "','" + dbId + "',"
                + "'" + dbIdCode + "','" + dbName + "',"
                + "TO_DATE('" + dbBirthDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbDegree + "','" + dbSex + "',"
                + "'" + dbTelno + "',"
                + "TO_DATE('" + dbInductDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbBackground + "','" + dbChooseItem + "',"
                + dbRank + ",'" + dbAbdicateCode + "','" + dbEmail + "',"
                + "TO_DATE('" + dbAbdicateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName36[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName36[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName36[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX04_log================================================
        String tbName37[] = {"WLX04_log"};
        for(int i=0;i<tbName37.length;i++){
          tbNameNow = tbName37[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.SEQ_NO,A.POSITION_CODE,A.ID,A.ID_CODE,A.NAME,"
              + "A.BIRTH_DATE,A.RANK,A.PASSPORT_AREA,A.PASSPORT_NO,"
              + "A.INDUCT_DATE,A.ABDICATE_CODE,A.ABDICATE_DATE,"
              + "A.APPOINTED_NUM,A.PERIOD_START,A.PERIOD_END,A.SEX,"
              + "A.DEGREE,A.BACKGROUND,A.PROFESSIONAL,A.TELNO,A.FINANCE_EXP,"
              + "A.EMAIL,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName37[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName37[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName37[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbPositionCode = RptCG002W.getTrimString(rs.getString("POSITION_CODE"));
            String dbId = RptCG002W.getTrimString(rs.getString("ID"));
            String dbIdCode = RptCG002W.getTrimString(rs.getString("ID_CODE"));
            String dbName = RptCG002W.getTrimString(rs.getString("NAME"));
            Timestamp tmTimestamp = rs.getTimestamp("BIRTH_DATE");
            String dbBirthDate = "";
            if(tmTimestamp != null){
              dbBirthDate = tmTimestamp.toString();
              dbBirthDate = dbBirthDate.substring(0, dbBirthDate.indexOf("."));
            }
            String dbRank = RptCG002W.getTrimString(rs.getString("RANK"),"0");
            String dbPassportArea = RptCG002W.getTrimString(rs.getString("PASSPORT_AREA"));
            String dbPassportNo = RptCG002W.getTrimString(rs.getString("PASSPORT_NO"));
            tmTimestamp = rs.getTimestamp("INDUCT_DATE");
            String dbInductDate = "";
            if(tmTimestamp != null){
              dbInductDate = tmTimestamp.toString();
              dbInductDate = dbInductDate.substring(0, dbInductDate.indexOf("."));
            }
            String dbAbdicateCode = RptCG002W.getTrimString(rs.getString("ABDICATE_CODE"));
            tmTimestamp = rs.getTimestamp("ABDICATE_DATE");
            String dbAbdicateDate = "";
            if(tmTimestamp != null){
              dbAbdicateDate = tmTimestamp.toString();
              dbAbdicateDate = dbAbdicateDate.substring(0, dbAbdicateDate.indexOf("."));
            }
            String dbAppointedNum = RptCG002W.getTrimString(rs.getString("APPOINTED_NUM"),"0");
            tmTimestamp = rs.getTimestamp("PERIOD_START");
            String dbPeriodStart = "";
            if(tmTimestamp != null){
              dbPeriodStart = tmTimestamp.toString();
              dbPeriodStart = dbPeriodStart.substring(0, dbPeriodStart.indexOf("."));
            }
            tmTimestamp = rs.getTimestamp("PERIOD_END");
            String dbPeriodEnd = "";
            if(tmTimestamp != null){
              dbPeriodEnd = tmTimestamp.toString();
              dbPeriodEnd = dbPeriodEnd.substring(0, dbPeriodEnd.indexOf("."));
            }
            String dbSex = RptCG002W.getTrimString(rs.getString("SEX"));
            String dbDegree = RptCG002W.getTrimString(rs.getString("DEGREE"));
            String dbBackground = RptCG002W.getTrimString(rs.getString("BACKGROUND"));
            String dbProfessional = RptCG002W.getTrimString(rs.getString("PROFESSIONAL"));
            String dbTelno = RptCG002W.getTrimString(rs.getString("TELNO"));
            String dbFinanceExp = RptCG002W.getTrimString(rs.getString("FINANCE_EXP"));
            String dbEmail = RptCG002W.getTrimString(rs.getString("EMAIL"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName37[i]+"_BAK("
                + "BANK_NO,SEQ_NO,POSITION_CODE,ID,ID_CODE,NAME,BIRTH_DATE,"
                + "RANK,PASSPORT_AREA,PASSPORT_NO,INDUCT_DATE,ABDICATE_CODE,"
                + "ABDICATE_DATE,APPOINTED_NUM,PERIOD_START,PERIOD_END,"
                + "SEX,DEGREE,BACKGROUND,PROFESSIONAL,TELNO,FINANCE_EXP,"
                + "EMAIL,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbSeqNo + "',"
                + "'" + dbPositionCode + "','" + dbId + "',"
                + "'" + dbIdCode + "','" + dbName + "',"
                + "TO_DATE('" + dbBirthDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + dbRank + ",'" + dbPassportArea + "',"
                + "'" + dbPassportNo + "',"
                + "TO_DATE('" + dbInductDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbAbdicateCode + "',"
                + "TO_DATE('" + dbAbdicateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + dbAppointedNum + ","
                + "TO_DATE('" + dbPeriodStart + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbPeriodEnd + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbSex + "','" + dbDegree + "',"
                + "'" + dbBackground + "','" + dbProfessional + "',"
                + "'" + dbTelno + "','" + dbFinanceExp + "',"
                + "'" + dbEmail + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName37[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName37[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName37[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }
        //WLX05_ATM_SETUP_log================================================
        String tbName38[] = {"WLX05_ATM_SETUP_log"};
        for(int i=0;i<tbName38.length;i++){
          tbNameNow = tbName38[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.SEQ_NO,A.SITE_Name,A.PROPERTY_NO,A.MACHINE_NAME,"
              + "A.HSIEN_ID,A.AREA_ID,A.ADDR,A.SETUP_DATE,A.CANCEL_TYPE,"
              + "A.CANCEL_DATE,A.COMMENT_M,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName38[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName38[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName38[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbSiteName = RptCG002W.getTrimString(rs.getString("SITE_Name"));
            String dbPropertyNo = RptCG002W.getTrimString(rs.getString("PROPERTY_NO"));
            String dbMachineName = RptCG002W.getTrimString(rs.getString("MACHINE_NAME"));
            String dbHsienId = RptCG002W.getTrimString(rs.getString("HSIEN_ID"));
            String dbAreaId = RptCG002W.getTrimString(rs.getString("AREA_ID"));
            String dbAddr = RptCG002W.getTrimString(rs.getString("ADDR"));
            Timestamp tmTimestamp = rs.getTimestamp("SETUP_DATE");
            String dbSetupDate = "";
            if(tmTimestamp != null){
              dbSetupDate = tmTimestamp.toString();
              dbSetupDate = dbSetupDate.substring(0, dbSetupDate.indexOf("."));
            }
            String dbCancelType = RptCG002W.getTrimString(rs.getString("CANCEL_TYPE"));
            tmTimestamp = rs.getTimestamp("CANCEL_DATE");
            String dbCancelDate = "";
            if(tmTimestamp != null){
              dbCancelDate = tmTimestamp.toString();
              dbCancelDate = dbCancelDate.substring(0, dbCancelDate.indexOf("."));
            }
            String dbCommentM = RptCG002W.getTrimString(rs.getString("COMMENT_M"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName38[i]+"_BAK("
                + "BANK_NO,SEQ_NO,SITE_Name,PROPERTY_NO,MACHINE_NAME,"
                + "HSIEN_ID,AREA_ID,ADDR,SETUP_DATE,CANCEL_TYPE,CANCEL_DATE,"
                + "COMMENT_M,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "'," + dbSeqNo + ","
                + "'" + dbSiteName + "','" + dbPropertyNo + "',"
                + "'" + dbMachineName + "','" + dbHsienId + "',"
                + "'" + dbAreaId + "','" + dbAddr + "',"
                + "TO_DATE('" + dbSetupDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbCancelType + "',"
                + "TO_DATE('" + dbCancelDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbCommentM + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName38[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName38[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName38[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX05_M_ATM_log================================================
        String tbName39[] = {"WLX05_M_ATM_log"};
        for(int i=0;i<tbName39.length;i++){
          tbNameNow = tbName39[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_NO,A.PUSH_DEBITCARD_CNT,"
              + "A.USE_DEBITCARD_CNT,A.PUSH_BINCARD_CNT,A.USE_BINCARD_CNT,"
              + "A.ATM_CNT,A.MONTH_TRAN_CNT,A.YEAR_ACCTRAN_CNT,"
              + "A.MONTH_TRAN_AMT,A.YEAR_ACCTRAN_AMT,A.USER_ID,A.USER_NAME,"
              + "A.UPDATE_DATE,A.CANC_DEBITCARD_CNT,A.CANC_BINCARD_CNT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName39[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName39[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName39[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbPushDebitcardCnt = RptCG002W.getTrimString(rs.getString("PUSH_DEBITCARD_CNT"));
            String dbUseDebitcardCnt = RptCG002W.getTrimString(rs.getString("USE_DEBITCARD_CNT"));
            String dbPushBincardCnt = RptCG002W.getTrimString(rs.getString("PUSH_BINCARD_CNT"));
            String dbUseBincardCnt = RptCG002W.getTrimString(rs.getString("USE_BINCARD_CNT"));
            String dbAtmCnt = RptCG002W.getTrimString(rs.getString("ATM_CNT"));
            String dbMonthTranCnt = RptCG002W.getTrimString(rs.getString("MONTH_TRAN_CNT"));
            String dbYearAcctranCnt = RptCG002W.getTrimString(rs.getString("YEAR_ACCTRAN_CNT"));
            String dbMonthTranAmt = RptCG002W.getTrimString(rs.getString("MONTH_TRAN_AMT"));
            String dbYearAcctranAmt = RptCG002W.getTrimString(rs.getString("YEAR_ACCTRAN_AMT"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbCancDebitcardCnt = RptCG002W.getTrimString(rs.getString("CANC_DEBITCARD_CNT"));
            String dbCancBincardCnt = RptCG002W.getTrimString(rs.getString("CANC_BINCARD_CNT"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName39[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_NO,PUSH_DEBITCARD_CNT,"
                + "USE_DEBITCARD_CNT,PUSH_BINCARD_CNT,USE_BINCARD_CNT,"
                + "ATM_CNT,MONTH_TRAN_CNT,YEAR_ACCTRAN_CNT,MONTH_TRAN_AMT,"
                + "YEAR_ACCTRAN_AMT,USER_ID,USER_NAME,UPDATE_DATE,"
                + "CANC_DEBITCARD_CNT,CANC_BINCARD_CNT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ",'" + dbBankNo + "',"
                + dbPushDebitcardCnt + "," + dbUseDebitcardCnt + ","
                + dbPushBincardCnt + "," + dbUseBincardCnt + ","
                + dbAtmCnt + "," + dbMonthTranCnt + ","
                + dbYearAcctranCnt + "," + dbMonthTranAmt + ","
                + dbYearAcctranAmt + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + dbCancDebitcardCnt + "," + dbCancBincardCnt + ","
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName39[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName39[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName39[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX06_M_OUTPUSH_log================================================
        String tbName40[] = {"WLX06_M_OUTPUSH_log"};
        for(int i=0;i<tbName40.length;i++){
          tbNameNow = tbName40[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.SEQ_NO,A.OUTCOMPANYNAME,A.OUTCONTRACTNAME,"
              + "A.OUTCONTRACTTEL,A.BANKCOMPLAINNAME,A.BANKCOMPLAINTEL,"
              + "A.OUTCOMMENT,A.OUT_BEGIN_DATE,A.OUT_END_DATE,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName40[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName40[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName40[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbOutcompanyname = RptCG002W.getTrimString(rs.getString("OUTCOMPANYNAME"));
            String dbOutcontractname = RptCG002W.getTrimString(rs.getString("OUTCONTRACTNAME"));
            String dbOutcontracttel = RptCG002W.getTrimString(rs.getString("OUTCONTRACTTEL"));
            String dbBankcomplainname = RptCG002W.getTrimString(rs.getString("BANKCOMPLAINNAME"));
            String dbBankcomplaintel = RptCG002W.getTrimString(rs.getString("BANKCOMPLAINTEL"));
            String dbOutcomment = RptCG002W.getTrimString(rs.getString("OUTCOMMENT"));
            Timestamp tmTimestamp = rs.getTimestamp("OUT_BEGIN_DATE");
            String dbOutBeginDate = "";
            if(tmTimestamp != null){
              dbOutBeginDate = tmTimestamp.toString();
              dbOutBeginDate = dbOutBeginDate.substring(0, dbOutBeginDate.indexOf("."));
            }
            tmTimestamp = rs.getTimestamp("OUT_END_DATE");
            String dbOutEndDate = "";
            if(tmTimestamp != null){
              dbOutEndDate = tmTimestamp.toString();
              dbOutEndDate = dbOutEndDate.substring(0, dbOutEndDate.indexOf("."));
            }
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName40[i]+"_BAK("
                + "BANK_NO,SEQ_NO,OUTCOMPANYNAME,OUTCONTRACTNAME,"
                + "OUTCONTRACTTEL,BANKCOMPLAINNAME,BANKCOMPLAINTEL,"
                + "OUTCOMMENT,OUT_BEGIN_DATE,OUT_END_DATE,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "'," + dbSeqNo + ","
                + "'" + dbOutcompanyname + "','" + dbOutcontractname + "',"
                + "'" + dbOutcontracttel + "','" + dbBankcomplainname + "',"
                + "'" + dbBankcomplaintel + "','" + dbOutcomment + "',"
                + "TO_DATE('" + dbOutBeginDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbOutEndDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName40[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName40[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName40[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX07_M_IMPORTANT_log================================================
        String tbName41[] = {"WLX07_M_IMPORTANT_log"};
        for(int i=0;i<tbName41.length;i++){
          tbNameNow = tbName41[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_NO,A.CHECKBANK_CNT,A.CHECKBANK_BAL,"
              + "A.CREDITMONTH_CNT,A.CREDITMONTH_CNT_ACC,A.CREDITMONTH_AMT,"
              + "A.CREDITMONTH_AMT_ACC,A.CHECKBANK_CNT_S,A.CHECKBANK_BAL_S,"
              + "A.CHECKBANK_CNT_N,A.CHECKBANK_BAL_N,A.CREDIT_BAL,"
              + "A.OVERCREDITMONTH_CNT,A.OVERCREDITMONTH_CNT_ACC,"
              + "A.OVERCREDITMONTH_AMT,A.OVERCREDITMONTH_AMT_ACC,"
              + "A.OVERCREDIT_BAL,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName41[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName41[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName41[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbCheckbankCnt = RptCG002W.getTrimString(rs.getString("CHECKBANK_CNT"));
            String dbCheckbankBal = RptCG002W.getTrimString(rs.getString("CHECKBANK_BAL"));
            String dbCreditmonthCnt = RptCG002W.getTrimString(rs.getString("CREDITMONTH_CNT"));
            String dbCreditmonthCntAcc = RptCG002W.getTrimString(rs.getString("CREDITMONTH_CNT_ACC"));
            String dbCreditmonthAmt = RptCG002W.getTrimString(rs.getString("CREDITMONTH_AMT"));
            String dbCreditmonthAmtAcc = RptCG002W.getTrimString(rs.getString("CREDITMONTH_AMT_ACC"));
            String dbCheckbankCntS = RptCG002W.getTrimString(rs.getString("CHECKBANK_CNT_S"));
            String dbCheckbankBalS = RptCG002W.getTrimString(rs.getString("CHECKBANK_BAL_S"));
            String dbCheckbankCntN = RptCG002W.getTrimString(rs.getString("CHECKBANK_CNT_N"));
            String dbCheckbankBalN = RptCG002W.getTrimString(rs.getString("CHECKBANK_BAL_N"));
            String dbCreditBal = RptCG002W.getTrimString(rs.getString("CREDIT_BAL"));
            String dbOvercreditmonthCnt = RptCG002W.getTrimString(rs.getString("OVERCREDITMONTH_CNT"));
            String dbOvercreditmonthCntAcc = RptCG002W.getTrimString(rs.getString("OVERCREDITMONTH_CNT_ACC"));
            String dbOvercreditmonthAmt = RptCG002W.getTrimString(rs.getString("OVERCREDITMONTH_AMT"));
            String dbOvercreditmonthAmtAcc = RptCG002W.getTrimString(rs.getString("OVERCREDITMONTH_AMT_ACC"));
            String dbOvercreditBal = RptCG002W.getTrimString(rs.getString("OVERCREDIT_BAL"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName41[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_NO,CHECKBANK_CNT,CHECKBANK_BAL,"
                + "CREDITMONTH_CNT,CREDITMONTH_CNT_ACC,CREDITMONTH_AMT,"
                + "CREDITMONTH_AMT_ACC,CHECKBANK_CNT_S,CHECKBANK_BAL_S,"
                + "CHECKBANK_CNT_N,CHECKBANK_BAL_N,CREDIT_BAL,"
                + "OVERCREDITMONTH_CNT,OVERCREDITMONTH_CNT_ACC,"
                + "OVERCREDITMONTH_AMT,OVERCREDITMONTH_AMT_ACC,"
                + "OVERCREDIT_BAL,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankNo + "'," + dbCheckbankCnt + ","
                + dbCheckbankBal + "," + dbCreditmonthCnt + ","
                + dbCreditmonthCntAcc + "," + dbCreditmonthAmt + ","
                + dbCreditmonthAmtAcc + "," + dbCheckbankCntS + ","
                + dbCheckbankBalS + "," + dbCheckbankCntN + ","
                + dbCheckbankBalN + "," + dbCreditBal + ","
                + dbOvercreditmonthCnt + "," + dbOvercreditmonthCntAcc + ","
                + dbOvercreditmonthAmt + "," + dbOvercreditmonthAmtAcc + ","
                + dbOvercreditBal + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName41[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName41[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName41[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX08_S_GAGE_APPLY_log================================================
        String tbName42[] = {"WLX08_S_GAGE_APPLY_log"};
        for(int i=0;i<tbName42.length;i++){
          tbNameNow = tbName42[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_Quarter,A.BANK_NO,A.APPLY_Cnt,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName42[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName42[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName42[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMQuarter = RptCG002W.getTrimString(rs.getString("M_Quarter"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbApplyCnt = RptCG002W.getTrimString(rs.getString("APPLY_Cnt"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName42[i]+"_BAK("
                + "M_YEAR,M_Quarter,BANK_NO,APPLY_Cnt,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMQuarter + ","
                + "'" + dbBankNo + "'," + dbApplyCnt + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName42[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName42[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName42[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX08_S_GAGE_log================================================
        String tbName43[] = {"WLX08_S_GAGE_log"};
        for(int i=0;i<tbName43.length;i++){
          tbNameNow = tbName43[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_QUARTER,A.BANK_NO,A.SEQ_NO,A.DUREASSURE_NO,"
              + "A.DEBTNAME,A.DUREDATE,A.DUREASSURESITE,A.ACCOUNTAMT,"
              + "A.APPLYDELAYYEAR,A.APPLYDELAYMONTH,A.APPLYDELAYREASON,"
              + "A.AUDIT_APPLYDELAYYEAR,A.AUDIT_APPLYDELAYMONTH,"
              + "A.AUDIT_DUREDATE,A.DAMAGE_YN,A.DISPOSAL_FACT_YN,"
              + "A.DISPOSAL_PLAN_YN,A.AUDITRESULT_YN,A.APPLYOK_DOCNO,"
              + "A.APPLYOK_DATE,A.REPORT_BOAF_DOCNO,A.REPORT_BOAF_DATE,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName43[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName43[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName43[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMQuarter = RptCG002W.getTrimString(rs.getString("M_QUARTER"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbSeqNo = RptCG002W.getTrimString(rs.getString("SEQ_NO"));
            String dbDureassureNo = RptCG002W.getTrimString(rs.getString("DUREASSURE_NO"));
            String dbDebtname = RptCG002W.getTrimString(rs.getString("DEBTNAME"));
            Timestamp tmTimestamp = rs.getTimestamp("DUREDATE");
            String dbDuredate = "";
            if(tmTimestamp != null){
              dbDuredate = tmTimestamp.toString();
              dbDuredate = dbDuredate.substring(0, dbDuredate.indexOf("."));
            }
            String dbDureassuresite = RptCG002W.getTrimString(rs.getString("DUREASSURESITE"));
            String dbAccountamt = RptCG002W.getTrimString(rs.getString("ACCOUNTAMT"));
            String dbApplydelayyear = RptCG002W.getTrimString(rs.getString("APPLYDELAYYEAR"));
            String dbApplydelaymonth = RptCG002W.getTrimString(rs.getString("APPLYDELAYMONTH"));
            String dbApplydelayreason = RptCG002W.getTrimString(rs.getString("APPLYDELAYREASON"));
            String dbAuditApplydelayyear = RptCG002W.getTrimString(rs.getString("AUDIT_APPLYDELAYYEAR"));
            String dbAuditApplydelaymonth = RptCG002W.getTrimString(rs.getString("AUDIT_APPLYDELAYMONTH"));
            tmTimestamp = rs.getTimestamp("AUDIT_DUREDATE");
            String dbAuditDuredate = "";
            if(tmTimestamp != null){
              dbAuditDuredate = tmTimestamp.toString();
              dbAuditDuredate = dbAuditDuredate.substring(0, dbAuditDuredate.indexOf("."));
            }
            String dbDamageYn = RptCG002W.getTrimString(rs.getString("DAMAGE_YN"));
            String dbDisposalFactYn = RptCG002W.getTrimString(rs.getString("DISPOSAL_FACT_YN"));
            String dbDisposalPlanYn = RptCG002W.getTrimString(rs.getString("DISPOSAL_PLAN_YN"));
            String dbAuditresultYn = RptCG002W.getTrimString(rs.getString("AUDITRESULT_YN"));
            String dbApplyokDocno = RptCG002W.getTrimString(rs.getString("APPLYOK_DOCNO"));
            tmTimestamp = rs.getTimestamp("APPLYOK_DATE");
            String dbApplyokDate = "";
            if(tmTimestamp != null){
              dbApplyokDate = tmTimestamp.toString();
              dbApplyokDate = dbApplyokDate.substring(0, dbApplyokDate.indexOf("."));
            }
            String dbReportBoafDocno = RptCG002W.getTrimString(rs.getString("REPORT_BOAF_DOCNO"));
            tmTimestamp = rs.getTimestamp("REPORT_BOAF_DATE");
            String dbReportBoafDate = "";
            if(tmTimestamp != null){
              dbReportBoafDate = tmTimestamp.toString();
              dbReportBoafDate = dbReportBoafDate.substring(0, dbReportBoafDate.indexOf("."));
            }
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName43[i]+"_BAK("
                + "M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,"
                + "DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,"
                + "APPLYDELAYYEAR,APPLYDELAYMONTH,APPLYDELAYREASON,"
                + "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"
                + "DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,AUDITRESULT_YN,"
                + "APPLYOK_DOCNO,APPLYOK_DATE,REPORT_BOAF_DOCNO,REPORT_BOAF_DATE,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMQuarter + ","
                + "'" + dbBankNo + "'," + dbSeqNo + ","
                + dbDureassureNo + ",'" + dbDebtname + "',"
                + "TO_DATE('" + dbDuredate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbDureassuresite + "'," + dbAccountamt + ","
                + dbApplydelayyear + "," + dbApplydelaymonth + ","
                + "'" + dbApplydelayreason + "'," + dbAuditApplydelayyear + ","
                + dbAuditApplydelaymonth + ","
                + "TO_DATE('" + dbAuditDuredate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbDamageYn + "','" + dbDisposalFactYn + "',"
                + "'" + dbDisposalPlanYn + "','" + dbAuditresultYn + "',"
                + "'" + dbApplyokDocno + "',"
                + "TO_DATE('" + dbApplyokDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbReportBoafDocno + "',"
                + "TO_DATE('" + dbReportBoafDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName43[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName43[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName43[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WLX09_S_WARNING_log================================================
        String tbName44[] = {"WLX09_S_WARNING_log"};
        for(int i=0;i<tbName44.length;i++){
          tbNameNow = tbName44[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_QUARTER,A.BANK_NO,A.WARNACCOUNT_TCNT,"
              + "A.WARNACCOUNT_TBAL,A.WARNACCOUNT_REMIT_TCNT,"
              + "A.WARNACCOUNT_REFUND_APPLY_CNT,A.WARNACCOUNT_REFUND_APPLY_AMT,"
              + "A.WARNACCOUNT_REFUND_CNT,A.WARNACCOUNT_REFUND_AMT,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName44[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName44[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName44[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMQuarter = RptCG002W.getTrimString(rs.getString("M_QUARTER"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbWarnaccountTcnt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_TCNT"));
            String dbWarnaccountTbal = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_TBAL"));
            String dbWarnaccountRemitTcnt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_REMIT_TCNT"));
            String dbWarnaccountRefundApplyCnt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_REFUND_APPLY_CNT"));
            String dbWarnaccountRefundApplyAmt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_REFUND_APPLY_AMT"));
            String dbWarnaccountRefundCnt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_REFUND_CNT"));
            String dbWarnaccountRefundAmt = RptCG002W.getTrimString(rs.getString("WARNACCOUNT_REFUND_AMT"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName44[i]+"_BAK("
                + "M_YEAR,M_QUARTER,BANK_NO,WARNACCOUNT_TCNT,"
                + "WARNACCOUNT_TBAL,WARNACCOUNT_REMIT_TCNT,"
                + "WARNACCOUNT_REFUND_APPLY_CNT,WARNACCOUNT_REFUND_APPLY_AMT,"
                + "WARNACCOUNT_REFUND_CNT,WARNACCOUNT_REFUND_AMT,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMQuarter + ","
                + "'" + dbBankNo + "'," + dbWarnaccountTcnt + ","
                + dbWarnaccountTbal + "," + dbWarnaccountRemitTcnt + ","
                + dbWarnaccountRefundApplyCnt + ","
                + dbWarnaccountRefundApplyAmt + ","
                + dbWarnaccountRefundCnt + "," + dbWarnaccountRefundAmt + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName44[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName44[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName44[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WML01_log================================================
        String tbName45[] = {"WML01_log"};
        for(int i=0;i<tbName45.length;i++){
          tbNameNow = tbName45[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.REPORT_NO,A.INPUT_METHOD,"
              + "A.ADD_USER,A.ADD_NAME,A.ADD_DATE,A.COMMON_CENTER,"
              + "A.UPD_METHOD,A.UPD_CODE,A.BATCH_NO,A.LOCK_STATUS,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName45[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName45[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName45[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankCode = RptCG002W.getTrimString(rs.getString("BANK_CODE"));
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORT_NO"));
            String dbInputMethod = RptCG002W.getTrimString(rs.getString("INPUT_METHOD"));
            String dbAddUser = RptCG002W.getTrimString(rs.getString("ADD_USER"));
            String dbAddName = RptCG002W.getTrimString(rs.getString("ADD_NAME"));
            Timestamp tmTimestamp = rs.getTimestamp("ADD_DATE");
            String dbAddDate = "";
            if(tmTimestamp != null){
              dbAddDate = tmTimestamp.toString();
              dbAddDate = dbAddDate.substring(0, dbAddDate.indexOf("."));
            }
            String dbCommonCenter = RptCG002W.getTrimString(rs.getString("COMMON_CENTER"));
            String dbUpdMethod = RptCG002W.getTrimString(rs.getString("UPD_METHOD"));
            String dbUpdCode = RptCG002W.getTrimString(rs.getString("UPD_CODE"));
            String dbBatchNo = RptCG002W.getTrimString(rs.getString("BATCH_NO"));
            String dbLockStatus = RptCG002W.getTrimString(rs.getString("LOCK_STATUS"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName45[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_CODE,REPORT_NO,INPUT_METHOD,"
                + "ADD_USER,ADD_NAME,ADD_DATE,COMMON_CENTER,"
                + "UPD_METHOD,UPD_CODE,BATCH_NO,LOCK_STATUS,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankCode + "','" + dbReportNo + "',"
                + "'" + dbInputMethod + "','" + dbAddUser + "',"
                + "'" + dbAddName + "',"
                + "TO_DATE('" + dbAddDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbCommonCenter + "','" + dbUpdMethod + "',"
                + "'" + dbUpdCode + "'," + dbBatchNo + ","
                + "'" + dbLockStatus + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName45[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName45[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName45[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WML02_log================================================
        String tbName46[] = {"WML02_log"};
        for(int i=0;i<tbName46.length;i++){
          tbNameNow = tbName46[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.REPORT_NO,"
              + "A.CANO,A.L_AMT,A.R_AMT,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName46[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName46[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName46[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankCode = RptCG002W.getTrimString(rs.getString("BANK_CODE"));
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORT_NO"));
            String dbCano = RptCG002W.getTrimString(rs.getString("CANO"));
            String dbLAmt = RptCG002W.getTrimString(rs.getString("L_AMT"),"0");
            String dbRAmt = RptCG002W.getTrimString(rs.getString("R_AMT"),"0");
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName46[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_CODE,REPORT_NO,"
                + "CANO,L_AMT,R_AMT,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankCode + "','" + dbReportNo + "',"
                + "'" + dbCano + "',"
                + dbLAmt + "," + dbRAmt + ","
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName46[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName46[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName46[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WML03_log================================================
        String tbName47[] = {"WML03_log"};
        for(int i=0;i<tbName47.length;i++){
          tbNameNow = tbName47[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.REPORT_NO,"
              + "A.SERIAL_NO,A.REMARK,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName47[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName47[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName47[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankCode = RptCG002W.getTrimString(rs.getString("BANK_CODE"));
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORT_NO"));
            String dbSerialNo = RptCG002W.getTrimString(rs.getString("SERIAL_NO"));
            String dbRemark = RptCG002W.getTrimString(rs.getString("REMARK"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName47[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_CODE,REPORT_NO,SERIAL_NO,REMARK,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankCode + "','" + dbReportNo + "',"
                + dbSerialNo + ",'" + dbRemark + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName47[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName47[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName47[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WTT07_ELM_log================================================
        String tbName48[] = {"WTT07_ELM_log"};
        for(int i=0;i<tbName48.length;i++){
          tbNameNow = tbName48[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.MUSER_ID,A.MUSER_PASSWORD,A.IP_ADDRESS,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName48[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName48[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName48[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMuserId = RptCG002W.getTrimString(rs.getString("MUSER_ID"));
            String dbMuserPassword = RptCG002W.getTrimString(rs.getString("MUSER_PASSWORD"));
            String dbIpAddress = RptCG002W.getTrimString(rs.getString("IP_ADDRESS"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName48[i]+"_BAK("
                + "MUSER_ID,MUSER_PASSWORD,IP_ADDRESS,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbMuserId + "','" + dbMuserPassword + "',"
                + "'" + dbIpAddress + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName48[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName48[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName48[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WZZ07_log================================================
        String tbName49[] = {"WZZ07_log"};
        for(int i=0;i<tbName49.length;i++){
          tbNameNow = tbName49[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.BANK_NO,A.PROGRAM_ID,A.MAINTAIN_ID,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName49[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName49[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName49[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbProgramId = RptCG002W.getTrimString(rs.getString("PROGRAM_ID"));
            String dbMaintainId = RptCG002W.getTrimString(rs.getString("MAINTAIN_ID"));
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName49[i]+"_BAK("
                + "BANK_NO,PROGRAM_ID,MAINTAIN_ID,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbProgramId + "',"
                + "'" + dbMaintainId + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName49[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName49[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName49[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WTT01_log=======================================================
        String tbName50[] = {"WTT01_log"};
        for(int i=0;i<tbName50.length;i++){
          tbNameNow = tbName50[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.MUSER_ID,A.MUSER_NAME,A.MUSER_PASSWORD,A.MUSER_I_O,"
              + "A.BANK_TYPE,A.TBANK_NO,A.BANK_NO,A.SUBDEP_ID,A.ADD_USER,"
              + "A.ADD_NAME,A.ADD_DATE,A.FIRSTLOGIN_MARK,A.LOCK_MARK,"
              + "A.DELETE_MARK,A.MUSER_TYPE,A.PASSWORD_UPDATE_DATE,A.PASSWORD_PRE,"
              + "A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName50[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName50[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName50[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMuserId = rs.getString("MUSER_ID");
            String dbMuserName = rs.getString("MUSER_NAME");
            String dbMuserPassword = rs.getString("MUSER_PASSWORD");
            String dbMuserIO = rs.getString("MUSER_I_O");
            String dbBankType = rs.getString("BANK_TYPE");
            String dbTbankNo = rs.getString("TBANK_NO");
            String dbBankNo = rs.getString("BANK_NO");
            String dbSubdepId = RptCG002W.getTrimString(rs.getString("SUBDEP_ID"));
            String dbAddUser = RptCG002W.getTrimString(rs.getString("ADD_USER"));
            String dbAddName = RptCG002W.getTrimString(rs.getString("ADD_NAME"));
            Timestamp tmTimestamp = rs.getTimestamp("ADD_DATE");
            String dbAddDate = "";
            if(tmTimestamp != null){
              dbAddDate = tmTimestamp.toString();
              dbAddDate = dbAddDate.substring(0, dbAddDate.indexOf("."));
            }
            String dbFirstloginMark = RptCG002W.getTrimString(rs.getString("FIRSTLOGIN_MARK"));
            String dbLockMark = RptCG002W.getTrimString(rs.getString("LOCK_MARK"));
            String dbDeleteMark = RptCG002W.getTrimString(rs.getString("DELETE_MARK"));
            String dbMuserType = RptCG002W.getTrimString(rs.getString("MUSER_TYPE"));
            tmTimestamp = rs.getTimestamp("PASSWORD_UPDATE_DATE");
            String dbPasswordUpdateDate = "";
            if(tmTimestamp != null){
              dbPasswordUpdateDate = tmTimestamp.toString();
              dbPasswordUpdateDate = dbPasswordUpdateDate.substring(0, dbPasswordUpdateDate.indexOf("."));
            }
            String dbPasswordPre = RptCG002W.getTrimString(rs.getString("PASSWORD_PRE"));
            String dbUserId = RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName = RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName50[i]+"_BAK("
                + "MUSER_ID,MUSER_NAME,MUSER_PASSWORD,MUSER_I_O,BANK_TYPE,"
                + "TBANK_NO,BANK_NO,SUBDEP_ID,ADD_USER,ADD_NAME,ADD_DATE,"
                + "FIRSTLOGIN_MARK,LOCK_MARK,DELETE_MARK,"
                + "MUSER_TYPE,PASSWORD_UPDATE_DATE,PASSWORD_PRE,"
                + "USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbMuserId + "','" + dbMuserName + "',"
                + "'" + dbMuserPassword + "','" + dbMuserIO + "',"
                + "'" + dbBankType + "','" + dbTbankNo + "',"
                + "'" + dbBankNo + "','" + dbSubdepId + "',"
                + "'" + dbAddUser + "','" + dbAddName + "',"
                + "TO_DATE('" + dbAddDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbFirstloginMark + "','" + dbLockMark + "',"
                + "'" + dbDeleteMark + "','" + dbMuserType + "',"
                + "TO_DATE('" + dbPasswordUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbPasswordPre + "',"
                + "'" + dbUserId + "','" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName50[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName50[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName50[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WTT07=======================================================
        String tbName51[] = {"WTT07"};
        for(int i=0;i<tbName51.length;i++){
          tbNameNow = tbName51[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.INPUT_DATE,A.MUSER_ID,A.MUSER_PASSWORD,"
              + "A.IP_ADDRESS,A.RESULT_P,"
              + "A.MUSER_ID AS USER_ID_C,W.MUSER_NAME AS USER_NAME_C,"
              + "A.INPUT_DATE AS UPDATE_DATE_C,"
              + "TO_CHAR(A.INPUT_DATE,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.INPUT_DATE,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName51[i]+" A ,WTT01 W "
              + "WHERE A.MUSER_ID = W.MUSER_ID "
              + "AND TO_CHAR(A.INPUT_DATE,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.INPUT_DATE,'yyyy'),"
              + "TO_CHAR(A.INPUT_DATE,'mm'),"
              + "A.MUSER_ID,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName51[i]+"_BAK "
                  + "WHERE TO_CHAR(INPUT_DATE,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName51[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            Timestamp tmTimestamp = rs.getTimestamp("INPUT_DATE");
            String dbInputDate = "";
            if(tmTimestamp != null){
              dbInputDate = tmTimestamp.toString();
              dbInputDate = dbInputDate.substring(0, dbInputDate.indexOf("."));
            }
            String dbMuserId = RptCG002W.getTrimString(rs.getString("MUSER_ID"));
            String dbMuserPassword = RptCG002W.getTrimString(rs.getString("MUSER_PASSWORD"));
            String dbIpAddress = RptCG002W.getTrimString(rs.getString("IP_ADDRESS"));
            String dbResultP = RptCG002W.getTrimString(rs.getString("RESULT_P"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType = rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName51[i]+"_BAK("
                + "INPUT_DATE,MUSER_ID,MUSER_PASSWORD,IP_ADDRESS,RESULT_P) "
                + "VALUES("
                + "TO_DATE('" + dbInputDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbMuserId + "','" + dbMuserPassword + "',"
                + "'" + dbIpAddress + "','" + dbResultP + "')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName51[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbResultP.equals("X00")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName51[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName51[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //WTT06=======================================================
        String tbName52[] = {"WTT06"};
        for(int i=0;i<tbName52.length;i++){
          tbNameNow = tbName52[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.SERIALNO,A.INPUT_DATE,A.MUSER_ID,A.TYPE,A.IP_ADDRESS,"
              + "A.MUSER_ID AS USER_ID_C,W.MUSER_NAME AS USER_NAME_C,"
              + "A.INPUT_DATE AS UPDATE_DATE_C,"
              + "TO_CHAR(A.INPUT_DATE,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.INPUT_DATE,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName52[i]+" A ,WTT01 W "
              + "WHERE A.MUSER_ID = W.MUSER_ID "
              + "AND TO_CHAR(A.INPUT_DATE,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.INPUT_DATE,'yyyy'),"
              + "TO_CHAR(A.INPUT_DATE,'mm'),"
              + "A.MUSER_ID,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName52[i]+"_BAK "
                  + "WHERE TO_CHAR(INPUT_DATE,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName52[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbSerialno = RptCG002W.getTrimString(rs.getString("SERIALNO"));
            Timestamp tmTimestamp = rs.getTimestamp("INPUT_DATE");
            String dbInputDate = "";
            if(tmTimestamp != null){
              dbInputDate = tmTimestamp.toString();
              dbInputDate = dbInputDate.substring(0, dbInputDate.indexOf("."));
            }
            String dbMuserId = RptCG002W.getTrimString(rs.getString("MUSER_ID"));
            String dbType = RptCG002W.getTrimString(rs.getString("TYPE"));
            String dbIpAddress = RptCG002W.getTrimString(rs.getString("IP_ADDRESS"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType = rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName52[i]+"_BAK("
                + "SERIALNO,INPUT_DATE,MUSER_ID,TYPE,IP_ADDRESS) "
                + "VALUES("
                + "'" + dbSerialno + "',"
                + "TO_DATE('" + dbInputDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbMuserId + "','" + dbType + "',"
                + "'" + dbIpAddress + "')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName52[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                loginNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
              loginNum++;
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName52[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName52[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        conn1.commit();
        conn2.commit();
        tbNameNow = "";
      }catch (Exception e) {
      System.out.println("tbNameNow = " + tbNameNow);
        System.out.println("//RptCG002W_W() Have Error.....");
        e.printStackTrace();
        System.out.println(e.toString());
        System.out.println("//-------------------------------------");
      }
      return tbNameNow;
    }

}
