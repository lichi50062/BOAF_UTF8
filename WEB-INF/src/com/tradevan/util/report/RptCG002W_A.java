package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

public class RptCG002W_A {

  /**
   * 備份DB--A組
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

        //A01_LOG,A02_LOG,A03_LOG,A04_LOG,A99_LOG============================
        String tbName[] = {"A01_LOG","A02_LOG","A03_LOG","A04_LOG","A99_LOG"};
        for(int i=0;i<tbName.length;i++){
          tbNameNow = tbName[i];
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
          sqlCmd = "SELECT A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.ACC_CODE,A.AMT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName[i]+" A,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),TO_CHAR(A.UPDATE_DATE_C,'mm')"
              + ",A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName[i]+"'"
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
            String dbAccCode = RptCG002W.getTrimString(rs.getString("ACC_CODE"));
            long dbAmt = RptCG002W.getTrimLong(rs.getString("AMT"));
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
            sqlCmd2 =
                "INSERT INTO "+tbName[i]+"_BAK(M_YEAR,M_MONTH,BANK_CODE,ACC_CODE,"
                + "AMT,USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES(" + dbMYear + "," + dbMMonth + ",'" + dbBankCode
                + "','" + dbAccCode + "'," + dbAmt + ",'" + dbUserIdC
                + "','" + dbUserNameC +"',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),'"
                + dbUpdateTypeC + "')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId + bankNo).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //A05_L0G=============================================================
        String tbName2[] = {"A05_LOG"};
        for(int i=0;i<tbName2.length;i++){
          tbNameNow = tbName2[i];
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
          sqlCmd = "SELECT A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.ACC_CODE,A.AMT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO,"
              + "A.AMT_NAME "
              + "FROM "+tbName2[i]+" A,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),TO_CHAR(A.UPDATE_DATE_C,'mm')"
              + ",A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName2[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName2[i]+"'"
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
            String dbAccCode = RptCG002W.getTrimString(rs.getString("ACC_CODE"));
            long dbAmt = RptCG002W.getTrimLong(rs.getString("AMT"));
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
            String dbAmtName = RptCG002W.getTrimString(rs.getString("AMT_NAME"));

            //新增_BAK的TABLE中
            sqlCmd2 =
                "INSERT INTO "+tbName2[i]+"_BAK(M_YEAR,M_MONTH,BANK_CODE,ACC_CODE,"
                + "AMT,USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C,AMT_NAME) "
                + "VALUES(" + dbMYear + "," + dbMMonth + ",'" + dbBankCode
                + "','" + dbAccCode + "'," + dbAmt + ",'" + dbUserIdC
                + "','" + dbUserNameC +"',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),'"
                + dbUpdateTypeC + "','"+dbAmtName+"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId + bankNo).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName2[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName2[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName2[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //A06_L0G=============================================================
        String tbName3[] = {"A06_LOG"};
        for(int i=0;i<tbName3.length;i++){
          tbNameNow = tbName3[i];
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
          sqlCmd = "SELECT A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.ACC_CODE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO,"
              + "A.AMT_3MONTH,A.AMT_6MONTH,A.AMT_1YEAR,A.AMT_2YEAR,"
              + "A.AMT_OVER2YEAR,A.AMT_TOTAL "
              + "FROM "+tbName3[i]+" A,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),TO_CHAR(A.UPDATE_DATE_C,'mm')"
              + ",A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName3[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName3[i]+"'"
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
            String dbAccCode = RptCG002W.getTrimString(rs.getString("ACC_CODE"));
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

            long dbAmt3Month = RptCG002W.getTrimLong(rs.getString("AMT_3MONTH"));
            long dbAmt6Month = RptCG002W.getTrimLong(rs.getString("AMT_6MONTH"));
            long dbAmt1Year = RptCG002W.getTrimLong(rs.getString("AMT_1YEAR"));
            long dbAmt2Year = RptCG002W.getTrimLong(rs.getString("AMT_2YEAR"));
            long dbAmtOver2Year = RptCG002W.getTrimLong(rs.getString("AMT_OVER2YEAR"));
            String dbAmtTotal = RptCG002W.getTrimString(rs.getString("AMT_TOTAL"));

            //新增_BAK的TABLE中
            sqlCmd2 =
                "INSERT INTO "+tbName3[i]+"_BAK(M_YEAR,M_MONTH,BANK_CODE,ACC_CODE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C,"
                + "AMT_3MONTH,AMT_6MONTH,AMT_1YEAR,AMT_2YEAR,"
                + "AMT_OVER2YEAR,AMT_TOTAL) "
                + "VALUES(" + dbMYear + "," + dbMMonth + ",'" + dbBankCode
                + "','" + dbAccCode + "','" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),'"
                + dbUpdateTypeC + "'," + dbAmt3Month + "," + dbAmt6Month
                + "," + dbAmt1Year + "," + dbAmt2Year + "," + dbAmtOver2Year
                + "," + dbAmtTotal + ")";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId + bankNo).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName3[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName3[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName3[i] + " "
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
        System.out.println("//RptCG002W_A() Have Error.....");
        e.printStackTrace();
        System.out.println(e.toString());
        System.out.println("//-------------------------------------");
      }

      return tbNameNow;
    }

}
