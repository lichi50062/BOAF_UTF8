package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

public class RptCG002W_M {

  /**
   * 備份DB--M組
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

        //M01_log、M02_log================================================
        String tbName21[] = {"M01_log","M02_log"};
        for(int i=0;i<tbName21.length;i++){
          tbNameNow = tbName21[i];
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

          String tmFNo = "";
          if(tbName21[i].equals("M01_log")){
            tmFNo = "GUARANTEE_ITEM_NO";
          }else{
            tmFNo = "LOAN_UNIT_NO";
          }

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,"
              + "A." + tmFNo + ","
              + "A.DATA_RANGE,A.GUARANTEE_CNT,A.LOAN_AMT,A.GUARANTEE_AMT,"
              + "A.LOAN_BAL,A.GUARANTEE_BAL,A.OVER_NOTPUSH_CNT,"
              + "A.OVER_NOTPUSH_BAL,A.OVER_OKPUSH_CNT,A.OVER_OKPUSH_BAL,"
              + "A.REPAY_TOT_CNT,A.REPAY_TOT_AMT,A.REPAY_BAL_CNT,"
              + "A.REPAY_BAL_AMT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName21[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName21[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName21[i]+"'"
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
            String dbTmpNo = RptCG002W.getTrimString(rs.getString(tmFNo));
            String dbDataRange = RptCG002W.getTrimString(rs.getString("DATA_RANGE"));
            String dbGuaranteeCnt = RptCG002W.getTrimString(rs.getString("GUARANTEE_CNT"));
            String dbLoanAmt = RptCG002W.getTrimString(rs.getString("LOAN_AMT"));
            String dbGuaranteeAmt = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT"));
            String dbLoanBal = RptCG002W.getTrimString(rs.getString("LOAN_BAL"));
            String dbGuaranteeBal = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL"));
            String dbOverNotpushCnt = RptCG002W.getTrimString(rs.getString("OVER_NOTPUSH_CNT"));
            String dbOverNotpushBal = RptCG002W.getTrimString(rs.getString("OVER_NOTPUSH_BAL"));
            String dbOverOkpushCnt = RptCG002W.getTrimString(rs.getString("OVER_OKPUSH_CNT"));
            String dbOverOkpushBal = RptCG002W.getTrimString(rs.getString("OVER_OKPUSH_BAL"));
            String dbRepayTotCnt = RptCG002W.getTrimString(rs.getString("REPAY_TOT_CNT"));
            String dbRepayTotAmt = RptCG002W.getTrimString(rs.getString("REPAY_TOT_AMT"));
            String dbRepayBalCnt = RptCG002W.getTrimString(rs.getString("REPAY_BAL_CNT"));
            String dbRepayBalAmt = RptCG002W.getTrimString(rs.getString("REPAY_BAL_AMT"));

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
            sqlCmd2 = "INSERT INTO "+tbName21[i]+"_BAK("
                + "M_YEAR,M_MONTH,"
                + tmFNo + ","
                + "DATA_RANGE,GUARANTEE_CNT,LOAN_AMT,GUARANTEE_AMT,"
                + "LOAN_BAL,GUARANTEE_BAL,OVER_NOTPUSH_CNT,"
                + "OVER_NOTPUSH_BAL,OVER_OKPUSH_CNT,OVER_OKPUSH_BAL,"
                + "REPAY_TOT_CNT,REPAY_TOT_AMT,REPAY_BAL_CNT,"
                + "REPAY_BAL_AMT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbTmpNo + "','" + dbDataRange + "',"
                + dbGuaranteeCnt + "," + dbLoanAmt + ","
                + dbGuaranteeAmt + "," + dbLoanBal + ","
                + dbGuaranteeBal + "," + dbOverNotpushCnt + ","
                + dbOverNotpushBal + "," + dbOverOkpushCnt + ","
                + dbOverOkpushBal + "," + dbRepayTotCnt + ","
                + dbRepayTotAmt + "," + dbRepayBalCnt + ","
                + dbRepayBalAmt + ","
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
                    + "VALUES('" + tbName21[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName21[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName21[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M03_log================================================
        String tbName22[] = {"M03_log"};
        for(int i=0;i<tbName22.length;i++){
          tbNameNow = tbName22[i];
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
              + "A.M_YEAR,A.M_MONTH,A.DIV_NO,A.GUARANTEE_CNT_MONTH,"
              + "A.LOAN_AMT_MONTH,A.GUARANTEE_AMT_MONTH,A.GUARANTEE_CNT_YEAR,"
              + "A.LOAN_AMT_YEAR,A.GUARANTEE_AMT_YEAR,A.GUARANTEE_BAL_TOTACC,"
              + "A.GUARANTEE_BAL_TOTACC_OVER,A.REPAY_BAL_TOTACC,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName22[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName22[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName22[i]+"'"
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
            String dbDivNo = RptCG002W.getTrimString(rs.getString("DIV_NO"));
            String dbGuaranteeCntMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_CNT_MONTH"));
            String dbLoanAmtMonth = RptCG002W.getTrimString(rs.getString("LOAN_AMT_MONTH"));
            String dbGuarnteeAmtMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_MONTH"));
            String dbGuaranteeCntYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_CNT_YEAR"));
            String dbLoanAmtYear = RptCG002W.getTrimString(rs.getString("LOAN_AMT_YEAR"));
            String dbGuaranteeAmtYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_YEAR"));
            String dbGuaranteeBalTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_TOTACC"));
            String dbGuaranteeBalTotaccOver = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_TOTACC_OVER"));
            String dbRepayBalTotacc = RptCG002W.getTrimString(rs.getString("REPAY_BAL_TOTACC"));

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
            sqlCmd2 = "INSERT INTO "+tbName22[i]+"_BAK("
                + "M_YEAR,M_MONTH,DIV_NO,GUARANTEE_CNT_MONTH,"
                + "LOAN_AMT_MONTH,GUARANTEE_AMT_MONTH,GUARANTEE_CNT_YEAR,"
                + "LOAN_AMT_YEAR,GUARANTEE_AMT_YEAR,GUARANTEE_BAL_TOTACC,"
                + "GUARANTEE_BAL_TOTACC_OVER,REPAY_BAL_TOTACC,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbDivNo + "',"
                + dbGuaranteeCntMonth + "," + dbLoanAmtMonth + ","
                + dbGuarnteeAmtMonth + "," + dbGuaranteeCntYear + ","
                + dbLoanAmtYear + "," + dbGuaranteeAmtYear + ","
                + dbGuaranteeBalTotacc + "," + dbGuaranteeBalTotaccOver + ","
                + dbRepayBalTotacc + ","
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
                    + "VALUES('" + tbName22[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName22[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName22[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M03_NOTE_log、M05_NOTE_log================================================
        String tbName23[] = {"M03_NOTE_log","M05_NOTE_log"};
        for(int i=0;i<tbName23.length;i++){
          tbNameNow = tbName23[i];
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
              + "A.M_YEAR,A.M_MONTH,A.NOTE_NO,A.NOTE_AMT_RATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName23[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName23[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName23[i]+"'"
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
            String dbNoteNo = RptCG002W.getTrimString(rs.getString("NOTE_NO"));
            String dbNoteAmtRate = RptCG002W.getTrimString(rs.getString("NOTE_AMT_RATE"));

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
            sqlCmd2 = "INSERT INTO "+tbName23[i]+"_BAK("
                + "M_YEAR,M_MONTH,NOTE_NO,NOTE_AMT_RATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbNoteNo + "'," + dbNoteAmtRate + ","
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
                    + "VALUES('" + tbName23[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName23[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName23[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M04_log================================================
        String tbName24[] = {"M04_log"};
        for(int i=0;i<tbName24.length;i++){
          tbNameNow = tbName24[i];
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
              + "A.M_YEAR,A.M_MONTH,A.LOAN_USE_NO,A.GUARANTEE_NO_MONTH,"
              + "A.GUARANTEE_NO_MONTH_P,A.LOAN_AMT_MONTH,A.LOAN_AMT_MONTH_P,"
              + "A.GUARANTEE_AMT_MONTH,A.GUARANTEE_AMT_MONTH_P,"
              + "A.GUARANTEE_NO_YEAR,A.GUARANTEE_NO_YEAR_P,A.LOAN_AMT_YEAR,"
              + "A.LOAN_AMT_YEAR_P,A.GUARANTEE_AMT_YEAR,A.GUARANTEE_AMT_YEAR_P,"
              + "A.GUARANTEE_NO_TOTACC,A.GUARANTEE_NO_TOTACC_P,"
              + "A.LOAN_AMT_TOTACC,A.LOAN_AMT_TOTACC_P,"
              + "A.GUARANTEE_AMT_TOTACC,A.GUARANTEE_AMT_TOTACC_P,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName24[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName24[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName24[i]+"'"
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
            String dbLoanUseNo = RptCG002W.getTrimString(rs.getString("LOAN_USE_NO"));
            String dbGuaranteeNoMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_MONTH"));
            String dbGuaranteeNoMonthP = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_MONTH_P"));
            String dbLoanAmtMonth = RptCG002W.getTrimString(rs.getString("LOAN_AMT_MONTH"));
            String dbLoanAmtMonthP = RptCG002W.getTrimString(rs.getString("LOAN_AMT_MONTH_P"));
            String dbGuaranteeAmtMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_MONTH"));
            String dbGuaranteeAmtMonthP = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_MONTH_P"));
            String dbGuaranteeNoYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_YEAR"));
            String dbGuaranteeNoYearP = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_YEAR_P"));
            String dbLoanAmtYear = RptCG002W.getTrimString(rs.getString("LOAN_AMT_YEAR"));
            String dbLoanAmtYearP = RptCG002W.getTrimString(rs.getString("LOAN_AMT_YEAR_P"));
            String dbGuaranteeAmtYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_YEAR"));
            String dbGuaranteeAmtYearP = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_YEAR_P"));
            String dbGuaranteeNoTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_TOTACC"));
            String dbGuaranteeNoTotaccP = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_TOTACC_P"));
            String dbLoanAmtTotacc = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC"));
            String dbLoanAmtTotaccP = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC_P"));
            String dbGuaranteeAmtTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_TOTACC"));
            String dbGuaranteeAmtTotaccP = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_TOTACC_P"));

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
            sqlCmd2 = "INSERT INTO "+tbName24[i]+"_BAK("
                + "M_YEAR,M_MONTH,LOAN_USE_NO,GUARANTEE_NO_MONTH,"
                + "GUARANTEE_NO_MONTH_P,LOAN_AMT_MONTH,LOAN_AMT_MONTH_P,"
                + "GUARANTEE_AMT_MONTH,GUARANTEE_AMT_MONTH_P,"
                + "GUARANTEE_NO_YEAR,GUARANTEE_NO_YEAR_P,LOAN_AMT_YEAR,"
                + "LOAN_AMT_YEAR_P,GUARANTEE_AMT_YEAR,GUARANTEE_AMT_YEAR_P,"
                + "GUARANTEE_NO_TOTACC,GUARANTEE_NO_TOTACC_P,"
                + "LOAN_AMT_TOTACC,LOAN_AMT_TOTACC_P,"
                + "GUARANTEE_AMT_TOTACC,GUARANTEE_AMT_TOTACC_P,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbLoanUseNo + "',"
                + dbGuaranteeNoMonth + "," + dbGuaranteeNoMonthP + ","
                + dbLoanAmtMonth + "," + dbLoanAmtMonthP + ","
                + dbGuaranteeAmtMonth + "," + dbGuaranteeAmtMonthP + ","
                + dbGuaranteeNoYear + "," + dbGuaranteeNoYearP + ","
                + dbLoanAmtYear + "," + dbLoanAmtYearP + ","
                + dbGuaranteeAmtYear + "," + dbGuaranteeAmtYearP + ","
                + dbGuaranteeNoTotacc + "," + dbGuaranteeNoTotaccP + ","
                + dbLoanAmtTotacc + "," + dbLoanAmtTotaccP + ","
                + dbGuaranteeAmtTotacc + "," + dbGuaranteeAmtTotaccP + ","
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
                    + "VALUES('" + tbName24[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName24[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName24[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M05_log================================================
        String tbName25[] = {"M05_log"};
        for(int i=0;i<tbName25.length;i++){
          tbNameNow = tbName25[i];
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
              + "A.M_YEAR,A.M_MONTH,"
              + "A.LOAN_UNIT_NO,A.PERIOD_NO,A.ITEM_NO,A.REPAY_CNT,"
              + "A.REPAY_AMT,A.RUN_NOTGOOD_CNT,A.RUN_NOTGOOD_AMT,"
              + "A.TURN_OUT_CNT,A.TURN_OUT_AMT,A.DIEASE_CNT,"
              + "A.DIEASEREPAY_AMT,A.DISASTER_CNT,A.DISASTER_AMT,"
              + "A.CORUN_OUT_CNT,A.CORUN_OUT_AMT,A.OTHER_CNT,A.OTHER_AMT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName25[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName25[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName25[i]+"'"
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
            String dbLoanUnitNo = RptCG002W.getTrimString(rs.getString("LOAN_UNIT_NO"));
            String dbPeriodNo = RptCG002W.getTrimString(rs.getString("PERIOD_NO"));
            String dbItemNo = RptCG002W.getTrimString(rs.getString("ITEM_NO"));
            String dbRepayCnt = RptCG002W.getTrimString(rs.getString("REPAY_CNT"));
            String dbRepayAmt = RptCG002W.getTrimString(rs.getString("REPAY_AMT"));
            String dbRunNotgoodCnt = RptCG002W.getTrimString(rs.getString("RUN_NOTGOOD_CNT"));
            String dbRunNotgoodAmt = RptCG002W.getTrimString(rs.getString("RUN_NOTGOOD_AMT"));
            String dbTurnOutCnt = RptCG002W.getTrimString(rs.getString("TURN_OUT_CNT"));
            String dbTurnOutAmt = RptCG002W.getTrimString(rs.getString("TURN_OUT_AMT"));
            String dbDieaseCnt = RptCG002W.getTrimString(rs.getString("DIEASE_CNT"));
            String dbDieaserepayAmt = RptCG002W.getTrimString(rs.getString("DIEASEREPAY_AMT"));
            String dbDisasterCnt = RptCG002W.getTrimString(rs.getString("DISASTER_CNT"));
            String dbDisasterAmt = RptCG002W.getTrimString(rs.getString("DISASTER_AMT"));
            String dbCorunOutCnt = RptCG002W.getTrimString(rs.getString("CORUN_OUT_CNT"));
            String dbCorunOutAmt = RptCG002W.getTrimString(rs.getString("CORUN_OUT_AMT"));
            String dbOtherCnt = RptCG002W.getTrimString(rs.getString("OTHER_CNT"));
            String dbOtherAmt = RptCG002W.getTrimString(rs.getString("OTHER_AMT"));

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
            sqlCmd2 = "INSERT INTO "+tbName25[i]+"_BAK("
                + "M_YEAR,M_MONTH,LOAN_UNIT_NO,PERIOD_NO,ITEM_NO,"
                + "REPAY_CNT,REPAY_AMT,RUN_NOTGOOD_CNT,RUN_NOTGOOD_AMT,"
                + "TURN_OUT_CNT,TURN_OUT_AMT,DIEASE_CNT,DIEASEREPAY_AMT,"
                + "DISASTER_CNT,DISASTER_AMT,CORUN_OUT_CNT,"
                + "CORUN_OUT_AMT,OTHER_CNT,OTHER_AMT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbLoanUnitNo + "','" + dbPeriodNo + "',"
                + "'" + dbItemNo + "',"
                + dbRepayCnt + "," + dbRepayAmt + ","
                + dbRunNotgoodCnt + "," + dbRunNotgoodAmt + ","
                + dbTurnOutCnt + "," + dbTurnOutAmt + ","
                + dbDieaseCnt + "," + dbDieaserepayAmt + ","
                + dbDisasterCnt + "," + dbDisasterAmt + ","
                + dbCorunOutCnt + "," + dbCorunOutAmt + ","
                + dbOtherCnt + "," + dbOtherAmt + ","
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
                    + "VALUES('" + tbName25[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName25[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName25[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M05_TOTACC_log================================================
        String tbName26[] = {"M05_TOTACC_log"};
        for(int i=0;i<tbName26.length;i++){
          tbNameNow = tbName26[i];
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
              + "A.M_YEAR,A.M_MONTH,A.LOAN_UNIT_NO,A.FIX_NO,"
              + "A.GUARANTEE_NO_TOTACC,A.GUARANTEE_AMT_TOTACC,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName26[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName26[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName26[i]+"'"
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
            String dbLoanUnitNo = RptCG002W.getTrimString(rs.getString("LOAN_UNIT_NO"));
            String dbFixNo = RptCG002W.getTrimString(rs.getString("FIX_NO"));
            String dbGuaranteeNoTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_TOTACC"),"0");
            String dbGuaranteeAmtTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_TOTACC"),"0");

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
            sqlCmd2 = "INSERT INTO "+tbName26[i]+"_BAK("
                + "M_YEAR,M_MONTH,LOAN_UNIT_NO,FIX_NO,"
                + "GUARANTEE_NO_TOTACC,GUARANTEE_AMT_TOTACC,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbLoanUnitNo + "','" + dbFixNo + "',"
                + dbGuaranteeNoTotacc + "," + dbGuaranteeAmtTotacc + ","
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
                    + "VALUES('" + tbName26[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName26[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName26[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M06_log、M07_log================================================
        String tbName27[] = {"M06_log","M07_log"};
        for(int i=0;i<tbName27.length;i++){
          tbNameNow = tbName27[i];
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
              + "A.M_YEAR,A.M_MONTH,A.AREA_NO,A.GUARANTEE_NO_MONTH,"
              + "A.GUARANTEE_AMT_MONTH,A.LOAN_AMT_MONTH,A.GUARANTEE_NO_YEAR,"
              + "A.GUARANTEE_AMT_YEAR,A.LOAN_AMT_YEAR,A.GUARANTEE_NO_TOTACC,"
              + "A.GUARANTEE_AMT_TOTACC,A.LOAN_AMT_TOTACC,A.GUARANTEE_BAL_NO,"
              + "A.GUARANTEE_BAL_AMT,A.GUARANTEE_BAL_P,A.LOAN_BAL,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName27[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName27[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName27[i]+"'"
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
            String dbAreaNo = RptCG002W.getTrimString(rs.getString("AREA_NO"));
            String dbGuaranteeNoMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_MONTH"));
            String dbGuaranteeAmtMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_MONTH"));
            String dbLoanAmtMonth = RptCG002W.getTrimString(rs.getString("LOAN_AMT_MONTH"));
            String dbGuaranteeNoYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_YEAR"));
            String dbGuaranteeAmtYear = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_YEAR"));
            String dbLoanAmtYear = RptCG002W.getTrimString(rs.getString("LOAN_AMT_YEAR"));
            String dbGuaranteeNoTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_TOTACC"));
            String dbGuaranteeAmtTotacc = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_TOTACC"));
            String dbLoanAmtTotacc = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC"));
            String dbGuaranteeBalNo = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_NO"));
            String dbGuaranteeBalAmt = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_AMT"));
            String dbGuaranteeBalP = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_P"));
            String dbLoanBal = RptCG002W.getTrimString(rs.getString("LOAN_BAL"));

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
            sqlCmd2 = "INSERT INTO "+tbName27[i]+"_BAK("
                + "M_YEAR,M_MONTH,AREA_NO,GUARANTEE_NO_MONTH,"
                + "GUARANTEE_AMT_MONTH,LOAN_AMT_MONTH,GUARANTEE_NO_YEAR,"
                + "GUARANTEE_AMT_YEAR,LOAN_AMT_YEAR,GUARANTEE_NO_TOTACC,"
                + "GUARANTEE_AMT_TOTACC,LOAN_AMT_TOTACC,GUARANTEE_BAL_NO,"
                + "GUARANTEE_BAL_AMT,GUARANTEE_BAL_P,LOAN_BAL,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ",'" + dbAreaNo + "',"
                + dbGuaranteeNoMonth + "," + dbGuaranteeAmtMonth + ","
                + dbLoanAmtMonth + "," + dbGuaranteeNoYear + ","
                + dbGuaranteeAmtYear + "," + dbLoanAmtYear + ","
                + dbGuaranteeNoTotacc + "," + dbGuaranteeAmtTotacc + ","
                + dbLoanAmtTotacc + "," + dbGuaranteeBalNo + ","
                + dbGuaranteeBalAmt + "," + dbGuaranteeBalP + ","
                + dbLoanBal + ","
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
                    + "VALUES('" + tbName27[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName27[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName27[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //M08_log================================================
        String tbName28[] = {"M08_log"};
        for(int i=0;i<tbName28.length;i++){
          tbNameNow = tbName28[i];
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
              + "A.M_YEAR,A.M_MONTH,A.ID_NO,A.DATA_RANGE,"
              + "A.GUARANTEE_NO_MONTH,A.LOAN_AMT_MONTH,"
              + "A.GUARANTEE_AMT_MONTH,A.GUARANTEE_BAL_MONTH,A.GUARANTEE_BAL_P,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName28[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName28[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName28[i]+"'"
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
            String dbIdNo = RptCG002W.getTrimString(rs.getString("ID_NO"));
            String dbDataRange = RptCG002W.getTrimString(rs.getString("DATA_RANGE"));
            String dbGuaranteeNoMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_NO_MONTH"));
            String dbLoanAmtMonth = RptCG002W.getTrimString(rs.getString("LOAN_AMT_MONTH"));
            String dbGuaranteeAmtMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_AMT_MONTH"));
            String dbGuaranteeBalMonth = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_MONTH"));
            String dbGuaranteeBalP = RptCG002W.getTrimString(rs.getString("GUARANTEE_BAL_P"));

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
            sqlCmd2 = "INSERT INTO "+tbName28[i]+"_BAK("
                + "M_YEAR,M_MONTH,ID_NO,DATA_RANGE,"
                + "GUARANTEE_NO_MONTH,LOAN_AMT_MONTH,"
                + "GUARANTEE_AMT_MONTH,GUARANTEE_BAL_MONTH,GUARANTEE_BAL_P,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ",'" + dbIdNo + "',"
                + "'" + dbDataRange + "'," + dbGuaranteeNoMonth + ","
                + dbLoanAmtMonth + "," + dbGuaranteeAmtMonth + ","
                + dbGuaranteeBalMonth + ","+ dbGuaranteeBalP + ","
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
                    + "VALUES('" + tbName28[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName28[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName28[i] + " "
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
        System.out.println("//RptCG002W_M() Have Error.....");
        e.printStackTrace();
        System.out.println(e.toString());
        System.out.println("//-------------------------------------");
      }
      return tbNameNow;
    }

}
