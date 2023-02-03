package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

public class RptCG002W_B {

  /**
   * 備份DB--B組
   * @param debug boolean 列印debug 訊息
   * @param isDelete boolean 是否刪除LOG檔的資料
   * @param sYear String 起始年份
   * @param eYear String 結束年份
   * @param sMonth String 起始月份
   * @param eMonth String 結束月份
   * @param conn1 Connection 原資料庫
   * @param conn2 Connection 備份資料庫
   * @return String 產生錯誤的TB_NAME
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

        //B01_log=============================================================
        String tbName4[] = {"B01_log"};
        for(int i=0;i<tbName4.length;i++){
          tbNameNow = tbName4[i];
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
              + "A.M_YEAR,A.M_MONTH,A.FUND_MASTER_NO,A.FUND_SUB_NO,"
              + "A.FUND_NEXT_NO,A.BUDGET_AMT,A.CREDIT_PAY_AMT,"
              + "A.CREDIT_PAY_RATE,A.REMARK,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName4[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName4[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName4[i]+"'"
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
            String dbFundMasterNo = RptCG002W.getTrimString(rs.getString("FUND_MASTER_NO"));
            String dbFundSubNo = RptCG002W.getTrimString(rs.getString("FUND_SUB_NO"));
            String dbFundNextNo = RptCG002W.getTrimString(rs.getString("FUND_NEXT_NO"));
            String dbBudgetAmt = RptCG002W.getTrimString(rs.getString("BUDGET_AMT"));
            String dbCreditPayAmt = RptCG002W.getTrimString(rs.getString("CREDIT_PAY_AMT"));
            String dbCreditPayRate = RptCG002W.getTrimString(rs.getString("CREDIT_PAY_RATE"));
            String dbRemark = RptCG002W.getTrimString(rs.getString("REMARK"));

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
            sqlCmd2 = "INSERT INTO "+tbName4[i]+"_BAK("
                + "M_YEAR,M_MONTH,FUND_MASTER_NO,FUND_SUB_NO,"
                + "FUND_NEXT_NO,BUDGET_AMT,CREDIT_PAY_AMT,"
                + "CREDIT_PAY_RATE,REMARK,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbFundMasterNo + "','" + dbFundSubNo + "',"
                + "'" + dbFundNextNo + "'," + dbBudgetAmt + ","
                + dbCreditPayAmt + "," + dbCreditPayRate + ","
                + "'" + dbRemark + "',"
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
                    + "VALUES('" + tbName4[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName4[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName4[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //B02_log=============================================================
        String tbName5[] = {"B02_log"};
        for(int i=0;i<tbName5.length;i++){
          tbNameNow = tbName5[i];
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
              + "A.M_YEAR,A.M_MONTH,A.RUN_MASTER_NO,A.RUN_SUB_NO,"
              + "A.RUN_NEXT_NO,A.LOAN_CNT_YEAR,A.LOAN_AMT_YEAR,"
              + "A.LOAN_CNT_TOTACC,A.LOAN_AMT_TOTACC,A.LOAN_CNT_BAL,"
              + "A.LOAN_AMT_BAL_SUBTOT,A.LOAN_AMT_BAL_FUND,A.LOAN_AMT_BAL_BANK,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName5[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName5[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName5[i]+"'"
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
            String dbRunMasterNo = RptCG002W.getTrimString(rs.getString("RUN_MASTER_NO"));
            String dbRunSubNo = RptCG002W.getTrimString(rs.getString("RUN_SUB_NO"));
            String dbRunNextNo = RptCG002W.getTrimString(rs.getString("RUN_NEXT_NO"));
            String dbLoanCntYear = RptCG002W.getTrimString(rs.getString("LOAN_CNT_YEAR"));
            String dbLoanAmtYear = RptCG002W.getTrimString(rs.getString("LOAN_AMT_YEAR"));
            String dbLoanCntTotacc = RptCG002W.getTrimString(rs.getString("LOAN_CNT_TOTACC"));
            String dbLoanAmtTotacc = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC"));
            String dbLoanCntBal = RptCG002W.getTrimString(rs.getString("LOAN_CNT_BAL"));
            String dbLoanAmtBalSubtot = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_SUBTOT"));
            String dbLoanAmtBalFund = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_FUND"));
            String dbLoanAmtBalBank = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_BANK"));

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
            sqlCmd2 = "INSERT INTO "+tbName5[i]+"_BAK("
                + "M_YEAR,M_MONTH,RUN_MASTER_NO,RUN_SUB_NO,"
                + "RUN_NEXT_NO,LOAN_CNT_YEAR,LOAN_AMT_YEAR,"
                + "LOAN_CNT_TOTACC,LOAN_AMT_TOTACC,LOAN_CNT_BAL,"
                + "LOAN_AMT_BAL_SUBTOT,LOAN_AMT_BAL_FUND,LOAN_AMT_BAL_BANK,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbRunMasterNo + "','" + dbRunSubNo + "',"
                + "'" + dbRunNextNo + "'," + dbLoanCntYear + ","
                + dbLoanAmtYear + "," + dbLoanCntTotacc + ","
                + dbLoanAmtTotacc + "," + dbLoanCntBal + ","
                + dbLoanAmtBalSubtot + "," + dbLoanAmtBalFund + ","
                + dbLoanAmtBalBank + ","
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
            //if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName5[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName5[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName5[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //B03_1_log=============================================================
        String tbName6[] = {"B03_1_log"};
        for(int i=0;i<tbName6.length;i++){
          tbNameNow = tbName6[i];
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
              + "A.M_YEAR,A.M_MONTH,A.FUNS_MASTER_NO,A.FUNS_SUB_NO,"
              + "A.FUNS_NEXT_NO,A.LOAN_CNT_TOTACC,A.LOAN_AMT_TOTACC_FUND,"
              + "A.LOAN_AMT_TOTACC_BANK,A.LOAN_AMT_TOTACC_TOT,"
              + "A.LOAN_CNT_BAL,A.LOAN_AMT_BAL_FUND,"
              + "A.LOAN_AMT_BAL_BANK,A.LOAN_AMT_BAL_TOT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName6[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName6[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName6[i]+"'"
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
            String dbFunsMasterNo = RptCG002W.getTrimString(rs.getString("FUNS_MASTER_NO"));
            String dbFunsSubNo = RptCG002W.getTrimString(rs.getString("FUNS_SUB_NO"));
            String dbFunsNextNo = RptCG002W.getTrimString(rs.getString("FUNS_NEXT_NO"));
            String dbLoanCntTotacc = RptCG002W.getTrimString(rs.getString("LOAN_CNT_TOTACC"));
            String dbLoanAmtTotaccFund = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC_FUND"));
            String dbLoanAmtTotaccBank = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC_BANK"));
            String dbLoanAmtTotaccTot = RptCG002W.getTrimString(rs.getString("LOAN_AMT_TOTACC_TOT"));
            String dbLoanCntBal = RptCG002W.getTrimString(rs.getString("LOAN_CNT_BAL"));
            String dbLoanAmtBalFund = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_FUND"));
            String dbLoanAmtBalBank = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_BANK"));
            String dbLoanAmtBalTot = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL_TOT"));

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
            sqlCmd2 = "INSERT INTO "+tbName6[i]+"_BAK("
                + "M_YEAR,M_MONTH,FUNS_MASTER_NO,FUNS_SUB_NO,"
                + "FUNS_NEXT_NO,LOAN_CNT_TOTACC,LOAN_AMT_TOTACC_FUND,"
                + "LOAN_AMT_TOTACC_BANK,LOAN_AMT_TOTACC_TOT,"
                + "LOAN_CNT_BAL,LOAN_AMT_BAL_FUND,"
                + "LOAN_AMT_BAL_BANK,LOAN_AMT_BAL_TOT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbFunsMasterNo + "','" + dbFunsSubNo + "',"
                + "'" + dbFunsNextNo + "'," + dbLoanCntTotacc + ","
                + dbLoanAmtTotaccFund + "," + dbLoanAmtTotaccBank + ","
                + dbLoanAmtTotaccTot + "," + dbLoanCntBal + ","
                + dbLoanAmtBalFund + "," + dbLoanAmtBalBank + ","
                + dbLoanAmtBalTot + ","
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
                    + "VALUES('" + tbName6[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName6[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName6[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //B03_2_log=============================================================
        String tbName7[] = {"B03_2_log"};
        for(int i=0;i<tbName7.length;i++){
          tbNameNow = tbName7[i];
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
              + "A.M_YEAR,A.M_MONTH,A.FUNS_MASTER_NO,A.FUNS_SUB_NO,"
              + "A.FUNS_NEXT_NO,A.LOAN_AMT_BAL,"
              + "A.LOAN_AMT_OVER,A.LOAN_RATE_OVER,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName7[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName7[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName7[i]+"'"
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
            String dbFunsMasterNo = RptCG002W.getTrimString(rs.getString("FUNS_MASTER_NO"));
            String dbFunsSubNo = RptCG002W.getTrimString(rs.getString("FUNS_SUB_NO"));
            String dbFunsNextNo = RptCG002W.getTrimString(rs.getString("FUNS_NEXT_NO"));
            String dbLoanAmtBal = RptCG002W.getTrimString(rs.getString("LOAN_AMT_BAL"));
            String dbLoanAmtOver = RptCG002W.getTrimString(rs.getString("LOAN_AMT_OVER"));
            String dbLoanRateOver = RptCG002W.getTrimString(rs.getString("LOAN_RATE_OVER"));

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
            sqlCmd2 = "INSERT INTO "+tbName7[i]+"_BAK("
                + "M_YEAR,M_MONTH,FUNS_MASTER_NO,FUNS_SUB_NO,"
                + "FUNS_NEXT_NO,LOAN_AMT_BAL,"
                + "LOAN_AMT_OVER,LOAN_RATE_OVER,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbFunsMasterNo + "','" + dbFunsSubNo + "',"
                + "'" + dbFunsNextNo + "'," + dbLoanAmtBal + ","
                + dbLoanAmtOver + "," + dbLoanRateOver + ","
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
                    + "VALUES('" + tbName7[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName7[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName7[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }


        //B03_3_log=============================================================
        String tbName8[] = {"B03_3_log"};
        for(int i=0;i<tbName8.length;i++){
          tbNameNow = tbName8[i];
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
              + "A.M_YEAR,A.M_MONTH,A.FUNO_MASTER_NO,A.FUNO_SUB_NO,"
              + "A.FUNO_NEXT_NO,A.FUNO_AMT,A.FUNO_RATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName8[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName8[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName8[i]+"'"
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
            String dbFunoMasterNo = RptCG002W.getTrimString(rs.getString("FUNO_MASTER_NO"));
            String dbFunoSubNo = RptCG002W.getTrimString(rs.getString("FUNO_SUB_NO"));
            String dbFunoNextNo = RptCG002W.getTrimString(rs.getString("FUNO_NEXT_NO"));
            String dbFunoAmt = RptCG002W.getTrimString(rs.getString("FUNO_AMT"));
            String dbFunoRate = RptCG002W.getTrimString(rs.getString("FUNO_RATE"));

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
            sqlCmd2 = "INSERT INTO "+tbName8[i]+"_BAK("
                + "M_YEAR,M_MONTH,FUNO_MASTER_NO,FUNO_SUB_NO,"
                + "FUNO_NEXT_NO,FUNO_AMT,FUNO_RATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbFunoMasterNo + "','" + dbFunoSubNo + "',"
                + "'" + dbFunoNextNo + "'," + dbFunoAmt + ","
                + dbFunoRate + ","
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
                    + "VALUES('" + tbName8[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName8[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName8[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }


        //B03_4_log=============================================================
        String tbName9[] = {"B03_4_log"};
        for(int i=0;i<tbName9.length;i++){
          tbNameNow = tbName9[i];
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
              + "A.M_YEAR,A.M_MONTH,A.BANK_NO,A.MACHINE_CNT,A.MACHINE_AMT,"
              + "A.LAND_CNT,A.LAND_AMT,A.HOUSE_CNT,A.HOUSE_AMT,"
              + "A.BUILD_CNT,A.BUILD_AMT,A.TOT_CNT,A.TOT_AMT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName9[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName9[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName9[i]+"'"
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
            String dbMachineCnt = RptCG002W.getTrimString(rs.getString("MACHINE_CNT"));
            String dbMachineAmt = RptCG002W.getTrimString(rs.getString("MACHINE_AMT"));
            String dbLandCnt = RptCG002W.getTrimString(rs.getString("LAND_CNT"));
            String dbLandAmt = RptCG002W.getTrimString(rs.getString("LAND_AMT"));
            String dbHouseCnt = RptCG002W.getTrimString(rs.getString("HOUSE_CNT"));
            String dbHouseAmt = RptCG002W.getTrimString(rs.getString("HOUSE_AMT"));
            String dbBuildCnt = RptCG002W.getTrimString(rs.getString("BUILD_CNT"));
            String dbBuildAmt = RptCG002W.getTrimString(rs.getString("BUILD_AMT"));
            String dbTotCnt = RptCG002W.getTrimString(rs.getString("TOT_CNT"));
            String dbTotAmt = RptCG002W.getTrimString(rs.getString("TOT_AMT"));

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
            sqlCmd2 = "INSERT INTO "+tbName9[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_NO,MACHINE_CNT,MACHINE_AMT,"
                + "LAND_CNT,LAND_AMT,HOUSE_CNT,HOUSE_AMT,"
                + "BUILD_CNT,BUILD_AMT,TOT_CNT,TOT_AMT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankNo + "'," + dbMachineCnt + ","
                + dbMachineAmt + "," + dbLandCnt + ","
                + dbLandAmt + "," + dbHouseCnt + ","
                + dbHouseAmt + "," + dbBuildCnt + ","
                + dbBuildAmt + "," + dbTotCnt + ","
                + dbTotAmt + ","
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
                    + "VALUES('" + tbName9[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName9[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName9[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //BN04_log=============================================================
        String tbName10[] = {"BN04_log"};
        for(int i=0;i<tbName10.length;i++){
          tbNameNow = tbName10[i];
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
              + "A.TBANK_NO,A.BANK_NO,A.BANK_NAME,A.BN_TYPE,A.BANK_TYPE,"
              + "A.ADD_USER,A.ADD_NAME,A.ADD_DATE,A.BANK_B_NAME,A.KIND_1,"
              + "A.KIND_2,A.BN_TYPE2,A.EXCHANGE_NO,A.USER_ID,A.USER_NAME,"
              + "A.UPDATE_DATE,A.UPDATE_KIND_C,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName10[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName10[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName10[i]+"'"
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
            String dbTbankNo =RptCG002W.getTrimString(rs.getString("TBANK_NO"));
            String dbBankNo =RptCG002W.getTrimString(rs.getString("BANK_NO"));
            String dbBankName =RptCG002W.getTrimString(rs.getString("BANK_NAME"));
            String dbBnType =RptCG002W.getTrimString(rs.getString("BN_TYPE"));
            String dbBankType =RptCG002W.getTrimString(rs.getString("BANK_TYPE"));
            String dbAddUser =RptCG002W.getTrimString(rs.getString("ADD_USER"));
            String dbAddName =RptCG002W.getTrimString(rs.getString("ADD_NAME"));
            Timestamp dbAddDateTimestamp = rs.getTimestamp("ADD_DATE");
            String dbAddDate = "";
            if(dbAddDateTimestamp != null){
              dbAddDate = dbAddDateTimestamp.toString();
              dbAddDate = dbAddDate.substring(0, dbAddDate.indexOf("."));
            }
            String dbBankBName =RptCG002W.getTrimString(rs.getString("BANK_B_NAME"));
            String dbKind1 =RptCG002W.getTrimString(rs.getString("KIND_1"));
            String dbKind2 =RptCG002W.getTrimString(rs.getString("KIND_2"));
            String dbBnType2 =RptCG002W.getTrimString(rs.getString("BN_TYPE2"));
            String dbExchangNo =RptCG002W.getTrimString(rs.getString("EXCHANGE_NO"));
            String dbUserId =RptCG002W.getTrimString(rs.getString("USER_ID"));
            String dbUserName =RptCG002W.getTrimString(rs.getString("USER_NAME"));
            Timestamp dbUpdateDateTimestamp = rs.getTimestamp("UPDATE_DATE");
            String dbUpdateDate = "";
            if(dbUpdateDateTimestamp != null){
              dbUpdateDate = dbUpdateDateTimestamp.toString();
              dbUpdateDate = dbUpdateDate.substring(0, dbUpdateDate.indexOf("."));
            }
            String dbUpdateKindC =RptCG002W.getTrimString(rs.getString("UPDATE_KIND_C"));

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
            sqlCmd2 = "INSERT INTO "+tbName10[i]+"_BAK("
                + "TBANK_NO,BANK_NO,BANK_NAME,BN_TYPE,BANK_TYPE,"
                + "ADD_USER,ADD_NAME,ADD_DATE,BANK_B_NAME,KIND_1,"
                + "KIND_2,BN_TYPE2,EXCHANGE_NO,USER_ID,USER_NAME,"
                + "UPDATE_DATE,UPDATE_KIND_C,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbTbankNo + "','" + dbBankNo + "',"
                + "'" + dbBankName + "','" + dbBnType + "',"
                + "'" + dbBankType + "','" + dbAddUser + "',"
                + "'" + dbAddName + "',"
                + "TO_DATE('" + dbAddDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbBankBName + "','" + dbKind1 + "',"
                + "'" + dbKind2 + "','" + dbBnType2 + "',"
                + "'" + dbExchangNo + "','" + dbUserId + "',"
                + "'" + dbUserName + "',"
                + "TO_DATE('" + dbUpdateDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateKindC + "',"
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
                    + "VALUES('" + tbName10[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName10[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName10[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //BANK_CMML_log=============================================================
        String tbName11[] = {"BANK_CMML_log"};
        for(int i=0;i<tbName11.length;i++){
          tbNameNow = tbName11[i];
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
              + "A.BANK_NO,A.BANK_NAME,A.BANK_ENGLISH,A.BUSINESS_PERSON,"
              + "A.HSIEN_ID,A.AREA_ID,A.ADDR,A.WEB_SITE,A.M_POSITION,"
              + "A.M_NAME,A.M_TELNO,A.M_CELLINO,A.M_FAX,A.M_EMAIL,A.M_SEX,"
              + "A.M_POSITION_OFFICER,A.M_NAME_OFFICER,A.M_TELNO_OFFICER,"
              + "A.M_CELLINO_OFFICER,A.M_FAX_OFFICER,A.M_EMAIL_OFFICER,"
              + "A.M_SEX_OFFICER,A.USER_ID,A.USER_NAME,A.UPDATE_DATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName11[i]+" A ,WTT01 W "
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
              sqlCmd2 = "DELETE "+tbName11[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName11[i]+"'"
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
            String dbBankName = RptCG002W.getTrimString(rs.getString("BANK_NAME"));
            String dbBankEnglish = RptCG002W.getTrimString(rs.getString("BANK_ENGLISH"));
            String dbBusinessPerson = RptCG002W.getTrimString(rs.getString("BUSINESS_PERSON"),"0");
            String dbHsineId = RptCG002W.getTrimString(rs.getString("HSIEN_ID"));
            String dbAreaId = RptCG002W.getTrimString(rs.getString("AREA_ID"));
            String dbAddr = RptCG002W.getTrimString(rs.getString("ADDR"));
            String dbWebSite = RptCG002W.getTrimString(rs.getString("WEB_SITE"));
            String dbMPosition = RptCG002W.getTrimString(rs.getString("M_POSITION"));
            String dbMName = RptCG002W.getTrimString(rs.getString("M_NAME"));
            String dbMTelno = RptCG002W.getTrimString(rs.getString("M_TELNO"));
            String dbMCellino = RptCG002W.getTrimString(rs.getString("M_CELLINO"));
            String dbMFax = RptCG002W.getTrimString(rs.getString("M_FAX"));
            String dbMEmail = RptCG002W.getTrimString(rs.getString("M_EMAIL"));
            String dbMSex = RptCG002W.getTrimString(rs.getString("M_SEX"));
            String dbMPositionOfficer = RptCG002W.getTrimString(rs.getString("M_POSITION_OFFICER"));
            String dbMNameOfficer = RptCG002W.getTrimString(rs.getString("M_NAME_OFFICER"));
            String dbMTelnoOfficer = RptCG002W.getTrimString(rs.getString("M_TELNO_OFFICER"));
            String dbMCellinoOfficer = RptCG002W.getTrimString(rs.getString("M_CELLINO_OFFICER"));
            String dbMFaxOfficer = RptCG002W.getTrimString(rs.getString("M_FAX_OFFICER"));
            String dbMEmailOfficer = RptCG002W.getTrimString(rs.getString("M_EMAIL_OFFICER"));
            String dbMSexOfficer = RptCG002W.getTrimString(rs.getString("M_SEX_OFFICER"));
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
            sqlCmd2 = "INSERT INTO "+tbName11[i]+"_BAK("
                + "BANK_NO,BANK_NAME,BANK_ENGLISH,BUSINESS_PERSON,HSIEN_ID,"
                + "AREA_ID,ADDR,WEB_SITE,M_POSITION,M_NAME,M_TELNO,M_CELLINO,"
                + "M_FAX,M_EMAIL,M_SEX,M_POSITION_OFFICER,M_NAME_OFFICER,"
                + "M_TELNO_OFFICER,M_CELLINO_OFFICER,M_FAX_OFFICER,"
                + "M_EMAIL_OFFICER,M_SEX_OFFICER,USER_ID,USER_NAME,UPDATE_DATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbBankNo + "','" + dbBankName + "',"
                + "'" + dbBankEnglish + "'," + dbBusinessPerson + ","
                + "'" + dbHsineId + "','" + dbAreaId + "',"
                + "'" + dbAddr + "','" + dbWebSite + "',"
                + "'" + dbMPosition + "','" + dbMName + "',"
                + "'" + dbMTelno + "','" + dbMCellino + "',"
                + "'" + dbMFax + "','" + dbMEmail + "',"
                + "'" + dbMSex + "','" + dbMPositionOfficer + "',"
                + "'" + dbMNameOfficer + "','" + dbMTelnoOfficer + "',"
                + "'" + dbMCellinoOfficer + "','" + dbMFaxOfficer + "',"
                + "'" + dbMEmailOfficer + "','" + dbMSexOfficer + "',"
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
                    + "VALUES('" + tbName11[i] + "'," + mYear + "," + mMonth
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
                + "VALUES('" + tbName11[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName11[i] + " "
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
        System.out.println("//RptCG002W_B() Have Error.....");
        e.printStackTrace();
        System.out.println(e.toString());
        System.out.println("//-------------------------------------");
      }
      return tbNameNow;
    }

}
