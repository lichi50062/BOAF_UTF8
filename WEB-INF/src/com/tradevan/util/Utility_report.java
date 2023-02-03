package com.tradevan.util;

import java.util.*;
import java.math.BigDecimal;

/**
 * <p>Title: BOAF</p>
 * <p>Description: </p>
 * <p>Copyright: Copyright (c) 2006</p>
 * <p>Company: ABYSS</p>
 * @author Brenda
 * @version 1.0
 */
public class Utility_report {
  public Utility_report() {}

  /**
   * 取得所有總機構資料
   * @return List
   */
  public static List getTatolBankType() {
    //查詢條件
    String sqlCmd = " Select HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE  from  BN01, WLX01  WHERE BN01.BANK_NO = WLX01.BANK_NO(+) AND bank_type in ( select cmuse_id from cdshareno where cmuse_div = '020' ) order by BANK_NO ";
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null, "");
    return dbData;
  }

  /**
   * 取得所有縣市
   * @return List
   */
  public static List getCity() {
    //查詢條件
    String sqlCmd =" SELECT HSIEN_id, HSIEN_name from cd01 order by input_order, hsien_id ";
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null, "");
    return dbData;
  }

  /**
   * 四捨五入至小數點第scale位
   * @param v double
   * @param scale int 小數點第幾位
   * @return double
   */
  public static double round(double v, int scale) {
    //System.out.println("Mathround round here");
    if (scale < 0) {
      throw new IllegalArgumentException("The scale must be a positive integer or zero");
    }
    BigDecimal b = new BigDecimal(Double.toString(v));
    BigDecimal one = new BigDecimal("1");

    return b.divide(one, scale, BigDecimal.ROUND_HALF_UP).doubleValue();
  }

  /**
   * 四捨五入至小數點第scale位
   * @param v double
   * @param scale int 小數點第幾位
   * @return double
   */
  public static String round(String v, int scale) {
    //System.out.println("Mathround round here");
    if (scale < 0) {
      throw new IllegalArgumentException("The scale must be a positive integer or zero");
    }
    BigDecimal b = new BigDecimal(v);
    BigDecimal one = new BigDecimal("1");
    String ans = b.divide(one, scale, BigDecimal.ROUND_HALF_UP).toString();

    if(scale > 0){
      String tmpAns = "";
      if (ans.indexOf(".") >= 0) {
        tmpAns = ans.substring(ans.indexOf(".") + 1);
      }else {
        ans += ".";
      }
      for (int i = tmpAns.length(); i < scale; i++) {
        ans += "0";
      }
    }

    return ans;
  }


  /**
   * 四捨五入至小數點第scale位
   * @param a 被除數
   * @param b 除數
   * @param scale 小數點第幾位
   */
  public static String round(String a, String b, int scale) {
    //System.out.println("Mathround round here");
    if (scale < 0) {
      throw new IllegalArgumentException(
          "The scale must be a positive integer or zero");
    }
    String ans = "0";
    if(a != null && b != null && !a.equals("") && !b.equals("") && !a.equals("0") && !b.equals("0")){
      BigDecimal aBigDecimal = new BigDecimal(a);
      BigDecimal bBigDecimal = new BigDecimal(b);
      ans = aBigDecimal.divide(bBigDecimal, scale, BigDecimal.ROUND_HALF_UP).toString();
    }

    if(scale > 0){
      String tmpAns = "";
      if (ans.indexOf(".") >= 0) {
        tmpAns = ans.substring(ans.indexOf(".") + 1);
      }else {
        ans += ".";
      }
      for (int i = tmpAns.length(); i < scale; i++) {
        ans += "0";
      }
    }

    return ans;
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
   * 將物件轉為字串格式，若為NULL 傳回tmp
   * @param str Object
   * @param tmp String
   * @return String
   */
  public static String getTrimString(Object str,String tmp) {
    return str != null ? ( (String) str).trim() : tmp;
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
   * 數值相加
   * @param a
   * @param b
   */
  public static String add(String a, String b) {
    String ans = "0";
    a = getTrimString(a,"0");
    b = getTrimString(b,"0");

    BigDecimal aBigDecimal = new BigDecimal(a);
    BigDecimal bBigDecimal = new BigDecimal(b);
    ans = aBigDecimal.add(bBigDecimal).toString();

    return ans;
  }

  /**
   * 數值相乘
   * @param a
   * @param b
   */
  public static String multiply(String a, String b) {
    String ans = "0";
    a = getTrimString(a,"0");
    b = getTrimString(b,"0");

    BigDecimal aBigDecimal = new BigDecimal(a);
    BigDecimal bBigDecimal = new BigDecimal(b);
    ans = aBigDecimal.multiply(bBigDecimal).toString();

    return ans;
  }
}
