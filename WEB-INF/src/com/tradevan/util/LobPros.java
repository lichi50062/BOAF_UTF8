package com.tradevan.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * 展示JDBC存取ORACLE大型資料物件LOB幾種情況的示範
 *
 * @author 雨亦奇(zhsoft88@sohu.com)
 * @version 1.0
 * @since 2003.05.28
 *
 * @modify jing.lin
 * @version 1.1
 * @date 2003.08.07
 * for 修改為 API 型式，原始檔案在 \MySamples\OracleSample\src_lob\ 目錄內。
 */
public class LobPros {
  private LobPros() {
  }

  /**
    * 往資料庫中插入一個新的CLOB物件
    * @param rs - 含有 COLB 的 sql.ResultSet
    * @param fname - 該 CLOB 欄位的名稱
    * @param sValue - 要新增的內容
    * @throws java.lang.Exception
    */
  public static void insertClob(ResultSet rs, String fname, String sValue)
    throws SQLException  {
    try {
    	if (rs == null) {
      	throw new SQLException("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new SQLException("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
      /* 取出此 CLOB 物件 */
     	oracle.sql.CLOB clob = (oracle.sql.CLOB)rs.getClob(fname);
     	/* 向 CLOB 物件中寫入資料 */
     	BufferedWriter out = new BufferedWriter(clob.getCharacterOutputStream());
     	BufferedReader in = new BufferedReader(new StringReader(sValue));
     	int c;
     	while ((c=in.read())!=-1) {
      	out.write(c);
     	}
     	in.close();
     	out.close();
   	} catch (Exception ex) {
      throw new SQLException( ex.getMessage() );
    }
  }

  /**
    * 修改CLOB物件（是在原CLOB物件基礎上進行覆蓋式的修改）
    * 更新時不建議，建議採用 replaceClob()
  	* @param rs - 含有 COLB 的 sql.ResultSet
    * @param fname - 該 CLOB 欄位的名稱
    * @param sValue - 要修改的內容
    * @throws java.lang.Exception
    */
  public static void modifyClob(ResultSet rs, String fname, String sValue) throws SQLException {
    try {
    	if (rs == null) {
      	throw new SQLException("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new SQLException("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
      /* 獲取此CLOB物件 */
      oracle.sql.CLOB clob = (oracle.sql.CLOB)rs.getClob(fname);
      /* 進行覆蓋式修改 */
      BufferedWriter out = new BufferedWriter(clob.getCharacterOutputStream());
      BufferedReader in = new BufferedReader(new StringReader(sValue));
      int c;
      while ((c=in.read())!=-1) {
        out.write(c);
      }
      in.close();
      out.close();
    }
    catch (Exception ex) {
      throw new SQLException( ex.getMessage() );
    }
  }

  /**
    * 替換CLOB物件（將原CLOB物件清除，換成一個全新的CLOB物件）
    *
    * @param rs - 含有 COLB 的 sql.ResultSet
    * @param fname - 該 CLOB 欄位的名稱
    * @param sValue - 要修改的內容
    * @throws java.lang.Exception
    */
  public static void replaceClob(ResultSet rs, String fname, String sValue)
    throws SQLException  {
    try {
    	if (rs == null) {
      	throw new SQLException("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new SQLException("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
      /* 獲取此CLOB物件 */
     	oracle.sql.CLOB clob = (oracle.sql.CLOB)rs.getClob(fname);
     	/* 更新資料 */
     	BufferedWriter out = new BufferedWriter(clob.getCharacterOutputStream());
     	BufferedReader in = new BufferedReader(new StringReader(sValue));
     	int c;
     	while ((c=in.read())!=-1) {
        out.write(c);
     	}
     	in.close();
      out.close();
    } catch (Exception ex) {
      throw new SQLException( ex.getMessage() );
    }
  }

 	/**
   * CLOB物件讀取
   *
   * @param rs - 含有 COLB 的 sql.ResultSet
   * @param fname - 該 CLOB 欄位的名稱
   * @throws java.lang.Exception
   * @return String - CLOB 內的資料
   */
  public static String readClob(ResultSet rs, String fname) throws SQLException {
    try {
     	if (rs == null) {
      	throw new SQLException("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new SQLException("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
     	/* 獲取CLOB物件 */
     	oracle.sql.CLOB clob = (oracle.sql.CLOB)rs.getClob(fname);
     	/* 以字元形式輸出 */
     	BufferedReader in = new BufferedReader(clob.getCharacterStream());
     	StringBuffer sbRet = new StringBuffer(4096);
     	int c;
     	while ((c=in.read())!=-1) {
        sbRet.append((char)c);
     	}
     	in.close();
     	return sbRet.toString();
   	}
   	catch (Exception ex) {
      return "<FONT COLOR=RED>此篇文章系統無法解析，對不起，請改用純文字模式。</FONT>";
      // throw new SQLException( ex.getMessage() );
    }
	}

  /**
   * 向資料庫中插入一個新的BLOB物件
   *
   * @param rs - 含有 COLB 的 sql.ResultSet
   * @param fname - 該 CLOB 欄位的名稱
   * @param istream - 要上載的資料
   * @throws java.lang.Exception
   */
  public static void insertBlob(ResultSet rs, String fname, InputStream istream) throws Exception {
    try {
    	if (rs == null) {
      	throw new Exception("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new Exception("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
     	/* 取出此BLOB物件 */
     	oracle.sql.BLOB blob = (oracle.sql.BLOB)rs.getBlob(fname);
     	/* 向BLOB物件中寫入資料 */
     	BufferedOutputStream out = new BufferedOutputStream(blob.getBinaryOutputStream());
     	BufferedInputStream in = new BufferedInputStream(istream);
     	int c;
     	while ((c=in.read())!=-1) {
        out.write(c);
     	}
     	in.close();
      out.close();
  	} catch (Exception ex) {
      throw ex;
    }
  }

  /**
   * 修改BLOB物件（是在原BLOB物件基礎上進行覆蓋式的修改）
   *
   * @param rs - 含有 COLB 的 sql.ResultSet
   * @param fname - 該 CLOB 欄位的名稱
   * @param istream - 要上載的資料
   * @throws java.lang.Exception
   */
  public static void modifyBlob(ResultSet rs, String fname, InputStream istream) throws Exception {
    try {
     	if (rs == null) {
      	throw new Exception("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new Exception("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
     	/* 取出此BLOB物件 */
     	oracle.sql.BLOB blob = (oracle.sql.BLOB)rs.getBlob(fname);
     	/* 向BLOB物件中寫入資料 */
     	BufferedOutputStream out = new BufferedOutputStream(blob.getBinaryOutputStream());
     	BufferedInputStream in = new BufferedInputStream(istream);
     	int c;
     	while ((c=in.read())!=-1) {
        out.write(c);
     	}
     	in.close();
      out.close();
   	} catch (Exception ex) {
      throw ex;
    }
  }

  /**
   * 替換BLOB物件（將原BLOB物件清除，換成一個全新的BLOB物件）
   *
   * @param rs - 含有 COLB 的 sql.ResultSet
   * @param fname - 該 CLOB 欄位的名稱
   * @param istream - 要上載的資料
   * @throws java.lang.Exception
   */
  public static void replaceBlob(ResultSet rs, String fname, InputStream istream) throws Exception {
    try {
    	if (rs == null) {
      	throw new Exception("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new Exception("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
     	/* 取出此BLOB物件 */
     	oracle.sql.BLOB blob = (oracle.sql.BLOB)rs.getBlob(fname);
     	/* 向BLOB物件中寫入資料 */
     	BufferedOutputStream out = new BufferedOutputStream(blob.getBinaryOutputStream());
     	BufferedInputStream in = new BufferedInputStream(istream);
     	int c;
     	while ((c=in.read())!=-1) {
        out.write(c);
     	}
     	in.close();
     	out.close();
    } catch (Exception ex) {
      throw ex;
    }
  }

  /**
   * BLOB物件讀取
   *
   * @param rs - 含有 COLB 的 sql.ResultSet
   * @param fname - 該 CLOB 欄位的名稱
   * @param ostream - 用來儲存 CLOB 欄位內容
   * @throws java.lang.Exception
   */
  public static void readBlob(ResultSet rs, String fname, OutputStream ostream) throws Exception {
   	try {
   		if (rs == null) {
      	throw new Exception("傳入的 ResultSet == null，請檢查程式！");
      }
    	if (fname == null || fname.length() == 0) {
        throw new Exception("傳入的 ResultSet fieldName == null or fieldName.length() == 0，請檢查程式！");
      }
     	/* 取出此BLOB物件 */
      oracle.sql.BLOB blob = (oracle.sql.BLOB)rs.getBlob(fname);
      /* 以二進位形式輸出 */
      BufferedOutputStream out = new BufferedOutputStream(ostream);
      BufferedInputStream in = new BufferedInputStream(blob.getBinaryStream());
      int c;
      while ((c=in.read())!=-1) {
         out.write(c);
      }
      in.close();
      out.close();
    } catch (Exception ex) {
      throw ex;
    }
  }
}

