package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.util.Date;
import java.sql.*;
import java.text.DateFormat;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;

public class Utility_WM {

	private static String dirpath = SetConfig.WM_dirpath;
	//SMTP_HOST如果跟FROM_ADDR不同會導致MAIL無法送出..WHY?
	public static String SMTP_HOST = SetConfig.SMTP_HOST;
	public static String FROM_ADDR = "pb@mail.boma.gov.tw";
	public static String FROM_ADDR = "pb@mail.boma.gov.tw";
	public static String FROM_ADDR = "pb@mail.boma.gov.tw";
	public static String USER_ID	= "pb";
	public static String PWD		= "";
	private static boolean SEND_MAIL = true;
	private static String BCC_addr	= "anson_tsai@hotmail.com";
	private static boolean SEND_BCC = false;

	/*
	 * get upload data's directory from PathConfig.dat
	 */
	public static String getDataDir() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("dataDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}
	public static String getDataDi1() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("dataDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}
	public static String getDataDir2() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("dataDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}
	/*
	 * get upload dataBK's directory from PathConfig.dat
	 */
	public static String getDataBKDir() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("dataBKDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}

	/*
	 * get unload dataFTP's directory from PathConfig.dat
	 */
	public static String getDataFTPDir() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("dataFTPDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}

	/*
	 * get unload dataFTP's directory from PathConfig.dat
	 */
	public static String getCheckOKDir() {

		String dir	= "";
		try {
			Properties p = new Properties();
			p.load(new FileInputStream(dirpath));
			dir = (String) p.get("checkOKDir");
		}
		catch (FileNotFoundException e) {}
		catch (IOException e) {}

		return dir;
	}

	/*
	 *  get DB date format;
	 *  0910212 -> 2002/02/12
	 *  910212  -> 2002/02/12
	 */
	public static String getDate(String cDate) {
		String c = cDate;
		String ndate = "";

		if (c == null)
			return "";
		else
		if ((c.trim().length() > 7) || (c.trim().length() < 6))
			return "";
		else
			c = cDate.trim();

		if (c.length() == 6)
			ndate = Integer.toString(Integer.parseInt(c.substring(0, 2)) + 1911) + "/" +
					c.substring(2, 4) + "/" + c.substring(4, 6);
		else
		if (c.length() == 7)
			ndate = Integer.toString(Integer.parseInt(c.substring(0, 3)) + 1911) + "/" +
					c.substring(3, 5) + "/" + c.substring(5, 7);

		return ndate;
	}
	/*
	 *  get DB date format;
	 *  0910212 -> 02/12/2002
	 *  910212  -> 02/12/2002
	 */
	public static String setDBdate(String cDate) {
		String c = cDate;
		String ndate = "";

		if (c == null)
			return "";
		else
		if ((c.trim().length() > 7) || (c.trim().length() < 6))
			return "";
		else
			c = cDate.trim();

		if (c.length() == 6)
			ndate = c.substring(2, 4) + "/" + c.substring(4, 6) + "/" +
					Integer.toString(Integer.parseInt(c.substring(0, 2)) + 1911);
		else
		if (c.length() == 7)
			ndate = c.substring(3, 5) + "/" + c.substring(5, 7) +  "/" +
					Integer.toString(Integer.parseInt(c.substring(0, 3)) + 1911);

		return ndate;
	}

	/*
	 *  get upload text filename
	 */
	public static String getFileName(String report_no, String bank_code, String m_year,
									String m_month) {
		String fileName	= null;

		return fileName = report_no + (bank_code.trim() + "0000000").substring(0, 7) +
						("000" + m_year).substring(m_year.length()) +
						("00" + m_month).substring(m_month.length());
	}

	/*
	 *  return 091010020000**
	 *  表示本檔案本月不用申報資料
	 */
	public static String getEmptyDataString(String bank_code, String m_year, String m_month) {
		String str = null;

		return str = ("000" + m_year).substring(m_year.length()) +
					("00" + m_month).substring(m_month.length()) +
					(bank_code.trim() + "0000000").substring(0, 7) +
					"**";
	}

	/*
	 * "00000-23444"	-> "-0000023444"
	 */
	public static String getNumber(String txtline) throws Exception {

		if (txtline.trim().equals(""))
			return "0";

		if (txtline.indexOf("-") != -1)
			return "-" + txtline.replace('-', '0');
		else
			return txtline;
	}

	/*
	 * get ruleno2's amt
	 * 無使用
	 */
	public static double getLRAmt(Connection conn, String bank_type, String cano,
									String LR, String tblname_t, String m_year, String m_month,
									String bank_code)
									throws SQLException, Exception {

		Statement	stmt	= null;
		ResultSet	rs		= null;
		String		sqlCmd	= null;
		double		amt		= 0.0;
		double		amt_tbl	= 0.0;	//table's amt
		boolean		plus1K	= false;//公式中出現*1000者,最後將金額round
		String		noop	= null;

		try {

			stmt = conn.createStatement();

			sqlCmd = "SELECT a.acc_code, amt, noop, nserial FROM ruleno2 a LEFT JOIN " + tblname_t +
					 " b ON a.acc_code = b.acc_code AND b.m_year=" + m_year + " AND b.m_month=" + m_month +
					 " AND b.bank_code='" + bank_code + "'" +
					 " WHERE a.acc_type='" + bank_type + "' AND cano='" + cano + "' AND left_flag='" +
					 LR + "' ORDER BY nserial";
			rs = stmt.executeQuery(sqlCmd);

			while (rs.next()) {
				if (rs.getString(1).substring(0, 3).equals("FIX"))
					amt_tbl = Double.parseDouble(ListArray.getCONST_VALUE(rs.getString(1)));
				else
					amt_tbl = rs.getDouble(2);

				if ((noop == null) || (noop.equals(""))) {
					amt	= amt_tbl;
				}
				else if (noop.equals("+")) {
					amt += amt_tbl;
				}
				else if (noop.equals("-")) {
					amt -= amt_tbl;
				}
				else if (noop.equals("*")) {
					amt *= amt_tbl;
				}
				else if (noop.equals("/")) {
					if (rs.getDouble(2) != 0)
						amt /= amt_tbl;
				}
				if ((noop != null) && noop.equals("*") && rs.getString(1).equals("FIX_K"))
					plus1K = true;

				noop = (rs.getString(3) == null ? "" : rs.getString(3).trim());
			} //while()
			if (plus1K)
				amt = Math.round(amt);

			return amt;
		}
		catch (SQLException e) {
			throw e;
		}
		catch (Exception e) {
			throw e;
		}
		finally {
			if (rs != null)		rs.close();
			if (stmt != null)	stmt.close();
		}
	} //getLRAmt

	/*
	 * get ruleno2's amt
	 * this function dedicated to A05
	 * 無使用
	 */
	public static double getLRAmt_A05(Connection conn, String cano, String LR,
									String tblname_t, String m_year, String m_month, String hsien_id,
									String bank_code)
									throws SQLException, Exception {

		Statement	stmt	= null;
		ResultSet	rs		= null;
		String		sqlCmd	= null;
		double		amt		= 0.0;
		double		amt_tbl	= 0.0;	//table's amt
		boolean		plus1K	= false;//公式中出現*1000者,最後將金額round
		String		noop	= null;

		try {
			if (hsien_id == null)	return 0;

			stmt = conn.createStatement();

			sqlCmd = "SELECT a.acc_code, amt, noop, nserial FROM ruleno2 a LEFT JOIN " + tblname_t +
					 " b ON a.acc_code = b.acc_code AND b.m_year=" + m_year + " AND b.m_month=" + m_month +
					 " AND b.hsien_id = '" + hsien_id + "' AND b.bank_code ='" + bank_code + "'" +
					 " WHERE a.acc_type='D' AND cano='" + cano + "' AND left_flag='" +
					 LR + "' ORDER BY nserial";
			rs = stmt.executeQuery(sqlCmd);

			while (rs.next()) {
				if (rs.getString(1).substring(0, 3).equals("FIX"))
					amt_tbl = Double.parseDouble(ListArray.getCONST_VALUE(rs.getString(1)));
				else
					amt_tbl = rs.getDouble(2);

				if ((noop == null) || (noop.equals(""))) {
					amt	= amt_tbl;
				}
				else if (noop.equals("+")) {
					amt += amt_tbl;
				}
				else if (noop.equals("-")) {
					amt -= amt_tbl;
				}
				else if (noop.equals("*")) {
					amt *= amt_tbl;
				}
				else if (noop.equals("/")) {
					if (rs.getDouble(2) != 0)
						amt /= amt_tbl;
				}
				noop = (rs.getString(3) == null ? "" : rs.getString(3).trim());
			} //while()
			if (plus1K)
				amt = Math.round(amt);

			return amt;
		}
		catch (SQLException e) {
			throw e;
		}
		catch (Exception e) {
			throw e;
		}
		finally {
			if (rs != null)		rs.close();
			if (stmt != null)	stmt.close();
		}
	} //getLRAmt

	/*
	 * get ruleno2's amt
	 * this function dedicated to F14
	 * bank_no 子公司代號
	 * 無使用
	 */
	public static double getLRAmt_F14(Connection conn, String cano, String LR,
									String tblname_t, String m_year, String m_month, String bank_no)
									throws SQLException, Exception {

		Statement	stmt	= null;
		ResultSet	rs		= null;
		String		sqlCmd	= null;
		double		amt		= 0.0;
		String		noop	= null;

		try {
			if (bank_no == null)	return 0;

			stmt = conn.createStatement();

			sqlCmd = "SELECT a.acc_code, amt, noop, nserial FROM ruleno2_fc a LEFT JOIN " + tblname_t +
					 " b ON a.acc_code = b.acc_code AND b.m_year=" + m_year + " AND b.m_month=" + m_month +
					 " AND b.bank_no = '" + bank_no + "'" +
					 " WHERE a.report_no='F14' AND cano='" + cano + "' AND left_flag='" +
					 LR + "' ORDER BY nserial";
			rs = stmt.executeQuery(sqlCmd);

			while (rs.next()) {
				if ((noop == null) || (noop.equals(""))) {
					amt	= rs.getDouble(2);
				}
				else if (noop.equals("+")) {
					amt += rs.getDouble(2);
				}
				else if (noop.equals("-")) {
					amt -= rs.getDouble(2);
				}
				else if (noop.equals("*")) {
					amt *= rs.getDouble(2);
				}
				else if (noop.equals("/")) {
					if (rs.getDouble(2) != 0)
						amt /= rs.getDouble(2);
				}
				noop = (rs.getString(3) == null ? "" : rs.getString(3).trim());
			} //while()
			return amt;
		}
		catch (SQLException e) {
			throw e;
		}
		catch (Exception e) {
			throw e;
		}
		finally {
			if (rs != null)		rs.close();
			if (stmt != null)	stmt.close();
		}
	} //getLRAmt

	public static double getLRAmt_FC(Connection conn, String report_no, String cano,
									String LR, String tblname_t, String m_year, String m_month,
									String bank_code)
									throws SQLException, Exception {

		Statement	stmt	= null;
		ResultSet	rs		= null;
		String		sqlCmd	= null;
		double		amt		= 0.0;
		double		amt_tbl	= 0.0;	//table's amt
		String		noop	= null;

		try {

			stmt = conn.createStatement();

			sqlCmd = "SELECT a.acc_code, amt, noop, nserial FROM ruleno2_fc a LEFT JOIN " + tblname_t +
					 " b ON a.acc_code = b.acc_code AND b.m_year=" + m_year + " AND b.m_month=" + m_month +
					 " AND b.bank_no ='" + bank_code + "'" +
					 " WHERE a.report_no='" + report_no + "' AND cano='" + cano + "' AND left_flag='" +
					 LR + "' ORDER BY nserial";
			rs = stmt.executeQuery(sqlCmd);

			while (rs.next()) {
				if (rs.getString(1).substring(0, 3).equals("FIX"))
					amt_tbl = Double.parseDouble(ListArray_FC.getCONST_VALUE(rs.getString(1)));
				else
					amt_tbl = rs.getDouble(2);

				if ((noop == null) || (noop.equals(""))) {
					amt	= amt_tbl;
				}
				else if (noop.equals("+")) {
					amt += amt_tbl;
				}
				else if (noop.equals("-")) {
					amt -= amt_tbl;
				}
				else if (noop.equals("*")) {
					amt *= amt_tbl;
				}
				else if (noop.equals("/")) {
					if (amt_tbl != 0)
						amt /= amt_tbl;
				}
				noop = (rs.getString(3) == null ? "" : rs.getString(3).trim());
			} //while()
			return amt;
		}
		catch (SQLException e) {
			throw e;
		}
		catch (Exception e) {
			throw e;
		}
		finally {
			if (rs != null)		rs.close();
			if (stmt != null)	stmt.close();
		}
	} //getLRAmt

	/*
	 * check whether txtline can be converted to number format
	 */
	public static boolean isNumber(String txtline) throws Exception {

		int i = -1;

		if (txtline.trim().equals(""))
			return false;

		try {
			if ((i = txtline.indexOf("-")) != -1) {
				String newstr = txtline.substring(i + 1);
				if (newstr.indexOf("-") != -1)		//show up two times
					return false;
			}
			if ((i = txtline.indexOf(".")) != -1) {
					return false;
			}
			String s = txtline.replace('-', '0');	//因為parseLong() '-'必須出現在第一位,所以先將其轉為0
			long l = Long.parseLong(s);
		}
		catch (NumberFormatException e) {
			return false;
		}
		catch (Exception e) {
			return false;
		}
		return true;
	}

	/*
	 * sendMailNotification
	 * suss_fail == true :更新成功; suss_fail == false : 更新失敗
	 * i == 0 : ZZ020W 更新; i == 1 : ZZ015W_ShowUpdate 更新
	 * input_method == "W" :線上編輯; input_method == "F" :檔案上傳
	 */
	public static void sendMailNotification(Statement stmt, String bank_no, String report_no,
											String m_year, String m_month, boolean suss_fail,
											int i, String im)
											throws Exception {
		Properties props = new Properties();
		Session session;
		Transport transport;

		try {
			if (!SEND_MAIL)	return;

			String filename = DownLoad.getFileName(report_no, bank_no, m_year, m_month);
			String input_method = im.equals("W") ? "線上編輯" : "檔案上傳";
			String input_date = getInput_date(stmt, bank_no, m_year, m_month, report_no);
			String txt	= "";
			String to	= getWZZ08EmailAddress(stmt, m_year, m_month, bank_no, im, report_no);
			if (to.equals(""))	return;

			props.put("mail.smtp.host", SMTP_HOST);
			props.put("mail.smtp.auth", "true");
			session = Session.getInstance(props, null);
//			session.setDebug(true);
//			Message msg = new MimeMessage(session);
			MimeMessage msg = new MimeMessage(session);
			msg.setFrom(new InternetAddress(FROM_ADDR));
			msg.setRecipient(Message.RecipientType.TO, new InternetAddress(to));
//			msg.setRecipient(Message.RecipientType.TO, new InternetAddress("gunson36@hotmail.com"));

			if (SEND_BCC)
				msg.setRecipient(Message.RecipientType.BCC, new InternetAddress(BCC_addr));

			if (i == 0)
				msg.setSubject("銀行局申報資料更新結果通知--" + filename, "ISO8859_1");
			else
				msg.setSubject("銀行局申報資料更新結果通知--" + filename, "ISO8859_1");
			msg.setSentDate(new Date());
	 		txt = "報表編號：" + report_no + "\n";
	 		txt += "資料基準日：" + m_year + "年" + m_month + "月\n";
	 		txt += "申報日期：" + input_date + "\n";
	 		txt += "申報方式：" + input_method + "\n";
	 		txt += "檢核結果：";
	 		if (suss_fail)
	 			txt += "成功\n";
	 		else {
	 			txt += "失敗\n請至 Https://ebank.boma.gov.tw 查詢原因及更正申報資料";
	 		}
	 		if (i == 0)
				msg.setText(txt, "ISO8859_1");
			else
				msg.setText(txt, "ISO8859_1");

			msg.setSentDate( new Date() );
			msg.saveChanges();
			transport = session.getTransport("smtp");
			transport.connect(SMTP_HOST, USER_ID, PWD);
			transport.sendMessage(msg, msg.getAllRecipients());
			transport.close();
		}
		catch (Exception e) {
			//先不丟出Exception or 會影響ZZ020W
//			throw new Exception(e + e.getMessage());
		}
	} // End of sendMailNotification()

	/*
	 * get email address from WZZ08
	 */
	private static String getWZZ08EmailAddress(Statement stmt, String m_year, String m_month,
												String bank_no, String im, String report_no)
												throws Exception {
		String	sqlCmd = "";
		String	emailaddress = "";
		String	input_user = "";
		ResultSet rs = null;

		try {
			if (im.equals("W")) {//線上編輯
				sqlCmd = "SELECT input_user FROM WML01 where m_year=" + m_year + " AND m_month=" + m_month +
						" AND bank_code='" + bank_no + "' AND report_no='" + report_no + "' AND upd_code IS NOT NULL";
				rs = stmt.executeQuery(sqlCmd);
				if (rs.next())
					input_user = rs.getString(1) == null ? "" : rs.getString(1).trim();
				if (!input_user.equals(""))	{
					sqlCmd = "SELECT maintain_mail FROM WZZ08 " +
							 "WHERE maintain_id='" + input_user + "'";
					rs = stmt.executeQuery(sqlCmd);
					if (rs.next())
						emailaddress = rs.getString(1) == null ? "" : rs.getString(1).trim();
				}
			}
			else {
				sqlCmd = "SELECT maintain_mail FROM WZZ07 a JOIN WZZ08 b ON a.maintain_id = b.maintain_id " +
						 "WHERE program_id = 'WM001W' AND bank_no ='" + bank_no + "'";
				rs = stmt.executeQuery(sqlCmd);
				if (rs.next())
					emailaddress = rs.getString(1) == null ? "" : rs.getString(1).trim();
			}
			if (rs != null)
				rs.close();
		}
		catch (Exception e) {
			//先不丟出Exception or 會影響ZZ020W
//			throw new Exception(e + e.getMessage());
		}
		return emailaddress;

	} //End of getWZZ08EmailAddress()

	private static String getInput_date(Statement stmt, String bank_no, String m_year, String m_month,
								String report_no) throws Exception {

		ResultSet rs = null;
		String input_date = "";
		String sqlCmd = "SELECT input_date FROM WML01 WHERE bank_code ='" + bank_no + "' AND report_no = '" +
						report_no + "' AND m_year=" + m_year + " AND m_month=" + m_month + "";
		try {
			rs = stmt.executeQuery(sqlCmd);
			if (rs.next())
				input_date = Utility.getCHTdate(rs.getString(1).substring(0, 10), 0) + "  " +
							rs.getString(1).substring(11, 19);
		}
		catch (Exception e) {
			throw new Exception(e + e.getMessage());
		}
		return input_date;
	} //End of getInput_date()

} //End of class