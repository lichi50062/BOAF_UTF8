package com.tradevan.util;

/* 加入Method getPercentNumber by Winnin 2004.11.22 PM*/
/* fix send mail ..自己傳給自己 by 2295 */
/* add getDatetoString 把"/"去掉 by 2295 2005.01.17*/
/* 94.01.18 add 讀取file data 儲存在List by 2295 */
/* 94.02.02 add 取得該月份的最後一天 by 2295*/
/* 94.02.02 ad 取得N月天後的日期 by 2295 */
/* 94.03.15 add sendMailNotification該申報資料已被鎖定,無法更新資料 by 2295*/
/* 94.03.15 add 取得已被鎖定的bank_code by 2295 */
/* 94.03.28 fix suss_fail boolean ->String by 2295*/
/* 94.04.14 fix 傳送email加上認証 by 2295*/
/* 94.04.15 fix send email 申報資料內容值皆為零 by 2295 */
/* 94.04.18 add toBig5Convert轉iso8859 by 2295 */
/* 94.04.25 add PathConfig.dat增加Auth=true/false */
/*              Auth=true需要ID/PWD做認証 */
/* 94.09.13 add 傳送e-mail失敗時,不throw exception */
/* 95.03.20 add addMailNotification by 2295 */
/* 95.04.04 fix 更改共用中心的e-mail主旨 by 2295 */
/* 95.04.17 fix send mail 檢核有誤時,只出現一次查詢原因 by 2295 */
/* 95.06.06 add 取得申報年/月 by 2295 */
/* 95.08.07 add 將dataList組合成a,b,c,d by  2295 */
/* 95.08.09 add 取得該用戶檢查追蹤權限的總機構資料 by 2295 */
/* 95.08.09 add 取得該用戶檢查追蹤權限的受檢單位資料 by 2295 */
/* 95.08.17 add getStringTokenizerData(String srcData,String delim)
/*              將來源字串,根據所設定的分隔符號,轉成List by 2295 */
/* 95.08.17 add 取得單位名稱 getUnitName(String unit) by 2295 */
/* 95.08.18 add getReportData by 2295 */ 
/* 95.09.05 add 將金額(amt)除以單位(Unit)四捨五入 by 2295 */
/* 95.11.03 add 檢核權限 by 2295 */
/* 95.11.13 add 取得可選機構代號權限設定(農.漁會機構) by 2295 */
/* 95.11.27 add 取得可選機構代號權限設定增加中央存保.可看到農.漁會的 by 2295 */
/* 96.01.12 fix 修改取得該用戶檢查追蹤權限的總機構資料.受檢單位資料,當縣市別不為空白時 */
/*              ,才加入屬於該縣市別(避免.全國農業金庫.農漁會共用中心.沒有縣市別)  by 2295 */ 
/* 96.03.21 add 加/解密模組 by 2295  */
/* 96.04.18 add 將List組合成a,b,c,d by  2295 */
/* 96.07.31 add 批次寫入WML03_LOG/WML02_LOG(區分檔案上傳/線上編輯)/WML01_LOG by 2295*/
/* 96.08.13 add 從WML01_UPLOAD取得對應的帳號.姓名 */
/* 96.08.15 add 批次寫入InsertWML03/updateAXX/InsertZeroAXX_List/InsertAXX_LOG by 2295 */
/* 96.11.21 fix InsertWML03_LOG區分檔案上傳/線上編輯 by 2295 */
/* 96.11.23 add 批次寫入updateAXX,增加寫入AMT_NAME by 2295 */
/* 96.12.18 add 檢核完成後.刪除暫存資料(delete AXX_TMP) by 2295 */
/* 98.08.11 add 共用sendMail(mail_to,mailsubject,sendMsg) by 2295 */
/* 99.02.10 add 取得BN01該總機構代號資料 by 2295 */
/* 99.02.10 add 取得是否有該程式細部功能權限 by 2295 */
/* 99.03.15 add 取得民國年/月getCHTYYMMDD(String yymmdd) by 2295 */
/* 99.03.10 add 取得所有總機構資料 by 2295 */
/* 99.03.10 add 取得所有縣市(99年跟100年縣市別) by 2295*/
/* 99.03.30 add getBankList(彈性報表用)區分新舊機構名稱 by 2295 */
/* 99.08.25 fix 取得所有縣市(99年跟100年縣市別),以報表的fr001w_output_order 排序為準*/    
/* 99.09.01 add getProperties_conf,取得cdao.conf設定值 by 2295 */
/* 99.09.02 add getPercentNumber_2 將傳入的數字轉換為百分比的格式 001155 --> 11.55 by 2295 */
/* 99.10.05 add getAcc_Code取得科目代號/名稱 by 2295 */
/* 99.10.15 fix getBN01區分100年/99年總機基本資料 by 2295 */
/* 99.10.20 add getPgName取得程式名稱 by 2295 */
/* 99.11.09 fix 修改機構名稱排列順序getBankList by 2295 */
/* 99.11.11 add InsertZeroAXX_List A08 zero data by 2295 */
/* 99.11.18 add 寫入WML03的參數create_dataList,回傳dataList by 2295 */
/* 99.11.19 add getWML03_count讀取WML03.count(*)--countdata by 2295 */
/* 99.11.19 add getCountZero讀取該申報資料all data 都為0的資料筆數 by 2295 */
/* 99.11.19 add getWML01讀取WML01 all data by 2295 */
/* 99.11.19 add Insert_UpdateWML01wml01不存在Insert,存在時Update by 2295 */
/* 99.11.19 add deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295 */
/* 101.07.19 add M106/M201/M206 by 2295 */
/* 101.08.20 add 寫入操作log by 2968 */
/* 101.09.21 parseLenBase64encode字串轉base64,以固定長度分割,回傳List by 2295 */
/* 102.01.21 add a02_log.amt_name by 2295 */ 
/* 102.02.01 fix InsertZeroAXX_List add a02.amt_name by 2295 */
/* 102.04.18 add InsertZeroAXX_List 漁會A01-103/01以後的資料.套用新的科目代號 by 2295 */
/* 102.04.25 add updateAXX增加A02.amt_name by 2295*/
/* 102.10.02 add InsertWML03_LOG/InsertWML02_LOG/InsertWML01_LOG/InsertWML03套用preparestatment by 2295 */
/* 102.10.09 add createEncZipFile壓縮成zip檔,並加密碼 by 2295 */
/* 102.11.08 add 調整QueryDB.改成QueryDB_SQLParam(使用preparestatment) by 2295*/
/* 103.02.11 add InsertZeroAXX_List 漁會A06-103/01以後的資料.套用新的科目代號 by 2295 */
/* 103.02.11 add InsertWML03 漁會A06-103/01以後的資料.套用新的科目代號 by 2295 */
/* 103.05.23 add 顯示檢核有誤的農漁會信部名稱 by 2295 */
/* 103.05.27 add createZipFile壓縮成zip檔案 by 2295 */
/* 104.03.09 add ISOtoUTF8轉編號 by 2295 */
/* 104.05.11 fix InsertZeroAXX_List.調整bank_code來自於axx_tmp by 2295*/
/* 104.06.03 fix parseLenBase64encode讀取編碼增加UTF-8 by 2295 */
/* 104.10.08 fix InsertAXX_LOG,調整A10_log增加新增欄位 by 2295 */
/* 105.09.12 add 共用 getLoanBank()取得貸款經辦機構名稱  by 2968*/
/* 109.06.10 add updateAXX增加A15 by 2295 */
/* 109.06.11 add InsertZeroAXX_List增加 A15 by 2295 */
/* 110.04.08 add 撤銷項目共用方法  by 6493*/
/* 110.04.09 fix InsertAXX_LOG/InsertZeroAXX_List,A02增加amt_name1/amt_name2 by 2295 */
/* 110.08.18 fix InsertZeroAXX_List/updateAXX/InsertAXX_LOG,A99增加amt_name by 2295[111.05.09取消] */
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.text.DecimalFormat; // Add by Winnin 2004.11.22
import java.util.*;
import javax.mail.*;
import javax.mail.internet.*;
import javax.servlet.ServletRequest;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
import com.sun.mail.smtp.*;
import com.tradevan.base64.Base64;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.TvEncrypt;
import java.math.BigInteger;
import java.net.URL;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;

public class Utility {
	//private static String dirpath = "D:\\workProject\\BOAF\\WEB-INF\\classes\\com\\tradevan\\util\\PathConfig.dat";
	private static String dirpath = "";
	private static String sendMsg = "";
	private static Properties p;
	private static Properties p_conf; 
	private static Date nowlog = new Date();			
	private static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	//private static String dirpath = "C:\\Sun\\WebServer6.1\\BOAF\\WEB-INF\\classes\\com\\tradevan\\util\\PathConfig.dat";
	 static{
         p = new Properties();
         p_conf = new Properties();
        try{
            ClassLoader cl = Thread.currentThread().getContextClassLoader();
            URL url = null;
            URL url_conf = null;
            if (cl != null) {
                //Environment.ini 該檔案放至於WEB-INF/classes/
                String filename = "PathConfig.dat";              
                url = cl.getResource(filename);
                url_conf = cl.getResource("conf"+System.getProperty("file.separator")+"cdao.conf");
            }            
            p.load(url.openStream());
            p_conf.load(url_conf.openStream());
              
        }catch(Exception e){
            e.printStackTrace();
        }
    }
	
	
	
	/******************************************************************
	 * 建立 subdir
	 */	
	public static boolean mkdirs(String dir){
		boolean mkOK=false;
		System.out.println("dir="+dir);
		
		try{
			File mdir = new File(dir);
			if(!mdir.exists()){
				mkOK=mdir.mkdirs();
			}	
			
			mkOK = true;
		}catch(Exception e){
			System.out.println("mkdirs.error:"+e.getMessage());			
		}
		return mkOK;
	}
	
	/******************************************************************8
	 * 檢查檔案是否存在該目錄下
	 */
	public static boolean CheckFileExist(String UploadDir,String UploadFileName){//檢查檔案是否存在該目錄下===
		String tmpFileName = "";
		if(UploadFileName.lastIndexOf("\\") != -1){
		   tmpFileName = UploadFileName.substring(UploadFileName.lastIndexOf("\\")+1,UploadFileName.length());
		}
		System.out.println("tmpFileName="+tmpFileName);
		File tmpFile = new File(UploadDir+System.getProperty("file.separator")+tmpFileName);
		System.out.println(UploadDir+System.getProperty("file.separator")+tmpFileName);
		if(tmpFile.exists()){
			return true;
		}
		
    	return false;		
	}
	
	/******************************************************************8
	 * 將"\"去掉,取實際檔名
	 */
	public static String parseFileName(String szFileName){
		String tmpFileName = "";
		if(szFileName.lastIndexOf("\\") != -1){
		   tmpFileName = szFileName.substring(szFileName.lastIndexOf("\\")+1,szFileName.length());
		}
		System.out.println("tmpFileName="+tmpFileName);
    	return tmpFileName;    	
	}
	
	/******************************************************************8
	 * copy file
	 */
	public static String CopyFile(String scSrcFile,String scDescFile)
	{
		String scErrMsg;
		try
		{	
			int iGetByte=0;
			FileInputStream SrcFileReader=new FileInputStream(scSrcFile);
			FileOutputStream DescFileWriter=new FileOutputStream(scDescFile);
			byte[] szData=new byte[8192];
			while ((iGetByte=SrcFileReader.read(szData,0,8192))>0)
			{
				DescFileWriter.write(szData,0,iGetByte);
				DescFileWriter.flush();
			}
			SrcFileReader.close();
			DescFileWriter.close();
			return "0";
		}
		catch (Exception e)
		{
			scErrMsg=e.toString();
			return "-1:"+scErrMsg;
		}
			
	}
	//get data format
	public static String getDateFormat(String dateFormat){
		try{						
			Calendar nowcalendar = Calendar.getInstance();
			SimpleDateFormat bkformat = new SimpleDateFormat(dateFormat,new Locale("en","US"));			 		
	   	    Date now = nowcalendar.getTime();
	   	    return(bkformat.format(now));
		}catch(Exception e){
			System.out.println(e.getMessage());
			return null;
		}
	}
	
	/******************************************************************8
	 * get property from PathConfig.dat
	 */
	public static String getProperties(String key) throws Exception{

		String value	= "";
		try {
			//Properties p = new Properties();
			//p.load(new FileInputStream(dirpath));
			System.out.println("@@@@@");
			value = (String) p.get(key);
		}catch (Exception e) {
			throw new Exception("Utility:Exception["+e.getMessage()+"]");
		}
		
		return value;
	}
	/******************************************************************8
	 * get conf property from cdao.conf
	 */
	public static String getProperties_conf(String key) throws Exception{

		String value	= "";
		try {
			value = (String) p_conf.get(key);
		}catch (Exception e) {
			throw new Exception("Utility_getProperties_conf:Exception["+e.getMessage()+"]");
		}
		
		return value;
	}
	/*
	 * get ruleno2's amt
	 */
	public static double getLRAmt(String bank_type,
			                      String cano,
								  String LR, 
								  String Report_no, 
								  String m_year, 
								  String m_month,
								  String bank_code)
	throws Exception {

		//Statement	stmt	= null;
		//ResultSet	rs		= null;
		String		sqlCmd	= null;
		double		amt		= 0.0;
		double		amt_tbl	= 0.0;	//table's amt
		boolean		plus1K	= false;//公式中出現*1000者,最後將金額round
		String		noop	= null;
		List paramList = new ArrayList();
		try {
			sqlCmd = "SELECT a.acc_code, amt, noop, nserial FROM ruleno2 a LEFT JOIN " + Report_no +
					 " b ON a.acc_code = b.acc_code AND b.m_year=? AND b.m_month= ?" +
					 " AND b.bank_code=?" +
					 " WHERE a.acc_type=? AND cano=? AND left_flag=? ORDER BY nserial";
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(bank_code);
			paramList.add(bank_type);
			paramList.add(cano);
			paramList.add(LR);
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"amt,nserial");
			
			for(int i=0;i<dbData.size();i++){
				if (((String)((DataObject)dbData.get(i)).getValue("acc_code")).substring(0, 3).equals("FIX")){
					amt_tbl = Double.parseDouble(ListArray.getCONST_VALUE((String)((DataObject)dbData.get(i)).getValue("acc_code")));
				}else{
					amt_tbl = Double.parseDouble((((DataObject)dbData.get(i)).getValue("amt")).toString());
				}		
				if ((noop == null) || (noop.equals(""))) {
					amt	= amt_tbl;
				}else if (noop.equals("+")) {
					amt += amt_tbl;
				}else if (noop.equals("-")) {
					amt -= amt_tbl;
				}else if (noop.equals("*")) {
					amt *= amt_tbl;
				}else if (noop.equals("/")) {
					if (Double.parseDouble((((DataObject)dbData.get(i)).getValue("amt")).toString()) != 0)
						amt /= amt_tbl;
				}
				
				if ((noop != null) && noop.equals("*") && ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("FIX_K"))
					plus1K = true;

				noop = (((DataObject)dbData.get(i)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(i)).getValue("noop")).trim());
			} //while()
			if (plus1K)
				amt = Math.round(amt);

			return amt;
		}catch (Exception e) {
			System.out.println("Utility.getLRAmt Error:"+e.getMessage());
			throw e;
		}		
	} //getLRAmt
	
	//93.12.03 add by 2295
	public static String CountDate(String InputDays){
		try{		
			int Records=0;
			String szSmallDate="";
			Date now = new Date();
			SimpleDateFormat t = new SimpleDateFormat("yyyy/MM/dd");
			Calendar nowcalendar = Calendar.getInstance();  
			now = nowcalendar.getTime(); 
			   
			long days = (24 * 60 * 60 ) * Integer.parseInt(InputDays);
			BigInteger daysBI = BigInteger.valueOf(days);
			BigInteger msecondBI = BigInteger.valueOf(1000);
			long beforedays =(daysBI.multiply(msecondBI)).longValue();
			
			Date BeforeDate = new Date(now.getTime() - beforedays);				
			   
			daysBI=null;
			msecondBI=null;			
			now=null;
			nowcalendar=null;
			return new String((t.format(BeforeDate)).toString());  
			
		}catch(Exception e){		   			      
		   	return "";
		}		
	}	
	
	/* 2004.12.06 add by 2295
	 * sendMailNotification
	 * suss_fail == true :更新成功; suss_fail == false : 更新失敗 
	 * input_method == "W" :線上編輯; input_method == "F" :檔案上傳;input_method=="C":檔案己被鎖定
	 * 94.03.28 fix suss_fail boolean ->String
	 */
	public static void sendMailNotification1(String bank_no, String report_no,
											String m_year, String m_month, String suss_fail,
											String szinput_method,String add_date,String filename,String userid)
											throws Exception {
		
		
		/*
		String SMTP_HOST = getProperties("SMTP_Host");
		String FROM_ADDR = getProperties("From_Addr");
		String SEND_MAIL = getProperties("Send_Mail");
		String WebSite = getProperties("WebSite");
		String Debug = getProperties("Debug");
		String input_method = szinput_method.equals("W") ? "線上編輯" : "檔案上傳";			
		String txt	= "";
		String m_email="";			
		
		System.out.println("Debug="+Debug);
		if(Debug.equals("true")){
			System.out.println("SMTP_HOST="+SMTP_HOST);
			System.out.println("FROM_ADDR="+FROM_ADDR);
			System.out.println("SEND_MAIL="+SEND_MAIL);				
		}	
		if(!SEND_MAIL.equals("true")) return;
		
		try{		
			List dbData = DBManager.QueryDB("select * from muser_data where muser_id='"+userid+"'","");
			if(dbData.size() != 0){
			   m_email = (String)((DataObject)dbData.get(0)).getValue("m_email");
			}
			System.out.println("m_email='"+m_email+"'");
			if (m_email.equals(""))	return;
			
			Notification ni = new Notification("MAIL");//91.04.22	   
	        ni.setSMTP(SMTP_HOST);//可不下此參數，若要下，<b>需在init之前</b><br>	        
	        ni.setDebug(Boolean.valueOf(Debug).booleanValue());//可不下此參數，若要下，<b>需在init之前</b>(預設為false)<br>       
	        ni.init(FROM_ADDR,m_email);        	
	        //ni.init(m_email,m_email);
	       	//要指定寄件方式，改用ni.init("寄信人Mail帳號","收信人Mail帳號一,收信人Mail帳號二","寄件方式");<br>	      	
	       	ni.setSubject("機構代號:"+bank_no+"--申報資料更新結果通知--" + filename);//主旨
	       	txt = "機構代號：" + bank_no + "\n";
	       	txt = "報表編號：" + report_no + "\n";
	       	txt += "資料基準日：" + m_year + "年" + m_month + "月\n";
	       	txt += "申報日期：" + add_date + "\n";
	       	txt += "申報方式：" + input_method + "\n";
	       	txt += "檢核結果：";
	       	//94.03.15 add 該申報資料已被鎖定,無法更新資料 by 2295
	       	if(szinput_method.equals("C")){
	       		txt += "該申報資料已被鎖定,無法更新資料\n";
	       	}else{
	       	   if (suss_fail.equals("true")){
	       	       txt += "檢核成功\n";
	       	   }else if (suss_fail.equals("Z")){
		       	       txt += "檢核為0\n";		       	   
	       	   }else {
	       	       txt += "檢核有誤\n請查詢原因後,至(網際網路申報系統)更正申報資料";
	       	   }	  
	       	}
	        ni.write(txt);
	        ni.send();
		}catch(Exception e){
			System.out.println("Utility.sendMailNotification Error:"+e+e.getMessage());	
		}
	    */    	
	} // End of sendMailNotification()
	
	/*******************************************************************************
	 * 2004.12.06 add by 2295
	 * sendMailNotification
	 * suss_fail == true :更新成功; suss_fail == false : 更新失敗 
	 * input_method == "W" :線上編輯; input_method == "F" :檔案上傳;input_method=="C":檔案己被鎖定
	 * sendMsg 欲傳送的內容
	 * 94.03.28 fix suss_fail boolean ->String
	 * 95.03.20 add sendMsg 欲傳送的內容 
	 */
	public static void sendMailNotification(String bank_no, String report_no,
				String m_year, String m_month, String suss_fail,
				String szinput_method,String add_date,String filename,String userid,String sendMsg)
	throws Exception {
				//throws Exception {
		
		String cc = null, bcc = null, url = null;
		String SMTP_HOST = getProperties("SMTP_Host").trim();
		String FROM_ADDR = getProperties("From_Addr").trim();
		String SEND_MAIL = getProperties("Send_Mail").trim();
		String UserID	= getProperties("UserID").trim();
		String PWD		= getProperties("PWD").trim();		
		String Auth		= (getProperties("Auth")==null?"false":getProperties("Auth")).trim();
		String WebSite = getProperties("WebSite").trim();
		String input_method = szinput_method.equals("W") ? "線上編輯" : "檔案上傳";
		String to="";
		String txt	= "";
		boolean debug = true;
		boolean verbose = true;
		boolean auth = Boolean.getBoolean(Auth);
		List paramList = new ArrayList();
		System.out.println("SMTP_HOST='"+SMTP_HOST+"'");
		System.out.println("FROM_ADDR='"+FROM_ADDR+"'");
		System.out.println("SEND_MAIL='"+SEND_MAIL+"'");
		System.out.println("UserID='"+UserID+"'");
		System.out.println("PWD='"+PWD+"'");
		System.out.println("Auth='"+Auth+"'");
        
		try {
				if (!SEND_MAIL.equals("true")){
					 System.out.println("Do no send mail");
					 return;
				}			
				paramList.add(userid);			
				List dbData = DBManager.QueryDB_SQLParam("select * from muser_data where muser_id=?",paramList,"");
				if(dbData.size() != 0){
					to = (String)((DataObject)dbData.get(0)).getValue("m_email");
				}
				System.out.println("m_email='"+to+"'");
				if (to.equals(""))	return;
                 
				//設定所要用的Mail 伺服器和所使用的傳送協定
				Properties props = System.getProperties();
				if (SMTP_HOST != null) props.put("mail.smtp.host", SMTP_HOST);
				if (auth) props.put("mail.smtp.auth", "true");

				// Get a Session object
				Session session = Session.getInstance(props, null);
				if (debug) session.setDebug(true);

				//construct the message
				Message msg = new MimeMessage(session);
				

				
				if (FROM_ADDR != null){					
					msg.setFrom(new InternetAddress(FROM_ADDR));					
				}else{
					msg.setFrom();
				} 	
				if(!to.equals("")){					
				    msg.setRecipients(Message.RecipientType.TO,
					InternetAddress.parse(to, false));				    
				}    
				if (cc != null){					
					msg.setRecipients(Message.RecipientType.CC,
				    InternetAddress.parse(cc, false));
				}	
				if (bcc != null){					
					msg.setRecipients(Message.RecipientType.BCC,
				    InternetAddress.parse(bcc, false));
				}
				if(sendMsg.equals("")){	
				   msg.setSubject("機構代號:"+bank_no+"--申報資料更新結果通知--" + filename);//主旨
				}else{
				    //95.04.04 fix 主旨為
				    //					貴中心(XXXXXXX) 執行95年3月 A05 代傳檢核結果清單如附內容!  
                    //					註:  XXXXXXX   為 該中心 的總機構代號 
				   msg.setSubject("貴中心["+bank_no+"]執行" +m_year + "年" + m_month + "月 "+report_no+ " 代傳檢核結果清單如附內容!");//主旨
				   //貴中心 執行95年1月 A01 代傳檢核結果清單如附內容!
				}
				msg.setSentDate(new Date());
				if(sendMsg.equals("")){
				   txt += "機構代號：" + bank_no + "\n"
				       +  "報表編號：" + report_no + "\n"
				       +  "資料基準日：" + m_year + "年" + m_month + "月\n"
				       +  "申報日期：" + add_date + "\n"
				       +  "申報方式：" + input_method + "\n"
				       +  "檢核結果：";

				       //94.03.15 add 該申報資料已被鎖定,無法更新資料 by 2295
				       if(szinput_method.equals("C")){
					      txt += "該申報資料已被鎖定,無法更新資料\n";
				       }else{
					      if (suss_fail.equals("true")){
						      txt += "檢核成功\n";
					      }else if (suss_fail.equals("Z")){
						      txt += "申報資料內容值皆為零\n";		       	   
					      }else {
						      txt += "檢核有誤\n請查詢原因後,至(網際網路申報系統)更正申報資料";
					      }	  
				       }
				}else{
				   txt = sendMsg; 
				   if(txt.indexOf("檢核有誤") != -1){//95.04.17 fix 檢核有誤時,只出現一次查詢原因
				       txt += "請查詢原因後,至(網際網路申報系統)更正申報資料\n";
				   }
				}    
				msg.setText(txt);
				
				SMTPTransport t = (SMTPTransport)session.getTransport("smtp");
			    try {
			    	if (auth){
			    		System.out.println("=========auth begin==============");
			    		//System.out.println("connect.SMPT_HOST= "+SMTP_HOST);
			    		//System.out.println("connect.UserID= "+UserID);
			    		//System.out.println("connect.PWD= "+PWD);
			    		t.connect(SMTP_HOST, UserID, PWD);
			    		System.out.println("=========auth end================");
			    	}else{
			    		t.connect();
			    	}    
			    	t.sendMessage(msg, msg.getAllRecipients());
			    } finally {
			    	//if (verbose) System.out.println("Response: " + t.getLastServerResponse());
			    	t.close();
			   }//end of try connect
		} catch (Exception e) {
			 if (verbose) e.printStackTrace();
			 if (e instanceof SendFailedException) {
			 	try{
					MessagingException sfe = (MessagingException)e;
				    if (sfe instanceof SMTPSendFailedException) {
					    SMTPSendFailedException ssfe = (SMTPSendFailedException)sfe;
					    System.out.println("SMTP SEND FAILED:");
					    if (verbose){
						    System.out.println(ssfe.toString());
					        System.out.println("  Command: " + ssfe.getCommand());
					        System.out.println("  RetCode: " + ssfe.getReturnCode());
					        System.out.println("  Response: " + ssfe.getMessage());
					    }
					} else {
					   if (verbose) System.out.println("Send failed: " + sfe.toString());
					}
					Exception ne;
					while ((ne = sfe.getNextException()) != null && ne instanceof MessagingException) {
					    sfe = (MessagingException)ne;
					    if (sfe instanceof SMTPAddressFailedException) {
						    SMTPAddressFailedException ssfe = (SMTPAddressFailedException)sfe;
						    System.out.println("ADDRESS FAILED:");
						    if (verbose){
						        System.out.println(ssfe.toString());
						        System.out.println("  Address: " + ssfe.getAddress());
						        System.out.println("  Command: " + ssfe.getCommand());
						        System.out.println("  RetCode: " + ssfe.getReturnCode());
						        System.out.println("  Response: " + ssfe.getMessage());
						    }
					    } else if (sfe instanceof SMTPAddressSucceededException) {
						    System.out.println("ADDRESS SUCCEEDED:");
						    SMTPAddressSucceededException ssfe = (SMTPAddressSucceededException)sfe;
						    if (verbose){
						        System.out.println(ssfe.toString());
						        System.out.println("  Address: " + ssfe.getAddress());
						        System.out.println("  Command: " + ssfe.getCommand());
						        System.out.println("  RetCode: " + ssfe.getReturnCode());
						        System.out.println("  Response: " + ssfe.getMessage());
						    }
					    }
					}//end of while
			 	}catch(Exception e1){
			 		System.out.println("Send Mail Error:"+e1.getMessage());
			 		//throw(new Exception("傳送email失敗"));
			 	}
			 }//end of instanceof SendFailedException
		}//end of catch
} // End of sendMailNotification()

	/*******************************************************************************
	 * 2006.03.20 add by 2295
	 * addMailNotification
	 * suss_fail == true :更新成功; suss_fail == false : 更新失敗 
	 * input_method == "W" :線上編輯; input_method == "F" :檔案上傳;input_method=="C":檔案己被鎖定
	 * 94.03.28 fix suss_fail boolean ->String
	 */
	public static void addMailNotification(String bank_no, String report_no,
				String m_year, String m_month, String suss_fail,
				String szinput_method,String add_date,String filename,String userid)
	throws Exception {
		String input_method = szinput_method.equals("W") ? "線上編輯" : "檔案上傳";		
		String txt	= "";
		
		try {			
		    
		    /* ex:
		     * 總機構代號   報表編號   資料基準日    申報日期              申報方式    檢核結果
		        XXXXXXX     A01       95年1月  2004/12/22 11:41:26   檔案上傳      失敗
		    */    
		       if(bank_no.equals("")){		          
		          txt += "總機構代號  報表編號   資料基準日     申報日期           申報方式      檢核結果\n"
		              +  "================================================================================\n";		          
		       }else{
		           txt += bank_no + "     "
		                + report_no + "        "
		                + m_year + "年" + m_month + "月     "
		                + add_date + "   "
		                + input_method + "      ";

		           //94.03.15 add 該申報資料已被鎖定,無法更新資料 by 2295
		           if(szinput_method.equals("C")){
		               txt += "該申報資料已被鎖定,無法更新資料\n";
		           }else{
		               if (suss_fail.equals("true")){
		                   txt += "檢核成功\n";
		               }else if (suss_fail.equals("Z")){
		                   txt += "申報資料內容值皆為零\n";		       	   
		               }else {
		                   txt += "檢核有誤\n";//95.04.17 fix 檢核有誤時,只出現一次查詢原因
		               }	  
		           }
		       }
			   sendMsg += txt;
		} catch (Exception e) {	
			System.out.println("addMailNotification Error:"+e.getMessage());
		}//end of catch
} // End of addMailNotification()
	
	/*******************************************************************************
	 * 2009.08.11 add by 2295
	 * sendMail
	 * mail_to mail address 
	 * mailsubject 主旨
	 * sendMsg 欲傳送的內容 
	 */
	public static void sendMail(String mail_to,String mailsubject,String sendMsg)
	throws Exception {				
		String cc = null, bcc = null, url = null;
		String SMTP_HOST = getProperties("SMTP_Host").trim();
		String FROM_ADDR = getProperties("From_Addr").trim();
		String SEND_MAIL = getProperties("Send_Mail").trim();
		String UserID	= getProperties("UserID").trim();
		String PWD		= getProperties("PWD").trim();		
		String Auth		= (getProperties("Auth")==null?"false":getProperties("Auth")).trim();
		String txt	= "";
		boolean debug = true;
		boolean verbose = true;
		boolean auth = Boolean.getBoolean(Auth);
		
		System.out.println("SMTP_HOST='"+SMTP_HOST+"'");
		System.out.println("FROM_ADDR='"+FROM_ADDR+"'");
		System.out.println("SEND_MAIL='"+SEND_MAIL+"'");
		System.out.println("UserID='"+UserID+"'");
		System.out.println("PWD='"+PWD+"'");
		System.out.println("Auth='"+Auth+"'");
        
		try {
				if (!SEND_MAIL.equals("true")){
					 System.out.println("Do no send mail");
					 return;
				}			
					
				System.out.println("email_to='"+mail_to+"'");
				if (mail_to.equals(""))	return;
                 
				//設定所要用的Mail 伺服器和所使用的傳送協定
				Properties props = System.getProperties();
				if (SMTP_HOST != null) props.put("mail.smtp.host", SMTP_HOST);
				if (auth) props.put("mail.smtp.auth", "true");

				// Get a Session object
				Session session = Session.getInstance(props, null);
				if (debug) session.setDebug(true);

				//construct the message
				Message msg = new MimeMessage(session);
				
				if (FROM_ADDR != null){					
					msg.setFrom(new InternetAddress(FROM_ADDR));					
				}else{
					msg.setFrom();
				} 	
				if(!mail_to.equals("")){					
				    msg.setRecipients(Message.RecipientType.TO,
					InternetAddress.parse(mail_to, false));				    
				}    
				if (cc != null){					
					msg.setRecipients(Message.RecipientType.CC,
				    InternetAddress.parse(cc, false));
				}	
				if (bcc != null){					
					msg.setRecipients(Message.RecipientType.BCC,
				    InternetAddress.parse(bcc, false));
				}
				
				msg.setSubject(mailsubject);//主旨
				
				msg.setSentDate(new Date());
				
				txt = sendMsg;
				msg.setText(txt);//內容
				
				SMTPTransport t = (SMTPTransport)session.getTransport("smtp");
			    try {
			    	if (auth){
			    		System.out.println("=========auth begin==============");
			    		t.connect(SMTP_HOST, UserID, PWD);
			    		System.out.println("=========auth end================");
			    	}else{
			    		t.connect();
			    	}    
			    	t.sendMessage(msg, msg.getAllRecipients());
			    } finally {
			    	//if (verbose) System.out.println("Response: " + t.getLastServerResponse());
			    	t.close();
			   }//end of try connect
		} catch (Exception e) {
			 if (verbose) e.printStackTrace();
			 if (e instanceof SendFailedException) {
			 	try{
					MessagingException sfe = (MessagingException)e;
				    if (sfe instanceof SMTPSendFailedException) {
					    SMTPSendFailedException ssfe = (SMTPSendFailedException)sfe;
					    System.out.println("SMTP SEND FAILED:");
					    if (verbose){
						    System.out.println(ssfe.toString());
					        System.out.println("  Command: " + ssfe.getCommand());
					        System.out.println("  RetCode: " + ssfe.getReturnCode());
					        System.out.println("  Response: " + ssfe.getMessage());
					    }
					} else {
					   if (verbose) System.out.println("Send failed: " + sfe.toString());
					}
					Exception ne;
					while ((ne = sfe.getNextException()) != null && ne instanceof MessagingException) {
					    sfe = (MessagingException)ne;
					    if (sfe instanceof SMTPAddressFailedException) {
						    SMTPAddressFailedException ssfe = (SMTPAddressFailedException)sfe;
						    System.out.println("ADDRESS FAILED:");
						    if (verbose){
						        System.out.println(ssfe.toString());
						        System.out.println("  Address: " + ssfe.getAddress());
						        System.out.println("  Command: " + ssfe.getCommand());
						        System.out.println("  RetCode: " + ssfe.getReturnCode());
						        System.out.println("  Response: " + ssfe.getMessage());
						    }
					    } else if (sfe instanceof SMTPAddressSucceededException) {
						    System.out.println("ADDRESS SUCCEEDED:");
						    SMTPAddressSucceededException ssfe = (SMTPAddressSucceededException)sfe;
						    if (verbose){
						        System.out.println(ssfe.toString());
						        System.out.println("  Address: " + ssfe.getAddress());
						        System.out.println("  Command: " + ssfe.getCommand());
						        System.out.println("  RetCode: " + ssfe.getReturnCode());
						        System.out.println("  Response: " + ssfe.getMessage());
						    }
					    }
					}//end of while
			 	}catch(Exception e1){
			 		System.out.println("Utility.Send Mail Error:"+e1.getMessage());
			 		//throw(new Exception("傳送email失敗"));
			 	}
			 }//end of instanceof SendFailedException
		}//end of catch
} // End of sendMailNotification()
	/***************************************************************************
	 * 讀出資料時使用
	 * Convert Big5 to ISO8859_1
	 */
	public static String Big5toISO(String s) {
		if (s == null)
			s = "";
		else
			s = s.trim();

  	    if (SetConfig.doConvert) {
			try {
				return new String(s.getBytes("Big5"), "ISO8859_1");
   		    }
   		    catch (UnsupportedEncodingException e) {
    		   	return s;
   		    }
   		}
   		else return s;
  	}

	/***************************************************************************
	 *
	 */
	public static String ChangeFormat(String Oring) {
		int pos;
		pos = Oring.indexOf(",");

	 	while (pos !=-1) {
        	Oring = (Oring).substring(0,pos) + (Oring).substring(pos+1, Oring.length());
            pos   = Oring.indexOf(",");
        }
		return Oring;
	}

	/***************************************************************************
	 * 將文字型數字中的","拿掉;
	 * ex:123,000 -> 123000
	 */
	public static String setNoCommaFormat(String str) {

		if (str.trim().equals("") || str == null)
			return "0";
		else {
	        StringTokenizer paser = new StringTokenizer(str.trim());
	        StringBuffer newstr	= new StringBuffer("");

	        while (paser.hasMoreTokens())
	            newstr.append(paser.nextToken(","));

	        return newstr.toString();

    	}
    }

	/***************************************************************************
	 * 寫入資料時使用
	 * Convert ISO8859_1 to Big5
	 */
  	public static String ISOtoBig5(String s) {
  		if (s == null)
			s = "";
		else
			s = s.trim();

        if (SetConfig.doConvert) {
    		try {
      			return new String(s.getBytes("ISO8859_1"), "Big5");
    		}
    		catch (UnsupportedEncodingException e) {
      			return s;
    		}
        }
        else return s;
    }
    //104.03.09 add
  	public static String ISOtoUTF8(String s) {
        if (s == null)
            s = "";
        else
            s = s.trim();

        if (SetConfig.doConvert) {
            try {
                return new String(s.getBytes("ISO8859_1"), "UTF-8");
            }
            catch (UnsupportedEncodingException e) {
                return s;
            }
        }
        else return s;
    }
  	
  	public  static String toBig5Convert(String s) 
	{
    
    		try {
      			return new String(s.getBytes(),"ISO8859_1");
    		}catch(UnsupportedEncodingException e) {
      			return s;
    		}	
  	}
  	
	/***************************************************************************
	 *  將傳入的數字型文字轉成千為單位的數字文字
	 *  ex: 1000000 -> 1,000,000
	 */
    //public static String setCommaFormat(String szNumString, PrintWriter out) {
    public static String setCommaFormat(String szNumString) {

        String szManufacture = szNumString;
        String szSign        = ""; //正負的符號
        String szDecimal     = ""; //放小數的變數
        //System.out.println("szNumStrin="+szNumString);
        try {
			if (szNumString == null)	return "";

            if (szNumString.trim().equals(""))
                return szNumString;

            if (szNumString.indexOf(".") != -1) { //有小數點
                szDecimal =  szNumString.substring(szNumString.indexOf("."),szNumString.length());// 將小數取出
                szNumString = szNumString.substring(0,szNumString.indexOf("."));
                //System.out.println("szDecimal="+szDecimal);
                //System.out.println("szNumString="+szNumString);
            }

            if (new Double(szNumString).doubleValue() < 0) {//若為負數
                szSign = "-"; 								// 符號為"-"
                szNumString =  szNumString.substring(1);	//把負號後的數字取出
            }

            int iTimes     = (szNumString.length() / 3) - 1;//要加","的次數(要扣掉第1個","號)
            int iFirstTime = szNumString.length() % 3;   	//第一個","的位置

            if (iFirstTime != 0 && iTimes >= 0)
                szManufacture = szNumString.substring(0,iFirstTime) + ",";
            else if (iTimes < 0) 							// 表示數字小於千
                    szManufacture = szNumString;
                 else 										//剛好為3的倍數
                    szManufacture = "";
            for (int i = 0; i < iTimes; i++) {
                szManufacture += szNumString.substring(iFirstTime, iFirstTime + 3) + ",";
                iFirstTime += 3;
            }

            szManufacture += szNumString.substring(iFirstTime);//把最後一段數字加上

            if (szSign.equals("-")) 						//若為負數，則加上負號
                szManufacture = "-" + szManufacture;

            szManufacture += szDecimal;

        }
        catch (Exception e) {
        	System.out.println(e);return szNumString;
        }

        return szManufacture;

    }
	/***************************************************************************
	 *  將傳入的日期字串轉換格式
	 * i==0 ex: 2002/09/23->91/09/23
	 * i==1 ex: 2002/09/23->91年09月23日
	 */
    public static String getCHTdate(String s, int i) {

		String CHTdate	= null;

		if ((s == null) || s.equals("")) return "";
		if (s.trim().length() != 10)	return s.trim();

		int	yr	= Integer.parseInt(s.substring(0,4)) - 1911;
		if (i == 0)
			CHTdate = Integer.toString(yr) + "/" + s.substring(5, 7) + "/" + s.substring(8, 10);
		if (i == 1)
			CHTdate = Integer.toString(yr) + "年" + s.substring(5, 7) + "月" + s.substring(8) + "日";

		return CHTdate;
	}
    
    /***************************************************************************
	 *  將傳入的日期字串把"/"去掉
	 * ex: 2002/09/23->20020923	  
	 */
    public static String getDatetoString(String s) {
		String date	= "";
		if ((s == null) || s.equals("")) return "";
        
		StringTokenizer st = new StringTokenizer(s,"/");
	     while (st.hasMoreTokens()) {
	         date += st.nextToken();
	     }		
		return date;
	}

    
	/***************************************************************************
	 *  將傳入的字串轉換為日期格式
	 * 91 11 12 -> 2002/11/12
	 */
    public static String getFullDate(String yr, String mn, String dt) {

		String fulldate = "";

		if (yr == null || mn == null || dt == null)	return "";
		if (yr.trim().equals("") || mn.trim().equals("") || dt.trim().equals("")) return "";

		int	year = Integer.parseInt(yr.trim()) + 1911;
		fulldate = Integer.toString(year) + "/" + mn + "/" + dt;

		return fulldate;
	}

	/***************************************************************************
	 *  將傳入的數字轉換為百分比的格式
	 * 00150 --> 0.15
	 */
    public static String getPercentNumber(String szNumString) {
        try {
	    	double dDouble=Double.parseDouble(szNumString)/1000;
	    	return String.valueOf(dDouble);
        }catch (Exception e) {
        	System.out.println(e);return szNumString;
        }
	}    
	
	/* 將0.15 --> 150 
		modify by 2354 2005.1.5 add when szNumString is 0 or null return 0*/
	public static String setNoPercentFormat(String szNumString) {
		if (szNumString.trim().equals("") || szNumString == null)
			return "0";
		else{
	        try {
				DecimalFormat df_md = new DecimalFormat("############0");
				double dDouble = Math.round(Double.parseDouble(szNumString)*1000);
				return df_md.format(dDouble).toString();
	        }catch (Exception e) {
	        	System.out.println(e);return szNumString;
	        }
		}
    }
	
	/***************************************************************************
	 *  將傳入的數字轉換為百分比的格式
	 * 001155 --> 11.55
	 */
    public static String getPercentNumber_2(String szNumString) {
        try {
	    	double dDouble=Double.parseDouble(szNumString)/100;
	    	DecimalFormat df_md = new DecimalFormat("############0.00");
	    	return df_md.format(dDouble).toString();	    	
        }catch (Exception e) {
        	System.out.println(e);return szNumString;
        }
	}
	//94.01.18 add 讀取file data 儲存在List by 2295
	/***************************************************************************
	 * 讀取file data 儲存在List 
	 */
	public static List getFileData(String szFileName){
		String sztemp = null;
		List data = new LinkedList();
		
		try{
			File WorkFile = new File(szFileName);	
			if(WorkFile.exists()){			
				FileInputStream fis = new FileInputStream(szFileName);	
				BufferedInputStream bis = new BufferedInputStream(fis);
				DataInputStream dis = new DataInputStream(fis);							
				while((sztemp = dis.readLine()) != null){
					data.add(sztemp);					
				}			
				dis.close();
				bis.close();
				fis.close();
			}
			WorkFile = null;
		}catch(Exception e){
			System.out.println("getFileData Error:"+e+e.getMessage());
		}
		return data;
	}
	
	/***************************************************************************
	 *  計算日期相距的天數
	 *  2005/01/10 ~ 2005/01/30 = 20 天
	 */
	public static int date_range(String begin_date,String end_date){
	       int range=0;
	       int year=0;
	       int month=0;
	       int day=0;
	       StringTokenizer token = null;
	       Calendar cal=null;
	       try{
	       		if (begin_date.equals("")) return -1;
	       		if (end_date.equals("")) return -1;
	       
	       		token = new StringTokenizer(begin_date,"/"); 
	            year = Integer.parseInt(token.nextToken()); 
	            month = Integer.parseInt(token.nextToken()); 
	            day = Integer.parseInt(token.nextToken()); 
	             
	            cal = Calendar.getInstance();
	            cal.set(year,month,day);       
	            Date d1 = cal.getTime();
	            System.out.println("d1="+d1);
	            token = new StringTokenizer(end_date,"/"); 
	            year = Integer.parseInt(token.nextToken()); 
	            month = Integer.parseInt(token.nextToken()); 
	            day = Integer.parseInt(token.nextToken());               
	            cal.set(year,month,day);         
	            Date d2 = cal.getTime();
	            System.out.println("d2="+d2);
	            long daterange = d2.getTime() - d1.getTime();
	            System.out.println("daterange="+daterange);
	            long time = 1000*3600*24; //A day in milliseconds       
	            range=(int)(daterange/time);
	       }catch(Exception e){
	       		System.out.println("date_range Error:"+e+e.getMessage());
	       		return -1;
	       }
	       return range;
	 }     
	
	/***************************************************************************
	 *  取得該月份的最後一天	   
	 */
	public static String getLastDay(String date,String dateformat){
		String lastDay = "-1";
		List paramList = new ArrayList();
		try{
		      paramList.add(date);
		      paramList.add(dateformat);
		      List dbData = DBManager.QueryDB_SQLParam("select LAST_DAY(Trunc(to_date(?,?),'MONTH'))+1-1/86400 Last_Day_Month FROM dual",paramList,"Last_Day_Month");
		      if(dbData != null && dbData.size() != 0){		  	     
		  	     lastDay = ((((DataObject)dbData.get(0)).getValue("last_day_month")).toString());
		      }
		}catch(Exception e)   {
			System.out.println("getLastDay Error:"+e+e.getMessage());
		}
		return lastDay;
	}

	/***************************************************************************
	 *  取得N月天後的日期	   
	 */
	public static String getAfterDay(int year,int month,int day,int afterday ){
		String afterDate = "-1";
		try{	
			//GregorianCalendar worldTour = new GregorianCalendar(2005, Calendar.FEBRUARY, 1); //->2005/2/1 
			//GregorianCalendar worldTour = new GregorianCalendar(2005, 1, 1); ->2005/2/1 ..從0開始算;0->1月:1->2月 
			GregorianCalendar worldTour = new GregorianCalendar(year, month-1, day);
			worldTour.add(GregorianCalendar.DATE, afterday); 
            Date d = worldTour.getTime(); 
            DateFormat df = DateFormat.getDateInstance(); 
            afterDate = df.format(d);             
		}catch(Exception e)   {
			System.out.println("getAfterDay Error:"+e+e.getMessage());
		}
		return afterDate;
	}	
	/***************************************************************************
	 *  取得已被鎖定的bank_code	   
	 */
	public static List getLockBank_Code(String m_year,String m_month,String Report_no){
		    List bank_code=null;
		    List paramList = new ArrayList();
	    	String sqlCmd = " select bank_code ,wml01_lock_status ,wml01_lock_lock_status"
	    		          + " FROM wml01_a_v"
					      + " where  m_year  = ?"  
					      + "	and   m_month = ?" 
					      + "	and   report_no = ?"
					      + "  and ( (wml01_lock_status is not null and wml01_lock_status = 'Y')"
					      + "	or (wml01_lock_lock_status is not null and (wml01_lock_lock_status = 'Y' or wml01_lock_lock_status = 'C')))";
	    	paramList.add(m_year);
	    	paramList.add(m_month);
	    	paramList.add(Report_no);
	    	
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			if(dbData != null && dbData.size() != 0){
			   bank_code = new LinkedList();
			   for(int i=0;i<dbData.size();i++){
			   	   bank_code.add((String)((DataObject)dbData.get(i)).getValue("bank_code"));
			   }
			}	
			return bank_code;	    
	}
	/***************************************************************************
	 *  設定sendMsg	   
	 */
    public static void setSendMsg(String szSendMsg){
        sendMsg = szSendMsg;
    }
    public static String getSendMsg(){
        return sendMsg;
    }
    
    /*add by Allen Chang 20060418 */
    // 取回字串, 如果為null 傳回""
    //這方法是幫助判斷字串物件是否為null , 如果是null 則回傳 空白, 不是則傳回原字串
    //同 String str = object != null ? (String) object : “”;
    public static String getTrimString(Object str) {
        return str != null ? ((String)str).trim() : "";
    }
     /*add by Allen Chang 20060418 */
    // 保存查詢條件
    public static Map saveSearchParameter(ServletRequest request) {
        Enumeration enu = request.getParameterNames();
        Map h = new HashMap();
        while(enu.hasMoreElements()) {
            String name = (String)enu.nextElement();
            String value = Utility.getTrimString(request.getParameter(name));
            System.out.println(" Save Request Attribute   :  " + name + " = " + value);
            h.put(name, value);
        }
        return h;
    }
    
    //95.06.06 add 取得申報年 by 2295
    public static String getYear(){        
        Calendar now = Calendar.getInstance();
       	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
       	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
        if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
           YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
           MONTH = "12";
        }else{    
          MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
        }
        return YEAR;
    }
    //95.06.06 add 取得申報月 by 2295
    public static String getMonth(){        
        Calendar now = Calendar.getInstance();
       	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
       	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
        if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
           YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
           MONTH = "12";
        }else{    
          MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
        }
        return MONTH;
    }
    

    /****************************************************************************
     * 取得民國年份/月份
     * yymmdd=yy:回傳民國年份
     * yymmdd=mm:回傳民國月份
     * yymmdd=dd:回傳民國日期
     */
    public static String getCHTYYMMDD(String yymmdd){   
    	String chtYM = "";
        Calendar now = Calendar.getInstance();
       	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
       	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
       	String DAY = String.valueOf(now.get(Calendar.DAY_OF_MONTH));//日期
        if(yymmdd.equals("yy")) chtYM = YEAR;
        if(yymmdd.equals("mm")) chtYM = MONTH;
        if(yymmdd.equals("dd")) chtYM = DAY;
        return chtYM;
    }
    
    //95.08.07 add 將dataList組合成a,b,c,d by  2295
    public static String getInString(List dataList,String strValue){
        String inString = "";          
        try{    
             for(int i=0;i<dataList.size();i++){
                 inString += " '"+((DataObject)dataList.get(i)).getValue(strValue)+"'";                 
                 if( i != dataList.size() -1){
                     inString += ",";
                 }                   
             }
        }catch(Exception e){
            System.out.println("getInString error:"+e.getMessage());
        }
        return inString;
    }    
    //96.04.18 add 將List組合成a,b,c,d by  2295
    public static String getCombinString(List dataList,String strValue){
        String inString = "";          
        try{    
             for(int i=0;i<dataList.size();i++){
                 inString += " '"+(String)dataList.get(i)+"'";                 
                 if( i != dataList.size() -1){
                     inString += strValue;
                 }                   
             }
        }catch(Exception e){
            System.out.println("getInString error:"+e.getMessage());
        }
        return inString;
    }    
    //95.08.09 add 取得該用戶檢查追蹤權限的總機構資料 by 2295
    //96.01.12 fix 修改sql,縣市別不為空白時,才加入屬於該縣市別
    public static List getTBankNO(HttpServletRequest request,String szmuser_id,String szbank_type,String szhsien_id){
        System.out.println("Utility.getTBankNO.szbank_type="+szbank_type);
        System.out.println("Utility.getTBankNO.szhsien_id="+szhsien_id);        
	    HttpSession session = request.getSession();    
	    List bank_type = (List)session.getAttribute("Bank_Type");	
	    List paramList = new ArrayList();
		String sqlCmd = " select wtt08.hsien_id, ba01.bank_no , ba01.bank_name, wtt08.bank_type  "
			   		  + " from  ba01, wlx01 ,wtt08 "
			   		  + " where ba01.bank_no = wlx01.bank_no(+)"
			   		  + " and wtt08.muser_id = ? ";
		paramList.add(szmuser_id);
        if(szbank_type.equals("")){//金融機構類別="",抓該用戶檢查追蹤管理的權限				   		  
			sqlCmd += " and ba01.bank_type in ("+Utility.getInString((List)session.getAttribute("Bank_Type"),"cmuse_id")+")";
		}else{
		    sqlCmd += " and ba01.bank_type = ?";
		    paramList.add(szbank_type);
		}	   		  
			sqlCmd += " and ba01.bank_no = wtt08.tbank_no ";		
		if(!szhsien_id.equals("")){//縣市別 !="",抓該用戶檢查追蹤管理的權限	
		    sqlCmd += " and wlx01.hsien_id = ?";
		    paramList.add(szhsien_id);
		}	   		  
			sqlCmd += " group by wtt08.hsien_id, ba01.bank_no , ba01.bank_name, wtt08.bank_type order by bank_no ";
		
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
	}
    
    //95.08.09 add 取得該用戶檢查追蹤權限的受檢單位資料 by 2295
    //96.01.12 fix 修改sql,改由查詢ba01,縣市別不為空白時,才加入屬於該縣市別
    public static List getBankNo(HttpServletRequest request,String szmuser_id,String szbank_type,String szhsien_id,String sztbank_no){
        System.out.println("Utility.getBankNo.szbank_type="+szbank_type);
        System.out.println("Utility.getBankNo.szhsien_id="+szhsien_id);
        System.out.println("Utility.getBankNo.sztbank_no="+sztbank_no);
	    HttpSession session = request.getSession();  
	    List paramList = new ArrayList();
		String sqlCmd = " select bank_no,bank_name,pbank_no  "
			   		  + " from  ba01 "
			   		  + " where bank_no in (select examine from wtt08 where muser_id=? ";
		paramList.add(szmuser_id);
		if(sztbank_no.equals("")){//總機構代號="",抓該用戶檢查追蹤管理的權限	   		  
		      sqlCmd += " and pbank_no in ("+Utility.getInString((List)session.getAttribute("TBank"),"bank_no")+")";
		}else{
		      sqlCmd += " and pbank_no=?";
		      paramList.add(sztbank_no);
		}
		/*
		if(szhsien_id.equals("")){//縣市別="",抓該用戶檢查追蹤管理的權限	   		  
			  sqlCmd += " and hsien_id in ("+Utility.getInString((List)session.getAttribute("hsien_id"),"hsien_id")+")";
		}else{
		*/
		if(!szhsien_id.equals("")){//縣市別!="",抓該用戶檢查追蹤管理的權限
			  sqlCmd += " and hsien_id=?";		
			  paramList.add(szhsien_id);
		}	  
		if(szbank_type.equals("")){//金融機構類別="",抓該用戶檢查追蹤管理的權限	
			  sqlCmd += " and bank_type in ("+Utility.getInString((List)session.getAttribute("Bank_Type"),"cmuse_id")+"))";
		}else{
		      sqlCmd += " and bank_type=?)";
		      paramList.add(szbank_type);
		}	   		  
			  sqlCmd += " order by bank_no ";		
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
	}
    
    //95.08.17 將來源字串,根據所設定的分隔符號,轉成List
    //         String srcData:來源字串
    //         String delim:分隔符號
    public static List getStringTokenizerData(String srcData,String delim){	    
	    StringTokenizer paserData = null;
	    List returnList = new LinkedList();
		try{
		    paserData = new StringTokenizer(srcData,delim);		     
	        while (paserData.hasMoreTokens()){
	            returnList.add(paserData.nextToken().trim());  
	        }          
		}catch(Exception e){
			System.out.println("getStringTokenizerData Error:"+e+e.getMessage());
		}
		return returnList;
	}
    //95.08.17 add 取得單位名稱 
    public static String getUnitName(String unit){
        String unit_name="";
        try{
            if (unit.equals("1")){
                unit_name="元";
            }else if (unit.equals("1000")){   	 	   		
                unit_name="千元";
            }else if (unit.equals("10000")){
                unit_name="萬元";
            }else if (unit.equals("1000000")){
 	 	  		unit_name="百萬元";
 	 	   	}else if (unit.equals("10000000")){
 	 	   	  	unit_name="千萬元";
        	}else if (unit.equals("100000000")){
        		unit_name="億元";
 	 	   	}
        }catch(Exception e){
            System.out.println("getUnitName Error:"+e.getMessage());
        }
        return unit_name;	
    }	
    //input  --> 110000+流動資產,110310+存放行庫-合作金庫
    //output --> [110000, 流動資產]
    //95.08.18 add getReportData 
    public static  List getReportData(String rptData){
	    List rptList = new LinkedList();
	    StringTokenizer paserData = null;
	    List rptDetail = null;
		try{
			StringTokenizer paser = new StringTokenizer(rptData.trim(),",");			
	        while (paser.hasMoreTokens()){
	            paserData = new StringTokenizer(paser.nextToken(","),"+");
	            rptDetail = new LinkedList();
	            while (paserData.hasMoreTokens()){
	                rptDetail.add(paserData.nextToken());  
	            }//end of have "+" data	            
	            rptList.add(rptDetail);
			}//end of have "," data 
		}catch(Exception e){
			System.out.println("getReportData Error:"+e+e.getMessage());
		}
		return rptList;
	}    
    //95.09.05 add 將金額(amt)除以單位(Unit)四捨五入
    public static String getRound(String amt,String Unit){
        String amt_convert="";
        if(amt.equals("") || amt == null) return amt_convert;
        double amt_d = 0.0;
        try{
            amt_d = Double.parseDouble(amt);                     
            //System.out.println(":amt.double="+amt_d);
            amt_d = amt_d / Double.parseDouble(Unit);
            //System.out.println(":amt.double="+amt_d);		                        
            //System.out.println(":amt.double="+Math.round(amt_d));
            amt_convert = String.valueOf(Math.round(amt_d));
        }catch(Exception e){
            System.out.println("Utility.getRound Error:"+e.getMessage());
        }
        return amt_convert;
    }
    //95.11.03 add 檢核權限
    public static boolean CheckPermission(HttpServletRequest request,String report_no){
	    boolean CheckOK=false;
	    HttpSession session = request.getSession();            
        Properties permission = ( session.getAttribute(report_no)==null ) ? new Properties() : (Properties)session.getAttribute(report_no);				                
        if(permission == null){
          System.out.println(report_no+".permission == null");
        }else{
           System.out.println(report_no+".permission.size ="+permission.size());
           
        }
        //只要有Query的權限,就可以進入畫面
    	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
    	   CheckOK = true;//Query
    	}
    	System.out.println("CheckOk="+CheckOK);        	
    	return CheckOK;
   }
    
    
    //95.11.13 add 取得可選機構代號權限設定(農.漁會機構)===================================================================================
    //99.03.30 add 區分新舊機構名稱
    //99.11.09 fix 修改機構名稱排列順序
    public static List getBankList(HttpServletRequest request){
        HttpSession session = request.getSession();
        List tbankList = null;
        List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
        try{
            //95.11.13 取得登入者資訊=================================================================================================
            String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
            String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
            String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");			
            String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
            String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
            //==============================================================================================================    	    
            String bank_type = (session.getAttribute("nowbank_type")==null)?"6":(String)session.getAttribute("nowbank_type");	    
    	
            sql.append(" select bn01.m_year,bn01.bn_type,bn01.bank_type,wlx01.hsien_id,bn01.bank_no,bn01.bank_name from bn01 ");
            sql.append(" LEFT JOIN WLX01 on bn01.bank_no = WLX01.bank_no");      
            //農會.漁會只顯示參加該共用中心.地方主管機關下的機構代碼============================	       
            if(muser_bank_type.equals("8") && (bank_type.equals("6") || bank_type.equals("7"))){//農漁會共用中心	       
            	sql.append(" and WLX01.center_no = ?");
            	paramList.add(muser_tbank_no);
                	  //+ " and   bn01.bank_type='"+bank_type+"'";//96.01.11拿掉   
            }else if(muser_bank_type.equals("B") && (bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("B"))){//地方主管機關  
            	sql.append(" and WLX01.m2_name = ?");
            	paramList.add(muser_tbank_no);
            }else if(muser_id.equals("A111111111") || muser_bank_type.equals("2")){//登入者為A11111111 or 農金局直接去抓在這分類(農會或漁會)所有的機構       
            	sql.append(" and bn01.bank_type = ?");
            	paramList.add(bank_type);
            }else if(muser_bank_type.equals("6") || muser_bank_type.equals("7")){//登入者為農漁會時,只能看到其所屬機構代號
            	sql.append(" and bn01.bank_no =? and bank_type =?");
                paramList.add(muser_tbank_no);
                paramList.add(muser_bank_type);
            }else if(bank_type.equals("3")){//95.11.27 add 中央存保.可看到農.漁會的
            	sql.append(" and bn01.bank_type in (?,?)");   
            	paramList.add("6");
            	paramList.add("7");
            }
            sql.append(" ,v_bank_location e");
            sql.append(" where wlx01.m_year=bn01.m_year ");            
            sql.append("   and bn01.m_year = e.m_year ");
            sql.append("   and bn01.bank_no = e.bank_no ");
            sql.append(" order by e.m_year,e.fr001w_output_order,bn01.bank_type,bn01.bank_no ");                       
            //====================================================================================================================
            tbankList =  DBManager.QueryDB_SQLParam(sql.toString(),paramList,"m_year");     
        }catch(Exception e){
            System.out.println("Utility.getBankList Error:"+e.getMessage());
        }
        return tbankList;
    }
    //將資料加密
    /*
     * 1.加/解密的用法
	     import com.tradevan.util.TvEncrypt;
	      加密後的密碼 = TvEncrypt.encode( 原始密碼 );
	      原始密碼 = TvEncrypt.decode( 加密後的密碼 );
	   2.雜湊的用法
	     import com.tradevan.util.TvHash;
	     String HashAlgorithm = "SHA-1";
	     加密後的密碼 = TvHash.Digest( 原始密碼 , HashAlgorithm);
    */
    public static String encode(String srcStr){
    	String encodeString = "";   	
    	//加密後的密碼 = TvEncrypt.encode( 原始密碼 );
    	//原始密碼 = TvEncrypt.decode( 加密後的密碼 );
        try{
        	encodeString = TvEncrypt.encode(srcStr);
        }catch(Exception e){
        	System.out.println("Utility.encode Error:"+e.getMessage());
        }
        return encodeString;
    	
    }
    //將資料解密
    public static String decode(String srcStr){
    	String decodeString = "";   	
    	//加密後的密碼 = TvEncrypt.encode( 原始密碼 );
    	//原始密碼 = TvEncrypt.decode( 加密後的密碼 );
        try{
        	decodeString = TvEncrypt.decode(srcStr);
        }catch(Exception e){
        	System.out.println("Utility.dencode Error:"+e.getMessage());
        }
        return decodeString;
    	
    }
    
    //96.07.31 add 批次寫入WML03_LOG
    //96.11.21 fix 區分檔案上傳/線上編輯
    //102.10.02 fix 套用preparestatment by 2295
	public static boolean InsertWML03_LOG(String m_year,String m_month,String user_id,String user_name,String report_no,String filename,String input_method,List bank_codeList){
		boolean updateOK=false;
		//String sqlCmd="";		
		Utility.printLogTime("InsertWML03_LOG begin time");
	
		StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList() ;
        List updateDBList = new ArrayList();//0:sql 1:data
        List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
        List updateDBDataList = new ArrayList();//儲存參數的List
		try{
			/*
			 select WML03.m_year,WML03.m_month,WML03.bank_code,WML03.report_no,WML03.serial_no,
			        WML03.remark,WML03.user_id,WML03.user_name,WML03.update_date,'A111111111','A111111111',sysdate,'D'
			 from WML03,(select m_year,m_month,bank_code,report_no 
			 		     from AXX_TMP 
			 		     where m_year=96 and m_month=2
			 		     and report_no='A01'
			 		     group by m_year,m_month,bank_code,report_no)AXX_TMP
			 where WML03.m_year= AXX_TMP.m_year
			 and WML03.m_month = AXX_TMP.m_month 
			 and WML03.m_year=96 AND WML03.m_month=2 
			 AND WML03.bank_code=AXX_TMP.bank_code 
			 AND WML03.report_no=AXX_TMP.report_no
			 */
			if(input_method.equals("F")){//檔案上傳
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select WML03.m_year,WML03.m_month,WML03.bank_code,WML03.report_no,WML03.serial_no,WML03.remark,WML03.user_id,WML03.user_name,WML03.update_date");
				sqlCmd.append(",?,?,sysdate,'D'");
				sqlCmd.append(" from WML03,(select m_year,m_month,bank_code,report_no"); 
				sqlCmd.append("	from AXX_TMP "); 
				sqlCmd.append(" where m_year=? and m_month=?"); 
				sqlCmd.append(" and report_no=?");
				sqlCmd.append(" and filename=?");
				sqlCmd.append(" group by m_year,m_month,bank_code,report_no)AXX_TMP");
				sqlCmd.append(" where WML03.m_year= AXX_TMP.m_year");
				sqlCmd.append(" and WML03.m_month = AXX_TMP.m_month"); 
				sqlCmd.append(" and WML03.m_year=? AND WML03.m_month=?");   
				sqlCmd.append(" AND WML03.bank_code=AXX_TMP.bank_code"); 
				sqlCmd.append(" AND WML03.report_no=AXX_TMP.report_no");
				paramList.add(user_id); 
				paramList.add(user_name);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				paramList.add(filename);
				paramList.add(m_year);
				paramList.add(m_month);
				
				
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		        updateDBList.add(updateDBSqlList);
		        
		        paramList = new ArrayList() ;		       
		        updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		        updateDBDataList = new ArrayList();//儲存參數的List
		        sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("DELETE FROM WML03 WHERE m_year=? AND m_month=?");   
				sqlCmd.append(" AND report_no=?");
				sqlCmd.append(" AND bank_code in (select bank_code"); 
				sqlCmd.append("	from AXX_TMP "); 
				sqlCmd.append(" where m_year=? AND m_month=?");   
				sqlCmd.append(" and report_no=?");
				sqlCmd.append(" group by bank_code)");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);	
			}else{//線上編輯
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select WML03.m_year,WML03.m_month,WML03.bank_code,WML03.report_no,WML03.serial_no,WML03.remark,WML03.user_id,WML03.user_name,WML03.update_date");
				sqlCmd.append(",?,?,sysdate,'D'");
				sqlCmd.append(" from WML03,(select m_year,m_month,bank_code"); 
				sqlCmd.append("	from "+report_no);
				sqlCmd.append(" where m_year=? AND m_month= ?");   
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" group by m_year,m_month,bank_code)t1");
				sqlCmd.append(" where WML03.m_year= t1.m_year");
				sqlCmd.append(" and WML03.m_month = t1.m_month"); 
				sqlCmd.append(" AND WML03.bank_code=t1.bank_code");
				sqlCmd.append(" and WML03.m_year=? AND WML03.m_month=?");  					   
				sqlCmd.append(" AND WML03.report_no=?");
				paramList.add(user_id);   
				paramList.add(user_name);				
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(0));
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);
                

                paramList = new ArrayList() ;              
                updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
                updateDBDataList = new ArrayList();//儲存參數的List
                sqlCmd.delete(0,sqlCmd.length());
							
				sqlCmd.append("DELETE FROM WML03 WHERE m_year=? AND m_month=?");  
				sqlCmd.append(" AND report_no=?");
				sqlCmd.append(" AND bank_code in (select bank_code"); 
				sqlCmd.append("	from "+report_no);
				sqlCmd.append(" where m_year=? and m_month=?");  
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" group by bank_code)");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(0));
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);          	
			}
			updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("Insert WML03_LOG OK??"+updateOK);
			if(!updateOK){
			    //errMsg = errMsg + "Insert WML03_LOG Fail:"+DBManager.getErrMsg()+"<br>";
			    System.out.println("Insert WML03_LOG Fail:"+DBManager.getErrMsg());
			}			
		}catch(Exception e){
			//errMsg = errMsg + "UpdateA01.InsertWML03_LOG Error:"+e.getMessage()+"<br>";
			System.out.println("Utility.InsertWML03_LOG Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertWML03_LOG end time");
    	return updateOK;
	}
	
	//96.07.31 add 批次寫入WML02_LOG(區分檔案上傳/線上編輯)
	//102.10.02 fix 套用preparestatment by 2295
	public static boolean InsertWML02_LOG(String m_year,String m_month,String user_id,String user_name,String report_no,String input_method,String filename,List bank_codeList){
		boolean updateOK=false;
		StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList() ;
        List updateDBList = new ArrayList();//0:sql 1:data
        List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
        List updateDBDataList = new ArrayList();//儲存參數的List
		Utility.printLogTime("InsertWML02_LOG begin time");
			
		try{
			/*
			 select WML02.m_year,WML02.m_month,WML02.bank_code,WML02.report_no,WML02.cano,l_amt,WML02.r_amt,WML02.user_id,WML02.user_name,WML02.update_date
			       ,'A111111111','A111111111',sysdate,'D'
			 from WML02,(select m_year,m_month,bank_code,report_no 
			 		     from AXX_TMP 
			 		     where m_year=96 and m_month=2
			 		     and report_no='A01'
			 		     group by m_year,m_month,bank_code,report_no)AXX_TMP
			 where WML02.m_year= AXX_TMP.m_year
			 and WML02.m_month = AXX_TMP.m_month 
			 and WML02.m_year=96 AND WML02.m_month=2 
			 AND WML02.bank_code=AXX_TMP.bank_code 
			 AND WML02.report_no=AXX_TMP.report_no
			 */			
			if(input_method.equals("F")){//檔案上傳
			   sqlCmd.append(" INSERT INTO WML02_LOG "); 
			   sqlCmd.append(" select WML02.m_year,WML02.m_month,WML02.bank_code,WML02.report_no,WML02.cano,");
			   sqlCmd.append("        WML02.l_amt,WML02.r_amt,WML02.user_id,WML02.user_name,WML02.update_date");
			   sqlCmd.append(",?,?,sysdate,'D'");
			   sqlCmd.append(" from WML02,(select m_year,m_month,bank_code,report_no"); 
			   sqlCmd.append("			   from AXX_TMP"); 
			   sqlCmd.append(" 			   where m_year=? AND m_month=?");  
			   sqlCmd.append(" 			   and report_no=?");
			   sqlCmd.append(" 			   and filename=?");
			   sqlCmd.append("  		   group by m_year,m_month,bank_code,report_no)AXX_TMP");
			   sqlCmd.append(" where WML02.m_year= AXX_TMP.m_year");
			   sqlCmd.append(" and WML02.m_month = AXX_TMP.m_month"); 
			   sqlCmd.append(" and WML02.m_year=? AND WML02.m_month=?");  
			   sqlCmd.append(" AND WML02.bank_code=AXX_TMP.bank_code"); 
			   sqlCmd.append(" AND WML02.report_no=AXX_TMP.report_no");
				paramList.add(user_id);
				paramList.add(user_name);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				paramList.add(filename);
				paramList.add(m_year);
				paramList.add(m_month);
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);                

                paramList = new ArrayList() ;              
                updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
                updateDBDataList = new ArrayList();//儲存參數的List
                sqlCmd.delete(0,sqlCmd.length());		
			    sqlCmd.append("DELETE FROM WML02 WHERE m_year=? AND m_month=?");  
			    sqlCmd.append(" AND report_no=?");
			    sqlCmd.append(" AND bank_code in (select bank_code"); 
			    sqlCmd.append("	from AXX_TMP");
			    sqlCmd.append(" where m_year=? AND m_month=?");  
			    sqlCmd.append(" and report_no=?");
			    sqlCmd.append(" and filename=?");
			    sqlCmd.append(" group by bank_code)");	
			    paramList.add(m_year);
			    paramList.add(m_month);
			    paramList.add(report_no);
			    paramList.add(m_year);
			    paramList.add(m_month);
			    paramList.add(report_no);
			    paramList.add(filename);
			    
                updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);         
			}else{//線上編輯
				sqlCmd.append(" INSERT INTO WML02_LOG "); 
				sqlCmd.append(" select WML02.m_year,WML02.m_month,WML02.bank_code,WML02.report_no,WML02.cano,");
				sqlCmd.append("        WML02.l_amt,WML02.r_amt,WML02.user_id,WML02.user_name,WML02.update_date");
				sqlCmd.append(",?,?,sysdate,'D'");
				sqlCmd.append(" from WML02,(select m_year,m_month,bank_code"); 
				sqlCmd.append(" from "+report_no); 
				sqlCmd.append(" where m_year=? AND m_month=?");   
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" group by m_year,m_month,bank_code)t1");
				sqlCmd.append(" where WML02.m_year= t1.m_year");
				sqlCmd.append(" and WML02.m_month = t1.m_month"); 
				sqlCmd.append(" AND WML02.bank_code=t1.bank_code");
				sqlCmd.append(" and WML02.m_year=? AND WML02.m_month=?");
				sqlCmd.append(" AND WML02.report_no=?");
				paramList.add(user_id);	
				paramList.add(user_name);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(0));
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList);                

                paramList = new ArrayList() ;              
                updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
                updateDBDataList = new ArrayList();//儲存參數的List
                sqlCmd.delete(0,sqlCmd.length());       			
				sqlCmd.append("DELETE FROM WML02 WHERE m_year=? AND m_month=?");  
				sqlCmd.append(" AND report_no=?");
				sqlCmd.append(" AND bank_code in (select bank_code"); 
				sqlCmd.append("	from "+report_no);
				sqlCmd.append(" where m_year=? AND m_month=?");  
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" group by bank_code)");			
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(0));
				updateDBDataList.add(paramList);//1:傳內的參數List
                updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql                  
                updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
                updateDBList.add(updateDBSqlList); 		
			}
			updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("Insert WML02_LOG OK??"+updateOK);
			if(!updateOK){
			    //errMsg = errMsg + "Insert WML02_LOG Fail:"+DBManager.getErrMsg()+"<br>";
			    System.out.println("Insert WML02_LOG Fail:"+DBManager.getErrMsg());
			}
		}catch(Exception e){
			//errMsg = errMsg + "UpdateA01.InsertWML02_LOG Error:"+e.getMessage()+"<br>";
			System.out.println("Utility.InsertWML02_LOG Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertWML02_LOG end time");
    	return updateOK;
	}
	
	//96.07.31 add 批次寫入WML01_LOG
	//102.10.02 add 套用preparestatement by 2295
	public static boolean InsertWML01_LOG(String m_year,String m_month,String user_id,String user_name,String report_no,String input_method,String filename,List bank_codeList){
		boolean updateOK=false;
		String sqlCmd="";	
		List paramList = new ArrayList() ;
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
	    List updateDBDataList = new ArrayList();//儲存參數的List
		Utility.printLogTime("InsertWML01_LOG begin time");
			
		try{
			/*
			select WML01.m_year,WML01.m_month,WML01.bank_code,WML01.report_no,WML01.input_method,WML01.add_user,WML01.add_name,WML01.add_date,WML01.common_center
			       ,WML01.upd_method,WML01.upd_code,WML01.batch_no,WML01.lock_status,WML01.user_id,WML01.user_name,WML01.update_date				 
			       ,'A111111111','A111111111',sysdate,'U'
			 from WML01,(select m_year,m_month,bank_code,report_no 
			 		     from AXX_TMP 
			 		     where m_year=96 and m_month=2
			 		     and report_no='A01'
			 		     group by m_year,m_month,bank_code,report_no)AXX_TMP
			 where WML01.m_year= AXX_TMP.m_year
			 and WML01.m_month = AXX_TMP.m_month 
			 and WML01.m_year=96 AND WML01.m_month=2 
			 AND WML01.bank_code=AXX_TMP.bank_code 
			 AND WML01.report_no=AXX_TMP.report_no
			 */
			if(input_method.equals("F")){//檔案上傳
			   sqlCmd = " INSERT INTO WML01_LOG " 
				      + " select WML01.m_year,WML01.m_month,WML01.bank_code,WML01.report_no,WML01.input_method,WML01.add_user,WML01.add_name,WML01.add_date,WML01.common_center"
			          + "        ,WML01.upd_method,WML01.upd_code,WML01.batch_no,WML01.lock_status,WML01.user_id,WML01.user_name,WML01.update_date"				 
			          + ",?,?,sysdate,'U'"
				      + " from WML01,(select m_year,m_month,bank_code,report_no" 
			 	      + "			   from AXX_TMP " 
			 	      + " 			   where m_year=? AND m_month= ?"  
				      + " 			   and report_no=?"
				      + " 			   and filename=?"
			 	      + "  		   group by m_year,m_month,bank_code,report_no)AXX_TMP"
			 	      + " where WML01.m_year= AXX_TMP.m_year"
				      + " and WML01.m_month = AXX_TMP.m_month" 
				      + " and WML01.m_year=? AND WML01.m_month=?"  
				      + " AND WML01.bank_code=AXX_TMP.bank_code" 
				      + " AND WML01.report_no=AXX_TMP.report_no";
			   paramList.add(user_id);
			   paramList.add(user_name);
			   paramList.add(m_year);
			   paramList.add(m_month);
			   paramList.add(report_no);
			   paramList.add(filename);
			   paramList.add(m_year);
			   paramList.add(m_month);
			}else{//線上編輯
				sqlCmd = " INSERT INTO WML01_LOG " 
					   + " select WML01.m_year,WML01.m_month,WML01.bank_code,WML01.report_no,WML01.input_method,WML01.add_user,WML01.add_name,WML01.add_date,WML01.common_center"
				       + "        ,WML01.upd_method,WML01.upd_code,WML01.batch_no,WML01.lock_status,WML01.user_id,WML01.user_name,WML01.update_date"				 
				       + ",?,?,sysdate,'U'"
					   + " from WML01,(select m_year,m_month,bank_code" 
				 	   + "			   from "+report_no 
				 	   + " 			   where m_year=? AND m_month=?"  
					   + " 			   and bank_code=?"
				 	   + "  		   group by m_year,m_month,bank_code)t1"
				 	   + " where WML01.m_year= t1.m_year"
					   + " and WML01.m_month = t1.m_month" 
					   + " AND WML01.bank_code=t1.bank_code" 
					   + " and WML01.m_year=? AND WML01.m_month=?" 
					   + " AND WML01.report_no=?";
				paramList.add(user_id);
				paramList.add(user_name);
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(0));
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(report_no);
			}
			updateDBDataList.add(paramList);//1:傳內的參數List
            updateDBSqlList.add(sqlCmd);//0:欲執行的sql                  
            updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
            updateDBList.add(updateDBSqlList);      
            updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("Insert WML01_LOG OK??"+updateOK);
			if(!updateOK){
			    //errMsg = errMsg + "Insert WML01_LOG Fail:"+DBManager.getErrMsg()+"<br>";
			    System.out.println("Insert WML01_LOG Fail:"+DBManager.getErrMsg());
			}
		}catch(Exception e){
			//errMsg = errMsg + "UpdateA01.InsertWML01_LOG Error:"+e.getMessage()+"<br>";
			System.out.println("Utility.InsertWML01_LOG Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertWML01_LOG end time");
    	return updateOK;
	}
	
	
    // 96.08.15 add 批次寫入AXX_LOG(使用AXX_TMP/preparestatement)
	// 99.11.10 add a08_log
	// 99.11.12 add a09_log/a10_log
	// 99.11.15 add a05_log
	//102.01.21 add a02_log.amt_name 
	//104.10.08 fix a10_log增加欄位
	//109.06.11 add a15_log
	//110.04.09 add a02_log.amt_name1/amt_name2
	//110.08.18 add a99.amt_name by 2295[111.05.09取消]
	public static boolean InsertAXX_LOG(String m_year,String m_month,String user_id,String user_name,String filename,String report_no){
		String sqlCmd="";
		boolean updateOK=false;
		List A03updateDBSqlList = new LinkedList();
		String condition = "";
    	String acc_code_String = "";
    	Utility.printLogTime("Insert"+report_no+"_LOG begin time");
			
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		try{
			
			/*
			select A99.m_year,A99.m_month,A99.bank_code,A99.acc_code,A99.amt ,'A111111111','A111111111',sysdate,'U' 
			from A99,AXX_TMP
			where A99.m_year=96 and A99.m_month=2
			and  AXX_TMP.m_year = 96 and AXX_TMP.m_month=2
			and A99.bank_code = AXX_TMP.bank_code
			and A99.acc_code = AXX_TMP.acc_code
			order by m_year,m_month,bank_code)    */	    	
			sqlCmd = " INSERT INTO "+report_no+"_LOG " 
				   + " select "+report_no+".m_year,"+report_no+".m_month,"+report_no+".bank_code,";
			if(!report_no.equals("A08") && !report_no.equals("A09") && !report_no.equals("A10")) sqlCmd += report_no+".acc_code,";
			if(report_no.equals("A05")){
			   sqlCmd += " a05.amt,a05.amt_name";	
			}else if(report_no.equals("A06")){
			   sqlCmd += " a06.amt_3month,a06.amt_6month,a06.amt_1year,a06.amt_2year,a06.amt_over2year,a06.amt_total";	
			}else if(report_no.equals("A08")){
			   sqlCmd += " a08.warnaccount_cnt,a08.limitaccount_cnt,a08.erroraccount_cnt,a08.otheraccount_cnt,a08.depositaccount_tcnt";
			}else if(report_no.equals("A09")){
			   sqlCmd += " over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate";
			}else if(report_no.equals("A10")){//104.10.08 fi
			   sqlCmd += " loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt,loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt"
	                  + ",build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt,baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108,property_loss";  
			}else if(report_no.equals("A15")){//109.06.11 add
			   sqlCmd += " a15.month_amt,a15.year_amt,user_id,user_name,update_date";
			}else{
			   sqlCmd += report_no+".amt";
			}
			    
			sqlCmd += ",?,?,sysdate,'U'";
			if(report_no.equals("A02")){//102.06.03 add
	           sqlCmd += ",a02.amt_name,a02.amt_name1,a02.amt_name2";   
	        }
			/*111.05.09取消
			if(report_no.equals("A99")){//110.08.16 add
		       sqlCmd += ",a99.amt_name";   
		    }
		    */
			sqlCmd += " from "+report_no+",AXX_TMP"
			      + " where "+report_no+".m_year=? and "+report_no+".m_month=?"
			      + " and AXX_TMP.m_year=? and AXX_TMP.m_month=?"
			      + " and AXX_TMP.filename=?"
			      + " and "+report_no+".bank_code = AXX_TMP.bank_code";
			if(!report_no.equals("A08") && !report_no.equals("A09") && !report_no.equals("A10")){
				sqlCmd +=" and "+report_no+".acc_code = AXX_TMP.acc_code";
			}
			sqlCmd += " order by m_year,m_month,bank_code";
			updateDBSqlList.add(sqlCmd);
			List dataList = new LinkedList();
			List detail = new LinkedList();
			detail.add(user_id);
			detail.add(user_name);
			detail.add(m_year);
			detail.add(m_month);
			detail.add(m_year);
			detail.add(m_month);
			detail.add(filename);
			dataList.add(detail);
			updateDBSqlList.add(dataList);
			updateDBList.add(updateDBSqlList);      	
			updateOK = DBManager.updateDB_ps(updateDBList);
		    System.out.println("Insert "+report_no+"_LOG OK??"+updateOK);
		    if(!updateOK){			   
			    System.out.println(DBManager.getErrMsg());
		    }			
		}catch(Exception e){			
			System.out.println("Update"+report_no+".Insert"+report_no+"_LOG Error:"+e.getMessage());
		}
		Utility.printLogTime("Insert"+report_no+"_LOG end time");
		return updateOK;
	}
	
	//96.08.15批次寫入zero data
	//99.11.09 add A06 zero data by 2295
	//99.11.11 add A08 zero data by 2295
    //99.11.12 add A09 zero data by 2295
	//99.11.12 add A10 zero data by 2295
	//99.11.15 add A05 zero data by 2295
    //99.11.11 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295 
	//102.02.01 add A02.amt_name by 2295
	//102.04.18 add 漁會A01-103/01以後的資料.套用新的科目代號 by 2295
	//103.02.11 add 漁會A06-103/01以後的資料.套用新的科目代號 by 2295
	//104.05.11 fix 調整bank_code來自於axx_tmp by 2295
	//109.06.11 add A15 zero data by 2295
	//110.04.09 add a02.amt_name1/amt_name2
	//110.08.18 add a99.amt_name by 2295[111.05.09取消]
	public static String InsertZeroAXX_List(String m_year,String m_month,List bank_code,String report_no,String condition){
		StringBuffer sqlCmd = new StringBuffer();	
		boolean updateOK=false;
		String errMsg="";
		String ncacno="ncacno";
		List AXXupdateDBSqlList = new LinkedList();
		Utility.printLogTime("InsertZero"+report_no+"_List begin time");	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		
		try{
			//ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";	
			/*若無資料在A99時,則新增zero至A99
			INSERT INTO A99
			select 96,3,a.bank_no,acc_code ,0 
			from ncacno,(select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			            (select bank_code from A99 where m_year=94 AND m_month=6 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A99已有資料的 
			where acc_tr_type = 'A99' and acc_div='99'
			and a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A99的bank_no
			order by a.bank_no,acc_range"
			
			若無資料在A08時,則新增zero至A08
			INSERT INTO A08
			select 99,8,a.bank_no,0,0,0,0,0 
			from (select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			     (select bank_code from A08 where m_year=99 AND m_month=8 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A08已有資料的 
			where a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A08的bank_no
			
			若無資料在A09時,則新增zero至A09
			INSERT INTO A09
			select 99,8,a.bank_no,0,0,0,0,0,0 
			from (select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			     (select bank_code from A09 where m_year=99 AND m_month=8 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A09已有資料的 
			where a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A09的bank_no
			若無資料在A10時,則新增zero至A10
			INSERT INTO A10
			select 99,8,a.bank_no,0,0,0,0,0,0,0,0,0 
			from (select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			     (select bank_code from A10 where m_year=99 AND m_month=8 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A10已有資料的 
			where a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A10的bank_no
			*/
			//96.12.17 add A01-97/01以後的資料.套用新的科目代號 by 2295
			if(report_no.equals("A01") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 9701)){
		         ncacno = "ncacno_rule";
		    }
			if(report_no.equals("A08")){
				sqlCmd.append(" INSERT INTO "+report_no); 
				sqlCmd.append(" select ?,?,a.bank_no,0,0,0,0,0 ");
				sqlCmd.append(" from (select bank_no from bn01 where bank_no in ("+Utility.getCombinString(bank_code,",")+")and m_year="+((Integer.parseInt(m_year) < 100)?"99":"100")+")a");//目前需更新的bank_code 
				sqlCmd.append("      ,(select bank_code from "+report_no);
				sqlCmd.append("         where m_year=? AND m_month=?");
				sqlCmd.append("           AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A08已有資料的 
				sqlCmd.append(" where a.bank_no = b.bank_code(+)");
				sqlCmd.append(" and b.bank_code is null ");//無資料在A08的bank_no									   
				dataList.add(m_year);
				dataList.add(m_month);
				dataList.add(m_year);
				dataList.add(m_month);
				updateDBDataList.add(dataList);
				
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				updateOK=DBManager.updateDB_ps(updateDBList);
				if(!updateOK){
					//errMsg = errMsg + "InsertZeroA99_List Fail:"+DBManager.getErrMsg()+"<br>";
					System.out.println(DBManager.getErrMsg());
				}
			}else if(report_no.equals("A09")){
					sqlCmd.append(" INSERT INTO "+report_no); 
					sqlCmd.append(" select ?,?,a.bank_no,0,0,0,0,0,0 ");
					sqlCmd.append(" from (select bank_no from bn01 where bank_no in ("+Utility.getCombinString(bank_code,",")+")and m_year="+((Integer.parseInt(m_year) < 100)?"99":"100")+")a");//目前需更新的bank_code 
					sqlCmd.append("      ,(select bank_code from "+report_no);
					sqlCmd.append("         where m_year=? AND m_month=?");
					sqlCmd.append("           AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A09已有資料的 
					sqlCmd.append(" where a.bank_no = b.bank_code(+)");
					sqlCmd.append(" and b.bank_code is null ");//無資料在A08的bank_no									   
					dataList.add(m_year);
					dataList.add(m_month);
					dataList.add(m_year);
					dataList.add(m_month);
					updateDBDataList.add(dataList);
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					updateOK=DBManager.updateDB_ps(updateDBList);
					if(!updateOK){
						//errMsg = errMsg + "InsertZeroA99_List Fail:"+DBManager.getErrMsg()+"<br>";
						System.out.println(DBManager.getErrMsg());
					}
			}else if(report_no.equals("A10")){
				sqlCmd.append(" INSERT INTO "+report_no); 
				sqlCmd.append(" select ?,?,a.bank_no,0,0,0,0,0,0,0,0,0  ");
				sqlCmd.append(" from (select bank_no from bn01 where bank_no in ("+Utility.getCombinString(bank_code,",")+")and m_year="+((Integer.parseInt(m_year) < 100)?"99":"100")+")a");//目前需更新的bank_code 
				sqlCmd.append("      ,(select bank_code from "+report_no);
				sqlCmd.append("         where m_year=? AND m_month=?");
				sqlCmd.append("           AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A09已有資料的 
				sqlCmd.append(" where a.bank_no = b.bank_code(+)");
				sqlCmd.append(" and b.bank_code is null ");//無資料在A08的bank_no									   
				dataList.add(m_year);
				dataList.add(m_month);
				dataList.add(m_year);
				dataList.add(m_month);
				updateDBDataList.add(dataList);
				
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);	
				updateOK=DBManager.updateDB_ps(updateDBList);
				if(!updateOK){
					//errMsg = errMsg + "InsertZeroA99_List Fail:"+DBManager.getErrMsg()+"<br>";
					System.out.println(DBManager.getErrMsg());
				}
			}else{//A01~A06/A15
			    //農會
			    sqlCmd.append(" INSERT INTO "+report_no); 
			    sqlCmd.append(" select ?,?,a.bank_no,acc_code ,0");
			    if(report_no.equals("A05")){//110.08.18 add A99.amt_name[111.05.09取消]	    	
				   sqlCmd.append(",''");//amt_name
			    }else if(report_no.equals("A02")){	//102.02.01 add A02.amt_name //110.04.09 add amt_name1/amt_name2		    	
				   sqlCmd.append(",'','',''");//amt_name/amt_name1/amt_name2
			    }else if(report_no.equals("A06")){
			       sqlCmd.append(",0,0,0,0,0");
			    //}else if(report_no.equals("A99")){	//110.08.17 add A99.amt_name[111.05.09取消] 		    	
				//   sqlCmd.append(",''");//amt_name  
			    }else if(report_no.equals("A15")){//109.06.11 add
				   sqlCmd.append(",0,'','',null");      
			    }
			    sqlCmd.append(" from "+ncacno+",(select bank_no from bn01 where bank_no in (select bank_code from axx_tmp where m_year=? and m_month=? and report_no=? )and bank_type='6' and m_year="+((Integer.parseInt(m_year) < 100)?"99":"100")+")a");//目前需更新的bank_code 
			    sqlCmd.append("            ,(select bank_code from "+report_no);
			    sqlCmd.append("              where m_year=? AND m_month=?");
			    sqlCmd.append("              AND bank_code in (select bank_code from axx_tmp where m_year=? and m_month=? and report_no=?)group by bank_code)b");//A99已有資料的 
			    sqlCmd.append(" where "+condition);
			    sqlCmd.append(" and a.bank_no = b.bank_code(+)");
			    sqlCmd.append(" and b.bank_code is null ");//無資料在A99的bank_no
			    sqlCmd.append(" order by a.bank_no,acc_range");		
			                    
			    dataList.add(m_year);
			    dataList.add(m_month);
			    
			    dataList.add(m_year);//104.05.11 add
                dataList.add(m_month);//104.05.11 add
                dataList.add(report_no);//104.05.11 add
			    
                dataList.add(m_year);
                dataList.add(m_month);
                
			    dataList.add(m_year);//104.05.11 add
                dataList.add(m_month);//104.05.11 add
                dataList.add(report_no);//104.05.11 add
			    
			   
			    updateDBDataList.add(dataList);
			    
			    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
			    updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			    updateDBList.add(updateDBSqlList);
			    updateOK=DBManager.updateDB_ps(updateDBList);
				if(!updateOK){
					//errMsg = errMsg + "InsertZero"+report_no +"_List_6 Fail:"+DBManager.getErrMsg()+"<br>";
					System.out.println("InsertZero"+report_no +"_List_6 Fail:"+DBManager.getErrMsg());
				}
			    //updateOK=DBManager.updateDB(AXXupdateDBSqlList);
			    //System.out.println("InsertZero"+report_no +"_List_6 OK??"+updateOK);
			    //if(!updateOK){				
			    //	System.out.println(DBManager.getErrMsg());
			    //}
			    //AXXupdateDBSqlList = new LinkedList();
			    //漁會
			    updateDBSqlList = new ArrayList();		
			    updateDBDataList = new ArrayList();//儲存參數的List		
			    updateDBList = new LinkedList();//0:sql 1:data		
			    sqlCmd.delete(0, sqlCmd.length());
			    sqlCmd.append(" INSERT INTO "+report_no); 
			    sqlCmd.append(" select ?,?,a.bank_no,acc_code ,0");
			    if(report_no.equals("A05")){//110.08.18 add A99.amt_name[111.05.09取消]		    	
				   sqlCmd.append(",''");//amt_name
			    }else  if(report_no.equals("A02")){	//102.02.01 add A02.amt_name //110.04.09 add amt_name1/amt_name2			    	
					sqlCmd.append(",'','',''");//amt_name/amt_name1/amt_name2   
			    }else if(report_no.equals("A06")){
			       sqlCmd.append(",0,0,0,0,0");
			    //}else if(report_no.equals("A99")){	//110.08.17 add A99.amt_name [111.05.09取消]			    	
				//   sqlCmd.append(",''");//amt_name
			    }else if(report_no.equals("A15")){//109.06.11 add
				   sqlCmd.append(",0,'','',null");      
				}
			    //102.04.18 add 漁會A01-103/01以後的資料.套用新的科目代號 by 2295
	            if(report_no.equals("A01") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301)){
	               ncacno = "ncacno_7_rule";
	            }else if(report_no.equals("A06") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301)){
	               ncacno = "ncacno_7_rule"; //103.02.11 add 漁會A06-103/01以後的資料.套用新的科目代號 by 2295
	            }else{
	               ncacno = "ncacno_7";
	            }
			    sqlCmd.append(" from "+ncacno+",(select bank_no from bn01 where bank_no in (select bank_code from axx_tmp where m_year=? and m_month=? and report_no=?)and bank_type='7' and m_year="+((Integer.parseInt(m_year) < 100)?"99":"100")+")a");//目前需更新的bank_code 
			    sqlCmd.append("            ,(select bank_code from "+report_no);
			    sqlCmd.append("              where m_year=? AND m_month=?");
			    sqlCmd.append("              AND bank_code in (select bank_code from axx_tmp where m_year=? and m_month=? and report_no=?)group by bank_code)b");//A99已有資料的 
			    sqlCmd.append(" where "+condition);
			    sqlCmd.append(" and a.bank_no = b.bank_code(+)");
			    sqlCmd.append(" and b.bank_code is null ");//無資料在A99的bank_no
			    sqlCmd.append(" order by a.bank_no,acc_range");	
			    updateDBDataList.add(dataList);			
			    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
			    updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			    updateDBList.add(updateDBSqlList);
			    updateOK=DBManager.updateDB_ps(updateDBList);
				if(!updateOK){
					//errMsg = errMsg + "InsertZeroA99_List Fail:"+DBManager.getErrMsg()+"<br>";
					System.out.println("InsertZero"+report_no +"_List_7 Fail:"+DBManager.getErrMsg());
				}
			    //AXXupdateDBSqlList.add(sqlCmd);
			    //updateOK=DBManager.updateDB(AXXupdateDBSqlList);
			    //System.out.println("InsertZero"+report_no +"_List_7 OK??"+updateOK);
			}
			
		}catch(Exception e){
			errMsg = errMsg + "Update"+report_no +".InsertZero"+report_no +"_List Error:"+e.getMessage()+"<br>";
			System.out.println("Update"+report_no +".InsertZero"+report_no +"_List Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertZero"+report_no+"_List end time");	
		return errMsg;
	}
	
	// 96.08.15 add 批次寫入AXX(將有異動的updte至AXX)--使用preparestatement	 
	// 96.11.23 add 增加寫入AMT_NAME 
	// 99.11.09 add 增加寫入A06.amt_3month(amt),amt_6month,amt_1year,amt_2year,amt_over2year,amt_total
	// 99.11.10 add 增加寫入A08
    // 99.11.12 add 增加寫入A09
	//102.04.25 add A02.amt_name by 2295
	//109.06.10 add 增加寫入A15
	//110.08.18 add A99.amt_name by 2295[111.05.09取消]
	public static String updateAXX(String m_year,String m_month,String report_no,String filename){
		boolean updateOK=false;		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		String errMsg="";
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List		
		List dataList = new LinkedList();
		DataObject obj = null;
		
		Utility.printLogTime("update"+report_no+" begin time");	
		try{
			//有其他檢核錯誤的(wml03).不更新至A99
			//96.05.01 使用preparestatment
			sqlCmd.append(" select bank_code,acc_code,amt,amt1,amt2,amt3,amt4,amt5,amt6,amt7,amt8,amt9,amt_name ");
			sqlCmd.append(" from axx_tmp");
			sqlCmd.append(" where m_year=? AND m_month=?");
			sqlCmd.append(" and report_no=? ");
			sqlCmd.append(" and bank_code not in (select bank_code from wml03 where m_year=? AND m_month=?");
			sqlCmd.append("                       and report_no=? group by bank_code)");
			sqlCmd.append(" and filename=?");
			sqlCmd.append(" order by bank_code,acc_code");
			paramList.add(m_year);//傳內的參數List
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(filename);
				
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt,amt1,amt2,amt3,amt4,amt5,amt6,amt7,amt8,amt9");			
			
			sqlCmd.delete(0,sqlCmd.length());
			
			if(report_no.equals("A05") || report_no.equals("A02")){//110.08.17 add amt_name[111.05.09A99取消]	
				sqlCmd.append(" UPDATE "+report_no+" SET amt= ? ");
				sqlCmd.append(" ,amt_name = ? ");
			}else if(report_no.equals("A06")){
				sqlCmd.append(" UPDATE "+report_no+" SET amt_3month = ? ");
				sqlCmd.append(" ,amt_6month = ? ");
				sqlCmd.append(" ,amt_1year = ? ");
				sqlCmd.append(" ,amt_2year = ? ");
				sqlCmd.append(" ,amt_over2year = ?");
				sqlCmd.append(" ,amt_total = ?");
			}else if(report_no.equals("A08")){
				sqlCmd.append(" UPDATE "+report_no+" SET warnaccount_cnt = ? ");
				sqlCmd.append(" ,limitaccount_cnt = ? ");
				sqlCmd.append(" ,erroraccount_cnt = ? ");
				sqlCmd.append(" ,otheraccount_cnt = ? ");
				sqlCmd.append(" ,depositaccount_tcnt = ?");
			}else if(report_no.equals("A09")){
				sqlCmd.append(" UPDATE "+report_no+" SET over_cnt = ? ");
				sqlCmd.append(" ,over_amt = ? ");
				sqlCmd.append(" ,push_over_amt = ? ");
				sqlCmd.append(" ,totalamt = ? ");
				sqlCmd.append(" ,push_totalamt = ?");	
				sqlCmd.append(" ,over_total_rate = ?");
			}else if(report_no.equals("A10")){
				sqlCmd.append(" UPDATE "+report_no+" SET loan2_amt = ? ");
				sqlCmd.append(" ,loan3_amt = ? ");
				sqlCmd.append(" ,loan4_amt = ? ");
				sqlCmd.append(" ,invest2_amt = ? ");
				sqlCmd.append(" ,invest3_amt = ?");	
				sqlCmd.append(" ,invest4_amt = ?");	
				sqlCmd.append(" ,other2_amt = ? ");
				sqlCmd.append(" ,other3_amt = ?");	
				sqlCmd.append(" ,other4_amt = ?");
			}else if(report_no.equals("A15")){//109.06.10 add
				sqlCmd.append(" UPDATE "+report_no+" SET month_amt = ? ");
				sqlCmd.append(" ,year_amt = ? ");				
			}else{//A01~A04
				sqlCmd.append(" UPDATE "+report_no+" SET amt= ? ");
			}
			
			sqlCmd.append(" where m_year=? and m_month=?");
			sqlCmd.append(" and bank_code=? ");
			if(!report_no.equals("A08") && !report_no.equals("A09") && !report_no.equals("A10")) sqlCmd.append(" and acc_code =? ");
			updateDBSqlList.add(sqlCmd.toString());
			//List A99List = new LinkedList();
			//List detail = new LinkedList();
			if(dbData != null && dbData.size() > 0){
				for(int i=0;i<dbData.size();i++){
				  obj = (DataObject) dbData.get(i);
				  dataList=new LinkedList();
				  dataList.add(obj.getValue("amt").toString());//A01~A05	
				  if(report_no.equals("A05") || report_no.equals("A02")){//A05.A02.A99 //110.08.17 add A99[111.05.09取消]	
				  	dataList.add(obj.getValue("amt_name") == null?"":(String)obj.getValue("amt_name"));
				  }else if (report_no.equals("A06")){//A06
				  	dataList.add(obj.getValue("amt1").toString());	
				  	dataList.add(obj.getValue("amt2").toString());
				  	dataList.add(obj.getValue("amt3").toString());
				  	dataList.add(obj.getValue("amt4").toString());
				  	dataList.add(obj.getValue("amt5").toString());
				  }else if (report_no.equals("A08")){//A08
				  	dataList.add(obj.getValue("amt1").toString());	
				  	dataList.add(obj.getValue("amt2").toString());
				  	dataList.add(obj.getValue("amt3").toString());
				  	dataList.add(obj.getValue("amt4").toString());				  	
				  }else if (report_no.equals("A09")){//A09
				  	dataList.add(obj.getValue("amt1").toString());	
				  	dataList.add(obj.getValue("amt2").toString());
				  	dataList.add(obj.getValue("amt3").toString());
				  	dataList.add(obj.getValue("amt4").toString());	
				  	dataList.add(obj.getValue("amt5").toString());	
				  }else if (report_no.equals("A10")){//A10
				  	dataList.add(obj.getValue("amt1").toString());	
				  	dataList.add(obj.getValue("amt2").toString());
				  	dataList.add(obj.getValue("amt3").toString());
				  	dataList.add(obj.getValue("amt4").toString());	
				  	dataList.add(obj.getValue("amt5").toString());		
				  	dataList.add(obj.getValue("amt6").toString());	
				  	dataList.add(obj.getValue("amt7").toString());	
				  	dataList.add(obj.getValue("amt8").toString());	
				  }else if (report_no.equals("A15")){//A15
					dataList.add(obj.getValue("amt1").toString());		
				  }
				  dataList.add(m_year);
				  dataList.add(m_month);
				  dataList.add((String)obj.getValue("bank_code"));
				  if(!report_no.equals("A08") && !report_no.equals("A09") && !report_no.equals("A10")) dataList.add((String)obj.getValue("acc_code"));
				  //System.out.println(detail);
				  updateDBDataList.add(dataList);//1:傳內的參數List
				}
			}	
			
			
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List			
			updateDBList.add(updateDBSqlList);
			
		    if(DBManager.updateDB_ps(updateDBList)){
        	   System.out.println(report_no+" UPDATE ok");			   	
        	}			
        	
		}catch(Exception e){
			errMsg = errMsg + "Update"+report_no+".update"+report_no+" OK Error:"+e.getMessage()+"<br>";
			System.out.println("Update"+report_no+".update"+report_no+" OK Error:"+e.getMessage());
		}
		Utility.printLogTime("update"+report_no+" end time");
		
    	return errMsg;
	}
	
	//96.08.15 add 批次寫入WML03(檢核其他錯誤.科目代號不存在)
	//102.10.02 add 套用preparestatement by 2295
	//103.02.11 add A06.漁會套用新表格(增加/異動科目代號)
	public static String InsertWML03(String m_year,String m_month,String user_id,String user_name,String report_no,String filename){
		boolean updateOK=false;
		String ncacno = "ncacno";
		String ncacno_7 = "ncacno_7";
		String errMsg="";
		String sqlCmd="";		
		Utility.printLogTime("InsertWML03 begin time");			
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List paramList = new ArrayList();
		
		//103.02.11 add A06.漁會套用新表格(增加/異動科目代號)
		if( report_no.equals("A06") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301) ){		    
		    ncacno_7 = "ncacno_7_rule";
        }
		try{
			/* A99.科目代號區分農漁會
			select * from
			(
				select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,ncacno.acc_name,amt
				from (
					  select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt
					  from axx_tmp,bn01	 
					  where m_year=96 and m_month=2 
					  and report_no='A99'
					  and axx_tmp.bank_code=bn01.BANK_NO 
					  and bn01.bank_type='6'
					  --and bank_code in ('6190196','6200167')
					  )axx_tmp,(select acc_code,acc_name from ncacno where acc_tr_type='A99')ncacno
					  where axx_tmp.acc_code = ncacno.acc_code(+))a
					  where acc_name is null
			union 
			select * from
			(
				select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,ncacno_7.acc_name,amt
				from (
					  select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt
					  from axx_tmp,bn01	 
					  where m_year=96 and m_month=2 
					  and report_no='A99'
					  and axx_tmp.bank_code=bn01.BANK_NO 
					  and bn01.bank_type='7'
					  --and bank_code in ('6190196','6200167')
					  )axx_tmp,(select acc_code,acc_name from ncacno_7 where acc_tr_type='A99')ncacno_7
					  where axx_tmp.acc_code = ncacno_7.acc_code(+))a
					  where acc_name is null		  
			 
			 */
			
			
			//檢核其他錯誤
			sqlCmd = " select * from(" 
				   + "        select axx_tmp.m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,"+ncacno+".acc_name,amt"
				   + "		  from (select axx_tmp.m_year,m_month,bank_code,bank_type,report_no,acc_code,amt"
				   + "				from axx_tmp,(select * from bn01 where m_year=?)bn01" 
				   + " 			    where axx_tmp.m_year=? AND m_month=?"  
				   + " 			    and report_no=?"	
				   + "              and filename=?"
				   + "		        and axx_tmp.bank_code=bn01.BANK_NO and bn01.bank_type='6')axx_tmp," 
				   + " 			    (select acc_code,acc_name from "+ncacno+" where acc_tr_type=?)"+ncacno
				   + "		        where axx_tmp.acc_code = "+ncacno+".acc_code(+)"
				   + "             )a"				 
				   + " where acc_name is null"
				   + " union "
				   + " select * from(" 
				   + "        select axx_tmp.m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,"+ncacno_7+".acc_name,amt"
				   + "		  from (select axx_tmp.m_year,m_month,bank_code,bank_type,report_no,acc_code,amt"
				   + "				from axx_tmp,(select * from bn01 where m_year=?)bn01" 
				   + " 			    where axx_tmp.m_year=? AND m_month=?"  
				   + " 			    and report_no=?"	
				   + "              and filename=?"
				   + "		        and axx_tmp.bank_code=bn01.BANK_NO and bn01.bank_type='7')axx_tmp," 
				   + " 			    (select acc_code,acc_name from "+ncacno_7+" where acc_tr_type=?)"+ncacno_7
				   + "		        where axx_tmp.acc_code = "+ncacno_7+".acc_code(+)"
				   + "				)a"				 
				   + " where acc_name is null";
			paramList.add(((Integer.parseInt(m_year) < 100)?"99":"100"));
			paramList.add(m_year);	  	
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(filename);
	        paramList.add(report_no);
			paramList.add(((Integer.parseInt(m_year) < 100)?"99":"100"));
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(filename);
			paramList.add(report_no);
			
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,amt");
			int errCount=-1;
			DataObject obj = null;
			//無此科目代號
			sqlCmd = "INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)"; 
			String prebank_code="";			
			if(dbData != null && dbData.size() != 0){
			   prebank_code=(String)((DataObject) dbData.get(0)).getValue("bank_code");
			   for(int i=0;i<dbData.size();i++){
				   obj = (DataObject) dbData.get(i);
				   if(prebank_code.equals((String)obj.getValue("bank_code"))){
					  errCount++;
				   }else{
				      prebank_code = (String)obj.getValue("bank_code");
					  errCount=0;
				   }
				   //無此科目代號
				   paramList = new ArrayList();
				   paramList.add(obj.getValue("m_year").toString());
				   paramList.add(obj.getValue("m_month").toString());
				   paramList.add((String)obj.getValue("bank_code"));
				   paramList.add(report_no);
				   paramList.add(String.valueOf(errCount));
				   paramList.add("科目代號[" + (String)obj.getValue("acc_code") + "]不存在");
				   paramList.add(user_id);
				   paramList.add(user_name);
				   updateDBDataList.add(paramList);
			   }//end of for
			   updateDBSqlList.add(sqlCmd);//0:欲執行的sql                 
               updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
               updateDBList.add(updateDBSqlList);
               updateOK=DBManager.updateDB_ps(updateDBList);
               
			   System.out.println("Insert WML03 OK??"+updateOK);
			   if(!updateOK){
			      errMsg = errMsg + "Insert WML03 Fail:"+DBManager.getErrMsg()+"<br>";
			      System.out.println("Insert WML03 Fail:"+DBManager.getErrMsg());
			   }
			}//end of 有其他檢核錯誤時
		}catch(Exception e){
			errMsg = errMsg + "Update"+report_no+".InsertWML03 Error:"+e.getMessage()+"<br>";
			System.out.println("Update"+report_no+".InsertWML03 Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertWML03 end time");		
		return errMsg;
	}
	//96.12.18 add 檢核完成後.刪除暫存資料 by 2295
	public static String deleteAXX_TMP(String m_year,String m_month,String report_no,String filename){		
	    String errMsg="";
	    List paramList = new ArrayList();
	    Utility.printLogTime("deleteAXX_TMP begin time");
		try{
		    paramList.add(m_year);
		    paramList.add(m_month);
		    paramList.add(filename);
			List dbData = DBManager.QueryDB_SQLParam("select count(*) as countdata from AXX_TMP where m_year=? and m_month=? and filename=?",paramList,"countdata");
			if(dbData != null && (Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) != 0)){
				System.out.println("AXX_TMP.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				List deleteList = new LinkedList();
				deleteList.add("delete AXX_TMP where m_year="+m_year+" and m_month="+m_month+" and filename='"+filename+"'");
				System.out.println("delete OK??"+DBManager.updateDB(deleteList));	
			}	
		}catch(Exception e){
			errMsg = errMsg + "delteAXX_TMP("+report_no+") Error:"+e.getMessage()+"<br>";
			System.out.println("delteAXX_TMP("+report_no+") Error Error:"+e.getMessage());
		}
		Utility.printLogTime("delteAXX_TMP end time");		
		return errMsg;
	}
    //從WML01_UPLOAD取得對應的帳號.姓名
	public static String[] getUser_Data(String filename){
		String user_data[]={"",""};
		String sqlCmd="";
		List paramList = new ArrayList();
		try{
			sqlCmd = "select * from WML01_UPLOAD where filename=?";
			paramList.add(filename.trim());
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			if(dbData.size() != 0){
			   user_data[0] = (String)((DataObject)dbData.get(0)).getValue("muser_id");
			   user_data[1] = (String)((DataObject)dbData.get(0)).getValue("muser_name");
			}
		}catch(Exception e){			
			System.out.println("getUser_Data Error:"+e.getMessage());
		}
		return user_data;	
	}
	
	public static void printLogTime(String report_no){
			Calendar logcalendar = Calendar.getInstance();
			nowlog = logcalendar.getTime();			
			System.out.println(report_no+"="+logformat.format(nowlog));	
	}
	
	
    
	//保存查詢條件(有複選的)，會將複選的值，以String[]形式存放
    public static Map saveSearchParameter(ServletRequest request, boolean hasMultiple) throws Exception {

        Map h = null;
        if (hasMultiple) {
            Map tmpMap = request.getParameterMap();
            h = new HashMap();
            Set s = tmpMap.keySet();
            Iterator it = s.iterator();
            while (it.hasNext()) {
                String key = (String) it.next();
                if ("action".equals(key)) {
                    continue;
                }
                Object value = tmpMap.get(key);
                h.put(key, value);
                // logger.debug("key : " + key);
                if (value instanceof String[]) {
                    String[] tmpValue = (String[]) value;
                    for (int i = 0; i < tmpValue.length; i++) {
                        // logger.debug("tmpValue : " + tmpValue[i]);
                    }
                } else {
                    // logger.debug("value : " + value);
                }
            }
        } else {
            Enumeration enu = request.getParameterNames();
            h = new HashMap();
            while (enu.hasMoreElements()) {
                String name = (String) enu.nextElement();
                String value = Utility.getTrimString(request.getParameter(name));
                if ("action".equals(name)) {
                    continue;
                }
                System.out.println(" Save Request Attribute   :  " + name + " = " + value);                
                h.put(name, value);
            }
        }
        return h;
    }
    //99.02.09取得機構類別
    public static List getBank_Kind(String cmuser_div,String cmuse_id) throws Exception {    	
    	List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
		
		sql.append(" select cmuse_id,cmuse_name from cdshareno where cmuse_div=? and cmuse_id not in (?) order by input_order"); 
		paramList.add(cmuser_div);
		paramList.add(cmuse_id);
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");                    
        return dbData;        
    }
    //99.02.10取得BN01該總機構代號資料
    //99.10.15區分100年/99年總機基本資料
    public static List getBN01(String bank_no) throws Exception{
    	List paramList = new ArrayList();
    	StringBuffer sql = new StringBuffer();
    	Calendar now = Calendar.getInstance();
	    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    //99.10.15 add 查詢年度100年以前.縣市別不同===============================         
  	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
        //===================================================================== 
    	sql.append(" select * from BN01 where m_year=? and bank_no=?");
    	paramList.add(wlx01_m_year);
    	paramList.add(bank_no);
    	List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"update_date");                    
        return dbData;
    } 
    	
    //99.02.10 add 取得是否有該程式細部功能權限
    public static boolean getPermission(HttpServletRequest request,String report_no,String func){
    	boolean CheckOK=false;
	    HttpSession session = request.getSession();            
        Properties permission = ( session.getAttribute(report_no)==null ) ? new Properties() : (Properties)session.getAttribute(report_no);				                
        System.out.println("report_no="+report_no);
        System.out.println("func="+func);
        if(permission != null && permission.get(func) != null && permission.get(func).equals("Y")){    	            
    	   CheckOK = true;
    	}        	
    	return CheckOK;
   }
  
    //99.03.10 add 取得所有總機構資料
    /****************************************************************************
     * 取得所有總機構資料      
     */
    public static List getALLTBank(String bank_type){
    		List paramList = new ArrayList();
    		StringBuffer sql = new StringBuffer();
    		sql.append(" select bn01.bn_type,HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE,bn01.m_year  ");
    		sql.append(" from  BN01, WLX01 ");
    		sql.append(" where BN01.BANK_NO = WLX01.BANK_NO(+) ");
    		sql.append(" and bank_type in ( ? ) ");
    		sql.append(" and wlx01.m_year = bn01.m_year ");    		
    		sql.append(" order by BANK_NO  ");    
    		
    		paramList.add(bank_type);
    		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"m_year");                    
    		return dbData;    		
    }
    //99.03.10 取得所有縣市(99年跟100年縣市別)
    /****************************************************************************
     * 取得所有縣市(99年跟100年縣市別)  
     * 99.08.25 以報表的排序為準    
     */
    public static List getCity(){
    		List paramList = new ArrayList();
    		StringBuffer sql = new StringBuffer();
    		sql.append(" select HSIEN_id, HSIEN_name,100 as m_year,fr001w_output_order from cd01 ");
    		sql.append(" union ");
    		sql.append(" select HSIEN_id, HSIEN_name,99 as m_year,fr001w_output_order from cd01_99 ");
    		sql.append(" order by m_year,fr001w_output_order, hsien_id ");
    		
    		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),null,"m_year");                    
    		return dbData;    		
    }
    
    //99.03.10 取得檢核結果與最後異動日期 	
    /****************************************************************************
     * 取得檢核結果與最後異動日期      
     */
    public static List getWML01UPD_CODE(String s_year,String s_month,String bank_code,String report_no){
    		List paramList = new ArrayList();
    		StringBuffer sql = new StringBuffer();
    		sql.append(" select UPD_CODE,  decode(UPD_CODE,'N','待檢核','E','檢核錯誤','U','檢核成功','') as UPD_CODE_NAME,");
    		sql.append(" to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE");
    		sql.append(" from WML01");		   		  
    		sql.append(" where M_YEAR=?");		   		   
    		sql.append("  and M_MONTH=?");
    		sql.append("  and BANK_CODE=?");
    		sql.append("  and REPORT_NO=?");
    		paramList.add(s_year);
    		paramList.add(s_month);
    		paramList.add(bank_code);
    		paramList.add(report_no);
    		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");                    
    		return dbData;    		
    }   
    //99.10.05 add 取得科目代號/名稱  
    public static List getAcc_Code(String ncacno,String acc_tr_type,String acc_div){
    	List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
		sql.append("select * from "+ncacno+" where acc_tr_type = ? and acc_div=? order by acc_range");		
		paramList.add(acc_tr_type);
		paramList.add(acc_div);		
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"update_date");                    
		return dbData;    		
    }
    //99.10.20 add 取得程式名稱
    public static String getPgName(String report_no){
    	List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
		String pgname = "";
		sql.append("select * from wtt03_1 where program_id = ?");		
		paramList.add(report_no);
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
		if(dbData.size() > 0){
		   pgname = (String)((DataObject)dbData.get(0)).getValue("program_name");
		}
		return pgname;        	
    }
    //99.11.18 add 寫入WML03的參數
    public static List create_dataList(String m_year,String m_month,String bank_code,String report_no,
			int errCount,String subErrMsg, String user_id,String user_name){
		
			List dataList =  new ArrayList();//儲存參數的data
			dataList.add(String.valueOf(Integer.parseInt(m_year)));
			dataList.add(String.valueOf(Integer.parseInt(m_month)));
			dataList.add(bank_code);
			dataList.add(report_no);
			dataList.add(String.valueOf(errCount));
			dataList.add(subErrMsg);
			dataList.add(user_id);
			dataList.add(user_name);
			return dataList;
	}
    //99.11.19 add 讀取WML03.count(*)--countdata
    public static List getWML03_count(String m_year,String m_month,String bank_code,String report_no){
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List dbData = null;
    	sqlCmd.delete(0,sqlCmd.length());
		sqlCmd.append("select count(*) as countdata from WML03 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
		paramList = new ArrayList();
		paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(bank_code);
		paramList.add(report_no);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");
		return dbData;
	}
    
    //99.11.19 add 讀取該申報資料all data 都為0的資料筆數
    public static List getCountZero(String m_year,String m_month,String bank_code,String report_no,String condition){
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List dbData = null;
    	sqlCmd.delete(0,sqlCmd.length());
    	sqlCmd.append("select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=?");
    	if(!condition.equals("")) sqlCmd.append(" and "+condition);
	 	paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(bank_code);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
		return dbData;
	}
    //99.11.19 add 讀取WML01 all data
    public static List getWML01(String m_year,String m_month,String bank_code,String report_no){
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List dbData = null;
    	sqlCmd.delete(0,sqlCmd.length());
    	sqlCmd.append("select * from WML01 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
	 	paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(bank_code);
		paramList.add(report_no);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
		return dbData;
	}
    //99.11.19 add wml01不存在Insert,存在時Update    
    public static List Insert_UpdateWML01(List dbData,String m_year,String m_month,String bank_code,String report_no,String input_method,String add_date,String user_id,String user_name,String common_center,String upd_method,String upd_code,int batch_no){
    	StringBuffer sqlCmd = new StringBuffer();
    	List updateDBDataList = new ArrayList();
		List updateDBSqlList = new ArrayList();
		List updateDBList = new ArrayList();
		List dataList = new ArrayList();//傳內的參數List		
		if(dbData.size() == 0){//WML01不存在時,Insert
			if(input_method.equals("F")){//檔案上傳
				//dataList_InsertWML01.add("to_date('"+add_date+"','YYYYMMDDHH24MISS')");
				sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
				
			    //99.09.24sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"++"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
			    //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
			}else{//線上編輯
				//dataList_InsertWML01.add("sysdate");						
				sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");						
			    //99.09.24sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
		        //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
			}
					
			dataList = new ArrayList();//傳內的參數List
			dataList.add(m_year); 
			dataList.add(m_month); 
			dataList.add(bank_code); 
			dataList.add(report_no); 
			dataList.add(input_method);
			dataList.add(user_id);
			dataList.add(user_name);
			dataList.add(common_center);
			dataList.add(upd_method);
			dataList.add(upd_code);
			dataList.add(String.valueOf(batch_no));
			dataList.add(user_id);
			dataList.add(user_name);
			
			updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List			
			updateDBList.add(updateDBSqlList);
		}else{//WML01存在時,做update
			/*96.04.26移至批次寫入A01~A99InsertWML01_LOG*/
			if(report_no.equals("F01") 
			|| report_no.equals("B01") || report_no.equals("B02") || report_no.equals("B03")
			|| report_no.equals("M01") || report_no.equals("M02") || report_no.equals("M03") || report_no.equals("M04")					
			|| report_no.equals("M05") || report_no.equals("M06") || report_no.equals("M07") || report_no.equals("M08")
			|| report_no.equals("M106") || report_no.equals("M201") || report_no.equals("M206") 
			){
			sqlCmd.append(" INSERT INTO WML01_LOG "); 
			sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,");
			sqlCmd.append(" upd_method,upd_code,batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'");
			sqlCmd.append(" from WML01");
			sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			dataList = new ArrayList();//傳內的參數List
			dataList.add(user_id); 
			dataList.add(user_name); 
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(bank_code);
			dataList.add(report_no);
			updateDBDataList.add(dataList);//1:傳內的參數List	
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			
		    sqlCmd.delete(0,sqlCmd.length());
			updateDBDataList = new ArrayList();
			updateDBSqlList = new ArrayList();
			}
		    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;			
			sqlCmd.append(" UPDATE WML01 SET "); 
			sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
			sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
					
			//99.09.27sqlCmd=" UPDATE WML01 SET " 
			//	  +" upd_method='"+upd_method+"',upd_code='"+upd_code+"',batch_no="+batch_no+",user_id='"+user_id+"',user_name='"+user_name+"',update_date=sysdate"   
			//	  +" where m_year=" + m_year + " AND m_month=" + m_month + " AND " 						
			//	  +" bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
				
			dataList = new ArrayList();//傳內的參數List
			dataList.add(upd_method); 
			dataList.add(upd_code); 
			dataList.add(String.valueOf(batch_no));
			dataList.add(user_id);
			dataList.add(user_name);
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(bank_code);
			dataList.add(report_no);					
			updateDBDataList.add(dataList);//1:傳內的參數List			
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
		}			
		return updateDBList;
	}
    //99.11.19 add 刪除上傳檔案紀錄
    public static List deleteWML01_UPLOAD(String filename){
    	StringBuffer sqlCmd = new StringBuffer();
    	List updateDBDataList = new ArrayList();
		List updateDBSqlList = new ArrayList();
		List dataList = new ArrayList();//傳內的參數List		
		sqlCmd.append(" DELETE FROM WML01_UPLOAD where filename=?"); 
		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
		dataList.add(filename);					
		updateDBDataList.add(dataList);//1:傳內的參數List
		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		//updateDBList.add(updateDBSqlList);						
		return updateDBSqlList;
	}
    //寫入操作log
    public static String insertDataToLog(HttpServletRequest request, String loginID, String pgId) {
        String actMsg = "";
        String errMsg = "";
        String act = Utility.getTrimString(request.getParameter("act" ));
        String sn_docno = Utility.getTrimString(request.getParameter("sn_docno" ));
        String rt_docno = Utility.getTrimString(request.getParameter("rt_docno" ));
        String reportno = Utility.getTrimString(request.getParameter("reportno" ));
        String reportno_seq = Utility.getTrimString(request.getParameter("reportno_seq" ));
        String sn_date = Utility.getTrimString(request.getParameter("sn_date" ));
        
        List paramList = new ArrayList() ;
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
        try {
            String ip = request.getRemoteAddr();
            System.out.println("ip_address:"+ip);
            StringBuffer sqlCmd = new StringBuffer() ;
            sqlCmd.append( "insert into EXOPERATE_LOG(muser_id, use_date, program_id, ip_address, sn_docno, rt_docno, reportno,reportno_seq, update_type) " +
                                                " values(?, sysdate, ?, ?, ?, ?, ?, ?, ?) " );
            paramList.add(loginID) ;
            paramList.add(pgId) ;
            paramList.add(ip) ;
            System.out.println("# request :"+sn_docno+", "+rt_docno+", "+reportno+", "+reportno_seq+", "+loginID+", "+pgId+", "+ip+", "+act);
            if(!"".equals(sn_docno)){
                paramList.add(sn_docno) ;
            }else{
                paramList.add("") ;
            }
            if(!"".equals(rt_docno)){
                paramList.add(rt_docno) ;
            }else{
                paramList.add("") ;
            }
            if(!"".equals(reportno)){
                paramList.add(reportno) ;
            }else{
                paramList.add("") ;
            }
            if(!"".equals(reportno_seq)){
                paramList.add(reportno_seq) ;
            }else{
                paramList.add("") ;
            }
            if("TC36".equals(pgId) && "Report".equals(act)){
                paramList.add("Q") ;
            }else{
            //操作類別 I-新增，U-查詢，D-刪除，Q-明細，P-列印
                if("Edit".equals(act) || "Qry2".equals(act) || "Qry".equals(act) ){
                    paramList.add("Q") ;
                }else if("Insert".equals(act)){
                    paramList.add("I") ;
                }else if("Update".equals(act)){
                    paramList.add("U") ;
                }else if("Delete".equals(act)){
                    paramList.add("D") ;
                }else if("Report".equals(act)|| "createRpt".equals(act) ){
                    paramList.add("P") ;
                }
            }
            updateDBSqlList.add(sqlCmd.toString()) ;
            updateDBDataList.add(paramList) ;
            updateDBSqlList.add(updateDBDataList) ;
            updateDBList.add(updateDBSqlList) ;
             if(DBManager.updateDB_ps(updateDBList)){
                   errMsg = "";
             }else{
                   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
              }
        }catch (Exception e){
                System.out.println(e+":"+e.getMessage());
                errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>";
        }
        return errMsg;
    }    
    //101.09.21 字串轉base64,以固定長度分割,回傳List by 2295
    //104.06.03 讀取編碼增加UTF-8 by 2295
    public static List parseLenBase64encode(String strSource,int parseLen){
        List content_list = new LinkedList();
        
        System.out.println("strSource="+strSource);
        try{
        byte[] srcDate = strSource.getBytes("UTF-8");
        String base64Data = Base64.encode(strSource.getBytes("UTF-8"));
        String substr = "";
        
        boolean checkLen = true;
        //System.out.println("base64Data="+base64Data);
        while(checkLen){
           if(base64Data.length() > parseLen){
               substr = base64Data.substring(0,parseLen);
               System.out.println("substr="+substr);
               content_list.add(substr);
               System.out.println("add="+substr);
               base64Data = base64Data.substring(parseLen,base64Data.length());
               System.out.println("last="+base64Data);
           }else{
               checkLen = false;
           }    
        }
        content_list.add(base64Data);
        System.out.println("content_list.size()="+content_list.size());
        
        String decode = "";
        for(int i=0;i<content_list.size();i++){
            decode += (String)content_list.get(i);   
        }
        System.out.println("decode="+decode);
        String data = new String(Base64.decode(decode.getBytes("UTF-8")),"UTF-8");
     
        System.out.println("data="+data);
        
        }catch(Exception e){
            System.out.println(e+e.getMessage());
        }
        return content_list;
        
    }
    /***************************************************************************
     *  執行壓縮加密檔案     
     *  lguser_id 使用者id 
     *  zipname 壓縮檔 (ex)D:\\7ZIP\\file.zip 
     *  files2zip 欲壓縮檔案  (ex)D:\\7ZIP\\file.xls  
     *  @return zipFile 
     */
    /*
    public static void createZipFile(String lguser_id,String zipname,String files2zip){
        try{
            StringBuffer sql = new StringBuffer();      
            List paramList = new ArrayList();
            sql.append(" select muser_password from wtt01 where muser_id=? ");
            paramList.add(lguser_id);
            List qList = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
            String pwd = Utility.decode((String)((DataObject)qList.get(0)).getValue("muser_password"));//取得原始密碼
            System.out.println("DOS Command => "+"D:\\7ZIP\\7ZA a -tzip "+zipname+" "+files2zip+" -p"+pwd);
            Process p = Runtime.getRuntime().exec("D:\\7ZIP\\7ZA a -tzip "+zipname+" "+files2zip+" -p"+pwd);
        } // end try
        catch(Exception e){
            System.out.println("createZipFile Exception e ="+e);
        } // end catch(Exception e)
    }
    */
    /***************************************************************************
     *  解壓縮檔案 
     * @param unzipFolder 解壓縮後檔案路徑  (ex)D:\\7ZIP\\unzipfiles
     * @param zipname 欲解壓縮檔案 (ex)D:\\7ZIP\\file.zip      
     * @return 
     */
    public static void unZipFile(String pwd,String unzipFolder,String zipname){
        try{
            System.out.println("DOS Command => "+"D:\\7ZIP\\7ZA x "+zipname+" -o"+unzipFolder+" -p"+pwd);
            Process p = Runtime.getRuntime().exec("D:\\7ZIP\\7ZA x "+zipname+" -o"+unzipFolder+" -p"+pwd);
        }// end try
        catch(Exception e){
            System.out.println("unZipFile Exception e ="+e);
        } // end catch(Exception e)
    }
    
    /***************************************************************************
     * 字串遮蔽功能
     * @param str              --欲遮蔽之字串 ex:A123456789
     * @param startChar        --遮蔽起始位置 ex:4
     * @param masklength       --遮蔽長度 ex:4
     * @param maskChar         --遮蔽符號 ex:*
     * @return Mask characters --ex:A12****789
     * @throws Exception
     */
    public static String maskChar(String str,int startChar,int masklength,String maskChar){
        String cover="";
        try{
            if(str.length()>0){
                for(int i=1;i<=masklength;i++) cover+=maskChar;
                str = str.substring(0, startChar-1)+cover
                        +str.substring(startChar-1+masklength, str.length());
            }
        }catch(Exception e){
            System.out.println("maskChar Exception e ="+e);
        } 
        return str;
    }
    /**
     *  執行壓縮加密檔案     
     *  inFileName 來源檔名 
     *  zipName 目的壓縮檔名
     *  strPwd 密碼 
     *  @return encFile 
     *  true 壓縮成功 
     *  false 壓縮失敗
     */
    public static boolean createEncZipFile(String inFileName,String zipName,String lguser_id) {
        boolean encFile = false;
        try {
            // Initiate ZipFile object with the path/name of the zip file.
            ZipFile zipFile = new ZipFile(Utility.getProperties("reportDir")+System.getProperty("file.separator")+zipName);
            // Build the list of files to be added in the array list
            // Objects of type File have to be added to the ArrayList
            String strPwd = "";
            StringBuffer sql = new StringBuffer();  
            List paramList = new ArrayList();
            sql.append(" select muser_password from wtt01 where muser_id=? ");
            paramList.add(lguser_id);
            List qList = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
            if(qList.size()>0){
                strPwd = Utility.decode((String)((DataObject)qList.get(0)).getValue("muser_password"));//取得原始密碼
            }
            ArrayList filesToAdd = new ArrayList();
            filesToAdd.add(new File(Utility.getProperties("reportDir")+System.getProperty("file.separator")+inFileName));
            //filesToAdd.add(new File("c:\\ZipTest\\myvideo.avi"));
            //filesToAdd.add(new File("c:\\ZipTest\\mysong.mp3"));
            
            // Initiate Zip Parameters which define various properties such
            // as compression method, etc. More parameters are explained in other
            // examples
            ZipParameters parameters = new ZipParameters();
            parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE); // set compression method to deflate compression
            
            // Set the compression level. This value has to be in between 0 to 9
            // Several predefined compression levels are available
            // DEFLATE_LEVEL_FASTEST - Lowest compression level but higher speed of compression
            // DEFLATE_LEVEL_FAST - Low compression level but higher speed of compression
            // DEFLATE_LEVEL_NORMAL - Optimal balance between compression level/speed
            // DEFLATE_LEVEL_MAXIMUM - High compression level with a compromise of speed
            // DEFLATE_LEVEL_ULTRA - Highest compression level but low speed
            parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL); 
            
            // Set the encryption flag to true
            // If this is set to false, then the rest of encryption properties are ignored
            parameters.setEncryptFiles(true);
            
            // Set the encryption method to AES Zip Encryption
            parameters.setEncryptionMethod(Zip4jConstants.ENC_METHOD_AES);
            
            // Set AES Key strength. Key strengths available for AES encryption are:
            // AES_STRENGTH_128 - For both encryption and decryption
            // AES_STRENGTH_192 - For decryption only
            // AES_STRENGTH_256 - For both encryption and decryption
            // Key strength 192 cannot be used for encryption. But if a zip file already has a
            // file encrypted with key strength of 192, then Zip4j can decrypt this file
            parameters.setAesKeyStrength(Zip4jConstants.AES_STRENGTH_256);
            
            // Set password
            if(strPwd != null && !"".equals(strPwd)) parameters.setPassword(strPwd);
            
            // Now add files to the zip file
            // Note: To add a single file, the method addFile can be used
            // Note: If the zip file already exists and if this zip file is a split file
            // then this method throws an exception as Zip Format Specification does not 
            // allow updating split zip files
            zipFile.addFiles(filesToAdd, parameters);
            encFile = true;
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (Exception e1){
            e1.printStackTrace();
        }
        return encFile;
    }
    /**
     *  執行壓縮檔案     
     *  workDir ZIP檔案工作目錄
     *  filenames 來源檔名List 
     *  zipName 目的壓縮檔名     
     *  @return zipFile 
     *  true 壓縮成功 
     *  false 壓縮失敗
     */
    public static boolean createZipFile(String workDir,List filenames,String zipName) {
        boolean encFile = false;
        try {
            // Initiate ZipFile object with the path/name of the zip file.
            ZipFile zipFile = new ZipFile(workDir+System.getProperty("file.separator")+zipName);
            // Build the list of files to be added in the array list
            // Objects of type File have to be added to the ArrayList
            
            ArrayList filesToAdd = new ArrayList();
            for(int i=0;i<filenames.size();i++){
            filesToAdd.add(new File(workDir+System.getProperty("file.separator")+(String)filenames.get(i)));
            }
            //filesToAdd.add(new File("c:\\ZipTest\\myvideo.avi"));
            //filesToAdd.add(new File("c:\\ZipTest\\mysong.mp3"));
            
            // Initiate Zip Parameters which define various properties such
            // as compression method, etc. More parameters are explained in other
            // examples
            ZipParameters parameters = new ZipParameters();
            parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE); // set compression method to deflate compression
            
            // Set the compression level. This value has to be in between 0 to 9
            // Several predefined compression levels are available
            // DEFLATE_LEVEL_FASTEST - Lowest compression level but higher speed of compression
            // DEFLATE_LEVEL_FAST - Low compression level but higher speed of compression
            // DEFLATE_LEVEL_NORMAL - Optimal balance between compression level/speed
            // DEFLATE_LEVEL_MAXIMUM - High compression level with a compromise of speed
            // DEFLATE_LEVEL_ULTRA - Highest compression level but low speed
            parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL); 
            
            // Set the encryption flag to true
            // If this is set to false, then the rest of encryption properties are ignored
            //parameters.setEncryptFiles(false);
            
            // Set the encryption method to AES Zip Encryption
            //parameters.setEncryptionMethod(Zip4jConstants.ENC_METHOD_AES);
            
            // Set AES Key strength. Key strengths available for AES encryption are:
            // AES_STRENGTH_128 - For both encryption and decryption
            // AES_STRENGTH_192 - For decryption only
            // AES_STRENGTH_256 - For both encryption and decryption
            // Key strength 192 cannot be used for encryption. But if a zip file already has a
            // file encrypted with key strength of 192, then Zip4j can decrypt this file
            //parameters.setAesKeyStrength(Zip4jConstants.AES_STRENGTH_256);
            
            // Set password
            //if(strPwd != null && !"".equals(strPwd)) parameters.setPassword(strPwd);
            
            // Now add files to the zip file
            // Note: To add a single file, the method addFile can be used
            // Note: If the zip file already exists and if this zip file is a split file
            // then this method throws an exception as Zip Format Specification does not 
            // allow updating split zip files
            zipFile.addFiles(filesToAdd, parameters);
            encFile = true;
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (Exception e1){
            e1.printStackTrace();
        }
        return encFile;
    }
    //103.05.23 add 檢核有誤的農漁會信部名稱
    public static String getWML01_Error(String m_year,String m_month,String bank_type,String report_no){
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList();//傳內的參數List   
        DataObject bean = null;
        String bank_name = "";
        String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
        sqlCmd.append(" select bn01.bank_no,bn01.bank_name");
        sqlCmd.append(" from WML01 left join (select * from bn01 where m_year=?)bn01 on wml01.bank_code =bn01.bank_no");                   
        sqlCmd.append(" where WML01.M_YEAR=?"); 
        sqlCmd.append(" and M_MONTH=?");
        sqlCmd.append(" and bank_type in (?)");
        sqlCmd.append(" and REPORT_NO=?");
        sqlCmd.append(" and upd_code =?");
                 
        paramList.add(wlx01_m_year);       
        paramList.add(m_year); 
        paramList.add(m_month); 
        paramList.add(bank_type); 
        paramList.add(report_no); 
        paramList.add("E"); 
        
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        
        for(int i=0;i<dbData.size();i++){
            bean = (DataObject)dbData.get(i);            
            bank_name += (String)bean.getValue("bank_no") +(String)bean.getValue("bank_name");
            bank_name += i<dbData.size()-1?":":"";
        }
        return  bank_name; 
    }
    /* 105.09.12 取得貸款經辦機構名稱 */
    public static List getLoanBank(){
    	List dbData = new ArrayList();
	    List paramList = new ArrayList();
	    String sqlCmd = 
	    		"select bn01.bn_type,HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE,bn01.m_year "
	    		+ "from  BN01, WLX01 "
	    		+ "where BN01.BANK_NO = WLX01.BANK_NO(+) " 
	    		+ "and bank_type in ('A','6','7') "//--機構類別:銀行/農會/漁會
	    		+ "and wlx01.m_year = bn01.m_year " 
	    		+ "and bn01.m_year=100 "//--都固定只抓100年度的          
	    		+ "order by hsien_id,bank_type,BANK_NO ";
		dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"bn_type,hsien_id,bank_no,bank_name,bank_type,m_year");
		return dbData;	    
    }
    //106.05.31 取得機構名稱
    /****************************************************************************
     * 取得所有機構名稱(99年跟100年縣市別)  
     */
    public static List getbn01Bank (String m_year) throws Exception{
            List paramList = new ArrayList();
            StringBuffer sql = new StringBuffer();
            sql.append(" select bank_no,")
                    .append(" bank_name ") //機構名稱
                    .append(" from bn01 ")
                    .append(" where m_year=? ")
                    .append(" and bank_type in ('6','7') ")
                    .append(" and bn_type !='2' ")
                    .append(" order by bank_no ");
            paramList.add(m_year);
            
            List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");                    
            return dbData;          
    }
    
    //110.04.08 add 撤銷項目 by 6493
  	//自110年4月份不開放申報990621。
    public static  List get_ncacno_990621(String S_YEAR, String S_MONTH){
  		StringBuffer sqlCmd = new StringBuffer();
  		List paramList = new ArrayList();
  		sqlCmd.append(" select acc_code ");//自某年月取消申報
  		sqlCmd.append(" ,to_char(cancel_year * 100 + cancel_month) ");
  		sqlCmd.append(" from ncacno ");
  		sqlCmd.append(" where acc_code='990621' ");
  		sqlCmd.append(" and to_char(cancel_year * 100 + cancel_month) > ? ");
  		paramList.add(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH));
  		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(), paramList, "");
  		return dbData;
  	}
  	//110.04.08 add 撤銷項目 by 6493
    public static  List getData_revoke(String bank_code){
  		StringBuffer sqlCmd = new StringBuffer();
  		List paramList = new ArrayList();
  		sqlCmd.append(" select acc_code ");
  		sqlCmd.append(" , sub_acc_code ");
  		sqlCmd.append(" , doc_no ");
  		sqlCmd.append(" from revoke_doc ");
  		sqlCmd.append(" where  bank_code = ?");
  		paramList.add(bank_code);
  		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(), paramList, "");
  		DataObject revokeBean = null;
  		String revokeDoc = "";
  		for(int j=0; j<dbData.size(); j++){
  			revokeBean = (DataObject)dbData.get(j);
  			revokeDoc = (String)revokeBean.getValue("doc_no");
  			revokeDoc = revokeDoc.substring(revokeDoc.indexOf("農授金字第")+5, revokeDoc.indexOf("號函"));
  			revokeBean.setValue("doc_no",revokeDoc);
  		}
  		return dbData;
  	}
  	//110.04.08 add acc明細項目 by 6493
    public static  List get_ncacno_detail_sub_acc(String acc_code){
  		StringBuffer sqlCmd = new StringBuffer();
  		List paramList = new ArrayList();
  		sqlCmd.append(" select sub_acc_code ");
  		sqlCmd.append(" , sub_acc_name ");
  		sqlCmd.append(" from ncacno_detail ");
  		sqlCmd.append(" where src_acc_code = ? ");
  		sqlCmd.append(" order by acc_range");
  		paramList.add(acc_code);
  		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(), paramList, "");
  		return dbData;
  	}
    
}
