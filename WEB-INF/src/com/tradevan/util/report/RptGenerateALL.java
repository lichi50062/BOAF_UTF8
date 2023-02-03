/*
	97.10.08 create 檢查局財務報表-合併單一檔案 by 2295
	97.11.25 fix 將中文檔名轉成ISO8859_1傳至檢查局中文字才會正常 by 2295	
	98.01.07 fix 將資料格式.轉換成數值格式儲存 by 2295
	               檔名更改成"信用部名稱"+西元年月日(該月份的最後一天),不產生流水號.每個月3次產生皆覆蓋原檔案 by 2295 
	98.02.11 fix 將資產負債表、損益表，當會計科目代號>0時才寫入cell  by 2479
	98.04.01 add 應予評估資產彙總表 by 2295
	98.04.09 產生農漁會信用部主要經營指標明細表 by 2295
	98.06.29 add 解釋函令/限制或核准業務/處分書/舞弊案件/檢舉書/理監事基本資料 上傳至檢查局 by 2295
	98.07.27 fix A06逾期放款統計表-A01放款合計.合計欄由990000逾期放款金額修正為 A01放款總額(120000+120800+150300) by 2295
	98.09.10 add 增加上傳檔案多層子目錄putFiles_multiSubDir by 2295
	98.09.25 fix 建立子目錄時,一層一層chang至該目錄 by 2295
	98.10.19 add 傳送e-mail by 2295
	99.01.05 add 檢查局上傳/下載作業,增加一組e-mail by 2295
	99.10.12 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			      使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
   102.07.15 add 103/01以後,漁會.資產負債表/損益表.套用新表格 by 2295
   102.09.23 add 各項法定比率增加A99.992710建築放款總額  by 2295
   102.10.09 add FR055W理監事基本資料表,IDN加密後,還原顯示 by 2295
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
   103.02.12 fix 資產負債表.逾放金額field_over/逾放比率field_over_rate/存放比率field_dc_rate.更改從A01_operation抓 by 2295
   103.03.13 fix 調整上傳目錄for utf8編碼 by 2295
   108.01.08 fix 調整A05報表格式 by 2295
   109.05.12 add 因檢查局FTP增加TLS設定,調整上傳使用FTPSClient by 2295
   111.01.11 fix 因檢查局FTPS只開放TLS1.2上傳,但TLS1.2不支援java1.6版本,改由FebFtps使用java1.8版本上傳檔案,此功能僅產生報表 by 2295
*/
package com.tradevan.util.report;



import org.apache.commons.net.PrintCommandListener;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPReply;
import org.apache.commons.net.ftp.FTPSClient;
import org.apache.commons.net.util.TrustManagerUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.security.KeyStore;
import java.text.SimpleDateFormat;

import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSocketFactory;
import javax.net.ssl.TrustManager;
import javax.net.ssl.TrustManagerFactory;
import javax.net.ssl.X509TrustManager;

import com.tradevan.util.ftp.JakartaFtpWrapper;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.ftp.MyFTPSClient;
import com.tradevan.util.xml.parseDOM;


public class RptGenerateALL {	
	static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	static SimpleDateFormat filenameformat = new SimpleDateFormat("yyyyMMdd");
	static Date nowlog = new Date();
	static Calendar logcalendar;
	static File xlsDir,febxlsDir,febBKxlsDir,febxlsDir_FR001WB,febBKxlsDir_FR001WB = null;
	static File febxlsDir_MCRptAll,febBKxlsDir_MCRptAll,febxlsDir_MC012W,febBKxlsDir_MC012W=null;
	static File febxlsDir_MC014W,febBKxlsDir_MC014W,febxlsDir_FR055W,febBKxlsDir_FR055W = null;
	
	public static void main(String args[]) { 
	    try {
	    	File logfile;
	    	FileOutputStream logos=null;    	
	    	BufferedOutputStream logbos = null;
	    	PrintStream logps = null;
	    	Date nowlog = new Date();
	    	SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
	    	SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	    	Calendar logcalendar;
	    	File logDir = null;
	    	
	    	
	    	   //申報上個月份的報表
	    	   Calendar now = Calendar.getInstance();
	    	   String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    	   String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
	    	   if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
	    	      YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
	    	      MONTH = "12";	
	    	   }else{    
	    	      MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
	    	   }	    	   
	    	 
	    	   logDir  = new File(Utility.getProperties("logDir"));
	    	   if(!logDir.exists()){
	    	        if(!Utility.mkdirs(Utility.getProperties("logDir"))){
	    	    	   System.out.println("目錄新增失敗");
	    	        }    
	    	   }
	    	   logfile = new File(logDir + System.getProperty("file.separator") + "MC008W."+ logfileformat.format(nowlog));						 
	    	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"MC008W."+ logfileformat.format(nowlog));
	    	   logos = new FileOutputStream(logfile,true);  		        	   
	    	   logbos = new BufferedOutputStream(logos);
	    	   logps = new PrintStream(logbos);			    
	    	 
	    	   String errMsg = "";
	    	   RptGenerateALL a = new RptGenerateALL();
	    	   errMsg += a.createRpt(YEAR,MONTH,logps);
	    	   System.out.println("errMsg = "+errMsg);   
	    	   logcalendar = Calendar.getInstance(); 
	    	   nowlog = logcalendar.getTime();
	    	   logps.println(logformat.format(nowlog)+errMsg);		    					    
	    	   logps.flush();
	    	   System.out.println("=============執行檢查局財務報表發佈結束===========");
	    
	    /*
	    JakartaFtpWrapper ftp = new JakartaFtpWrapper();
        String nowDir = "";
        String rptIP_feb=Utility.getProperties("rptIP_feb");         
        String rptID_feb=Utility.getProperties("rptID_feb");            
        String rptPwd_feb=Utility.getProperties("rptPwd_feb");  
        String feb_serverRptDir=Utility.getProperties("feb_serverRptDir");
        
        if(ftp.connectAndLogin(rptIP_feb, rptID_feb, rptPwd_feb)) {//connect成功時
             System.out.println("Connected to " + rptIP_feb);  
                 System.out.println("Welcome message:\n" + ftp.getReplyString());
                 System.out.println("Current Directory: " + ftp.printWorkingDirectory());                    
                 //System.out.println("parentDir=" + parentDir);
                 //97.10.02要先切換到/boaf/目錄下.才可切換到其他下層目錄
                 //System.out.println("change home dir [/boaf/] ?? " + ftp.changeWorkingDirectory("/boaf/"));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("UTF-8"),"ISO8859_1") + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes("UTF-8"),"ISO8859_1")));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("UTF-8")) + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes("UTF-8"),"big5")));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("UTF-8")) + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes(),"big5")));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("UTF-8")) + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes(),"ISO8859_1")));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes("ISO8859_1"),"UTF-8")));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes("ISO8859_1"))));
                 System.out.println("change home dir [" + new String(feb_serverRptDir.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftp.changeWorkingDirectory(new String(feb_serverRptDir.getBytes("ISO8859_1"),"big5")));
           
        } 
        */
	    } catch(Exception e) {
            System.out.println(e);
            System.out.println(e.getMessage());
            e.printStackTrace();                    
        }   
              
	  } 
	 
    public static String createRpt(String S_YEAR,String S_MONTH,PrintStream logps){
		String errMsg = "";
		String putMsg = "";	
		List dbData = null;		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		String bank_code="";
		String bank_name = "";
		String bank_type = "";	
		String filename = "";
		FileInputStream finput = null;
		FileOutputStream fout = null; 
		POIFSFileSystem fs = null;
		String[] fname;	
		File tmpFile = null;
		String copyResult = "";
		String serialNo="01";
		List filename_List  = new LinkedList();	
		List filename_List_FR001WB  = new LinkedList();//主要經營指標/各會員別放款金額一覽表	
		List filename_List_MCRptAll = new LinkedList();//解釋函令/限制或核准業務/處分書
		List filename_List_MC012W = new LinkedList();//舞弊案件
		List filename_List_MC014W = new LinkedList();//檢舉書
		List filename_List_FR055W = new LinkedList();//理監事基本資料
		String errRptMsg = "";
		HSSFWorkbook wb = null;
		DataObject bean = null;
		try{			 
			Utility.printLogTime("RptGenerateALL begin time");
			String rptIP_feb=Utility.getProperties("rptIP_feb");			
	        String rptID_feb=Utility.getProperties("rptID_feb");			
	        String rptPwd_feb=Utility.getProperties("rptPwd_feb");	
	        xlsDir = new File(Utility.getProperties("xlsDir"));
        	febxlsDir = new File(Utility.getProperties("febxlsDir"));//財務報表-合併檔案   
        	febBKxlsDir = new File(Utility.getProperties("febBKxlsDir"));   
        	febxlsDir_FR001WB = new File(Utility.getProperties("febxlsDir_FR001WB"));//主要經營指標/各會員別放款金額一覽表 
        	febBKxlsDir_FR001WB = new File(Utility.getProperties("febBKxlsDir_FR001WB"));  
        	febxlsDir_MCRptAll = new File(Utility.getProperties("febxlsDir_MCRptAll"));//解釋函令/限制或核准業務/處分書
        	febBKxlsDir_MCRptAll = new File(Utility.getProperties("febBKxlsDir_MCRptAll")); 
        	febxlsDir_MC012W = new File(Utility.getProperties("febxlsDir_MC012W"));//舞弊案件
        	febBKxlsDir_MC012W = new File(Utility.getProperties("febBKxlsDir_MC012W")); 
        	febxlsDir_MC014W = new File(Utility.getProperties("febxlsDir_MC014W"));//檢舉書
        	febBKxlsDir_MC014W = new File(Utility.getProperties("febBKxlsDir_MC014W")); 
        	febxlsDir_FR055W = new File(Utility.getProperties("febxlsDir_FR055W"));//理監事基本資料
        	febBKxlsDir_FR055W = new File(Utility.getProperties("febBKxlsDir_FR055W")); 
    	
	        errMsg += mkDir();//建立報表目錄    
	        //99.10.12 add 查詢年度100年以前.縣市別不同===============================
		    String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
		    String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
		    //===================================================================== 
		  
    		/*97.11.25只能上傳非中文的檔案 
    		putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir"), Utility.getProperties("febxlsDir")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir"),filename_List,logps);
    		*/
    		
            //dbData = DBManager.QueryDB("select bank_type,bank_no,bank_name from bn01 where bank_type in ('6','7') and bank_no='"+bank_code+"'",""); 
    		
    		//allen debug
    		//dbData = DBManager.QueryDB("select bank_type,bank_no,bank_name from bn01 where bank_no in ('6030016','6200189','6040017','6230012','6080011') and bank_type in ('6','7')  and bank_no <> '8888888' and bn01.bn_type <> '2' order by bank_type,bank_no","");
            
		    //正式
		    sqlCmd.append("select bank_type,bank_no,bank_name from bn01 where m_year = ? and bank_type in ('6','7')  and bank_no <> '8888888' and bn01.bn_type <> '2' order by bank_type,bank_no");//正式
		    //測試
		    //sqlCmd.append("select bank_type,bank_no,bank_name from bn01 where m_year = ? and bank_type in ('6','7')  and bank_no <> '8888888' and bank_no in ('6030016') and bn01.bn_type <> '2' order by bank_type,bank_no");//測試
		    
		    paramList.add(wlx01_m_year);
	        dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    		
    		//dbData = DBManager.QueryDB("select bank_type,bank_no,bank_name from bn01 where bank_type in ('6','7')  and bank_no <> '8888888' and bank_no in ('6040017') and bn01.bn_type <> '2' order by bank_type,bank_no","");
    			
    		errMsg += moveBKDir(logps);//將febxlsDir下的檔案.搬到febBKxlsDir下
    		
            //98.01.07 add 檔名更改成"信用部名稱"+西元年月日(該月份的最後一天),不產生流水號.每個月3次產生皆覆蓋原檔案 by 2295
            serialNo = Utility.getLastDay(String.valueOf(Integer.parseInt(S_YEAR)+1911)+(S_MONTH.length()==1?"0":"")+S_MONTH+"01","yyyyMMdd");
            System.out.println("月份最後一天="+serialNo);
            Date d = java.sql.Date.valueOf(serialNo.substring(0,10));
            serialNo = filenameformat.format(d);
            System.out.println("serialNo="+serialNo);
            
            //產生所有農漁會信用部的財務報表檔=====================================================            
    		for(int i=0;i<dbData.size();i++){
    			bean = (DataObject)dbData.get(i);
    			bank_code=(String)bean.getValue("bank_no");
                bank_name=(String)bean.getValue("bank_name");
                bank_type=(String)bean.getValue("bank_type");
                filename = bank_name+serialNo;
                //if(!bank_code.equals("5030019")){
                    //continue;
                //}
                //filename = bank_code+"_"+Utility.getDateFormat("yyyyMMdd");
                //98.01.07 不產生流水號.每個月3次產生皆覆蓋原檔案
                //serialNoLoop:
                //if(i==0){//檢查檔案流水號.最多編到99
                //	for(int k=1;k<100;k++){
                //	    tmpFile = new File(febBKxlsDir+System.getProperty("file.separator")+filename+(k<10?"0":"")+k+".xls");
                //	    System.out.println("check file name="+febBKxlsDir+System.getProperty("file.separator")+filename+(k<10?"0":"")+k+".xls");
                //	    printRptMsg(logps,"","check file name="+febBKxlsDir+System.getProperty("file.separator")+filename+(k<10?"0":"")+k+".xls");
                //	    if(tmpFile.exists()){
                //	    	continue;
                //	    }else{
                //	    	serialNo = (k<10?"0":"")+String.valueOf(k);
                //	    	break serialNoLoop;
                //	    }
                //	}                	
                //}
                //filename += serialNo;        
             
                //add 103/01以後,漁會資產負債表/損益表.套用新表格(增加/異動科目代號) 
                if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){ 
                    finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"財務資料_檢查局格式_"+(bank_type.equals("6")?"農會":"漁會_10301")+".xls" );
                    System.out.println("財務資料_檢查局格式_10301.xls");
                }else{
                    finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"財務資料_檢查局格式_"+(bank_type.equals("6")?"農會":"漁會")+".xls" );
                    System.out.println("財務資料_檢查局格式.xls");
                }

			   
			   
	  	        //設定FileINputStream讀取Excel檔
	  		    fs = new POIFSFileSystem( finput );
	  		    wb = new HSSFWorkbook(fs);
	  		    
	  		    //各項法定比率
	  		    HSSFSheet sheet = wb.getSheetAt(0);
	  		    HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定	        	
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比
	            HSSFFooter footer = sheet.getFooter();
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4	
 			    errRptMsg = A02Rpt(S_YEAR,S_MONTH,bank_code,bank_name,wb,sheet);
 			    printRptMsg(logps,"各項法定比率",bank_code+bank_name+errRptMsg);
 			   
	            //淨值占風險性資產比率
	  		    sheet = wb.getSheetAt(1);
	  		    ps = sheet.getPrintSetup(); //取得設定	      		
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比	        
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4	
	            errRptMsg = A05Rpt(S_YEAR,S_MONTH,bank_type,bank_code,bank_name,wb,sheet);
	            printRptMsg(logps,"淨值占風險性資產比率",bank_code+bank_name+errRptMsg);
	            
	            //損益表
	  		    sheet = wb.getSheetAt(2);
	  		    ps = sheet.getPrintSetup(); //取得設定	       	
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比	        
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	            errRptMsg = A01Rpt_rpt2(S_YEAR,S_MONTH,bank_type,bank_code,bank_name,wb,sheet);
	            printRptMsg(logps,"損益表",bank_code+bank_name+errRptMsg);
	            
	            //資產負債表
	  		    sheet = wb.getSheetAt(3);
	  		    ps = sheet.getPrintSetup(); //取得設定	       	
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比	        
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	            errRptMsg = A01Rpt_rpt1(S_YEAR,S_MONTH,bank_type,bank_code,bank_name,wb,sheet);
	            printRptMsg(logps,"資產負債表",bank_code+bank_name+errRptMsg);
	            
	            //逾期放款統計表
	  		    sheet = wb.getSheetAt(4);
	  		    ps = sheet.getPrintSetup(); //取得設定	       	
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比	        
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4 
	            errRptMsg = A06Rpt(S_YEAR,S_MONTH,bank_type,bank_code,bank_name,wb,sheet);
	            printRptMsg(logps,"逾期放款統計表",bank_code+bank_name+errRptMsg);
	            
	            //應予評估資產彙總表
	  		    sheet = wb.getSheetAt(5);
	  		    ps = sheet.getPrintSetup(); //取得設定	       	
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )100 ); //列印縮放百分比	        
	            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4 
	            errRptMsg = A10Rpt(S_YEAR,S_MONTH,bank_type,bank_code,bank_name,wb,sheet);
	            printRptMsg(logps,"應予評估資產彙總表",bank_code+bank_name+errRptMsg);
	            
  	            fout=new FileOutputStream(febxlsDir + System.getProperty("file.separator")+filename+".xls");
	            wb.write(fout);
	            //儲存 
	            fout.close();
	            //======================================================================================
	            
                //98.06.25限制或核准業務_解釋函令_處分書===================================================
                finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"限制或核准業務_解釋函令_處分書_合併檔案_檢查局.xls" );
    		    System.out.println("限制或核准業務_解釋函令_處分書_合併檔案_檢查局.xls");
      	        //設定FileINputStream讀取Excel檔
      		    fs = new POIFSFileSystem( finput );
      		    wb = new HSSFWorkbook(fs);
        		
        		//產生農漁會信用部解釋函令=========================================================    		
        		errRptMsg = RptMC011W.createRpt(bank_code,"限制或核准業務_解釋函令_處分書_合併檔案",wb);
        		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
        		printRptMsg(logps,"解釋函令",bank_code+bank_name+errRptMsg);
        		
        		//限制或核准業務函令=========================================================    		
        		errRptMsg = RptMC010W.createRpt(bank_code,"限制或核准業務_解釋函令_處分書_合併檔案",wb);
        		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
        		printRptMsg(logps,"限制或核准業務函令",bank_code+bank_name+errRptMsg);
        		
        		//處分書===============================================================    		
        		errRptMsg = RptMC013W.createRpt(bank_code,"限制或核准業務_解釋函令_處分書_合併檔案",wb);
        		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
        		printRptMsg(logps,"處分書",bank_code+bank_name+errRptMsg);
        		
        		fout=new FileOutputStream(febxlsDir_MCRptAll + System.getProperty("file.separator")+bank_name+".xls");
    	        wb.write(fout);
    	        //儲存 
    	        fout.close();    
    	        
    	        //產生舞弊案件==================================================================
    	        errRptMsg = RptMC012W.createRpt(bank_code);    	        
        		if(errRptMsg.equals("")){
        			copyResult = Utility.CopyFile(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"舞幣案件_檢查局.xls",febxlsDir_MC012W+System.getProperty("file.separator")+bank_name+".xls");
        			errRptMsg = "報表產生完成";
        		}
        		printRptMsg(logps,"舞弊案件",bank_code+bank_name+errRptMsg);
        		
        		//產生檢舉書==================================================================
    	        errRptMsg = RptMC014W.createRpt(bank_code);    	        
        		if(errRptMsg.equals("")){
        			copyResult = Utility.CopyFile(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"檢舉書_檢查局.xls",febxlsDir_MC014W+System.getProperty("file.separator")+bank_name+".xls");
        			errRptMsg = "報表產生完成";
        		}
        		printRptMsg(logps,"檢舉書",bank_code+bank_name+errRptMsg);
        		
        		//理監事基本資料==================================================================
        		//區身上半年/下半年.ex:xxx農會信用部98上半年.xls / xxx農會信用部98下半年.xls
        		//102.10.09 add IDN加密後,還原顯示 by 2295
    	        errRptMsg = RptFR055W.createRpt(bank_code);    	        
        		if(errRptMsg.equals("")){
        			copyResult = Utility.CopyFile(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"理監事基本資料_檢查局.xls",febxlsDir_FR055W+System.getProperty("file.separator")+bank_name+S_YEAR+(Integer.parseInt(S_MONTH)<=6?"上半年":"下半年")+".xls");
        			errRptMsg = "報表產生完成";
        		}
        		printRptMsg(logps,"理監事基本資料",bank_code+bank_name+errRptMsg);
    		}//================================================================================
    		
            //財務資料_檢查局格式_農漁會綜合指標
    		finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"財務資料_檢查局格式_農漁會綜合指標.xls" );
		    System.out.println("財務資料_檢查局格式_農漁會綜合指標.xls");
  	        //設定FileINputStream讀取Excel檔
  		    fs = new POIFSFileSystem( finput );
  		    wb = new HSSFWorkbook(fs);
    		
    		//98.06.16產生農漁會信用部主要經營指標明細表=========================================================    		
    		//農會
    		errRptMsg = RptFR001WB.createRpt(S_YEAR,S_MONTH,"1000","6","1","農漁會綜合指標",wb);//1明細表
    		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
    		printRptMsg(logps,"全体農會信用部主要經營指標明細表",errRptMsg);    		
    		//漁會
    		errRptMsg = RptFR001WB.createRpt(S_YEAR,S_MONTH,"1000","7","1","農漁會綜合指標",wb);//1明細表    	
    		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
    		printRptMsg(logps,"全体漁會信用部主要經營指標明細表",errRptMsg);
    		
    		//98.06.16 產生全體農漁會信用部各會員別放款金額一覽表=====================================================
    		errRptMsg = RptFR054W.createRpt(S_YEAR,S_MONTH,"1000","農漁會綜合指標",wb);
    		if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
    		printRptMsg(logps,"全體農漁會信用部各會員別放款金額一覽表",errRptMsg);
    		fout=new FileOutputStream(febxlsDir_FR001WB + System.getProperty("file.separator")+S_YEAR+(S_MONTH.length()<2?"0":"")+S_MONTH+"農漁會綜合指標.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();    		
    		//===================================================================================================    		
	        
    		if(febxlsDir.exists() && febxlsDir.isDirectory()){
    			fname= febxlsDir.list(); //財務報表-合併檔案
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){
                    	filename_List.add(fname[c]);
                        System.out.println(fname[c]);
                    }
                }
                fname= febxlsDir_FR001WB.list(); //財務報表-主要經營指標/各會員別放款金額一覽表	
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir_FR001WB")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){ 
                    	filename_List_FR001WB.add(fname[c]);
                    	System.out.println("FR001WB.add="+fname[c]);                    	
                    }
                }
                fname= febxlsDir_MCRptAll.list(); //解釋函令/限制或核准業務/處分書
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir_MCRptAll")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){ 
                    	filename_List_MCRptAll.add(fname[c]);
                    	System.out.println("MCRptAll.add="+fname[c]);                    	
                    }
                }
                
                fname= febxlsDir_MC012W.list(); //舞弊案件
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir_MC012W")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){ 
                    	filename_List_MC012W.add(fname[c]);
                    	System.out.println("MC012W.add="+fname[c]);                    	
                    }
                }
                fname= febxlsDir_MC014W.list(); //檢舉書
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir_MC014W")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){ 
                    	filename_List_MC014W.add(fname[c]);
                    	System.out.println("MC014W.add="+fname[c]);                    	
                    }
                }
                fname= febxlsDir_FR055W.list(); //理監事基本資料
                for(int c=0;c<fname.length;c++){
                	tmpFile = new File(Utility.getProperties("febxlsDir_FR055W")+System.getProperty("file.separator")+fname[c]);
                    if(!tmpFile.isDirectory()){ 
                    	filename_List_FR055W.add(fname[c]);
                    	System.out.println("FR055W.add="+fname[c]);                    	
                    }
                }
                //MyFTPSClient ftpC = new MyFTPSClient(rptIP_feb, rptID_feb, rptPwd_feb);//109.05.12 暫時拿掉  
                
                //putFiles(String remote_path, String local_path,String workDir,List filename)
                //putMsg = ftpC.putFiles(Utility.getProperties("feb_serverRptDir"), Utility.getProperties("febxlsDir")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir"),filename_List);
                
                //putFiles(server_host,username,password,remote_path,local_path,workDir,filename,logps)
                /*111.01.11改由 FebFtps提供上傳檔案
                //上傳合併檔案財務資料
                putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir"), Utility.getProperties("febxlsDir")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir"),filename_List,logps);
                //98.04.09 上傳主要經營指標明細表/各會員別放款金額一覽表                
                if(putMsg == null){//上傳檔案成功		  
                   putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_FR001WB"), Utility.getProperties("febxlsDir_FR001WB")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_FR001WB"),filename_List_FR001WB,logps);
                }                
                //98.06.26 上傳解釋函令/限制或核准業務/處分書
                if(putMsg == null){//上傳檔案成功		  
                   putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_MCRptAll"), Utility.getProperties("febxlsDir_MCRptAll")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_MCRptAll"),filename_List_MCRptAll,logps);
                }
                
                //98.06.26 上傳舞弊案件
                if(putMsg == null){//上傳檔案成功		  
                   putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_MC012W"), Utility.getProperties("febxlsDir_MC012W")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_MC012W"),filename_List_MC012W,logps);
                }
                
                //98.06.26 上傳檢舉書
                if(putMsg == null){//上傳檔案成功		  
                   putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_MC014W"), Utility.getProperties("febxlsDir_MC014W")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_MC014W"),filename_List_MC014W,logps);
                }
                
                //98.06.29 上傳理監事基本資料
                if(putMsg == null){//上傳檔案成功		  
                   putMsg = putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_FR055W"), Utility.getProperties("febxlsDir_FR055W")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_FR055W"),filename_List_FR055W,logps);
                }
                
                logcalendar = Calendar.getInstance(); 
 			    nowlog = logcalendar.getTime();
 			    
                
                //ftpC=null;//109.05.12 fix
                if(putMsg == null){//上傳檔案成功		                     
                   System.out.println("檔案上傳成功");
                   //將febxlsDir目錄下產生完成的檔案.並上傳成功的.搬移至所對應的bkdir================================= 
                   errMsg += moveBKDir_complete(logps,filename_List,filename_List_FR001WB,
                   									  filename_List_MCRptAll,filename_List_MC012W,
													  filename_List_MC014W,filename_List_FR055W);
                   
                   errMsg +="<br>上列檢查局各式報表,檔案上傳完成";
                   printRptMsg(logps,"","上列檢查局各式報表,檔案上傳完成"); 
                   send_Mail(S_YEAR+"年"+S_MONTH+"月檢查局財務報表上傳成功","檢查局各式報表,檔案上傳完成");//98.10.19 add 傳送e-mail
                }else{//end of 上傳檔案成功
                   errMsg += "<br>報表產生完成,但上傳至Sever未成功"+putMsg;   
                   send_Mail(S_YEAR+"年"+S_MONTH+"月檢查局財務報表上傳失敗","報表產生完成,但上傳至Sever未成功\n錯誤原因:"+putMsg+"\n煩請至MIS管理系統及檢查缺失追蹤管理系統平台->監理資訊共享平台->檢查局財務報表產生及上傳(MC008W)\n重新執行報表產生及上傳作業!!");//98.10.19 add 傳送e-mail
                }
                */
	        }//end of febxlsDir存在
    		
	        Utility.printLogTime("RptGenerateALL end time");	
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
			send_Mail(S_YEAR+"年"+S_MONTH+"月檢查局財務報表上傳失敗",e+e.getMessage());//98.10.19 add 傳送e-mail
		}
		return errMsg;
	}
    
    //建立報表存放目錄
    private static String mkDir(){
    	String errMsg="";
    	try{
    	    if(!xlsDir.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
 		       		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
 		    	}    
		    }    		
		    //檢查局財務報表檔
		    if(!febxlsDir.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir"))){
 		       		errMsg +=Utility.getProperties("febxlsDir")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir")+"目錄新增失敗";
 		    	}    
		    }
		    if(!febxlsDir_FR001WB.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir_FR001WB"))){
 		       		errMsg +=Utility.getProperties("febxlsDir_FR001WB")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir_FR001WB.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_FR001WB"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir_FR001WB")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febxlsDir_MCRptAll.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir_MCRptAll"))){
 		       		errMsg +=Utility.getProperties("febxlsDir_MCRptAll")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir_MCRptAll.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_MCRptAll"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir_MCRptAll")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febxlsDir_MC012W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir_MC012W"))){
 		       		errMsg +=Utility.getProperties("febxlsDir_MC012W")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir_MC012W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_MC012W"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir_MC012W")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febxlsDir_MC014W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir_MC014W"))){
 		       		errMsg +=Utility.getProperties("febxlsDir_MC014W")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir_MC014W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_MC014W"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir_MC014W")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febxlsDir_FR055W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febxlsDir_FR055W"))){
 		       		errMsg +=Utility.getProperties("febxlsDir_FR055W")+"目錄新增失敗";
 		    	}    
		    }
		    
		    if(!febBKxlsDir_FR055W.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_FR055W"))){
 		       		errMsg +=Utility.getProperties("febBKxlsDir_FR055W")+"目錄新增失敗";
 		    	}    
		    }
		
    	}catch(Exception e){
    		errMsg = "RptGenerateALL.mkDir Error:"+e+e.getMessage();
    	}
    	return errMsg;
    }
    //將原本已存在的檔案move至備份目錄下 
    private static String moveBKDir(PrintStream logps){
    	String errMsg="";
    	String[] fname;	
		File tmpFile = null;
		String copyResult = "";
    	try{
    	
		    //財務報表-合併檔案
		    //將febxlsDir下的檔案.搬到febBKxlsDir下
		    fname= febxlsDir.list(); //====列出此目錄下的所有檔案===================
		    
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir+System.getProperty("file.separator")+fname[j],febBKxlsDir+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
            
		    //財務報表-主要經營指標/各會員放款金額一覽表    		
		    fname= febxlsDir_FR001WB.list(); 
		    copyResult = "";
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir_FR001WB+System.getProperty("file.separator")+fname[j],febBKxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_FR001WB+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
            
            //解釋函令/限制或核准業務/處分書-合併檔案
		    fname= febxlsDir_MCRptAll.list(); 
		    copyResult = "";
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j],febBKxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MCRptAll+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
            
            //舞弊案件
		    fname= febxlsDir_MC012W.list(); 
		    copyResult = "";
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir_MC012W+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir_MC012W+System.getProperty("file.separator")+fname[j],febBKxlsDir_MC012W+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir_MC012W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MC012W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir_MC012W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MC012W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
            
            //檢舉書
		    fname= febxlsDir_MC014W.list(); 
		    copyResult = "";
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir_MC014W+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir_MC014W+System.getProperty("file.separator")+fname[j],febBKxlsDir_MC014W+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir_MC014W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MC014W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir_MC014W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_MC014W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
            //理監事基本資料
		    fname= febxlsDir_FR055W.list(); 
		    copyResult = "";
            for(int j=0;j<fname.length;j++){
            	tmpFile = new File(febxlsDir_FR055W+System.getProperty("file.separator")+fname[j]);
            	copyResult = Utility.CopyFile(febxlsDir_FR055W+System.getProperty("file.separator")+fname[j],febBKxlsDir_FR055W+System.getProperty("file.separator")+fname[j]);
            	 if(!tmpFile.isDirectory() && tmpFile.exists()){
            	 	System.out.println(febxlsDir_FR055W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_FR055W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);
            	 	printRptMsg(logps,"",febxlsDir_FR055W+System.getProperty("file.separator")+fname[j]+" copy to "+febBKxlsDir_FR055W+System.getProperty("file.separator")+fname[j]+" success ?? "+copyResult);            	 	
            	 	if(copyResult.equals("0")) tmpFile.delete();
            	 }
            }
    	}catch(Exception e){
    		System.out.println("RptGenerateALL.moveBKDir Error:"+e+e.getMessage());
    		errMsg += "RptGenerateALL.moveBKDir Error:"+e+e.getMessage();
    	}
    	return errMsg;
    }
    //將產生完成.並上傳至Server成功後的檔案.搬移至備份目錄
    private static String moveBKDir_complete(PrintStream logps,
    										 List filename_List,
											 List filename_List_FR001WB,
											 List filename_List_MCRptAll,
											 List filename_List_MC012W,
											 List filename_List_MC014W,
											 List filename_List_FR055W
											 ){
    	String errMsg="";    	
		File tmpFile = null;
		String copyResult = "";
    	try{
    		//財務資料-合併檔案區
            for(int i=0;i<filename_List.size();i++){
            	logps.println(logformat.format(nowlog)+(String)filename_List.get(i));
	            logps.flush();
                errMsg +="<br>財務資料-合併檔案:"+(String)filename_List.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir")+System.getProperty("file.separator")+(String)filename_List.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
                   copyResult = Utility.CopyFile(febxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i),febBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i));
             	   System.out.println(febxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" copy to "+febBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" success ?? "+copyResult);
             	   printRptMsg(logps,"",febxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" copy to "+febBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" success ?? "+copyResult);                    	  
             	   if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            //98.04.09 add 主要經營指標明細表
            for(int i=0;i<filename_List_FR001WB.size();i++){
         	    logps.println(logformat.format(nowlog)+(String)filename_List_FR001WB.get(i));
         	    logps.flush();
         	    errMsg +="<br>主要經營指標明細表/各會員別身份放款金額一覽表:"+(String)filename_List_FR001WB.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir_FR001WB")+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
             	  copyResult = Utility.CopyFile(febxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i),febBKxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i));
             	  System.out.println(febxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i)+" copy to "+febBKxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i)+" success ?? "+copyResult);
             	  printRptMsg(logps,"",febxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i)+" copy to "+febBKxlsDir_FR001WB+System.getProperty("file.separator")+(String)filename_List_FR001WB.get(i)+" success ?? "+copyResult);                    	  
             	  if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            
            //98.06.26 add 解釋函令/限制或核准業務/處分書
            for(int i=0;i<filename_List_MCRptAll.size();i++){
         	    logps.println(logformat.format(nowlog)+(String)filename_List_MCRptAll.get(i));
         	    logps.flush();
         	    errMsg +="<br>解釋函令/限制或核准業務/處分書:"+(String)filename_List_MCRptAll.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir_MCRptAll")+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
             	  copyResult = Utility.CopyFile(febxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i),febBKxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i));
             	  System.out.println(febxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i)+" copy to "+febBKxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i)+" success ?? "+copyResult);
             	  printRptMsg(logps,"",febxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i)+" copy to "+febBKxlsDir_MCRptAll+System.getProperty("file.separator")+(String)filename_List_MCRptAll.get(i)+" success ?? "+copyResult);                    	  
             	  if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            //98.06.26 add 舞弊案件
            for(int i=0;i<filename_List_MC012W.size();i++){
         	    logps.println(logformat.format(nowlog)+(String)filename_List_MC012W.get(i));
         	    logps.flush();
         	    errMsg +="<br>舞弊案件:"+(String)filename_List_MC012W.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir_MC012W")+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
             	  copyResult = Utility.CopyFile(febxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i),febBKxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i));
             	  System.out.println(febxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i)+" copy to "+febBKxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i)+" success ?? "+copyResult);
             	  printRptMsg(logps,"",febxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i)+" copy to "+febBKxlsDir_MC012W+System.getProperty("file.separator")+(String)filename_List_MC012W.get(i)+" success ?? "+copyResult);                    	  
             	  if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            
            //98.06.26 add 檢舉書
            for(int i=0;i<filename_List_MC014W.size();i++){
         	    logps.println(logformat.format(nowlog)+(String)filename_List_MC014W.get(i));
         	    logps.flush();
         	    errMsg +="<br>檢舉書:"+(String)filename_List_MC014W.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir_MC014W")+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
             	  copyResult = Utility.CopyFile(febxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i),febBKxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i));
             	  System.out.println(febxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i)+" copy to "+febBKxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i)+" success ?? "+copyResult);
             	  printRptMsg(logps,"",febxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i)+" copy to "+febBKxlsDir_MC014W+System.getProperty("file.separator")+(String)filename_List_MC014W.get(i)+" success ?? "+copyResult);                    	  
             	  if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            
            //98.06.29 add 理監事基本資料
            for(int i=0;i<filename_List_FR055W.size();i++){
         	    logps.println(logformat.format(nowlog)+(String)filename_List_FR055W.get(i));
         	    logps.flush();
         	    errMsg +="<br>理監事基本資料:"+(String)filename_List_FR055W.get(i);
                tmpFile = new File(Utility.getProperties("febxlsDir_FR055W")+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
             	  copyResult = Utility.CopyFile(febxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i),febBKxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i));
             	  System.out.println(febxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i)+" copy to "+febBKxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i)+" success ?? "+copyResult);
             	  printRptMsg(logps,"",febxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i)+" copy to "+febBKxlsDir_FR055W+System.getProperty("file.separator")+(String)filename_List_FR055W.get(i)+" success ?? "+copyResult);                    	  
             	  if(copyResult.equals("0")) tmpFile.delete();
                }
            }
    	}catch(Exception e){
    		System.out.println("RptGenerateALL.moveBKDir_complete Error:"+e+e.getMessage());
    		errMsg += "RptGenerateALL.moveBKDir_complete Error:"+e+e.getMessage();    		
    	}
    	return errMsg;
    }
    
    //各項法定比率
    //102.09.23 add A99.992710建築放款總額
    private static String A02Rpt(String S_YEAR,String S_MONTH,String bank_code,String bank_name,
    						   HSSFWorkbook wb,HSSFSheet sheet){
    	short i = 0;
    	HSSFRow row = null;
		HSSFCell cell = null;
    	int rowNum=0;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
    	String acc_code = "";
    	String amt = "";
    	Properties A02Data = new Properties();
    	String errMsg = "";
    	try {
    	     sqlCmd.append(" select * from a02 where m_year=? and m_month=? and bank_code=? order by acc_code");    	     
    	     paramList.add(S_YEAR);
    	     paramList.add(S_MONTH);
    	     paramList.add(bank_code);
    	     List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
    	     for(i=0;i<dbData.size();i++){
		     	 acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
		     	 amt = (((DataObject)dbData.get(i)).getValue("amt")).toString();
       	         A02Data.setProperty(acc_code,amt);
       	         //System.out.println("acc_code="+acc_code+":amt="+amt);
             }
    	     sqlCmd.setLength(0) ;    	     
    	     sqlCmd.append(" select * from a99 where m_year=? and m_month=? and bank_code=? order by acc_code");
    	     dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
             for(i=0;i<dbData.size();i++){
                 acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
                 amt = (((DataObject)dbData.get(i)).getValue("amt")).toString();
                 A02Data.setProperty(acc_code,amt);
                 //System.out.println("acc_code="+acc_code+":amt="+amt);
             }
    	     row = sheet.getRow(0);
             cell = row.getCell((short) 0);
             //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellValue(bank_name+"各項法定比率");
             row = sheet.getRow(1);
             cell = row.getCell((short) 0);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
             if (dbData.size() == 0) {			 	
			 	 cell.setCellValue("中華民國"+S_YEAR+"年"+S_MONTH+"月無資料存在");				
			 } else { 
			 	 cell.setCellValue("中華民國"+S_YEAR+"年"+S_MONTH+"月");
    	         for(i=3;i<40;i++){    		
	             	row=sheet.getRow(i);	    		
	             	cell=row.getCell((short)1);	    	
	             	//System.out.print((int)cell.getNumericCellValue()+"=");
	             	amt = A02Data.getProperty(String.valueOf((int)cell.getNumericCellValue()));
	             	//System.out.println(amt);
	             	cell=row.getCell((short)3);
	             	cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	             	cell.setCellValue(Double.parseDouble(amt));//98.01.07以數值格式儲存				
	             	if(i==37){
	             		row=sheet.getRow(++i);	    		
		             	cell=row.getCell((short)1);	    
	             		amt = A02Data.getProperty(cell.getStringCellValue());
		             	//System.out.println("99141Y="+amt);
		             	cell=row.getCell((short)3);
		             	cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		             	cell.setCellValue(Double.parseDouble(amt));
	             	}
	             }
			 }
    	} catch (Exception e) {
    		errMsg = "A02Rpt Error:" + e + e.getMessage();
			System.out.println("A02Rpt Error:" + e + e.getMessage());
		}
    	return errMsg;
    }
		
    
    //淨值占風險性資產比率
    private static String A05Rpt(String S_YEAR,String S_MONTH,String bank_type,String bank_code,String bank_name,
			   				   HSSFWorkbook wb,HSSFSheet sheet){	
		String[] table = {"910101", "910102", "910103", "910104", "910105", "910106", "910108", "910109", "910110", "910199", "910201", "910202", "910299", "910300", "910401", "910402", "910403", "910404", "910400", "910500", "91060P"};
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		String ncacno="";
		String unit="1";
		BigDecimal tmpZero=new BigDecimal("0");		
		String errMsg = "";
		try {			
			 ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";				
			 HSSFRow row = null;// 宣告一列
			 HSSFCell cell = null;// 宣告一個儲存格
			 short rowNo = 0;			
			 //95.09.25 fix 若單位為1000時,利率會為0 by 2295
			 sqlCmd.append(" select m_year,m_month,"+ncacno+".acc_code,"+ncacno+".acc_name,");
   			         //+ " round(decode(substr("+ncacno+".acc_code,length("+ncacno+".acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05.amt,0))/"+unit+",0) as amt, "
			 sqlCmd.append(" decode(substr("+ncacno+".acc_code,length("+ncacno+".acc_code)),'P',nvl(a05.amt,0)/1000,round(nvl(a05.amt,0)/"+unit+",0)) as amt,"); 
			 sqlCmd.append(" nvl(a05.amt_name,'') as amt_name ");
			 sqlCmd.append(" from "+ncacno+" left join a05 on "+ncacno+".acc_code=a05.acc_code ");
			 sqlCmd.append(" and "+ncacno+".acc_code like '91%' " + " and   m_year = ? and   m_month = ? and  a05.bank_code = ? ");
			 sqlCmd.append(" where acc_tr_type=?");
			 sqlCmd.append(" order by "+ncacno+".acc_range");
			 paramList.add(S_YEAR);
			 paramList.add(String.valueOf(Integer.parseInt(S_MONTH)) );
			 paramList.add(bank_code);
			 paramList.add("A05");
			 dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, "m_year,m_month,amt");
			 System.out.println("A05Data.size="+dbData.size());
			 // 取出資料存入MAP
			 HashMap h = new HashMap();
			 for (int k = 0; k < dbData.size(); k++) {
			 	DataObject obj = (DataObject) dbData.get(k);
			 	h.put(obj.getValue("acc_code"), obj.getValue("amt").toString());
			 }
			 
			 rowNo = 0;			
			 row = sheet.getRow(rowNo);
			 cell = row.getCell((short) 0);
			 // 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
			 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 cell.setCellValue(S_YEAR+"年"+S_MONTH+"月"+bank_name + "淨值占風險性資產比率");
			 
             rowNo=1;
			 if (dbData.size() == 0) {
			 	row = sheet.getRow(rowNo);
			 	cell = row.getCell((short) 0);
			 	// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			 	cell.setCellValue("無資料存在");				
			 } else {
			 	rowNo = 3;
			 	BigDecimal tmpAmt = null;
			 	BigDecimal tmpAmt_910500 = null;
			 	BigDecimal tmpAmt_910500_910203 = null;
			 	BigDecimal tmpAmt_910202 = null;
			 	BigDecimal tmpAmt_910204 = null;//96.07.05 add 
			 	for (int i = 0; i < table.length; i++) {
			 		//System.out.print("Row =" + rowNo+":"+table[i]);
			 		row = sheet.getRow(rowNo++);
			 		cell = row.getCell((short) 1);
			 		String amt = "0";
			 		if(table[i].equals("910108")){					    
			 		   tmpAmt = new BigDecimal(((String) h.get(table[i])));//910108
			 		   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get(table[i+1]))).negate());//910108-910109
			 		   //amt = Utility.setCommaFormat(tmpAmt.toString());
			 		   amt = tmpAmt.toString();
			 		   i++;
			 		}else if(table[i].equals("910202")){
			 		    tmpAmt_910202 = new BigDecimal(((String) h.get(table[i])));//910202	
			 		   /*95.09.25 fix  
			 		   System.out.println("910203="+(String) h.get("910203"));
			 		   tmpAmt = new BigDecimal(((String) h.get(table[i])));//910202
			 		   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("910203"))).negate());//910202-910203
			 		   //95.08.17 add 910202-910203 < 0 ,則910202-910203等於0
			 		   if(tmpAmt.compareTo(tmpZero) == -1){//910202-910203 < 0 
			 		       tmpAmt = new BigDecimal("0");
			 		   }
			 		   */
			 		   //amt = Utility.setCommaFormat(tmpAmt.toString());
			 		    amt = tmpAmt_910202.toString();//108.01.08 fix 
			 	    }else{
			 	       if(table[i].equals("910500")){
			 	          tmpAmt_910500 = new BigDecimal(((String) h.get(table[i])));//910500  
			 	          tmpAmt_910500_910203 = tmpAmt_910500.add( new BigDecimal(((String) h.get("910203"))));//910500+910203
			 	       }
			 		   //amt = Utility.setCommaFormat((String) h.get(table[i]));    
			 	       amt = (String) h.get(table[i]);
			 		}
			 		//System.out.println(":amt =" + amt);
			 		//if (!amt.equals("0")) {			 			
			 			cell.setCellValue(Double.parseDouble(amt));		
			 		//}
			 	}
			 	
			 	/*96.07.05 if【910204】< (【910500】* 1.25% )
			 	              Temp_910202 = 【910202】-【910203】
			 	           Else (為910204 = 910500 * 1.25%)
			 	              Temp_910202 = 【910204】*/   
			 	/*108.01.08 取消因910204已為min([910202-910203],910500*1.75%)
			 	row = sheet.getRow((short)14);//910202
			 	cell = row.getCell((short) 1);
			 	
			 	String amt = "0";
			 	BigDecimal tmp125=new BigDecimal("0.0125");	
			 	tmpAmt_910204 = h.get("910204") == null ? new BigDecimal("0") : new BigDecimal(((String) h.get("910204")));//910204
			 	//910500*1.25%->取到整數,小數點下以無條件捨去
			 	tmpAmt_910500 = new BigDecimal(tmpAmt_910500.multiply(tmp125).intValue());
			     //System.out.println("910500*1.25%="+tmpAmt_910500);	
			     
			     if(tmpAmt_910204.compareTo(tmpAmt_910500) == -1){//if【910204】< (【910500】* 1.25%
			        //System.out.println("910203="+(String) h.get("910203"));
			 	   tmpAmt = new BigDecimal(((String) h.get("910202")));//910202
			 	   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("910203"))).negate());//910202-910203				   
			     }else{//為910204 = 910500 * 1.25%
			        tmpAmt = new BigDecimal(((String) h.get("910204")));//910204
			     }
			     //add 910202-910203 or 910204 < 0 ,則設為0
			 	if(tmpAmt.compareTo(tmpZero) == -1){//910202-910203 or 910204 < 0 
			 	   tmpAmt = new BigDecimal("0");
			 	}
			 	//amt = Utility.setCommaFormat(tmpAmt.toString());				
			 	amt = tmpAmt.toString();			 	
			 	cell.setCellValue(Double.parseDouble(amt));		
			 	*/
			 	//==============================================================================
			 	String amt320200="0";
			 	sqlCmd.delete(0,sqlCmd.length());
			 	paramList = new ArrayList();
			 	sqlCmd.append(" select m_year,m_month,acc_code,round(nvl(amt,0)/?,0) as amt");  
			 	sqlCmd.append(" from a01  ");
			 	sqlCmd.append(" where m_year = ?");  
			 	sqlCmd.append(" and   m_month = ?"); 
			 	sqlCmd.append(" and   bank_code = ?");
			 	sqlCmd.append(" and   acc_code='320200'");
			 	paramList.add(unit);
			 	paramList.add(S_YEAR);
			 	paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			 	paramList.add(bank_code);
			 	
			 	dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, "m_year,m_month,amt");
			 	if(dbData != null && dbData.size() != 0){
			 	    amt320200 = ((DataObject)dbData.get(0)).getValue("amt").toString();
			 	}
			 	//System.out.println("Row =" + rowNo);
			 	row = sheet.getRow(rowNo++);
			 	cell = row.getCell((short) 0);
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 	//96.01.22 fix 備註加上單位別 
			 	cell.setCellValue("註1:「(7)累積盈虧」有加計「上期損益: "+amt320200+"元」及扣除「備抵呆帳、損失準備");
			 	row = sheet.getRow(rowNo++);
			 	cell = row.getCell((short) 0);
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 	cell.setCellValue("    及營業準備提列不足  "+(String) h.get("910109")+"元」");
			 	//95.08.17 add 2:備抵呆帳、損失準備及營業準備(第一類資本)-910202 ==============================================================
			 	row = sheet.getRow(rowNo++);
			 	cell = row.getCell((short) 0);
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 	cell.setCellValue("  2:帳列備抵呆帳、損失準備及營業準備 "+(String) h.get("910202")+"元");
			 	//========================================================================================================================
			 	row = sheet.getRow(rowNo++);
			 	cell = row.getCell((short) 0);
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 	cell.setCellValue("  3:特定損失所提列備抵呆帳、損失準備及營業準備  "+(String) h.get("910203")+"元");
			 	row = sheet.getRow(rowNo++);
			 	cell = row.getCell((short) 0);
			 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			 	cell.setCellValue("  4:原風險性資產總額(未扣除特定損失)   "+tmpAmt_910500_910203.toString()+"元");//910500+910203
			 }//end of 有資料存在
			
		} catch (Exception e) {
			errMsg = "A05Rpt Error:" + e + e.getMessage();
			System.out.println("A05Rpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}
    //損益表
    //102.07.15 add 103/01以後,漁會.套用新表格 by 2295
    private static String A01Rpt_rpt2(String S_YEAR,String S_MONTH,String bank_type,String bank_code,String bank_name,
			   						HSSFWorkbook wb,HSSFSheet sheet){		
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		String unit="1";		
		FileInputStream finput = null;
		String ncacno="ncacno";
		int rowNum=0;		
		String errMsg = "";
		try{	
			ncacno = bank_type.equals("6")?"ncacno_rule":"ncacno_7";//96.12.19 add 97/01以後,套用新表格(增加/異動科目代號)
			//add 103/01以後,漁會.套用新表格(增加/異動科目代號) 
			if(bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){ 
			   ncacno = "ncacno_7_rule";
			}

			HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  			  		
	  		sqlCmd.append("select A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code, ");			
	  		sqlCmd.append(" round(amt/?,0) as amt");
	  	    sqlCmd.append(" from A01 LEFT JOIN "+ncacno +" ON A01.acc_code = "+ncacno+".acc_code");
	  		sqlCmd.append(" where A01.m_year=?");
	  		sqlCmd.append("   and A01.m_month=?");
	  		sqlCmd.append("   and A01.bank_code=?");
	  		sqlCmd.append("   and "+ncacno+".acc_div='02'");
	  		sqlCmd.append(" order by "+ncacno+".acc_range");
	  		paramList.add(unit);
	  		paramList.add(S_YEAR);
	  		paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
	  		paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,b_amt,month_amt");	  	         

  			for(i=0;i<dbData.size();i++){
  			    amt = "0";  			    
  				acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
  				amt = (((DataObject)dbData.get(i)).getValue("amt")).toString();
  				//System.out.println("acc_code="+acc_code+":amt="+amt);
  				A01Data.setProperty(acc_code,amt);
	       	}
  			
  			row=sheet.getRow(0);
  			cell=row.getCell((short)0);	
  		    //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	 		cell.setCellValue(bank_name+"損益表");
  	 		row=sheet.getRow(1);
  	 		cell=row.getCell((short)0);	       	
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
  	        if(dbData.size() == 0){
   	       	    cell.setCellValue("中華民國"+S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	 	}else{  	 	    
	  	 		cell.setCellValue("中華民國"+S_YEAR +"年" +S_MONTH +"月");
	  	 		//以巢狀迴圈讀取所有儲存格資料 
	  	 		System.out.println("total row ="+sheet.getLastRowNum());	  
	  	 		rowNum=bank_type.equals("6")?30:31;
	  	 		for(i=4;i<rowNum;i++){
	  	 		    row=sheet.getRow(i);
	  	 		    cell=row.getCell((short)3);
	  	 		    if((int)cell.getNumericCellValue()>0){
    		          	//System.out.print((int)cell.getNumericCellValue()+"=");		          		
    		          	amt = Utility.getTrimString(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
    		          	//System.out.println(amt);	
    		          	cell=row.getCell((short)4);
    		            cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
    		          	cell.setCellValue(Double.parseDouble(amt));
	  	 		    }
		          	cell=row.getCell((short)8);
		          	if((int)cell.getNumericCellValue()>0){
    	          	    //System.out.print((int)cell.getNumericCellValue()+"="); 
    	          	    amt = Utility.getTrimString(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
    	          	    //System.out.println(amt);
    	          	    cell=row.getCell((short)9);		          	
    	          	    cell.setCellValue(Double.parseDouble(amt));
		          	}
	  	 		}
	  	 	}
		}catch(Exception e){
			errMsg = "A01Rpt_rpt2 Error:"+e+e.getMessage();
			System.out.println("A01Rpt_rpt2 Error:"+e+e.getMessage());
		}
		return errMsg;
	}
    //資產負債表
    //102.07.15 add 103/01以後,漁會.套用新表格 by 2295
    private static String A01Rpt_rpt1(String S_YEAR,String S_MONTH,String bank_type,String bank_code,String bank_name,
								    HSSFWorkbook wb,HSSFSheet sheet){	
		List dbData = null;
		List dbData_other = null;
		StringBuffer sqlCmd = new StringBuffer();
		StringBuffer sqlCmd_other = new StringBuffer();
		List paramList = new ArrayList();	
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		String unit="1";
		FileInputStream finput = null;
		String ncacno="ncacno";
		int rowNum=0;	
		String errMsg = "";
		try{
			ncacno = bank_type.equals("6")?"ncacno_rule":"ncacno_7";//96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) 	
			//add 103/01以後,漁會.套用新表格(增加/異動科目代號) 
            if(bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){ 
               ncacno = "ncacno_7_rule";
            }
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		//99.10.12 add 查詢年度100年以前.縣市別不同===============================
		    String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
		    String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
		    //===================================================================== 
		  
	  		short i=0;
	  		short y=0;  
	  		
	  		sqlCmd.append(" select A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code, ");
	  		sqlCmd.append(" round(amt/?,0) as amt");
	  	    sqlCmd.append(" from A01 LEFT JOIN "+ncacno+" ON A01.acc_code = "+ncacno+".acc_code");
	  		sqlCmd.append(" where A01.m_year=?");
	  		sqlCmd.append("   and A01.m_month=?");
	  	    sqlCmd.append("   and A01.bank_code=?");
	  		sqlCmd.append("   and "+ncacno+".acc_div='01'");
	  		sqlCmd.append(" order by "+ncacno+".acc_range");
	  		paramList.add(unit);
	  		paramList.add(S_YEAR);
	  		paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
	  		paramList.add(bank_code);
			
	  		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");	
	  		int dbDatasize=dbData.size();
            System.out.println("a01.size()="+dbData.size());
	  		//94.12.06 add 逾期放款金額/逾放比率/存放比率===============================================
            //103.02.12 fix 逾放金額field_over/逾放比率field_over_rate/存放比率field_dc_rate.從A01_operation抓
            paramList = new ArrayList();	
            sqlCmd_other.append(" select bank_code,");
            sqlCmd_other.append(" round(sum(decode(a01_operation.acc_code,'field_over',amt,0))/?,0) as field_over,"); 
            sqlCmd_other.append(" sum(decode(a01_operation.acc_code,'field_over_rate',amt,0)) as field_over_rate,"); 
            sqlCmd_other.append(" sum(decode(a01_operation.acc_code,'field_dc_rate',amt,0)) as field_dc_rate");
            sqlCmd_other.append(" from a01_operation");  
            sqlCmd_other.append(" where m_year = ? and m_month=?");  
            sqlCmd_other.append(" and bank_code= ?");
            sqlCmd_other.append(" group by bank_code"); 
            paramList.add(unit);
            paramList.add(S_YEAR);
            paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
            paramList.add(bank_code);
            /*
	  		sqlCmd_other.append(" select round(field_OVER / "+ unit+ ",0)    as field_OVER,"); //逾期放款金額                     
	  		sqlCmd_other.append(" decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,"); //逾放比率
	  		sqlCmd_other.append(" decode(a01.fieldI_Y,0,0, round((a01.fieldI_XA + ");
	  		sqlCmd_other.append(" decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,");
	  		sqlCmd_other.append("       (a01.fieldI_XB1 - a01.fieldI_XB2))        +");
	  		sqlCmd_other.append(" decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,");
	  		sqlCmd_other.append("	   (a01.fieldI_XC1 - a01.fieldI_XC2))           +");
	  		sqlCmd_other.append(" decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,");
	  		sqlCmd_other.append("	   (a01.fieldI_XD1 - a01.fieldI_XD2))           +");
	  		sqlCmd_other.append(" decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,");
	  		sqlCmd_other.append("	    (a01.fieldI_XE1 - a01.fieldI_XE2))           -");
	  		sqlCmd_other.append(" decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,"); //95.01.24 fix 
	  		sqlCmd_other.append("	    (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))");
	  		sqlCmd_other.append(" )");
	  		sqlCmd_other.append(" /    a01.fieldI_Y * 100,2))        as     field_DC_RATE"); //存放比率
	  		sqlCmd_other.append(" from (");
	  		sqlCmd_other.append(" select");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)        as field_OVER,");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0)        as field_DEBIT," );                   
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");
	  		sqlCmd_other.append(" round(sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt, "); 
	  		sqlCmd_other.append("                                                       '120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0) "); 
	  		sqlCmd_other.append("                                                   ,'7',decode(a01.acc_code,'120101',amt,'120102',amt, "); 
	  		sqlCmd_other.append("                                                      '120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)), ");
	  		sqlCmd_other.append("                            '103',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,");  
	  		sqlCmd_other.append("                                 '120302',amt,'120700',amt,'150200',amt,0),0)");
	  		sqlCmd_other.append(" ) /1,0)     as fieldI_XA,");  
	  		sqlCmd_other.append(" round(sum(decode(year_type,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)),");
	  		sqlCmd_other.append("                            '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0)");
	  		sqlCmd_other.append(" ) /1,0)     as fieldI_XB1,");  
	  		sqlCmd_other.append(" sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)),");
	  		sqlCmd_other.append("                      '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0)");
	  		sqlCmd_other.append(" )  as fieldI_XB2,"); 	  		               
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0) as fieldI_XC1,");
	  		sqlCmd_other.append(" round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),");  
	  		sqlCmd_other.append("                                                     '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)),");
	  		sqlCmd_other.append("                              '103', decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0)  ");            
	  		sqlCmd_other.append(" ) /1,0)     as fieldI_XC2, ");	          
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0) as fieldI_XD1,");
	  		sqlCmd_other.append(" round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0),");
	  		sqlCmd_other.append("                                                    '7',decode(a01.acc_code,'240300',amt,0)),");
	  		sqlCmd_other.append("                             '103',decode(a01.acc_code,'240200',amt,0),0)");                                  
	  		sqlCmd_other.append(" ) /1,0)   as fieldI_XD2, ");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0) as fieldI_XE1,");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0) as fieldI_XE2,");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1,");
	  		sqlCmd_other.append(" round(sum( decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0),");
	  		sqlCmd_other.append("                             '103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0)");//95.01.24 add 
	  		sqlCmd_other.append(" ) /1,0)  as fieldI_XF3, ");	  		
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0) as fieldI_XF2,");                    
	  		sqlCmd_other.append(" round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,");
	  		sqlCmd_other.append("           					 '220300',amt,'220400',amt,");
	  		sqlCmd_other.append("					             '220500',amt,'220600',amt,");
	  		sqlCmd_other.append("				                 '220700',amt,'220800',amt,");
	  		sqlCmd_other.append("				                 '220900',amt,'221000',amt,0))-");
	  		sqlCmd_other.append(" round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y"); 
	  		sqlCmd_other.append(" from (select  (CASE WHEN (a01.m_year <= 102) THEN '102'");
	  		sqlCmd_other.append("                WHEN (a01.m_year > 102) THEN '103'");
	  		sqlCmd_other.append("                ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt");
	  		sqlCmd_other.append("       from a01 )a01, (select * from bn01 where m_year=?)bn01");
	  		sqlCmd_other.append(" where (a01.m_year  = ?"); 
	  		sqlCmd_other.append(" and    a01.m_month = ?) ");
	  		paramList.add(wlx01_m_year);
	  		paramList.add(S_YEAR);
	  		paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
	  		if (!bank_code.equals("ALL")) {
	  		    sqlCmd_other.append(" and   (a01.bank_code  = ?)");
	  		    paramList.add(bank_code);
	  		}
        	sqlCmd_other.append(" and   (a01.bank_code=bn01.bank_no  and bn01.bank_type=?)");//農會
        	paramList.add(bank_type);        	
        	sqlCmd_other.append(" ) a01");
        	*/
        	dbData_other = DBManager.QueryDB_SQLParam(sqlCmd_other.toString(),paramList,"field_over,field_over_rate,field_dc_rate");
        	System.out.println("dbData_other.size()="+dbData.size());
	  		//=================================================================================================
  			for(i=0;i<dbData.size();i++){
  				acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
  				amt = (((DataObject)dbData.get(i)).getValue("amt")).toString();
	       	    A01Data.setProperty(acc_code,amt);
	       	    //System.out.println("acc_code="+acc_code+":amt="+amt);
	        }
  			row=sheet.getRow(0);
  			cell=row.getCell((short)0);	
  		    //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	 		cell.setCellValue(S_YEAR+"年"+S_MONTH+"月"+bank_name+"資產負債表");
  	 		row=sheet.getRow(1);
  	 		cell=row.getCell((short)0);	       	
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
  	        if(dbDatasize == 0){
   	   	       	cell.setCellValue("無資料存在");
	  	    }else{
	  		  	//以巢狀迴圈讀取所有儲存格資料 
	  		  	System.out.println("total row ="+sheet.getLastRowNum());
	  		    //96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) 
	  		  	rowNum=bank_type.equals("6")?110:107;
	  		    //add 103/01以後,漁會套用新表格(增加/異動科目代號) 
	  		    if(bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){ 
	  		       rowNum = 111;
	            }
	  		    System.out.println("rowNum="+rowNum);
	  		  	for(i=4;i<rowNum;i++){    		
	    	    	row=sheet.getRow(i);	    		
	    	    	cell=row.getCell((short)3);
	    	    	if((int)cell.getNumericCellValue()>0){
    	    	    	//System.out.print((int)cell.getNumericCellValue()+"=");
    	    	    	amt = Utility.getTrimString(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
    	    	    	//System.out.println(amt);
    	    	    	cell=row.getCell((short)4);
    	    	    	cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
    	    	    	cell.setCellValue(Double.parseDouble(amt));
	    	    	}
	    	    	cell=row.getCell((short)8);
	    	    	if((int)cell.getNumericCellValue()>0){
    	    	    	//System.out.print((int)cell.getNumericCellValue()+"=");
    	    	    	amt = Utility.getTrimString(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));	       		
    	    	    	cell=row.getCell((short)9);
    	    	    	cell.setCellValue(Double.parseDouble(amt));		 
    	    	    	//System.out.println(amt);
	    	    	}
	  		  	}
	  		  	
	  		    //94.12.06 add 逾期放款金額/逾放比率/存放比率
                if(dbData_other != null && dbData_other.size() != 0){
                    //逾期放款金額
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt = (((DataObject) dbData_other.get(0)).getValue("field_over")).toString();                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(amt));		
                    //逾放比率
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt =(((DataObject) dbData_other.get(0)).getValue("field_over_rate")).toString();                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(amt));		
                    //存放比率
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt = (((DataObject) dbData_other.get(0)).getValue("field_dc_rate")).toString();                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(amt));		
                }else{
                   System.out.println("dbData_other is null"); 
                }
	  	    }//end of 有資料
		}catch(Exception e){
			errMsg = "A01Rpt_rpt1 Error:"+e+e.getMessage();
			System.out.println("A01Rpt_rpt1 Error:"+e+e.getMessage());			
		}
		return errMsg;
	}	

    //逾期放款統計表
    private static String A06Rpt(String S_YEAR,String S_MONTH,String bank_type,String bank_code,String bank_name,
							   HSSFWorkbook wb,HSSFSheet sheet){
		
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();	
		StringBuffer a01table = new StringBuffer();
		List paramList = new ArrayList();	
		List paramList_a01table = new ArrayList();	
		List A06Data = new LinkedList();		
		String acc_code = "";
		String amt = "";		
		String YEAR = "";
		String unit="1";
		String errMsg = "";
		try{
			
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;
	  		String ncacno = (bank_type.equals("6"))?"ncacno":"ncacno_7";
	  		/*
	  		  單一信用部
	  		  select A06.m_year, A06.m_month,
	  		  	     ncacno.acc_range,A06.acc_code,ncacno.acc_name,
	  		  	     round(amt_3month/1,0) as amt_3month, --未滿3個月
	  		  	     round(amt_6month/1,0) as amt_6month, --3個月~6個月
	  		  	     round(amt_1year/1,0) as amt_1year, --6個月~1年
	  		  	     round(amt_2year/1,0) as amt_2year, --1年~2年
	  		  	     round(amt_over2year/1,0) as amt_over2year, --2年以上
	  		  	     round(amt_total/1,0) as amt_total -- 逾放合計
	  		  	     ,round(A01.amt/1,0) as amt -- a01放款合計
				     ,round(round(amt_total/1,0)/ decode(round(A01.amt/1,0),0,1,round(A01.amt/1,0)) * 100,3) as a01_per --逾放比率 
              from A06 LEFT JOIN ncacno ON A06.acc_code = ncacno.acc_code and ncacno.acc_div='08' 
              	       LEFT JOIN  (select m_year,m_month,bank_code,
					               decode(acc_code,'990000','970000',acc_code) as acc_code
                                   ,amt from a01
								   )A01 ON A01.bank_code = A06.bank_code 
              		   			        and A01.m_year = A06.m_year and A01.m_month = A06.m_month 
									    and A01.acc_code = A06.acc_code              	              
              where A06.m_year=94 and A06.m_month=6
              and A06.bank_code = '6030016'
              order by ncacno.acc_range
                漁會的a01取代成
              (select 94 as m_year,6 as m_month,'5030019' as bank_code,acc_code,amt
			   from a01
			   where acc_code not in ('120700','120900')
			   and m_year=94 and m_month=6 and bank_code='5030019'
			   union
			   select 94 as m_year,6 as m_month,'5030019' as bank_code,'120700',sum(amt) amt from a01  
			   where acc_code in('120700','120900') 
			   and m_year=94 and m_month=6 and bank_code='5030019') a01  
             
                總表               
              select a06.m_year,a06.m_month,a06.acc_range,a06.acc_code,a06.acc_name,a06.amt_3month,a06.AMT_6MONTH,
              		 a06.AMT_1YEAR,a06.AMT_2YEAR,a06.AMT_OVER2YEAR,a06.AMT_TOTAL,a01.amt
              		 ,round(a06.AMT_TOTAL/ decode(a01.amt,0,1,a01.amt) * 100,3) as a01_per --逾放比率
              from	   
              	    (select A06.m_year, A06.m_month, ncacno.acc_range, a06.acc_code,ncacno.acc_name 
              	            ,round(sum(amt_3month)/1,0) as amt_3month --未滿3個月
              	            ,round(sum(amt_6month)/1,0) as amt_6month --3個月~6個月
              	            ,round(sum(amt_1year)/1,0) as amt_1year --6個月~1年
              	            ,round(sum(amt_2year)/1,0) as amt_2year --1年~2年
              	            ,round(sum(amt_over2year)/1,0) as amt_over2year --2年以上
              	            ,round(sum(amt_total)/1,0) as amt_total -- 逾放合計
                     from A06 LEFT JOIN ncacno ON A06.acc_code = ncacno.acc_code
                     	  ,bn01 
                     where A06.m_year=94 
                       and A06.m_month=6
                       and ncacno.acc_div='08'
                       and A06.bank_code = bn01.bank_no 
                       and bn01.bank_type='6'
                     group by A06.m_year, A06.m_month, ncacno.acc_range, A06.acc_code,ncacno.acc_name
                     order by A06.m_year, A06.m_month, ncacno.acc_range, A06.acc_code,ncacno.acc_name )a06
                     left join
                     (select A01.m_year, A01.m_month, ncacno.acc_range, 
                     	     decode(a01.acc_code,'990000','970000',a01.acc_code) as acc_code,ncacno.acc_name 
                     	     round(sum(amt)/1,0) as amt -- a01放款合計
                      from A01 LEFT JOIN ncacno ON A01.acc_code = ncacno.acc_code
                      	   ,bn01 
                      where A01.m_year=94 
                        and A01.m_month=6
                        and (ncacno.acc_div='08' or ncacno.acc_code='990000')
                        and A01.bank_code = bn01.bank_no 
                        and bn01.bank_type='6'
                      group by A01.m_year, A01.m_month, ncacno.acc_range, A01.acc_code,ncacno.acc_name
                      order by A01.m_year, A01.m_month, ncacno.acc_range, A01.acc_code,ncacno.acc_name)a01
                      ON A01.m_year = A06.m_year 
                      and A01.m_month = A06.m_month 
					  and A01.acc_code = A06.acc_code  
                   order by a06.m_year,a06.m_month,a06.acc_range			 
              

                漁會總表A01取代成
                 (select m_year,m_month,bank_code,acc_code,amt
				  from a01
				  where acc_code not in ('120700','120900') and m_year=94 and m_month=6
				  union
				  select m_year,m_month,bank_code,'120700' as acc_code,sum(amt) amt from a01  
				  where acc_code in('120700','120900') and m_year=94 and m_month=6
				  group by m_year,m_month,bank_code
				  order by m_year,m_month,bank_code ) a01      
	  		 */
	  		
	  		
	  		if(!bank_code.equals("ALL")){//單一信用部
	  		    if(bank_type.equals("6")){//農會		  		  
		  		   a01table.append(" (select "+S_YEAR+" as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,acc_code,amt ");
		  		   a01table.append("  from a01 ");
		  		   a01table.append(" where acc_code not in ('120000','120800','150300') ");
		  		   a01table.append(" and m_year=? and m_month=? and bank_code=? ");
		  		   a01table.append(" union ");
							//98.07.27 add A01放款總額(120000+120800+150300)
		  		   a01table.append(" select "+S_YEAR+"  as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,'120000',sum(amt) amt from a01 ");  
		  		   a01table.append(" where acc_code in('120000','120800','150300')"); 
		  		   a01table.append(" and m_year=? and m_month=? and bank_code=?) a01 ");
		  		   paramList_a01table.add(S_YEAR);
		  		   paramList_a01table.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  		   paramList_a01table.add(bank_code);
		  		   paramList_a01table.add(S_YEAR);
		  		   paramList_a01table.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  		   paramList_a01table.add(bank_code);

		  		}else{//漁會
		  		   a01table.append(" (select "+S_YEAR+" as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,acc_code,amt ");
		  		   a01table.append("  from a01 ");
		  		   a01table.append(" where acc_code not in ('120700','120900','120000') ");
		  		   a01table.append(" and m_year=? and m_month=? and bank_code=?");
		  		   a01table.append(" union ");
		  		   a01table.append(" select "+S_YEAR+"  as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,'120700',sum(amt) amt from a01 ");  
		  		   a01table.append(" where acc_code in('120700','120900')" );
		  		   a01table.append(" and m_year=? and m_month=? and bank_code=? ");
		  		   a01table.append(" union ");
						    //98.07.27 add A01放款總額(120000+120800+150300)
		  		   a01table.append(" select "+S_YEAR+"  as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,'120000',sum(amt) amt from a01 ");  
		  		   a01table.append(" where acc_code in('120000','120800','150300')" );
		  		   a01table.append(" and m_year=? and m_month=? and bank_code=?) a01 "); 
		  		   paramList_a01table.add(S_YEAR);
		  		   paramList_a01table.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  		   paramList_a01table.add(bank_code);
		  		   paramList_a01table.add(S_YEAR);
		  		   paramList_a01table.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  		   paramList_a01table.add(bank_code);
		  		   paramList_a01table.add(S_YEAR);
		  		   paramList_a01table.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  		   paramList_a01table.add(bank_code);

		  		}
	  		    sqlCmd.append(" select A06.m_year, A06.m_month, ");
	  		    sqlCmd.append(" "+ncacno+".acc_range,A06.acc_code, "+ncacno+".acc_name,");
	  		    sqlCmd.append(" round(amt_3month/?,0) as amt_3month, "); //未滿3個月
	  		  	sqlCmd.append(" round(amt_6month/?,0) as amt_6month, "); //3個月~6個月
	  		  	sqlCmd.append(" round(amt_1year/?,0) as amt_1year, "); //6個月~1年
	  		  	sqlCmd.append(" round(amt_2year/?,0) as amt_2year, "); //1年~2年
	  		  	sqlCmd.append(" round(amt_over2year/?,0) as amt_over2year, "); //2年以上
	  		  	sqlCmd.append(" round(amt_total/?,0) as amt_total, "); //逾放合計
	  		  	sqlCmd.append(" round(A01.amt/?,0) as amt, "); //a01放款合計
	  		  	sqlCmd.append(" round(round(amt_total/?,0)/ decode(round(A01.amt/?,0),0,1,round(A01.amt/?,0)) * 100,3) as a01_per "); //逾放比率 
	  		    sqlCmd.append(" from A06 LEFT JOIN  "+ncacno+" ON A06.acc_code =  "+ncacno+".acc_code and  "+ncacno+".acc_div='08' "); 
	  		    sqlCmd.append(" LEFT JOIN  (select m_year,m_month,bank_code, ");
	  		    sqlCmd.append("        			  decode(acc_code,'120000','970000',acc_code) as acc_code,amt ");
	  		    sqlCmd.append("			   from "+a01table);                       
	  		    sqlCmd.append("			   )A01 ON A01.bank_code = A06.bank_code "); 
	  		    sqlCmd.append("   			    and A01.m_year = A06.m_year and A01.m_month = A06.m_month "); 
	  		    sqlCmd.append("			    	and A01.acc_code = A06.acc_code ");              	              
	  		    sqlCmd.append(" where  A06.m_year=? and A06.m_month= ?");
	  		    sqlCmd.append(" and A06.bank_code = ?");
	  		    sqlCmd.append(" order by  "+ncacno+".acc_range ");
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    paramList.add(unit);
	  		    for(int iparam=0;iparam<paramList_a01table.size();iparam++){
            	  paramList.add(paramList_a01table.get(iparam));
                }
	  		    paramList.add(S_YEAR);
	  		    paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
	  		    paramList.add(bank_code);
	  		}
	  		
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total,amt,a01_per");	  	         
			List amt_List =null;
  			for(i=0;i<dbData.size();i++){  			      	
  			    amt_List = new LinkedList();
  				acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
  				amt_List.add(acc_code);
  				amt_List.add((String)((DataObject)dbData.get(i)).getValue("acc_name"));  				
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_3month")).toString());
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_6month")).toString());  				
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_1year")).toString());  				
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_2year")).toString());  				
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_over2year")).toString());  				
  				amt_List.add((((DataObject)dbData.get(i)).getValue("amt_total")).toString());
  				if(((DataObject)dbData.get(i)).getValue("amt") != null){
  				   amt_List.add((((DataObject)dbData.get(i)).getValue("amt")).toString());
  				}else{
  				   amt_List.add("0"); 
  				}
  				if(((DataObject)dbData.get(i)).getValue("a01_per") != null){  				 
  				    if(((((DataObject)dbData.get(i)).getValue("a01_per")).toString()).indexOf(".") != -1){
  				       //取到小數第3位
  				       amt_List.add((((DataObject)dbData.get(i)).getValue("a01_per")).toString());  				       
  				    }else{
  				       amt_List.add("0");				       
  				    }   
  				}else{
  				   amt_List.add("0");
  				}  				
  				//System.out.println(amt_List);
  				A06Data.add(amt_List);	         		       	    	       	    
	       	}
  			
  			row=sheet.getRow(0);
  			cell=row.getCell((short)0);	
  		    //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	 		cell.setCellValue(bank_name+"逾期放款統計表");
  	 		row=sheet.getRow(1);
  	 		cell=row.getCell((short)0);	       	
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	        if(dbData.size() == 0){	
   	       	    cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月無資料存在");
	  	 	}else{ 	 		
   	          	cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月");
	  	 		
	  	 		for(i=3;i<19;i++){
	  	 		    if(i==18){//合計
	  	 		       row=sheet.getRow(i+4);
		          	}else{
	  	 		       row=sheet.getRow(i);
		          	}   
	  	 		    //System.out.println("i="+i);
	  	 		    amt_List = (List)A06Data.get(i-3);
	  	 		    acc_code = ((String)amt_List.get(0)).trim();//科目代號
	  	 		    //System.out.println(amt_List);
	  	 		    //System.out.println("acc_code="+acc_code);
	  	 		    cell=row.getCell((short)0);
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	  	 		    if(acc_code.equals("120700") && bank_type.equals("7")){
	  	 		       cell.setCellValue((String)amt_List.get(1)+"(內含120900)");//項目
	  	 		    }else{    
	  	 		       cell.setCellValue((String)amt_List.get(1));//項目
	  	 		    }
	  	 		    //System.out.print(":"+(String)amt_List.get(1));
	  	 		    cell=row.getCell((short)1);//未滿3個月
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(2);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.print(":"+(String)amt_List.get(2));
		          	cell=row.getCell((short)2);//3個月~6個月
		            cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(3);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.print(":"+(String)amt_List.get(3));
		          	cell=row.getCell((short)3);//6個月~1年
	                cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(4);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.print(":"+(String)amt_List.get(4));
		          	cell=row.getCell((short)4);//1年~2年
		            cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(5);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.print(":"+(String)amt_List.get(5));
		          	cell=row.getCell((short)5);//2年以上
		            cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(6);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.print(":"+(String)amt_List.get(6));
		          	cell=row.getCell((short)6);//逾放合計
		            cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = (String)amt_List.get(7);
		          	cell.setCellValue(Double.parseDouble(amt));		
		          	//System.out.println(":"+(String)amt_List.get(7));
		          	
		          	if(!acc_code.equals("960500")){
		          	   cell=row.getCell((short)7);//A01放款合計
		          	   cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	   amt = (String)amt_List.get(8);
		          	   cell.setCellValue(Double.parseDouble(amt));		
		          	   //System.out.print(":"+(String)amt_List.get(8));
		          	} 
	  	 		}
	  	 	}//end of monthExist
	        
		}catch(Exception e){
			errMsg = "A06Rpt Error:"+e+e.getMessage();
			System.out.println("A06Rpt Error:"+e+e.getMessage());
		}	
		return errMsg;
	}	
    
    //應予評估資產彙總表
    private static String A10Rpt(String S_YEAR,String S_MONTH,String bank_type,String bank_code,String bank_name,
							   HSSFWorkbook wb,HSSFSheet sheet){
	    String errMsg = "";
	    List dbData = null;
	    StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
	    int rowNum=0;
	    String unit = "1";
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	   
	    String data_year="";
	    String data_month="";
	    String loan2_amt="";//放款-列二類
	  	String loan3_amt="";//放款-列三類     
	  	String loan4_amt="";//放款-列四類   
	  	String loan_sum="";//放款-合計  
	  	
	  	String invest2_amt="";//投資-列二類   
	  	String invest3_amt="";//投資-列三類   
	  	String invest4_amt="";//投資-列四類   
	  	String invest_sum="";//投資-合計
	  	
	  	String other2_amt="";//其他-列二類    
	  	String other3_amt="";//其他-列三類    
	  	String other4_amt="";//其他-列四類    
	  	String other_sum="";//其他-合計
	  	
	  	String type2_sum="";//列二類合計
	  	String type3_sum="";//列三類合計
	  	String type4_sum="";//列三類合計
	  	String type_sum="";//合計
	  	
	    try {

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格

	      short i = 0;
	      short y = 0;
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      //99.10.12 add 查詢年度100年以前.縣市別不同===============================
		  String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
		  String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
		  //===================================================================== 
	      
	   	  /*
	      select m_year,m_month,bank_name,loan2_amt,loan3_amt,loan4_amt,loan2_amt+loan3_amt+loan4_amt as loan_sum,
	      		 invest2_amt,invest3_amt,invest4_amt,invest2_amt+invest3_amt+invest4_amt as invest_sum,
	      		 other2_amt,other3_amt,other4_amt,other2_amt+other3_amt+other4_amt as other_sum,
	      		 loan2_amt+invest2_amt+other2_amt as type2_sum,
	      		 loan3_amt+invest3_amt+other3_amt as type3_sum,
	      		 loan4_amt+invest4_amt+other4_amt as type4_sum,
	      		 loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt as type_sum
          from a10 left join bn01 on a10.bank_code = bn01.bank_no
          where m_year=97
          and m_month=6
          and bank_code='6030016'  
          */
	      sqlCmd.append(" select a10.m_year,m_month,bank_name,");
	      sqlCmd.append(" round(loan2_amt/?,0) as loan2_amt,");
	      sqlCmd.append(" round(loan3_amt/?,0) as loan3_amt,");
	      sqlCmd.append(" round(loan4_amt/?,0) as loan4_amt,");
	      sqlCmd.append(" round((loan2_amt+loan3_amt+loan4_amt)/?,0) as loan_sum,");
	      sqlCmd.append(" round(invest2_amt/?,0) as invest2_amt,");
	      sqlCmd.append(" round(invest3_amt/?,0) as invest3_amt,");
	      sqlCmd.append(" round(invest4_amt/?,0) as invest4_amt,");
	      sqlCmd.append(" round((invest2_amt+invest3_amt+invest4_amt)/?,0) as invest_sum,");
	      sqlCmd.append(" round(other2_amt/?,0) as other2_amt,");
	      sqlCmd.append(" round(other3_amt/?,0) as other3_amt,");
	      sqlCmd.append(" round(other4_amt/?,0) as other4_amt,");
	      sqlCmd.append(" round((other2_amt+other3_amt+other4_amt)/?,0) as other_sum,");
	      sqlCmd.append(" round((loan2_amt+invest2_amt+other2_amt)/?,0) as type2_sum,");
	      sqlCmd.append(" round((loan3_amt+invest3_amt+other3_amt)/?,0) as type3_sum,");
	      sqlCmd.append(" round((loan4_amt+invest4_amt+other4_amt)/?,0) as type4_sum,");
	      sqlCmd.append(" round((loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt)/?,0) as type_sum");
	      sqlCmd.append(" from a10 left join (select * from bn01 where m_year=?)bn01 on a10.bank_code = bn01.bank_no "); 
	      sqlCmd.append(" where a10.m_year=?");
	      sqlCmd.append(" and a10.m_month=?");
	      sqlCmd.append(" and a10.bank_code=?");  
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(wlx01_m_year);
	      paramList.add(S_YEAR);
	      paramList.add(S_MONTH);
	      paramList.add(bank_code);

	      dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan2_amt,loan3_amt,loan4_amt,loan_sum,invest2_amt,invest3_amt,invest4_amt,invest_sum,other2_amt,other3_amt,other4_amt,other_sum,type2_sum,type3_sum,type4_sum,type_sum");

	      
	      System.out.println("dbData.size=" + dbData.size());
	     
	      //設定報表表頭資料============================================
	      
	      row=sheet.getRow(1);
	      cell=row.getCell((short)1);	       	
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue("民國"+S_YEAR+"年"+S_MONTH+"月"+((dbData == null || dbData.size() ==0)?"無資料存在":""));  	
	      
	      row = sheet.getRow(2);
	      cell = row.getCell( (short) 2);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	      if(dbData != null && dbData.size() != 0){
	      	 bean = (DataObject)dbData.get(0);
	      	 bank_name = (String)bean.getValue("bank_name");
	      }
	      cell.setCellValue(bank_name);
	      
	      if (dbData == null || dbData.size() ==0) {      	      
	      }else {
	      	rowNum = 6;
	      	bean = (DataObject)dbData.get(0);
	      	data_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();
	        data_month=(bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();
	        loan2_amt=(bean.getValue("loan2_amt") == null)?"0":(bean.getValue("loan2_amt")).toString();//放款-列二類
			loan3_amt=(bean.getValue("loan3_amt") == null)?"0":(bean.getValue("loan3_amt")).toString();//放款-列三類
			loan4_amt=(bean.getValue("loan4_amt") == null)?"0":(bean.getValue("loan4_amt")).toString();//放款-列四類
			loan_sum=(bean.getValue("loan_sum") == null)?"0":(bean.getValue("loan_sum")).toString();//放款-合計
			invest2_amt=(bean.getValue("invest2_amt") == null)?"0":(bean.getValue("invest2_amt")).toString();//投資-列二類
			invest3_amt=(bean.getValue("invest3_amt") == null)?"0":(bean.getValue("invest3_amt")).toString();//投資-列三類		
			invest4_amt=(bean.getValue("invest4_amt") == null)?"0":(bean.getValue("invest4_amt")).toString();//投資-列四類  
			invest_sum=(bean.getValue("invest_sum") == null)?"0":(bean.getValue("invest_sum")).toString();//投資-合計
			other2_amt=(bean.getValue("other2_amt") == null)?"0":(bean.getValue("other2_amt")).toString();//其他-列二類
			other3_amt=(bean.getValue("other3_amt") == null)?"0":(bean.getValue("other3_amt")).toString();//其他-列三類		
			other4_amt=(bean.getValue("other4_amt") == null)?"0":(bean.getValue("other4_amt")).toString();//其他-列四類 	
			other_sum=(bean.getValue("other_sum") == null)?"0":(bean.getValue("other_sum")).toString();//其他-合計
			type2_sum=(bean.getValue("type2_sum") == null)?"0":(bean.getValue("type2_sum")).toString();//列二類-合計
			type3_sum=(bean.getValue("type3_sum") == null)?"0":(bean.getValue("type3_sum")).toString();//列三類-合計
			type4_sum=(bean.getValue("type4_sum") == null)?"0":(bean.getValue("type4_sum")).toString();//列四類-合計
			type_sum=(bean.getValue("type_sum") == null)?"0":(bean.getValue("type_sum")).toString();//總合計-合計
			
			//System.out.println("loan2_amt="+loan2_amt);
			//System.out.println("type_sum="+type_sum);
			//列款
			row = sheet.getRow(rowNum);
			for(int cellcount=2;cellcount<6;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);		     		
		    	if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(loan2_amt));	 
		    	if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(loan3_amt));
		    	if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(loan4_amt));
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(loan_sum));	     		
			}//end of cellcount	 		  	  	 	
			
			//投資
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=2;cellcount<6;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);		     		
		    	if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(invest2_amt));	 
		    	if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(invest3_amt));
		    	if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(invest4_amt));
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(invest_sum));	     		
			}//end of cellcount
			//其他
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=2;cellcount<6;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);		     		
		    	if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(other2_amt));	 
		    	if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(other3_amt));
		    	if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(other4_amt));
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(other_sum));	     		
			}//end of cellcount
			//合計
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=2;cellcount<6;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);		     		
		    	if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(type2_sum));	 
		    	if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(type3_sum));
		    	if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(type4_sum));
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(type_sum));	     		
			}//end of cellcount
	      } //end of else dbData.size() != 0
	      
	    }
	    catch (Exception e) {
	      System.out.println("A10Rpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
    
    
    
    
    public static void printRptMsg(PrintStream logps,String rptKind,String errRptMsg){
    	if(!errRptMsg.equals("")){
	       logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();
	       logps.println(logformat.format(nowlog)+rptKind+":"+errRptMsg);
	       logps.flush();
	    }
    }
    //109.05.12 add 調整使用FTPSClient上傳檔案
    public static String putFiles(String server_host, String username, String password, 
			   String remote_path, String local_path, String workDir, List filename, 
			   PrintStream logps){    
           try{    
         	   //System.setProperty("https.protocols", "TLSv1,TLSv1.1,TLSv1.2");
        	   System.setProperty("https.protocols", "TLSv1.2");//110.11.29 金管會改只開放TLS1.2//110.12.29        	   
        	   FTPSClient ftps = new FTPSClient("TLS",true);//隱含式TLS//原使用方式
        	   ftps.addProtocolCommandListener(new PrintCommandListener(new PrintWriter(System.out),true));//將command輸出
                              
               ftps.connect(server_host);

               int reply = ftps.getReplyCode();
               
               System.out.println("getReplyCode()="+reply);
               
               if (!FTPReply.isPositiveCompletion(reply)) {
                   ftps.disconnect();
                   printRptMsg(logps,"","FTP server refused connection(" + server_host+")");   
                   return "FTP server refused connection";
               }else{
            	   System.out.println("Connected to " + server_host);					    	
                   printRptMsg(logps,"","Connected to " + server_host);   
               }

               if (!ftps.login(username, password)) {
                   ftps.logout();
                   printRptMsg(logps,"","Login failed!");   
                   return "Login failed!";
               }else{
                   System.out.println("Login success");
                   printRptMsg(logps,"","Login success");   
               }
               printRptMsg(logps,"","default port is: " + ftps.getDefaultPort());   
               System.out.println("default port is: " + ftps.getDefaultPort());
               
               // set transport mode to binary
               ftps.setFileType(FTP.BINARY_FILE_TYPE);
               System.out.println("setFileType ok");
               printRptMsg(logps,"","setFileType: binary");   
               ftps.enterLocalPassiveMode();
               System.out.println("enterLocalPassiveMode ok");
               printRptMsg(logps,"","enterLocalPassiveMode ok");
               ftps.setControlKeepAliveTimeout(300);
               //useEpsvWithIPv4
               ftps.setUseEPSVwithIPv4(true);//109.05.12 add
               // Set protection buffer size
               ftps.execPBSZ(0);//109.05.08 add
               // Set data channel protection to private
               ftps.execPROT("P"); //109.05.08 add
               System.out.println("Set data channel protection to private ");
               
               //109.05.12 add begin
               ftps.setTrustManager(TrustManagerUtils.getAcceptAllTrustManager());   
               String[] cipher = ftps.getEnabledCipherSuites();
               for(int i=0;i<cipher.length;i++){
               System.out.println("Remote system is " + cipher[i]);
               }
               System.out.println("SSL: " + ftps.getEnableSessionCreation());
               // PROTOCOLOS
               String[] Protocols = ftps.getEnabledProtocols();
               for(int i=0;i<Protocols.length;i++){
            	   System.out.println("Protocols: " + Protocols[i]);   
               }
                              
               // AUTH
               boolean Auth = ftps.getNeedClientAuth();
               System.out.println("Auth: " + Auth);
               ftps.getWantClientAuth();
               ftps.getTrustManager();
               ftps.feat();              
               //109.05.12 add end                              
               
               String nowDir = "";
               printRptMsg(logps,"","filename=" + filename);    
           
                   
                  try {
                  		System.out.println("Welcome message:\n" + ftps.getReplyString());
                  		System.out.println("Current Directory: " + new String(ftps.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));     			 	
                  		//System.out.println("parentDir=" + parentDir);
                  		//97.10.02要先切換到/boaf/目錄下.才可切換到其他下層目錄(檢查局根目錄為boaf)
                  		System.out.println("change home dir [/boaf/] ?? " + ftps.changeWorkingDirectory("/boaf/"));
                  		System.out.println("change home dir [" + new String(remote_path.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftps.changeWorkingDirectory(new String((new String(remote_path.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1")));
                  		System.out.println("begin change dir to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
                  		//change remote path
                  		if(!ftps.changeWorkingDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                
                  			System.out.println("begin change dir failed!! ");
                  			printRptMsg(logps,"","begin change dir failed!! " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));  
                  			if(!ftps.makeDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                     
                  				System.out.println("ftp.makeDirectory(" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + ") failed!!");
                  				ftps.logout();
                  				return("Cannot create working directory to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
                  			} 	
                  		}
                  		printRptMsg(logps,"",":begin change dir to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
                  		printRptMsg(logps,"",":change home dir [" + new String(remote_path.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftps.changeWorkingDirectory(new String((new String(remote_path.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1")));
                    
                  		System.out.println("create dir success[" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + "]");
                  		logps.println(logformat.format(nowlog) + ":create dir success[" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + "]");
                  		logps.flush();
                  		if(!ftps.changeWorkingDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                 
                  			ftps.logout();
                  			return("Cannot change working directory to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
                  		}
                   
                  		logps.println(logformat.format(nowlog) + "Current Directory: " + new String(ftps.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));
                  		logps.flush();
                  		System.out.println("Current Directory: " + new String(ftps.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));
                  		/*109.05.12
                  		ftp.setPassiveMode(true);
                  		System.out.println("setPassiveMode ok");
                  		ftp.binary();
                  		System.out.println("binary");
                  		*/
                  		boolean success = false;  
                  		//success = ftp.uploadFile(local_path + "2008112504.xls", workDir+System.getProperty("file.separator")+"2008112504.xls");
                  		
                  		for(int i=0;i<filename.size();i++){//有檔案需上傳時
                  			success = false;
                  			System.out.println("put file begin==============================");
                  			//success = ftp.uploadFile(local_path + (String)filename.get(i),(String)filename.get(i));                  			
                  			//System.out.println((new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));97.11.25將中文檔名轉成ISO8859_1傳至檢查局中文字才會正常
                  			
                  			//success = ftps.uploadFile(local_path + (String)filename.get(i),(new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));
                  			
                  			FileInputStream in = new FileInputStream(local_path + (String)filename.get(i));//109.05.12 add    		
                  			success = ftps.storeFile((new String (((String)filename.get(i)).getBytes(),"ISO8859_1")), in);//109.05.12 add
                    		in.close();//109.05.12 add
                  			
                  			System.out.println(":put " + (String)filename.get(i) + " success? " + success);
                  			printRptMsg(logps,"","put " + local_path + (String)filename.get(i) + " to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + (String)filename.get(i) + " success? " + success);                  			
                  			//if(!success) return("put filename " + filename + " failed");//97.10.02拿掉.即使有某個檔案上傳失敗.其他的繼續上傳
                  		}//end of filename
                  		
                  } catch (Exception ftpe) {
                  		System.out.println("putFiles Error:"+ftpe.getMessage());
                  		printRptMsg(logps,"",ftpe.getMessage());
                  		ftpe.printStackTrace();
                  		return "upload Fail";
                  } finally {
                  		//disconnect
                  		if (ftps != null && ftps.isConnected()) {
                  			try {
                  				ftps.logout();
                  				ftps.disconnect();		                    
                  			} catch (IOException f) { }
                  		}				
                  }//end of finally
              
               System.out.println("Finished");
           } catch(Exception e) {
               System.out.println(e);
               System.out.println(e.getMessage());
               e.printStackTrace();
               printRptMsg(logps,"",e+e.getMessage()); 
               return "putFiles Error:"+e+e.getMessage();			
           }
           return null;
    }
    
    /*原使用FTPClient上傳檔案
    public static String putFiles(String server_host, String username, String password, 
			   String remote_path, String local_path, String workDir, List filename, 
			   PrintStream logps){    
        try{    
            JakartaFtpWrapper ftp = new JakartaFtpWrapper();
            String nowDir = "";
            printRptMsg(logps,"","filename=" + filename);     
        
            if(ftp.connectAndLogin(server_host, username, password)) {//connect成功時
               System.out.println("Connected to " + server_host);					    	
               printRptMsg(logps,"","Connected to " + server_host);   
               
               try {
               		System.out.println("Welcome message:\n" + ftp.getReplyString());
               		System.out.println("Current Directory: " + new String(ftp.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));     			 	
               		//System.out.println("parentDir=" + parentDir);
               		//97.10.02要先切換到/boaf/目錄下.才可切換到其他下層目錄(檢查局根目錄為boaf)
               		System.out.println("change home dir [/boaf/] ?? " + ftp.changeWorkingDirectory("/boaf/"));
               		System.out.println("change home dir [" + new String(remote_path.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftp.changeWorkingDirectory(new String((new String(remote_path.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1")));
               		System.out.println("begin change dir to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
               		//change remote path
               		if(!ftp.changeWorkingDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                
               			System.out.println("begin change dir failed!! ");
               			printRptMsg(logps,"","begin change dir failed!! " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));  
               			if(!ftp.makeDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                     
               				System.out.println("ftp.makeDirectory(" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + ") failed!!");
               				ftp.logout();
               				return("Cannot create working directory to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
               			} 	
               		}
               		printRptMsg(logps,"",":begin change dir to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
               		printRptMsg(logps,"",":change home dir [" + new String(remote_path.getBytes("ISO8859_1"),"UTF-8") + "] ?? " + ftp.changeWorkingDirectory(new String((new String(remote_path.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1")));
                 
               		System.out.println("create dir success[" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + "]");
               		logps.println(logformat.format(nowlog) + ":create dir success[" + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + "]");
               		logps.flush();
               		if(!ftp.changeWorkingDirectory(new String((new String(workDir.getBytes("ISO8859_1"),"UTF-8")).getBytes(),"ISO8859_1"))){                 
               			ftp.logout();
               			return("Cannot change working directory to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8"));
               		}
                
               		logps.println(logformat.format(nowlog) + "Current Directory: " + new String(ftp.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));
               		logps.flush();
               		System.out.println("Current Directory: " + new String(ftp.printWorkingDirectory().getBytes("ISO8859_1"),"UTF-8"));
               		ftp.setPassiveMode(true);
               		System.out.println("setPassiveMode ok");
               		ftp.binary();
               		System.out.println("binary");
               		boolean success = false;  
               		//success = ftp.uploadFile(local_path + "2008112504.xls", workDir+System.getProperty("file.separator")+"2008112504.xls");
               		
               		for(int i=0;i<filename.size();i++){//有檔案需上傳時
               			System.out.println("put file begin==============================");
               			//success = ftp.uploadFile(local_path + (String)filename.get(i),(String)filename.get(i));                  			
               			//System.out.println((new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));97.11.25將中文檔名轉成ISO8859_1傳至檢查局中文字才會正常
               			
               			success = ftp.uploadFile(local_path + (String)filename.get(i),(new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));
               			System.out.println(":put " + (String)filename.get(i) + " success? " + success);
               			printRptMsg(logps,"","put " + local_path + (String)filename.get(i) + " to " + new String(workDir.getBytes("ISO8859_1"),"UTF-8") + (String)filename.get(i) + " success? " + success);                  			
               			//if(!success) return("put filename " + filename + " failed");//97.10.02拿掉.即使有某個檔案上傳失敗.其他的繼續上傳
               		}//end of filename
               		
               } catch (Exception ftpe) {
               		System.out.println("putFiles Error:"+ftpe.getMessage());
               		printRptMsg(logps,"",ftpe.getMessage());
               		ftpe.printStackTrace();
               		return "upload Fail";
               } finally {
               		//disconnect
               		if (ftp != null && ftp.isConnected()) {
               			try {
               				ftp.logout();
               				ftp.disconnect();		                    
               			} catch (IOException f) { }
               		}				
               }//end of finally
            } else {//無法connect至Server時
               System.out.println("Unable to connect to " + server_host);
               printRptMsg(logps,"","Unable to connect to " + server_host); 
               return "Unable to connect to " + server_host;			
            }
            System.out.println("Finished");
        } catch(Exception e) {
            System.out.println(e);
            System.out.println(e.getMessage());
            e.printStackTrace();
            printRptMsg(logps,"",e+e.getMessage()); 
            return "putFiles Error:"+e+e.getMessage();			
        }
        return null;
    }
    */
    
    /* String server_host:主機ip
     * String username:使用者帳號 
     * String password:使用者名稱
     * String remote_path:主機端目錄 ex:/export/home/pboafmgr/AgriBankDir/
     * String local_path:欲上傳檔案的目錄 ex:D:\\workProject\\BOAF\\AgriBankData\\
     * String workDir:次目錄名稱 ex:AgriRpt01/9808/007
     * List filename:欲上傳檔名List
     * PrintStream logps:print log 
     */
    
    public static String putFiles_multiSubDir(String server_host, String username, String password, 
			   String remote_path, String local_path, String workDir, List filename, 
			   PrintStream logps){    
        try{    
            JakartaFtpWrapper ftp = new JakartaFtpWrapper();
           
            String nowDir = "";
            String moveDir="";
            String serverRptDir="";	
            printRptMsg(logps,"","filename=" + filename);     
                    
            if(ftp.connectAndLogin(server_host, username, password)) {//connect成功時
               System.out.println("Connected to " + server_host);					    	
               printRptMsg(logps,"","Connected to " + server_host);  
              
               try {
               		System.out.println("Welcome message:\n" + ftp.getReplyString());
               		System.out.println("Current Directory: " + ftp.printWorkingDirectory());
               		serverRptDir=Utility.getProperties("serverRptDir");		
               		System.out.println("change home dir [" + serverRptDir + "] ?? " + ftp.changeWorkingDirectory(serverRptDir));
               		System.out.println("change home dir [" + remote_path + "] ?? " + ftp.changeWorkingDirectory(remote_path));
               		System.out.println("begin change dir to " + workDir);
               		List dir_List = Utility.getStringTokenizerData(workDir,"/");
               		moveDir = remote_path; 
               		for(int i = 0; i < dir_List.size(); i++){//建立其subdir
                		nowDir = (String)dir_List.get(i);                		
                		System.out.println("Current Directory: " + ftp.printWorkingDirectory());
                		//moveDir+="/"+nowDir;
                		if(!nowDir.equals("")){
                			if(!ftp.changeWorkingDirectory(nowDir)){  
                			   if(!ftp.makeDirectory(nowDir)){
                				  printRptMsg(logps,"",":create sub dir fail[" + nowDir + "]");
                				  System.out.println("create sub dir fail[" + nowDir + "]");
                			   } else{
                				  printRptMsg(logps,"",":create sub dir success[" + nowDir + "]");
                				  System.out.println("create sub dir success[" + nowDir + "]");
                			   }
                			   ftp.changeWorkingDirectory(nowDir);
                			}                			
                		}    
                	}

               		logps.println(logformat.format(nowlog) + "Current Directory: " + ftp.printWorkingDirectory());
               		logps.flush();
               		System.out.println("Current Directory: " + ftp.printWorkingDirectory());
               		ftp.setPassiveMode(true);
               		System.out.println("setPassiveMode ok");
               		ftp.binary();
               		System.out.println("binary");
               		boolean success = false;  
                 		
               		for(int i=0;i<filename.size();i++){//有檔案需上傳時
               			System.out.println("put file begin==============================");
               			//success = ftp.uploadFile(local_path + (String)filename.get(i),(String)filename.get(i));                  			
               			//System.out.println((new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));97.11.25將中文檔名轉成ISO8859_1傳至檢查局中文字才會正常
               			
               			success = ftp.uploadFile(local_path + (String)filename.get(i),(new String (((String)filename.get(i)).getBytes(),"ISO8859_1")));
               			System.out.println(":put " + (String)filename.get(i) + " success? " + success);
               			printRptMsg(logps,"","put " + local_path + (String)filename.get(i) + " to " + remote_path+workDir + (String)filename.get(i) + " success? " + success);                  			
               			//if(!success) return("put filename " + filename + " failed");//97.10.02拿掉.即使有某個檔案上傳失敗.其他的繼續上傳
               		}//end of filename
               		
               } catch (Exception ftpe) {
               		System.out.println("putFiles_multiSubDir Error:"+ftpe.getMessage());
               		printRptMsg(logps,"",ftpe.getMessage());
               		ftpe.printStackTrace();
               		return "upload Fail";
               } finally {
               		//disconnect
               		if (ftp != null && ftp.isConnected()) {
               			try {
               				ftp.logout();
               				ftp.disconnect();		                    
               			} catch (IOException f) { }
               		}				
               }//end of finally
            } else {//無法connect至Server時
               System.out.println("Unable to connect to " + server_host);
               printRptMsg(logps,"","Unable to connect to " + server_host); 
               return "Unable to connect to " + server_host;			
            }
            System.out.println("Finished");
        } catch(Exception e) {
            System.out.println(e);
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
        return null;
    }
 
    //98.10.19 有失敗時.發送e-mail
    //99.01.05 add 檢查局上傳/下載作業,增加一組e-mail
    public static void send_Mail(String Subject,String messageText){      
      try {     
           String SMTP_HOST = Utility.getProperties("SMTP_Host").trim();
           String FROM_ADDR = Utility.getProperties("From_Addr").trim();
           String SEND_MAIL = Utility.getProperties("Send_Mail").trim();
           String UserID     = Utility.getProperties("UserID").trim();
           String PWD    = Utility.getProperties("PWD").trim();
           String feb_Addr1   = Utility.getProperties("feb_Addr1").trim();
           String feb_Addr2   = Utility.getProperties("feb_Addr2").trim();
           String feb_Addr3   = Utility.getProperties("feb_Addr3").trim();
           String feb_Addr4   = Utility.getProperties("feb_Addr4").trim();
           String feb_Addr5   = Utility.getProperties("feb_Addr5").trim();
           String Auth        = (Utility.getProperties("Auth")==null?"false":Utility.getProperties("Auth")).trim();
           boolean auth = Boolean.getBoolean(Auth);
           boolean sessionDebug = true;
           if(SEND_MAIL.equals("true")){
              if(Subject.equals("")){
                  Subject = "檢查局財務報表產生及上傳作業失敗";
              }
              //設定所要用的Mail 伺服器和所使用的傳送協定
              Properties props = System.getProperties();
              if (SMTP_HOST != null) props.put("mail.smtp.host", SMTP_HOST);
              if (auth) props.put("mail.smtp.auth", "true");
              //產生新的Session 服務
              Session mailSession = Session.getInstance(props, null);
              if (sessionDebug) mailSession.setDebug(true);    
              Message msg = new MimeMessage(mailSession);       
              //設定傳送郵件的發信人
              msg.setFrom(new InternetAddress(FROM_ADDR));
              //設定傳送郵件至收信人的信箱
              msg.setRecipients(Message.RecipientType.TO,InternetAddress.parse(feb_Addr1+","+feb_Addr2+","+feb_Addr3, false));
              msg.setRecipients(Message.RecipientType.CC,InternetAddress.parse(feb_Addr4+","+feb_Addr5, false));
              //設定信中的主題 
              msg.setSubject(Subject);
              //設定送信的時間
              msg.setSentDate(new Date());
              Multipart mp = new MimeMultipart();
              msg.setText(messageText);   
              //Transport.send(msg);      
              Transport transport = mailSession.getTransport("smtp"); 
              transport.connect(SMTP_HOST, UserID, PWD); 
              transport.sendMessage(msg, msg.getAllRecipients()); 
              transport.close();
           }
          return ;
      }catch (Exception e){
        System.out.println("RptGenerateALL.send_Mail Error:"+e.getMessage());
      }
    }
}
