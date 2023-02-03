/*
  111.01.10 檢查局財務報表檔案上傳須使用java1.8版本及commons-net-3.6.jar才能使用FTPSClient上傳檔案 by 2295            
*/
package com.tradevan.util.ftp;

import java.io.*;
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

import org.apache.commons.net.PrintCommandListener;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPReply;
import org.apache.commons.net.ftp.FTPSClient;
import org.apache.commons.net.util.TrustManagerUtils;

import com.tradevan.util.Utility;


public class FebFtps {	
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
	    	   logfile = new File(logDir + System.getProperty("file.separator") + "FebFtps."+ logfileformat.format(nowlog));						 
	    	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"FebFtps."+ logfileformat.format(nowlog));
	    	   logos = new FileOutputStream(logfile,true);  		        	   
	    	   logbos = new BufferedOutputStream(logos);
	    	   logps = new PrintStream(logbos);			    
	    	   System.out.println("=============執行檢查局財務報表發佈開始===========");
	    	   String errMsg = "";
	    	   FebFtps a = new FebFtps();
	    	   errMsg += a.upload(YEAR,MONTH,logps);
	    	   System.out.println("errMsg = "+errMsg);   
	    	   logcalendar = Calendar.getInstance(); 
	    	   nowlog = logcalendar.getTime();
	    	   logps.println(logformat.format(nowlog)+errMsg);		    					    
	    	   logps.flush();
	    	   System.out.println("=============執行檢查局財務報表發佈結束===========");
		 	} catch(Exception e) {
	            System.out.println(e);
	            System.out.println(e.getMessage());
	            e.printStackTrace();                    
	        }   
	} 
	 
	public static String upload(String S_YEAR,String S_MONTH,PrintStream logps){
			String errMsg = "";
			String putMsg = "";	
			
	    try{ 
	    	   String[] fname;	
	    	   File tmpFile = null;
	    	   String rptIP_feb=Utility.getProperties("rptIP_feb");			
	           String rptID_feb=Utility.getProperties("rptID_feb");			
	           String rptPwd_feb=Utility.getProperties("rptPwd_feb");
	           List filename_List  = new LinkedList();	
	           List filename_List_FR001WB  = new LinkedList();//主要經營指標/各會員別放款金額一覽表	
	           List filename_List_MCRptAll = new LinkedList();//解釋函令/限制或核准業務/處分書
	           List filename_List_MC012W = new LinkedList();//舞弊案件
	           List filename_List_MC014W = new LinkedList();//檢舉書
	           List filename_List_FR055W = new LinkedList();//理監事基本資料
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
	                
		        }//end of febxlsDir存在
	    } catch(Exception e) {
            System.out.println(e);
            System.out.println(e.getMessage());
            e.printStackTrace();
            printRptMsg(logps,"",e+e.getMessage()); 
            return "upload Error:"+e+e.getMessage();			
        }
        return null;
	           
	}		
	//109.05.12 add 調整使用FTPSClient上傳檔案
    public static String putFiles(String server_host, String username, String password, 
			   String remote_path, String local_path, String workDir, List filename, 
			   PrintStream logps){    
           try{ 
        	   /*
        	   //108.12.26
        	     會出現錯誤訊息 Unrecognized SSL message, plaintext connection?Unrecognized SSL message, plaintext connection?
        	     
               X509TrustManager sunJSSEX509TrustManager;
               // 加載 Keytool 生成的証書文件
               char[] kphrase;
               String p = "changeit";
               kphrase = p.toCharArray();
               File file = new File("C:/jdk1.6.0_10/jre/lib/security/cacerts");
               System.out.println("Loading KeyStore " + file + "...");
               InputStream in_ca = new FileInputStream(file);
               KeyStore ks = KeyStore.getInstance(KeyStore.getDefaultType());
               ks.load(in_ca, kphrase);
               in_ca.close();
               for(int pi=0;pi < kphrase.length;pi++){
              	 kphrase[pi]=' ';
               }
               System.out.println("test1="+ks.getCertificate("febsr11-RootCA")); //檢查局的root憑證     
               System.out.println("test2="+ks.getCertificate("examweb"));//網站的憑證                  
               TrustManagerFactory tmf;
               tmf = TrustManagerFactory.getInstance(TrustManagerFactory.getDefaultAlgorithm());
               
               tmf.init(ks);
               TrustManager tms [] = tmf.getTrustManagers();
               
               System.out.println("FebFtps:java.version="+(System.getProperties()).getProperty("java.version")); 
             
               // 使用構造好的 TrustManager 訪問相應的 https 站點
               //SSLContext sslContext = SSLContext.getInstance("SSL", "SunJSSE");
               System.out.println("default protocol="+(SSLContext.getDefault()).getProtocol());
               SSLContext sslContext = SSLContext.getInstance("TLSv1.2");
               sslContext.init(null, tms, new java.security.SecureRandom());
               
               SSLSocketFactory ssf = sslContext.getSocketFactory();
               System.out.println("sslContext.getProtocol()="+sslContext.getProtocol());
        	   
               FTPSClient ftps = new FTPSClient(true,sslContext);//隱含式TLS//110.12.06 add
               ftps.setSocketFactory(ssf);//110.12.06 add
               */
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
	
	/*
    public static String upload(PrintStream logps){
		String errMsg = "";
		        
		try{	
			String rptIP_coa=Utility.getProperties("rptIP_coa");			
	        String rptID_coa=Utility.getProperties("rptID_coa");			
	        String rptPwd_coa=Utility.getProperties("rptPwd_coa");	
	        List filename_List  = new LinkedList();
	        File rptDir = new File(Utility.getProperties("coaxlsDir"));
	        coaxlsDir = new File(Utility.getProperties("coaxlsDir"));
	        coaBKxlsDir = new File(Utility.getProperties("coaBKxlsDir"));  
            File rptFile = null;
            String[] fname1;	
            String FTP_DIRECTORY = "";
            
            printRptMsg(logps,"============執行上傳農委會 OpenData所需報表開始============"); 
           
           
		    if(!coaxlsDir.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("coaxlsDir"))){ 		       		
 		       	    printRptMsg(logps,Utility.getProperties("coaxlsDir")+"目錄新增失敗"); 
 		    	}    
		    }
		    
		    if(!coaBKxlsDir.exists()){
 		    	if(!Utility.mkdirs(Utility.getProperties("coaBKxlsDir"))){ 		       		
 		       	    printRptMsg(logps,Utility.getProperties("coaBKxlsDir")+"目錄新增失敗"); 
 		    	}    
		    }
            
            
            boolean uploadSuccess = false;
            if(rptDir.exists() && rptDir.isDirectory()){
               MySFTPClient msftp = new MySFTPClient(rptIP_coa, rptID_coa, rptPwd_coa);
               fname1= rptDir.list(); //====列出此目錄下的所有檔案===================
               for(int c=0;c<fname1.length;c++){
                   rptFile = new File(Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+fname1[c]);
                   uploadSuccess = false;
                   if(!rptFile.isDirectory()){
                      uploadSuccess = msftp.sendMyFiles(FTP_DIRECTORY, Utility.getProperties("coaxlsDir")+System.getProperty("file.separator"), fname1[c]);
                      filename_List.add(fname1[c]);
                      printRptMsg(logps,fname1[c]+(uploadSuccess==true?"檔案上傳完成":"檔案上傳失敗")); 
                   }
               }               
           }   
            
           //將coaxlsDir目錄下上傳成功的.搬移至所對應的bkdir================================= 
           errMsg += moveBKDir_complete(logps,filename_List); 
           //errMsg +="<br>上列檢查局各式報表,檔案上傳完成";
           printRptMsg(logps,errMsg+"檔案已搬移至"+coaBKxlsDir);  
           printRptMsg(logps,errMsg+"============執行上傳農委會 OpenData所需報表結束============");          
    	
		}catch(Exception e){
			System.out.println("CoaFtp.upload Error:"+e+e.getMessage());			
		}
		return errMsg;
	}
    
    //將上傳至Server成功後的檔案.搬移至備份目錄
    private static String moveBKDir_complete(PrintStream logps,List filename_List){
    	String errMsg="";    	
		File tmpFile = null;
		String copyResult = "";
    	try{
    		
            for(int i=0;i<filename_List.size();i++){            	
                errMsg +="<br>"+(String)filename_List.get(i);
                tmpFile = new File(Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+(String)filename_List.get(i));                       
                if(!tmpFile.isDirectory() && tmpFile.exists()){
                   copyResult = Utility.CopyFile(coaxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i),coaBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i));
             	   System.out.println(coaxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" copy to "+coaBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" success ?? "+copyResult);
             	   printRptMsg(logps,coaxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" copy to "+coaBKxlsDir+System.getProperty("file.separator")+(String)filename_List.get(i)+" success ?? "+copyResult);                    	  
             	   if(copyResult.equals("0")) tmpFile.delete();
                }
            }
            
    	}catch(Exception e){
    		System.out.println("CoaFtp.moveBKDir_complete Error:"+e+e.getMessage());
    		errMsg += "CoaFtp.moveBKDir_complete Error:"+e+e.getMessage();    		
    	}
    	return errMsg;
    }
    */
    
    public static void printRptMsg(PrintStream logps,String rptKind,String errRptMsg){
    	if(!errRptMsg.equals("")){
	       logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();
	       logps.println(logformat.format(nowlog)+rptKind+":"+errRptMsg);
	       logps.flush();
	    }
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

