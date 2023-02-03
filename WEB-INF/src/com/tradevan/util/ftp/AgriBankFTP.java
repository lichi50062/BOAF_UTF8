/* 
  要記得在本機端根目錄下面加上四各檔案名稱--ATM200607.DAT,CHK200607.DAT,FGN200607.DAT,INT200603.DAT
                             FRM201401.DAT,ITEM201401.DAT -->2014.04.02 add 專案農貸
  96.01.03 fix 晶片卡流通張數:數字（43-49）by 2295
  96.01.03 fix 牌告利率取到小數後3位 by 2295
  96.01.05 fix ATM.CHG.FGN.INT.重新轉檔,模組化並加上Log by 2295
  96.01.10 fix 清空static變數 by 2295
  96.01.16 add 重新至金庫取檔 by 2295 
  96.02.07 避免sql injection by 2295 
  96.09.28 add 牌告利率增加欄位.改為月報 by 2295
  96.11.22 add ATM修改傳輸規格 by 2295
  97.02.18 fix ATM轉檔.當每年一月份時.本年累計=當月份交易資料 by 2295
  98.03.30 add 牌告利率增加欄位.基準利率-指標利率（月調）/基準利率（月調）/指數型房貸指標利率（月調） by 2295
 100.01.13 fix sql injection 使用preparestatement by 2295 
 103.04.02 add 專案農貸明細資料/專案農貸貸款項目資料  by 2295
 109.02.17 fix 處理F01排程更新資料庫問題 by 2295
*/

package com.tradevan.util.ftp;
import org.apache.commons.net.ftp.*;
import org.apache.commons.vfs2.FileObject;
import org.apache.commons.vfs2.FileSystemOptions;
import org.apache.commons.vfs2.Selectors;
import org.apache.commons.vfs2.impl.StandardFileSystemManager;
import org.apache.commons.vfs2.provider.sftp.SftpFileSystemConfigBuilder;
import java.io.*;
import java.util.*;
import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.DownLoad;

import java.util.Properties;
import java.util.Date;
import javax.mail.*;
import javax.mail.internet.*;
import java.text.SimpleDateFormat;

public class  AgriBankFTP{
    static List updateDBSqlList = new LinkedList();
    //農漁會在台無住所外國人新台幣存款資料
    static String  ACCT_CNT_TM="",DEP_TYPE="",ACCT_TYPE="",BAL_LM="",DEP_TM="",WTD_TM="",BAL_TM="";
    static double ACCT_CNT_TM_sum=0.0,BAL_LM_sum =0.0,DEP_TM_sum=0.0,WTD_TM_sum=0.0,BAL_TM_sum=0.0;//F01.小計
    //F01.合計
    static double[] ACCT_CNT_TM_idx = {0.0,0.0,0.0,0.0,0.0,0.0};
    static double[] BAL_LM_idx = {0.0,0.0,0.0,0.0,0.0,0.0};
    static double[] DEP_TM_idx = {0.0,0.0,0.0,0.0,0.0,0.0};
    static double[] WTD_TM_idx = {0.0,0.0,0.0,0.0,0.0,0.0};
    static double[] BAL_TM_idx = {0.0,0.0,0.0,0.0,0.0,0.0};
    static String preBANK_NO = "";
    static String BANK_NO="";
    static int cnt=0;
    static String INPUT_METHOD="F";//F:檔案上傳
    static String COMMON_CENTER="N";//Y:表示是由共用中心代傳入
    static String UPD_METHOD="A";//A:自動(排程)
    static String UPD_CODE="N";//N:待檢核
    static String BATCH_NO="1",LOCK="";
    static String sqlCmd = "";
    static File logfile;
    static FileOutputStream logos=null;    	
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    static String mailTitle = "";
    
public static int getDataFiles(String FTP_SERVER, String FTP_USER,String FTP_PASSWORD, String FTP_DIRECTORY, String LOCAL_DIRECTORY, String TargetFile,String getFiles)throws IOException{    
      String frist_type = "";
      String M_YEAR="",M_MONTH="",M_Quarter="";
      String getfr = "";
      String UPDATE_DATE="";
      int NumCount=0;
      BufferedReader fr = null;
      List paramList =new ArrayList() ;
      List updateDBList = new ArrayList();//0:sql 1:data
      List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
  	  List updateDBDataList = new ArrayList();//儲存參數的List
  	  List updateDBDataList_WML01 = new ArrayList();//儲存參數的List
  	  List updateDBDataList_F01 = new ArrayList();//儲存參數的List
   try{ 
   	   FTP_SERVER = FTP_SERVER.replaceAll("'","''");
	   FTP_USER = FTP_USER.replaceAll("'","''");
	   FTP_PASSWORD = FTP_PASSWORD.replaceAll("'","''");
	   FTP_DIRECTORY = FTP_DIRECTORY.replaceAll("'","''");
	   LOCAL_DIRECTORY = LOCAL_DIRECTORY.replaceAll("'","''");
	   TargetFile = TargetFile.replaceAll("'","''");
	   getFiles = getFiles.replaceAll("'","''");
   	   //96.01.10 清空static變數===================================================================
	   paramList =new ArrayList() ;
	   updateDBList = new ArrayList();//0:sql 1:data
	   updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
	   updateDBDataList = new ArrayList();//儲存參數的List
	   updateDBDataList_WML01 = new ArrayList();//儲存參數的List
	   updateDBDataList_F01 = new ArrayList();//儲存參數的List
       
       System.out.println("paramList.size()="+paramList.size());
       System.out.println("updateDBList.size()="+updateDBList.size());
       System.out.println("updateDBSqlList.size()="+updateDBSqlList.size());
       System.out.println("updateDBDataList.size()="+updateDBDataList.size());
       System.out.println("updateDBDataList_WML01.size()="+updateDBDataList_WML01.size());
       System.out.println("updateDBDataList_F01.size()="+updateDBDataList_F01.size());
       
       ACCT_CNT_TM="";DEP_TYPE="";ACCT_TYPE="";BAL_LM="";DEP_TM="";WTD_TM="";BAL_TM="";
       ACCT_CNT_TM_sum=0.0;BAL_LM_sum =0.0;DEP_TM_sum=0.0;WTD_TM_sum=0.0;BAL_TM_sum=0.0;//F01.小計
       //F01.合計
       for(int i=1;i<ACCT_CNT_TM_idx.length;i++){           
           ACCT_CNT_TM_idx[i]=0.0;                          
		      BAL_LM_idx[i]=0.0;
		      DEP_TM_idx[i]=0.0;
		      WTD_TM_idx[i]=0.0;
		      BAL_TM_idx[i]=0.0;
       }       
       preBANK_NO = "";
       BANK_NO="";
       cnt=0;
       sqlCmd="";
       //=========================================================================================
       logDir  = new File(Utility.getProperties("logDir"));
 	   if(!logDir.exists()){
           if(!Utility.mkdirs(Utility.getProperties("logDir"))){
      		  System.out.println("目錄新增失敗");
      	  }    
        }
        
 	   logfile = new File(logDir + System.getProperty("file.separator") +"AgriBankFTP."+ logfileformat.format(nowlog));						 
 	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + "AgriBankFTP."+ logfileformat.format(nowlog));
 	   logos = new FileOutputStream(logfile,true);  		        	   
 	   logbos = new BufferedOutputStream(logos);
 	   logps = new PrintStream(logbos);			          
       logcalendar = Calendar.getInstance(); 
 	   nowlog = logcalendar.getTime();			    	
 	   logps.println(logformat.format(nowlog)+" "+"執行金庫轉檔程式開始");		    					    
 	   logps.flush();
 	  //96.01.16 add 重新至金庫取檔 by 2295 ========================================================================== 
 	  if(getFiles.equals("true")){
         //ftp 下載檔案 	     
         //int iftpFileResult = ftpgetFile(FTP_SERVER,FTP_USER,FTP_PASSWORD,FTP_DIRECTORY,LOCAL_DIRECTORY,TargetFile);
         //sftp 下載檔案
         int isftpFileResult = sftpGetFile(FTP_SERVER,FTP_USER,FTP_PASSWORD,FTP_DIRECTORY,LOCAL_DIRECTORY,TargetFile);
         if(isftpFileResult != 1) return 0;
         logcalendar = Calendar.getInstance(); 
 		 nowlog = logcalendar.getTime();			    	
 		 logps.println(logformat.format(nowlog)+" "+LOCAL_DIRECTORY+File.separator+TargetFile+"檔案下載成功");		    
 		 logps.flush();
 	  }
 	  //========================================================================================================
      //parse TargetFile
      File file = new File(LOCAL_DIRECTORY+File.separator+TargetFile);
      fr= new BufferedReader(new FileReader(file));      
      System.out.println("file="+LOCAL_DIRECTORY+File.separator+TargetFile);            
      String sub_TargetFile = TargetFile.substring(0,3);
      if(sub_TargetFile.equals("ITE")) sub_TargetFile = TargetFile.substring(0,4);
      logcalendar = Calendar.getInstance(); 
	  nowlog = logcalendar.getTime();			    	
	  logps.println(logformat.format(nowlog)+" "+LOCAL_DIRECTORY+File.separator+TargetFile+"檔案讀取成功");		    
	  logps.flush();
 	  
      whileLoop:
      while(fr.ready()){//如果檔案沒有讀完，就繼續處理            
          	getfr = fr.readLine(); //取得一行輸入
          	if(getfr == null) continue;
          	System.out.println(getfr);          	
          	frist_type = getfr.substring(0,1);//1:表頭.2:表身.3:表尾
          	if(!frist_type.equals("2")){
          	   BANK_NO = getfr.substring(1,8);
		       UPDATE_DATE = getfr.substring(8,16);
          	}
		    //表頭
		    if(frist_type.equals("1")){
		        System.out.println("BANK_NO="+BANK_NO);
			    System.out.println("UPDATE_DATE="+UPDATE_DATE);
			    logcalendar = Calendar.getInstance(); 
				nowlog = logcalendar.getTime();			    	
				logps.println(logformat.format(nowlog)+" "+"發件單位:"+BANK_NO+":發件日期"+UPDATE_DATE);		    
				logps.flush();
		        continue;		    		    
		    }
		    if(frist_type.equals("3")){//表尾
		       NumCount = Integer.parseInt(getfr.substring(16,22));
		       System.out.println("BANK_NO="+BANK_NO);
			   System.out.println("UPDATE_DATE="+UPDATE_DATE);
		       System.out.println("NumCount="+NumCount);
		       logcalendar = Calendar.getInstance(); 
			   nowlog = logcalendar.getTime();			    	
			   logps.println(logformat.format(nowlog)+" "+"發件單位:"+BANK_NO+":發件日期"+UPDATE_DATE+"總筆數:"+NumCount);		    
			   logps.flush();
			   
		       if(sub_TargetFile.equals("FGN")){//農漁會在台無住所外國人新台幣存款資料
			      System.out.println("write last FGN.... : ");
			      M_YEAR = M_YEAR.replaceAll("'","''");
              	  M_MONTH = M_MONTH.replaceAll("'","''");
              	  preBANK_NO = preBANK_NO.replaceAll("'","''");              	     	  
	              for(int i=1;i<ACCT_CNT_TM_idx.length;i++){	              	  
	                  //sqlCmd = "INSERT INTO F01 VALUES(?,?,?,?,?,?,?,?,?,?)";
	                  paramList = new ArrayList();    
	                  paramList.add(M_YEAR);
	                  paramList.add(M_MONTH);
	                  paramList.add(preBANK_NO);
	                  paramList.add("E");
	                  paramList.add(String.valueOf(i));
	                  paramList.add(Double.toString(ACCT_CNT_TM_idx[i]));
	                  paramList.add(Double.toString(BAL_LM_idx[i]));
	                  paramList.add(Double.toString(DEP_TM_idx[i]));
	                  paramList.add(Double.toString(WTD_TM_idx[i]));
	                  paramList.add(Double.toString(BAL_TM_idx[i]));
	                  updateDBDataList_F01.add(paramList); 
	              }
	             
	        	  //WML01用
	        	  paramList = new ArrayList();    
	        	  paramList.add(M_YEAR);
	        	  paramList.add(M_MONTH);
	        	  paramList.add(preBANK_NO);
	        	  paramList.add(INPUT_METHOD);
	        	  paramList.add(COMMON_CENTER);
	        	  paramList.add(UPD_METHOD);
	        	  paramList.add(UPD_CODE);
	        	  paramList.add(BATCH_NO);
	        	  paramList.add(LOCK);
	        	  updateDBDataList_WML01.add(paramList);	        	  
	           }		       
		       continue;
		    } 
		    if(frist_type.equals("2")){//表身			      
		       if(!sub_TargetFile.equals("ITEM")){
		          M_YEAR="";M_MONTH="";
		          M_YEAR  = Integer.toString(Integer.parseInt(getfr.substring(1,5))-1911);//年度:數字（2-5）
		          M_MONTH = Integer.toString(Integer.parseInt(getfr.substring(5,7)));//月份:數字（6-7）
	      	      //M_Quarter = Integer.toString(Integer.parseInt(getfr.substring(5,7)));//月份:數字（6-7）
	      	      BANK_NO = getfr.substring(7,14);//金融總機構代號:文數字（8-14）	      	   
	      	      if(cnt==0){
	          	     preBANK_NO = BANK_NO;
	              }
	              cnt++;
	              System.out.print("M_YEAR="+M_YEAR);
		          System.out.print(":M_MONTH="+M_MONTH); 
		          System.out.print(":BANK_NO="+BANK_NO);
		       }		       
		       if(sub_TargetFile.equals("ATM")){//農漁會金融卡發卡情形及ATM裝設情形統計		
		          mailTitle = "農漁會金融卡發卡情形及ATM裝設情形統計";
		       	  updateDBDataList.add(parseATM(getfr,M_YEAR,M_MONTH,BANK_NO));		          
		       }//end of ATM		     
		       if(sub_TargetFile.equals("CHK")){//農漁會支票存款資料
		          mailTitle = "農漁會支票存款資料";
		       	  updateDBDataList.add(parseCHK(getfr,M_YEAR,M_MONTH,BANK_NO));			      
		       }//end of CHK  
		       if(sub_TargetFile.equals("FGN")){//農漁會在台無住所外國人新台幣存款資料
		          mailTitle = "農漁會在台無住所外國人新台幣存款資料";
		       	  parseFGN(getfr,M_YEAR,M_MONTH,BANK_NO,updateDBDataList_F01,updateDBDataList_WML01);
               }//end of FGN		     
		       if(sub_TargetFile.equals("INT")){//農漁會信用部牌告利率申報資料
		          mailTitle = "農漁會信用部牌告利率申報資料";
		       	  updateDBDataList.add(parseINT(getfr,M_YEAR,M_MONTH,BANK_NO));				  
		       }//end of INT
		       if(sub_TargetFile.equals("FRM")){//103.04.02 add專案農貸明細資料
		          mailTitle = "專案農貸明細資料";
	              updateDBDataList.add(parseFRM(getfr,M_YEAR,M_MONTH,BANK_NO));               
	           }//end of FRM
		       if(sub_TargetFile.equals("ITEM")){//103.04.02 add專案農貸貸款項目代碼資料	
		          mailTitle = "專案農貸貸款項目代碼資料";
                  updateDBDataList.add(parseFRMITEM(getfr));               
               }//end of ITEM
		       logcalendar = Calendar.getInstance(); 
			   nowlog = logcalendar.getTime();			    	
			   logps.println(logformat.format(nowlog)+" "+"機構代號:"+BANK_NO);		    
			   logps.flush();		     
		    }//end of 2:表身
      }//while fr.ready()
      //96.11.22 add 
	  
	  if(sub_TargetFile.equals("FGN")){//農漁會在台無住所外國人新台幣存款資料	  	 
	  	 if(updateDBDataList_F01.size() >= 1){
	  	    sqlCmd = "INSERT INTO F01 VALUES(?,?,?,?,?,?,?,?,?,?)";	  
	  	    updateDBSqlList = new ArrayList();//109.02.17 add 
	  	    updateDBSqlList.add(sqlCmd);
	  	    updateDBSqlList.add(updateDBDataList_F01);
	  	    updateDBList.add(updateDBSqlList);
	  	 }   
	  	 if(updateDBDataList_WML01.size() >= 1){
	        sqlCmd = "INSERT INTO WML01 VALUES(?,?,?,'F01',?,'A111111111','A111111111',sysdate,?,?,?,?,?,'A111111111','A111111111',sysdate)";
	        updateDBSqlList = new ArrayList();
	      	updateDBSqlList.add(sqlCmd);
            updateDBSqlList.add(updateDBDataList_WML01);
            updateDBList.add(updateDBSqlList);
	     }
	  }else{//end of FGN	
	  	  if(sub_TargetFile.equals("ATM")){//農漁會金融卡發卡情形及ATM裝設情形統計		
		  	sqlCmd = "INSERT INTO WLX05_M_ATM VALUES(?,?,?,?,?,?,?,?,?,'',?,'','A111111111','A111111111',sysdate,?,?,?,'',?,'')";	  	
		  }
		  if(sub_TargetFile.equals("CHK")){//農漁會支票存款資料
		  	sqlCmd = "INSERT INTO WLX07_M_CHECKBANK VALUES(?,?,?,?,?,?,?,?,?,'A111111111','A111111111',sysdate)";   	     	  	
	      }//end of CHK
		  if(sub_TargetFile.equals("INT")){//農漁會信用部牌告利率申報資料      	
	      	sqlCmd = "INSERT INTO WLX_S_RATE VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'A111111111','A111111111',sysdate,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
	      }	 
		  if(sub_TargetFile.equals("FRM")){//103.04.02 add專案農貸明細資料
	        sqlCmd = "INSERT INTO AGRI_LOAN VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'A111111111','A111111111',sysdate)";
	      }  
		  if(sub_TargetFile.equals("ITEM")){//103.04.02 add專案農貸貸款項目代碼資料
	         sqlCmd = "INSERT INTO AGRI_LOAN_ITEM VALUES(?,?,'','','A111111111','A111111111',sysdate)";
	      } 
		  updateDBSqlList = new ArrayList();//109.02.17 add
		  updateDBSqlList.add(sqlCmd);
		  updateDBSqlList.add(updateDBDataList);
		  updateDBList.add(updateDBSqlList);
	  }
      
      if(!DBManager.updateDB_ps(updateDBList)){		
          logcalendar = Calendar.getInstance(); 
    	  nowlog = logcalendar.getTime();
    	  logps.println(logformat.format(nowlog)+" "+"更新資料庫失敗"+DBManager.getErrMsg());		    
    	  logps.flush();
    	  return 0;
      }
      
      System.out.println("sub_TargetFile="+sub_TargetFile);
	  if(sub_TargetFile.equals("ATM")){//農漁會金融卡發卡情形及ATM裝設情形統計--統計年累計
	  	 if(updateSumYear(M_YEAR,M_MONTH) != 0){//統計ATM交易次數/交易金額;金融卡交易次數/交易金額本年度累計
	  	 	logps.println(logformat.format(nowlog)+" "+"ATM更新資料庫失敗"+DBManager.getErrMsg());		    
	    	logps.flush();
	    	return 0;
	  	 }    	  	 
	  }
	  
      logcalendar = Calendar.getInstance(); 
	  nowlog = logcalendar.getTime();			    	
	  logps.println(logformat.format(nowlog)+" "+"執行金庫轉檔程式完成");		    
	  logps.flush();   
      send_Mail(mailTitle+"-至農業金庫轉檔成功",mailTitle+"-至全國農業金庫轉檔完成");
  }catch (Exception e){
   	   System.out.println(e+":"+e.getMessage()); 
       return 0;
  }finally{
  	if(fr != null) fr.close(); //close file  	
  }
  return 1;
}

private static int ftpgetFile(String FTP_SERVER, String FTP_USER,String FTP_PASSWORD, String FTP_DIRECTORY, String LOCAL_DIRECTORY, String TargetFile){
    int iResult=-1;
    String msgtext="";
    FileOutputStream out = null;
    try{
        FTPClient ftp = new FTPClient();
        FTPClientConfig conf = new FTPClientConfig(FTPClientConfig.SYST_UNIX);
        ftp.configure(conf);
        ftp.connect(FTP_SERVER);    
        if(!ftp.login(FTP_USER,FTP_PASSWORD)){
            System.out.println("無法連接到FTP,請查看FTP是否開啟或是網址已改變或是ID/PWD不符合,登入不成功!!");
            msgtext="無法連接到FTP,請查看FTP是否開啟或是網址已改變或是ID/PWD不符合,登入不成功!!";
            send_Mail("",msgtext);           
            return iResult=0;
        }                  
        ftp.changeWorkingDirectory(FTP_DIRECTORY);         
        System.out.println("Workdir >>" +  ftp.printWorkingDirectory()); 
               
        FTPFile[] files = ftp.listFiles();
        System.out.println( "files.... : " + files);
        System.out.println( "files.length.... : " + files.length);         		 			
		System.out.println("Current Directory: " + ftp.printWorkingDirectory());
		//下載檔案
		System.out.println("下載檔案開始:" );	
		System.out.println("下載位置:" +LOCAL_DIRECTORY+File.separator+TargetFile);		
		out = new FileOutputStream(LOCAL_DIRECTORY+File.separator+TargetFile);
		System.out.println("out:"+out );			
		if(!ftp.retrieveFile(TargetFile,out)){
			System.out.println("下載傳檔失敗,請將"+TargetFile+"重新傳送!!");
			msgtext="下載傳檔失敗,請將"+TargetFile+"重新傳送";
			send_Mail("",msgtext);
			return iResult;
		}				
		//out.close();
		
		ftp.logout();
		ftp.disconnect();
		iResult = 1;		
    }catch(Exception e){
        System.out.println("ftpgetFile ERROR:"+e.getMessage());
    }finally{
    	try{
    	if(out != null) out.close();    	 
    	}catch(IOException ioe){
    		System.out.println("ftpgetFile finally ERROR:"+ioe.getMessage());
    	}
    	out = null;
    }
    return iResult;
}
//103.05.19 add SFTP
private static int sftpGetFile(String FTP_SERVER, String FTP_USER,String FTP_PASSWORD, String FTP_DIRECTORY, String LOCAL_DIRECTORY, String TargetFile){
    int iResult=-1;
    String msgtext="";   
    StandardFileSystemManager manager = new StandardFileSystemManager();
    try{
        manager.init();
         
        //Setup our SFTP configuration
        FileSystemOptions opts = new FileSystemOptions();
        SftpFileSystemConfigBuilder.getInstance().setStrictHostKeyChecking(opts, "no");
        SftpFileSystemConfigBuilder.getInstance().setUserDirIsRoot(opts, true);
        SftpFileSystemConfigBuilder.getInstance().setTimeout(opts, 10000);
         
        //Create the SFTP URI using the host name, userid, password,  remote path and file name
        String sftpUri = "sftp://" + FTP_USER + ":" + FTP_PASSWORD +  "@" + FTP_SERVER + "/" +
                FTP_DIRECTORY + "/" + TargetFile;
        //check local dir, if not exist then make local dir
        File local_dir = new File(LOCAL_DIRECTORY);
        if (!local_dir.exists()) {
            local_dir.mkdirs();
        } 
        // Create local file object
        String filepath = LOCAL_DIRECTORY+File.separator+TargetFile;
        File file = new File(filepath);
        FileObject localFile = manager.resolveFile(file.getAbsolutePath());
        System.out.println("localFile="+localFile);
        // Create remote file object
        FileObject remoteFile = manager.resolveFile(sftpUri, opts);
        //下載檔案
        System.out.println("下載檔案開始:" ); 
        System.out.println("下載位置:" +LOCAL_DIRECTORY+File.separator+TargetFile);    
        // Copy local file to sftp server
        localFile.copyFrom(remoteFile, Selectors.SELECT_SELF);
        System.out.println(" File download successful");       
        
        iResult = 1;        
    }catch(Exception e){
        System.out.println("sftpGetFile ERROR:"+e.getMessage());        
        msgtext="下載傳檔失敗,請將"+TargetFile+"重新傳送";
        send_Mail("",msgtext);
        return iResult;
    }finally{
        try{
            manager.close();  
        }catch(Exception ex){
            System.out.println("sftpGetFile finally ERROR:"+ex.getMessage());
        }        
    }
    return iResult;
}

//農漁會金融卡發卡情形及ATM裝設情形統計
private static List parseATM(String getfr,String M_YEAR,String M_MONTH,String BANK_NO){     
      String Month_Tran_AMT="",Year_AccTran_AMT="",PUSH_DebitCard_CNT="",USE_DebitCard_CNT="",CANC_DebitCard_CNT="",PUSH_BinCard_CNT="",USE_BinCard_CNT="",CANC_BinCard_CNT="",ATM_CNT="",Month_Tran_CNT="",Year_AccTran_CNT="";
      String sqlCmd = "";	
      String DebitCard_Month_Tran_AMT="",DebitCard_Year_AccTran_AMT="",DebitCard_Month_Tran_CNT="",DebitCard_Year_AccTran_CNT="";
      List paramList =new ArrayList() ;      
      try{
          PUSH_DebitCard_CNT  = Integer.toString(Integer.parseInt(getfr.substring(14,21)));//磁條金融卡發行張數:數字（15-21）
          USE_DebitCard_CNT  = Integer.toString(Integer.parseInt(getfr.substring(21,28)));//磁條金融卡流通張數:數字（22-28）
          CANC_DebitCard_CNT = Integer.toString(Integer.parseInt(getfr.substring(28,35)));//磁條金融卡停卡張數:數字（29-35）
          PUSH_BinCard_CNT  = Integer.toString(Integer.parseInt(getfr.substring(35,42)));//晶片卡發行張數:數字（36-42）
          USE_BinCard_CNT  = Integer.toString(Integer.parseInt(getfr.substring(42,49)));//晶片卡流通張數:數字（43-49）//96.01.03 fix 流通張數.位數
          CANC_BinCard_CNT = Integer.toString(Integer.parseInt(getfr.substring(49,56)));//晶片卡停卡張數:數字（50-56）
          ATM_CNT  = Integer.toString(Integer.parseInt(getfr.substring(56,61)));//ATM裝設台數:數字（57-61）
          Month_Tran_CNT  = Integer.toString(Integer.parseInt(getfr.substring(61,68)));//ATM本月交易次數:數字（62-68）          
          Month_Tran_AMT  = Integer.toString(Integer.parseInt(getfr.substring(68,82)));//ATM本月交易金額 (元):數字（69-82）
          DebitCard_Month_Tran_CNT  = Integer.toString(Integer.parseInt(getfr.substring(82,89)));//金融卡本月交易次數:數字（83-89）          
          DebitCard_Month_Tran_AMT  = Integer.toString(Integer.parseInt(getfr.substring(89,103)));//金融卡本月交易金額 (元):數字（90-103）
                    
//        Year_AccTran_CNT  = Integer.toString(Integer.parseInt(getfr.substring(68,77)));//本年累計交易次數:數字（69-77）
          
          //String str=getfr.substring(91,105);//本年累計交易金額 (元):數字（92-105）
          //Year_AccTran_AMT=str.trim();	
          //96.02.07 避免sql injection by 2295 ==================================================
          M_YEAR = M_YEAR.replaceAll("'","''");
      	  M_MONTH = M_MONTH.replaceAll("'","''");
      	  BANK_NO = BANK_NO.replaceAll("'","''");
          PUSH_DebitCard_CNT = PUSH_DebitCard_CNT.replaceAll("'","''");
		  USE_DebitCard_CNT  = USE_DebitCard_CNT.replaceAll("'","''");
		  CANC_DebitCard_CNT = CANC_DebitCard_CNT.replaceAll("'","''");
		  PUSH_BinCard_CNT   = PUSH_BinCard_CNT.replaceAll("'","''");
		  USE_BinCard_CNT    = USE_BinCard_CNT.replaceAll("'","''");
		  CANC_BinCard_CNT   = CANC_BinCard_CNT.replaceAll("'","''");
		  ATM_CNT            = ATM_CNT.replaceAll("'","''");
		  Month_Tran_CNT     = Month_Tran_CNT.replaceAll("'","''");//ATM本月交易次數
		  Month_Tran_AMT     = Month_Tran_AMT.replaceAll("'","''");//ATM本月交易金額 (元)
		  DebitCard_Month_Tran_CNT = DebitCard_Month_Tran_CNT.replaceAll("'","''");//金融卡本月交易次數
		  DebitCard_Month_Tran_AMT = DebitCard_Month_Tran_AMT.replaceAll("'","''");//金融卡本月交易金額 (元)
		  //Year_AccTran_CNT   = Year_AccTran_CNT.replaceAll("'","''");
		  
		  //Year_AccTran_AMT   = Year_AccTran_AMT.replaceAll("'","''");
          //====================================================================================	  
          System.out.print(":PUSH_DebitCard_CNT="+PUSH_DebitCard_CNT); 
          System.out.print(":USE_DebitCard_CNT="+USE_DebitCard_CNT);
          System.out.print(":CANC_DebitCard_CNT="+CANC_DebitCard_CNT);
          System.out.print(":PUSH_BinCard_CNT="+PUSH_BinCard_CNT); 
          System.out.print(":USE_BinCard_CNT="+USE_BinCard_CNT);
          System.out.print(":CANC_BinCard_CNT="+CANC_BinCard_CNT);
          System.out.print(":ATM_CNT="+ATM_CNT); 
          System.out.print(":Month_Tran_CNT="+Month_Tran_CNT);
          System.out.print(":Month_Tran_AMT="+Month_Tran_AMT);
          System.out.print(":DebitCard_Month_Tran_CNT="+DebitCard_Month_Tran_CNT);
          System.out.print(":DebitCard_Month_Tran_AMT="+DebitCard_Month_Tran_AMT);
          //System.out.print(":Year_AccTran_CNT="+Year_AccTran_CNT);          
          
          //System.out.println(":Year_AccTran_AMT="+Year_AccTran_AMT);
          /*
          sqlCmd = "INSERT INTO WLX05_M_ATM VALUES('"+ M_YEAR+ "','"+M_MONTH+"','"+BANK_NO+"'"
   		  +",'"+PUSH_DebitCard_CNT+"','"+USE_DebitCard_CNT+"','"+PUSH_BinCard_CNT+"','"+USE_BinCard_CNT+"','"+ATM_CNT+"','"
          + Month_Tran_CNT+"','','"+Month_Tran_AMT+"','','A111111111"+"','A111111111',sysdate,'"+CANC_DebitCard_CNT+"','"+CANC_BinCard_CNT+"','"
          + DebitCard_Month_Tran_CNT+"','','"+DebitCard_Month_Tran_AMT+"','')";
          */	
          	
          paramList.add(M_YEAR);
          paramList.add(M_MONTH);
          paramList.add(BANK_NO);
          paramList.add(PUSH_DebitCard_CNT);
          paramList.add(USE_DebitCard_CNT);
          paramList.add(PUSH_BinCard_CNT);
          paramList.add(USE_BinCard_CNT);
          paramList.add(ATM_CNT);
          paramList.add(Month_Tran_CNT);
          paramList.add(Month_Tran_AMT);
          paramList.add(CANC_DebitCard_CNT);
          paramList.add(CANC_BinCard_CNT);
          paramList.add(DebitCard_Month_Tran_CNT);
          paramList.add(DebitCard_Month_Tran_AMT);          
          
      }catch(Exception e){
          System.out.println("parseATM Error:"+e.getMessage());
      }
      return paramList;
}

//96.11.22 統計ATM交易次數/交易金額;金融卡交易次數/交易金額本年度累計
private static int updateSumYear(String m_year,String m_month){
	String sqlCmd = "",bank_no="";
	String DebitCard_Year_AccTran_AMT="",DebitCard_Year_AccTran_CNT="";
	String Year_AccTran_AMT="",Year_AccTran_CNT="";
	DataObject bean = null;	
	int returnCode = -1;
	
	List paramList =new ArrayList() ;
    List updateDBList = new ArrayList();//0:sql 1:data
    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
	List updateDBDataList = new ArrayList();//儲存參數的List
	try{
		if(Integer.parseInt(m_month) > 1){
		    sqlCmd = " select wlx05_m_atm.bank_no,sum1.year_acctran_cnt,sum1.year_acctran_amt,"
		    	   + " 		  sum1.DebitCard_year_acctran_cnt,sum1.DebitCard_year_acctran_amt "
		    	   + " from wlx05_m_atm left join ( select bank_no,"
	        	   + "		  				      		   sum(month_tran_cnt) as year_acctran_cnt,"
	        	   + " 		  							   sum(month_tran_amt) as year_acctran_amt,"
	        	   + "		  							   sum(DebitCard_month_tran_cnt) as DebitCard_year_acctran_cnt,"
	        	   + " 		  							   sum(DebitCard_month_tran_amt) as DebitCard_year_acctran_amt"
		    	   + " 						        from WLX05_M_ATM " 
		    	   + " 								where to_char(m_year * 100 + m_month) >= ?"
		    	   + " 								and to_char(m_year * 100 + m_month) <= ?"
		    	   + " 							    group by bank_no )sum1 on wlx05_m_atm.bank_no = sum1.bank_no "
		    	   + " where wlx05_m_atm.m_year=?"
		    	   + " and wlx05_m_atm.m_month=?";
		    paramList.add(m_year+"01");
		    paramList.add(m_year+DownLoad.fillStuff(m_month, "L", "0", 2));		    
		    paramList.add(m_year);
		    paramList.add(m_month);
		    
		    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"year_acctran_cnt,year_acctran_amt,debitcard_year_acctran_cnt,debitcard_year_acctran_amt");  
		    for(int i=0;i<dbData.size();i++){
		    	 bean = (DataObject)dbData.get(i);
		    	 bank_no = (String)bean.getValue("bank_no");
		    	 Year_AccTran_CNT = bean.getValue("year_acctran_cnt").toString();//ATM本年累計交易次數
		    	 Year_AccTran_AMT = bean.getValue("year_acctran_amt").toString();//ATM本年累計交易金額
		    	 DebitCard_Year_AccTran_CNT = bean.getValue("debitcard_year_acctran_cnt").toString();//金融卡本年累計交易次數
		    	 DebitCard_Year_AccTran_AMT = bean.getValue("debitcard_year_acctran_amt").toString();//金融卡本年累計交易金額
		    	 /*
		    	 sqlCmd = " update WLX05_M_ATM"
		    	 	    + " set year_acctran_cnt='"+Year_AccTran_CNT+"',"
			   		    + "     year_acctran_amt='"+Year_AccTran_AMT+"',"
			   		    + "     debitcard_year_acctran_cnt='"+DebitCard_Year_AccTran_CNT+"',"
			   		    + "     debitcard_year_acctran_amt='"+DebitCard_Year_AccTran_AMT+"'"
			   		    + " where bank_no='"+bank_no+"'"
			   		    + " and   m_year="+m_year
			   		    + " and   m_month="+m_month;
			   	 */	    
		    	 //updateDBList.add(sqlCmd);
		    	 sqlCmd = " update WLX05_M_ATM"
	    	 	    	+ " set year_acctran_cnt=?,"
						+ "     year_acctran_amt=?,"
						+ "     debitcard_year_acctran_cnt=?,"
						+ "     debitcard_year_acctran_amt=?"
						+ " where bank_no=?"
						+ " and   m_year=?"
						+ " and   m_month=?";
		    	 paramList = new ArrayList();
		    	 paramList.add(Year_AccTran_CNT);
		    	 paramList.add(Year_AccTran_AMT);
		    	 paramList.add(DebitCard_Year_AccTran_CNT);
		    	 paramList.add(DebitCard_Year_AccTran_AMT);
		    	 paramList.add(bank_no);
		    	 paramList.add(m_year);
		    	 paramList.add(m_month);
		    	 updateDBDataList.add(paramList);
		    }
		 }else{//當每年一月份時.本年累計=當月份交易資料
		 	sqlCmd = " update WLX05_M_ATM "
		 		   + " set year_acctran_cnt=month_Tran_CNT,"
				   + " 	   year_acctran_amt=month_Tran_AMT,"
				   + "     debitcard_year_acctran_cnt=DebitCard_month_Tran_CNT,"
				   + "	   debitcard_year_acctran_amt=DebitCard_month_Tran_AMT"
				   + " where m_year=?"
				   + " and   m_month=?";
		 	paramList = new ArrayList();
		 	paramList.add(m_year);
		 	paramList.add(m_month);
		 	updateDBDataList.add(paramList);		 	
		 }
		 updateDBSqlList.add(sqlCmd);
	  	 updateDBSqlList.add(updateDBDataList);
	  	 updateDBList.add(updateDBSqlList);
		 if(DBManager.updateDB_ps(updateDBList)){
		 	returnCode = 0;
		 }
	}catch(Exception e){
		System.out.println("updateSumYear Error:"+e+e.getMessage());
	}
	return returnCode;
}
//農漁會支票存款資料
private static List parseCHK(String getfr,String M_YEAR,String M_MONTH,String BANK_NO){    	      	
	  String CheckBank_Bal="",CheckBank_Bal_S="",CheckBank_Bal_N="",CheckBank_Cnt="",CheckBank_Cnt_S="",CheckBank_Cnt_N="";
	  String sqlCmd = "";
	  List paramList =new ArrayList() ;    
	  //ex:22006085030019000010200000002046574000001300000001937903000003400000004941317
	  //M_YEAR  = Integer.toString(Integer.parseInt(getfr.substring(1,5))-1911);
      //M_MONTH = Integer.toString(Integer.parseInt(getfr.substring(5,7)));
	  //BANK_NO = getfr.substring(7,14);  
	  try{
	      CheckBank_Cnt  = Integer.toString(Integer.parseInt(getfr.substring(14,21)));//正會員支票存款戶數:數字（15-21）
	      CheckBank_Bal  = Integer.toString(Integer.parseInt(getfr.substring(21,35)));//正會員支票存款餘額 (元):數字（22-35）
	      CheckBank_Cnt_S  = Integer.toString(Integer.parseInt(getfr.substring(35,42)));//贊助會員支票存款戶數:數字（36-42）
	      CheckBank_Bal_S  = Integer.toString(Integer.parseInt(getfr.substring(42,56)));//贊助會員支票存款餘額 (元):數字（43-56）
	      CheckBank_Cnt_N  = Integer.toString(Integer.parseInt(getfr.substring(56,63)));//非會員支票存款戶數:數字（57-63）
	      CheckBank_Bal_N  = Integer.toString(Integer.parseInt(getfr.substring(63,77)));//非會員支票存款餘額 (元):數字（64-77）
	      //96.02.07 避免sql injection by 2295 ==================================================
	      M_YEAR = M_YEAR.replaceAll("'","''");
      	  M_MONTH = M_MONTH.replaceAll("'","''");
      	  BANK_NO = BANK_NO.replaceAll("'","''");
	      CheckBank_Cnt = CheckBank_Cnt.replaceAll("'","''");
	      CheckBank_Bal = CheckBank_Bal.replaceAll("'","''");
	      CheckBank_Cnt_S = CheckBank_Cnt_S.replaceAll("'","''");
	      CheckBank_Bal_S = CheckBank_Bal_S.replaceAll("'","''");
	      CheckBank_Cnt_N = CheckBank_Cnt_N.replaceAll("'","''");
	      CheckBank_Bal_N = CheckBank_Bal_N.replaceAll("'","''");		  
          //====================================================================================	  
         
	      System.out.print(":CheckBank_Cnt="+CheckBank_Cnt); 
	      System.out.print(":CheckBank_Bal="+CheckBank_Bal);
	      System.out.print(":CheckBank_Cnt_S="+CheckBank_Cnt_S); 
	      System.out.print(":CheckBank_Bal_S="+CheckBank_Bal_S);
	      System.out.print(":CheckBank_Cnt_N="+CheckBank_Cnt_N); 
	      System.out.println(":CheckBank_Bal_N="+CheckBank_Bal_N);
	      /*
	      sqlCmd = "INSERT INTO WLX07_M_CHECKBANK VALUES('"
	          	 + M_YEAR + "','" + M_MONTH +"','"+ BANK_NO +"','" + CheckBank_Cnt +"','" + CheckBank_Bal +"','" + CheckBank_Cnt_S +"','"
	          	 + CheckBank_Bal_S +"','"+ CheckBank_Cnt_N +"','" + CheckBank_Bal_N +"','A111111111','A111111111',sysdate)";    
          */
	     
	      paramList.add(M_YEAR);
	      paramList.add(M_MONTH);
	      paramList.add(BANK_NO);
	      paramList.add(CheckBank_Cnt);
	      paramList.add(CheckBank_Bal);
	      paramList.add(CheckBank_Cnt_S);
	      paramList.add(CheckBank_Bal_S);
	      paramList.add(CheckBank_Cnt_N);
	      paramList.add(CheckBank_Bal_N);
		  
	  }catch(Exception e){
        System.out.println("parseCHK Error:"+e.getMessage());
	  }
      return paramList;
}

//農漁會信用部牌告利率申報資料
private static List parseINT(String getfr,String M_YEAR,String M_MONTH,String BANK_NO){
      String Period_1_FIX_Rate="",Period_1_Var_Rate="",Period_3_FIX_Rate="",Period_3_Var_Rate="",Period_6_FIX_Rate="",Period_6_Var_Rate="",Period_9_FIX_Rate="",Period_9_Var_Rate="",Period_12_FIX_Rate="",Period_12_Var_Rate="",Basic_Pay_Var_Rate="",Period_House_Var_Rate="",Base_Mark_Rate="",Base_Fix_Rate="",Base_Base_Rate="";
      //96.09.28增加的利率==================================================================================
      String Period_24_FIX_Rate="",Period_24_Var_Rate="",Period_36_FIX_Rate="",Period_36_Var_Rate="";
      String Deposit_12_FIX_Rate="",Deposit_12_Var_Rate="",Deposit_24_FIX_Rate="",Deposit_24_Var_Rate="",Deposit_36_FIX_Rate="",Deposit_36_Var_Rate="";
      String Deposit_Var_Rate,Save_Var_Rate="";
      //98.03.30增加基準利率-指標利率（月調）/基準利率（月調）/指數型房貸指標利率（月調）=======================================================================================================
      String Base_Mark_Rate_Month="";//基準利率-指標利率（月調）
	  String Base_Base_Rate_Month="";//基準利率（月調）
	  String Period_House_Var_Rate_Month="";//指數型房貸指標利率（月調）
	  //=============================================================================================================
      String sqlCmd = "";
      List paramList =new ArrayList() ;    
      //ex:22006045030019017050170501785017850194501925020750205502200021550744002180021800150003680
      //M_YEAR  = Integer.toString(Integer.parseInt(getfr.substring(1,5))-1911);//年度:數字（2-5）
      //M_Quarter = Integer.toString(Integer.parseInt(getfr.substring(5,7)));//月份:數字（6-7）
      //BANK_NO = getfr.substring(7,14);//金融總機構代號:文數字（8-14）	    
	  try{
		  //96.01.03 fix 牌告利率取到小數後3位 by 2295
	  	  Deposit_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(22,27))/1000);//活期存款機動利率:數字（23-27）
	  	  Save_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(27,32))/1000);//活期儲蓄存款機動利率:數字（28-32）
	  	  System.out.println("活期存款機動利率="+Deposit_Var_Rate);
	  	  System.out.println("活期儲蓄存款機動利率="+Save_Var_Rate);
	  	  //機動利率
	  	  Period_1_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(32,37))/1000);//定期存款-1個月-機動:數字（33-37）
		  Period_3_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(37,42))/1000);//定期存款-3個月-機動:數字（38-42）		  
		  Period_6_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(42,47))/1000);//定期存款-6個月-機動:數字（43-47）		  
		  Period_9_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(47,52))/1000);//定期存款-9個月-機動:數字（48-52）		  
		  Period_12_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(52,57))/1000);//定期存款-12個月-機動:數字（53-57）
		  Period_24_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(57,62))/1000);//定期存款-24個月-機動:數字（58-62）
		  Period_36_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(62,67))/1000);//定期存款-36個月-機動:數字（63-67）
		  
		  Deposit_12_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(67,72))/1000);//定期儲蓄存款-12個月_機動:數字（68-72）
		  Deposit_24_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(72,77))/1000);//定期儲蓄存款-24個月_機動:數字（73-77）
		  Deposit_36_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(77,82))/1000);//定期儲蓄存款-36個月_機動:數字（78-82）
		  
	  	  //固定利率	  	  
		  Period_1_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(82,87))/1000);//定期存款-1個月-固定:數字（83-87）		  
		  Period_3_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(87,92))/1000);//定期存款-3個月-固定:數字（88-92）
		  Period_6_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(92,97))/1000);//定期存款-6個月-固定:數字（93-97）
		  Period_9_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(97,102))/1000);//定期存款-9個月-固定:數字（98-102）
		  Period_12_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(102,107))/1000);//定期存款-12個月-固定:數字（103-107）
		  Period_24_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(107,112))/1000);//定期存款-24個月-固定:數字（108-112）
		  Period_36_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(112,117))/1000);//定期存款-36個月-固定:數字（113-117）
		  
		  Deposit_12_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(117,122))/1000);//定期儲蓄存款-12個月_固定:數字（118-122）
		  Deposit_24_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(122,127))/1000);//定期儲蓄存款-24個月_固定:數字（123-127）
		  Deposit_36_FIX_Rate  = Double.toString(Double.parseDouble(getfr.substring(127,132))/1000);//定期儲蓄存款-36個月_固定:數字（128-132）
		  		  	  
		  Basic_Pay_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(132,137))/1000);//基本放款利率（機動）:數字（133-137）
		  Period_House_Var_Rate  = Double.toString(Double.parseDouble(getfr.substring(137,142))/1000);//指數型房貸指標利率:數字（138-142）
		  Base_Mark_Rate   = Double.toString(Double.parseDouble(getfr.substring(142,147))/1000);//基準利率-指標利率(1):數字（143-147）
		  Base_Fix_Rate  = Double.toString(Double.parseDouble(getfr.substring(147,152))/1000);//基準利率-一定比率(2):數字（148-152）
		  Base_Base_Rate  = Double.toString(Double.parseDouble(getfr.substring(152,157))/1000);//基準利率=(1)+(2):數字（153-157）		  
	      //98.03.30 add =====================================================================================================================
		  Base_Mark_Rate_Month  = Double.toString(Double.parseDouble(getfr.substring(157,162))/1000);//基準利率-指標利率（月調）:數字（158-162）
		  Base_Base_Rate_Month  = Double.toString(Double.parseDouble(getfr.substring(162,167))/1000);//基準利率（月調）:數字（163-167）
		  Period_House_Var_Rate_Month  = Double.toString(Double.parseDouble(getfr.substring(167,172))/1000);//指數型房貸指標利率（月調）:數字（168-172）
		  //96.02.07 避免sql injection by 2295 ==================================================
		  M_YEAR = M_YEAR.replaceAll("'","''");
      	  M_MONTH = M_MONTH.replaceAll("'","''");
      	  BANK_NO = BANK_NO.replaceAll("'","''");
      	  Deposit_Var_Rate = Deposit_Var_Rate.replaceAll("'","''");
      	  Save_Var_Rate = Save_Var_Rate.replaceAll("'","''");
		  Period_1_FIX_Rate = Period_1_FIX_Rate.replaceAll("'","''");
		  Period_1_Var_Rate = Period_1_Var_Rate.replaceAll("'","''");
		  Period_3_FIX_Rate = Period_3_FIX_Rate.replaceAll("'","''");
		  Period_3_Var_Rate = Period_3_Var_Rate.replaceAll("'","''");
		  Period_6_FIX_Rate = Period_6_FIX_Rate.replaceAll("'","''");
		  Period_6_Var_Rate = Period_6_Var_Rate.replaceAll("'","''");
		  Period_9_FIX_Rate = Period_9_FIX_Rate.replaceAll("'","''");
		  Period_9_Var_Rate = Period_9_Var_Rate.replaceAll("'","''");
		  Period_12_FIX_Rate = Period_12_FIX_Rate.replaceAll("'","''");
		  Period_12_Var_Rate = Period_12_Var_Rate.replaceAll("'","''");
		  Period_24_FIX_Rate = Period_24_FIX_Rate.replaceAll("'","''");
		  Period_24_Var_Rate = Period_24_Var_Rate.replaceAll("'","''");
		  Period_36_FIX_Rate = Period_36_FIX_Rate.replaceAll("'","''");
		  Period_36_Var_Rate = Period_36_Var_Rate.replaceAll("'","''");
		  Deposit_12_FIX_Rate = Deposit_12_FIX_Rate.replaceAll("'","''");
		  Deposit_12_Var_Rate = Deposit_12_Var_Rate.replaceAll("'","''");
		  Deposit_24_FIX_Rate = Deposit_24_FIX_Rate.replaceAll("'","''");
		  Deposit_24_Var_Rate = Deposit_24_Var_Rate.replaceAll("'","''");
		  Deposit_36_FIX_Rate = Deposit_36_FIX_Rate.replaceAll("'","''");
		  Deposit_36_Var_Rate = Deposit_36_Var_Rate.replaceAll("'","''");
		  Basic_Pay_Var_Rate = Basic_Pay_Var_Rate.replaceAll("'","''");
		  Period_House_Var_Rate = Period_House_Var_Rate.replaceAll("'","''");
		  Base_Mark_Rate = Base_Mark_Rate.replaceAll("'","''");
		  Base_Fix_Rate = Base_Fix_Rate.replaceAll("'","''");
		  Base_Base_Rate = Base_Base_Rate.replaceAll("'","''");
          //98.03.30 add ====================================================================================
		  Base_Mark_Rate_Month = Base_Mark_Rate_Month.replaceAll("'","''");;//基準利率-指標利率（月調）
		  Base_Base_Rate_Month = Base_Base_Rate_Month.replaceAll("'","''");;//基準利率（月調）
		  Period_House_Var_Rate_Month = Period_House_Var_Rate_Month.replaceAll("'","''");;//指數型房貸指標利率（月調）
          //====================================================================================
		  System.out.print(":Deposit_Var_Rate="+Deposit_Var_Rate); 
		  System.out.print(":Save_Var_Rate="+Save_Var_Rate);
		  System.out.print(":Period_1_FIX_Rate="+Period_1_FIX_Rate); 
		  System.out.print(":Period_1_Var_Rate="+Period_1_Var_Rate);		      		
		  System.out.print(":Period_3_FIX_Rate="+Period_3_FIX_Rate);
		  System.out.print(":Period_3_Var_Rate="+Period_3_Var_Rate);
		  System.out.print(":Period_6_FIX_Rate="+Period_6_FIX_Rate);
		  System.out.print(":Period_6_Var_Rate="+Period_6_Var_Rate); 
		  System.out.print(":Period_9_FIX_Rate="+Period_9_FIX_Rate);
		  System.out.print(":Period_9_Var_Rate="+Period_9_Var_Rate);
		  System.out.print(":Period_12_FIX_Rate="+Period_12_FIX_Rate);
		  System.out.print(":Period_12_Var_Rate="+Period_12_Var_Rate);
		  System.out.print(":Period_24_FIX_Rate="+Period_24_FIX_Rate);
		  System.out.print(":Period_24_Var_Rate="+Period_24_Var_Rate);  
		  System.out.print(":Period_36_FIX_Rate="+Period_36_FIX_Rate);
		  System.out.print(":Period_36_Var_Rate="+Period_36_Var_Rate);  
		  System.out.print(":Deposit_12_FIX_Rate="+Deposit_12_FIX_Rate);
		  System.out.print(":Deposit_12_Var_Rate="+Deposit_12_Var_Rate);
		  System.out.print(":Deposit_24_FIX_Rate="+Deposit_24_FIX_Rate);
		  System.out.print(":Deposit_24_Var_Rate="+Deposit_24_Var_Rate);  
		  System.out.print(":Deposit_36_FIX_Rate="+Deposit_36_FIX_Rate);
		  System.out.print(":Deposit_36_Var_Rate="+Deposit_36_Var_Rate);  
		  System.out.print(":Basic_Pay_Var_Rate="+Basic_Pay_Var_Rate);
		  System.out.print(":Period_House_Var_Rate="+Period_House_Var_Rate); 
		  System.out.print(":Base_Mark_Rate="+Base_Mark_Rate);
		  System.out.print(":Base_Fix_Rate="+Base_Fix_Rate);
		  System.out.println(":Base_Base_Rate="+Base_Base_Rate);
		  System.out.println(":Base_Mark_Rate_Month="+Base_Mark_Rate_Month);
		  System.out.println(":Base_Base_Rate_Month="+Base_Base_Rate_Month);
		  System.out.println(":Period_House_Var_Rate_Month="+Period_House_Var_Rate_Month);
		  /*
		  sqlCmd = "INSERT INTO WLX_S_RATE VALUES('"
				   + M_YEAR + "','" + M_MONTH +"','"+ BANK_NO +"','" + Period_1_FIX_Rate +"','" + Period_1_Var_Rate +"','" + Period_3_FIX_Rate +"','"
			       + Period_3_Var_Rate +"','"+ Period_6_FIX_Rate +"','" + Period_6_Var_Rate +"','" + Period_9_FIX_Rate +"','" 
			       + Period_9_Var_Rate +"','"+ Period_12_FIX_Rate +"','" + Period_12_Var_Rate+"','" + Basic_Pay_Var_Rate +"','" + Period_House_Var_Rate +"','" 
			       + Base_Mark_Rate +"','"+ Base_Fix_Rate +"','" + Base_Base_Rate +"','A111111111','A111111111',sysdate,"
				   + "'"+ Period_24_FIX_Rate +"','" + Period_24_Var_Rate+"','" + Period_36_FIX_Rate +"','" + Period_36_Var_Rate+"','"
				   + Deposit_12_FIX_Rate +"','" + Deposit_12_Var_Rate+"','" + Deposit_24_FIX_Rate +"','" + Deposit_24_Var_Rate+"','" + Deposit_36_FIX_Rate +"','" + Deposit_36_Var_Rate+"','"
				   + Deposit_Var_Rate +"','" + Save_Var_Rate+"','"+Base_Mark_Rate_Month+"','"+Base_Base_Rate_Month+"','"+Period_House_Var_Rate_Month+"'"
				   +")";
		  */
		  
		  paramList.add(M_YEAR);    
		  paramList.add(M_MONTH);
		  paramList.add(BANK_NO);
		  paramList.add(Period_1_FIX_Rate);
		  paramList.add(Period_1_Var_Rate);
		  paramList.add(Period_3_FIX_Rate);
		  paramList.add(Period_3_Var_Rate);
		  paramList.add(Period_6_FIX_Rate);
		  paramList.add(Period_6_Var_Rate);
		  paramList.add(Period_9_FIX_Rate);
		  paramList.add(Period_9_Var_Rate);
		  paramList.add(Period_12_FIX_Rate);
		  paramList.add(Period_12_Var_Rate);
		  paramList.add(Basic_Pay_Var_Rate);
		  paramList.add(Period_House_Var_Rate);
		  paramList.add(Base_Mark_Rate);
		  paramList.add(Base_Fix_Rate);
		  paramList.add(Base_Base_Rate);
		  paramList.add(Period_24_FIX_Rate);
		  paramList.add(Period_24_Var_Rate);
		  paramList.add(Period_36_FIX_Rate);
		  paramList.add(Period_36_Var_Rate);
		  paramList.add(Deposit_12_FIX_Rate);
		  paramList.add(Deposit_12_Var_Rate);
		  paramList.add(Deposit_24_FIX_Rate);
		  paramList.add(Deposit_24_Var_Rate);
		  paramList.add(Deposit_36_FIX_Rate);
		  paramList.add(Deposit_36_Var_Rate);
		  paramList.add(Deposit_Var_Rate);
		  paramList.add(Save_Var_Rate);
		  paramList.add(Base_Mark_Rate_Month);
		  paramList.add(Base_Base_Rate_Month);
		  paramList.add(Period_House_Var_Rate_Month);
	  }catch(Exception e){
      System.out.println("parseINT Error:"+e.getMessage());
	  }
    return paramList;
}


//農漁會在台無住所外國人新台幣存款資料
private static void parseFGN(String getfr,String M_YEAR,String M_MONTH,String BANK_NO,List updateDBDataList_F01,List updateDBDataList_WML01){                 
	  String ACCT_CNT_TM="",DEP_TYPE="",ACCT_TYPE="",BAL_LM="",DEP_TM="",WTD_TM="",BAL_TM="";
	  String sqlCmd = "";
	  List paramList =new ArrayList() ;    
      //ex:22006086060019A1000002100000000062409000000000009500000000003605900000000027300
      //M_YEAR  = Integer.toString(Integer.parseInt(getfr.substring(1,5))-1911);
      //M_MONTH = Integer.toString(Integer.parseInt(getfr.substring(5,7)));
	  //BANK_NO = getfr.substring(7,14);
      try{
          if(!preBANK_NO.equals(BANK_NO)){                   
              System.out.println("preBANK_NO="+preBANK_NO);
              System.out.println("BANK_NO="+BANK_NO);
              for(int i=1;i<ACCT_CNT_TM_idx.length;i++){
              	  /*
              	  sqlCmd = "INSERT INTO F01 VALUES('" + M_YEAR + "','" + M_MONTH +"','"+ preBANK_NO +"','E','"+i+"','" + Double.toString(ACCT_CNT_TM_idx[i]) +"','"
	                     + Double.toString(BAL_LM_idx[i]) +"','"+ Double.toString(DEP_TM_idx[i]) +"','" + Double.toString(WTD_TM_idx[i]) +"','" + Double.toString(BAL_TM_idx[i]) +"')";
	              */                         
                  
                  paramList = new ArrayList();
                  paramList.add(M_YEAR);
                  paramList.add(M_MONTH);
                  paramList.add(preBANK_NO);
                  paramList.add("E");
                  paramList.add(String.valueOf(i));
                  paramList.add(Double.toString(ACCT_CNT_TM_idx[i]));
                  paramList.add(Double.toString(BAL_LM_idx[i]));
                  paramList.add(Double.toString(DEP_TM_idx[i]));
                  paramList.add(Double.toString(WTD_TM_idx[i]));
                  paramList.add(Double.toString(BAL_TM_idx[i]));
                  updateDBDataList_F01.add(paramList);
                  
                  ACCT_CNT_TM_idx[i]=0.0;                          
      		      BAL_LM_idx[i]=0.0;
      		      DEP_TM_idx[i]=0.0;
      		      WTD_TM_idx[i]=0.0;
      		      BAL_TM_idx[i]=0.0;
                  
              }
              //96.02.07 避免sql injection by 2295 ==================================================
    		  M_YEAR = M_YEAR.replaceAll("'","''");
          	  M_MONTH = M_MONTH.replaceAll("'","''");
          	  preBANK_NO = preBANK_NO.replaceAll("'","''");
          	  //====================================================================================
              /*
      	      sqlCmd = "INSERT INTO WML01 VALUES('"
      			     + M_YEAR + "','" + M_MONTH +"','" + preBANK_NO +"','F01','" + INPUT_METHOD +"','A111111111','A111111111',sysdate,'" 
      			     + COMMON_CENTER +"','" + UPD_METHOD +"','" + UPD_CODE +"','" + BATCH_NO +"','" + LOCK +"','A111111111','A111111111',sysdate)";
      	      */
          	  paramList = new ArrayList();
      	      paramList.add(M_YEAR);
      	      paramList.add(M_MONTH);
      	      paramList.add(preBANK_NO);
      	      paramList.add(INPUT_METHOD);
      	      paramList.add(COMMON_CENTER);
      	      paramList.add(UPD_METHOD);
      	      paramList.add(UPD_CODE);
      		  paramList.add(BATCH_NO);
      		  paramList.add(LOCK);
      		  updateDBDataList_WML01.add(paramList);
      		  
      	      preBANK_NO = BANK_NO;
      	      System.out.println("preBANK_NO="+preBANK_NO);
          }
      	  
          DEP_TYPE = getfr.substring(14,15);//存款類型:文數字（15-15）
          ACCT_TYPE = getfr.substring(15,16);//帳戶類型:文數字（16-16）
          ACCT_CNT_TM = Double.toString(Double.parseDouble(getfr.substring(16,23)));//本月底戶數:數字（17-23）
          BAL_LM = Double.toString(Double.parseDouble(getfr.substring(23,37)));//上月底餘額 (元):數字（24-37）
          DEP_TM = Double.toString(Double.parseDouble(getfr.substring(37,51)));//本月存入金額 (元):數字（38-51）
          WTD_TM = Double.toString(Double.parseDouble(getfr.substring(51,65)));//本月提出金額 (元):數字（52-65）
          BAL_TM = Double.toString(Double.parseDouble(getfr.substring(65,79)));//本月底餘額 (元):數字（66-79）
          //本月底戶數.
          ACCT_CNT_TM_idx[Integer.parseInt(ACCT_TYPE)] += Double.parseDouble(getfr.substring(16,23));//合計
          ACCT_CNT_TM_sum += Double.parseDouble(getfr.substring(16,23));//小計
          //上月底餘額 (元)
          BAL_LM_idx[Integer.parseInt(ACCT_TYPE)] += Double.parseDouble(getfr.substring(23,37));
          BAL_LM_sum += Double.parseDouble(getfr.substring(23,37));
          //本月存入金額 (元)
          DEP_TM_idx[Integer.parseInt(ACCT_TYPE)] += Double.parseDouble(getfr.substring(37,51));
          DEP_TM_sum += Double.parseDouble(getfr.substring(37,51));
          //本月提出金額 (元)
          WTD_TM_idx[Integer.parseInt(ACCT_TYPE)] += Double.parseDouble(getfr.substring(51,65));
          WTD_TM_sum += Double.parseDouble(getfr.substring(51,65));
          //本月底餘額 (元)
          BAL_TM_idx[Integer.parseInt(ACCT_TYPE)] += Double.parseDouble(getfr.substring(65,79));
          BAL_TM_sum += Double.parseDouble(getfr.substring(65,79));      
          /*
          System.out.print(":DEP_TYPE=" + DEP_TYPE);
          System.out.print(":ACCT_TYPE=" + ACCT_TYPE);
          System.out.print(":ACCT_CNT_TM=" + ACCT_CNT_TM);
          System.out.print(":BAL_LM=" + BAL_LM);
          System.out.print(":DEP_TM=" +DEP_TM);
          System.out.print(":WTD_TM=" + WTD_TM);
          System.out.println(":BAL_TM=" + BAL_TM);
          */
          //96.02.07 避免sql injection by 2295 ==================================================
		  M_YEAR = M_YEAR.replaceAll("'","''");
      	  M_MONTH = M_MONTH.replaceAll("'","''");
      	  BANK_NO = BANK_NO.replaceAll("'","''");
      	  DEP_TYPE = DEP_TYPE.replaceAll("'","''");
      	  ACCT_TYPE = ACCT_TYPE.replaceAll("'","''");
      	  ACCT_CNT_TM = ACCT_CNT_TM.replaceAll("'","''");
      	  BAL_LM = BAL_LM.replaceAll("'","''");
      	  DEP_TM = DEP_TM.replaceAll("'","''");
      	  WTD_TM = WTD_TM.replaceAll("'","''");
      	  BAL_TM = BAL_TM.replaceAll("'","''");
      	  //====================================================================================
  	      /* 
          sqlCmd = "INSERT INTO F01 VALUES('" + M_YEAR + "','" + M_MONTH +"','"+ BANK_NO +"','" + DEP_TYPE +"','" + ACCT_TYPE +"','" + ACCT_CNT_TM +"','"
      	   		 + BAL_LM +"','"+ DEP_TM +"','" + WTD_TM +"','" + BAL_TM +"')";
      	  */
      
      	  paramList = new ArrayList();
      	  paramList.add(M_YEAR);
      	  paramList.add(M_MONTH);
      	  paramList.add(BANK_NO);
      	  paramList.add(DEP_TYPE);
      	  paramList.add(ACCT_TYPE);
      	  paramList.add(ACCT_CNT_TM);
      	  paramList.add(BAL_LM);
      	  paramList.add(DEP_TM);
      	  paramList.add(WTD_TM);
      	  paramList.add(BAL_TM);
      	  updateDBDataList_F01.add(paramList);
          
          if(ACCT_TYPE.equals("4")){//將個別合計寫入DB
          	 /*
             sqlCmd = "INSERT INTO F01 VALUES('" + M_YEAR + "','" + M_MONTH +"','"+ BANK_NO +"','" + DEP_TYPE +"','5','" + Double.toString(ACCT_CNT_TM_sum) +"','"
         		    + Double.toString(BAL_LM_sum) +"','"+ Double.toString(DEP_TM_sum) +"','" + Double.toString(WTD_TM_sum) +"','" + Double.toString(BAL_TM_sum) +"')";
             */
         
              paramList = new ArrayList();
         	  paramList.add(M_YEAR);
         	  paramList.add(M_MONTH);
         	  paramList.add(BANK_NO);
         	  paramList.add(DEP_TYPE);
         	  paramList.add("5");
         	  paramList.add(Double.toString(ACCT_CNT_TM_sum));
         	  paramList.add(Double.toString(BAL_LM_sum));
         	  paramList.add(Double.toString(DEP_TM_sum));
         	  paramList.add(Double.toString(WTD_TM_sum));
         	  paramList.add(Double.toString(BAL_TM_sum));
         	  updateDBDataList_F01.add(paramList);
         	
             ACCT_CNT_TM_idx[5] += ACCT_CNT_TM_sum;
             BAL_LM_idx[5] +=BAL_LM_sum;
             DEP_TM_idx[5] += DEP_TM_sum;
             WTD_TM_idx[5] += WTD_TM_sum;
             BAL_TM_idx[5] += BAL_TM_sum;
             ACCT_CNT_TM_sum=0.0;BAL_LM_sum =0.0;DEP_TM_sum=0.0;WTD_TM_sum=0.0;BAL_TM_sum=0.0;
          }	  
      }catch(Exception e){
          System.out.println("parseFGN Error:"+e.getMessage());
	  }    
}

//103.04.02 add專案農貸明細資料 
private static List parseFRM(String getfr,String M_YEAR,String M_MONTH,String BANK_NO){     
    String loan_item="",loan_bal_cnt="",loan_bal_amt="",over6m_loan_bal_cnt="",over6m_loan_bal_amt="",delay_loan_cnt="",delay_loan_amt="",over3p_loan_cnt="",over3p_loan_amt="",again_loan_cnt="",again_loan_bal_amt="",again_loan_bal_amt_1th="";
    String sqlCmd = "";   
    List paramList =new ArrayList() ;      
    try{
        
        loan_item           = getfr.substring(14,16);//(5)貸款項目代碼:文字（15-16）  
        loan_bal_cnt        = Integer.toString(Integer.parseInt(getfr.substring(16,22)));//(6)貸放餘額-戶數:數字（17-22）
        loan_bal_amt        = Double.toString(Double.parseDouble(getfr.substring(22,33)));//(7)貸放餘額-金額 :數字（23-33）   
        over6m_loan_bal_cnt = Integer.toString(Integer.parseInt(getfr.substring(33,39)));//(8)逾期六個月放款餘額-戶數:數字（34-39）  
        over6m_loan_bal_amt = Double.toString(Double.parseDouble(getfr.substring(39,50)));//(9)逾期六個月放款餘額-金額:數字（40-50）   
        delay_loan_cnt      = Integer.toString(Integer.parseInt(getfr.substring(50,56)));//(10)當月核准延期還款-件數:數字（51-56）  
        delay_loan_amt      = Double.toString(Double.parseDouble(getfr.substring(56,67)));//(11)當月核准延期還款-金額:數字（57-67）  
        over3p_loan_cnt     = Integer.toString(Integer.parseInt(getfr.substring(67,73)));//(12)同戶3人以上申貸-戶數:數字（68-73）  
        over3p_loan_amt     = Double.toString(Double.parseDouble(getfr.substring(73,84)));//(13)同戶3人以上申貸-金額:數字（74-84）  
        again_loan_cnt      = Integer.toString(Integer.parseInt(getfr.substring(84,90)));//(14)一借再借的-戶數:數字（85-90） 
        again_loan_bal_amt  = Double.toString(Double.parseDouble(getfr.substring(90,101)));//(15)一借再借的-貸放餘額:數字（91-101）  
        again_loan_bal_amt_1th = Double.toString(Double.parseDouble(getfr.substring(101,112)));//(16)一借再借的-首筆貸放餘額:數字（102-112） 
 
        //避免sql injection by 2295 ==================================================
        M_YEAR = M_YEAR.replaceAll("'","''");
        M_MONTH = M_MONTH.replaceAll("'","''");
        BANK_NO = BANK_NO.replaceAll("'","''");
        loan_item = loan_item.replaceAll("'","''");
        loan_bal_cnt = loan_bal_cnt.replaceAll("'","''");
        loan_bal_amt = loan_bal_amt.replaceAll("'","''");
        over6m_loan_bal_cnt = over6m_loan_bal_cnt.replaceAll("'","''");
        over6m_loan_bal_amt = over6m_loan_bal_amt.replaceAll("'","''");
        delay_loan_cnt = delay_loan_cnt.replaceAll("'","''");
        delay_loan_amt =delay_loan_amt.replaceAll("'","''");
        over3p_loan_cnt = over3p_loan_cnt.replaceAll("'","''");
        over3p_loan_amt = over3p_loan_amt.replaceAll("'","''");
        again_loan_cnt = again_loan_cnt.replaceAll("'","''");
        again_loan_bal_amt = again_loan_bal_amt.replaceAll("'","''");
        again_loan_bal_amt_1th = again_loan_bal_amt_1th.replaceAll("'","''");
        //====================================================================================      
        System.out.print(":loan_item="+loan_item); 
        System.out.print(":loan_bal_cnt="+loan_bal_cnt);
        System.out.print(":loan_bal_amt="+loan_bal_amt);
        System.out.print(":over6m_loan_bal_cnt="+over6m_loan_bal_cnt); 
        System.out.print(":over6m_loan_bal_amt="+over6m_loan_bal_amt);
        System.out.print(":delay_loan_cnt="+delay_loan_cnt);
        System.out.print(":delay_loan_amt="+delay_loan_amt); 
        System.out.print(":over3p_loan_cnt="+over3p_loan_cnt);//僅農綜貸(08)須報送
        System.out.print(":over3p_loan_amt="+over3p_loan_amt);//僅農綜貸(08)須報送
        System.out.print(":again_loan_cnt="+again_loan_cnt);//僅農綜貸(08)須報送
        System.out.print(":again_loan_bal_amt="+again_loan_bal_amt);//僅農綜貸(08)須報送
        System.out.println(":again_loan_bal_amt_1th="+again_loan_bal_amt_1th);//僅農綜貸(08)須報送
        
        paramList.add(M_YEAR);
        paramList.add(M_MONTH);
        paramList.add(BANK_NO);
        paramList.add(loan_item);
        paramList.add(loan_bal_cnt);
        paramList.add(loan_bal_amt);
        paramList.add(over6m_loan_bal_cnt);
        paramList.add(over6m_loan_bal_amt);
        paramList.add(delay_loan_cnt);
        paramList.add(delay_loan_amt);
        paramList.add(over3p_loan_cnt);
        paramList.add(over3p_loan_amt);
        paramList.add(again_loan_cnt);
        paramList.add(again_loan_bal_amt);          
        paramList.add(again_loan_bal_amt_1th);     
    }catch(Exception e){
        System.out.println("parseFRM Error:"+e.getMessage());
    }
    return paramList;
}


//103.04.02 add專案農貸貸款項目資料
private static List parseFRMITEM(String getfr){     
  String loan_item="",loan_item_name="";
  String sqlCmd = "";   
  List paramList =new ArrayList() ;      
  try{
      loan_item="";loan_item_name="";
      //System.out.println("len="+Utility.toBig5Convert(getfr).length());
      loan_item      = (Utility.toBig5Convert(getfr)).substring(1,3);//貸款項目代碼:文字（2-3）  
      loan_item_name = (Utility.toBig5Convert(getfr)).substring(3,63);//貸款項目名稱:文字（4-63）      
      loan_item_name = Utility.ISOtoBig5(loan_item_name);     
      //避免sql injection by 2295 ==================================================
      loan_item = loan_item.replaceAll("'","''");
      loan_item_name = loan_item_name.replaceAll("'","''").trim();  
      //====================================================================================      
      System.out.print(":loan_item="+loan_item); 
      System.out.println(":loan_item_name="+loan_item_name);     
     
      paramList.add(loan_item);
      paramList.add(loan_item_name);
    
  }catch(Exception e){
      System.out.println("parseFRMITEM Error:"+e.getMessage());
  }
  return paramList;
}


public static void send_Mail(String Subject,String messageText){      
  try {     
      String SMTP_HOST = Utility.getProperties("SMTP_Host").trim();
	  String FROM_ADDR = Utility.getProperties("From_Addr").trim();
	  String SEND_MAIL = Utility.getProperties("Send_Mail").trim();
	  String UserID	   = Utility.getProperties("UserID").trim();
	  String PWD	   = Utility.getProperties("PWD").trim();
	  String AgriBank_Addr1	= Utility.getProperties("AgriBank_Addr1").trim();
	  String AgriBank_Addr2	= Utility.getProperties("AgriBank_Addr2").trim();
	  String AgriBank_Addr3	= Utility.getProperties("AgriBank_Addr3").trim();
	  String Auth		= (Utility.getProperties("Auth")==null?"false":Utility.getProperties("Auth")).trim();
	  boolean auth = Boolean.getBoolean(Auth);
	  boolean sessionDebug = true;
	  if(SEND_MAIL.equals("true")){
	     if(Subject.equals("")){
	        Subject = "急件!!送至農金局的傳輸檔案失敗(至全國農業金庫轉檔失敗)....";
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
	     msg.setRecipients(Message.RecipientType.TO,InternetAddress.parse(AgriBank_Addr1+","+AgriBank_Addr2, false));
	     msg.setRecipients(Message.RecipientType.CC,InternetAddress.parse(AgriBank_Addr3, false));  	  		
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
    System.out.println("send_Mail Error:"+e.getMessage());
  }
}
}
