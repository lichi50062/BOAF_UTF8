package com.tradevan.util;
import java.io.*;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.net.URL;
import java.net.URLConnection;
import org.apache.commons.net.ftp.*;

import com.tradevan.util.ftp.JakartaFtpWrapper;

/**
* @author User
*
* TODO To change the template for this generated type comment go to
* Window - Preferences - Java - Code Style - Code Templates
*/
public class AutoSm001W_Billing {
public static void main(String[] args){
	System.out.println("AutoSm001W_Billing begin");
	AutoSm001W_Billing a = new AutoSm001W_Billing();
	a.exeSm001W_Billing_URL(args[0],args[1]);
	//a.exeSm001W_Billing(args[0],args[1]);
	System.out.println("AutoSm001W_Billing end");
}

public void exeSm001W_Billing_URL(String UID,String PWD){
	 File logfile;
	 FileOutputStream logos=null;    	
	 BufferedOutputStream logbos = null;
	 PrintStream logps = null;
	 Date nowlog = new Date();
	 SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
	 SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    Calendar logcalendar;
    File dataFile = null;
	 try{
	       logfile = new File("/APSCGF/JAVA/sm001W.log");			   
		   logos = new FileOutputStream("/APSCGF/JAVA/sm001W.log",false);  		        	   
		   logbos = new BufferedOutputStream(logos);
		   logps = new PrintStream(logbos);			          
	       logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();			    	
		   logps.println(logformat.format(nowlog)+"轉檔開始");		    					    
		   logps.flush();
	       URL myURL=new URL("http://172.28.153.230:6009/APSCGF/Sm001W_Billing.jsp?UID="+UID+"&PWD="+PWD+"&flag=yes");
	       URLConnection raoURL;
	       raoURL=myURL.openConnection();
	       raoURL.setDoInput(true);
	       raoURL.setDoOutput(true);
	       raoURL.setUseCaches(false);
	   	   BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
    	   String raoInputString;	     	  
    	   while((raoInputString=infromURL.readLine())!=null){
      	          System.out.println(raoInputString);
    	   }
    	   infromURL.close();
    	   logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();			    	
		   logps.println(logformat.format(nowlog)+"轉檔結束");		    					    
		   logps.flush();
    	   
    	   dataFile = new File("/APSCGF/JAVA/SMBCS");
    	   String putMsg = "";
    	   if(dataFile.exists()){//有產生檔案
    	      nowlog = logcalendar.getTime();			    	
		      logps.println(logformat.format(nowlog)+"檔案產生成功");		    					    
		      logps.flush(); 
    	      putMsg= putFiles("172.20.5.22", "pbillftp", "ftppbill","/PBIL/ftp/APLOG","/APSCGF/JAVA/","/PBIL/ftp/APLOG","SMBCS");    	   	
              if(putMsg == null){//上傳檔案成功		                     
                 System.out.print("檔案上傳成功");
                 nowlog = logcalendar.getTime();			    	
    		     logps.println(logformat.format(nowlog)+"檔案上傳成功");		    					    
    		     logps.flush(); 	         					            		               
              }else{//end of 上傳檔案成功
                 nowlog = logcalendar.getTime();			    	
     		     logps.println(logformat.format(nowlog)+"檔案上傳失敗");		    					    
     		     logps.flush(); 	                	       	
              }
    	  }else{//end of 有產生檔案
    	  	 nowlog = logcalendar.getTime();			    	
		     logps.println(logformat.format(nowlog)+"檔案產生失敗");		    					    
		     logps.flush(); 
    	  } 
	 }catch (Exception e){
		    System.out.println(e.getMessage());
	 }
}
public String putFiles(String server_host, String username, String password,String remote_path, String local_path,String workDir,String filename) {
	try {
		  JakartaFtpWrapper ftp = new JakartaFtpWrapper();			
		  if(ftp.connectAndLogin(server_host, username, password)) {
			 System.out.println("Connected to " + server_host);
			try {
				System.out.println("Welcome message:\n" + ftp.getReplyString());
				System.out.println("Current Directory: " + ftp.printWorkingDirectory());
				//change remote path
				System.out.println("change home dir ["+remote_path+"] ?? "+ftp.changeWorkingDirectory(remote_path));
				System.out.println("begin change dir to "+remote_path+workDir);
	            //if (!ftp.changeWorkingDirectory(remote_path+workDir)) {
				if (!ftp.changeWorkingDirectory(workDir)) {
	                System.out.println("begin change dir failed!! ");
	                //create remote dir	                
	                //95.03.08 fix 防止sub dir 建不起來.先建parent dir 		
	                System.out.println("create parent dir["+workDir.substring(0,workDir.indexOf("/"))+"]="+ftp.makeDirectory(workDir.substring(0,workDir.indexOf("/"))));
	                
	                if(!ftp.makeDirectory(workDir)){
	                    System.out.println("ftp.makeDirectory("+workDir+") failed!!");
	                    ftp.logout();		                    
	                    return "Cannot create working directory to " + workDir;		                    
	                }else{
	                    //System.out.println("create dir success["+workDir+"]");
	                    System.out.println("create dir success["+remote_path+workDir+"]");
	                    if (!ftp.changeWorkingDirectory(remote_path+workDir)) {
	                        ftp.logout();
	                        return "Cannot change working directory to " + remote_path+workDir;
	                    }
	                }                
	            }
				System.out.println("Current Directory: " + ftp.printWorkingDirectory());
				ftp.setPassiveMode(true);
				System.out.println("setPassiveMode ok");
				//ftp.ascii();
				ftp.binary();//95.03.22 改用binary code
				System.out.println("binary");
				//starting put localfile
				FTPFile [] files = ftp.listFiles();
	           		            
		        System.out.println("put file begin==============================");			                
		        boolean  success = ftp.uploadFile (local_path+filename,filename);			                
		        System.out.print(":put " + filename + " success? " + success);
		        if(!success) return "put filename "+filename+" failed";
		        
			} catch (Exception ftpe) {
			    System.out.println("putFiles Error:"+ftpe.getMessage());
				ftpe.printStackTrace();
			} finally {
				//disconnect
	            if (ftp != null && ftp.isConnected()) {
	               try {
	               		ftp.logout();
	               		ftp.disconnect();		                    
	               } catch (IOException f) { }
	            }
				
			}
		} else {
			System.out.println("Unable to connect to" + server_host);
			return "Unable to connect to" + server_host;
		}
		System.out.println("Finished");
	} catch(Exception e) {
	    System.out.println(e);
	    System.out.println(e.getMessage());
		e.printStackTrace();
	}
	return null;
}


}
