/*
  109.09.09 農委會open data報表檔案上傳排程 by 2295
*/
package com.tradevan.util.ftp;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.Utility;
import com.tradevan.util.sftp.*;


public class CoaFtp {	
	static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	static SimpleDateFormat filenameformat = new SimpleDateFormat("yyyyMMdd");
	static Date nowlog = new Date();
	static Calendar logcalendar;
	static File xlsDir= null;
	static File coaxlsDir,coaBKxlsDir = null;
	
	public static void main(String args[]) { 
	    	 
	} 
	 
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
    
    
    public static void printRptMsg(PrintStream logps,String errRptMsg){
    	if(!errRptMsg.equals("")){
	       logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();
	       logps.println(logformat.format(nowlog)+errRptMsg);
	       logps.flush();
	    }
    }
    
}
