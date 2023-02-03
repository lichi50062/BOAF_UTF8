//97.10.02 final version by 2295
//98.08.03 add 單一connect.disconnect by 2295
package UploadHtml;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class UploadHtml
{
	 File logfile;
	 FileOutputStream logos;
	 BufferedOutputStream logbos;
	 PrintStream logps;
	 Date nowlog;
	 SimpleDateFormat logformat;
	 SimpleDateFormat dirformat;
	 SimpleDateFormat logfileformat;
	 Calendar logcalendar;
	 //private static String dirpath = "/APSCGF/JAVA/htmlDir.properties";//正式
	 private static String dirpath = "D:\\workProject\\BOAF\\WEB-INF\\classes\\UploadHtml\\htmlDir.properties";//測試
	 
	 private static String server_host="172.20.24.17";
	 
	 //正式帳號
	 /*
	 private static String username="ftpPSCGF";
	 private static String password="*8uw6JmzE";
	 */ 
     //測試帳號
	 private static String username="smbmgr";
	 private static String password="tvsmbmgr"; 
	 JakartaFtpWrapper ftp = new JakartaFtpWrapper();
    public UploadHtml()
    {
        logos = null;
        logbos = null;
        logps = null;
        nowlog = new Date();
        logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
        dirformat = new SimpleDateFormat("yyyy/MM/dd");
        logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    }

    public static void main(String args1[])
    {
    }

    public String exeUploadHtml(String zipFileName)
    {
        String returnMsg;
        returnMsg = "";
        String fname1[];
        String putMsg = null;
        File uploadFile = null;
        File tmpFile = null;
        File subfile = null;
        List filename_List = new LinkedList();
        String nowDir = "";
        String zipfiledate = "";
        try
        {
            //logfile = new File("/APSCGF/JAVA/uploadHtml.log");//正式
            //logos = new FileOutputStream("/APSCGF/JAVA/uploadHtml.log", true);//正式
        	logfile = new File("C:\\test2\\uploadHtml.log");//測試
            logos = new FileOutputStream("C:\\test2\\uploadHtml.log", true);//測試
            logbos = new BufferedOutputStream(logos);
            logps = new PrintStream(logbos);
            logcalendar = Calendar.getInstance();
            nowlog = logcalendar.getTime();
            uploadFile = new File(getProperties("homeDir"));
            //讀取zip檔名.當做上傳目錄名稱
            if(zipFileName.indexOf(".") != -1)
                zipfiledate = zipFileName.substring(0, zipFileName.indexOf("."));
            else
                zipfiledate = getCHTdate(dirformat.format(nowlog), 0);
            
            if(ftp.connectAndLogin(server_host, username, password)) {//connect成功時
          	   System.out.println("Connected to " + server_host);	
          	   //logcalendar = Calendar.getInstance();
                //nowlog = logcalendar.getTime();
          	   pringLog("Connected to " + server_host);
          	   pringLog("Welcome message:\n" + ftp.getReplyString());
          	   pringLog("Current Directory: " + ftp.printWorkingDirectory());          	   
          	   //System.out.println("Welcome message:\n" + ftp.getReplyString());
  		 	   //System.out.println("Current Directory: " + ftp.printWorkingDirectory());    
            }else{
            	return "Connected to " + server_host + " Failed !!";
            }
            if(uploadFile.exists() && uploadFile.isDirectory()){
                fname1 = uploadFile.list();//====列出此目錄下的所有檔案===================
                //上傳uploadDir根目錄下的檔案
                for(int c = 0; c < fname1.length; c++){
                    tmpFile = new File(getProperties("homeDir") + System.getProperty("file.separator") + fname1[c]);
                    if(!tmpFile.isDirectory())
                        filename_List.add(fname1[c]);
                }

                if(filename_List.size() != 0){
                	System.out.println("filename_List.size()="+filename_List.size());
                    //putMsg = putFiles("172.20.24.17", "ftpPSCGF", "*8uw6JmzE", "/export/home/smbmgr/" + zipfiledate, getProperties("homeDir") + System.getProperty("file.separator"), "/export/home/smbmgr/" + zipfiledate, filename_List, "0");//正式
                	putMsg = putFiles("172.20.24.17", "smbmgr", "tvsmbmgr", "/export/home/smbmgr/2295test/" + zipfiledate, getProperties("homeDir") + System.getProperty("file.separator"), "/export/home/smbmgr/2295test/" + zipfiledate, filename_List, "0");//測試
                    filename_List = new LinkedList();
                    if(putMsg != null) pringLog(putMsg);
                    
                    pringLog(" thread sleep begin");
                    Thread.currentThread().sleep(1000);
                    pringLog(" thread sleep end");
                }
            }
            for(int i = 1; i <= Integer.parseInt(getProperties("dir_count")); i++){
                nowDir = getProperties("dir" + i);
                if(!nowDir.equals("")){
                    uploadFile = new File(getProperties("homeDir") + System.getProperty("file.separator") + nowDir);
                    System.out.print("uploadFile=" + getProperties("homeDir") + System.getProperty("file.separator") + nowDir);
                    if(uploadFile.exists() && uploadFile.isDirectory()){
                        fname1 = uploadFile.list();
                        System.out.println(":fname1.length=" + fname1.length);
                        for(int c = 0; c < fname1.length; c++){
                            tmpFile = new File(getProperties("homeDir") + System.getProperty("file.separator") + nowDir + System.getProperty("file.separator")+ fname1[c]);
                            if(!tmpFile.isDirectory()){
                            	System.out.println("add filename="+fname1[c]);
                                filename_List.add(fname1[c]);
                            }    
                        }

                        if(filename_List.size() != 0){
                        	System.out.println("filename_List.size()="+filename_List.size());
                            putMsg = null;
                            //putMsg = putFiles("172.20.24.17", "ftpPSCGF", "*8uw6JmzE", "/export/home/smbmgr/" + zipfiledate, getProperties("homeDir") + System.getProperty("file.separator") + nowDir + System.getProperty("file.separator"), nowDir, filename_List, "1");//正式
                            putMsg = putFiles("172.20.24.17", "smbmgr", "tvsmbmgr", "/export/home/smbmgr/2295test/" + zipfiledate, getProperties("homeDir") + System.getProperty("file.separator") + nowDir + System.getProperty("file.separator"), nowDir, filename_List, "1");//測試
                            if(putMsg != null) pringLog(putMsg);
                            
                            pringLog(" thread sleep begin");
                            Thread.currentThread().sleep(1000);
                            pringLog(" thread sleep end");
                        }
                    } else{
                        System.out.println(getProperties("homeDir") + System.getProperty("file.separator") + nowDir + " has not exists");
                    }
                } else{
                    System.out.println("dir" + i + " has not found");
                    pringLog("dir" + i + " has not found");
                }
                filename_List = new LinkedList();
            }
            pringLog("============file upload complete===================");
            
            returnMsg = "檔案上傳成功";
            File updateSMBCGFFile = new File(getProperties("homeDir") + System.getProperty("file.separator") + "updateSMBCGF.flg");
            if(updateSMBCGFFile.exists()) updateSMBCGFFile.delete();
        }catch(Exception e){
            System.out.println("exeUploadHtml Error:" + e + e.getMessage());
            returnMsg = "exeUploadHtml Error:" + e + e.getMessage();
        }finally{
            try{
                if(logos != null) logos.close();
                if(logbos != null) logbos.close();
                if(logps != null)  logps.close();
                //disconnect
                if (ftp != null && ftp.isConnected()) {
                   try {
                   		ftp.logout();
                   		ftp.disconnect();		                    
                   } catch (IOException f) { }
                   pringLog(" disconnect " + server_host);
                }
            }catch(Exception ioe){
                System.out.println(ioe.getMessage());
            }
        }
        return returnMsg;
    }

    public String putFiles(String server_host, String username, String password, 
    					   String remote_path, String local_path, String workDir, List filename, 
						   String parentDir){/* parentDir=0:根目錄下的檔案.parentDir=1:根目錄下所有子目錄的檔案*/    
    try{    
    	
        //logfile = new File("/APSCGF/JAVA/uploadHtml.log");//正式
        //logos = new FileOutputStream("/APSCGF/JAVA/uploadHtml.log", true);//正式
    	/*980803
        logfile = new File("C:\\test2\\uploadHtml.log");//測試
        logos = new FileOutputStream("C:\\test2\\uploadHtml.log", true);//測試
        logbos = new BufferedOutputStream(logos);
        logps = new PrintStream(logbos);
        */
        
        String nowDir = "";
        pringLog("filename=" + filename);
        
        
        /*980803
		if(ftp.connectAndLogin(server_host, username, password)) {//connect成功時
     	   System.out.println("Connected to " + server_host);	
     	   //logcalendar = Calendar.getInstance();
           //nowlog = logcalendar.getTime();
     	   pringLog("Connected to " + server_host);
     	   //logps.println(logformat.format(nowlog)+"Connected to " + server_host);		    					    
     	   //logps.flush();     	    
     	    */
     	   try {
     		 	System.out.println("Welcome message:\n" + ftp.getReplyString());
     		 	System.out.println("Current Directory: " + ftp.printWorkingDirectory());     			 	
     			System.out.println("parentDir=" + parentDir);
     			//97.10.02要先切換到/export/home/smbmgr/目錄下.才可切換到其他下層目錄
     			System.out.println("change home dir [/export/home/smbmgr/] ?? " + ftp.changeWorkingDirectory("/export/home/smbmgr/"));
     			System.out.println("change home dir [" + remote_path + "] ?? " + ftp.changeWorkingDirectory(remote_path));
     			System.out.println("begin change dir to " + workDir);
     			//change remote path
     			if(!ftp.changeWorkingDirectory(workDir)){                
     			 	System.out.println("begin change dir failed!! ");
     			 	pringLog("begin change dir failed!! " + workDir);
     			}
     			pringLog(":begin change dir to " + workDir);
     			pringLog(":change home dir [" + remote_path + "] ?? " + ftp.changeWorkingDirectory(remote_path));
     			
                if(parentDir.equals("0")){//為上傳uploadDir下的檔案.需建立uploadDir下的子目錄                     
                    if(!ftp.makeDirectory(workDir)){                     
                    	System.out.println("ftp.makeDirectory(" + workDir + ") failed!!");
                    	//98.08.03ftp.logout();
                        return("Cannot create working directory to " + workDir);
                    } 	
                }
                System.out.println("create dir success[" + workDir + "]");
                pringLog(":create dir success[" + workDir + "]");
                
                if(!ftp.changeWorkingDirectory(workDir)){                 
                	//98.08.03ftp.logout();
                	return("Cannot change working directory to " + workDir);
                }
                if(parentDir.equals("0")){//為根目錄.先將子目錄建起來        
                	System.out.println("Current Directory: " + ftp.printWorkingDirectory());
                	System.out.println("begin create sub dir ");
                	for(int i = 1; i <= Integer.parseInt(getProperties("dir_count")); i++){
                		nowDir = getProperties("dir" + i);
                		System.out.println("dir_count"+i+"[nowDir]=" + nowDir);
                		System.out.println("Current Directory: " + ftp.printWorkingDirectory());
                		if(!nowDir.equals("")){
                			if(!ftp.makeDirectory(remote_path + "/" + nowDir)){
                				pringLog(":create sub dir fail[" + nowDir + "]");
                				System.out.println("create sub dir fail[" + nowDir + "]");
                			} else{
                				pringLog(":create sub dir success[" + nowDir + "]");
                				System.out.println("create sub dir success[" + nowDir + "]");
                			}
                		}    
                	}
                	System.out.println("end create sub dir ");
                } else {
                	System.out.println("don't need to create sub dir");
                }
                pringLog("Current Directory: " + ftp.printWorkingDirectory());
                
                System.out.println("Current Directory: " + ftp.printWorkingDirectory());
                ftp.setPassiveMode(true);
                System.out.println("setPassiveMode ok");
                ftp.binary();
                System.out.println("binary");
                boolean success = false;          
                for(int i=0;i<filename.size();i++){//有檔案需上傳時
                	System.out.println("put file begin==============================");
                	success = ftp.uploadFile(local_path + (String)filename.get(i), (String)filename.get(i));
	                System.out.println(":put " + (String)filename.get(i) + " success? " + success);
	                pringLog(":put " + local_path + (String)filename.get(i) + " to " + (parentDir.equals("0") ? "" : remote_path + System.getProperty("file.separator")) + workDir + System.getProperty("file.separator") + (String)filename.get(i) + " success? " + success);
		            
		            //if(!success) return("put filename " + filename + " failed");//97.10.02拿掉.即使有某個檔案上傳失敗.其他的繼續上傳
                }//end of filename
             } catch (Exception ftpe) {
     		    System.out.println("putFiles Error:"+ftpe.getMessage());
     		    pringLog(ftpe.getMessage());
     			ftpe.printStackTrace();
     			return "upload Fail";
     		}/* finally {
     			//disconnect
                 if (ftp != null && ftp.isConnected()) {
                    try {
                    		ftp.logout();
                    		ftp.disconnect();		                    
                    } catch (IOException f) { }
                 }	
                 pringLog(" disconnect " + server_host);
     		 }//end of finally
     		 */
	      /*
          } else {//無法connect至Server時	     	
			System.out.println("Unable to connect to " + server_host);	
			pringLog("Unable to connect to " + server_host);
			return "Unable to connect to " + server_host;
		 }
		 */
		 System.out.println("Finished");
     } catch(Exception e) {
	     System.out.println(e);
	     System.out.println(e.getMessage());
		 e.printStackTrace();
	 }
	return null;
	}
    
    
    
    public static String getProperties(String key) throws Exception{
        String value = "";
        String homeDir = "";
        try{
            Properties p = new Properties();
            p.load(new FileInputStream(dirpath));            
            homeDir = (String)p.get("homeDir");           
            p = new Properties();
            p.load(new FileInputStream(homeDir + System.getProperty("file.separator") + "htmlDir.properties"));
            value = (String)p.get(key);
        }catch(FileNotFoundException e){
            throw new Exception("UploadHtml.getProperties:FileNotFoundException[" + e.getMessage() + "]");
        }catch(IOException e){
            throw new Exception("UploadHtml.getProperties:IOException[" + e.getMessage() + "]");
        }
        return value;
    }

    public static String getCHTdate(String s, int i){
        String CHTdate = null;
        if(s == null || s.equals(""))
            return "";
        if(s.trim().length() != 10)
            return s.trim();
        int yr = Integer.parseInt(s.substring(0, 4)) - 1911;
        if(i == 0)
            CHTdate = Integer.toString(yr) + s.substring(5, 7) + s.substring(8, 10);
        if(i == 1)
            CHTdate = Integer.toString(yr) + "年" + s.substring(5, 7) + "月" + s.substring(8) + "日";
        return CHTdate;
    }
    private void pringLog(String data){
    	 logcalendar = Calendar.getInstance();
         nowlog = logcalendar.getTime();
   	     logps.println(logformat.format(nowlog)+data);		    					    
   	     logps.flush();
    }
}