//95.03.08 fix 防止sub dir 建不起來.先建parent dir
//95.03.22 fix 改用binary code 方式 by 2295
//98.08.18 fix 檔名重覆時.不另存至LOG_Vxx目錄下,原檔案更名成file.年月日時分秒 by 2295
//98.09.10 add getFiles:get回來的檔名.由ISO8859_1轉->Big5 by 2295
package com.tradevan.util.ftp;
import org.apache.commons.net.ftp.*;

import com.tradevan.util.Utility;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class MyFTPClient{
    private String server_host;
    private String username;
    private String password;
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Date nowlog = new Date();
	static Calendar logcalendar;
    public MyFTPClient(String server_host, String username, String password) {
        this.server_host = server_host;
        this.username = username;
        this.password = password;
    }

    public String getFiles(String remote_path, String local_path,List filename) {
        FTPClient ftp = null;
        try {

            ftp = new FTPClient();

            ftp.connect(server_host);

            int reply = ftp.getReplyCode();

            if (!FTPReply.isPositiveCompletion(reply)) {
                ftp.disconnect();
                return "FTP server refused connection";
            }

            if (!ftp.login(username, password)) {
                ftp.logout();
                return "Login failed!";
            }else{
                System.out.println("Login success");
            }

            // set transport mode to binary
            ftp.setFileType(FTP.BINARY_FILE_TYPE);
            System.out.println("setFileType ok");
            ftp.enterLocalPassiveMode();
            System.out.println("enterLocalPassiveMode ok");
            // change remote path
            if (!ftp.changeWorkingDirectory(remote_path)) {
                System.out.println("Cannot change working directory to " + remote_path);
                ftp.logout();
                return "Cannot change working directory to " + remote_path;
            }
            
            // check local dir, if not exist then make local dir
            File local_dir = new File(local_path);
            if (!local_dir.exists()) {
                local_dir.mkdirs();
            }
            // starting get remote file
            System.out.println("local_dir="+local_path);
            FTPFile [] files = ftp.listFiles();
            boolean success = false;
            boolean findFile=false;
            FileOutputStream fos = null;
            int iRptCount=0;//儲存完成的報表數
            
            String tmp_files="";
            System.out.println("filename.size()="+filename.size());
            fileLoop:
            for(int j=0;j<filename.size();j++){
                findFile=false;
                for (int i=0;i<files.length;i++) {
                     if (files[i].isFile()) {
                     	  //98.09.10 add 中文檔名.get回來的檔名.由ISO8859_1轉->Big5
                     	  tmp_files = Utility.ISOtoBig5(files[i].getName());
                     	  
                          if(files[i].getName().equals((String)filename.get(j))){
                          System.out.println("begin get file ="+files[i].getName());
                          findFile=true;
                          fos = new FileOutputStream(local_path + "/" + tmp_files);
                          success = ftp.retrieveFile(files[i].getName(), fos);
                          fos.close();                           
                          if(!success){
                              System.out.println("get " + files[i].getName() + " success? " + success);
                              return "get " + files[i].getName() + " failed ";
                          }else{
                              System.out.println("get " + files[i].getName() + " success? " + success);
                          }
                          iRptCount++;
                          if(filename.size() == 1) break fileLoop;
                          if(iRptCount == filename.size()) break fileLoop;                             
                        }
                     }
                }//end of server file
                if(!findFile) return "can not find file "+(String)filename.get(j) + " in "+remote_path;
            }//end of filename            
            ftp.logout();     
            
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(e);
            System.out.println(e.getMessage());
            return e.toString();
            
        } finally {
            // disconnect
            if (ftp != null && ftp.isConnected()) {
               try {
                   ftp.disconnect();
                   ftp = null;
               } catch (IOException f) { }
           }

        }
        return null;
    }    
    public String putFiles(String remote_path, String local_path,String workDir,List filename) {
		try {
			JakartaFtpWrapper ftp = new JakartaFtpWrapper();			
			if (ftp.connectAndLogin(server_host, username, password)) {
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
		                //if(!ftp.makeDirectory(workDir)){
		                   // System.out.println("ftp.makeDirectory("+workDir+") failed!!");
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
		            boolean findFile=false;
					boolean success = false; 
					String[] idx = new String[3];
					String strPrefix = "LOG_V";
			        for(int i=0;i<filename.size();i++){
			            //find file in workdir=========================================
			            if(findFile(files,(String)filename.get(i))){//找到該檔案
			               System.out.println((String)filename.get(i)+" has in "+remote_path+workDir);
			               logcalendar = Calendar.getInstance(); 
			 			   nowlog = logcalendar.getTime();
			 			   //98.08.18重覆產生時,原檔案更名成file.年月日時分秒
			               System.out.println((String)filename.get(i)+" rename to "+(String)filename.get(i)+"."+logformat.format(nowlog)+" ??"+ftp.rename((String)filename.get(i),(String)filename.get(i)+"."+logformat.format(nowlog)));
			               /*98.08.18 fix 重覆產生時.不另存至LOG_Vxx目錄下
			               LOG_VLoop:
			               for(int k=1;k<=100;){
			                   strPrefix = "LOG_V";
			                   if(String.valueOf(k).length() == 1){			                       
			                      strPrefix += "00";			                      
			                   }else if(String.valueOf(k).length() == 2){
			                      strPrefix += "0";
			                   }		                   
			                   //change remote path
			       		       if (!ftp.changeWorkingDirectory(remote_path+workDir+strPrefix +String.valueOf(k))) {
			       		            System.out.println(remote_path+workDir+strPrefix +String.valueOf(k)+" has not exists");
			       		            System.out.println("workDir="+workDir+strPrefix +String.valueOf(k));
			       		            ftp.changeWorkingDirectory(remote_path);
			       		            System.out.println("Current Directory: " + ftp.printWorkingDirectory());
			       		            //create remote dir
			       		            if(!ftp.makeDirectory(workDir+strPrefix +String.valueOf(k))){
			       		                ftp.logout();
			       		                System.out.println("Cannot create working directory to " + workDir+strPrefix +String.valueOf(k));
			       		                return "Cannot create working directory to " + workDir+strPrefix +String.valueOf(k); 
			       		            }else{
			       		                System.out.println("create dir success["+workDir+strPrefix +String.valueOf(k)+"]");
   			       		                if (!ftp.changeWorkingDirectory(remote_path+workDir+strPrefix +String.valueOf(k))) {
			       		                    ftp.logout();
			       		                    return "Cannot change working directory to " + remote_path+workDir+strPrefix +String.valueOf(k);
			       		                }
			       		            }                
			       		       }//end of change remote path
			       			   System.out.println("Current Directory: " + ftp.printWorkingDirectory());
			       			   ftp.setPassiveMode(true);
			       			   ftp.ascii();
			       			   files = ftp.listFiles();
			       			   if(findFile(files,(String)filename.get(i))){//找到該檔案			       			       
			       			      System.out.println((String)filename.get(i)+" has in "+remote_path+workDir+strPrefix +String.valueOf(k));
			       				  k++; 
			       			   }else{
			       			      break LOG_VLoop;
			       			   }
			               }//end of k=1~100
			               */			                
			            }
			            //else{//find file
			               ftp.changeWorkingDirectory(remote_path+workDir); 
			            //}
			            //============================================================			            
			            System.out.println("put file begin==============================");			                
			            success = ftp.uploadFile (local_path+(String)filename.get(i),(String)filename.get(i));			                
			            System.out.print(":put " + (String)filename.get(i) + " success? " + success);
			            if(!success) return "put filename "+(String)filename.get(i)+" failed";
			        }//end of filename
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
    
    private boolean findFile(FTPFile [] files,String filename){
        boolean finded = false;
        for (int j=0;j<files.length;j++) {
            if (files[j].isFile()) {			                      
               if(files[j].getName().equals(filename)){
                   finded = true;
               }
            }
        }    
        return finded;
    }
    
    public boolean CreateZIP(String local_path,List filenames){
        // These are the files to include in the ZIP file
        //String[] filenames = new String[]{"filename1", "filename2"};
        boolean bSuccess = false;
        //Create a buffer for reading the files
        byte[] buf = new byte[1024];
        System.out.println("CreateZIP local_path="+local_path);
        try {
            // Create the ZIP file
            String outFilename = local_path+System.getProperty("file.separator")+"outfile.zip";
            ZipOutputStream out = new ZipOutputStream(new FileOutputStream(outFilename));    
            //Compress the files
            for (int i=0; i<filenames.size(); i++) {
                FileInputStream in = new FileInputStream((String)filenames.get(i));
                System.out.println("add file "+(String)filenames.get(i));
                // Add ZIP entry to output stream.
                out.putNextEntry(new ZipEntry((String)filenames.get(i)));    
                // Transfer bytes from the file to the ZIP file
                int len;
                while ((len = in.read(buf)) > 0) {
                    out.write(buf, 0, len);
                }    
                // Complete the entry
                out.closeEntry();
                in.close();
            }
            
            // Complete the ZIP file
            out.close();
            bSuccess = true;
        } catch (IOException e) {            
            System.out.println("CreateZIP Error:"+e.getMessage());
        } 
        return bSuccess;
    }	
      


    
    
    public static void main(String args[]) {
        MyFTPClient ftpC = new MyFTPClient("server", "user", "password");
        //System.out.println(ftpC.getFiles("/tmp", "c:/temp/ftp",""));
    }

}
