//103.05.19 create SFTP上傳及下載檔案 by 2295
//104.04.02 fix 上傳檔案使用sftp並自動建目錄 by 2295
//110.06.29 add 設定上傳/下載使用絕對目錄 by 2295
package com.tradevan.util.sftp;

import org.apache.commons.vfs2.FileObject;
import org.apache.commons.vfs2.FileSystemOptions;
import org.apache.commons.vfs2.Selectors;
import org.apache.commons.vfs2.impl.StandardFileSystemManager;
import org.apache.commons.vfs2.provider.sftp.SftpFileSystemConfigBuilder;

import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException; 

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class MySFTPClient{
    private String server_host;
    private String username;
    private String password;
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Date nowlog = new Date();
	static Calendar logcalendar;
    public MySFTPClient(String server_host, String username, String password) {
        this.server_host = server_host;
        this.username = username;
        this.password = password;
    }
    public static void main(String args[]) {
        //MySFTPClient sftpC = new MySFTPClient("172.20.29.105", "pboafmgr", "tvpboafmgr");
        //MySFTPClient sftpC = new MySFTPClient("172.28.51.3", "root", "tvroot123");
    	MySFTPClient sftpC = new MySFTPClient("172.28.51.3", "pboafmgr", "tvpboafmgr");
    	//MySFTPClient sftpC = new MySFTPClient("172.22.64.30", "root", "tvroot123");
        
    	//sftpC.connect("172.22.64.30", 22, "pboafmgr","tvpboafmgr"); 
        //sftpC.getMyFiles("/APBOAF","C:\\","alter_table.sql");
    	sftpC.getMyFiles("/ORADATA/APBOAF/FR001WB_STOT/10607","C:\\BOAF\\GenerateRptDir\\10607","6T_10607.xls");
    	sftpC.sendMyFiles("/ORADATA/APBOAF/FR001WB_STOT/10607", "C:\\BOAF\\GenerateRptDir\\10607", "6T_10607.xls");  
        
        //System.out.println(ftpC.getFiles("/tmp", "c:/temp/ftp",""));
    }

    
    public boolean getMyFiles(String remote_path, String local_path,String fileToDownload){        
       
        StandardFileSystemManager manager = new StandardFileSystemManager();
        
        try { 
          
         //Initializes the file manager
         manager.init();
          
         //Setup our SFTP configuration
         //Create SFTP options
         FileSystemOptions opts = new FileSystemOptions();
         //SSH Key checking
         SftpFileSystemConfigBuilder.getInstance().setStrictHostKeyChecking(opts, "no");
         //Root directory set to user home
         //SftpFileSystemConfigBuilder.getInstance().setUserDirIsRoot(opts, true);//設定預設在登入者的目錄下
         
         //This line tells VFS to treat the URI as the absolute path and not relative
         SftpFileSystemConfigBuilder.getInstance( ).setUserDirIsRoot( opts, false );//設定使用絕對路徑
         //Timeout is count by Milliseconds
         SftpFileSystemConfigBuilder.getInstance().setTimeout(opts, 10000);
          
        
         
         //Create the SFTP URI using the host name, userid, password,  remote path and file name
         // sftpUri: "sftp://username:password@server_host/remote_path/filename
         String sftpUri = "sftp://" + username + ":" + password +  "@" + server_host + "/" +
                 remote_path + "/" + fileToDownload;
         
         //check local dir, if not exist then make local dir
         File local_dir = new File(local_path);
         if (!local_dir.exists()) {
             local_dir.mkdirs();
         } 
         // Create local file object
         String filepath = local_path+"\\" +  fileToDownload;
         File file = new File(filepath);
         FileObject localFile = manager.resolveFile(file.getAbsolutePath());
         System.out.println("localFile="+localFile);
         // Create remote file object
         
         FileObject remoteFile = manager.resolveFile(sftpUri, opts);
       
         // Create remote file object
         
         // Copy local file to sftp server
        
         System.out.println("remoteFile.getURL()="+remoteFile.getURL());
         localFile.copyFrom(remoteFile, Selectors.SELECT_SELF);
         System.out.println(" File download successful");
       
        }
        catch (Exception ex) {
         ex.printStackTrace();
         System.out.println(ex);
         System.out.println(ex.getMessage());         
         return false;
        }
        finally {
         manager.close();
        }
       
        return true;
       }
       
      
    
    public boolean sendMyFiles(String remote_path, String local_path, String fileToFTP){        
       
        StandardFileSystemManager manager = new StandardFileSystemManager();
       
        try {        
       
         //check if the file exists
         String filepath = local_path+"\\" + fileToFTP;
         File file = new File(filepath);
         System.out.println(file.getAbsolutePath()+":"+file.getName());
         if (!file.exists())  throw new RuntimeException("Error. Local file not found");
       
         //Initializes the file manager
         manager.init();
          
         //Setup our SFTP configuration
         //Create SFTP options
         FileSystemOptions opts = new FileSystemOptions();
         //SSH Key checking
         SftpFileSystemConfigBuilder.getInstance().setStrictHostKeyChecking(opts, "no");
         
         // This line tells VFS to treat the URI as the absolute path and not relative
         SftpFileSystemConfigBuilder.getInstance( ).setUserDirIsRoot( opts, false );//設定使用絕對路徑

         // Retrieve the file from the remote FTP server    
         //FileObject realFileObject = fileSystemManager.resolveFile( fileSystemUri, opts );
         //=====================================
         
         //Root directory set to user home
         //SftpFileSystemConfigBuilder.getInstance().setUserDirIsRoot(opts, true);//設定預設在登入者的目錄下
         
         //SftpFileSystemConfigBuilder.getInstance().setRootURI(opts, "sftp://" + username + ":" + password +  "@" + server_host + ":22/ORADATA/APBOAF/");
         
        
         //Timeout is count by Milliseconds
         SftpFileSystemConfigBuilder.getInstance().setTimeout(opts, 10000);
          
        // SftpFileSystemConfigBuilder.getInstance().setRootURI(opts, "/ORADATA/APBOAF/");
         
         //Create the SFTP URI using the host name, userid, password,  remote path and file name
         //sftpUri: "sftp://username:password@server_host/remote_path/filename
         String sftpUri = "sftp://" + username + ":" + password +  "@" + server_host + ":22/" +
                 remote_path + "/" + fileToFTP;
        
         System.out.println("sftpUri="+sftpUri);        
         // Create local file object
         FileObject localFile = manager.resolveFile(file.getAbsolutePath());
         System.out.println("localFile="+localFile.getURL());
         // Create remote file object
         FileObject remoteFile = manager.resolveFile(sftpUri, opts);
         
        
         //=====================================
         
         
         //System.out.println("remoteFile.getParent()="+remoteFile.getParent());
         //System.out.println("remoteFile.getFileSystem()="+remoteFile.getFileSystem());
         System.out.println("remoteFile.getURL()="+remoteFile.getURL());
         // Copy local file to sftp server
         remoteFile.copyFrom(localFile, Selectors.SELECT_SELF);
         System.out.println(fileToFTP+" File upload successful");
         
       
        }catch (Exception ex) {
            ex.printStackTrace();
            System.out.println(ex);
            System.out.println(ex.getMessage());         
            return false;
        }finally {
            manager.close();
        }
       
        return true;
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
    
    public boolean uploadFiles(String remote_path, List filename){     
        int port = 22;    
        String directory = "/APBOAF/FR001WB_STOT/";
        String subDir="";    
        ChannelSftp sftp=null;
        try{
             sftp= connect(server_host, port, username, password);
             boolean hasDir = false;
             System.out.println("remote_path="+remote_path);
             try{//確認該年月資料夾有無建立        
                 sftp.cd(remote_path);
                 System.out.println("cd "+remote_path + " ok");                 
                 hasDir = true;
             }catch(Exception e1){
                 e1.printStackTrace();        
             }
     
             if(!hasDir){//若無年月資料夾時,自動建立該資料夾
                 directory = remote_path.substring(0,remote_path.length()-5);
                 System.out.println("directory="+directory);        
                 subDir = remote_path.substring(remote_path.lastIndexOf("/")+1,remote_path.length());        
                 System.out.println("subDir="+subDir);
                 sftp.cd(directory);
                 System.out.println("cd "+directory + " ok");
                 sftp.mkdir(subDir);
                 sftp.cd(remote_path);
                 System.out.println("cd "+remote_path + " ok");        
             }
             //檔案上傳
             for(int i=0;i<filename.size();i++){
                 upload(remote_path, (String)filename.get(i), sftp);
                 System.out.println((String)filename.get(i)+" upload OK");
             }
             return true;
        }catch(Exception e){
            e.printStackTrace();
            return false;
        } finally {
            // disconnect
            if (sftp != null && sftp.isConnected()) {
               try {
                   sftp.disconnect();
                   sftp = null;
               } catch (Exception f) { }
           }
        }
    }
    
    
    public boolean getFiles(String remote_path, String local_path,String filename){     
        int port = 22;
        ChannelSftp sftp=null;              
        try{
             sftp= connect(server_host, port, username, password);
             System.out.println("remote_path="+remote_path);
             sftp.cd(remote_path);
             System.out.println("cd "+remote_path + " ok");     
             //檔案下載            
             download(remote_path, filename,local_path+filename, sftp);                 
             System.out.println(filename+" get OK");             
             return true;
        }catch(Exception e){
            e.printStackTrace();
            return false;
        } finally {
            // disconnect
            if (sftp != null && sftp.isConnected()) {
               try {
                   sftp.disconnect();
                   sftp = null;
               } catch (Exception f) { }
           }
        }    
    }
    
    /**
    * 連接sftp服務器
    * @param host 主機
    * @param port 端口
    * @param username 用戶名
    * @param password 密碼
    * @return
    */
    public ChannelSftp connect(String host, int port, String username,String password) {
        ChannelSftp sftp = null;
        try {
            JSch jsch = new JSch();
            
            java.util.Properties configuration = new java.util.Properties();
            
            //configuration.put("KexAlgorithms", "diffie-hellman-group1-sha1");
                       //configuration.put("KexAlgorithms", "diffie-hellman-group1-sha1,diffie-hellman-group14-sha1,diffie-hellman-group-exchange-sha1,diffie-hellman-group-exchange-sha256");
            //configuration.put("StrictHostKeyChecking", "no");
            
            Session sshSession = jsch.getSession(username, host, port);
            sshSession.setPassword("password");
            //configuration.put("kex","diffie-hellman-group1-sha1,diffie-hellman-group14-sha1,diffie-hellman-group-exchange-sha1,diffie-hellman-group-exchange-sha256");
            //configuration.put("kex","diffie-hellman-group1-sha1,curve25519-sha256@libssh.org,ecdh-sha2-nistp256,ecdh-sha2-nistp384,ecdh-sha2-nistp521,diffie-hellman-group-exchange-sha256,diffie-hellman-group14-sha1");
            //newSession.setConfig("kex", "diffie-hellman-group1-sha1,diffie-hellman-group14-sha1,diffie-hellman-group-exchange-sha1,diffie-hellman-group-exchange-sha256");
            //configuration.put("StrictHostKeyChecking","no");
            
            configuration.put("kex","diffie-hellman-group1-sha1,diffie-hellman-group14-sha1,diffie-hellman-group-exchange-sha1,diffie-hellman-group-exchange-sha256");
            configuration.put("StrictHostKeyChecking", "no");
            
            sshSession.setConfig(configuration);
            
            //sshSession.setConfig(configuration);
            sshSession.connect();
            
            /*
           jsch.getSession(username, host, port);
            Session sshSession = jsch.getSession(username, host, port);
            System.out.println("Session created.");
            sshSession.setPassword(password);
            Properties sshConfig = new Properties();
            sshConfig.put("StrictHostKeyChecking", "no");
            sshSession.setConfig(sshConfig);
            */
            sshSession.connect();
            System.out.println("Session connected.");
            System.out.println("Opening Channel.");
            Channel channel = sshSession.openChannel("sftp");
            channel.connect();
            sftp = (ChannelSftp) channel;
            System.out.println("Connected to " + host + ".");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return sftp;
    }

    /**
    * 上傳文件
    * @param directory 上傳的目錄
    * @param uploadFile 要上傳的文件
    * @param sftp
    */
    public void upload(String directory, String uploadFile, ChannelSftp sftp) {
        try {
            sftp.cd(directory);
            File file=new File(uploadFile);
            FileInputStream tmp = new FileInputStream(file);
            sftp.put(tmp, file.getName());
            tmp.close();
            /*
            if(file.exists()){
                System.out.println(uploadFile+" delete OK ?"+file.delete());
            }
            */
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
    * 下載文件
    * @param directory 下載目錄
    * @param downloadFile 下載的文件
    * @param saveFile 存在本地的路徑
    * @param sftp
    */
    public void download(String directory, String downloadFile,String saveFile, ChannelSftp sftp) {
        try {
            sftp.cd(directory);
            File file=new File(saveFile);
            sftp.get(downloadFile, new FileOutputStream(file));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
    * 刪除文件
    * @param directory 要刪除文件所在目錄
    * @param deleteFile 要刪除的文件
    * @param sftp
    */
    public void delete(String directory, String deleteFile, ChannelSftp sftp) {
        try {
            sftp.cd(directory);
            sftp.rm(deleteFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
    * 列出目錄下的文件
    * @param directory 要列出的目錄
    * @param sftp
    * @return
    * @throws SftpException
    */
    public Vector listFiles(String directory, ChannelSftp sftp) throws SftpException{
        return sftp.ls(directory);
    }
/*
    public static void main(String[] args) {
    MySFTP sf = new MySFTP(); 
    String host = "192.168.0.1";
    int port = 22;
    String username = "root";
    String password = "root";
    String directory = "/home/httpd/test/";
    String uploadFile = "D:\\tmp\\upload.txt";
    String downloadFile = "upload.txt";
    String saveFile = "D:\\tmp\\download.txt";
    String deleteFile = "delete.txt";
    ChannelSftp sftp=sf.connect(host, port, username, password);
    sf.upload(directory, uploadFile, sftp);
    sf.download(directory, downloadFile, saveFile, sftp);
    sf.delete(directory, deleteFile, sftp);
    try{
    sftp.cd(directory);
    sftp.mkdir("ss");
    System.out.println("finished");
    }catch(Exception e){
    e.printStackTrace();
    } 
    } 
    }
}
    */
    
    
    
}
