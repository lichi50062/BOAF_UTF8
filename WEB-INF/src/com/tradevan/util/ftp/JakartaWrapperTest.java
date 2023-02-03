package com.tradevan.util.ftp;

/* 
 */
//import JakartaFtpWrapper;
import java.io.*;

/**
  * a very simple example of using the JakartaFtpWrapper class,
  * available at http://www.nsftools.com/tips/JavaFtp.htm
  */

public class JakartaWrapperTest {
	public static void main (String[] args) {
		try {
			new JakartaWrapperTest().test("","");
			
			//FTPSClient
			/*
			JakartaFtpSWrapper ftp = new JakartaFtpSWrapper();
			String serverName = "172.16.2.91";
			if (ftp.connectAndLogin(serverName, "boafftp", "boaf+1080605!")) {
				System.out.println("Connected to " + serverName);
				try {
					System.out.println("Welcome message:\n" + ftp.getReplyString());
					System.out.println("Current Directory: " + ftp.printWorkingDirectory());
					ftp.setPassiveMode(true);
					System.out.println("Files in this directory:\n" + ftp.listFileNamesString());
					System.out.println("Subdirectories in this directory:\n" + ftp.listSubdirNamesString());
					System.out.println("Downloading file robots.txt");
					ftp.ascii();
					ftp.downloadFile("b.sql", "C:\\b.sql");
				} catch (Exception ftpe) {
					ftpe.printStackTrace();
				} finally {
					ftp.logout();
					ftp.disconnect();
				}
			} else {
				System.out.println("Unable to connect to" + serverName);
			}
			System.out.println("Finished");
			*/
			/*FTPClient
			 JakartaFtpWrapper ftp = new JakartaFtpWrapper();
			String serverName = "172.20.5.22";
			if (ftp.connectAndLogin(serverName, "pboafmgr", "tvpboafmgr")) {
				System.out.println("Connected to " + serverName);
				try {
					System.out.println("Welcome message:\n" + ftp.getReplyString());
					System.out.println("Current Directory: " + ftp.printWorkingDirectory());
					ftp.setPassiveMode(true);
					System.out.println("Files in this directory:\n" + ftp.listFileNamesString());
					System.out.println("Subdirectories in this directory:\n" + ftp.listSubdirNamesString());
					System.out.println("Downloading file robots.txt");
					ftp.ascii();
					ftp.downloadFile("b.sql", "C:\\b.sql");
				} catch (Exception ftpe) {
					ftpe.printStackTrace();
				} finally {
					ftp.logout();
					ftp.disconnect();
				}
			} else {
				System.out.println("Unable to connect to" + serverName);
			}
			System.out.println("Finished");
			 */
			
			
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	public void test(String remotepath,String local_path){
		try {
			JakartaFtpWrapper ftp = new JakartaFtpWrapper();
			String serverName = "172.16.2.91";
			if (ftp.connectAndLogin(serverName, "boafftp", "boaf+1080605!")) {
				System.out.println("Connected to " + serverName);
				try {
					System.out.println("Welcome message:\n" + ftp.getReplyString());					
					if(!ftp.changeWorkingDir(remotepath)){
						System.out.println("Cannot change working directory to "+remotepath);
					}else{
						System.out.println("Current Directory: " + ftp.printWorkingDirectory());
						ftp.setPassiveMode(true);
						System.out.println("Files in this directory:\n" + ftp.listFileNamesString());
						System.out.println("Subdirectories in this directory:\n" + ftp.listSubdirNamesString());
						System.out.println("Downloading file b.sql");
						ftp.ascii();
						//check local dir, if not exist then make local dir
			            File local_dir = new File(local_path);
			            if (!local_dir.exists()) {
			                local_dir.mkdirs();
			            }
					   ftp.downloadFile("c1.sql", "C:\\c111.sql");
					   ftp.uploadFile ("C:\\test55.java", "test11.java");
					}
				} catch (Exception ftpe) {
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
				System.out.println("Unable to connect to" + serverName);
			}
			System.out.println("Finished");
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}


