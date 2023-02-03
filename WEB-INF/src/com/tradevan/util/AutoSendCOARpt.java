/* 
 * 109.09.08 農委會OpenData報表檔案上傳至SFTP Server排程 by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import com.tradevan.util.Utility;


public class AutoSendCOARpt {
	public static void main(String[] args){
		System.out.println("AutoSendCOARpt begin");
		AutoSendCOARpt a = new AutoSendCOARpt();
		a.exeRpt();
		System.out.println("AutoSendCOARpt end");
	}
	public void exeRpt(){
	    try{
		       
		       URL myURL=new URL(Utility.getProperties("MISWebSite")+"WMAutoSendCOARpt.jsp");
		       URLConnection raoURL;
		       raoURL=myURL.openConnection();
		       raoURL.setDoInput(true);
		       raoURL.setDoOutput(true);
		       raoURL.setUseCaches(false);
		   	   BufferedReader infromURL=new BufferedReader(
	               	   new InputStreamReader(raoURL.getInputStream()));
	     	    String raoInputString;	     	  
	     	    while((raoInputString=infromURL.readLine())!=null){
	       	          System.out.println(raoInputString);
	     	    }
	     	    infromURL.close();
     	}catch (Exception e){
			    System.out.println(e.getMessage());
		}
    }		
}

