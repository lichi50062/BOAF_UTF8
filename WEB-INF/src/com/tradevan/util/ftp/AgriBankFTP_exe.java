//95.09.12 create by 金庫轉檔2495
//96.01.05 fix 無法連結網站問題 by 2295
//96.02.27 fix link 至農金局網頁使用https://localhost by 2295
//98.09.01 add weburl使用設定檔BOAFWebSite by 2295
package com.tradevan.util.ftp;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

import com.tradevan.util.Utility;


public class  AgriBankFTP_exe{
public static void main(String[] args){
		System.out.println("AgriBankFTP_exe begin");
		AgriBankFTP_exe a = new AgriBankFTP_exe();
		a.exeCheck();
		System.out.println("AgriBankFTP_exe end");
	}
	public void exeCheck(){
	    try{
	    	   URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"AutoAgriBankFtp.jsp");  
	    	   //URL myURL=new URL("https://ebankweb.boaf.gov.tw/pages/AutoAgriBankFtp.jsp");//正式環境
	    	   //URL myURL=new URL("http://localhost:82/pages/AutoAgriBankFtp.jsp");//測試環境
	    	   	    	   
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