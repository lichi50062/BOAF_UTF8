/* 
  103.05.27 add 上傳A01/A02/A03/A04/A05/WLX01/WLX01_M/WLX02/WLX02_M至金庫  by 2295
*/
package com.tradevan.util.ftp;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import com.tradevan.util.Utility;


public class  AutoAgriBankWM{
public static void main(String[] args){
		System.out.println("AutoAgriBankWM begin");
		AutoAgriBankWM a = new AutoAgriBankWM();
		a.exeCheck();
		System.out.println("AutoAgriBankWM end");
	}
	public void exeCheck(){
	    try{
	    	   URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"AutoAgriBankWM.jsp");
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