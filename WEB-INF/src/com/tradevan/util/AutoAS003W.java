/*
 * 106.11.28 create 配賦代號轉換作業.排程 by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoAS003W {
	public static void main(String[] args){
		System.out.println("AutoAS003W begin");
		AutoAS003W a = new AutoAS003W();
		a.exeAS003W();
		System.out.println("AutoAS003W end");
	}
	public void exeAS003W(){	
	    
	    try{
		       //URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoAS003W.jsp");
	           URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoAS003W.jsp");
		       System.out.println("urlurlurl-----> "+myURL);
		       //URL myURL=new URL("https://localhost/pages/WMAutoAS003W.jsp");
		       //System.out.println(Utility.getProperties("BOAFWebSite")+"WMAutoAS003W.jsp");
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

