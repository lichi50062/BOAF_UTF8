/*
 * Created on 2015/08
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoFR072W {
	public static void main(String[] args){
		System.out.println("AutoFR072W begin");
		AutoFR072W a = new AutoFR072W();
		a.exeFR072W();
		System.out.println("AutoFR072W end");
	}
	public void exeFR072W(){	
	    
	    try{
		       //URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoFR072W.jsp");
	           URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoFR072W.jsp");
   	           //URL myURL=new URL("https://ebankweb.boaf.gov.tw/pages/WMAutoFR072W.jsp");
		       System.out.println("urlurlurl-----> "+myURL);
		       //URL myURL=new URL("https://ebankweb.boaf.gov.tw/pages/WMAutoFR072W.jsp");
		       //System.out.println(Utility.getProperties("BOAFWebSite")+"WMAutoFR072W.jsp");
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

