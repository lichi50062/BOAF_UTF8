/*
 * Created on 2012/07
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoSumBn01 {
	public static void main(String[] args){
		System.out.println("AutoSumBn01 begin");
		AutoSumBn01 a = new AutoSumBn01();
		a.exeSumBn01();
		System.out.println("AutoSumBn01 end");
	}
	public void exeSumBn01(){	
	    
	    try{
		       
	           URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoSumBn01.jsp");
		       System.out.println("urlurlurl-----> "+myURL);
		       //URL myURL=new URL("https://localhost/pages/WMAutoSumBn01.jsp");
		       //System.out.println(Utility.getProperties("BOAFWebSite")+"WMAutoSumBn01.jsp");
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

