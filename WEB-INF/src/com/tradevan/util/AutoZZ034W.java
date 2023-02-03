/*
 * Created on 2009/09/24 申報資料跨表檢核排程作業 by 2205
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoZZ034W {
	public static void main(String[] args){
		System.out.println("AutoZZ034W begin");
		AutoZZ034W a = new AutoZZ034W();
			
		String S_YEAR="";
		String S_MONTH="";
		if(args.length == 2){
		   if(args[0] != null) S_YEAR = args[0];
		   if(args[1] != null) S_MONTH = args[1];
		}
		a.exeZZ034W(S_YEAR,S_MONTH);
		System.out.println("AutoZZ034W end");
	}
	public void exeZZ034W(String S_YEAR,String S_MONTH){
	    try{
		       URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"ZZ034W.jsp?act=checkALL&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
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

