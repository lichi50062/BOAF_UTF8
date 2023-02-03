/*
 * Created on 2007/10/13
 * 96.02.27 fix link 至農金局網頁使用https://localhost by 2295
 * 98.09.01 add weburl使用設定檔MISWebSite by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoMC008W {
	public static void main(String[] args){
		System.out.println("AutoMC008W begin");
		AutoMC008W a = new AutoMC008W();		
		
		String lguser_id=args[0];
		String lguser_name=args[1];		
		String S_YEAR="";
		String S_MONTH="";
		if(args.length == 4){
		   if(args[2] != null) S_YEAR = args[2];
		   if(args[3] != null) S_MONTH = args[3];
		}
		a.exeMC008W(lguser_id,lguser_name,S_YEAR,S_MONTH);
		System.out.println("AutoMC008W end");
	}
	public void exeMC008W(String lguser_id, String lguser_name,String S_YEAR,String S_MONTH){
	    try{
		       //URL myURL=new URL("http://localhost:81/pages/AutoMC008W.jsp?lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
	    	   URL myURL=new URL(Utility.getProperties("MISWebSite")+"AutoMC008W.jsp?lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
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

