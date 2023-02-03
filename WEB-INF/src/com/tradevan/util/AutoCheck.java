/*
 * Created on 2005/1/4
 * 96.02.27 fix link 至農金局網頁使用https://localhost by 2295
 * 98.09.01 申報資料排程檢核作業使用設定檔BOAFWebSite by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoCheck {
	public static void main(String[] args){
		System.out.println("AutoCheck begin");
		try{
		AutoCheck a = new AutoCheck();
		//System.out.println("https:/localhost/pages/WMAutoFileCheck.jsp?report_no="+args[0]);
		System.out.println(Utility.getProperties("BOAFWebSite")+"WMAutoFileCheck.jsp?report_no="+args[0]);
		a.exeCheck(args[0]);
		System.out.println("AutoCheck end");
		}catch(Exception e){System.out.println("AutoCheck Error"+e+e.getMessage());}
	}
	public void exeCheck(String report_no){		      	     
	    try{
		       //URL myURL=new URL("https:/localhost/pages/WMAutoFileCheck.jsp?report_no="+report_no);
		       URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoFileCheck.jsp?report_no="+report_no);
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

