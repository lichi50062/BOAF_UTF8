/*
 * Created on 2005/1/4-各機構基本資料
 *  98.09.01 add weburl使用設定檔BOAFWebSite by 2295
 * 100.07.21 fix 更改jsp名稱 by 2295
*/
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoHeadPerson_exe {
	public static void main(String[] args){
		System.out.println("AutoHeadPerson_exe begin");
		AutoHeadPerson_exe a = new AutoHeadPerson_exe();
		a.exeCheck();
		System.out.println("AutoHeadPerson_exe end");
	}
	public void exeCheck(){
	    try{
		       URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoHeadPerson.jsp");
		       //URL myURL=new URL("https://ebankweb.boaf.gov.tw/pages/test.jsp");//農金局正式環境
	    	   //URL myURL=new URL("http://localhost/pages/test.jsp");//測試環境
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

