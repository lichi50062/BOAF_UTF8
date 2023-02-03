/*
 * Created on 2006/2/10
 * 96.02.27 fix link 至農金局網頁使用https://localhost by 2295
 * 98.09.01 add 排程鎖定作業weburl使用設定檔BOAFWebSite by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoZZ031W {
	public static void main(String[] args){
		System.out.println("AutoZZ031W begin");
		AutoZZ031W a = new AutoZZ031W();
		a.exeZZ031W(args[0],args[1],args[2],args[3]);
		System.out.println("AutoZZ031W end");
	}
	public void exeZZ031W(String report_no,String lguser_id,String bank_type,String lock_status){		      	     
	    try{
		       URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"WMAutoZZ031W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&bank_type="+bank_type+"&lock_status="+lock_status);
		       //URL myURL=new URL("https://localhost/pages/WMAutoZZ031W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&bank_type="+bank_type+"&lock_status="+lock_status);
		       //System.out.println(Utility.getProperties("BOAFWebSite")+"WMAutoZZ031W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&bank_type="+bank_type+"&lock_status="+lock_status);
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

