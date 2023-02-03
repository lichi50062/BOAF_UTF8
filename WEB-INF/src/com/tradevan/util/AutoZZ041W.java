/*
 * Created on 2006/2/10
 * 96.02.27 fix link 至農金局網頁使用https://localhost by 2295
 * 98.09.01 管理報表產生及發佈(ZZ041W)使用設定檔MISWebSite by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import com.tradevan.util.Utility;

public class AutoZZ041W {
	public static void main(String[] args){
		System.out.println("AutoZZ041W begin");
		AutoZZ041W a = new AutoZZ041W();
		
		String report_no=args[0];
		String lguser_id=args[1];
		String lguser_name=args[2];
		String unit=args[3];
		String S_YEAR="";
		String S_MONTH="";
		if(args.length == 6){
		   if(args[4] != null) S_YEAR = args[4];
		   if(args[5] != null) S_MONTH = args[5];
		}
		a.exeZZ041W(report_no,lguser_id,lguser_name,unit,S_YEAR,S_MONTH);
		System.out.println("AutoZZ041W end");
	}
	public void exeZZ041W(String report_no,String lguser_id, String lguser_name,String unit,String S_YEAR,String S_MONTH){
	    try{
		       //URL myURL=new URL("http://localhost:81/pages/WMAutoZZ041W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&unit="+unit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
		       URL myURL=new URL(Utility.getProperties("MISWebSite")+"WMAutoZZ041W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&unit="+unit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
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

