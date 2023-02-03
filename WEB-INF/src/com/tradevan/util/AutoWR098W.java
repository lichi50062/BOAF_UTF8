/*
 * Created on 2013/08/06
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import com.tradevan.util.Utility;

public class AutoWR098W {
	public static void main(String[] args){
		System.out.println("AutoWR098W begin");
		AutoWR098W a = new AutoWR098W();
		
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
		a.exeWR098W(report_no,lguser_id,lguser_name,unit,S_YEAR,S_MONTH);
		System.out.println("AutoWR098W end");
	}
	public void exeWR098W(String report_no,String lguser_id, String lguser_name,String unit,String S_YEAR,String S_MONTH){
	    try{
		       //URL myURL=new URL("http://localhost:81/pages/WMAutoWR098W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&unit="+unit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
		       URL myURL=new URL(Utility.getProperties("MISWebSite")+"WMAutoWR098W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&unit="+unit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
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

