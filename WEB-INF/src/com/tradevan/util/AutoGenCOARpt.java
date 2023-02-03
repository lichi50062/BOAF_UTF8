/*
 * 105.12.07 農委會OpenData報表發佈排程 by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import com.tradevan.util.Utility;


public class AutoGenCOARpt {
	public static void main(String[] args){
		System.out.println("AutoGenCOARpt begin");
		AutoGenCOARpt a = new AutoGenCOARpt();
		if(args.length == 2){		
		    System.out.println("s_year="+args[0]);
	        System.out.println("s_month="+args[1]);
            a.exeRpt(args[0],args[1]);
		}else{
		    a.exeRpt("","");
		}
		
		
		System.out.println("AutoGenCOARpt end");
	}
	public void exeRpt(String S_YEAR,String S_MONTH){
	    try{
		       //URL myURL=new URL("http://localhost:81/pages/WMAutoZZ041W.jsp?report_no="+report_no+"&lguser_id="+lguser_id+"&lguser_name="+lguser_name+"&unit="+unit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
		       URL myURL=new URL(Utility.getProperties("MISWebSite")+"WMAutoGenCOARpt.jsp?S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
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

