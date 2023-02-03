/*
  105.03.09 add 排程產檔至官網提供官網公示揭露查詢 by 2968
*/
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class AutoOffice_exe {
	public static void main(String[] args) {
		System.out.println("AutoOffice_exe begin");
		AutoOffice_exe a = new AutoOffice_exe();
		a.exeCheck();
		System.out.println("AutoOffice_exe end");
	}

	public void exeCheck() {
		try {
			URL myURL = new URL(Utility.getProperties("BOAFWebSite") + "WMAutoOffice.jsp");
			// URL myURL=new
			// URL("https://ebankweb.boaf.gov.tw/pages/test.jsp");//農金局正式環境
			// URL myURL=new URL("http://localhost/pages/test.jsp");//測試環境
			URLConnection raoURL;
			raoURL = myURL.openConnection();
			raoURL.setDoInput(true);
			raoURL.setDoOutput(true);
			raoURL.setUseCaches(false);
			BufferedReader infromURL = new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
			String raoInputString;
			while ((raoInputString = infromURL.readLine()) != null) {
				System.out.println(raoInputString);
			}
			infromURL.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
