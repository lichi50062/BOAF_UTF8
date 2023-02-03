
/*
  94.09.27 add 各機構基本資料 by 2495
 100.07.21 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			         使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 101.09.04 fix 信用部員工人數,與BR002W(從事信用業務之員工人數(含總機構及分支機構))規則相同 by 2295
*/
package com.tradevan.util;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class AutoHeadPerson{	
	public static void main(String[] args){
		System.out.println("AutoHeadPerson begin");
		AutoHeadPerson a = new AutoHeadPerson();
		a.exeHeadPerson();
		System.out.println("AutoHeadPerson end");
	}


	public void exeHeadPerson()
	{
	 String errMsg = "",sqlCmd="";	
	 String cd01_table = "";
	 String cd02_table = "";
     String wlx01_m_year = "";
     List paramList = new ArrayList(); 
     String bank_name = "";
     String bank_no = "";
     String bank_name_a = ""; 
     String bank_name_b = ""; 
     String name_1 = "";
     String name_2 = "";
     String name_3 = "";
     String name_4 = "";
     String addr_1 = ""; 
     String addr_2 = ""; 
     String addr_3 = ""; 
     String telno = "";
     String fax = "";
     String web_site = "";
     String email = "";
     String tstaff_num = "";
	try{
		//100.07.21 add 查詢年度100年以前.縣市別不同===============================
	    cd01_table = (Integer.parseInt(Utility.getYear()) < 100)?"cd01_99":""; 
	    cd02_table = (Integer.parseInt(Utility.getYear()) < 100)?"cd02_99":"";
	    wlx01_m_year = (Integer.parseInt(Utility.getYear()) < 100)?"99":"100"; 
	   	//=====================================================================    

      sqlCmd =    " Select  *  from ("      
             +" Select  replace(nvl(bd.BANK_NAME, ' '), '!', ' ') as BANK_NAME, "
		     +" replace(nvl(bd.BANK_NO,' '), '!', ' ')		as BANK_NO," 
			 +" replace(nvl(bd.BANK_NAME_A,' '), '!', ' ')	as BANK_NAME_A,"  
			 +" replace(nvl(bd.BANK_NAME_B,' '), '!', ' ')  as BANK_NAME_B," 
          	 +" replace(nvl(bd.addr_1,' '), '!', ' ')       as addr_1,"  
			 +" replace(nvl(bd.addr_2,' '), '!', ' ')       as addr_2,"
			 +" replace(nvl(bd.addr_3,' '), '!', ' ')       as addr_3,"    
			 +" replace(nvl(bd.TELNO,' '), '!', ' ')        as TELNO," 
			 +" replace(nvl(bd.FAX,' '), '!', ' ')          as FAX,"
			 +" replace(nvl(bd.WEB_SITE,' '), '!', ' ')     as WEB_SITE," 
			 +" replace(nvl(bd.EMAIL,' '), '!', ' ')        as EMAIL,"
	         +" replace(nvl(bd.name_1,' '), '!', ' ')       as name_1," 
			 +" replace(nvl(bd.name_2,' '), '!', ' ')       as name_2,"  
			 +" replace(nvl(WLX01_M.name,' '), '!', ' ')    as name_3,"
			 +" ' ' as name_4,"  
			 +" replace(nvl(bd.TSTAFF_NUM,'0'), '!', '0')   as TSTAFF_NUM"
		     +" from (select bc.* , WLX01_M.name 			as name_2"
			 +" from (select bb.* , WLX01_M.name 			as name_1"
			 +" from (select bn01_1.BANK_NAME,wlx01.BANK_NO,bn01_2.BANK_NAME as BANK_NAME_A,"  
			 +" bn01.BANK_NAME as BANK_NAME_B,"
			 +" cd01.HSIEN_NAME as addr_1," 
			 +" cd02.AREA_NAME as addr_2," 
			 +" wlx01.ADDR  as addr_3," 
			 +" wlx01.TELNO, wlx01.FAX, wlx01.WEB_SITE , wlx01.EMAIL,wlx01.TSTAFF_NUM"
			 +" from ( select wlx01.*, INP_NO.Tstaff_num " 
             +" from ( select wlx01.*" 
             +" from (select * from bn01 where m_year=?)bn01, (select * from wlx01 where m_year=?)wlx01 where bn01.bank_type in ('6', '7') "  
             +" and  bn01.bank_no = wlx01.bank_no  and wlx01.CANCEL_NO <> 'Y') wlx01  "
             +" left join (select BANK_NO,  sum(CREDIT_STAFF_NUM)   as Tstaff_num "
             +" from (select BANK_NO, CREDIT_STAFF_NUM " //101.09.04 fix 信用部員工人數,與BR002W(從事信用業務之員工人數(含總機構及分支機構))規則相同
             +" from (select * from wlx01 where m_year=?)wlx01  WHERE wlx01.CANCEL_NO <> 'Y' "
             +" )INP_NO  group by bank_no ) INP_NO "
             +" on wlx01.bank_no = INP_NO.BANK_NO)   wlx01"
             +" left join (select * from bn01 where m_year=?)bn01_1  on  wlx01.BANK_NO=bn01_1.BANK_NO and  bn01_1.bank_type in ('6', '7') "
             +" left join (select * from bn01 where m_year=?)bn01 on  wlx01.CENTER_NO = bn01.BANK_NO AND bn01.bank_type = '8'" 		   
             +" left join "+cd01_table+" cd01 on  wlx01.HSIEN_ID = CD01.HSIEN_ID"
	         +" left join "+cd02_table+"cd02 on  wlx01.AREA_ID =  CD02.AREA_ID"
             +" left join (select * from bn01 where m_year=?)bn01_2 on wlx01.M2_NAME =  bn01_2.BANK_NO"
             +" and  bn01_2.bank_type = 'B') bb "
             +" left join WLX01_M  on bb.BANK_NO =  WLX01_M.BANK_NO "
             +" and  WLX01_M.ABDICATE_CODE <> 'Y' and  WLX01_M.POSITION_CODE = '1') bc"
             +" left join WLX01_M  on bc.BANK_NO =  WLX01_M.BANK_NO"
             +" and  WLX01_M.ABDICATE_CODE <> 'Y'"
             +" and  WLX01_M.POSITION_CODE = '2') bd"
             +" left join WLX01_M  on bd.BANK_NO =  WLX01_M.BANK_NO"
             +" and  WLX01_M.ABDICATE_CODE <> 'Y'"
             +" and WLX01_M.POSITION_CODE = '3') Search_End"
             +" where BANK_NO <> ' '   order by  BANK_NO";	
        paramList.add(wlx01_m_year);
        paramList.add(wlx01_m_year);
        paramList.add(wlx01_m_year);
        paramList.add(wlx01_m_year);
        paramList.add(wlx01_m_year);
        paramList.add(wlx01_m_year);
        
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"bank_name,bank_no,bank_name_a,bank_name_b,name_1,name_2,name_3,name_4,addr_1,addr_2,addr_3,telno,fax,web_site,email,tstaff_num");
        
		String Path = Utility.getProperties("headPersonDir");	
		File objFile = new File(Path);    	 
               
        if(dbData.size()==0)
        {
		objFile.delete();
        }
        else
        {		        
       		String Line="";
                FileWriter objWriter = new FileWriter(Path);
		Line =Integer.toString(dbData.size());
 		objWriter.write(Line);
		objWriter.write(13);
 		objWriter.write(10);
		for(int i=0;i<dbData.size();i++)
		{
			bank_name = "";bank_no = "";bank_name_a = ""; bank_name_b = ""; name_1 = "";
			name_2 = "";name_3 = "";name_4 = "";addr_1 = ""; addr_2 = ""; addr_3 = ""; 
			telno = "";fax = "";web_site = "";email = "";tstaff_num = "";
			
			bank_name = (((DataObject)dbData.get(i)).getValue("bank_name")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("bank_name");	  	   		
			bank_no = (((DataObject)dbData.get(i)).getValue("bank_no")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("bank_no");
			//System.out.println("bank_no ="+bank_no);
			bank_name_a = (((DataObject)dbData.get(i)).getValue("bank_name_a")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("bank_name_a");
			//System.out.println("bank_name_a ="+bank_name_a);
			
			bank_name_b = (((DataObject)dbData.get(i)).getValue("bank_name_b")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("bank_name_b");
			//System.out.println("bank_name_b ="+bank_name_b);
			name_1 = (((DataObject)dbData.get(i)).getValue("name_1")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("name_1");
			//System.out.println("name_1 ="+name_1);
			name_2 = (((DataObject)dbData.get(i)).getValue("name_2")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("name_2");	
			//System.out.println("name_2 ="+name_2);
			name_3 = (((DataObject)dbData.get(i)).getValue("name_3")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("name_3");
			//System.out.println("name_3 ="+name_3);
			name_4 = (((DataObject)dbData.get(i)).getValue("name_4")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("name_4");	
			//System.out.println("name_4 ="+name_4);
			addr_1 = (((DataObject)dbData.get(i)).getValue("addr_1")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("addr_1");	  	   		
			//System.out.println("addr_1 ="+addr_1);
			addr_2 = (((DataObject)dbData.get(i)).getValue("addr_2")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("addr_2");	  	   		
			//System.out.println("addr_2 ="+addr_2);
			addr_3 = (((DataObject)dbData.get(i)).getValue("addr_3")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("addr_3");	  	   		
			//System.out.println("addr_3 ="+addr_3);
			telno = (((DataObject)dbData.get(i)).getValue("telno")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("telno");
			//System.out.println("telno ="+telno);
			fax = (((DataObject)dbData.get(i)).getValue("fax")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("fax");
			//System.out.println("fax ="+fax);
			web_site = (((DataObject)dbData.get(i)).getValue("web_site")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("web_site");
			//System.out.println("web_site ="+web_site);
			email = (((DataObject)dbData.get(i)).getValue("email")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("email");
			//System.out.println("email ="+email);
			tstaff_num = (((DataObject)dbData.get(i)).getValue("tstaff_num")==null ) ? " " : (String)((DataObject)dbData.get(i)).getValue("tstaff_num");	
			//System.out.println("tstaff_num ="+tstaff_num);
			
			Line = bank_name+"!"+bank_no+"!"+bank_name_a+"!"+bank_name_b+"!"+addr_1+"!"+addr_2+"!"+addr_3+"!"+telno+"!"+fax+"!"+web_site+"!"+email+"!"+name_1+"!"+name_2+"!"+name_3+"!"+name_4+"!"+tstaff_num;
			System.out.println("Line ="+Line);
			
			objWriter.write(Line);
			objWriter.write(13);
 			objWriter.write(10);
		}
 		objWriter.close();
       }
     }catch (Exception e){
			    System.out.println(e.getMessage());
     } 
    
     }		
}



