//99.11.16 add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update by 2295
package com.tradevan.util;

import java.util.*;
import java.text.SimpleDateFormat;

import com.tradevan.util.dao.DataObject;


public class UpdateB02{
	private static String errMsg = "";
	
	public String getErrMsg(){
		return errMsg;  
	}
	
	public static synchronized String doParserReport_B02(String report_no, String m_year,
		String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type,String szuser_id,String szuser_name,int batch_no) throws Exception {
		//boolean parserResult=true;
		errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		int		errCount = 0;		
		String  user_id="";//本次異動者帳號
		String  user_name="";//本次異動者姓名
		String[] user_data = {"",""};
		String  bank_code=srcbank_code; 	
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 
		
		//99.11.16 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":"";
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		DataObject bean = null;
	    //=====================================================================
		
		try { 

			user_id = szuser_id;
		    user_name = szuser_name;
			errCount = 0;
			upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗		
			dbData =  Utility.getWML01(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data*/
			//99.11.19 fix
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,String.valueOf(Integer.parseInt(m_year)), String.valueOf(Integer.parseInt(m_month)),bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);
			for(int i=0;i<getUpdateDBList.size();i++){
			    updateDBList.add(getUpdateDBList.get(i));
			}
			/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update*/
			
			//傳送email至bank_cmml的e-mail信箱
			System.out.println("send email begin");		
			//線上編輯
			Utility.sendMailNotification(bank_code,report_no,
			m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");
			System.out.println("send email end");
			
			if(updateDBList != null && updateDBList.size()!=0){		
				/*
				for(int j=0;j<updateDBList.size();j++){
		    		System.out.println((String)((List)updateDBList.get(j)).get(0));
		    		System.out.println((List)((List)updateDBList.get(j)).get(1));
		    	}
		    	*/				
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateB02.doParserReport_B02 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
				
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateB02.doParserReport_B02 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateB02.doParserReport_B02="+e.getMessage());			
		}		
		return errMsg;
		//return parserResult;
	}
}


