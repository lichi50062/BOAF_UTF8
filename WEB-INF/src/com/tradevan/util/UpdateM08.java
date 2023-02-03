//94.04.22 fix 只有線上編輯模式,都為"0"值時,顯示檢核為"0" by 2295
//99.11.18 寫入WML03的參數移至共用create_dataList;add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update by 2295

package com.tradevan.util;


import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateM08 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M08(String report_no, String m_year,
			String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type,String szuser_id,String szuser_name,int batch_no)
			throws Exception {
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
		
		List M08List = new LinkedList();//M08細部資料
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true";
		boolean nonZero=false;//94.04.22 檢查內容是否都為"0"
		
		//99.11.18 add 查詢年度100年以前.縣市別不同================================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":"";
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data		
	    //=====================================================================
		
		try { 

			user_id = szuser_id;
		    user_name = szuser_name;
			errCount = 0;

			//clear WML03(檢核其他錯誤)
			dbData = Utility.getWML03_count(m_year,m_month,bank_code,report_no);//99.11.19 fix
			/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)*/
			if(dbData.size() != 0){//WML03有資料時,先清掉
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D'");
				sqlCmd.append(" from WML03");
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				updateDBDataList = new ArrayList();
				dataList =  new ArrayList();
				dataList.add(user_id);
				dataList.add(user_name);	
				dataList.add(String.valueOf(Integer.parseInt(m_year)));	
				dataList.add(String.valueOf(Integer.parseInt(m_month)));	
				dataList.add(bank_code);	
				dataList.add(report_no);
				updateDBDataList.add(dataList);
				updateDBSqlList.add(sqlCmd.toString());
				updateDBSqlList.add(updateDBDataList);
				updateDBList.add(updateDBSqlList);
				
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				
				sqlCmd.append("DELETE FROM WML03 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				dataList =  new ArrayList();				
				dataList.add(String.valueOf(Integer.parseInt(m_year)));	
				dataList.add(String.valueOf(Integer.parseInt(m_month)));	
				dataList.add(bank_code);	
				dataList.add(report_no);
				updateDBDataList.add(dataList);
				updateDBSqlList.add(sqlCmd.toString());
				updateDBSqlList.add(updateDBDataList);
				updateDBList.add(updateDBSqlList);	
			}
			
			sqlCmd.delete(0,sqlCmd.length());
			paramList = new ArrayList();			
		    sqlCmd.append(" select id_no,data_range,guarantee_no_month,loan_amt_month,guarantee_amt_month,");
		    sqlCmd.append("        guarantee_bal_month,guarantee_bal_p");
		    sqlCmd.append("   from M08");
		    sqlCmd.append(" WHERE m_year=? AND m_month=?");
		    paramList.add(String.valueOf(Integer.parseInt(m_year)));
			paramList.add(String.valueOf(Integer.parseInt(m_month)));      
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_no_month,loan_amt_month,guarantee_amt_month,guarantee_bal_month,guarantee_bal_p");     
	    	checkZeroLoop:
		    if(dbData != null && dbData.size() != 0){
		    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
		    		if(	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_month")).toString()) > 0
					||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_month")).toString()) > 0
					||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_month")).toString()) > 0
					||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_month")).toString()) > 0
					||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_p")).toString()) > 0											
		    		){
			    	   nonZero = true;
			    	   break checkZeroLoop;
			    	}
		    	}
		    }
	    	System.out.println("線上編輯.nonZero="+nonZero);
			
			//WML03 (檢查合計和小計)
			if(nonZero && errCount == 0){//無其他錯誤時,才檢核合計金額
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append(" select decode(id_no,'0','0','1') as \"type\", ");
                sqlCmd.append("       a.data_range, ");
                sqlCmd.append("       data_range_name || '加總與總計不合' as \"errmsg\", ");
                sqlCmd.append("       sum(guarantee_no_month) as \"guarantee_no_month\", ");
                sqlCmd.append("       sum(loan_amt_month) as \"loan_amt_month\", ");
                sqlCmd.append("       sum(guarantee_amt_month) as \"guarantee_amt_month\", ");
                sqlCmd.append("       sum(guarantee_bal_month) as \"guarantee_bal_month\", ");
                sqlCmd.append("       sum(guarantee_bal_p) as \"guarantee_bal_p\" ");
                sqlCmd.append(" from  M08 a,m00_data_range_item b ");
                sqlCmd.append(" where m_year=? and m_month=?");
                sqlCmd.append(" and   substr(a.data_range,4,2)=b.data_range ");
                sqlCmd.append(" and   b.report_no='M08' ");
                sqlCmd.append(" group by decode(id_no,'0','0','1'),a.data_range,b.data_range_name || '加總與總計不合' ");
                sqlCmd.append(" order by decode(id_no,'0','0','1') desc,data_range_name || '加總與總計不合',a.data_range,b.data_range_name || '加總與總計不合' ");
                paramList.add(String.valueOf(Integer.parseInt(m_year)));
                paramList.add(String.valueOf(Integer.parseInt(m_month)));
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,loan_amt_month,guarantee_amt_month,guarantee_bal_month,guarantee_bal_p");		
				for(int j=0;j<3;j++){
					long guarantee_no_month	 = Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_no_month")).toString());
					long loan_amt_month		 = Long.parseLong((((DataObject)dbData.get(j)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(j)).getValue("loan_amt_month")).toString());
					long guarantee_amt_month = Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_amt_month")).toString());
					long guarantee_bal_month = Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_bal_month")	==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_bal_month")).toString());
					long guarantee_bal_p  	 = Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_bal_p")).toString());

					long t_guarantee_no_month  = Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_no_month")).toString());
					long t_loan_amt_month	   = Long.parseLong((((DataObject)dbData.get(j+3)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(j+3)).getValue("loan_amt_month")).toString());
					long t_guarantee_amt_month = Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_amt_month")).toString());
					long t_guarantee_bal_month = Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_bal_month")	==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_bal_month")).toString());
					long t_guarantee_bal_p     = Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_bal_p")).toString());
					
					
					updateDBDataList = new ArrayList();	
					System.out.println("guarantee_no_month="+guarantee_no_month);
					System.out.println("t_guarantee_no_month="+t_guarantee_no_month);
					if(guarantee_no_month  != t_guarantee_no_month  ) {
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證件數:"+(String)((DataObject)dbData.get(j)).getValue("errmsg"),user_id,user_name));								
						errCount++ ;						
					}

					if(loan_amt_month	   != t_loan_amt_month		) {
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"融資金額:"+(String)((DataObject)dbData.get(j)).getValue("errmsg"),user_id,user_name));								
						errCount++ ;						
					}

					if(guarantee_amt_month != t_guarantee_amt_month ) {
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證金額:"+(String)((DataObject)dbData.get(j)).getValue("errmsg"),user_id,user_name));								
						errCount++ ;						
					}

					if(guarantee_bal_month != t_guarantee_bal_month	) {
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證餘額:"+(String)((DataObject)dbData.get(j)).getValue("errmsg"),user_id,user_name));								
						errCount++ ;						
					}

					if(guarantee_bal_p     != t_guarantee_bal_p   	) {
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證餘額(結構比):"+(String)((DataObject)dbData.get(j)).getValue("errmsg"),user_id,user_name));								
						errCount++ ;						
					}
					if(updateDBDataList != null && updateDBDataList.size()!=0){
						updateDBSqlList = new ArrayList();	
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
						updateDBSqlList.add(sqlCmd.toString());
						updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
					}	
				}
			}//enf of errCount==0
			dbData =  Utility.getWML01(m_year,m_month,bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data*/
			upd_code = (errCount == 0)?"U":"E";//U 檢核成功;E 檢核失敗
			checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
			//94.04.22 fix Z檢核為0====================================
			if(!nonZero && errCount == 0){
				upd_code = "Z";//Z檢核為0
				checkResult="Z";//Z檢核為0
			}
			//========================================================
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);//99.11.19 fix
			for(int j=0;j<getUpdateDBList.size();j++){
			    updateDBList.add(getUpdateDBList.get(j));
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
				errMsg = errMsg + "UpdateM08.doParserReport_M08 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
		}catch (Exception e) {			
			errMsg = errMsg + "UpdateM08.doParserReport_M08 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM08.doParserReport_M08="+e.getMessage());			
		}
		return errMsg;		
	}
}
