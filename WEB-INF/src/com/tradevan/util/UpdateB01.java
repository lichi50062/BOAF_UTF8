//99.11.16 add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  寫入WML03的參數移至共用create_dataList by 2295
package com.tradevan.util;


import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateB01{
	private static String errMsg = "";
	
	public String getErrMsg(){
		return errMsg;  
	}
	
	public static synchronized String doParserReport_B01(String report_no, String m_year,
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
			dbData = Utility.getWML03_count(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);//99.11.19 fix
			/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)*/
			Utility.printLogTime("B01-1 time");	
			if(dbData.size() != 0){//WML03有資料時,先清掉
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D'");
				sqlCmd.append(" from WML03");
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
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
			Utility.printLogTime("B01-2 time");	
			
			//WML03 (檢查合計和小計)			
			updateDBDataList = new ArrayList();			
			if(errCount == 0){//無其他錯誤時,才檢核合計金額
				System.out.println("無其他錯誤時,才檢核合計金額");
				String subErrMsg="";
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				//檢核小計(不含農業發展基金，即1、農機貸款+2、擴大家庭農場經營規模協助農民購買耕地貸款+3、輔導修建農宅貸款+4、加速農村建設貸款)
				System.out.println("@@檢核小計");
				
				sqlCmd.append("select decode(fund_master_no||fund_sub_no||fund_next_no,'010490','0','1') as \"fund_no\", ");
				sqlCmd.append(" sum(budget_amt) as \"budget_amt\", ");
				sqlCmd.append(" sum(credit_pay_amt) as \"credit_pay_amt\", ");
				sqlCmd.append(" sum(credit_pay_rate) as \"credit_pay_rate\" ");
				sqlCmd.append(" from	B01 ");
				sqlCmd.append("where m_year=? and m_month=?");
				sqlCmd.append("	and fund_master_no||fund_sub_no||fund_next_no in ('010100','010200','010300','010400','010490') ");
				sqlCmd.append("group by decode(fund_master_no||fund_sub_no||fund_next_no,'010490','0','1') ");
				sqlCmd.append("order by decode(fund_master_no||fund_sub_no||fund_next_no,'010490','0','1') desc");
				
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,budget_amt, credit_pay_amt,credit_pay_rate");
				long s_budget_amt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("budget_amt")).toString());
				long s_credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_amt")).toString());
				long s_credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_rate")).toString());
				
				long st_budget_amt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("budget_amt")).toString());
				long st_credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_amt")).toString());
				long st_credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_rate")).toString());
				
							
				
				if(s_budget_amt != st_budget_amt	) { //比對筆數金額					  
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"預算加總與小計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(s_credit_pay_amt != st_credit_pay_amt ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放加總與小計列不合",user_id,user_name));								
					errCount++ ;										
				}
				if(s_credit_pay_rate	!= st_credit_pay_rate	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放率加總與小計列不合",user_id,user_name));								
					errCount++ ;					
				}
				Utility.printLogTime("B01-3 time");	
				//檢查加速農村建設貸款
				System.out.println("@@檢查加速農村建設貸款");
				sqlCmd.delete(0,sqlCmd.length());				
				sqlCmd.append("select decode(fund_master_no||fund_sub_no||fund_next_no,'010400','0','1') as \"fund_no\", ");
				sqlCmd.append( " sum(budget_amt) as \"budget_amt\", ");
				sqlCmd.append(" sum(credit_pay_amt) as \"credit_pay_amt\", ");
				sqlCmd.append(" sum(credit_pay_rate) as \"credit_pay_rate\" ");
				sqlCmd.append(" from B01 ");
				sqlCmd.append("where m_year=? and m_month=?");
				sqlCmd.append("	and fund_master_no||fund_sub_no||fund_next_no not in ('010000','010100','010200','010300','010490','020000','030000') ");
				sqlCmd.append("group by decode(fund_master_no||fund_sub_no||fund_next_no,'010400','0','1') ");
				sqlCmd.append("order by decode(fund_master_no||fund_sub_no||fund_next_no,'010400','0','1') desc");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,budget_amt, credit_pay_amt,credit_pay_rate");
				long vs_budget_amt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("budget_amt")).toString());
				long vs_credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_amt")).toString());
				long vs_credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_rate")).toString());
				
				long vst_budget_amt			= Long.parseLong((((DataObject)dbData.get(1)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("budget_amt")).toString());
				long vst_credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_amt")).toString());
				long vst_credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_rate")).toString());
				if(vs_budget_amt != vst_budget_amt	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"預算加總與加速農村建設貸款金額不合",user_id,user_name));								
					errCount++ ;					
				}
				if(vs_credit_pay_amt != vst_credit_pay_amt ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放加總與加速農村建設貸款金額不合",user_id,user_name));								
					errCount++ ;										
				}
				if(vs_credit_pay_rate	!= vst_credit_pay_rate	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放率加總與加速農村建設貸款金額不合",user_id,user_name));								
					errCount++ ;
				}
				Utility.printLogTime("B01-4 time");					
				//檢查合計列資料
				System.out.println("@@檢查合計列資料");
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append("select decode(fund_master_no||fund_sub_no||fund_next_no,'030000','0','1') as \"type\", ");
				sqlCmd.append("      sum(budget_amt) as \"budget_amt\", ");
				sqlCmd.append("      sum(credit_pay_amt) as \"credit_pay_amt\", ");
				sqlCmd.append("      sum(credit_pay_rate) as \"credit_pay_rate\" ");
				sqlCmd.append(" from	B01 ");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" 	and fund_master_no||fund_sub_no||fund_next_no in ('010490','020000','030000') ");
				sqlCmd.append("group by decode(fund_master_no||fund_sub_no||fund_next_no,'030000','0','1') ");
				sqlCmd.append("order by decode(fund_master_no||fund_sub_no||fund_next_no,'030000','0','1') desc");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,budget_amt, credit_pay_amt,credit_pay_rate");		
					
				long budget_amt			= Long.parseLong((((DataObject)dbData.get(0)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("budget_amt")).toString());
				long credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_amt")).toString());
				long credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(0)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(0)).getValue("credit_pay_rate")).toString());
					
				long t_budget_amt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("budget_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("budget_amt")).toString());
				long t_credit_pay_amt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_amt")).toString());
				long t_credit_pay_rate	= Long.parseLong((((DataObject)dbData.get(1)).getValue("credit_pay_rate")	==null?"0":((DataObject)dbData.get(1)).getValue("credit_pay_rate")).toString());
				if(budget_amt != t_budget_amt	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"預算加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}
				if(credit_pay_amt != t_credit_pay_amt ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(credit_pay_rate	!= t_credit_pay_rate	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放率加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}
				Utility.printLogTime("B01-5 time");		
				if(updateDBDataList != null && updateDBDataList.size()!=0){
				   updateDBSqlList = new ArrayList();	
				   sqlCmd.delete(0,sqlCmd.length());
				   sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
				   updateDBSqlList.add(sqlCmd.toString());
				   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
				}
			}//enf of errCount==0
			dbData =  Utility.getWML01(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data*/
			Utility.printLogTime("B01-6 time");		
			upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
			checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
			
			//99.11.19 fix
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,String.valueOf(Integer.parseInt(m_year)), String.valueOf(Integer.parseInt(m_month)),bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);
			for(int i=0;i<getUpdateDBList.size();i++){
			    updateDBList.add(getUpdateDBList.get(i));
			}
			/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update*/
			Utility.printLogTime("B01-7 time");	
						
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
				errMsg = errMsg + "UpdateB01.doParserReport_B01 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			Utility.printLogTime("B01-8 time");		
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateB01.doParserReport_B01 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateB01.doParserReport_B01="+e.getMessage());
		}
		
		return errMsg;
		//return parserResult;
	}
}


