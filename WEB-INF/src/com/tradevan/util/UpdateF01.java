//在台無住所外國人新台幣存款只有線上編輯,無檔案上傳
//99.11.16 add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295

package com.tradevan.util;

import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateF01 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_F01(String report_no, String m_year,
			String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type,String szuser_id,String szuser_name,int batch_no)
			throws Exception {
		//boolean parserResult=true;
		errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();					
		int		errCount = 0;		
		String  user_id="";//本次異動者帳號
		String  user_name="";//本次異動者姓名
		String[] user_data = {"",""};
		String  bank_code=srcbank_code;//農業信用保證基金 	
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
		boolean nonZero=false;//94.04.22 檢查內容是否都為"0"		
		
		//99.11.15 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":"";
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		DataObject bean = null;
	    //=====================================================================
			
		
		try { 
			/*
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				M01List = getM01FileData(report_no,filename);
				user_data = getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
			}else{*/
				user_id = szuser_id;
				user_name = szuser_name;
			//}
			
			errCount = 0;
			dbData = Utility.getWML03_count(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);//99.11.19 fix
			/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)
			
			sqlCmd.append("select * from WML03 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			
			paramList.add(String.valueOf(Integer.parseInt(m_year)));
			paramList.add(String.valueOf(Integer.parseInt(m_month)));
			paramList.add(bank_code);
			paramList.add(report_no);			
			
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,serial_no,update_date");
			*/
			Utility.printLogTime("F01-1 time");	
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
			
			Utility.printLogTime("F01-2 time");	
			//clear WML03(檢核其他錯誤)			
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
				nonZero = false;
				dbData = Utility.getCountZero(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no,"(acct_cnt_tm > 0 or bal_lm > 0 or dep_tm >0 or wtd_tm >0 or bal_tm >0)");//99.11.19 fix
				/*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數				
				
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
		    	
		    	//99.11.16檢核內容是否都為0
				sqlCmd.append(" select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? ");
				sqlCmd.append(" and (acct_cnt_tm > 0 or bal_lm > 0 or dep_tm >0 or wtd_tm >0 or bal_tm >0)");
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month))	);
				paramList.add(bank_code);
				
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
		    	*/		    		     
				if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					        nonZero = true;
					   }
				}  
		    	System.out.println("線上編輯.nonZero="+nonZero);
		    }//end of 線上編輯check non zero
			Utility.printLogTime("F01-3 time");	
		    //94.04.22 fix 只要有非"0"值時,才執行檢核			
			//WML03 (檢查合計和小計)
			if(nonZero && errCount == 0){//無其他錯誤時,才檢核合計金額				
				String[] column={"acct_cnt_tm","bal_lm","dep_tm","wtd_tm","bal_tm"};                          
				long[][] sumE = new long[4][4];
				long[][] E = new long[4][4];
				int j=0;
				long tmpAmt=0;
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append(" select dep_type,acct_type,acct_cnt_tm,bal_lm,dep_tm,wtd_tm"); 
				sqlCmd.append(" from   F01 ");
				sqlCmd.append(" where m_year=? and m_month=? and bank_code=?");
				sqlCmd.append(" order by dep_type,acct_type ");
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month))	);
				paramList.add(bank_code);
				
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acct_cnt_tm,bal_lm,dep_tm,wtd_tm");
				System.out.println("dbData.size()="+dbData.size());
				for(int i=0;i<dbData.size();i++){
					bean = (DataObject)dbData.get(i);
					if(i == 4 || i==9 || i==14 || i==19 || i==24){
						j=0;
						continue;
					}else if(i >= 20 && i < 24){						
						j=(i==20)?0:++j;						
						E[j][0] = Long.parseLong((bean.getValue("acct_cnt_tm")==null?"0":bean.getValue("acct_cnt_tm")).toString());
						E[j][1] = Long.parseLong((bean.getValue("bal_lm")==null?"0":bean.getValue("bal_lm")).toString());
						E[j][2] = Long.parseLong((bean.getValue("dep_tm")==null?"0":bean.getValue("dep_tm")).toString());
						E[j][3] = Long.parseLong((bean.getValue("wtd_tm")==null?"0":bean.getValue("wtd_tm")).toString());
						System.out.println("i="+i+" E["+j+"][0]="+E[j][0]);
						System.out.println("i="+i+" E["+j+"][1]="+E[j][1]);
						System.out.println("i="+i+" E["+j+"][2]="+E[j][2]);
						System.out.println("i="+i+" E["+j+"][3]="+E[j][3]);
					}else{
					    sumE[j][0] +=  Long.parseLong((bean.getValue("acct_cnt_tm")==null?"0":bean.getValue("acct_cnt_tm")).toString());
					    sumE[j][1] +=  Long.parseLong((bean.getValue("bal_lm")==null?"0":bean.getValue("bal_lm")).toString());
					    sumE[j][2] +=  Long.parseLong((bean.getValue("dep_tm")==null?"0":bean.getValue("dep_tm")).toString());
					    sumE[j][3] +=  Long.parseLong((bean.getValue("wtd_tm")==null?"0":bean.getValue("wtd_tm")).toString());
					    System.out.println("i="+i+" sumE["+j+"][0]="+sumE[j][0]);
						System.out.println("i="+i+" sumE["+j+"][1]="+sumE[j][1]);
						System.out.println("i="+i+" sumE["+j+"][2]="+sumE[j][2]);
						System.out.println("i="+i+" sumE["+j+"][3]="+sumE[j][3]);
					   j++;
					}
				}
				System.out.println("get data end");
				for(int k =0;k<4;k++){					 
					System.out.println("sumE["+k+"][1]="+sumE[k][0]);
					System.out.println("sumE["+k+"][2]="+sumE[k][1]);
					System.out.println("sumE["+k+"][3]="+sumE[k][2]);
					System.out.println("sumE["+k+"][4]="+sumE[k][3]);					
				}
				for(int k =0;k<4;k++){					 
					System.out.println("E["+k+"][1]="+E[k][0]);
					System.out.println("E["+k+"][2]="+E[k][1]);
					System.out.println("E["+k+"][3]="+E[k][2]);
					System.out.println("E["+k+"][4]="+E[k][3]);					
				}
				String subErrMsg="";	
				String subtitle[] = {"個人","公司、行號、團體","外國專業投資機構","外國銀行"};
				String item[] = {"本月底戶數","上月底餘額","本月存入","本用提出"};
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();				
				sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
				for(int k =0;k<4;k++){
					for(int k1 =0;k1<4;k1++){
				        if(E[k][k1] != sumE[k][k1]){				    	
				    	   subErrMsg=item[k1]+"的"+"("+subtitle[k]+")合計數:與活期、活期儲蓄、定期、支票等存款的加總不符";
				        }
					    if(subErrMsg.length()!=0){
					        dataList =  new ArrayList();
					        dataList.add(String.valueOf(Integer.parseInt(m_year)));
							dataList.add(String.valueOf(Integer.parseInt(m_month)));
							dataList.add(bank_code);
							dataList.add(report_no);
							dataList.add(String.valueOf(errCount));//errCount
							dataList.add(subErrMsg);
							dataList.add(user_id);
							dataList.add(user_name);
							updateDBDataList.add(dataList);//1:傳內的參數List		
							
							errCount++;
						    subErrMsg = "";
					   }
					}//column
				}//row 
				if(updateDBDataList != null && updateDBDataList.size()!=0){
				   updateDBSqlList.add(sqlCmd.toString());
				   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
				}
			}//enf of errCount==0
			Utility.printLogTime("F01-4 time");	
			dbData =  Utility.getWML01(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data
			sqlCmd.delete(0,sqlCmd.length());
			paramList = new ArrayList();
			sqlCmd.append("select * from WML01 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
		 	paramList.add(String.valueOf(Integer.parseInt(m_year)));
			paramList.add(String.valueOf(Integer.parseInt(m_month)));
			paramList.add(bank_code);
			paramList.add(report_no);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
			*/	
			upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
			checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
			//94.04.22 fix Z檢核為0====================================
			if(!nonZero && errCount == 0){
				upd_code = "Z";//Z檢核為0
				checkResult="Z";//Z檢核為0
			}
			//========================================================
			//99.11.19 fix
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,String.valueOf(Integer.parseInt(m_year)), String.valueOf(Integer.parseInt(m_month)),bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);
			for(int i=0;i<getUpdateDBList.size();i++){
			    updateDBList.add(getUpdateDBList.get(i));
			}
			/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update			
			sqlCmd.delete(0,sqlCmd.length());
			updateDBDataList = new ArrayList();
			updateDBSqlList = new ArrayList();
			if(dbData.size() == 0){//WML01不存在時,Insert
				//if(input_method.equals("F")){//檔案上傳
				//   sqlCmd="INSERT INTO WML01 VALUES (" +String.valueOf(Integer.parseInt(m_year))+","+String.valueOf(Integer.parseInt(m_month))+",'"+bank_code+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
				//         + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
				//}else{
				//線上編輯
				sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");
						
				dataList = new ArrayList();//傳內的參數List
				dataList.add(String.valueOf(Integer.parseInt(m_year))); 
				dataList.add(String.valueOf(Integer.parseInt(m_month))); 
				dataList.add(bank_code); 
				dataList.add(report_no); 
				dataList.add(input_method);
				dataList.add(user_id);
				dataList.add(user_name);
				dataList.add(common_center);
				dataList.add(upd_method);
				dataList.add(upd_code);
				dataList.add(String.valueOf(batch_no));
				dataList.add(user_id);
				dataList.add(user_name);
				updateDBDataList.add(dataList);//1:傳內的參數List
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				//99.11.16sqlCmd="INSERT INTO WML01 VALUES (" +String.valueOf(Integer.parseInt(m_year))+","+String.valueOf(Integer.parseInt(m_month))+",'"+bank_code+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
			    //         + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
				//}
			}else{//WML01存在時,做update				
				sqlCmd.append(" INSERT INTO WML01_LOG "); 
				sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,");
				sqlCmd.append(" upd_method,upd_code,batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'");
				sqlCmd.append(" from WML01");
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				dataList = new ArrayList();//傳內的參數List
				dataList.add(user_id); 
				dataList.add(user_name); 
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add(bank_code);
				dataList.add(report_no);
				updateDBDataList.add(dataList);//1:傳內的參數List	
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				
			    sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
			    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
			    sqlCmd.append(" UPDATE WML01 SET "); 
				sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
						
				dataList = new ArrayList();//傳內的參數List
				dataList.add(upd_method); 
				dataList.add(upd_code); 
				dataList.add(String.valueOf(batch_no));
				dataList.add(user_id);
				dataList.add(user_name);
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add(bank_code);
				dataList.add(report_no);					
				updateDBDataList.add(dataList);//1:傳內的參數List
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				//99.11.16sqlCmd=" UPDATE WML01 SET " 
				//	  +" upd_method='"+upd_method+"',upd_code='"+upd_code+"',batch_no="+batch_no+",user_id='"+user_id+"',user_name='"+user_name+"',update_date=sysdate"   
				//	  +" where m_year=" + String.valueOf(Integer.parseInt(m_year)) + " AND m_month=" + String.valueOf(Integer.parseInt(m_month)) + " AND " 						
				//	  +" bank_code='" + bank_code + "' AND report_no='" + report_no + "'";
			}
			*/
			Utility.printLogTime("F01-5 time");	
			//傳送email至bank_cmml的e-mail信箱
			System.out.println("send email begin");
			/*
			if(input_method.equals("F")){//檔案上傳
				Utility.sendMailNotification(bank_code,report_no,
				m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id);
			}else{*/
			//線上編輯  		
				Utility.sendMailNotification(bank_code,report_no,
				m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");							
			//}		
			System.out.println("send email end");
			
			/*	
			if(input_method.equals("F")){//檔案上傳
			   //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			   updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
			}*/
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
				errMsg = errMsg + "UpdateF01.doParserReport_F01 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			Utility.printLogTime("F01-6 time");	
			errMsg = upd_code +":" + errMsg;//94.11.14
			System.out.println("errMsg="+errMsg);
			/*
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}*/	
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateF01.doParserReport_F01 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateF01.doParserReport_F01="+e.getMessage());
			System.out.println("UpdateF01.doParserReport_F01="+e.toString());
		}
		return errMsg;
		//return parserResult;
	}	
}
