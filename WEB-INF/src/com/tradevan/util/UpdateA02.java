//94.03.02 金額改用BigDecimal去轉 by 2295
//94.03.02 找出該bank_code的bank_type by 2295
//94.03.07 fix 有負號("-")時,把之前的"0"去掉 by 2295
//94.03.22 add 若被鎖定,則不執行檢核,直接send mail通知 by 2295
//94.03.22 fix 改善performance by 2295
//94.03.25 fix 若A02的值,都為"0"值時,顯示檢核為"0" by 2295
//95.03.20 add 共用中心傳送彙總檢核結果e-mail by 2295
//95.03.27 fix 共用中心傳送彙總e-mail格式 by 2295
//95.05.17 add 公式檢核/增加額外的科目代號 by 2295
//         add 95.04.20以後申報的資料才要做公式檢核 by 2295
//95.05.29 fix can=040010 是 95.4.20以後申報的.才要檢核
//             990220 <= 990210 其他的,直接檢核 by 2295
//96.08.13 若為檔案上傳批次寫入A04 ZeroData / 有修改的A04寫入A04_log by 2295
//96.08.14 寫入暫存table AXX_TMP/有異動的資料寫入A04.使用preparestatement by 2295
//         批次寫入WML01_Log/WML02_Log/WML03_Log/WML03 by 2295
//96.12.17 fix 批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯) by 2295
//96.12.21 刪除AXX_TMP暫存檔 by 2295 
//99.11.08 add InsertWML03/ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.10 add AXX_TMP增加amt1~amt9 by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//			   		  Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					  Utility.getWML01讀取WML01 all data
//				      Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//101.08.16 add 檔案上傳.990811信用部固定資產淨值，無超過其淨值,預設勾選 by 2295
//          fix 990811/990812/990813/990814為勾選項目,不算在金額皆為零 by 2295
//102.02.01 fix InsertZeroExtraA02_List add A02.amt_name by 2295
//102.04.25 add 增加990421/990621檔案上傳 by 2295
//104.02.13 add 增加990422/990622檔案上傳 by 2295
//108.01.03 add 增加檢核990720=990721*1/2+990722 by 2295
//110.04.09 fix InsertZeroExtraA02_List add A02.amt_name1/amt_name2 by 2295
//111.08.24 add 增加991311/991321檔案上傳 by 2295
package com.tradevan.util;

import com.tradevan.util.dao.DataObject;
import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.text.SimpleDateFormat;


public class UpdateA02 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A02(String report_no, String m_year,
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
		String  bank_code="";//機構代碼
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				
		List A02List = new LinkedList();//A02細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true";
		boolean nonZero=false;//94.03.25 檢查A02的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.08 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":""; 
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
	    List updateDBList = new ArrayList();//0:sql 1:data		
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
	    //=====================================================================    
		if(input_method.equals("F")){
		  System.out.println("tmpFile.lastModified()="+tmpFile.lastModified());
		}
		add_date=bkformat.format(filedate);
		
		try { 
			Utility.printLogTime("A02 begin time");		
			
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來						
				List allList = getA02FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table
				bank_codeList = (List)allList.get(0);
				A02List = (List)allList.get(1);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
			}else{//線上編輯,將bank_code加到bank_codeList				
				bank_codeList.add(srcbank_code);
				user_id = szuser_id;
				user_name = szuser_name;
			}
			Utility.printLogTime("A02-1 time");		
			//若為"Y"表示由共用中心代傳入
			sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
			paramList.add(srcbank_code);
			paramList.add(wlx01_m_year);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				
			if(dbData.size() != 0){//ba01有此筆bank_no的資料
			   if((((DataObject)dbData.get(0)).getValue("bank_type") != null) && ((String)((DataObject)dbData.get(0)).getValue("bank_type")).equals("8")){			   	 
			   	  common_center="Y";	
			   	  Utility.setSendMsg("");//95.03.20清空共用中心的mail msg
			   	  Utility.addMailNotification("",report_no,m_year,m_month, checkResult,"",emailformat.format(filedate),filename,user_id);//95.03.27 add 彙總title
			   }
			}
			
			Utility.printLogTime("A02-2 time");	
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
						
			//96.08.13 若為檔案上傳批次寫入A02 ZeroData / 有修改的A02寫入A02_log
			if(input_method.equals("F")){//在A02中無資料的先將insert Zero data 到A02
		    	errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A02' and acc_div='04' ");//在A02中無資料的先將insert Zero data 到A02
		    	InsertZeroExtraA02_List(m_year,m_month,bank_codeList);
		    	Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A02_LOG		    	
		    	errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤		    	
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A02		    	
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			//批次寫入WML01_LOG(檢核結果)(區分檔案上傳/線上編輯)
			Utility.InsertWML01_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			
			
			List lockBank_CodeList = Utility.getLockBank_Code(m_year,m_month,report_no);
			bank_codeLoop:
			for(int i=0;i<bank_codeList.size();i++){
				//94.03.22 add 若被鎖定,則不執行檢核,直接send mail通知==================================
				if(lockBank_CodeList != null){
				   for(int lockidx=0;lockidx<lockBank_CodeList.size();lockidx++){					
				       if(((String)bank_codeList.get(i)).equals((String)lockBank_CodeList.get(lockidx))){
				           if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容
				               Utility.addMailNotification((String)bank_codeList.get(i),report_no,
										m_year,m_month, checkResult,"C",emailformat.format(filedate),filename,user_id);
				           }else{    
					         Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
								     m_year,m_month, checkResult,"C",emailformat.format(filedate),filename,user_id,"");
				           }
				           continue bank_codeLoop;	
				       }
				   }
				}
				//=================================================================================
				errCount = 0;
				System.out.println("bank_codeList["+i+"]="+(String)bank_codeList.get(i));
				//94.03.02 找出該bank_code的bank_type by 2295=============================================
				
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
				paramList = new ArrayList();
				paramList.add((String)bank_codeList.get(i));
				paramList.add(wlx01_m_year);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				
				if(dbData != null && dbData.size() != 0){
					bank_type = (String)((DataObject)dbData.get(0)).getValue("bank_type");
					System.out.println("bank_type="+bank_type);
				}
				//=================================================================================
				Utility.printLogTime("A02-3 time");	
				
				//clear WML03(檢核其他錯誤)
				/*96.08.13移至InsertWML03(科目代號不存在)*/
				Utility.printLogTime("A02-4 time");	
				//clear WML02(檢核公式)
				dbData = Utility.getWML03_count(m_year,m_month,(String)bank_codeList.get(i),report_no);//99.11.19 fix
				/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select count(*) as countdata from WML03 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				paramList = new ArrayList();
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(i));
				paramList.add(report_no);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");	
			    */
				if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) == 0){//96.08.13檢查WML03有無其他錯誤,無其他錯誤時,才檢核公式
				   System.out.println("bank_code="+(String)bank_codeList.get(i)+".WML03.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				
				   //if(errCount == 0){//無其他錯誤時,才檢核公式
				   /*96.08.07 移至InsertWML02_LOG*/
				    Utility.printLogTime("A02-5 time");
				    
				    nonZero = false;
				    //101.08.16 fix 990811/990812/990813/990814為勾選項目,不算在金額皆為零 by 2295
				    dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"amt > 0 AND acc_code not in ('990811','990812','990813','990814')");//99.11.19 fix
				    /*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數					
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
					
				    if(input_method.equals("W")){
				    	sqlCmd.append("select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? and amt > 0");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));
						
						//99.11.08dbData = DBManager.QueryDB("select count(*) as countzero from A02 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
						// 	    			       "bank_code='" + (String)bank_codeList.get(i) + "' and amt > 0","countzero");
					}else{
						sqlCmd.append("select count(*) as countzero from axx_tmp where m_year=? AND m_month=? AND bank_code=? and report_no=? and amt > 0 and filename=?");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));
						paramList.add(report_no);
						paramList.add(filename);		
						//99.11.08dbData = DBManager.QueryDB("select count(*) as countzero from axx_tmp where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
						//						   "bank_code='" + (String)bank_codeList.get(i) + "' and report_no='"+report_no+"' and amt > 0 and filename='"+filename+"'","countzero");				    	
					}	
					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
					*/
					if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					       nonZero = true;
					   }
					}   
					System.out.println("nonZero="+nonZero);					
					Utility.printLogTime("A02-6 time");
					
				    //95.05.17 fix 若A02的值,只要有非"0"值時,才執行檢核				    
				    if(nonZero){
					   //執行公式檢核				       		       	
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       if(ruleList.size() > 0){//99.11.08有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
					       
					       //99.09.24for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
					    	   //updateDBSqlList.add((String)ruleList.get(ruleidx));
					    	   //errCount ++;
					       //}
					   }				       
				       Utility.printLogTime("A02-7 time");			
				       //108.01.03 add 檢核990720 = 990721*1/2+990722==================================================
                       ruleList = new LinkedList();
                       ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
                       if(ruleList.size() > 0){//有檢核失敗時,才加入
                           updateDBList.add((List)ruleList.get(0));
                           errCount += ((List)((List)ruleList.get(0)).get(1)).size();//累計參數List的size
                       }
                       Utility.printLogTime("A02-8 time");  
				    }				   
				}else{//enf of errCount==0
				   errCount = Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString());	
				}//enf of errCount==0
				
				dbData =  Utility.getWML01(m_year,m_month,(String)bank_codeList.get(i),report_no);
				/*99.11.19移至Utility.getWML01讀取WML01 all data
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select * from WML01 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			 	paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(i));
				paramList.add(report_no);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				
				//99.11.08dbData = DBManager.QueryDB("select * from WML01 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +						
				//		 "bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'","m_year,m_month,add_date,batch_no,update_date");
				*/
				upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
				checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
				//94.03.28 fix Z檢核為0====================================
				if(!nonZero && errCount == 0){
					upd_code = "Z";//Z檢核為0
					checkResult="Z";//Z檢核為0
				}
				//========================================================
				List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,(String)bank_codeList.get(i),report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);//99.11.19 fix
				for(int j=0;j<getUpdateDBList.size();j++){
				    updateDBList.add(getUpdateDBList.get(j));
				}
				/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update
				if(dbData.size() == 0){//WML01不存在時,Insert
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					if(input_method.equals("F")){//檔案上傳
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
						
					   //99.11.08sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					  sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");						
					  //99.11.08sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
				      //99.11.08       + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
					dataList = new ArrayList();//傳內的參數List
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add((String)bank_codeList.get(i)); 
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
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}else{//WML01存在時,做update
					//96.08.13移至批次寫入InsertWML01_LOG
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				
					//99.11.08sqlCmd=" UPDATE WML01 SET " 
					//	  +" upd_method='"+upd_method+"',upd_code='"+upd_code+"',batch_no="+batch_no+",user_id='"+user_id+"',user_name='"+user_name+"',update_date=sysdate"   
					//	  +" where m_year=" + m_year + " AND m_month=" + m_month + " AND " 						
					//	  +" bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
					dataList = new ArrayList();//傳內的參數List
					dataList.add(upd_method); 
					dataList.add(upd_code); 
					dataList.add(String.valueOf(batch_no));
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(m_year);
					dataList.add(m_month);
					dataList.add((String)bank_codeList.get(i));
					dataList.add(report_no);					
					updateDBDataList.add(dataList);//1:傳內的參數List					
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
				*/
				Utility.printLogTime("A02-9 time");
				
				//傳送email至bank_cmml的e-mail信箱
			    System.out.println("send email begin");
				if(input_method.equals("F")){//檔案上傳
				    if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容				           
				       Utility.addMailNotification((String)bank_codeList.get(i),report_no,
					           m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id);
				    }else{
				        Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
								m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,"");
				    }
				}else{//線上編輯  		
				    if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容
					   Utility.addMailNotification((String)bank_codeList.get(i),report_no,
					           m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id);
				    }else{
				        Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
								m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");
				    }
				}		
				System.out.println("send email end");
			}//end of for -->bank_codeList
			Utility.printLogTime("A02-10 time");
				
			if(input_method.equals("F")){//檔案上傳
			    //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.08updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19				
			    //96.12.21 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			}
			//99.11.08if(updateDBSqlList != null && updateDBSqlList.size()!=0) updateOK=DBManager.updateDB(updateDBSqlList);
			
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA02.doParserReport_A02 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
			if(common_center.equals("Y")){//95.03.20由共用中心傳入的,傳送彙總mail	
			   if(input_method.equals("F")){//檔案上傳			   
			        Utility.sendMailNotification(srcbank_code,report_no,
							m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,Utility.getSendMsg());
			    }else{//線上編輯
			        Utility.sendMailNotification(srcbank_code,report_no,
							m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,Utility.getSendMsg());
			    }
			}
			Utility.printLogTime("A02-11 time");			
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA02.doParserReport_A02 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA02.doParserReport_A02="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	// 96.08.13 寫入AXX_TMP暫存table,使用preparestatement
    // 99.11.08 add 套用DAO.preparestatment,並列印轉換後的SQL 
	// 99.11.10 add amt1~amt9 by 2295
	//101.08.16 add 990811信用部固定資產淨值，無超過其淨值,預設勾選 by 2295
	//102.04.25 add 990421/990621.農金局文號檔案上傳 by 2295
	//111.08.24 add 991311/991321.農金局文號檔案上傳 by 2295
	private static List getA02FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			//List A02List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAmt="";
			String amt_name = "";
			//96.08.13 寫入AXX_TMP暫存table,使用preparestatement
			List updateDBList = new LinkedList();//0:sql 1:data		
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new LinkedList();//儲存參數的data
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				sqlCmd.append("select count(*) as countdata from AXX_TMP where m_year=? and m_month=? and filename=?");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(filename);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");
				if(dbData != null && (Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) != 0)){
					System.out.println("AXX_TMP.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
					//List deleteList = new LinkedList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete AXX_TMP where m_year=? and m_month=? and filename=?");
					updateDBDataList.add(paramList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					//deleteList.add("delete AXX_TMP where m_year="+m_year+" and m_month="+m_month+" and filename='"+filename+"'");
					//System.out.println("delete OK??"+DBManager.updateDB(deleteList));
				}
				
				//96.08.13 寫入AXX_TMP暫存table
				sqlCmd.delete(0, sqlCmd.length());
				dataList =  new ArrayList();//儲存參數的data
				updateDBDataList = new ArrayList();//儲存參數的List	
				updateDBSqlList = new ArrayList();	
				
				sqlCmd.append("INSERT INTO AXX_TMP  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
				
				doLoop://將txt檔案儲存至LinkedList
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
							detail=new LinkedList();
							detail.add(m_year);
							detail.add(m_month);
							if(!bank_codeList.contains(txtline.substring(5,12))){
								bank_codeList.add(txtline.substring(5,12));
								/* 102.02.01 990811未上線.暫時拿掉
								detail.add(txtline.substring(5,12));//bank_code
	                            detail.add("990811");//科目代號
	                            detail.add("1");//990811信用部固定資產淨值，無超過其淨值,預設勾選 
	                            detail.add("0");//99.11.11 add amt1
	                            detail.add("0");//99.11.11 add amt2
	                            detail.add("0");//99.11.11 add amt3
	                            detail.add("0");//99.11.11 add amt4
	                            detail.add("0");//99.11.11 add amt5
	                            detail.add("0");//99.11.11 add amt6
	                            detail.add("0");//99.11.11 add amt7
	                            detail.add("0");//99.11.11 add amt8
	                            detail.add("0");//99.11.11 add amt9         
	                            detail.add(report_no);//96.08.13
	                            detail.add(filename);//96.08.13
	                            //System.out.println("detail="+detail);
	                            //A02List.add(detail);
	                            updateDBDataList.add(detail);//1:傳內的參數List
	                            detail=new LinkedList();
	                            detail.add(m_year);
	                            detail.add(m_month);
	                            */
							}
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號
							if(txtline.substring(12,18).equals("990421") || txtline.substring(12,18).equals("990621") || txtline.substring(12,18).equals("990422") || txtline.substring(12,18).equals("990622") || txtline.substring(12,18).equals("991311")  || txtline.substring(12,18).equals("991321")){							   
	                           amt_name = txtline.substring(18,txtline.length());
	                           if(!amt_name.equals("")){
	                               tmpAmt = "1";
	                           }else{
	                               tmpAmt = "0";
	                           }
							}else{
							    //94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
							    tmpAmt = txtline.substring(18,32);//金額
							    if(tmpAmt.indexOf("-") != -1){
							        tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							    }
							    amt_name = "";
							}
							//=========================================							
							detail.add(tmpAmt);//金額
							detail.add("0");//99.11.11 add amt1
							detail.add("0");//99.11.11 add amt2
							detail.add("0");//99.11.11 add amt3
							detail.add("0");//99.11.11 add amt4
							detail.add("0");//99.11.11 add amt5
							detail.add("0");//99.11.11 add amt6
							detail.add("0");//99.11.11 add amt7
							detail.add("0");//99.11.11 add amt8
							detail.add("0");//99.11.11 add amt9		
							detail.add(amt_name);//102.04.25 add amt_name
							detail.add(report_no);//96.08.13
							detail.add(filename);//96.08.13
							//System.out.println("detail="+detail);
							//A02List.add(detail);
							updateDBDataList.add(detail);//1:傳內的參數List
						}
				}	
				in.close();
				f.close();
				//96.08.13 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
			}catch(Exception e){
				errMsg = errMsg + "UpdateA02.getA02FileData Error:"+e.getMessage()+"<br>";
			}
			allList.add(bank_codeList);
			allList.add(updateDBDataList);
			return allList;
	}
	
	//95.05.17 add 公式檢核	
	//99.11.08 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List CheckRule(String bank_type,String report_no,String m_year,String m_month,
								 String bank_code,String user_id,String user_name){
		String  amt="";		
		double	amt_L = 0.0, amt_R = 0.0,amt_tbl = 0.0;
		String	quop = null;
		String  cano="";
		String noop="";
			
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.08 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		
		try{
			//check ruleno1, ruleno2
			sqlCmd.append(" select r1.cano,r1.quop,r2.left_flag,"+report_no+".amt,r2.noop,"+report_no+".acc_code ");
			sqlCmd.append(" from "+report_no+",ruleno1 r1,ruleno2 r2 ");
			sqlCmd.append(" where "+report_no+".acc_code=r2.acc_code ");				   
		    sqlCmd.append(" and r1.CANO = r2.cano");
			sqlCmd.append(" and r1.acc_type = r2.acc_type");
			sqlCmd.append(" and r1.acc_type = ?");
		    sqlCmd.append(" and "+report_no+".m_year=?"); 
			sqlCmd.append(" and "+report_no+".m_month=?");
			sqlCmd.append(" and "+report_no+".bank_code=?");
			sqlCmd.append(" order by r1.cano,r2.left_flag,r2.nserial");
			
			paramList.add(bank_type);
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
			
			cano=(String)((DataObject)dbData.get(0)).getValue("cano");
			quop = ((String)((DataObject)dbData.get(0)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());
	
			System.out.println("check rule begin");
			System.out.println("quop:"+quop);
			amt_L = 0.0; amt_R = 0.0;
			amt_tbl=0.0;
			sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append("INSERT INTO WML02 VALUES (?,?,?,?,?,?,?,?,?,sysdate)");
			for(int r=0;r<dbData.size();r++){
			    //System.out.println("cano:"+cano);			   
				//System.out.println("r="+r);
				if(!(((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals(cano)){//與前一個公式不同時				    
				    //if ((quop.equals("=") && (amt_L == amt_R))){				   
				    if ((quop.equals("=") &&  ( Math.abs(amt_L - amt_R) <= 10 ))){//95.02.07 add -10~10均算檢核成功       
				         //System.out.println("amt_L="+amt_L);
					     //System.out.println("amt_R="+amt_R);
					     //System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));					
					} else if ((quop.equals(">") && (amt_L > amt_R))) {
					} else if ((quop.equals("<") && (amt_L < amt_R))) {
					} else if ((quop.equals(">=") && (amt_L >= amt_R))) {
					} else if ((quop.equals("<=") && (amt_L <= amt_R))) {
					} else if ((quop.equals("!=") && (amt_L != amt_R))) {
					}else {
					    //System.out.println(sqlCmd);
					    if(cano.equals("040010")){//95.05.29 fix can=040010 是 95.4.20以後申報的.才要檢核
					       if((Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month)) >= 9504 ){
						        System.out.println("WML02 have error:cano="+cano);
						        dataList = new ArrayList();//傳內的參數List		 	           				   
								dataList.add(m_year); 
								dataList.add(m_month); 
								dataList.add(bank_code); 
								dataList.add(report_no); 
								dataList.add(cano); 
								dataList.add(String.valueOf(amt_L)); 
								dataList.add(String.valueOf(amt_R)); 
								dataList.add(user_id); 
								dataList.add(user_name); 
								updateDBDataList.add(dataList);//1:傳內的參數List
						        //sqlCmd = "INSERT INTO WML02 VALUES (" + m_year + "," + m_month + ",'" + bank_code + "','" +
						        //99.11.08        report_no+"','"+cano + "'," + amt_L + "," + amt_R + ",'"+user_id+"','"+user_name+"',sysdate)";
						        //updateDBSqlList.add(sqlCmd); 
					       }    
					    }else{
					    	System.out.println("WML02 have error:cano="+cano);
					        dataList = new ArrayList();//傳內的參數List		 	           				   
							dataList.add(m_year); 
							dataList.add(m_month); 
							dataList.add(bank_code); 
							dataList.add(report_no); 
							dataList.add(cano); 
							dataList.add(String.valueOf(amt_L)); 
							dataList.add(String.valueOf(amt_R)); 
							dataList.add(user_id); 
							dataList.add(user_name); 
							updateDBDataList.add(dataList);//1:傳內的參數List
					    
						    //99.11.08sqlCmd = "INSERT INTO WML02 VALUES (" + m_year + "," + m_month + ",'" + bank_code + "','" +
						    //       report_no+"','"+cano + "'," + amt_L + "," + amt_R + ",'"+user_id+"','"+user_name+"',sysdate)";					        
						    //updateDBSqlList.add(sqlCmd);    
					    }
					}
					//把這次的公式編號及quop儲存;並把金額清空
					cano=(String)((DataObject)dbData.get(r)).getValue("cano");
					quop = ((String)((DataObject)dbData.get(r)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("quop")).trim());					
					amt_L = 0.0; amt_R = 0.0;
					amt_tbl=0.0;
				}	
				//System.out.println("cano:"+cano);
				//System.out.println("quop:"+quop);	
				amt_tbl = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString());
				//System.out.println("acc_code="+(String)((DataObject)dbData.get(r)).getValue("acc_code"));
				//System.out.println("noop="+noop);
				//System.out.println("amt_tbl="+amt_tbl);
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("0")){//左式					
					if ((noop == null) || (noop.equals(""))) {
						amt_L = amt_tbl;
					}else if (noop.equals("+")) {
						amt_L += amt_tbl;
					}else if (noop.equals("-")) {
						amt_L -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_L *= amt_tbl;
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_L /= amt_tbl;
					}//end of noop
				}//end of left_flag=0-->左式金額加總	
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("1")){//右式
					if ((noop == null) || (noop.equals(""))) {
						amt_R = amt_tbl;
					}else if (noop.equals("+")) {
						amt_R += amt_tbl;
					}else if (noop.equals("-")) {
						amt_R -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_R *= amt_tbl;
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_R /= amt_tbl;
					}//end of noop
				}//end of left_flag=0-->右式金額加總
				//94.03.04 fix 把最後一筆的檢核失敗寫入db=================================
				if(r==dbData.size()-1){
				   if ((quop.equals("=") && (amt_L == amt_R))){
				   //if ((quop.equals("=") &&  ( Math.abs(amt_L - amt_R) <= 10 ))){//95.02.07 add -10~10均算檢核成功
				       //System.out.println("amt_L="+amt_L);
				       //System.out.println("amt_R="+amt_R);
				       //System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));
				   } else if ((quop.equals(">") && (amt_L > amt_R))) {
				   } else if ((quop.equals("<") && (amt_L < amt_R))) {
				   } else if ((quop.equals(">=") && (amt_L >= amt_R))) {
				   } else if ((quop.equals("<=") && (amt_L <= amt_R))) {
				   } else if ((quop.equals("!=") && (amt_L != amt_R))) {
				   }else {
					   //System.out.println(sqlCmd);
				       if(cano.equals("040010")){//95.05.29 fix can=040010 是 95.4.20以後申報的.才要檢核
					       if((Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month)) >= 9504 ){
					            System.out.println("WML02 have error:cano="+cano);
						        dataList = new ArrayList();//傳內的參數List		 	           				   
								dataList.add(m_year); 
								dataList.add(m_month); 
								dataList.add(bank_code); 
								dataList.add(report_no); 
								dataList.add(cano); 
								dataList.add(String.valueOf(amt_L)); 
								dataList.add(String.valueOf(amt_R)); 
								dataList.add(user_id); 
								dataList.add(user_name); 
								updateDBDataList.add(dataList);//1:傳內的參數List
						 
						       //99.11.08sqlCmd = "INSERT INTO WML02 VALUES (" + m_year + "," + m_month + ",'" + bank_code + "','" +
						       //         report_no+"','"+cano + "'," + amt_L + "," + amt_R + ",'"+user_id+"','"+user_name+"',sysdate)";
						       //updateDBSqlList.add(sqlCmd);
					       }    
					   }else{
					        System.out.println("WML02 have error:cano="+cano);
					        dataList = new ArrayList();//傳內的參數List		 	           				   
							dataList.add(m_year); 
							dataList.add(m_month); 
							dataList.add(bank_code); 
							dataList.add(report_no); 
							dataList.add(cano); 
							dataList.add(String.valueOf(amt_L)); 
							dataList.add(String.valueOf(amt_R)); 
							dataList.add(user_id); 
							dataList.add(user_name); 
							updateDBDataList.add(dataList);//1:傳內的參數List
						    //99.11.08sqlCmd = "INSERT INTO WML02 VALUES (" + m_year + "," + m_month + ",'" + bank_code + "','" +
						    //       report_no+"','"+cano + "'," + amt_L + "," + amt_R + ",'"+user_id+"','"+user_name+"',sysdate)";
						    //updateDBSqlList.add(sqlCmd);
					   }
				   }
				}
				//==========================================================
				noop = (((DataObject)dbData.get(r)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("noop")).trim());
				//System.out.println("r="+r+"end");				
			} //end of for()	
			System.out.println("WML02 have error.updateDBDataList.size()="+updateDBDataList.size());
			if(updateDBDataList.size() > 0){//99.11.08當有檢核失敗時,才加入
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			System.out.println("check rule end");
		}catch(Exception e){
			errMsg = errMsg + "UpdateA02.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA02.CheckRule Error:"+e.getMessage());
		}
		return updateDBList;
	}
		
	//96.08.13 ad 批次增加額外科目代號
    //99.11.08 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//102.02.01 add a02.amt_name 
	private static boolean InsertZeroExtraA02_List(String m_year,String m_month,List bank_code){
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.08 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
		boolean updateOK=false;
		List A02updateDBSqlList = new LinkedList();
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertZeroExtraA02_List begin time="+logformat.format(nowlog));	
		try{
			
			/*若無資料在A02時,則新增zero至A02
			INSERT INTO A02
			select 96,3,a.bank_no,acc_code ,0
			from ncacno,(select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			            (select bank_code from A05 where m_year=94 AND m_month=6 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A05已有資料的
			where acc_tr_type = 'A02' and acc_div='04' and acc_code in ('990511','990512','990611')
			and a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A05的bank_no
			order by a.bank_no,acc_range"
			*/
			//農會
			sqlCmd.append(" INSERT INTO A02"); 
			sqlCmd.append(" select "+m_year+","+m_month+",a.bank_no,acc_code ,0,'','','' "); 
			sqlCmd.append(" from ncacno,(select bank_no from bn01 where m_year= ?  and bank_no in ("+Utility.getCombinString(bank_code,",")+"))a");//目前需更新的bank_code 
		    sqlCmd.append("            ,(select bank_code from A05");
			sqlCmd.append("              where m_year=? AND m_month=?");
		    sqlCmd.append("              AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A05已有資料的 
			sqlCmd.append(" where acc_tr_type = 'A02' and acc_div='04' and acc_code in ('990511','990512','990611')"); 
		    sqlCmd.append(" and a.bank_no = b.bank_code(+)");
			sqlCmd.append(" and b.bank_code is null ");//無資料在A05的bank_no
			sqlCmd.append(" order by a.bank_no,acc_range");		
			dataList = new ArrayList();//傳內的參數List	
			dataList.add(wlx01_m_year);
			dataList.add(m_year);
			dataList.add(m_month);
			updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			
			//漁會			
			sqlCmd.delete(0, sqlCmd.length());
			dataList =  new ArrayList();//儲存參數的data
			updateDBDataList = new ArrayList();//儲存參數的List	
			updateDBSqlList = new ArrayList();							
			
			sqlCmd.append(" INSERT INTO A02");
			sqlCmd.append(" select "+m_year+","+m_month+",a.bank_no,acc_code ,0,'','','' "); 
			sqlCmd.append(" from ncacno_7,(select bank_no from bn01 where m_year = ? and bank_no in ("+Utility.getCombinString(bank_code,",")+"))a");//目前需更新的bank_code 
			sqlCmd.append("            ,(select bank_code from A05");
			sqlCmd.append("              where m_year=? AND m_month=?");
			sqlCmd.append("              AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A05已有資料的 
			sqlCmd.append(" where acc_tr_type = 'A02' and acc_div='04' and acc_code in ('990511','990512','990611')");
			sqlCmd.append(" and a.bank_no = b.bank_code(+)");
			sqlCmd.append(" and b.bank_code is null ");//無資料在A05的bank_no
			sqlCmd.append(" order by a.bank_no,acc_range");	
			dataList.add(wlx01_m_year);
			dataList.add(m_year);
			dataList.add(m_month);
			updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
				
			System.out.println("InsertZeroExtraA02_List OK??"+updateOK);
			if(!updateOK){
				//errMsg = errMsg + "InsertZeroExtraA05_List Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA02.InsertZeroExtraA02_List Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA02.InsertZeroExtraA02_List Error:"+e.getMessage());
		}
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertZeroExtraA02_List end time="+logformat.format(nowlog));	
    	return updateOK;
	}
	
	//108.01.03 add 增加檢核990720=990721*1/2+990722
    private static List CheckOtherRule(int errCount,String bank_type,String report_no,String m_year,String m_month,
                                       String bank_code,String user_id,String user_name){
        BigDecimal tmp990720=null;
        BigDecimal tmp990720_sum=null;
        BigDecimal tmp990721=null;
        BigDecimal tmp990722=null;
        String  amt="";
        List dbData = null;
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList();
        //99.11.15 add 
        List updateDBList = new ArrayList();//0:sql 1:data      
        List updateDBSqlList = new ArrayList();
        List updateDBDataList = new ArrayList();//儲存參數的List
        List dataList =  new ArrayList();//儲存參數的data
        try{
             sqlCmd.append("select acc_code,amt from A02 where bank_code=? and m_year=? and m_month=?");
             paramList.add(bank_code);
             paramList.add(m_year);
             paramList.add(m_month);
             dbData=DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
             if(dbData != null && dbData.size() != 0){
                 for(int i=0;i<dbData.size();i++){
                     if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("990720")){//95.05.29 add
                         tmp990720=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));                        
                     }
                     if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("990721")){
                         tmp990721=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));                        
                     }
                     if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("990722")){                
                         tmp990722=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
                     }                     
                 }
             }
             System.out.println("CheckOtherRule.bank_code="+bank_code);
             System.out.println("990720="+tmp990720);
             System.out.println("990721="+tmp990721);
             System.out.println("990722="+tmp990722);
             System.out.println("errCount="+errCount);
             
             
             //990720=990721*1/2+990722->取到整數,小數點下以無條件捨去
             tmp990720_sum = new BigDecimal(tmp990721.multiply(new BigDecimal("0.5")).intValue());
             System.out.println("990721*1/2="+tmp990720_sum);
             tmp990720_sum = tmp990720_sum.add(tmp990722);
             System.out.println("990721*1/2+990722="+tmp990720_sum);
            
             if(tmp990720.compareTo(tmp990720_sum) != 0){//990720 !=990721*1/2+990722
                 sqlCmd.delete(0,sqlCmd.length());
                 sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
                 dataList.add(m_year); 
                 dataList.add(m_month); 
                 dataList.add(bank_code);   
                 dataList.add(report_no);  
                 dataList.add(String.valueOf(errCount));  
                 dataList.add("檢核發現(990720)["+Utility.setCommaFormat(tmp990720.toString())+"元]不等於(990721*1/2+990722)["+Utility.setCommaFormat(tmp990720_sum.toString())+"元]");
                 dataList.add(user_id);  
                 dataList.add(user_name);
                 updateDBDataList.add(dataList);//1:傳內的參數List
                 errCount++ ;                                           
             }
             
              
        }catch(Exception e){
            errMsg = errMsg + "UpdateA02.CheckOtherRule Error:"+e.getMessage()+"<br>";
            System.out.println("UpdateA02.CheckOtherRule Error:"+e.getMessage());
        }
        if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
           updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql   
           updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
           updateDBList.add(updateDBSqlList);
        }
        return updateDBList;
    }
    
}
