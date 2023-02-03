// 95.05.25 create by 2295
// 95.05.29 若為檔案上傳時,將wlx01.credit_staff_num寫回992300 by 2295
// 96.08.15 若為檔案上傳批次寫入A03 ZeroData / 有修改的A03寫入A03_log by 2295
// 96.08.15 寫入暫存table AXX_TMP/有異動的資料寫入A99.使用preparestatement by 2295
//          批次寫入WML01_Log/WML02_Log/WML03_Log/WML03 by 2295
// 96.12.17 fix 批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯) by 2295
// 99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
// 99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					   Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					   Utility.getWML01讀取WML01 all data
//				       Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					   Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//100.06.27 fix 檢核992110農(漁)會全體淨值不可為空值或為0 by 2295
//100.08.04 fix 檢核992100農(漁)會全體淨值不可為空值或為0或負數 by 2295
//100.09.07 fix 992100農(漁)會全體資產總額用BigDecimal取值 by 2295
//101.06.29 add 檔案上傳,當992300為0時,才將wlx01.credit_staff_num回寫 by 2295
//101.06.29 add 增加992300信用部員工人數與原總機構基本資料維護之從事信用業務員工人數 by 2295
//101.07.09 add 回寫總機構人數至A99時,先commit by 2295
//101.10.12 add 增加檢核.從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數 = 本機構員工總人數 + 分支機構員工總人數 by 2295
//110.08.16 add 增加檢核  992450=Max(992452,992454,992456) / 992460=Max(992462,992464,992466) by 2295[111.05.09取消]
//110.08.09 add 增加上傳amt_name[111.06.14取消]
//110.10.26 add 增加檢核 992452>992454>992456 / 992462>992464>992466 by 2295[111.05.09取消]
//110.10.28 add 增加檢核 992440>(992452+992454+992456)+(992462+992464+992466) by 2295[111.05.09取消]
//111.07.15 add 增加公式檢核 992710=992711+992712+992713 by 2295

package com.tradevan.util;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;

public class UpdateA99 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A99(String report_no, String m_year,
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
		
		List A99List = new LinkedList();//A99細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 		
		boolean nonZero=false;//94.03.25 檢查A99的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.15 add 查詢年度100年以前.縣市別不同===============================
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
			Utility.printLogTime("A99 begin time");
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA99FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table									
				bank_codeList = (List)allList.get(0);
				A99List = (List)allList.get(1);
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
			Utility.printLogTime("A99-1 time");
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
			
			Utility.printLogTime("A99-2 time");
			
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
			
			
   		    //96.08.15 若為檔案上傳批次寫入A99 ZeroData / 有修改的A99寫入A99_log
			if(input_method.equals("F")){//在A99中無資料的先將insert Zero data 到A99		    	
				errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A99' and acc_div='99' ");//在A99中無資料的先將insert Zero data 到A99
				Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A99_LOG		    	
		    	errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤		    	
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A99		    	
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
				
				Utility.printLogTime("A99-3 time");
				System.out.println("bank_codeList["+i+"]="+(String)bank_codeList.get(i));
				
				errCount = 0;
				//clear WML03(檢核其他錯誤)	
				/*96.08.15 移至Utlity.InsertWML03_LOG*/
				//(檢核其他錯誤)
				/*96.08.15移至InsertWML03*/
				Utility.printLogTime("A99-4 time");	
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
		    	if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) == 0){//96.08.15檢查WML03有無其他錯誤,無其他錯誤時,才檢核公式
				   System.out.println("bank_code="+(String)bank_codeList.get(i)+".WML03.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				   /*96.08.15移至InsertWML02_LOG
				    * 96.08.14移至InsertZeroA99_List*/
				  
					Utility.printLogTime("A99-5 time");
					
					nonZero = false;
					dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"amt > 0");//99.11.19 fix
					/*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數				
						
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
					
					//99.11.15檢核內容是否都為0
					sqlCmd.append(" select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? and amt > 0");                          
					paramList.add(m_year);
					paramList.add(m_month);
					paramList.add((String)bank_codeList.get(i));					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
					*/
					if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					       nonZero = true;
					   }
					}
					System.out.println("nonZero="+nonZero);
					
					Utility.printLogTime("A99-6 time");	
				    // 95.05.29將wlx01.credit_staff_num寫回992300=========================================================\
					//101.06.29 add 檔案上傳,當992300為0時,才將wlx01.credit_staff_num回寫 by 2295
					//101.07.09 add回寫總機構人數至A99時,先commit by 2295
				    ReWrite992300(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);
				    /*
				    ruleList = ReWrite992300(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);
				    if(ruleList.size() > 0){
					   updateDBList.add((List)ruleList.get(0));					   
					}   
				    */
				    if(nonZero){
					   //執行公式檢核
				       ruleList = new LinkedList();
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       if(ruleList.size() > 0){//99.11.08有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size					      
					   }			
				       
				       //檢核原總機構基本資料維護中從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數為0
				       //101.06.29 add 增加992300信用部員工人數與原總機構基本資料維護之從事信用業務員工人數 by 2295
				       //101.10.12 add 增加檢核.從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數 = 本機構員工總人數 + 分支機構員工總人數 by 2295
				       ruleList = new LinkedList();
				       ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);				    
					   if(ruleList.size() > 0){//99.11.15有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
					   }   
					   //99.11.15ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);
					   //for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
					   //   updateDBSqlList.add((String)ruleList.get(ruleidx));
					   //	   errCount ++;
					   //}
					   Utility.printLogTime("A99-7 time");
				    }
				    ruleList = new LinkedList();
				    //100.06.27 fix 檢核992110農(漁)會全體淨值不可為空值或為0
				    //100.08.04 fix 檢核992100農(漁)會全體淨值不可為空值或為0或負數
				    ruleList = CheckOtherRule1(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);				    
					if(ruleList.size() > 0){
					   updateDBList.add((List)ruleList.get(0));
					   errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
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
		    	
				//99.11.15dbData = DBManager.QueryDB("select * from WML01 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +						
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
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				if(dbData.size() == 0){//WML01不存在時,Insert					
					if(input_method.equals("F")){//檔案上傳
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
						
					   //99.11.15sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");
					   //99.11.15sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
				       //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
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
					//96.08.15移至批次寫入InsertWML01_LOG
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
					
					//99.11.15sqlCmd=" UPDATE WML01 SET " 
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
			
			Utility.printLogTime("A99-8 time");	
			if(input_method.equals("F")){//檔案上傳
			    //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.15updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19
			    //99.11.15 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			}
			
			for(int j=0;j<updateDBList.size();j++){
	    		System.out.println((String)((List)updateDBList.get(j)).get(0));
	    		System.out.println((List)((List)updateDBList.get(j)).get(1));
	    	}
		
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}			
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA99.doParserReport_A99 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
			Utility.printLogTime("A99-9 time");	
			if(common_center.equals("Y")){//95.03.20由共用中心傳入的,傳送彙總mail	
			   if(input_method.equals("F")){//檔案上傳			   
			        Utility.sendMailNotification(srcbank_code,report_no,
							m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,Utility.getSendMsg());
			    }else{//線上編輯
			        Utility.sendMailNotification(srcbank_code,report_no,
							m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,Utility.getSendMsg());
			    }
			}		
			
			Utility.printLogTime("A99-10 time");
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA99.doParserReport_A99 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA99.doParserReport_A99="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	// 99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL
	//110.08.09 add 增加上傳amt_name[111.06.14取消] 
	private static List getA99FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List A99List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAmt="";
			//String amt_name = "";//110.08.09 add
			//96.08.15 寫入AXX_TMP暫存table,使用preparestatement
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
			List updateDBList = new LinkedList();//0:sql 1:data		
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new LinkedList();//儲存參數的data
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				
				
				//96.08.15 若AXX_TMP暫存table,有資料則先delete
				sqlCmd.append("select count(*) as countdata from AXX_TMP where m_year=? and m_month=? and filename=?");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(filename);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");				
				if(dbData != null && (Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) != 0)){
					System.out.println("AXX_TMP.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete AXX_TMP where m_year=? and m_month=? and filename=?");
					updateDBDataList.add(paramList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);					
				}				
							
				//99.11.15 寫入AXX_TMP暫存table
				sqlCmd.delete(0, sqlCmd.length());
				dataList =  new ArrayList();//儲存參數的data
				updateDBDataList = new ArrayList();//儲存參數的List	
				updateDBSqlList = new ArrayList();
				
				sqlCmd.append("INSERT INTO AXX_TMP  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?,?)");
			
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
								
							}
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號.acc_code
							/*111.06.14取消
							if(txtline.substring(12,18).equals("992451") || txtline.substring(12,18).equals("992453") || txtline.substring(12,18).equals("992455") 
							|| txtline.substring(12,18).equals("992461") || txtline.substring(12,18).equals("992463") || txtline.substring(12,18).equals("992465")){							   
		                       amt_name = txtline.substring(18,txtline.length());
		                       if(!amt_name.equals("")){
		                           tmpAmt = "1";
		                       }else{
		                           tmpAmt = "0";
		                       }
							}else{
							*/
														
								//94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
								tmpAmt = txtline.substring(18,32);//金額
								if(tmpAmt.indexOf("-") != -1){
									tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
								}
								//amt_name = "";
							//}
							//=========================================							
							detail.add(String.valueOf(tmpAmt));//金額axx_tmp.amt
							detail.add("0");//99.11.15 add amt1
							detail.add("0");//99.11.15 add amt2
							detail.add("0");//99.11.15 add amt3
							detail.add("0");//99.11.15 add amt4
							detail.add("0");//99.11.15 add amt5
							detail.add("0");//99.11.15 add amt6
							detail.add("0");//99.11.15 add amt7
							detail.add("0");//99.11.15 add amt8
							detail.add("0");//99.11.15 add amt9		
							//detail.add(amt_name);//110.08.09 add amt_name
							detail.add(report_no);//96.08.15
							detail.add(filename);//96.08.15
							updateDBDataList.add(detail);//1:傳內的參數List
							A99List.add(detail);
						}
				}	
				in.close();
				f.close();
				//99.11.15 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				
				updateDBList.add(updateDBSqlList);            	
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
				
			}catch(Exception e){
				errMsg = errMsg + "UpdateA99.getA99FileData Error:"+e.getMessage()+"<br>";
			}
			allList.add(bank_codeList);
			allList.add(A99List);
			return allList;
	}
	
	//95.05.17 add 公式檢核	
	//99.11.08 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//A99無公式檢核//111.07.15 增加公式檢核		
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
				errMsg = errMsg + "UpdateA99.CheckRule Error:"+e.getMessage()+"<br>";
				System.out.println("UpdateA99.CheckRule Error:"+e.getMessage());
			}
			return updateDBList;
	}
	
	// 99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL
	//101.06.29 add 增加992300信用部員工人數與原總機構基本資料維護之從事信用業務員工人數 by 2295
	//101.10.12 add 檢核.從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數 = 本機構員工總人數 + 分支機構員工總人數 by 2295
	private static List CheckOtherRule(int errCount,String bank_type,String report_no,String m_year,String m_month,
			   String bank_code,String user_id,String user_name,String input_method){
	    List dbData = null;	
	    List dbData_992300 = null;    
	    StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		int wlx01_credit_staff_num = 0;//(B)從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數CREDIT_STAFF_NUM
		int wlx01_staff_num = 0;//(A)本機構員工總人數STAFF_NUM 101.10.12 add
		int wlx02_staff_num = 0;//(C)分支機構員工總人數 101.10.12 add
		int wlx01_wlx02_staff_num = 0;//A+C
		int a99_992300 = 0;
	    try{
	    	sqlCmd.append("select credit_staff_num,staff_num from wlx01 where m_year=? and bank_no=?");
	    	paramList.add(wlx01_m_year);
	    	paramList.add(bank_code);	  
	    	dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"credit_staff_num,staff_num");
	        //99.11.15dbData=DBManager.QueryDB("select credit_staff_num from wlx01 where bank_no='"+bank_code+"'","credit_staff_num");
	    	sqlCmd.delete(0,sqlCmd.length());
            sqlCmd.append(" select amt from A99 where m_year=? and m_month=? and bank_code=?");
            sqlCmd.append(" and acc_code = '992300'");
            paramList = new ArrayList();
            paramList.add(m_year);
            paramList.add(m_month);     
            paramList.add(bank_code);     
            dbData_992300 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
            
            sqlCmd.delete(0,sqlCmd.length());
            paramList = new ArrayList();
           
            sqlCmd.append(" select sum(decode(ba01.bank_no,?,0,(wlx02.staff_num))) as wlx02staff_num");
            sqlCmd.append(" from (select * from ba01 where m_year=? )ba01,(select * from wlx02 where m_year=?)wlx02");
            sqlCmd.append(" where (ba01.pbank_no=? )");
            sqlCmd.append(" and (ba01.bank_no=wlx02.bank_no(+) and (wlx02.CANCEL_NO <> 'Y' OR wlx02.CANCEL_NO IS NULL)) ");//95.08.25 add 分支機構總人數只統計未被裁撤的分支機構          
            paramList.add(bank_code) ;                   
            paramList.add(wlx01_m_year);
            paramList.add(wlx01_m_year);
            paramList.add(bank_code) ;
            List dbData_wlx02 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"wlx02staff_num");            
                        
            
	        if(dbData != null && dbData.size() != 0){
	            System.out.println("errCount="+errCount);
	            if(((DataObject)dbData.get(0)).getValue("credit_staff_num") == null
                || ((((DataObject)dbData.get(0)).getValue("credit_staff_num")).toString()).equals("0")){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現原總機構基本資料維護中從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數為0");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List
	                //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
	                //m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
	                //errCount + ",'檢核發現原總機構基本資料維護中從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數為0','"+user_id+"','"+user_name+"',sysdate)";
	                errCount++ ;	
	                //updateDBSqlList.add(sqlCmd);	            
	            }else{
	                //(B)從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數CREDIT_STAFF_NUM
	                wlx01_credit_staff_num = Integer.parseInt( (((DataObject)dbData.get(0)).getValue("credit_staff_num")).toString() );
	                //(A)本機構員工總人數STAFF_NUM 101.10.12 add
	                wlx01_staff_num = Integer.parseInt( (((DataObject)dbData.get(0)).getValue("staff_num")).toString() );
	                
	            }
	        }
	        //(C)分支機構員工總人數 101.10.12 add       
	        if(((DataObject)dbData_wlx02.get(0)).getValue("wlx02staff_num") != null){	                   
	            wlx02_staff_num = Integer.parseInt( (((DataObject)dbData_wlx02.get(0)).getValue("wlx02staff_num")).toString() );               
	        }
	        
	        if(dbData_992300 != null && dbData_992300.size() != 0){
                System.out.println("errCount="+errCount);
                if(Integer.parseInt( ((((DataObject)dbData_992300.get(0)).getValue("amt")).toString()) ) > 0){
                    a99_992300 = Integer.parseInt( ((((DataObject)dbData_992300.get(0)).getValue("amt")).toString()) );
                }
	        }
	        System.out.println("(B)wlx01_credit_staff_num="+wlx01_credit_staff_num);
	        System.out.println("a99_992300="+a99_992300);
	        if(wlx01_credit_staff_num != a99_992300){//101.06.29 add 檢核992300與總機構不同時,顯示錯誤訊息
	            //System.out.println("992300信用部員工人數["+a99_992300+"],與原總機構基本資料維護之從事信用業務員工人數["+wlx01_credit_staff_num+"]不一致!!若需修正總機構基本資料維護之從事信用業務員工人數,請至總機構基本資料維護(FX001W)進行修正!!");
	            dataList = new ArrayList();//傳內的參數List                                    
                dataList.add(m_year); 
                dataList.add(m_month); 
                dataList.add(bank_code); 
                dataList.add(report_no); 
                dataList.add(String.valueOf(errCount)); 
                dataList.add("992300信用部員工人數["+a99_992300+"],與原總機構基本資料維護之從事信用業務員工人數["+wlx01_credit_staff_num+"]不一致!!若需修正總機構基本資料維護之從事信用業務員工人數,請至總機構基本資料維護(FX001W)進行修正!!");                   
                dataList.add(user_id); 
                dataList.add(user_name); 
                updateDBDataList.add(dataList);//1:傳內的參數List
                errCount++ ;    
	        }
	        //A+C
	        wlx01_wlx02_staff_num = wlx01_staff_num + wlx02_staff_num;
	        System.out.println("(A)wlx01_staff_num="+wlx01_staff_num);
	        System.out.println("(C)wlx02_staff_num="+wlx02_staff_num);
	        //101.10.12 add A+C != B顯示錯誤訊息 
	        if(wlx01_wlx02_staff_num != wlx01_credit_staff_num){//A+C != B
	            System.out.println("A+C != B");
	            dataList = new ArrayList();//傳內的參數List                                    
                dataList.add(m_year); 
                dataList.add(m_month); 
                dataList.add(bank_code); 
                dataList.add(report_no); 
                dataList.add(String.valueOf(errCount)); 
                dataList.add("原總機構基本資料維護(FX001W)之從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數["+wlx01_credit_staff_num+"],與本機構員工總人數["+wlx01_staff_num+"]與分支機構員工總人數["+wlx02_staff_num+"]加總["+wlx01_wlx02_staff_num+"]不合("+wlx01_credit_staff_num+" 不等於 "+wlx01_staff_num+"+"+wlx02_staff_num+ " ),若需修正分支機構員工總人數,請至國內營業分支構機基本資料維護(FX002W)進行資料修改");                   
                dataList.add(user_id); 
                dataList.add(user_name); 
                updateDBDataList.add(dataList);//1:傳內的參數List
                errCount++ ;    	            
	        }
	       
	        if(updateDBDataList.size() > 0){//101.06.29當有檢核失敗時,才加入
	            sqlCmd.delete(0,sqlCmd.length());
	            sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");             
	            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql   
	            updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
	            updateDBList.add(updateDBSqlList);
	         }
	    }catch(Exception e){
	        errMsg = errMsg + "UpdateA99.CheckOtherRule Error:"+e.getMessage()+"<br>";
	        System.out.println("UpdateA99.CheckOtherRule Error:"+e.getMessage());
	    }	    
	    return updateDBList;
	}
 
	//100.06.27 fix 檢核992110農(漁)會全體淨值不可為空值或為0
	//100.08.04 fix 檢核992100農(漁)會全體資產總額不可為空值或為0或負數
	//110.08.16 fix 992450=Max(992452,992454,992456)[111.05.09取消] 
	//              992460=Max(992462,992464,992466)[111.05.09取消]

	private static List CheckOtherRule1(int errCount,String bank_type,String report_no,String m_year,String m_month,
			   String bank_code,String user_id,String user_name,String input_method){
	    List dbData = null;	
	    //List dbData_992450 = null;//110.08.16 add	
	    //List dbData_992460 = null;//110.08.16 add
	    StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		DataObject bean = null;
		BigDecimal amt992100 = null;
		/*111.05.09 fix
		BigDecimal amt992450 = null;//110.08.16 add
		BigDecimal amt992460 = null;//110.08.16 add
		BigDecimal amt992450_max = null;//110.08.16 add
		BigDecimal amt992460_max = null;//110.08.16 add
		BigDecimal amt992452 = null;//110.10.25 add
		BigDecimal amt992454 = null;//110.10.25 add
		BigDecimal amt992456 = null;//110.10.25 add
		BigDecimal amt992462 = null;//110.10.25 add
		BigDecimal amt992464 = null;//110.10.25 add
		BigDecimal amt992466 = null;//110.10.25 add
		BigDecimal amt992440 = null;//110.10.28 add
		BigDecimal amt992440_sum = null;//110.10.28 add
		*/
	    try{
	    	
	     	sqlCmd.append("select sum(decode(acc_code,'992100',amt,0)) as amt992100,");
	        sqlCmd.append("       sum(decode(acc_code,'992110',amt,0)) as amt992110 ");
	        /*111.05.09 fix
	        sqlCmd.append("       sum(decode(acc_code,'992440',amt,0)) as amt992440, ");//110.10.28 add
	        sqlCmd.append("       sum(decode(acc_code,'992452',amt,'992454',amt,'992456',amt,'992462',amt,'992464',amt,'992466',amt,0)) as amt992440_sum, ");//110.10.28 add
	        sqlCmd.append("       sum(decode(acc_code,'992450',amt,0)) as amt992450, ");
	        sqlCmd.append("       sum(decode(acc_code,'992452',amt,0)) as amt992452, ");
	        sqlCmd.append("       sum(decode(acc_code,'992454',amt,0)) as amt992454, ");
	        sqlCmd.append("       sum(decode(acc_code,'992456',amt,0)) as amt992456, ");
	        sqlCmd.append("       sum(decode(acc_code,'992460',amt,0)) as amt992460, ");
	        sqlCmd.append("       sum(decode(acc_code,'992462',amt,0)) as amt992462, ");
	        sqlCmd.append("       sum(decode(acc_code,'992464',amt,0)) as amt992464, ");
	        sqlCmd.append("       sum(decode(acc_code,'992466',amt,0)) as amt992466  ");
	        */
	        sqlCmd.append(" from a99 where m_year=? and m_month=? and bank_code=? and acc_code in ('992100','992110')");
	    	paramList.add(m_year);
	    	paramList.add(m_month);
	    	paramList.add(bank_code);	  
	    	dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt992100,amt992110");
	        
	    	//110.08.16 取得992452/992454/992456最大值[111.05.09取消]
	    	/*
	    	sqlCmd.delete(0, sqlCmd.length());	    		    	
	    	sqlCmd.append(" select max(amt) as amt992450_max");
	        sqlCmd.append(" from a99 where m_year=? and m_month=? and bank_code=? and acc_code in ('992452','992454','992456') ");
	        dbData_992450 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt992450_max");
	        System.out.println("dbData_992450.size="+dbData_992450.size());
	        if(dbData_992450 != null && dbData_992450.size() != 0){
	        	amt992450_max =BigDecimal.valueOf(Long.parseLong((((DataObject)dbData_992450.get(0)).getValue("amt992450_max")).toString()));//110.08.16 add
	        	System.out.println("amt992450_max="+amt992450_max);
	        }
	        */
	        //110.08.16 取得992462/992464/992466最大值[111.05.09取消]
	    	/*
	        sqlCmd.delete(0, sqlCmd.length());	    		    	
	    	sqlCmd.append(" select max(amt) as amt992460_max");
	        sqlCmd.append(" from a99 where m_year=? and m_month=? and bank_code=? and acc_code in ('992462','992464','992466')  ");
	        dbData_992460 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt992460_max");
	        System.out.println("dbData_992460.size="+dbData_992460.size());
	        if(dbData_992460 != null && dbData_992460.size() != 0){
	        	amt992460_max =BigDecimal.valueOf(Long.parseLong((((DataObject)dbData_992460.get(0)).getValue("amt992460_max")).toString()));//110.08.16 add
	        	System.out.println("amt992460_max="+amt992460_max);
	        }
	    	*/
	        if(dbData != null && dbData.size() != 0){
	        	System.out.println("errCount="+errCount);
		        bean = (DataObject)dbData.get(0);
		        amt992100 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992100")).toString()));
		        /*111.05.09取消
		        amt992440 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992440")).toString()));//110.10.28 add
		        amt992440_sum =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992440_sum")).toString()));//110.10.28 add
		        amt992450 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992450")).toString()));//110.08.16 add
		        amt992460 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992460")).toString()));//110.08.16 add
		        
		        amt992452 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992452")).toString()));//110.10.25 add
		        amt992454 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992454")).toString()));//110.10.25 add
		        amt992456 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992456")).toString()));//110.10.25 add
		        amt992462 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992462")).toString()));//110.10.25 add
		        amt992464 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992464")).toString()));//110.10.25 add
		        amt992466 =BigDecimal.valueOf(Long.parseLong((bean.getValue("amt992466")).toString()));//110.10.25 add
		        
		        System.out.println("amt992450="+amt992450+":amt992452="+amt992452+":amt992454="+amt992454+":amt992456="+amt992456);
		        System.out.println("amt992460="+amt992460+":amt992462="+amt992462+":amt992464="+amt992464+":amt992466="+amt992466);
		        System.out.println("amt992440="+amt992440+":amt992440_sum="+amt992440_sum);
		        */
		        //System.out.println("amt992100="+(bean.getValue("amt992100")).toString()+" <0?"+amt992100.compareTo(BigDecimal.ZERO));
	            //檢核992100農(漁)會全體資產總額必須申報,且不可為0/負數
	            System.out.println("errCount="+errCount);
	           
	            if( bean.getValue("amt992100") == null
                || ((bean.getValue("amt992100")).toString()).equals("0")				
				|| amt992100.compareTo(BigDecimal.ZERO) < 0                
				){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現[992100]農(漁)會全體資產總額不可為空值或為0或負數 ");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }	
	            
	            
	            //檢核992110農(漁)會全體淨值不可為空值或為0	           
	            if(bean.getValue("amt992110") == null || ((bean.getValue("amt992110")).toString()).equals("0")){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現[992110]農(漁)會全體淨值為空值或為0 ");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            /*111.05.09取消
	            //檢核992450=Max(992452,992454,992456) 
	            if(amt992450.compareTo(amt992450_max) != 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992450值為"+amt992450+"不等於992452、992454、992456的最大值為"+amt992450_max);					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            //992460=Max(992462,992464,992466) 
	            if(amt992460.compareTo(amt992460_max) != 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992460值為"+amt992460+"不等於992462、992464、992466的最大值為"+amt992460_max);					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            
	            //992452>992454>992456
	            if( amt992452.compareTo(amt992454) < 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992454大於992452");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            
	            if( amt992454.compareTo(amt992456) < 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992456大於992454");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            
	            //992462>992464>992466
	            if( amt992462.compareTo(amt992464) < 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992464大於992462");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            if( amt992464.compareTo(amt992466) < 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992466大於992464");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }
	            
	            //992440>(992452+992454+992456)+(992462+992464+992466)
	           	if( amt992440.compareTo(amt992440_sum) < 0){
	            	dataList = new ArrayList();//傳內的參數List		 	           				   
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add(bank_code); 
					dataList.add(report_no); 
					dataList.add(String.valueOf(errCount)); 
					dataList.add("檢核發現992440小於992452、992454、992456、992462、992464、992466之合計數");					 
					dataList.add(user_id); 
					dataList.add(user_name); 
					updateDBDataList.add(dataList);//1:傳內的參數List	                
	                errCount++ ;    
	            }	    
	            */
	        }
	        
	        if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
           	   sqlCmd.delete(0,sqlCmd.length());
			   sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");				
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
	    }catch(Exception e){
	        errMsg = errMsg + "UpdateA99.CheckOtherRule1 Error:"+e.getMessage()+"<br>";
	        System.out.println("UpdateA99.CheckOtherRule1 Error:"+e.getMessage());
	    }	    
	    return updateDBList;
	}
	//95.05.29 若為檔案上傳,將wlx01.credit_staff_num寫回992300======================================================
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//101.06.29 add 檔案上傳,當992300為0時,才將wlx01.credit_staff_num回寫 by 2295
	//101.07.09 add回寫總機構人數至A99時,先commit by 2295
	private static void ReWrite992300(int errCount,String bank_type,String report_no,String m_year,String m_month,
			   String bank_code,String user_id,String user_name,String input_method){	    		
	    		
	    List dbData = null;
	    List dbData_992300 = null;
	    StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
	    try{
	    	sqlCmd.append("select credit_staff_num from wlx01 where m_year=? and bank_no=?");
	    	paramList.add(wlx01_m_year);
	    	paramList.add(bank_code);	  
	    	dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"credit_staff_num");
	    	
	    	sqlCmd.delete(0,sqlCmd.length());
	    	sqlCmd.append(" select amt from A99 where m_year=? and m_month=? and bank_code=?");
            sqlCmd.append(" and acc_code = '992300'");
            paramList = new ArrayList();
            paramList.add(m_year);
            paramList.add(m_month);     
            paramList.add(bank_code);     
            dbData_992300 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
	    	
	        if(dbData != null && dbData.size() != 0){
	            if((((DataObject)dbData.get(0)).getValue("credit_staff_num") != null)
	            &&    input_method.equals("F"))/*檔案上傳*/
	            {
	                //101.06.29 add 當992300為0時,才將wlx01.credit_staff_num回寫 by 2295
	                System.out.println(bank_code+".992300="+(((DataObject)dbData_992300.get(0)).getValue("amt")).toString());
	                if(Integer.parseInt( (((DataObject)dbData_992300.get(0)).getValue("amt")).toString() ) == 0){	                
	                    System.out.println(bank_code+".992300.amt=0");
	                    sqlCmd.delete(0,sqlCmd.length());
	                    sqlCmd.append(" UPDATE A99 SET amt=? ");
	                    sqlCmd.append(" where m_year=? and m_month=? and bank_code=?");
	                    sqlCmd.append(" and acc_code = '992300'");
	                    dataList = new ArrayList();//傳內的參數List	
	            	    dataList.add((((DataObject)dbData.get(0)).getValue("credit_staff_num")).toString());
	            	    dataList.add(m_year); 
	            	    dataList.add(m_month); 
	            	    dataList.add(bank_code); 
	            	    updateDBDataList.add(dataList);//1:傳內的參數List
	                }
	            }
	        }
	        if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			   //回寫總機構人數至A99時,先commit 101.07.09 add
	           if(updateDBList != null && updateDBList.size()!=0){
	              boolean updateOK=DBManager.updateDB_ps(updateDBList);                     
	              System.out.println("ReWrite992300 update OK??"+updateOK+DBManager.getErrMsg());                
	           }
			}
	        
	    }catch(Exception e){
	        errMsg = errMsg + "UpdateA99.ReWrite992300 Error:"+e.getMessage()+"<br>";
	        System.out.println("UpdateA99.ReWrite992300 Error:"+e.getMessage());
	    }
	    //return updateDBList;
	}	
}
