//95.05.03 fix 修改顯示檢核結果顯示訊息 by 2295
//95.05.16 fix 檢核公式修改.不累加120601/120602/120603/120604 by 2295
//95.05.25 fix 科目代號不為150200.960500.970000且逾放期數為6個月~未滿1年.1年~未滿2年.2年以上皆設為"0" by 2295
//95.08.10 fix 公式檢核拿掉減項:950600 by 2295
//95.08.16 fix 6個月~未滿一年開放輸入/上傳 by 2295
//         fix checkRuleError-amt_L/amt_R改用bigDecimal顯示 by 2295
//99.11.10 若為檔案上傳批次寫入A06 ZeroData / 有修改的A06寫入A06_log by 2295
//         寫入暫存table AXX_TMP/有異動的資料寫入A06使用preparestatement by 2295
//         批次寫入WML01_Log/WML02_Log/WML03_Log/InsertWML03 by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//				      Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					  Utility.getWML01讀取WML01 all data
//				      Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//103.02.11 add 103/01以後.漁會套用新科目代號 by 2295
package com.tradevan.util;

import java.io.*;
import java.lang.Long;
import java.math.BigDecimal;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateA06 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A06(String report_no, String m_year,
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
		
		List A06List = new LinkedList();//A06細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 
		
		boolean nonZero=false;//94.03.25 檢查A06的內容是否都為"0"
		long[] sum_total = new long[6];
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		String[] amt_name = {"amt_3month","amt_6month","amt_1year","amt_2year","amt_over2year","amt_total"};
		String ncacno="ncacno";
		//99.11.09 add 查詢年度100年以前.縣市別不同===============================
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
		System.out.println("UpdateA06 begin");
		try { 
						
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA06FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table							
				bank_codeList = (List)allList.get(0);
				A06List = (List)allList.get(1);
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
			
			Utility.printLogTime("A06-1 time");	
			
			//若為"Y"表示由共用中心代傳入			
			sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
			paramList.add(srcbank_code);
			paramList.add(wlx01_m_year);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			if(dbData.size() != 0){//ba01有此筆bank_no的資料
			   System.out.println("common_center dbdata.size="+dbData.size());	
			   if((((DataObject)dbData.get(0)).getValue("bank_type") != null) && ((String)((DataObject)dbData.get(0)).getValue("bank_type")).equals("8")){			   	 
			   	  common_center="Y";
			   	  Utility.setSendMsg("");//95.03.20清空共用中心的mail msg			   	
	              Utility.addMailNotification("",report_no,m_year,m_month, checkResult,"",emailformat.format(filedate),filename,user_id);//95.03.27 add 彙總title
			   }
			}
			System.out.println("common_center="+common_center);
			System.out.println("bank_codeList.size="+bank_codeList.size());
			
			Utility.printLogTime("A06-2 time");	
			
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
  		    //96.04.19 若為檔案上傳批次寫入A01 ZeroData / 有修改的A01寫入A01_log
			if(input_method.equals("F")){//在A01中無資料的先將insert Zero data 到A01
		    	errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A06' and acc_div='08' ");//在A06中無資料的先將insert Zero data 到A06
		    	Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A01_LOG		    			    	
		    	errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤(科目代號不存在)		    	
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A01
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			//批次寫入WML01_LOG(檢核結果)(區分檔案上傳/線上編輯)
			Utility.InsertWML01_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
		
			
			List lockBank_CodeList = Utility.getLockBank_Code(m_year,m_month,report_no);
			bank_codeLoop:
			for(int i=0;i<bank_codeList.size();i++){
				//94.03.15 add 若被鎖定,則不執行檢核,直接send mail通知==================================
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
				//94.01.04 找出該bank_code的bank_type by 2295=============================================
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
				paramList = new ArrayList();
				paramList.add((String)bank_codeList.get(i));
				paramList.add(wlx01_m_year);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				
				if(dbData != null && dbData.size() != 0){
					bank_type = (String)((DataObject)dbData.get(0)).getValue("bank_type");
					System.out.println("bank_type="+bank_type);
					ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
					if( bank_type.equals("7") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301) ){
					    ncacno = "ncacno_7_rule";
		            }		            
				}
				//=================================================================================
				errCount = 0;
				Utility.printLogTime("A06-3 time");	
				//clear WML03(檢核其他錯誤)
				/*99.11.09移至InsertWML03(科目代號不存在)*/	
				dbData = Utility.getWML03_count(m_year,m_month,(String)bank_codeList.get(i),report_no);//99.11.19 fix
				/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)	
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select count(*) as data      from wml03 where m_year=? and m_month=? and bank_code =? and report_no = ?");				
				paramList = new ArrayList();
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(i));
				paramList.add(report_no);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
				*/
				if(dbData != null && dbData.size() > 0){
				   errCount =Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString());//99.11.09 先取得insertwml03的檢核錯誤總數				  
				}
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				//檢核合計數加總不符
				if(input_method.equals("F")){//檔案上傳					
					String acc_name="";					
					sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
					 
					for(int j=0;j<A06List.size();j++){						
						acc_name="";
						//m_year,m_month,bank_code_acc_code,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total
						if(((String)((List)A06List.get(j)).get(2)).equals((String)bank_codeList.get(i))){
							//System.out.println("bank_code="+(String)((List)A06List.get(j)).get(2));
							//System.out.println("acc_code="+(String)((List)A06List.get(j)).get(3));							
							
							sum_total[0] = Long.parseLong((String)((List)A06List.get(j)).get(4)==null?"0":(String)((List)A06List.get(j)).get(4));//amt_3month
							sum_total[1] = Long.parseLong((String)((List)A06List.get(j)).get(5)==null?"0":(String)((List)A06List.get(j)).get(5));//amt_6month
							sum_total[2] = Long.parseLong((String)((List)A06List.get(j)).get(6)==null?"0":(String)((List)A06List.get(j)).get(6));//amt_1year
							sum_total[3] = Long.parseLong((String)((List)A06List.get(j)).get(7)==null?"0":(String)((List)A06List.get(j)).get(7));//amt_2year
							sum_total[4] = Long.parseLong((String)((List)A06List.get(j)).get(8)==null?"0":(String)((List)A06List.get(j)).get(8));//amt_over2year
							sum_total[5] = Long.parseLong((String)((List)A06List.get(j)).get(9)==null?"0":(String)((List)A06List.get(j)).get(9));//amt_total
							acc_name = getAcc_Code(ncacno,(String)((List)A06List.get(j)).get(3));
							
							if(sum_total[0]+sum_total[1]+sum_total[2]+sum_total[3]+sum_total[4] != sum_total[5]){
							    dataList =  new ArrayList();//儲存參數的data
							    dataList.add(m_year);
						        dataList.add(m_month);
						        dataList.add((String)((List)A06List.get(j)).get(2));//bank_code
						        dataList.add(report_no);
						        dataList.add(String.valueOf(errCount));
								if(((String)((List)A06List.get(j)).get(1)).equals("970000")){
							        dataList.add(acc_name + "項目的逾放合計金額:與該科目未滿3個月、3個月~未滿6個月、6個月~未滿1年、1年~未滿2年、2年以上等合計逾放金額的加總不符");							      
							    }else{							        
							    	dataList.add(acc_name + "逾放合計金額:與該科目未滿3個月、3個月~未滿6個月、6個月~未滿1年、1年~未滿2年、2年以上等逾放金額的加總不符");    
							    }
								dataList.add(user_id);
						        dataList.add(user_name);
						    	updateDBDataList.add(dataList);//1:傳內的參數List					
								
								errCount++ ;		
								
							}
						}//end of bank_code相同時					
					}//end of A06List
				}else{
					//input_method=="W"-->線上編輯
				    ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
                    if( bank_type.equals("7") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301) ){
                        ncacno = "ncacno_7_rule";
                    }   
					sqlCmd.append(" select a06.acc_code,"+ncacno+".acc_name, amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
					sqlCmd.append(" from a06 ");  
					sqlCmd.append(" left join "+ncacno+" on "+ncacno+ ".acc_code=a06.acc_code ");
					sqlCmd.append(" and "+ncacno+".acc_div='08'");
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=?");
					
					paramList = new ArrayList();
					paramList.add(m_year);
					paramList.add(m_month);
					paramList.add((String)bank_codeList.get(i));					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
					
				    if(dbData != null && dbData.size() != 0){					        
				        System.out.println("線上編輯.A06.size()="+dbData.size());	
				        sqlCmd.delete(0,sqlCmd.length());
					    sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");					   
				    	for(int idx=0;idx < dbData.size();idx++){
				    	    sum_total[0] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_3month")).toString());//amt_3month
				    	    sum_total[1] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_6month")).toString());//amt_6month
				    	    sum_total[2] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_1year")).toString());//amt_1year
				    	    sum_total[3] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_2year")).toString());//amt_2year
				    	    sum_total[4] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_over2year")).toString());//amt_over2year
				    	    sum_total[5] = Long.parseLong((((DataObject)dbData.get(idx)).getValue("amt_total")).toString());//amt_total
				    	    
				    	    if(sum_total[0]+sum_total[1]+sum_total[2]+sum_total[3]+sum_total[4] != sum_total[5]){
							    dataList =  new ArrayList();//儲存參數的data
							    dataList.add(m_year);
						        dataList.add(m_month);
						        dataList.add((String)bank_codeList.get(i));//bank_code
						        dataList.add(report_no);
						        dataList.add(String.valueOf(errCount));
						        if(((String)((DataObject)dbData.get(idx)).getValue("acc_code")).equals("970000")){
							        dataList.add((String)((DataObject)dbData.get(idx)).getValue("acc_name") + "項目的逾放合計金額:與該科目未滿3個月、3個月~未滿6個月、6個月~未滿1年、1年~未滿2年、2年以上等合計逾放金額的加總不符");							      
							    }else{							        
							    	dataList.add((String)((DataObject)dbData.get(idx)).getValue("acc_name") + "逾放合計金額:與該科目未滿3個月、3個月~未滿6個月、6個月~未滿1年、1年~未滿2年、2年以上等逾放金額的加總不符");    
							    }
								dataList.add(user_id);
						        dataList.add(user_name);
						    	updateDBDataList.add(dataList);//1:傳內的參數List					
								
								errCount++ ;		
								
							}
				    	}
				    }
			    	System.out.println("線上編輯.nonZero="+nonZero); 
				}//end of 線上編輯
				//99.11.10 有合計數加總不符
				if(updateDBDataList.size() != 0){
				   updateDBSqlList.add(sqlCmd.toString());
				   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);				   
				}
				
				Utility.printLogTime("A06-4 time");	
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
				    //無其他錯誤時,才檢核公式
					/*96.04.19 移至InsertWML02_LOG */				   
				    Utility.printLogTime("A06-5 time");	
				    
					nonZero = false;
					dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"(amt_3month > 0 or amt_6month > 0 or amt_1year > 0 or amt_2year > 0 or amt_over2year > 0 or amt_total > 0)");//99.11.19 fix
				    /*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
										
					 if(input_method.equals("W")){//線上編輯
					 	sqlCmd.append("select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? ");
					 	sqlCmd.append(" and (amt_3month > 0 or amt_6month > 0 ");
					 	sqlCmd.append(" or amt_1year > 0 or amt_2year > 0");
					 	sqlCmd.append(" or amt_over2year > 0 or amt_total > 0)");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));
					}else{//檔案上傳
						sqlCmd.append("select count(*) as countzero from axx_tmp where m_year=? AND m_month=? AND bank_code=? and report_no=? ");
						sqlCmd.append(" and (amt > 0 or amt1 > 0 or amt2 > 0 or amt3 > 0 or amt4 >0 or amt5 > 0)"); 
						sqlCmd.append(" and filename=?");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));
						paramList.add(report_no);
						paramList.add(filename);						
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
				   	
				   	Utility.printLogTime("A06-6 time");	
					
					if(nonZero){//94.03.25 fix 若A06的值,只要有非"0"值時,才執行檢核
				       //執行檢核
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,errCount);
				       if(ruleList.size() > 0){//99.11.10有檢核失敗時,才加入				       	  
				          updateDBList.add((List)ruleList.get(0));				          
				          errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.11.10累計參數List的size
				       }
				       Utility.printLogTime("A06-7 time");
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
			    //99.11.10
				//dbData = DBManager.QueryDB("select * from WML01 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +						
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
						
					   //99.11.10sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");							
					   //99.11.10sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
				       //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
					}
						
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
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}else{//WML01存在時,做update
					//99.11.10移至批次寫入InsertWML01_LOG
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			
					//99.11.10sqlCmd=" UPDATE WML01 SET " 
					//	  +" upd_method='"+upd_method+"',upd_code='"+upd_code+"',batch_no="+batch_no+",user_id='"+user_id+"',user_name='"+user_name+"',update_date=sysdate"   
					//	  +" where m_year=" + m_year + " AND m_month=" + m_month + " AND " 						
					//	  +" bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
					
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
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
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
			
			Utility.printLogTime("A06-8 time");	
			if(input_method.equals("F")){//檔案上傳
			   //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.10updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19
			    //99.11.10 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			}
			//updateOK=DBManager.updateDB(updateDBSqlList);
			if(updateDBList != null && updateDBList.size()!=0){
			   updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA06.doParserReport_A06 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("W"))/*線上編輯*/ errMsg = upd_code +":" + errMsg;//94.11.14
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
			Utility.printLogTime("A06-9 time");	
			if(common_center.equals("Y")){//95.03.20由共用中心傳入的,傳送彙總mail
			    if(input_method.equals("F")){//檔案上傳
			       Utility.sendMailNotification(srcbank_code,report_no,
				           m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,Utility.getSendMsg());			    
			    }else{//線上編輯
				   Utility.sendMailNotification(srcbank_code,report_no,
				           m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,Utility.getSendMsg());
			    }   
			}	
			Utility.printLogTime("A06-10 time");	
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA06.doParserReport_A06 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA06.doParserReport_A06="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	//95.05.25 fix 科目代號不為150200.960500.970000且逾放期數為6個月~未滿1年.1年~未滿2年.2年以上皆設為"0"
	private static List getA06FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List A06List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAmt="";
			int[][] amtIdx = {{18,32},/*未滿3個月*/
			        		  {32,46},/*3個月~未滿6個月*/
			        		  {46,60},/*6個月~未滿1年*/
			        		  {60,74},/*1年~未滿2年*/
			        		  {74,88},/*2年以上*/
			        		  {88,102}/*逾放合計*/};
			//99.11.09 寫入AXX_TMP暫存table,使用preparestatement
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
				
				//98.11.09 寫入AXX_TMP暫存table
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
								System.out.println("add bank_code list="+txtline.substring(5,12));
							}
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號
							
							for(int i=0;i<amtIdx.length;i++){
							    tmpAmt = txtline.substring(amtIdx[i][0],amtIdx[i][1]);//金額
							    //95.05.25 fix 科目代號不為150200.960500.970000且逾放期數為6個月~未滿1年.1年~未滿2年.2年以上皆設為"0"
							    //95.08.16 fix 6個月~未滿一年開放輸入/上傳 by 2295
							    if((!(txtline.substring(12,18).equals("150200") ||
							       txtline.substring(12,18).equals("960500") ||
							       txtline.substring(12,18).equals("970000"))) && (/*i==2 ||*/ i==3 || i==4)){
							       tmpAmt = "0";   
							    }
							    //94.03.07 fix 有負號("-")時,把之前的"0"去掉..
							    if(tmpAmt.indexOf("-") != -1){
							       tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							    }
							    //=========================================
							    detail.add(tmpAmt);							    
							}
							detail.add("0");
							detail.add("0");
							detail.add("0");
							detail.add("0");
							detail.add(report_no);//99.11.09
							detail.add(filename);//99.11.09								
							updateDBDataList.add(detail);//1:傳內的參數List
							A06List.add(detail);
						}
				}	
				in.close();
				f.close();
				//99.11.09 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
			}catch(Exception e){
				errMsg = errMsg + "UpdateA06.getA06FileData Error:"+e.getMessage()+"<br>";
			}
			System.out.println("bank_codeList.size="+bank_codeList.size());
			System.out.println("A06List.size="+A06List.size());
			updateDBSqlList = null;
			allList.add(bank_codeList);
			allList.add(A06List);
			return allList;
	}
	
	//公式檢核
	//99.11.10 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List CheckRule(String bank_type,String report_no,String m_year,String m_month,
								 String bank_code,String user_id,String user_name,int errCount){
		
		String  amt="";		
		//double	amt_L = 0.0, amt_R = 0.0,
		//double amt_tbl = 0.0;		
		double amt_L[]= new double[6];
		double amt_R[]= new double[6];
		double amt_tbl[]= new double[6];
		String	quop = null;
		String  cano="";
		String noop="";
		String acc_code="";
		String lstr = ""; 
		String rstr = "";
			
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.10 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		List ruleList = new ArrayList();
		try{
		    /*
		     select r1.cano,r1.quop,r2.left_flag,
		     a06.amt_3month,a06.amt_6month,a06.amt_1year,a06.amt_2year,a06.amt_over2year,a06.amt_total,r2.noop,
		     A06.acc_code
		     from a06,ruleno1 r1,ruleno2 r2
		     where a06.acc_code=r2.acc_code 				   
		     and r1.CANO = r2.cano
		     and r1.CANO like '08%'
		     and r1.acc_type = r2.acc_type
		     and r1.acc_type = '6'
		     and a06.m_year=94 
		     and a06.m_month=1
		     and a06.bank_code='6030016'
		     order by r1.cano,r2.left_flag,r2.nserial
		     */
		    
			//check ruleno1, ruleno2
			sqlCmd.append(" select r1.cano,r1.quop,r2.left_flag,");
			sqlCmd.append(" a06.amt_3month,a06.amt_6month,a06.amt_1year,a06.amt_2year,a06.amt_over2year,a06.amt_total,"); 
			sqlCmd.append("  r2.noop,a06.acc_code ");
			sqlCmd.append(" from a06,ruleno1 r1,ruleno2 r2 ");
			sqlCmd.append(" where a06.acc_code=r2.acc_code ");			   
			sqlCmd.append(" and r1.cano = r2.cano");
			sqlCmd.append(" and r1.cano like '08%'"); //a06的div='08',公式起始為'08'開頭
			sqlCmd.append(" and r1.acc_type = r2.acc_type");
			sqlCmd.append(" and r1.acc_type = ?");
			sqlCmd.append(" and a06.m_year=?"); 
			sqlCmd.append(" and a06.m_month=?");
			sqlCmd.append(" and a06.bank_code=?");
		    sqlCmd.append(" order by r1.cano,r2.left_flag,r2.nserial");
			
		    paramList.add(bank_type);
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
			
			cano=(String)((DataObject)dbData.get(0)).getValue("cano");
			
			sqlCmd.delete(0,sqlCmd.length());
			paramList = new ArrayList();
			sqlCmd.append(" select acc_code, noop, left_flag, nserial FROM ruleno2 ");
			sqlCmd.append(" where acc_type=?");
			sqlCmd.append(" and   cano=? ORDER BY 3, 4");
			paramList.add(bank_type);
			paramList.add(cano);
		    List dbData2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"nserial");    
		    System.out.println("dbData2.size="+dbData2.size());
		    int j = 0;
		    while (j < dbData2.size()) {
			      if (((String)((DataObject)dbData2.get(j)).getValue("acc_code")).substring(0, 3).equals("FIX")){	//如果是常數，就將常數值取出
				      acc_code = "(" + ListArray_FC.getCONST_VALUE((String)((DataObject)dbData2.get(j)).getValue("acc_code")) + ")";
			      }else{
				      acc_code = (String)((DataObject)dbData2.get(j)).getValue("acc_code");
			      }
			      if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("0"))
				      lstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());
			      if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("1"))
				      rstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());										
			      //System.out.println("lstr="+lstr);	
			      //System.out.println("rstr="+rstr);	
			      j ++ ;	
		    } //end of while(dbData2--ruleno2)
		    //System.out.println("lstr="+lstr);	
	        //System.out.println("rstr="+rstr);
			quop = ((String)((DataObject)dbData.get(0)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());
			System.out.println("check rule begin");
			//amt_L = 0.0; amt_R = 0.0;
			//amt_tbl=0.0;
			for(int r=0;r<dbData.size();r++){
				//System.out.println("r="+r);
				if(!(((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals(cano)){//與前一個公式不同時				    
				    //取得檢核結果
				    for(int periodidx=0;periodidx<amt_L.length;periodidx++){
				    	   ruleList = new ArrayList();
				    	   ruleList = checkRuleError(periodidx,quop, amt_L[periodidx],amt_R[periodidx],bank_code,m_year,m_month,report_no,lstr,rstr,user_id,user_name,errCount,cano);
				    	   if(ruleList.size() >0){
				    	   	  System.out.println("have checkRuleError");
				    	   	  updateDBDataList.add(ruleList);	           
							  //System.out.println("公式不同="+sqlCmd);							   
							  errCount++;
					       }
					}   
					//把這次的公式編號及quop儲存;並把金額清空
					cano=(String)((DataObject)dbData.get(r)).getValue("cano");
					quop = ((String)((DataObject)dbData.get(r)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());
					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" select acc_code, noop, left_flag, nserial FROM ruleno2 ");
					sqlCmd.append(" where acc_type=?");
					sqlCmd.append(" and   cano=? ORDER BY 3, 4");
					dbData2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"nserial");    
					System.out.println("dbData2.size="+dbData2.size());
					j = 0;
					lstr="";
					rstr="";
					while (j < dbData2.size()) {
					       if (((String)((DataObject)dbData2.get(j)).getValue("acc_code")).substring(0, 3).equals("FIX")){	//如果是常數，就將常數值取出
							    acc_code = "(" + ListArray_FC.getCONST_VALUE((String)((DataObject)dbData2.get(j)).getValue("acc_code")) + ")";
						   }else{
							    acc_code = (String)((DataObject)dbData2.get(j)).getValue("acc_code");
						   }
						   if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("0"))
							     lstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());
						   if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("1"))
							     rstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());										
						   //System.out.println("lstr="+lstr);	
						   //System.out.println("rstr="+rstr);	
						   j ++ ;	
					} //end of while(dbData2--ruleno2)
					//System.out.println("lstr="+lstr);	
					//System.out.println("rstr="+rstr);
				}	

				amt_tbl[0] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_3month")).toString());
				amt_tbl[1] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_6month")).toString());
				amt_tbl[2] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_1year")).toString());
				amt_tbl[3] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_2year")).toString());
				amt_tbl[4] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_over2year")).toString());
				amt_tbl[5] = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_total")).toString());
				/*
				System.out.println("acc_code="+(String)((DataObject)dbData.get(r)).getValue("acc_code"));
				System.out.println("noop="+noop);
				System.out.println("amt_tbl[0]="+amt_tbl[0]);
				System.out.println("amt_tbl[1]="+amt_tbl[1]);
				System.out.println("amt_tbl[2]="+amt_tbl[2]);
				System.out.println("amt_tbl[3]="+amt_tbl[3]);
				System.out.println("amt_tbl[4]="+amt_tbl[4]);
				*/
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("0")){//左式					
					if ((noop == null) || (noop.equals(""))) {
						amt_L[0] = amt_tbl[0];
						amt_L[1] = amt_tbl[1];
						amt_L[2] = amt_tbl[2];
						amt_L[3] = amt_tbl[3];
						amt_L[4] = amt_tbl[4];
						amt_L[5] = amt_tbl[5];
					}else if (noop.equals("+")) {
						amt_L[0] += amt_tbl[0];
						amt_L[1] += amt_tbl[1];
						amt_L[2] += amt_tbl[2];
						amt_L[3] += amt_tbl[3];
						amt_L[4] += amt_tbl[4];						
						amt_L[5] += amt_tbl[5];
					}else if (noop.equals("-")) {
						amt_L[0] -= amt_tbl[0];
						amt_L[1] -= amt_tbl[1];
						amt_L[2] -= amt_tbl[2];
						amt_L[3] -= amt_tbl[3];
						amt_L[4] -= amt_tbl[4];
						amt_L[5] -= amt_tbl[5];
					}else if (noop.equals("*")) {
						amt_L[0] *= amt_tbl[0];
						amt_L[1] *= amt_tbl[1];
						amt_L[2] *= amt_tbl[2];
						amt_L[3] *= amt_tbl[3];
						amt_L[4] *= amt_tbl[4];
						amt_L[5] *= amt_tbl[5];
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_3month")).toString()) != 0){
							amt_L[0] /= amt_tbl[0];
					    }
					    if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_6month")).toString()) != 0){
						    amt_L[1] /= amt_tbl[1];
				        }
				        if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_1year")).toString()) != 0){
					        amt_L[2] /= amt_tbl[2];
			            }
			            if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_2year")).toString()) != 0){
				            amt_L[3] /= amt_tbl[3];
		                }
		                if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_over2year")).toString()) != 0){
			                amt_L[4] /= amt_tbl[4];
	                    }
		                if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_total")).toString()) != 0){
			                amt_L[5] /= amt_tbl[5];
	                    }
					}//end of noop
				}//end of left_flag=0-->左式金額加總	
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("1")){//右式
				    if ((noop == null) || (noop.equals(""))) {
						amt_R[0] = amt_tbl[0];
						amt_R[1] = amt_tbl[1];
						amt_R[2] = amt_tbl[2];
						amt_R[3] = amt_tbl[3];
						amt_R[4] = amt_tbl[4];
						amt_R[5] = amt_tbl[5];
					}else if (noop.equals("+")) {
						amt_R[0] += amt_tbl[0];
						amt_R[1] += amt_tbl[1];
						amt_R[2] += amt_tbl[2];
						amt_R[3] += amt_tbl[3];
						amt_R[4] += amt_tbl[4];						
						amt_R[5] += amt_tbl[5];
					}else if (noop.equals("-")) {
						amt_R[0] -= amt_tbl[0];
						amt_R[1] -= amt_tbl[1];
						amt_R[2] -= amt_tbl[2];
						amt_R[3] -= amt_tbl[3];
						amt_R[4] -= amt_tbl[4];
						amt_R[5] -= amt_tbl[5];
					}else if (noop.equals("*")) {
						amt_R[0] *= amt_tbl[0];
						amt_R[1] *= amt_tbl[1];
						amt_R[2] *= amt_tbl[2];
						amt_R[3] *= amt_tbl[3];
						amt_R[4] *= amt_tbl[4];
						amt_R[5] *= amt_tbl[5];
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_3month")).toString()) != 0){
							amt_R[0] /= amt_tbl[0];
					    }
					    if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_6month")).toString()) != 0){
						    amt_R[1] /= amt_tbl[1];
				        }
				        if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_1year")).toString()) != 0){
					        amt_R[2] /= amt_tbl[2];
			            }
			            if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_2year")).toString()) != 0){
				            amt_R[3] /= amt_tbl[3];
		                }
		                if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_over2year")).toString()) != 0){
			                amt_R[4] /= amt_tbl[4];
	                    }
		                if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt_total")).toString()) != 0){
			                amt_R[5] /= amt_tbl[5];
	                    }
					}//end of noop					
				}//end of left_flag=0-->右式金額加總
				//94.03.04 fix 把最後一筆的檢核失敗寫入db=================================
				if(r==dbData.size()-1){//檢核結果
				   for(int periodidx=0;periodidx<amt_L.length;periodidx++){
				       ruleList = new ArrayList();
			    	   ruleList = checkRuleError(periodidx,quop, amt_L[periodidx],amt_R[periodidx],bank_code,m_year,m_month,report_no,lstr,rstr,user_id,user_name,errCount,cano);
			    	   if(ruleList.size() >0){
			    	   	  updateDBDataList.add(ruleList);	
			    	   	  System.out.println("最後一筆=have checkRuleError");
				          //System.out.println("最後一筆="+sqlCmd);						   
						  errCount++;
				       }				      
				   }   
				}
				//==========================================================
				noop = (((DataObject)dbData.get(r)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("noop")).trim());
				//System.out.println("r="+r+"end");				
			} //end of for()
			sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
			if(updateDBDataList.size() > 0){
				updateDBSqlList.add(sqlCmd.toString());
				updateDBSqlList.add(updateDBDataList);
				updateDBList.add(updateDBSqlList);
			}
			System.out.println("check rule end");
		}catch(Exception e){
			errMsg = errMsg + "UpdateA06.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA06.CheckRule Error:"+e.getMessage());
		}
		return updateDBList;
	}
	
	
	private static List checkRuleError(int idx,String quop,double amt_L,double amt_R,String bank_code,String m_year,String m_month,String report_no,String lstr,String rstr,String user_id,String user_name,int errCount,String cano){
	    //95.08.16 fix amt_L/amt_R改用bigDecimal顯示 by 2295
		//99.11.10 add
		List dataList =  new ArrayList();//儲存參數的data
	    String[] period = {"未滿3個月","3個月~未滿6個月","6個月~未滿1年","1年~未滿2年","2年以上","逾放合計"};
	    try{
	        if ((quop.equals("=") && (amt_L == amt_R))){				   
		        //System.out.println("amt_L="+amt_L);
		        //System.out.println("amt_R="+amt_R);
		        //System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));
		    } else if ((quop.equals(">") && (amt_L > amt_R))) {
		    } else if ((quop.equals("<") && (amt_L < amt_R))) {
		    } else if ((quop.equals(">=") && (amt_L >= amt_R))) {
		    } else if ((quop.equals("<=") && (amt_L <= amt_R))) {
		    } else if ((quop.equals("!=") && (amt_L != amt_R))) {
		    }else {
			    System.out.println("WML03 have error");
			    System.out.println("cano="+cano);
			    //sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");			   
			    dataList.add(m_year);
		        dataList.add(m_month);
		        dataList.add(bank_code);
		        dataList.add(report_no);
		        dataList.add(String.valueOf(errCount));
		        if(cano.equals("080010")){
			        dataList.add(period[idx]+"的合計金額["+BigDecimal.valueOf((new Double(amt_L)).longValue())+"]:與同逾放期數的科目金額["+BigDecimal.valueOf((new Double(amt_R)).longValue())+"]加總不符");							      
			    }else{							        
			    	dataList.add(period[idx]+"的「農發基金放款-小計」["+BigDecimal.valueOf((new Double(amt_L)).longValue())+"]:與同逾放期數的「農發基金-XX放款」["+BigDecimal.valueOf((new Double(amt_R)).longValue())+"]各個科目加總不符");    
			    }
				dataList.add(user_id);
		        dataList.add(user_name);		    	
		    }
	    }catch(Exception e){
	        System.out.println("UpdateA06.checkRuleError:"+e.getMessage());
	    }
	    //System.out.println("sqlCmd="+sqlCmd);
	 	return dataList;
	}
	
    public static String getAcc_Code(String ncacno,String acc_code){
    	String acc_name = "";
    	List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
		sql.append("select * from "+ncacno+" where acc_tr_type='A06' and acc_code = ? order by acc_range");		
		paramList.add(acc_code);		
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
		if(dbData != null && dbData.size() > 0){
		  acc_name = (String)((DataObject)dbData.get(0)).getValue("acc_name");
		}		
		return acc_name;    		
    }    
}
