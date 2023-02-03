//104.10.12 create A13 by 2295 
package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;
import java.lang.Long;
import java.math.BigDecimal;


public class UpdateA13 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A13(String report_no, String m_year,
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
		List A05List = new LinkedList();//A05細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 		
		boolean nonZero=false;//94.03.28 檢查A13的內容是否都為"0"
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
			Utility.printLogTime("A13 begin time");	
			
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA13FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table									
				bank_codeList = (List)allList.get(0);
				A05List = (List)allList.get(1);
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
			Utility.printLogTime("A13-1 time");	
			
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
			Utility.printLogTime("A13-2 time");
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
			//96.08.15 若為檔案上傳批次寫入A13 ZeroData / 有修改的A13寫入A13_log
			if(input_method.equals("F")){//在A13中無資料的先將insert Zero data 到A13
				errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A13' and acc_div in ('12','13') ");//在A05中無資料的先將insert Zero data 到A05		    	
				Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A13_LOG
		    	errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A03		    	
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
				Utility.printLogTime("A13-3 time");
				
				System.out.println("bank_codeList["+i+"]="+(String)bank_codeList.get(i));
				errCount = 0;
				//94.03.22 找出該bank_code的bank_type by 2295=============================================
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
				//clear WML03(檢核其他錯誤)
				/*96.08.15移至InsertWML03*/
				Utility.printLogTime("A13-4 time");
				
				//clear WML02(檢核公式)
				dbData = Utility.getWML03_count(m_year,m_month,(String)bank_codeList.get(i),report_no);//99.11.19 fix
				
			    if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) == 0){//96.08.15檢查WML03有無其他錯誤,無其他錯誤時,才檢核公式
				    System.out.println("bank_code="+(String)bank_codeList.get(i)+".WML03.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				    /*96.08.15 移至InsertWML02_LOG*/
				    //96.04.18移至InsertZeroA05_List/InsertZeroExtraA05_List
				    //96.04.19 移至InsertA05_LOG
				    Utility.printLogTime("A13-5 time");
				    
				    nonZero = false;
				    dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"amt > 0");//99.11.19 fix
				    
					if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					       nonZero = true;
					   }
					}
					
				    Utility.printLogTime("A13-6 time");
				    	
				    if(nonZero){//94.03.28 fix 若A05的值,只要有非"0"值時,才執行檢核
				       //執行檢核
				       
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
					   if(ruleList.size() > 0){//99.11.15有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
					   }
				       
				       Utility.printLogTime("A13-7 time");
				       //====================================================================================				      
				    }				
				}else{//enf of errCount==0
				   errCount = Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString());	
				}//enf of errCount==0
			    dbData =  Utility.getWML01(m_year,m_month,(String)bank_codeList.get(i),report_no);
					    
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
				Utility.printLogTime("A13-9 time");
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
			Utility.printLogTime("A13-10 time");
				
			if(input_method.equals("F")){//檔案上傳
			 	//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19				
			    //96.12.20 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);			   
			}
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA13.doParserReport_A13 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			Utility.printLogTime("A13-11 time");
			
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA13.doParserReport_A13 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA13.doParserReport_A13="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List getA13FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List A13List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAcc_Code="";
			String tmpAmt="";
			//96.08.15 寫入AXX_TMP暫存table,使用preparestatement
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
			List updateDBList = new LinkedList();//0:sql 1:data		
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new LinkedList();//儲存參數的data
			Utility.printLogTime("getA13FileData begin time");
				
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
				Map values = new HashMap();//96.04.10 add
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
				
				//96.08.14 寫入AXX_TMP暫存table
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
								
							}
							
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號
							tmpAcc_Code=txtline.substring(12,18);
							//94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
							tmpAmt = txtline.substring(18,32);//金額
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							//=========================================									
							detail.add(tmpAmt);//金額 axx_tmp.amt
							detail.add("0");//99.11.15 add amt1
							detail.add("0");//99.11.15 add amt2
							detail.add("0");//99.11.15 add amt3
							detail.add("0");//99.11.15 add amt4
							detail.add("0");//99.11.15 add amt5
							detail.add("0");//99.11.15 add amt6
							detail.add("0");//99.11.15 add amt7
							detail.add("0");//99.11.15 add amt8
							detail.add("0");//99.11.15 add amt9		
							detail.add("");//amt_name
							detail.add(report_no);//96.08.15
							detail.add(filename);//96.08.15							
							
							updateDBDataList.add(detail);//1:傳內的參數List
							A13List.add(detail);
						}
				}	
				in.close();
				f.close();
				//96.08.15 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				updateDBSqlList.add(A13List);
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}				
			}catch(Exception e){
				errMsg = errMsg + "UpdateA13.getA13FileData Error:"+e.getMessage()+"<br>";
				return null;
			}
					
			allList.add(bank_codeList);
			allList.add(A13List);
			Utility.printLogTime("getA05FileData end time");
				
			return allList;
	}
	
	//公式檢核
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List CheckRule(String bank_type,String report_no,String m_year,String m_month,
								 String bank_code,String user_id,String user_name){
		
		String  amt="";		
		double	amt_L = 0.0, amt_R = 0.0,amt_tbl = 0.0,amt_assumed=0.0;
		String	quop = null;
		String  cano="";
		String acc_code="";
		String noop="";		
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logformat.format(nowlog);
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckRule begin time="+logformat.format(nowlog));	
		try{
		    //check ruleno1, ruleno2
            sqlCmd.append(" select r1.cano,r1.quop,r2.left_flag,a13.amt,r2.noop,a13.acc_code ");
            sqlCmd.append(" from a13,ruleno1 r1,ruleno2 r2 ");
            sqlCmd.append(" where a13.acc_code=r2.acc_code ");             
            sqlCmd.append(" and r1.CANO = r2.cano");
            sqlCmd.append(" and r1.acc_type = r2.acc_type");
            sqlCmd.append(" and r1.acc_type = ?");
            sqlCmd.append(" and (r1.CANO like '12%' or r1.CANO like '13%') ");//95.06.01 fix 抓A13的公式代號為12/13開頭的
            sqlCmd.append(" and a13.m_year=?"); 
            sqlCmd.append(" and a13.m_month=?");
            sqlCmd.append(" and a13.bank_code=?");
            sqlCmd.append(" order by r1.cano,r2.left_flag,r2.nserial"); 
            paramList.add(bank_type);
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
			cano=(String)((DataObject)dbData.get(0)).getValue("cano");
			quop = ((String)((DataObject)dbData.get(0)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());
			System.out.println("check rule begin");
			amt_L = 0.0; amt_R = 0.0;
			amt_tbl=0.0;
			
			sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append("INSERT INTO WML02 VALUES (?,?,?,?,?,?,?,?,?,sysdate)");
			
			ruleLoop:
			for(int r=0;r<dbData.size();r++){				
				//System.out.println("r="+r);
				
				if(!(((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals(cano)){//與前一個公式不同時
					//System.out.println("amt_L="+amt_L);
					//System.out.println("amt_R="+amt_R);
					if ((quop.equals("=") && (amt_L == amt_R))){
					} else if ((quop.equals(">") && (amt_L > amt_R))) {
					} else if ((quop.equals("<") && (amt_L < amt_R))) {
					} else if ((quop.equals(">=") && (amt_L >= amt_R))) {
					} else if ((quop.equals("<=") && (amt_L <= amt_R))) {
					} else if ((quop.equals("!=") && (amt_L != amt_R))) {
					}else {	    	
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
					}
					//把這次的公式編號及quop儲存;並把金額清空
					cano=(String)((DataObject)dbData.get(r)).getValue("cano");
					quop = ((String)((DataObject)dbData.get(r)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("quop")).trim());					
					amt_L = 0.0; amt_R = 0.0;
					amt_tbl=0.0;
				}
				acc_code = (((DataObject)dbData.get(r)).getValue("acc_code")).toString();
				amt_tbl = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString());
				
				
				//System.out.println("acc_code="+acc_code);
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
						//System.out.println("now is /");
						//System.out.println("amt_L="+amt_L);
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
						//System.out.println("now is /");
						//System.out.println("amt_R="+amt_R);
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_R /= amt_tbl;											
					}//end of noop
				}//end of left_flag=0-->右式金額加總
				
				//94.03.04 fix 把最後一筆的檢核失敗寫入db=================================
				if(r==dbData.size()-1){
				   if ((quop.equals("=") && (amt_L == amt_R))){
				   } else if ((quop.equals(">") && (amt_L > amt_R))) {
				   } else if ((quop.equals("<") && (amt_L < amt_R))) {
				   } else if ((quop.equals(">=") && (amt_L >= amt_R))) {
				   } else if ((quop.equals("<=") && (amt_L <= amt_R))) {
				   } else if ((quop.equals("!=") && (amt_L != amt_R))) {
				   }else {	    	
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
				   }
				}
				//==========================================================
				/*if(cano.equals("070060") && acc_code.substring(acc_code.length()-1).equals("P"))	//略過下一個為該資產之帳面金額
					r++;*/
				noop = (((DataObject)dbData.get(r)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("noop")).trim());
				//System.out.println("r="+r+"end");				
			} //end of for()
			System.out.println("WML02 have error.updateDBDataList.size()="+updateDBDataList.size());
			if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			System.out.println("check rule end");
		}catch(Exception e){
			errMsg = errMsg + "UpdateA13.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA13.CheckRule Error:"+e.getMessage());
		}
		
		return updateDBList;
	}
}
