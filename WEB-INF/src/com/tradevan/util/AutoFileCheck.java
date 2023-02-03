/*
 * Created on 2005/1/4
 *  96.02.27 fix link 至農金局網頁使用https://localhost by 2295
 * 101.10.12 add A06/A08/A09/A10/A99/F01 by 2295
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package com.tradevan.util;

import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import java.util.*;
import java.util.LinkedList;
import com.tradevan.util.UpdateA01;
import com.tradevan.util.UpdateA02;	//add by 2354 2004.12.17
import com.tradevan.util.UpdateA03;	//add by 2354 2004.12.24
import com.tradevan.util.UpdateA04;	//add by 2354 2004.12.17
import com.tradevan.util.UpdateA05;	//add by 2354 2004.12.17
import com.tradevan.util.UpdateB01;	//add by egg 2004.12.21
import com.tradevan.util.UpdateB03;	//add by egg 2004.12.21
import com.tradevan.util.UpdateM01;	//add by 2354 2004.12.17
import com.tradevan.util.UpdateM02;	//add by 2354 2004.12.17
import com.tradevan.util.UpdateM03;
import com.tradevan.util.UpdateM04;	//add by egg 2004.12.21
import com.tradevan.util.UpdateM05;
import com.tradevan.util.UpdateM06;	//add by egg 2004.12.21
import com.tradevan.util.UpdateM07;	//add by egg 2004.12.21
import com.tradevan.util.Utility; 
import java.util.Date;
import java.io.File;
import java.lang.Integer;
import java.util.Calendar;
import java.text.SimpleDateFormat;
import java.io.FileOutputStream;
import java.io.BufferedOutputStream;
import java.io.PrintStream;

public class AutoFileCheck {
	public static String FileCheck(String Report_no,String bank_type){
		String S_YEAR = "";
		String S_MONTH = "";
		String bank_code = "";
		String errMsg = "";
		File logfile;
		FileOutputStream logos=null;    	
		BufferedOutputStream logbos = null;
		PrintStream logps = null;
		Date nowlog = new Date();
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
		SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	    Calendar logcalendar;
	    File logDir = null;
	    String bank_name = "";
	    List paramList = new ArrayList();
		try{
			String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no;
			logcalendar = Calendar.getInstance(); 
		    nowlog = logcalendar.getTime();   
	        logDir  = new File(Utility.getProperties("logDir"));
	        if(!logDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("logDir"))){
     				System.out.println("目錄新增失敗");
     			}    
    		}
	        logfile = new File(logDir + System.getProperty("file.separator") + Report_no +"_"+bank_type+"."+ logfileformat.format(nowlog));			
			 
			logos = new FileOutputStream(logfile,true);  		        	   
			logbos = new BufferedOutputStream(logos);
			logps = new PrintStream(logbos);
			Date today = new Date();		
			int	batch_no = today.hashCode();
			System.out.println("batch_no="+batch_no);
			File tmpFile = new File(WMdataDir);
			String[] fileList = tmpFile.list();
			List checkFile = new LinkedList();
			paramList.add(bank_type);
			List dbData = DBManager.QueryDB_SQLParam("select * from cdshareno where cmuse_Div='001' and cmuse_id=?",paramList,"");   
			if(dbData != null && dbData.size() != 0){
				bank_name = (String)((DataObject)dbData.get(0)).getValue("cmuse_name");
			}
		    for(int i=0;i<fileList.length;i++){//把要check的file加到checkFile的list
		        File checkfile = new File(WMdataDir+System.getProperty("file.separator")+fileList[i]);    		    
			    if (!checkfile.isDirectory()){	    			   
			        System.out.println("filename="+WMdataDir+System.getProperty("file.separator")+fileList[i]);
			        if((!Report_no.equals("")) && ((fileList[i].substring(0,3)).equals(Report_no))){			          
					    System.out.println("bank_code==all");
				        List bn01Data  = DBManager.QueryDB_SQLParam("select bank_no, bank_name from bn01 where bank_type=? and bn_type <> '2' ORDER BY bank_no",paramList,"");	    			       					   
				        bank_codeLoop:
				        for(int j=0;j<bn01Data.size();j++){					       
				            if(Report_no.substring(0,1).equals("M") || (fileList[i].substring(3,10)).equals((String)((DataObject)bn01Data.get(j)).getValue("bank_no"))){	//modify by 2354 2004.12.17				                 
				               checkFile.add(fileList[i]);
					  	       System.out.println("add file "+fileList[i]);	
				               break bank_codeLoop;			                     
				            }//end of bank_no
				        }//end of for
			        }//enf of Report_no     
			    }//end of is file
		    }//end of for把要check的file加到checkFile的list	
		
		    for(int i=0;i<checkFile.size();i++){	
		    	errMsg = "";
		    	logcalendar = Calendar.getInstance(); 
			    nowlog = logcalendar.getTime();
			    logps.println(logformat.format(nowlog)+" "+"檢核報表:"+Report_no);			    
			    logps.println(logformat.format(nowlog)+" "+"機構類別:"+bank_type+" "+bank_name);
		    	logps.println(logformat.format(nowlog)+" "+(String)checkFile.get(i));	
			    logps.flush();
			    
		        Report_no=((String)checkFile.get(i)).substring(0,3);
		        if(Report_no.substring(0,1).equals("M")){	// add by 2354 MXX報表檔名: M01YYYMM
		    	    bank_code="9700002";
		    	    S_YEAR=((String)checkFile.get(i)).substring(3,6);
		    	    S_MONTH=((String)checkFile.get(i)).substring(6,8);
		        }else{
		    	    bank_code=((String)checkFile.get(i)).substring(3,10);
		    	    S_YEAR=((String)checkFile.get(i)).substring(10,13);
		    	    S_MONTH=((String)checkFile.get(i)).substring(13,15);
		        }

		        System.out.println("Report_no="+Report_no);
		        System.out.println("bank_code="+bank_code);
		        System.out.println("S_YEAR="+S_YEAR);
		        System.out.println("S_MONTH="+S_MONTH); 	   		    
			    System.out.println("begin check i="+i);
			    					
				//moidfy  and add by winnin 2004.11.17
				if(Report_no.equals("A01")){
					errMsg = UpdateA01.doParserReport_A01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    						    				    						     
					//doParserReport_A01(String report_no, String m_year,String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type)
				}else if(Report_no.equals("A02")){
    				errMsg = UpdateA02.doParserReport_A02(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A03")){
    				errMsg = UpdateA03.doParserReport_A03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A04")){
					errMsg = UpdateA04.doParserReport_A04(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A05")){
					errMsg = UpdateA05.doParserReport_A05(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A06")){//101.10.12 add by 2295
				    errMsg = UpdateA06.doParserReport_A06(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A08")){//101.10.12 add by 2295
                    errMsg = UpdateA08.doParserReport_A08(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("A09")){//101.10.12 add by 2295
                    errMsg = UpdateA09.doParserReport_A09(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);                    
				}else if(Report_no.equals("A10")){//101.10.12 add by 2295
                    errMsg = UpdateA10.doParserReport_A10(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);                    
				}else if(Report_no.equals("A99")){//101.10.12 add by 2295
                    errMsg = UpdateA99.doParserReport_A99(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);                    
				}else if(Report_no.equals("F01")){//101.10.12 add by 2295
                    errMsg = UpdateF01.doParserReport_F01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);                 
                }else if(Report_no.equals("B01")){
					errMsg = UpdateB01.doParserReport_B01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("B03")){
					errMsg = UpdateB03.doParserReport_B03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("M01")){
					errMsg = UpdateM01.doParserReport_M01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("M02")){
					errMsg = UpdateM02.doParserReport_M02(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("M03")){
					errMsg = UpdateM03.doParserReport_M03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
				}else if(Report_no.equals("M04")){
					errMsg = UpdateM04.doParserReport_M04(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    			}else if(Report_no.equals("M05")){
    				errMsg = UpdateM05.doParserReport_M05(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    			}else if(Report_no.equals("M06")){
    				errMsg = UpdateM06.doParserReport_M06(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    			}else if(Report_no.equals("M07")){
    				errMsg = UpdateM07.doParserReport_M07(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    	
    			}
    			if(errMsg.equals("")){//檢核成功
    			   logps.println(logformat.format(nowlog)+" "+"執行檢核完成");				   
    			}else{//檢核失敗
    			   logps.println(logformat.format(nowlog)+" "+"執行檢核失敗:"+errMsg);
    			}
			    logps.flush();
			    System.out.println("end check i="+i);							
		    }//end of for 執行檢核	   		
		    if(checkFile.size() == 0){
	    	   logcalendar = Calendar.getInstance(); 
			   nowlog = logcalendar.getTime();  			   
			   logps.println(logformat.format(nowlog)+" "+"檢核報表:"+Report_no);			    
			   logps.println(logformat.format(nowlog)+" "+"機構類別:"+bank_type+" "+bank_name);		    	
		       logps.println(logformat.format(nowlog)+" "+"無符合條件之上傳檔案以供檢核");	
			   logps.flush();		       		   
		    }	
		}catch(Exception e){
		   System.out.println("FileCheck Error:"+e.getMessage());
  		   logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();   
	       logps.println(logformat.format(nowlog)+" "+"FileCheck Error:"+e + "\n"+e.getMessage());	
		   logps.flush();		   
		}finally{
			try{
			   if (logos  != null) logos.close();
 	           if (logbos != null) logbos.close();
 	           if (logps  != null) logps.close();
			}catch(Exception ioe){
				System.out.println(ioe.getMessage());
		    }
		}
		return new String(logDir + System.getProperty("file.separator") + Report_no +"_"+bank_type+"."+ logfileformat.format(nowlog));
	}//end of FileCheck
	
}
