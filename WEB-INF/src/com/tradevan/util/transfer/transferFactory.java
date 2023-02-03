/* 99.09.01 create 金庫BIS報表轉檔 by 2295 
 * */
package com.tradevan.util.transfer;


import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.tradevan.util.*;
import com.tradevan.util.dao.DataObject;

public class transferFactory {
	 static String errMsg = "";
	 File agriBankDir = null;  
	 List dbData = null;        
	 StringBuffer sql = new StringBuffer();		
	 List paramList = new ArrayList();
	 FileInputStream finput = null;	
	 POIFSFileSystem fs = null;
	 HSSFWorkbook wb = null;
	 HSSFSheet sheet = null;//讀取工作表，宣告其為sheet
     HSSFRow row = null;//宣告一列
     HSSFCell cell = null;//宣告一個儲存格
     DataObject bean = null;	
     double cellValue = 0.0;
     int data_yy = 0;
     String cellString = "";
     List updateDBList = new LinkedList();	
     List updateDBSqlList = new LinkedList();//0:sql 1:data	
     
     List updateDBDataList = new LinkedList();//儲存參數的List
     List dataList = new LinkedList();//參數detail
     short[][] rptidx = null;
     
     static File logfile;
     static FileOutputStream logos=null;      
     static BufferedOutputStream logbos = null;
     static PrintStream logps = null;
     static Date nowlog = new Date();
     static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
     static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
     static Calendar logcalendar;
     static File logDir = null;
     short[][] rptidx_1_A1 = {
     		  {(short) 7, (short)0}/*1*/, {(short)11,(short)0}/*2*/, {(short)15,(short)0}/*3*/, {(short)22,(short)0}/*4*/,
			  {(short) 7, (short)1}/*5*/, {(short)11,(short)1}/*6*/, {(short)15,(short)1}/*7*/,
			  {(short) 7, (short)2}/*8*/, {(short)10,(short)2}/*9*/, {(short)12,(short)2}/*10*/,
			  {(short) 7, (short)3}/*11*/,{(short) 9,(short)3}/*12*/,{(short)11,(short)3}/*13*/,{(short)13,(short)3}/*14*/,
			  {(short)15, (short)3}/*15*/,{(short)17,(short)3}/*16*/,{(short)19,(short)3}/*17*/, 
			  {(short) 7, (short)4}/*18*/,{(short) 9,(short)4}/*19*/,{(short)11,(short)4}/*20*/,{(short)21,(short)4}/*21*/,
			  {(short) 7, (short)5}/*22*/,{(short) 9,(short)5}/*23*/,{(short)30,(short)4}/*24*/
	 };
     
     short[][] rptidx_1_B = {
   		      {(short) 7,(short)1}/*1*/, {(short) 8,(short)1}/*2*/, {(short) 9,(short)1}/*3*/, {(short)10,(short)1}/*4*/,
			  {(short)11,(short)1}/*5*/, {(short)12,(short)1}/*6*/, {(short)13,(short)1}/*7*/, {(short)14,(short)1}/*8*/, 
			  {(short)15,(short)1}/*9*/, {(short)16,(short)1}/*10*/,{(short)17,(short)1}/*11*/,{(short)18,(short)1}/*12*/,
			  {(short)19,(short)1}/*13*/,{(short)20,(short)1}/*14*/,{(short)22,(short)1}/*15*/,{(short)23,(short)1}/*16*/,
			  {(short)24,(short)1}/*17*/,{(short)25,(short)1}/*18*/,{(short)26,(short)1}/*19*/,{(short)27,(short)1}/*20*/,
			  {(short)28,(short)1}/*21*/,{(short)29,(short)1}/*22*/,{(short)30,(short)1}/*23*/,{(short)33,(short)1}/*24*/,
			  {(short)34,(short)1}/*25*/,{(short)35,(short)1}/*26*/,{(short)37,(short)1}/*27*/,{(short)38,(short)1}/*28*/,
			  {(short)39,(short)1}/*29*/,{(short)40,(short)1}/*30*/
	 };
     
     short[][] rptidx_1_B1 = {
     		  {(short) 6,(short)1}/*1*/, {(short) 7,(short)1}/*2*/, {(short)8,(short)1}/*3*/,{(short)9,(short)1}/*4*/,
			  {(short)10,(short)1}/*5*/, {(short) 6,(short)2}/*6*/, {(short)7,(short)2}/*7*/,{(short)8, (short)2}/*8*/,
			  {(short) 9,(short)2}/*9*/, {(short)10,(short)2}/*10*/
	 };
     short[][] rptidx_1_C = {
     		  {(short) 5, (short)6}/*1*/, {(short) 6,(short)6}/*2*/,  {(short) 7,(short)6}/*3*/, {(short) 8,(short)6}/*4*/,
			  {(short) 9, (short)6}/*5*/, {(short)10,(short)6}/*6*/,  {(short)12,(short)6}/*7*/, {(short)14,(short)6}/*8*/, 
			  {(short)16, (short)6}/*9*/, {(short)18,(short)6}/*10*/, {(short)20,(short)6}/*11*/,{(short)22,(short)6}/*12*/,
			  {(short)24, (short)6}/*13*/,{(short)26,(short)6}/*14*/, {(short)28,(short)6}/*15*/
	 };
     short[][] rptidx_2_A = {
   		  	  {(short) 6,(short)1}/*1*/, {(short) 7,(short)1}/*2*/, {(short) 8,(short)1}/*3*/,{(short) 9,(short)1}/*4*/,
			  {(short)10,(short)1}/*5*/, {(short)11,(short)1}/*6*/, {(short)12,(short)1}/*7*/,{(short)13,(short)1}/*8*/,
			  {(short)14,(short)1}/*9*/, {(short)15,(short)1}/*10*/
	 };
     short[][] rptidx_2_F = {
 		  	  {(short) 6,(short)3}/*1*/, {(short) 7,(short)3}/*2*/, {(short) 8,(short)3}/*3*/,{(short) 9,(short)3}/*4*/,
			  {(short)10,(short)3}/*5*/, {(short)11,(short)3}/*6*/, {(short)12,(short)3}/*7*/,{(short) 6,(short)4}/*8*/,
			  {(short) 7,(short)4}/*9*/, {(short) 8,(short)4}/*10*/,{(short) 9,(short)4}/*11*/,
			  {(short)10,(short)4}/*12*/,{(short)11,(short)4}/*12*/,{(short)12,(short)4}/*14*/
	 };
     short[][] rptidx_4_H = {
 		  	  {(short) 7,(short)1}/*1*/, {(short) 8,(short)1}/*2*/, {(short) 9,(short)1}/*3*/,{(short)10,(short)1}/*4*/,
			  {(short) 7,(short)2}/*5*/, {(short) 8,(short)2}/*6*/, {(short) 9,(short)2}/*7*/,{(short)10,(short)2}/*8*/
	 };
     
     short[][] rptidx_6_A = {
		  	  {(short) 8,(short)1}, {(short) 9,(short)1}, {(short)22,(short)1},
			  {(short) 8,(short)2}, {(short) 9,(short)2}, {(short)22,(short)2},
			  {(short) 8,(short)3}, {(short) 9,(short)3}, {(short)22,(short)3},
			  {(short) 8,(short)4}, {(short) 9,(short)4},
			  {(short) 8,(short)5}, {(short) 9,(short)5}, {(short)22,(short)5},
			  {(short) 8,(short)6}, {(short) 9,(short)6}, {(short)22,(short)6},
			  {(short)27,(short)5}, {(short)29,(short)5}
	 };
     /* 6-A2-a(NTD/USD)其他各別獨立部份*/
     short[][] rptidx_6_A2_a_extra = {
		  	  {(short)10,(short)11}, {(short)13,(short)11}, {(short)17,(short)11},
			  {(short)10,(short)12}, {(short)13,(short)12}, {(short)17,(short)12},
			  {(short)10,(short)13}, {(short)15,(short)14}, {(short)10,(short)15},
			  {(short)27,(short) 8}
	 };
     short[][] rptidx_6_A1 = {     		  
     		  /*資產市價*/
		  	  {(short) 8,(short)4},{(short)11,(short)4},{(short)12,(short)4},{(short)15,(short)4},
			  {(short)16,(short)4},{(short)17,(short)4},{(short)18,(short)4},{(short)19,(short)4},
			  {(short)20,(short)4},{(short)21,(short)4},{(short)22,(short)4},{(short)23,(short)4},
			  {(short)24,(short)4},{(short)25,(short)4},{(short)26,(short)4},{(short)36,(short)4},
			  {(short)37,(short)4},{(short)38,(short)4},{(short)39,(short)4},{(short)40,(short)4},
			  {(short)41,(short)4},{(short)42,(short)4},{(short)43,(short)4},{(short)44,(short)4},
			  {(short)45,(short)4},{(short)46,(short)4},{(short)47,(short)4},{(short)48,(short)4},
			  {(short)49,(short)4},{(short)50,(short)4},{(short)51,(short)4},{(short)52,(short)4},
			  {(short)53,(short)4},{(short)54,(short)4},
			  {(short)56,(short)4},{(short)57,(short)4},{(short)58,(short)4},{(short)59,(short)4},
			  /*資本計提*/
			  {(short) 8,(short)5},{(short)11,(short)5},{(short)12,(short)5},{(short)15,(short)5},
              {(short)16,(short)5},{(short)17,(short)5},{(short)18,(short)5},{(short)19,(short)5},
              {(short)20,(short)5},{(short)21,(short)5},{(short)22,(short)5},{(short)23,(short)5},
              {(short)24,(short)5},{(short)25,(short)5},{(short)26,(short)5},{(short)36,(short)5},
              {(short)37,(short)5},{(short)38,(short)5},{(short)39,(short)5},{(short)40,(short)5},
              {(short)41,(short)5},{(short)42,(short)5},{(short)43,(short)5},{(short)44,(short)5},
              {(short)45,(short)5},{(short)46,(short)5},{(short)47,(short)5},{(short)48,(short)5},
              {(short)49,(short)5},{(short)50,(short)5},{(short)51,(short)5},{(short)52,(short)5},
              {(short)53,(short)5},{(short)54,(short)5},
              {(short)56,(short)5},{(short)57,(short)5},{(short)58,(short)5},{(short)59,(short)5},
			  /*資本扣除*/
			  {(short)41,(short)6},{(short)43,(short)6},{(short)44,(short)6},
			  {(short)49,(short)6},{(short)51,(short)6},{(short)52,(short)6},
			  {(short)53,(short)6},{(short)54,(short)6},{(short)59,(short)6}
			  
	 };
     
     short[][] rptidx_6_B = {
		  	  {(short) 8,(short)2}/*1*/, {(short) 8,(short)4}/*2*/, {(short) 8,(short)5}/*3*/,
			  {(short) 8,(short)6}/*4*/, {(short)29,(short)5}/*5*/, {(short)31,(short)5}/*6*/
	 }; 
     
     short[][] rptidx_6_B1= {
		  	  {(short)15,(short)2}, {(short)16,(short)2}, {(short)17,(short)2},{(short)31,(short)2},
			  {(short)15,(short)3}, {(short)16,(short)3}, {(short)17,(short)3},{(short)31,(short)3},
			  {(short)15,(short)4}, {(short)16,(short)4}, {(short)17,(short)4},{(short)31,(short)4},
			  {(short)15,(short)5}, {(short)16,(short)5}, {(short)17,(short)5},{(short)31,(short)5},
			  {(short)15,(short)6}, {(short)16,(short)6}, {(short)17,(short)6},{(short)31,(short)6},
			  {(short)13,(short)7}, {(short)31,(short)7}
	 };
     
     short[][] rptidx_6_B2 = {
		  	  {(short) 8,(short)0}/*1*/, {(short) 8,(short)3}/*2*/, {(short) 8,(short)6}/*3*/,
			  {(short)15,(short)1}/*4*/, {(short)15,(short)7}/*5*/
	 };
     short[][] rptidx_6_C = {
		  	  {(short) 7,(short)1}/*1*/, {(short) 8,(short)1}/*2*/, {(short) 9,(short)1}/*3*/,
			  {(short)10,(short)1}/*4*/, {(short)12,(short)1}/*5*/, {(short)16,(short)1}/*6*/,
			  {(short)17,(short)1}/*7*/
	 }; 
     short[][] rptidx_6_C2 = {
		  	  {(short) 4,(short)0},/*幣別*/
			  /*長部位*/
			  {(short) 9,(short)1}, {(short)10,(short)1}, {(short)11,(short)1}, {(short)12,(short)1}, 
			  {(short)13,(short)1}, {(short)14,(short)1}, {(short)19,(short)1},
			  /*短部位*/
			  {(short) 9,(short)2}, {(short)10,(short)2}, {(short)11,(short)2}, {(short)12,(short)2}, 
			  {(short)13,(short)2}, {(short)14,(short)2}, {(short)19,(short)2}
	 }; 
     short[][] rptidx_6_E = {
     		  /*單一部位*/
		  	  {(short) 5,(short)3},/*幣別*/			 
			  {(short) 7,(short)1}, {(short) 7,(short)2}, {(short) 7,(short)3}, {(short) 7,(short)4},{(short) 7,(short)5}, 
			  {(short) 8,(short)1}, {(short) 8,(short)2}, {(short) 8,(short)3}, {(short) 8,(short)4},{(short) 8,(short)5}, 
			  {(short) 9,(short)1}, {(short) 9,(short)2}, {(short) 9,(short)3}, {(short) 9,(short)4},{(short) 9,(short)5}, 
			  {(short)10,(short)1}, {(short)10,(short)2}, {(short)10,(short)3}, {(short)10,(short)4},{(short)10,(short)5}, 
			  {(short)13,(short)1}, {(short)13,(short)2}, {(short)13,(short)3}, {(short)13,(short)4},{(short)13,(short)5}, 
			  /*避險部位*/
			  {(short)19,(short)3},/*幣別*/			 
			  {(short)21,(short)1}, {(short)21,(short)2}, {(short)21,(short)3}, {(short)21,(short)4},{(short)21,(short)5}, 
			  {(short)22,(short)1}, {(short)22,(short)2}, {(short)22,(short)3}, {(short)22,(short)4},{(short)22,(short)5}, 
			  {(short)23,(short)1}, {(short)23,(short)2}, {(short)23,(short)3}, {(short)23,(short)4},{(short)23,(short)5}, 
			  {(short)24,(short)1}, {(short)24,(short)2}, {(short)24,(short)3}, {(short)24,(short)4},{(short)24,(short)5}, 
			  {(short)27,(short)1}, {(short)27,(short)2}, {(short)27,(short)3}, {(short)27,(short)4},{(short)27,(short)5}, 
			  /*資本計提總額*/
			  {(short)33,(short)4}
	 }; 
     short[][] rptidx_6_G = {
		  	  {(short) 6,(short)1}/*1*/, {(short) 7,(short)1}/*2*/, {(short) 8,(short)1}/*3*/,{(short)9,(short)1}/*4*/, 
		  	  {(short) 6,(short)2}/*5*/, {(short) 7,(short)2}/*6*/, {(short) 8,(short)2}/*7*/,{(short)9,(short)2}/*8*/			  
	 };
     
     /*101.09.17舊格式101年度以前
     short[][] rptidx_4_A_1  = { {(short) 7,(short)22}//rowNum.start/end , 
             					 {(short) 3,(short)11}//cellNum.start/end };
     */
     short[][] rptidx_4_A_1  = { {(short) 7,(short)29}/* rowNum.start/end */, 
                                 {(short) 4,(short)12}/* cellNum.start/end */ };
     short[][] rptidx_2_B    = { {(short) 6,(short)58}/* rowNum.start/end */, 
							     {(short) 2,(short) 5}/* cellNum.start/end */ };
     short[][] rptidx_2_C    = { {(short) 7,(short)51}/* rowNum.start/end */, 
			   				     {(short) 2,(short)10}/* cellNum.start/end */ };
     short[][] rptidx_2_D    = { {(short) 7,(short)43}/* rowNum.start/end */, 
			   				     {(short) 2,(short) 8}/* cellNum.start/end */ };
     short[][] rptidx_2_E    = { {(short) 7,(short)44}/* rowNum.start/end */, 
			   					 {(short) 2,(short) 8}/* cellNum.start/end */ };
     short[][] rptidx_2_E1   = { {(short) 8,(short)45}/* rowNum.start/end */, 
     		                     {(short) 2,(short) 4}/* cellNum.start/end */ };
     short[][] rptidx_2_E2   = { {(short) 7,(short)43}/* rowNum.start/end */, 
             				     {(short) 3,(short) 5}/* cellNum.start/end */ };
     short[][] rptidx_4_D_1  = { {(short) 7,(short)11}/* rowNum.start/end */, 
			    			     {(short) 3,(short) 4}/* cellNum.start/end */ };
     short[][] rptidx_4_D_2  = { {(short) 7,(short)11}/* rowNum.start/end */, 
     					         {(short) 3,(short) 4}/* cellNum.start/end */ };
     short[][] rptidx_5_A    = { {(short) 5,(short)17}/* rowNum.start/end */, 
		        			     {(short) 1,(short) 3}/* cellNum.start/end */ };
     short[][] rptidx_6_A2_a = { {(short) 8,(short)24}/* rowNum.start/end */, 
     						     {(short) 5,(short)10}/* cellNum.start/end */ };
    public boolean InsertZero(String m_year,String m_month,String report_no){
   		List dbData = null;        
   		StringBuffer sql = new StringBuffer();
   		List paramList = new ArrayList();
		
   		List updateDBList = new LinkedList();	
        List updateDBSqlList = new LinkedList();//0:sql 1:data	
        
        List updateDBDataList = new LinkedList();//儲存參數的List
        List dataList = new LinkedList();//參數detail   		
   		
   		boolean updateOK=false;
   		
   		DataObject bean = null; 
   		try{
   			sql.append(" select count(*) as data from AgriBank_rpt where m_year = ? and m_month=? and acc_tr_type = ? and data_extra = ?");
		    paramList.add(m_year);
		    paramList.add(m_month);
		    paramList.add(report_no);
		    paramList.add(String.valueOf(data_yy));
		    dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"data");     			      
	 		System.out.println("dbData.size="+dbData.size());
	 		bean = (DataObject)dbData.get(0);
		    if(Integer.parseInt(bean.getValue("data").toString()) == 0){
		       sql = new StringBuffer();
		       sql.append(" insert into agribank_rpt ");
		       sql.append(" select ?,?,'0180012',acc_code,?,0,? from ncacno where acc_tr_type = ? ");
		       if(report_no.equals("6-C1")){
		       	  //if(data_yy == 999 ){
		       	  //   sql.append(" and acc_code in ('001000','110000','120000','130000','140000')");
		       	  //}else{	
		       	     sql.append(" and acc_code in ('001000','111000','121000','131000','141000')");
		       	  //}
		       }
		       
		       if(report_no.equals("6-A")){
		       	  if(data_yy == 999 ){
		       	     sql.append(" and acc_code in ('001000','150000','160000','170000','180000')");
		       	  }else{	
		       	     sql.append(" and acc_code in ('001000','111000','121000','131000','141000','151000','161000')");
		       	  }
		       }
		       
		       if(report_no.equals("6-B")){
		       	  if(data_yy == 999 ){
		       	     sql.append(" and acc_code in ('001000','111000','121000','131000','141000','150000','160000')");
		       	  }else{	
		       	     sql.append(" and acc_code in ('001000','111000','121000','131000','141000')");
		       	  }
		       }
	
		       sql.append(" order by acc_range");
		       updateDBSqlList.add(sql.toString());
		       dataList.add(m_year);
		       dataList.add(m_month);
		       dataList.add(report_no);
		       dataList.add(String.valueOf(data_yy));
		       dataList.add(report_no);		       
	           updateDBDataList.add(dataList);
	           updateDBSqlList.add(updateDBDataList);
		       updateDBList.add(updateDBSqlList);
		       updateOK = DBManager.updateDB_ps(updateDBList);
		       if(updateOK){
		           System.out.println("AgriBank Insert Zero OK");	
		       }
		    }else{
		       System.out.println("AgriBank not need Insert Zero");
		       updateOK = true;
		    }
			if(!updateOK){
				errMsg = errMsg + "AgriBank Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
   		}catch(Exception e){
   			errMsg = errMsg + "doParserReport.InsertZeroAgriBank Error:"+e.getMessage()+"<br>";
   			System.out.println("doParserReport.InsertZeroAgriBank Error:"+e.getMessage());
   		}
   		return updateOK;
   }  
   
   public static void printLog(PrintStream logps,String errRptMsg){
        if(!errRptMsg.equals("")){
           logcalendar = Calendar.getInstance(); 
           nowlog = logcalendar.getTime();
           logps.println(logformat.format(nowlog)+errRptMsg);
           logps.flush();
        }
   }
}

