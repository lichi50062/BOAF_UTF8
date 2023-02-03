/*  99.09.01 create 金庫BIS報表轉檔(1-A1/1-B/1-B1/1-C/2-A/2-F/4-H/4-A-1/2-B) by 2295 
 *  99.12.28 fix 金庫資料金額單位皆為仟元
 *           add 1-A1 026000普通股占風險性資產比率=1-B(001000)普通股/1-A(004000)風險性資產合計 by 2295
 * 100.06.01-09 add BIS報表轉檔(R0521/6-A1(NTD)/6-A1(USD)/6-B1/6-B2/6-G/2-C/2-D/2-E/2-E1/2-E2 by 2295
 * 100.06.10-14 add BIS報表轉檔(4-D-1/4-D-2/5-A/5-A/6-A by 2295
 * 100.06.15-17	add BIS報表轉檔(6-A2-a(NTD)/6-A2-a(USD)/6-B/6-C/6-C1/6-C2/6-E) by 2295
 * 100.07.28 fix 6-A/6-C1 轉檔規則 by 2295  
 * 100.08.03 fix 6-B 轉檔規則 by 2295
 * 100.12.16 fix 6-A2-a(NTD)/(USD)轉檔規則 by 2295
 * 101.09.17 add 4-A-1增修欄位 by 2295
 * 101.10.04 add 6-A1(NTD)增修欄位 by 2295
 * 101.10.05 add 6-A1(USD)增修欄位 by 2295
 * 101.10.11 add 4-D-1增修欄位 by 2295
 * 101.10.12 add 4-D-2增修欄位 by 2295
 */
package com.tradevan.util.transfer;


import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.tradevan.util.*;
import com.tradevan.util.dao.DataObject;

public class parseAgriBankRpt extends transferFactory {
	 
	 public static void main(String args[]) {      
	 	parseAgriBankRpt a = new parseAgriBankRpt();
	    System.out.println("read data="+a.doParserRpt("101","06","4-D-2"));	    
	 } 
	  
	
		
   public String doParserRpt(String m_year,String m_month,String report_no){
   	    short[][] rptidx = null;
   	    short cellNum_start = 0;
   	    short cellNum_end = 0;
   	    short rowNum_start = 0;
   	    short rowNum_end = 0;
		try {
			sql.delete(0,sql.length());
			paramList = new ArrayList();
			updateDBList = new LinkedList();	
		    updateDBSqlList = new LinkedList();//0:sql 1:data
		    updateDBDataList = new LinkedList();//儲存參數的List
		    dataList = new LinkedList();//參數detail
		    
			agriBankDir = new File(Utility.getProperties("AgriBank_transferDir"));
			if (!agriBankDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("AgriBank_transferDir"))) {
					errMsg += Utility.getProperties("AgriBank_transferDir") + "目錄新增失敗";
				}
			}
            
			logDir  = new File(Utility.getProperties("logDir"));
	        if(!logDir.exists()){
	            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
	               System.out.println("目錄新增失敗");
	            }    
	        }
		    logfile = new File(logDir + System.getProperty("file.separator") + "parseAgriBank_"+report_no+"."+ logfileformat.format(nowlog));                       
		    System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseAgriBank_"+report_no+"."+ logfileformat.format(nowlog));
		    logos = new FileOutputStream(logfile,true);                         
		    logbos = new BufferedOutputStream(logos);
		    logps = new PrintStream(logbos);   
		    			
			finput = new FileInputStream(agriBankDir+ System.getProperty("file.separator") + "BIS_"+m_year+m_month+".xls");	
			fs = new POIFSFileSystem(finput);
			wb = new HSSFWorkbook(fs);
			System.out.println("report_no="+report_no);
			System.out.println("wb.size()="+wb.getNumberOfSheets());
					
			sheetLoop:
			for(int i=0;i<wb.getNumberOfSheets();i++){
			    sheet = wb.getSheetAt(i);//讀取工作表，宣告其為sheet	
			    System.out.println("now.sheet="+wb.getSheetName(i));
			    if(wb.getSheetName(i).equals(report_no)){
			       printLog(logps,"讀取工作表["+report_no+"]成功");
				   System.out.println("now parser "+report_no);
				   System.out.println("讀取工作表["+report_no+"]成功");
				   break sheetLoop;
			    }
			}  
			
			rptidx = getRptIdx(report_no);
			
			//100.06.14 fix 5-A表.每份表格有3份資料,先不塞入0值
			if(!report_no.equals("5-A") && !report_no.equals("6-A") && !report_no.equals("6-C1") && InsertZero(m_year,m_month,report_no)){
				printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
			}
						
     		sql.append(" select acc_code,acc_name from ncacno where acc_tr_type=? ");
     		if(report_no.equals("4-A-1")){//小計欄位,不需轉檔		
     		   sql.append(" and acc_code not in (");
     		   sql.append(" '111000','112000','115000','116000','113000','121000','122000','125000','126000','123000','131000','132000','135000','136000','133000',");
     		   sql.append("	'141000','142000','145000','146000','143000','151000','152000','155000','156000','153000','161000','162000','165000','166000','163000',");
     		   sql.append("	'171000','172000','175000','176000','173000','181000','182000','185000','186000','183000','191000','192000','195000','196000','193000'");
     		   sql.append(" )");
     		}
     		sql.append(" order by acc_range ");
     		paramList.add(report_no);
     		
     		dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");     			      
	 		System.out.println("dbData.size="+dbData.size());
	 		cellValue = 0.0; 		
	 		
	 		sql.delete(0,sql.length());
	 		sql.append(" update agribank_rpt set amt = ?");
        	sql.append(" where m_year = ? and m_month= ?");        	
        	sql.append("  and bank_code='0180012'");
        	sql.append("  and acc_code =? ");
        	sql.append("  and acc_tr_type =? ");
        	sql.append("  and data_extra = ? ");
        	updateDBSqlList.add(sql.toString());
        	
        	double cellValue_4 = 0.0;
        	double cellValue_18 = 0.0;
        	double cellValue_tmp1 = 0.0;
        	double cellValue_tmp2 = 0.0;
        	short rowNum=7;
        	if(report_no.equals("4-A-1") || report_no.equals("2-B")         || report_no.equals("2-C")   ||
        	   report_no.equals("2-D")   || report_no.equals("2-E")         || report_no.equals("2-E1")  ||
        	   report_no.equals("2-E2")  || report_no.equals("4-D-1")       || report_no.equals("4-D-2") || 
        	   report_no.equals("5-A")   || report_no.equals("6-A2-a(NTD)") || report_no.equals("6-A2-a(USD)") 
			  )
        	{    
        		int dataCount = 0;
            	rowNum_start = rptidx[0][0];
            	rowNum_end = rptidx[0][1];
            	cellNum_start = rptidx[1][0];
            	cellNum_end = rptidx[1][1];
            	//if(report_no.equals("6-A2-a(NTD)") || report_no.equals("6-A2-a(USD)")){
            	//	dataCount = 10;//100.12.16 fix
            	//}
                for(short cellNum=cellNum_start;cellNum<=cellNum_end;cellNum++){
              	   for(rowNum=rowNum_start;rowNum<=rowNum_end;rowNum++){  
              	       if(report_no.equals("4-A-1")){
              	   	       //if(cellNum <= 10 && (rowNum==12 || rowNum==17)) continue;
              	   	       //if(cellNum == 11 && (rowNum==11 || rowNum==16)) continue;//資本扣除金額
              	       }
              	       if(report_no.equals("5-A")){
              	       	  if(rowNum == 16) continue;              	       	  
              	       	  if(dataCount == 12){
              	       	  	dataCount = 0;
              	       	  	System.out.print("cellNum="+cellNum);
              	       	    System.out.println(":rowNum="+rowNum);
              	       	  }
              	       	  
              	       }
              	       if((report_no.equals("6-A2-a(NTD)") || report_no.equals("6-A2-a(USD)")) && rowNum == 23) continue;
              	       
              	   	   System.out.print("dataCount="+dataCount);              	   	   
              	       bean = (DataObject)dbData.get(dataCount);        	       	   
              	       //System.out.println("acc_code="+(String)bean.getValue("acc_code"));
              	       cellValue = 0.0;
                   	   row = sheet.getRow(rowNum);
                   	   if(report_no.equals("5-A") && rowNum == 17){
                   	      cell = row.getCell((short)1);
                   	   }else{
      			          cell = row.getCell(cellNum);
                   	   }
      			       System.out.print(":row="+rowNum+":cell="+cellNum);
      			       //System.out.print(":celltype="+cell.getCellType());
      			       if(cell != null){
      			          if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){  	
      			             System.out.print(":CELL_TYPE_STRING");
      			             if(report_no.equals("5-A") && rowNum == 5){
      			             	cellString = cell.getStringCellValue();
      			             	if(cellString.indexOf("年度") != -1){    
      			             		data_yy = Integer.parseInt(cellString.substring(0,cellString.indexOf("年度")).trim());
      			             		cellValue = Double.parseDouble(cellString.substring(0,cellString.indexOf("年度")).trim());
      			             		if(InsertZero(m_year,m_month,report_no)){      			             			
      			      				    printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
      			      			    }      			             		
      			             	}
      			             }else{
      			             	if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
      			             		cellValue = Double.parseDouble(cell.getStringCellValue());
      			             	}   
      			             }
      			          }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){
      			              System.out.print(":CELL_TYPE_NUMERIC"); 
      			              cellValue = cell.getNumericCellValue();			           	  
      			          }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){
      			              System.out.print(":CELL_TYPE_FORMULA");
      			              errMsg += "無法讀取公式內容";
                       	      return errMsg;
                            }
      			       }//end of cell != null
      			       System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
      			       if(report_no.equals("5-A") && rowNum == 5){//年度別    			          
    			          updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));
    			       }else{
    			       	  updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
    			       }
      			           			      
      	               printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));
      			       dataCount++;      			       
              	   }//end of rowNum 
              	   rowNum = rowNum_start;
                }//end of cellNum
                //其他獨立cell部份
                if(report_no.equals("6-A2-a(NTD)") || report_no.equals("6-A2-a(USD)")){
                   rptidx = rptidx_6_A2_a_extra;                   
                   dataLoop:
    			   for(int i=0;i<=dbData.size()-1;i++){	    			   	   
   			    	   if(i==10) break dataLoop;
   			     	   cellValue = 0.0; 
   			           bean = (DataObject)dbData.get(dataCount);//100.12.16 fix 
   			           row = sheet.getRow(rptidx[i][0]);
   			           System.out.print("i="+i+":row="+rptidx[i][0]);
   			           cell = row.getCell(rptidx[i][1]);
   			           System.out.print(":cell="+rptidx[i][1]);
   			           System.out.print(":celltype="+cell.getCellType());
   			           if(cell != null){
   	                      if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){	
   	                   	  //System.out.println("CELL_TYPE_STRING");	                	  	  
   	                   	  if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
   	                    	     cellValue = Double.parseDouble(cell.getStringCellValue());	                	
   	                    	  }
   	                      }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){	                 		
   	                    	  cellValue = cell.getNumericCellValue();				   
   	                      }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){	                 		
   	                   	     errMsg += "無法讀取公式內容";
   	                   	     return errMsg;
   	                      }
   			           }//end of cell != null
   	                   
   	                   System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
   	                   updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
   	                  
   	                   printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));
   	                   dataCount++;
   			       }//end of for    
                }//end of 6-A2-a(NTD/USD)
        	}else if (report_no.equals("R0521")){
        	   int dataCount = 0;           	   
           	   for(rowNum=8;rowNum<=105;rowNum++){  
           	   	 if(rowNum == 19 || rowNum == 20 || rowNum == 22 || rowNum == 23 || rowNum == 74 || 
           	   	 	rowNum == 75 || rowNum == 97 || rowNum == 98 || rowNum == 99 || rowNum == 100){
           	   	 	continue;
           	   	 }
           	     for(short cellNum=2;cellNum<=5;cellNum++){
           	     	   if(rowNum >= 101 && rowNum <= 105){
           	     	   	  if(cellNum >= 4) continue;
           	     	   }
           	       	   bean = (DataObject)dbData.get(dataCount);        	       	   
           	           //System.out.println("acc_code="+(String)bean.getValue("acc_code"));
           	           cellValue = 0.0;
                	   row = sheet.getRow(rowNum);
   			           cell = row.getCell(cellNum);
   			           System.out.println("row="+rowNum+";cell="+cellNum);
   			           System.out.println("celltype="+cell.getCellType());
   			           if(cell != null){
   			             if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
   			                if(cell.getStringCellValue().indexOf("$") == -1){	 
   			           	       System.out.println("CELL_TYPE_STRING");
   			           	       cellValue = Double.parseDouble(cell.getStringCellValue());
   			                }   
   			             }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){
   			           	    System.out.println("CELL_TYPE_NUMERIC"); 
   			           	    cellValue = cell.getNumericCellValue();			           	  
   			             }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){
   			           	    System.out.println("CELL_TYPE_FORMULA");
   			           	    errMsg += "無法讀取公式內容";
                    	    return errMsg;
                         }
   			           }//end of cell != null
   			           System.out.println("data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
   			           updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
	                  
   	                   printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));
   			           dataCount++;
           	   	 }//end of cellNum            	       
           	   }//end of rowNum
        	}else if(report_no.equals("6-A")){
        		int dataCount = 0;   
        		rowLoop:
            	for(rowNum=8;rowNum<=29;rowNum++){
            		if(rowNum >=23 && rowNum <=26) continue;
            		if(rowNum == 28) continue;
            		cellLoop:
            	    for(short cellNum=0;cellNum<=6;cellNum++){
            	     	bean = (DataObject)dbData.get(dataCount);  
            	     	if(rowNum==22 && (cellNum > 0 && cellNum < 5)) continue;//匯率沒有合計欄
     			        if(rowNum==27 && cellNum != 5) continue;//(27,5)有值
     			        if(rowNum==29 && cellNum != 5) continue;//(29,5)有值
            	        //System.out.println("acc_code="+(String)bean.getValue("acc_code"));
            	        cellValue = 0.0;
                 	    row = sheet.getRow(rowNum);
    			        cell = row.getCell(cellNum);
    			        System.out.print("row="+rowNum+":cell="+cellNum);
    			        System.out.print(":celltype="+cell.getCellType());
    			       
    			        if(cell != null){
    			           if(cell.getCellType()==HSSFCell.CELL_TYPE_BLANK){
    			           	  dataCount = 0;
    			           	  break cellLoop;
    			           }else if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
    			           	 System.out.print(":CELL_TYPE_STRING");
    			           	 if(cellNum == 0){
      			             	cellString = cell.getStringCellValue();
      			             	if(cellString.indexOf("合計") == -1){    
      			             	   data_yy = Integer.parseInt(getUnit(cellString.trim()));
      			             	   cellValue = Double.parseDouble(getUnit(cellString.trim()));
      			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
  			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
  			      			       } 
      			             	}else{
      			             	   data_yy = 999;
       			             	   cellValue = (double)999;
       			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
   			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
   			      			       } 
       			             	  
      			             	}
      			             }else{
      			             	if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
      			             		cellValue = Double.parseDouble(cell.getStringCellValue());
      			             	}   
      			             }
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){
    			           	    System.out.print(":CELL_TYPE_NUMERIC"); 
    			           	    cellValue = cell.getNumericCellValue();			           	  
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){
    			           	    System.out.print(":CELL_TYPE_FORMULA");
    			           	    errMsg += "無法讀取公式內容";
                     	        return errMsg;
                            }
    			        }//end of cell != null
    			        System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
    			        if(cellNum == 0){//幣別    			          
      			          updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));
      			          printLog(logps,(String)bean.getValue("acc_name")+"="+getDollar(String.valueOf(cellValue)));
      			        }else{      			          
      			       	  updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
      			       	  printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));    	                
      			        }
    			         	                   
    	                if(data_yy == 999 && cellValue == 999){
    	                   dataCount = dataCount + 7;
    	                }else{
    			           dataCount++;
    	                }
    			        if(cellNum == 6 && data_yy != 999) dataCount = 0;
    			        if("180000".equals((String)bean.getValue("acc_code"))){
    			        	dataCount = 0;
    			        }
    			        System.out.println("dataCount="+dataCount);
            	 }//end of cellNum            	       
               }//end of rowNum   
        	}else if(report_no.equals("6-B")){
        		int dataCount = 0;   
        		rowLoop:
            	for(rowNum=8;rowNum<=31;rowNum++){
            		if(rowNum >=25 && rowNum <=28) continue;
            		if(rowNum == 30) continue;
            		cellLoop:
            	    for(short cellNum=0;cellNum<=6;cellNum++){
            	     	bean = (DataObject)dbData.get(dataCount);  
            	     	if(cellNum == 1 || cellNum == 3) continue;
     			        if(rowNum==29 && cellNum != 5) continue;//(29,5)有值
     			        if(rowNum==31 && cellNum != 5) continue;//(31,5)有值
            	        //System.out.println("acc_code="+(String)bean.getValue("acc_code"));
            	        cellValue = 0.0;
                 	    row = sheet.getRow(rowNum);
                 	    cell = row.getCell(cellNum);
                 	        			        
    			        System.out.print("row="+rowNum+":cell="+cellNum);
    			        System.out.print(":celltype="+cell.getCellType());
    			       
    			        if(cell != null){    			        	
    			           if(cell.getCellType()==HSSFCell.CELL_TYPE_BLANK){
    			           	  dataCount = 0;
    			           	  break cellLoop;
    			           }else if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
    			           	 System.out.print(":CELL_TYPE_STRING");
    			           	 if(cellNum == 0){
      			             	cellString = cell.getStringCellValue();
      			             	if(cellString.indexOf("合計") == -1){    
      			             	   data_yy = Integer.parseInt(getCountry(cellString.trim(),"0"));
      			             	   cellValue = Double.parseDouble(getCountry(cellString.trim(),"0"));//取得國別代碼
      			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
  			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
  			      			       } 
      			             	}else{
      			             	   data_yy = 999;
       			             	   cellValue = (double)999;
       			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
   			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
   			      			       } 
       			             	  
      			             	}
      			             }else{
      			             	if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
      			             		cellValue = Double.parseDouble(cell.getStringCellValue());
      			             	}   
      			             }
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){
    			           	    System.out.print(":CELL_TYPE_NUMERIC"); 
    			           	    cellValue = cell.getNumericCellValue();			           	  
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){
    			           	    System.out.print(":CELL_TYPE_FORMULA");
    			           	    errMsg += "無法讀取公式內容";
                     	        return errMsg;
                            }
    			        }//end of cell != null
    			        System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
    			        if(cellNum == 0){//國家別    			          
      			          updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));
      			          printLog(logps,(String)bean.getValue("acc_name")+"="+getDollar(String.valueOf(cellValue)));
      			        }else{      			          
      			       	  updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
      			       	  printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));    	                
      			        }
    			         	                   
    			        dataCount++;    	                
    			        if(cellNum == 6 && data_yy != 999) dataCount = 0;
    			        if("160000".equals((String)bean.getValue("acc_code"))){
    			        	dataCount = 0;
    			        }
    			        System.out.println("dataCount="+dataCount);
            	 }//end of cellNum            	       
               }//end of rowNum   		
        	}else if(report_no.equals("6-C1")){
        		int dataCount = 0;   
        		rowLoop:
            	for(rowNum=6;rowNum<=16;rowNum++){
            		cellLoop:
            	    for(short cellNum=0;cellNum<=4;cellNum++){
            	     	bean = (DataObject)dbData.get(dataCount);        	       	   
            	        //System.out.println("acc_code="+(String)bean.getValue("acc_code"));
            	        cellValue = 0.0;
                 	    row = sheet.getRow(rowNum);
    			        cell = row.getCell(cellNum);
    			        System.out.print("row="+rowNum+":cell="+cellNum);
    			        System.out.print(":celltype="+cell.getCellType());
    			        if(cell != null){
    			           if(cell.getCellType()==HSSFCell.CELL_TYPE_BLANK){
    			           	  dataCount = 0;
    			           	  break cellLoop;
    			           }else if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
    			           	 System.out.print(":CELL_TYPE_STRING");
    			           	 if(cellNum == 0){
      			             	cellString = cell.getStringCellValue();
      			             	if(cellString.indexOf("合計") == -1){    
      			             	   data_yy = Integer.parseInt(getUnit(cellString.trim()));
      			             	   cellValue = Double.parseDouble(getUnit(cellString.trim()));
      			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
  			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
  			      			       } 
      			             	}else{
      			             	   data_yy = 999;
       			             	   cellValue = (double)999;
       			             	   if(InsertZero(m_year,m_month,report_no)){      			             			
   			      				      printLog(logps,m_year+"年"+m_month+"月["+report_no+"]-Insert Zero ok");
   			      			       } 
       			             	  
      			             	}
      			             }else{
      			             	if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
      			             		cellValue = Double.parseDouble(cell.getStringCellValue());
      			             	}   
      			             }
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){
    			           	    System.out.print(":CELL_TYPE_NUMERIC"); 
    			           	    cellValue = cell.getNumericCellValue();			           	  
    			           }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){
    			           	    System.out.print(":CELL_TYPE_FORMULA");
    			           	    errMsg += "無法讀取公式內容";
                     	        return errMsg;
                            }
    			        }//end of cell != null
    			        System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
    			        if(cellNum == 0){//幣別    			          
      			          updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));      			        
						  printLog(logps,(String)bean.getValue("acc_name")+"="+getDollar(String.valueOf(cellValue)));
      			        }else{
      			       	  updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
      			       	  printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));  
      			        }
    	                
    			        dataCount++;    	                
    			        if(cellNum == 4) dataCount = 0;
            	 }//end of cellNum            	       
               }//end of rowNum
		    }else{//1-A1/1-B/1-B1/1-C/2-A/2-F/4-H/6-A1(NTD/USD)/6-B1/6-C/6-C2/6-E
        	     dataLoop:
			     for(int i=0;i<=dbData.size()-1;i++){		
			     	cellValue = 0.0;
			     	if(i==24 && report_no.equals("1-A1")) break dataLoop;
			        bean = (DataObject)dbData.get(i);
			        row = sheet.getRow(rptidx[i][0]);
			        System.out.print("i="+i+":row="+rptidx[i][0]);
			        cell = row.getCell(rptidx[i][1]);
			        System.out.print(":cell="+rptidx[i][1]);
			        System.out.print(":celltype="+cell.getCellType());
			        if(cell != null){
	                   if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){	
	                	  //System.out.println("CELL_TYPE_STRING");	        
	                   	  if(report_no.equals("6-C2") && i == 0){
	                   	     cellString = cell.getStringCellValue();
	                   	     cellValue = Double.parseDouble(getUnit(cellString.trim()));
	                   	  }else if(report_no.equals("6-E") && (i == 0 || i == 26)){ 
	                   	  	 cellString = cell.getStringCellValue();
	                   	  	 cellValue = Double.parseDouble(getUnit(cellString.trim()));
	                   	  }else if(cell.getStringCellValue().indexOf("$") == -1 && cell.getStringCellValue().indexOf("-") == -1){
	                 	     cellValue = Double.parseDouble(cell.getStringCellValue());	                	
	                 	  }
	                   }else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC){	                 		
	                 	  cellValue = cell.getNumericCellValue();				   
	                   }else if(cell.getCellType()== HSSFCell.CELL_TYPE_FORMULA){	                 		
	                	  errMsg += "無法讀取公式內容";
	                	  return errMsg;
	                  }
			        }//end of cell != null
	                if(report_no.equals("1-A1")){
	                	if(i==23) cellValue = cellValue * 10000;//024000合格自有資本與風險性資產比率	               
	                	if(i== 3) cellValue_4 = cellValue * 1000;//99.12.28 fix 金庫BIS報表皆為仟元
	                	if(i==17) cellValue_18 = cellValue * 1000;//99.12.28 fix 金庫BIS報表皆為仟元
	                }
	                if(report_no.equals("1-B")){//儲存(1-B)acc_code=001000普通股
	                	if(i==0) cellValue_tmp1 = cellValue * 1000;//99.12.28 fix 金庫BIS報表皆為仟元
	                }
	                System.out.println(":data="+cellValue+":acc_code="+(String)bean.getValue("acc_code"));
	                if(report_no.equals("1-A1") && i==23){		                  
		               updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));//024000合格自有資本與風險性資產比率
		            }else if(report_no.equals("6-C2") && i==0){
		               updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));//幣別
		            }else if(report_no.equals("6-E") && (i == 0 || i == 26)){ 
		               updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));//幣別
		            }else{	
		               updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue * 1000,bean));//99.12.28 fix 金庫BIS報表皆為仟元
		            }
	                	                
	                printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue * 1000));	                
			     }//end of for   
        	}//end of 1-A1/1-B/1-B1/1-C/2-A/2-F/4-H
        	//printLog(logps,(String)bean.getValue("acc_name")+"="+String.valueOf(cellValue));
			if(report_no.equals("1-A1")){
				//計算第一類資本占風險性資產比率=(18)/(4)
				cellValue = (cellValue_18/cellValue_4)*10000;//025000第一類資本占風險性資產比率
				bean = (DataObject)dbData.get(24);//acc_code='025000'
				updateDBDataList.add(newDataList(m_year,m_month,report_no,data_yy,cellValue,bean));
				
				printLog(logps,"第一類資本占風險性資產比率=("+cellValue_18+"/"+cellValue_4+")*10000="+String.valueOf(cellValue));
			}
			
			if(report_no.equals("1-B")){
				//計算普通股占風險性資產比率=(1-B)001000/(1-A1)004000
				sql.delete(0,sql.length());
				paramList = new ArrayList();
				sql.append(" select amt ");
				sql.append(" from agribank_rpt");
				sql.append(" where m_year=? and m_month=? ");
				sql.append(" and acc_code=?");
				sql.append(" and acc_tr_type=?");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add("004000");
				paramList.add("1-A1");	     		
	     		dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"amt");
	     		if(dbData != null && dbData.size() > 0){
	     			cellValue_tmp2 = Double.parseDouble((((DataObject)dbData.get(0)).getValue("amt")).toString());
	     		}     		
	     		
				cellValue = (cellValue_tmp1/cellValue_tmp2)*10000;//026000普通股占風險性資產比率				
				dataList = new LinkedList();
				dataList.add(String.valueOf(cellValue));
				dataList.add(m_year);
				dataList.add(m_month);
				dataList.add("026000");
				dataList.add("1-A1");
				updateDBDataList.add(dataList);
				printLog(logps,"1-A1普通股占風險性資產比率=("+cellValue_tmp1+"/"+cellValue_tmp2+")*10000="+String.valueOf(cellValue));
			}			
            
	        updateDBSqlList.add(updateDBDataList);
	        updateDBList.add(updateDBSqlList);
	        if(DBManager.updateDB_ps(updateDBList)){
	           System.out.println(" UPDATE ok");	
	           printLog(logps,"更新資料庫成功");
	        }
	        
		 } catch (Exception e) {
            System.out.println("parseRpt.doParserRpt Error:" + e + e.getMessage());
            printLog(logps,"parseRpt.doParserRpt Error:" + e + e.getMessage());
        }
		return errMsg;		
	}    
      
   public String getErrMsg(){
	      return errMsg;  
   }
   
   private short[][] getRptIdx(String report_no){
   		short[][] rptidx = null;
   		/*cellNum.rowNum各別設定*/
		if(report_no.equals("1-A1")){
			rptidx = rptidx_1_A1;
		}else if(report_no.equals("1-B")){			
			rptidx = rptidx_1_B;			
		}else if(report_no.equals("1-B1")){			
			rptidx = rptidx_1_B1;				         
		}else if(report_no.equals("1-C")){			
			rptidx = rptidx_1_C;	
		}else if(report_no.equals("2-A")){			
			rptidx = rptidx_2_A;	
		}else if(report_no.equals("2-F")){			
			rptidx = rptidx_2_F;	
		}else if(report_no.equals("4-H")){			
			rptidx = rptidx_4_H;	
		}else if(report_no.equals("6-A1(NTD)")){/*新台幣*/			
			rptidx = rptidx_6_A1;	
		}else if(report_no.equals("6-A1(USD)")){/*美元*/			
			rptidx = rptidx_6_A1;	
		}else if(report_no.equals("6-B")){			
			rptidx = rptidx_6_B;		
		}else if(report_no.equals("6-B1")){			
			rptidx = rptidx_6_B1;	
		}else if(report_no.equals("6-B2")){			
			rptidx = rptidx_6_B2;		
		}else if(report_no.equals("6-G")){			
			rptidx = rptidx_6_G;
		//}else if(report_no.equals("6-A")){			
		//	rptidx = rptidx_6_A;	
		}else if(report_no.equals("6-C")){			
			rptidx = rptidx_6_C;	
		}else if(report_no.equals("6-C2")){			
			rptidx = rptidx_6_C2;	
		}else if(report_no.equals("6-E")){			
			rptidx = rptidx_6_E;		
		}	
		/*設定cellNum.rowNum.起始.結束*/
		if(report_no.equals("4-A-1")){			
			rptidx = rptidx_4_A_1;			
		}else if(report_no.equals("2-B")){			
			rptidx = rptidx_2_B;	
		}else if(report_no.equals("2-C")){			
			rptidx = rptidx_2_C;	
		}else if(report_no.equals("2-D")){			
			rptidx = rptidx_2_D;	
		}else if(report_no.equals("2-E")){			
			rptidx = rptidx_2_E;	
		}else if(report_no.equals("2-E1")){			
			rptidx = rptidx_2_E1;	
		}else if(report_no.equals("2-E2")){			
			rptidx = rptidx_2_E2;	
		}else if(report_no.equals("4-D-1")){			
			rptidx = rptidx_4_D_1;		
		}else if(report_no.equals("4-D-2")){			
			rptidx = rptidx_4_D_1;		
		}else if(report_no.equals("5-A")){			
			rptidx = rptidx_5_A;
		}else if(report_no.equals("6-A2-a(NTD)")){/*新台幣*/			
			rptidx = rptidx_6_A2_a;	
		}else if(report_no.equals("6-A2-a(USD)")){/*美元*/			
			rptidx = rptidx_6_A2_a;		
		}
		
   		return rptidx;
   }
   private List newDataList(String m_year,String m_month,String report_no,int data_yy,double cellValue,DataObject bean){
    	List dataList = new LinkedList();//參數detail
    	dataList.add(String.valueOf(cellValue));
		dataList.add(m_year);
		dataList.add(m_month);				
		dataList.add((String)bean.getValue("acc_code"));
		dataList.add(report_no);
		dataList.add(String.valueOf(data_yy));//100.06.14 add 資料年度
    	return dataList;
   }
   private List newDataList_acc_code(String m_year,String m_month,String report_no,int data_yy,double cellValue,String acc_code){
   		List dataList = new LinkedList();//參數detail
   		dataList.add(String.valueOf(cellValue));
   		dataList.add(m_year);
   		dataList.add(m_month);				
   		dataList.add(acc_code);
   		dataList.add(report_no);
   		dataList.add(String.valueOf(data_yy));//100.06.14 add 資料年度
   		return dataList;
   }
   public String getUnit(String unit){
   	   String unit_code = "0";
   	   String[][] unit_array = {{"0","新台幣"},{"1","美元"},{"2","英鎊"},{"3","日幣"},{"4","馬克"}};
   	   
   	   for(int i=0;i<unit_array.length;i++){
   	   	   if(unit.indexOf(unit_array[i][1]) != -1){
   	   	   	  unit_code = unit_array[i][0];
   	   	   }
   	   }
   	   return unit_code;
   }
   public static String getDollar(String unit){
	   String unit_code = "0";
	   String[][] unit_array = {{"新台幣","0"},{"美元","1"},{"英鎊","2"},{"日幣","3"},{"馬克","4"},{"合計","999"}};
	   
	   for(int i=0;i<unit_array.length;i++){
	   	   if(unit.indexOf(unit_array[i][1]) != -1){
	   	   	  unit_code = unit_array[i][0];
	   	   }
	   }
	   return unit_code;
   }
  
   public static String getCountry(String name,String code){
	   String unit_code = "0";
	   String[][] unit_array = {{"0","中華民國"},{"1","美國"},{"999","合計"}};
	   if("0".equals(code)){//回傳國別代碼
	      for(int i=0;i<unit_array.length;i++){
	   	    if(name.indexOf(unit_array[i][1]) != -1){
	   	   	  unit_code = unit_array[i][0];
	   	    }
	      }
	   }else{//回傳國別名稱
	   	  for(int i=0;i<unit_array.length;i++){
	   	    if(name.indexOf(unit_array[i][0]) != -1){
	   	   	   unit_code = unit_array[i][1];
	   	    }
	      }
	   }
	   
	   return unit_code;
   }
   
}

