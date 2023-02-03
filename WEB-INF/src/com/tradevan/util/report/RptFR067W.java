/*
   102.07.03 created by2968	
   104.08.13 fix 因在農金局無法下載報表但在公司內部環境下載OK,調整SQL及寫法 by 2295		      
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR067W {
    public static String createRpt(String s_year,String s_month,String s_year1,String s_year4,String s_year3,String s_year2){
    	System.out.println("inpute s_month = "+s_month);
		String errMsg = "";
		StringBuffer sqlCmd1 = new StringBuffer();//資本適足率.range家數統計
		List paramList1 = new ArrayList();
		List dbData1 = null;
		StringBuffer sqlCmd2 = new StringBuffer();//農.漁會家數及平均值
        List paramList2 = new ArrayList();
        List dbData2 = null;
		DataObject bean1 = null;
		DataObject bean2 = null;
		int rowNum=0;
		int rowTmp1=0;
		int rowTmp2=0;
		String s_year_last="";
		String s_month_last="";
		String filename="N年內及目前月份資本適足率家數統計.xls";
		String m_year = "";
		String m_month = "";
		String bank_type = "";
		String last_bank_type = "";
		String last_m_year = "";
		String range_type = "";
		String cnt = "";
		String L1cnt = "";
		String L2cnt = "";
		String L3cnt = ""; 
		String L4cnt = "";        
		String bank_sum_6 = "";
		String bank_sum_7 = "";
		String avg_percent = "";
		String tmp = "";
		String tmpYear1 = "0";
		String tmpYear2 = "0";
		int tmpLength= Integer.parseInt(s_year1)+1;
		reportUtil reportUtil = new reportUtil();
		//99.09.16 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"cd01"; 
	    String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    
		try{
			File xlsDir = new File(Utility.getProperties("xlsDir"));
        	File reportDir = new File(Utility.getProperties("reportDir"));

    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}
    		}
    		
    		String openfile="N年內及目前月份資本適足率家數統計.xls";
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );

	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )90 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();

	  		HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格
	  		HSSFFont f = wb.createFont();
	  		String div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
	  		if(Integer.parseInt(s_month) == 1) {
			    s_year_last   =  String.valueOf(Integer.parseInt(s_year) - 1);
			    s_month_last  =  "12";
			}else{
			    s_year_last   =  s_year;
			    s_month_last = String.valueOf(Integer.parseInt(s_month) - 1);
			}
	  		//資本適足率.range家數統計
	  		sqlCmd1.append(" select m_year,m_month,bank_type,sum(decode(range_type,'L1',cnt,0)) as l1cnt,");
	  		sqlCmd1.append(" sum(decode(range_type,'L2',cnt,0)) as l2cnt,");
	  		sqlCmd1.append(" sum(decode(range_type,'L3',cnt,0)) as l3cnt,");
	  		sqlCmd1.append(" sum(decode(range_type,'L4',cnt,0)) as l4cnt ");
	  		sqlCmd1.append(" from ");
	  		sqlCmd1.append("( ");
	  		sqlCmd1.append("select to_char(m_year) m_year, to_char(m_month) m_month, to_char(bank_type) bank_type, range_type, to_char(count(*)) cnt ");
	  		sqlCmd1.append("from (select (CASE WHEN (a01.amt < ?) THEN 'L4' "); //--第四個range家數
	  		sqlCmd1.append("                   WHEN (a01.amt >= ? and a01.amt < ?) THEN 'L3' "); //--第三個range家數
	  		sqlCmd1.append("                   WHEN (a01.amt >= ? and a01.amt < ?) THEN 'L2' "); //--第二range家數
	  		sqlCmd1.append("                   WHEN (a01.amt >= ? ) THEN 'L1' "); //--第一個range家數
	  		sqlCmd1.append("              ELSE '00' END) as RANGE_TYPE,m_year,m_month,bank_type,bank_code,acc_code,amt "); 
	  		paramList1.add(s_year4);
	  		paramList1.add(s_year4);
	  		paramList1.add(s_year3);
	  		paramList1.add(s_year3);
	  		paramList1.add(s_year2);
	  		paramList1.add(s_year2);
	  		sqlCmd1.append("      from a01_operation a01 ");
	  		sqlCmd1.append("      where acc_code = ? ");
	  		paramList1.add("field_captial_rate");
	  		sqlCmd1.append("      and to_char(m_year * 100 + m_month) in (");
	  		if(!"".equals(s_year1)){
                for(int k=Integer.parseInt(s_year)-Integer.parseInt(s_year1);k<Integer.parseInt(s_year);k++){
                    sqlCmd1.append("?,");
                    paramList1.add(k+"12");
                    if(k==Integer.parseInt(s_year)-Integer.parseInt(s_year1)){
                        tmp= String.valueOf(k) ;
                    }else{
                        tmp= tmp+","+String.valueOf(k);
                    }
                }
            }
	  		sqlCmd1.append("      ?) ");
	  		paramList1.add(s_year+addZeroForNum(s_month,2));
	  		if(!"".equals(tmp)){
                tmp= tmp+","+s_year;
            }else{
                tmp = s_year;
            }
	  		sqlCmd1.append("      and bank_type  in  ('6','7') ");
	  		sqlCmd1.append("      and bank_code != 'ALL' ");
	  		sqlCmd1.append("     )a01_operation ");          
	  		sqlCmd1.append("group by m_year,m_month,bank_type,range_type ");	
	  		sqlCmd1.append("order by m_year,m_month,bank_type,range_type ");
	  		sqlCmd1.append(" )");  
	  		sqlCmd1.append(" group by m_year,m_month,bank_type ");
	  		sqlCmd1.append(" order by m_year,m_month,bank_type ");
            dbData1 = DBManager.QueryDB_SQLParam(sqlCmd1.toString(),paramList1,"m_year,m_month,l1cnt,l2cnt,l3cnt,l4cnt");
            System.out.println("dbData1.size()="+dbData1.size());
            //農.漁會家數及平均值
            sqlCmd2.append("select to_char(a01_operation.m_year) as m_year, ");
            sqlCmd2.append("to_char(a01_operation.m_month) as m_month, ");
            sqlCmd2.append("bank_type, ");//--6農.7漁會
            sqlCmd2.append("to_char(bn01_month.bank_sum_6) as bank_sum_6, ");//--農會家數
            sqlCmd2.append("to_char(bn01_month.bank_sum_7) as bank_sum_7, ");//--漁會家數
            sqlCmd2.append("to_char(decode(bank_type,'6',round(a01_operation.amt /  bn01_month.bank_sum_6,2),'7',round(a01_operation.amt /  bn01_month.bank_sum_7,2),0)) as avg_percent ");//--資本適足率平均值
            sqlCmd2.append("from ( ");
            sqlCmd2.append("select m_year,m_month,bank_type,sum(amt) as amt ");
            sqlCmd2.append("from (select * ");
            sqlCmd2.append("     from a01_operation a01 ");
            sqlCmd2.append("     where acc_code = 'field_captial_rate' ");
            sqlCmd2.append("     and to_char(m_year * 100 + m_month) in (");
            if(!"".equals(s_year1)){
                for(int k=Integer.parseInt(s_year)-Integer.parseInt(s_year1);k<Integer.parseInt(s_year);k++){
                    sqlCmd2.append("?,");
                    paramList2.add(k+"12");
                }
            }
            sqlCmd2.append("     ?) ");
            paramList2.add(s_year+addZeroForNum(s_month,2));
            sqlCmd2.append("     and bank_type  in  ('6','7') ");
            sqlCmd2.append("     and bank_code != 'ALL' ");
            sqlCmd2.append("     )a01_operation ");          
            sqlCmd2.append("group by m_year,m_month,bank_type ");       
            sqlCmd2.append(")a01_operation "); 
            sqlCmd2.append("left join bn01_month on a01_operation.m_year=bn01_month.m_year and a01_operation.m_month = bn01_month.m_month ");
            sqlCmd2.append(" order by m_year,m_month,bank_type");
            dbData2 = DBManager.QueryDB_SQLParam(sqlCmd2.toString(),paramList2,"");
            System.out.println("dbData2.size()="+dbData2.size());
            //列印表頭
            row = sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year1+"年內及目前月份資本適足率家數統計");
			//列印欄位
			row = sheet.getRow(4);
			cell=row.getCell((short)3);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("BIS≧"+s_year2+"%");
            cell=row.getCell((short)4);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year3+"%≦BIS<"+s_year2+"%");
            cell=row.getCell((short)5);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year4+"%≦BIS<"+s_year3+"%");
            cell=row.getCell((short)6);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("BIS<"+s_year4+"%");
           
            rowNum = 5;           
            if(dbData1 != null && dbData1.size() != 0){
                for(int j=0;j<tmpLength;j++){
                    for(int r=1;r<=2;r++){
                        row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);                        
                        //System.out.println("row="+row.getRowNum());
                        for(short k=1;k<=6;k++){
                            cell=(row.getCell(k)==null)? row.createCell(k) : row.getCell(k);                           
                            if(k==1){
                                insertCell_tail("",wb,row,k);
                            }else{
                                insertCell_tail("0",wb,row,k);
                            }
                        }
                        if(j==tmpLength-1){
                            insertCell_tail(s_year+"年"+s_month+"月",wb,row,(short)0);
                        }else{
                            insertCell_tail(splitStr(tmp,j)+"年底",wb,row,(short)0);
                        }
                        //合併儲存格，參數分別為起始行、起始列、結束行、結束列 
                        if(r==2){
                          //System.out.println("rowNum="+rowNum);  
                          sheet.addMergedRegion(new Region(rowNum-1, (short)0, rowNum, (short)0));
                        }  
                        rowNum++;
                       
                    }
                }
                                       
                    rowNum=5;//104.08.12                    
                    for(int i=0;i<dbData1.size();i++){
                        bean1 = (DataObject)dbData1.get(i);
                        m_year = (bean1.getValue("m_year") == null)?"":(bean1.getValue("m_year")).toString();
                        m_month = (bean1.getValue("m_month") == null)?"":(bean1.getValue("m_month")).toString();  
                        bank_type = (bean1.getValue("bank_type") == null)?"":(String)bean1.getValue("bank_type");
                        L1cnt = (bean1.getValue("l1cnt") == null)?"0":(bean1.getValue("l1cnt")).toString();
                        L2cnt = (bean1.getValue("l2cnt") == null)?"0":(bean1.getValue("l2cnt")).toString();
                        L3cnt = (bean1.getValue("l3cnt") == null)?"0":(bean1.getValue("l3cnt")).toString();
                        L4cnt = (bean1.getValue("l4cnt") == null)?"0":(bean1.getValue("l4cnt")).toString();
                        System.out.println("m_year="+m_year+":m_month="+m_month+":bank_type="+bank_type+":L1cnt="+L1cnt+":L2cnt="+L2cnt+":L3cnt="+L3cnt+":L4cnt="+L4cnt);
                        
                        yearLoop:
                        for(int j=0;j<tmpLength;j++){
                            if((splitStr(tmp,j)).equals(m_year)){
                                row = sheet.getRow(rowNum);
                                System.out.println("insert.rowNum="+rowNum);
                                insertCell_tail(L1cnt,wb,row,(short)3);                              
                                insertCell_tail(L2cnt,wb,row,(short)4);                               
                                insertCell_tail(L3cnt,wb,row,(short)5);
                                insertCell_tail(L4cnt,wb,row,(short)6);
                                ++rowNum;
                                break yearLoop;
                            }
                        }
                    }
                    rowNum=5;
                    for(int k=0;k<dbData2.size();k++){
                        bean2 = (DataObject)dbData2.get(k);
                        m_year = (bean2.getValue("m_year") == null)?"":(bean2.getValue("m_year")).toString();  
                        m_month = (bean2.getValue("m_month") == null)?"":(bean2.getValue("m_month")).toString();
                        bank_sum_6 = (bean2.getValue("bank_sum_6") == null)?"":(bean2.getValue("bank_sum_6")).toString();  
                        bank_sum_7 = (bean2.getValue("bank_sum_7") == null)?"":(bean2.getValue("bank_sum_7")).toString();  
                        avg_percent = (bean2.getValue("avg_percent") == null)?"":(bean2.getValue("avg_percent")).toString();  
                        bank_type = (bean2.getValue("bank_type") == null)?"":(bean2.getValue("bank_type")).toString();
                        System.out.println("m_year="+m_year+":m_month="+m_month+":bank_type"+bank_type+":bank_sum_6="+bank_sum_6+":bank_sum_7"+bank_sum_7+":avg_percent="+avg_percent);
                        yearLoop1:
                        for(int j=0;j<tmpLength;j++){ 
                            if((splitStr(tmp,j)).equals(m_year)){
                                //row = ("6".equals(bank_type)? sheet.getRow(rowNum):sheet.getRow(++rowNum));
                                row = sheet.getRow(rowNum);
                                System.out.println("insert.rowNum="+rowNum);
                                insertCell_tail(("6".equals(bank_type)?"(農)"+bank_sum_6:"(漁)"+bank_sum_7),wb,row,(short)1);                           
                                insertCell_tail(avg_percent,wb,row,(short)2);
                                ++rowNum;
                                break yearLoop1;
                            }
                        }
                    }
                
            }//end if(dbData1 != null && dbData1.size() != 0){
            
	        HSSFFooter footer = sheet.getFooter();
	        footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}

	public static void insertCell_tail(String Item,HSSFWorkbook wb,HSSFRow row,short c)
	{
		try{
			HSSFCell cell=(row.getCell(c)==null)? row.createCell(c) : row.getCell(c);
			HSSFCellStyle cs1 = wb.createCellStyle();
			//HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
			//設置邊框
			cs1.setBorderTop(HSSFCellStyle.BORDER_THIN); //上邊框
		    cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下邊框
		    cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN); //左邊框
		    cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); //右邊框
			HSSFFont f = wb.createFont();
		    f.setFontHeightInPoints((short) 12);//字體大小
		    f.setFontName("標楷體");//字體格式
		    cs1.setFont(f);
		    cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平置中
		    cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
		    cell.setCellStyle(cs1);
			cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
			cell.setCellValue(Item);
			//System.out.println("insertCell_tail ok:cell="+c+":row="+row.getRowNum()+":value="+Item);
		}catch(Exception e){
			System.out.println("insertCell_tail Error:cell="+c+":row="+row.getRowNum()+":value="+Item+":"+e+e.getMessage());
		}
	}
	public static String addZeroForNum(String str, int strLength) {
        int strLen = str.length();
        if (strLen < strLength) {
            while (strLen < strLength) {
                StringBuffer sb = new StringBuffer();
                sb.append("0").append(str);//左補零
                //sb.append(str).append("0");//右補零
                str = sb.toString();
                strLen = str.length();
            }
        }
        return str;
    }
	public static String splitStr(String str, int i) {
        String[] strarray=str.split(",");
        return strarray[i];
    }
}
