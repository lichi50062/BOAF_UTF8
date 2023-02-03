/*
   102.07.03 created by2968     
   106.10.17 add 不為已裁撤,才計算家數 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR065W {
    public static String createRpt(String s_year,String s_month,String s_year1,String s_year4,String s_year3,String s_year2,String s_year0,String bank_type){
    	System.out.println("inpute s_month = "+s_month);
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List dbData = null;
		DataObject bean = null;
		int rowNum=0;
		String s_year_last="";
		String s_month_last="";
		String filename="N年內及目前月份逾放比家數統計.xls";
		String bank_type_name = "";
		String m_year = "";
		String m_month = "";
		String range_type = "";
		String cnt = "";
		String tmp = "";
		String tmpYear = "";
		int tmpLength= Integer.parseInt(s_year1)+1;
		reportUtil reportUtil = new reportUtil();
		if("6".equals(bank_type)){
		    bank_type_name = "農會";
        }else if("7".equals(bank_type)){
            bank_type_name = "漁會";
        }else{
            bank_type_name = "農漁會";
        }
		
		System.out.println("s_year1="+s_year1);
		System.out.println("s_year4="+s_year4);
		  System.out.println("s_year3="+s_year3);
		  System.out.println("s_year2="+s_year2);
		  System.out.println("s_year0="+s_year0);
		  
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
    		
    		String openfile="N年內及目前月份逾放比家數統計.xls";
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
	  		
	  		sqlCmd.append("select to_char(m_year) m_year, to_char(m_month) m_month, range_type, to_char(count(*)) cnt ");
	  		sqlCmd.append("from (select (CASE WHEN (a01.amt <= ?) THEN 'L5' ");//--第五個range家數
	  		sqlCmd.append("            WHEN (a01.amt > ? and a01.amt <= ?) THEN 'L4' ");//--第四個range家數
	  		sqlCmd.append("            WHEN (a01.amt > ? and a01.amt <= ?) THEN 'L3' ");//--第三個range家數
	  		sqlCmd.append("            WHEN (a01.amt > ? and a01.amt <= ?) THEN 'L2' ");//--第二range家數
	  		sqlCmd.append("            WHEN (a01.amt > ? ) THEN 'L1' ");//--第一個range家數
	  		sqlCmd.append("            ELSE '00' END) as RANGE_TYPE,m_year,m_month,bank_code,acc_code,amt "); 
	  		paramList.add(s_year4);
	  		paramList.add(s_year4);
	  		paramList.add(s_year3);
	  		paramList.add(s_year3);
	  		paramList.add(s_year2);
	  		paramList.add(s_year2);
	  		paramList.add(s_year0);
	  		paramList.add(s_year0);
	  		sqlCmd.append("      from a01_operation a01 ");
	  		sqlCmd.append("      left join (select bank_no,bn_type from bn01 where m_year=?)bn01 ");
	  		sqlCmd.append("      on a01.bank_code=bn01.bank_no "); 
	  		paramList.add("100");
	  		sqlCmd.append("      where acc_code = ? ");
	  		paramList.add("field_over_rate");
	  		sqlCmd.append("      and to_char(m_year * 100 + m_month) in (");
	  		if(!"".equals(s_year1)){
    	  		for(int k=Integer.parseInt(s_year)-Integer.parseInt(s_year1);k<Integer.parseInt(s_year);k++){
    	  		    sqlCmd.append("?,");
    	  		    paramList.add(k+"12");
    	  		    if(k==Integer.parseInt(s_year)-Integer.parseInt(s_year1)){
    	  		        tmp= String.valueOf(k) ;
    	  		    }else{
    	  		        tmp= tmp+","+String.valueOf(k);
    	  		    }
    	  		}
	  		}
	  		sqlCmd.append("      ?) ");
	  		paramList.add(s_year+addZeroForNum(s_month,2));
	  		if(!"".equals(tmp)){
	  		    tmp= tmp+","+s_year;
	  		}else{
	  		    tmp = s_year;
	  		}
	  		if("ALL".equals(bank_type)){
	  		    sqlCmd.append("      and bank_type  in  ('6','7') ");
	  		}else{
	  		    sqlCmd.append("      and bank_type  in  (?) ");
	  		    paramList.add(bank_type);
	  		}
	  		sqlCmd.append("      and bank_code != 'ALL' ");
	  		sqlCmd.append("      and bn_type !='2' ");//106.10.17 add 不為已裁撤,才計算家數
	  		sqlCmd.append("     )a01_operation ");          
	  		sqlCmd.append("group by m_year,m_month,range_type ");
            dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
            System.out.println("dbData.size()="+dbData.size());
			//列印表頭
            row = sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(bank_type_name+s_year1+"年內及目前月份逾比家數統計");
			//列印欄位
			row = sheet.getRow(3);
			cell=row.getCell((short)1);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("逾放比＞"+s_year0+"%");
            cell=row.getCell((short)2);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year2+"%<逾放比≦"+s_year0+"%");
            cell=row.getCell((short)3);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year3+"%＜逾放比≦"+s_year2+"%");
            cell=row.getCell((short)4);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year4+"%＜逾放比≦"+s_year3+"%");
            cell=row.getCell((short)5);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("逾放比≦"+s_year4+"%");
            if(dbData != null && dbData.size() != 0){
                rowNum = 3;
                for(int j=0;j<tmpLength;j++){
                    rowNum++;
                    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
                    for(short k=0;k<=5;k++){
                        cell=(row.getCell(k)==null)? row.createCell(k) : row.getCell(k);
                        insertCell_tail("0",wb,row,k);
                    }
                    if(j==tmpLength-1){
                        insertCell_tail(s_year+"年"+s_month+"月",wb,row,(short)0);
                    }else{
                        insertCell_tail(splitStr(tmp,j)+"年底",wb,row,(short)0);
                    }
                    for(int i=0;i<dbData.size();i++){
                        bean = (DataObject)dbData.get(i);
                        m_year = (bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();  
                        m_month = (bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();  
                        range_type = (bean.getValue("range_type") == null)?"":(bean.getValue("range_type")).toString(); 
                        cnt = (bean.getValue("cnt") == null)?"0":(bean.getValue("cnt")).toString();
                        if((splitStr(tmp,j)).equals(m_year)){
                            if("L1".equals(range_type)){
                                insertCell_tail(cnt,wb,row,(short)1);
                            }else if("L2".equals(range_type)){
                                insertCell_tail(cnt,wb,row,(short)2);
                            }else if("L3".equals(range_type)){
                                insertCell_tail(cnt,wb,row,(short)3);
                            }else if("L4".equals(range_type)){
                                insertCell_tail(cnt,wb,row,(short)4);
                            }else if("L5".equals(range_type)){
                                insertCell_tail(cnt,wb,row,(short)5);
                            }
                        }else{
                            if("".equals(tmpYear)){
                                insertCell_tail("0",wb,row,(short)1);
                                insertCell_tail("0",wb,row,(short)2);
                                insertCell_tail("0",wb,row,(short)3);
                                insertCell_tail("0",wb,row,(short)4);
                                insertCell_tail("0",wb,row,(short)5);
                            }else if(tmpYear.equals(String.valueOf(Integer.parseInt(splitStr(tmp,j))-1)) && !tmpYear.equals(m_year)){
                                row=sheet.getRow(rowNum);
                                insertCell_tail("0",wb,row,(short)1);
                                insertCell_tail("0",wb,row,(short)2);
                                insertCell_tail("0",wb,row,(short)3);
                                insertCell_tail("0",wb,row,(short)4);
                                insertCell_tail("0",wb,row,(short)5);
                            }
                        }
                        tmpYear=splitStr(tmp,j);
                    }
                }
            }

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

	public static void insertCell_tail(String Item,HSSFWorkbook wb,HSSFRow row,short c){
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
		}catch(Exception e){
			System.out.println("insertCell_tail Error:"+e+e.getMessage());
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
