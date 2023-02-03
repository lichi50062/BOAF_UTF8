/*
 * create 農漁會信用部財營運狀況警訊報表-各別指標明細 by 2968
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.awt.Font;
import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.sun.corba.se.impl.javax.rmi.CORBA.Util;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.*; 

public class RptWR100W {
    
	static DecimalFormat df_md = new DecimalFormat("############0.00");//顯示小數點至第2位,不足者補0
	static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    public static String createRpt(String s_year,String s_month,String unit,String rptType,String wr_rpt){
          String errMsg = "";
          String last_year = "";
          String last_month = "";
          String lastSeason_year = "";
          String lastSeason_month = "";
          FileOutputStream fileOut = null;          
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
          //99.09.16 add 查詢年度100年以前.縣市別不同===============================
          String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
          //===================================================================== 
          s_month = String.valueOf(Integer.parseInt(s_month));
          last_year  = (Integer.parseInt(s_month) == 1) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          last_month = (Integer.parseInt(s_month) == 1) ? "12" : String.valueOf(Integer.parseInt(s_month) - 1);
          lastSeason_year  = (Integer.parseInt(s_month) == 3) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          if(Integer.parseInt(s_month) <= 3 && Integer.parseInt(s_month)>=1){
              lastSeason_month = "12";
          }else if(Integer.parseInt(s_month) <= 6 && Integer.parseInt(s_month)>=4){
              lastSeason_month = "3";
          }else if(Integer.parseInt(s_month) <= 9 && Integer.parseInt(s_month)>=7){
              lastSeason_month = "6";
          }else if(Integer.parseInt(s_month) <= 12 && Integer.parseInt(s_month)>=10){
              lastSeason_month = "9";
          }
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
            int table = 0;
            String openfile="";//要去開啟的範本檔
            if("1".equals(rptType)){
                openfile = "農漁會信用部營運狀況警訊報表-各項指標明細.xls";
                table = 8;
            }else if("2".equals(rptType)){
                openfile="農漁會信用部營運狀況警訊報表-專案農貸-各項指標明細.xls";
                table = 3;
            }
            System.out.println("開啟檔:" + openfile);
            FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );         
            //設定FileINputStream讀取Excel檔
            POIFSFileSystem fs = new POIFSFileSystem( finput );//新增一個xls unit
            if(fs==null){
                System.out.println("open 範本檔失敗");
            } else{ 
                System.out.println("open 範本檔成功");
            }
            HSSFWorkbook wb = new HSSFWorkbook(fs);//新增一個sheet
            if(wb==null){
                System.out.println("open工作表失敗");
            } else {
                System.out.println("open 工作表成功");
            }
            
            
            HSSFCellStyle style = wb.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setFontHeightInPoints((short) 14);
            style.setFont(font);
            HSSFCellStyle style1 = wb.createCellStyle();
            HSSFFont font1 = wb.createFont();
            font1.setFontHeightInPoints((short) 12);
            font1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//加粗  
            style1.setFont(font1);
            
            for(int t=1;t<=table;t++){
                String strT=String.valueOf(t);
                HSSFSheet sheet = wb.getSheetAt(t-1);//讀取第t-1個工作表，宣告其為sheet 
                if(sheet==null){
                    System.out.println("open sheet 失敗");
                }else {
                    System.out.println("open sheet 成功");
                }
                //做屬性設定
                HSSFPrintSetup ps = sheet.getPrintSetup();  //取得設定
                //sheet.setZoom(80, 100);                   //螢幕上看到的縮放大小
                //sheet.setAutobreaks(true);                //自動分頁
                
                //設定頁面符合列印大小
                sheet.setAutobreaks( false );
                /*
                if("1".equals(rptType)){
                    ps.setScale( ( short )75 );                 //列印縮放百分比
                }else if("2".equals(rptType)){
                    ps.setScale( ( short )100 );
                }
                */
                ps.setPaperSize( ( short )9 );              //設定紙張大小 A4
                
                //設定表頭 為固定 先設欄的起始再設列的起始
                //wb.setRepeatingRowsAndColumns(0, 1, 17, 2, 3);
                finput.close();
                HSSFFooter footer = sheet.getFooter();            
                if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
                footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                        HSSFFooter.numPages() );                                        
                footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
                
                //設定表頭===============================================================================
                String title0 = s_year+"年"+s_month+"月份農漁會信用部營運狀況警訊報表";
                String title1 = "";
                if("0".equals(wr_rpt)){
                    title1 = "與上月份("+last_year+"年"+last_month+"月)比較";
                }else if("1".equals(wr_rpt)){
                    title1 = "與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較";
                }else if("2".equals(wr_rpt)){
                    title1 = "與上一年度同期("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較";
                }else if("3".equals(wr_rpt)){
                    title0 += "（專案農貸部分）";
                    title1 = "與上月份("+last_year+"年"+last_month+"月)比較";
                }else if("4".equals(wr_rpt)){
                    title0 += "（專案農貸部分）";
                    title1 = "與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較";
                }else if("5".equals(wr_rpt)){
                    title0 += "（專案農貸部分）";
                    title1 = "與上一年度同期("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較";
                }
                row=sheet.getRow((short)0);
                cell = row.getCell((short)0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(title0);
                row=sheet.getRow((short)1);
                cell = row.getCell((short)0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(title1);
                
                row=sheet.getRow((short)2);
                cell = row.getCell((short)0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue("單位：新台幣"+Utility.getUnitName(unit)+",％,百分點");
                String subT = "";
                if("1".equals(rptType)){
                    if("4".equals(strT)){
                        row=sheet.getRow((short)7);
                        if("0".equals(wr_rpt)){
                            subT = "本("+s_month+")月份之(建築放款/信用部上年度決算淨值)-上("+last_month+")月份之(建築放款/信用部上年度決算淨值)";
                        }else if("1".equals(wr_rpt)){
                            subT = "本季("+s_month+"月)之(建築放款/信用部上年度決算淨值)-上季("+lastSeason_month+"月)之(建築放款/信用部上年度決算淨值)";
                        }else if("2".equals(wr_rpt)){
                            subT = "本年("+s_year+"年"+s_month+"月)之(建築放款/信用部上年度決算淨值)-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)之(建築放款/信用部上年度決算淨值)";
                        }
                        cell = row.getCell((short)8);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                        row=sheet.getRow((short)9);
                        if("0".equals(wr_rpt)){
                            subT = "與上("+last_month+")月份超逾100%者比較，本("+s_month+")月份增加超逾100%者";
                        }else if("1".equals(wr_rpt)){
                            subT = "與上季("+lastSeason_month+"月)超逾100%者比較，本季("+s_month+"月)增加超逾100%者";
                        }else if("2".equals(wr_rpt)){
                            subT = "與上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)超逾100%者比較，本年("+s_year+"年"+s_month+"月)增加超逾100%者";
                        }
                        cell = row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                    }else if("5".equals(strT)){
                        row=sheet.getRow((short)7);
                        if("0".equals(wr_rpt)){
                            subT = "本("+s_month+")月份逾放比率-上("+last_month+")月份逾放比率";
                        }else if("1".equals(wr_rpt)){
                            subT = "本季("+s_month+"月)逾放比率-上季("+lastSeason_month+"月)逾放比率";
                        }else if("2".equals(wr_rpt)){
                            subT = "本年("+s_year+"年"+s_month+"月)逾放比率-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)逾放比率";
                        }
                        cell = row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                    }else if("8".equals(strT)){
                        row=sheet.getRow((short)7);
                        if("0".equals(wr_rpt)){
                            subT = "本("+s_month+")月份放款覆蓋率-上("+last_month+")月份放款覆蓋率";
                        }else if("1".equals(wr_rpt)){
                            subT = "本季("+s_month+"月)放款覆蓋率-上季("+lastSeason_month+"月)放款覆蓋率";
                        }else if("2".equals(wr_rpt)){
                            subT = "本年("+s_year+"年"+s_month+"月)放款覆蓋率-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)放款覆蓋率";
                        }
                        cell = row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                    }
                }else{
                    String tmpT = "月";
                    if("4".equals(wr_rpt)){
                        tmpT = "季";
                    }else if("5".equals(wr_rpt)){
                        tmpT = "年度";
                    }
                    row=sheet.getRow((short)7);
                    if("1".equals(strT)){
                        subT = "本"+tmpT+"專案農貸放款餘額-上"+tmpT+"專案農貸放款餘額";
                        cell = row.getCell((short)4);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                    }else if("2".equals(strT)){
                        subT = "本"+tmpT+"專案農貸逾放比率-上"+tmpT+"專案農貸逾放比率";
                        cell = row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellValue(subT);
                    }
                }
                
                //取得資料===============================================================================
                List dbData = getData(wlx01_m_year,s_year,s_month,unit,wr_rpt,String.valueOf(t),rptType);
                //寫入資料=======================================================================================
                int rowNum=10;
                if(dbData != null && dbData.size()>0){
                    for(int i=0;i<dbData.size();i++){
                        DataObject bean = (DataObject)dbData.get(i);
                        String bank_code= (bean.getValue("bank_code")==null)?"":bean.getValue("bank_code").toString();
                        String bank_name= (bean.getValue("bank_name")==null)?"":bean.getValue("bank_name").toString();
                        String wr_count= (bean.getValue("wr_count")==null)?"":Utility.setCommaFormat(bean.getValue("wr_count").toString());
                        String sort1= (bean.getValue("sort1")==null)?"":Utility.setCommaFormat(bean.getValue("sort1").toString());
                        String field_wr1= (bean.getValue("field_wr1")==null)?"":Utility.setCommaFormat(bean.getValue("field_wr1").toString());
                        String sort2= (bean.getValue("sort2")==null)?"":Utility.setCommaFormat(bean.getValue("sort2").toString());
                        String field_wr2= (bean.getValue("field_wr2")==null)?"":Utility.setCommaFormat(bean.getValue("field_wr2").toString());
                        String sort3= (bean.getValue("sort3")==null)?"":Utility.setCommaFormat(bean.getValue("sort3").toString());
                        String field_wr3= (bean.getValue("field_wr3")==null)?"":Utility.setCommaFormat(bean.getValue("field_wr3").toString());
                        String field_over_rate= (bean.getValue("field_over_rate")==null)?"":Utility.setCommaFormat(bean.getValue("field_over_rate").toString());
                        String field_backup_over_rate= (bean.getValue("field_backup_over_rate")==null)?"":Utility.setCommaFormat(bean.getValue("field_backup_over_rate").toString());
                        String field_backup_credit_rate= (bean.getValue("field_backup_credit_rate")==null)?"":Utility.setCommaFormat(bean.getValue("field_backup_credit_rate").toString());
                        String field_992710_990230_rate= (bean.getValue("field_992710_990230_rate")==null)?"":Utility.setCommaFormat(bean.getValue("field_992710_990230_rate").toString());
                        
                        row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
                        insertCell(String.valueOf(i+1),wb,row,(short)0,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_CENTER,true);
                        insertCell(bank_code,wb,row,(short)1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,false);
                        insertCell(bank_name+"("+wr_count+")",wb,row,(short)2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
                        if("1".equals(rptType)){
                            insertCell(sort1,wb,row,(short)3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_CENTER,true);
                            insertCell(field_wr1,wb,row,(short)4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
                            insertCell(sort2,wb,row,(short)5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_CENTER,true);
                            insertCell(field_wr2,wb,row,(short)6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
                            if("5".equals(strT)||"6".equals(strT)){
                                insertCell(field_over_rate,wb,row,(short)8,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
                            }else if("8".equals(strT)){
                                insertCell(field_backup_over_rate,wb,row,(short)8,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
                                insertCell(field_backup_credit_rate,wb,row,(short)9,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
                            }else if("4".equals(strT)||"7".equals(strT)){
                                insertCell(sort3,wb,row,(short)7,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_CENTER,true);
                                insertCell(field_wr3,wb,row,(short)8,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
                                if("4".equals(strT)){
                                    insertCell(field_992710_990230_rate,wb,row,(short)11,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
                                }else if("7".equals(strT)){
                                    insertCell(field_over_rate,wb,row,(short)10,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
                                }
                            }
                            if("4".equals(strT)||"7".equals(strT)){
                                insertCell("",wb,row,(short)9,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
                            }else{
                                insertCell("",wb,row,(short)7,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
                            }
                        }else if("2".equals(rptType)){
                            insertCell(sort1,wb,row,(short)3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_CENTER,true);
                            insertCell(field_wr1,wb,row,(short)4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
                            if("2".equals(strT)){
                                insertCell(sort2,wb,row,(short)5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_CENTER,true);
                                insertCell(field_wr2,wb,row,(short)6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
                                insertCell("",wb,row,(short)7,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
                            }else{
                                insertCell("",wb,row,(short)5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
                            }
                        }
                        rowNum++;
                    }
                }
                rowNum = rowNum+3;
                short c = 7;
                if("1".equals(rptType) && ("4".equals(strT)||"7".equals(strT))){
                    c = 9;
                }else if("2".equals(rptType) && ("1".equals(strT)||"3".equals(strT))){
                    c = 5;
                }
                row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
                insertCell("附表"+t,wb,row,c,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT,true);
            }
           // Write the output to a file============================   
           fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+openfile );
           wb.write( fileOut );
           fileOut.close();            
           System.out.println("儲存完成");
        } catch ( Exception e ) {            
           e.printStackTrace();
           
        } finally {
           try {
               if ( fileOut != null ) {
                   fileOut.close();
               }
           } catch ( Exception e ) {
                 System.out.println(e.getMessage() );
           }
        }
      return errMsg;
    }
    public static List getData(String wlx01_m_year,String s_year,String s_month,String unit,String wr_rpt,String table,String rptType){
        List qList = new ArrayList();
        if("1".equals(rptType)){
            if("4".equals(table)||"7".equals(table)){
                qList = getRemarkList2(wlx01_m_year,s_year,s_month,unit,wr_rpt,table);
            }else{
                qList = getRemarkList1(wlx01_m_year,s_year,s_month,unit,wr_rpt,table);
            }
        }else if("2".equals(rptType)){
            if("2".equals(table)){
                qList = getRemarkList3(wlx01_m_year,s_year,s_month,unit,wr_rpt,table);
            }else{
                qList = getRemarkList4(wlx01_m_year,s_year,s_month,unit,wr_rpt,table);
            }
        }
        return qList;
    }
    //各項警示指標:附表1/附表2/附表3/附表5/附表6/附表8[2項警示指標]查詢SQL
    public static List getRemarkList1(String wlx01_m_year,String s_year,String s_month,String unit,String wr_rpt,String table){
        StringBuffer sqlCmd = new StringBuffer(); 
        List paramList = new ArrayList();
        String code1="";
        String code2="";
        int serial1 = 0;
        int serial2 = 0;
        sqlCmd.append("select a1.bank_code,a1.bank_name,"); 
        sqlCmd.append("       b.wr_count,");//警示項目加總
        sqlCmd.append("       sort1,");//第一項指標排名
        sqlCmd.append("       field_wr1,");//第一項指標
        sqlCmd.append("       sort2,");//第二項指標排名
        sqlCmd.append("       a1.field_wr2,");//第二項指標      
        sqlCmd.append("       c.field_over_rate,");//逾放比:附表5/附表6使用
        sqlCmd.append("       c.field_backup_over_rate,");//備呆覆蓋率 :附表8使用      
        sqlCmd.append("       c.field_backup_credit_rate ");//放款覆蓋率 :附表8使用       
        sqlCmd.append("from ");
        sqlCmd.append(" (select bank_code,bank_name,");
        sqlCmd.append("         decode(sum(sort1),'','',sum(sort1)) as sort1,");
        sqlCmd.append("         sum(decode(field_wr1,null,'',field_wr1)) as field_wr1,");
        sqlCmd.append("         decode(sum(sort2),'','',sum(sort2)) as sort2,");
        sqlCmd.append("         sum(decode(field_wr2,null,'',field_wr2)) as field_wr2 ");
        sqlCmd.append("    from ");
        sqlCmd.append("     (select a.bank_code,a.bank_name,");       
        sqlCmd.append("             sort1,");//警示項目1排名
        sqlCmd.append("             field_wr1,");//警示項目1
        sqlCmd.append("             sort2,");//警示項目2排名
        sqlCmd.append("             field_wr2 ");//警示項目1
        sqlCmd.append("        from ");
        sqlCmd.append("         (select bank_code,bank_name,sort1,field_wr1,sort2,field_wr2 ");
        sqlCmd.append("        from (select t.*,rownum as sort1,null as sort2 ");//該警示指標排名       
        sqlCmd.append("            from(select * ");                
        sqlCmd.append("                   from (select bank_code,");//機構代號           
        sqlCmd.append("                                bn01.bank_name,");//機構名稱
        if("1".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,1,amt,''),'')) as field_wr1,");//存款總額.增加比率前5名:附表1,才需加入
            code1 = "field_debit_rate";
            serial1 = 1;
        }else if("2".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,3,amt,''),'')) as field_wr1,");//放款總額.增加比率前5名:附表2,才需加入
            code1 = "field_credit_rate";
            serial1 = 3;
        }else if("3".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_990610_rate',decode(wr_range_serial,5,amt,''),'')) as field_wr1,");//非會員放款.增加比率前5名:附表3,才需加入
            code1 = "field_990610_rate";
            serial1 = 5;
        }else if("5".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_diff_over',decode(wr_range_serial,16,round(amt /?,0),''),'')) as field_wr1,");//逾期放款增加金額前5名:附表5,才需加入
            paramList.add(unit);
            code1 = "field_diff_over";
            serial1 = 16;
        }else if("6".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_diff_992530',decode(wr_range_serial,19,round(amt /?,0),''),'')) as field_wr1,");//逾放-非會員.增加金額前5名:附表6,才需加入
            paramList.add(unit);
            code1 = "field_diff_992530";
            serial1 = 19;
        }else if("8".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_diff_backup',decode(wr_range_serial,27,round(amt /?,0),''),'')) as field_wr1,");//備抵呆帳.減少金額前5名:附表8,才需加入
            paramList.add(unit);
            code1 = "field_diff_backup";
            serial1 = 27;
        }
        sqlCmd.append("                                null as field_wr2 ");
        sqlCmd.append("                           from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
        sqlCmd.append("                          where wr_operation.m_year=? ");
        sqlCmd.append("                            and m_month=? ");
        sqlCmd.append("                            and wr_rpt=? "); //0:與上月比較/1:與上季比較/2:與上年度同期比較
        sqlCmd.append("                            and acc_code=?");//依不同表格,變動變數設定
        sqlCmd.append("                            and wr_range_serial=").append(serial1);//依不同表格,變動變數設定
        sqlCmd.append("                            and warn_type=? ");
        paramList.add(wlx01_m_year);
        paramList.add(s_year);
        paramList.add(s_month);
        paramList.add(wr_rpt);
        paramList.add(code1);
        paramList.add("Y");
        sqlCmd.append("                          group by bank_code,bank_name ");
        sqlCmd.append("                        ) ");
        sqlCmd.append("                  order by field_wr1 ");
        if("8".equals(table)){//desc除了附表8時,用asc其餘皆用desc
            sqlCmd.append(" asc ");
        }else{
            sqlCmd.append(" desc ");
        }
        sqlCmd.append("                )t ");
        sqlCmd.append("           union all ");
        sqlCmd.append("          select t1.*,null,rownum as sort2 ");
        sqlCmd.append("            from(select * ");
        sqlCmd.append("                   from (select bank_code,");//機構代號
        sqlCmd.append("                                bn01.bank_name,");//機構名稱
        sqlCmd.append("                                null as field_wr1,");
        if("1".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,2,amt,''),'')) as field_wr2 ");//存款總額.減少比率前5名:附表1,才需加入
            code2 = "field_debit_rate";
            serial2 = 2;
        }else if("2".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,4,amt,''),'')) as field_wr2 ");//放款總額.減少比率前5名:附表2,才需加入
            code2 = "field_credit_rate";
            serial2 = 4;
        }else if("3".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_noassure_rate',decode(wr_range_serial,7,amt,''),'')) as field_wr2 ");//無擔保放款.增加比率前5名:附表3才需加入
            code2 = "field_noassure_rate";
            serial2 = 7;
        }else if("5".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_diff_over_rate',decode(wr_range_serial,18,amt,''),'')) as field_wr2 ");//本月份逾放比率-上月份逾放比率增加百分點前5名:附表5才需加入
            code2 = "field_diff_over_rate";
            serial2 = 18;
        }else if("6".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_992720_992710_rate',decode(wr_range_serial,21,amt,''),'')) as field_wr2 ");//建築放款中之逾放/建築放款.前5名:附表6才需加入
            code2 = "field_992720_992710_rate";
            serial2 = 21;
        }else if("8".equals(table)){
            sqlCmd.append("                            sum(decode(acc_code,'field_diff_backup_credit_rate',decode(wr_range_serial,28,amt,''),'')) as field_wr2 ");//本月份之放款覆蓋率-上月份放款覆蓋率.最低前5名:附表8才需加入
            code2 = "field_diff_backup_credit_rate";
            serial2 = 28;
        }
        sqlCmd.append("                           from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
        sqlCmd.append("                          where wr_operation.m_year=? ");
        sqlCmd.append("                            and m_month=? ");
        sqlCmd.append("                            and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較
        sqlCmd.append("                            and acc_code=? ");//依不同表格,變動變數設定
        sqlCmd.append("                            and wr_range_serial=").append(serial2);//依不同表格,變動變數設定
        sqlCmd.append("                            and warn_type=? ");
        paramList.add(wlx01_m_year);
        paramList.add(s_year);
        paramList.add(s_month);
        paramList.add(wr_rpt);
        paramList.add(code2);
        paramList.add("Y");
        sqlCmd.append("                          group by bank_code,bank_name ");
        sqlCmd.append("                         ) ");
        sqlCmd.append("                   order by field_wr2 ");
        if("3".equals(table)||"5".equals(table)||"6".equals(table)){//asc附表1/附表2/附表8-asc;附表3/附表5/附表6-desc
            sqlCmd.append(" desc ");
        }else{
            sqlCmd.append(" asc ");
        }
        sqlCmd.append("                 )t1 ");
        sqlCmd.append("          ) ");
        sqlCmd.append("          order by bank_code ");
        sqlCmd.append("        )a ");
        sqlCmd.append("     ) ");
        sqlCmd.append("   group by bank_code,bank_name ");
        sqlCmd.append("   )a1,");
        sqlCmd.append("   (select bank_code,");//機構代號
        sqlCmd.append("           bn01.bank_name,");//機構名稱
        sqlCmd.append("           count(*) as wr_count ");//
        sqlCmd.append("      from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
        sqlCmd.append("     where wr_operation.m_year=? ");
        sqlCmd.append("       and m_month=? ");
        sqlCmd.append("       and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較
        sqlCmd.append("       and warn_type=? ");
        paramList.add(wlx01_m_year);
        paramList.add(s_year);
        paramList.add(s_month);
        paramList.add(wr_rpt);
        paramList.add("Y");
        sqlCmd.append("     group by bank_code,bank_name ");
        sqlCmd.append("     order by bank_code ");
        sqlCmd.append("   )b,");
        sqlCmd.append("   (select bank_code,");//機構代號
        sqlCmd.append("       bn01.bank_name,");//機構名稱     
        sqlCmd.append("       sum(decode(acc_code,'field_over_rate',amt,'')) as field_over_rate,");//逾放比     
        sqlCmd.append("       sum(decode(acc_code,'field_backup_over_rate',amt,'')) as field_backup_over_rate,");//備呆覆蓋率      
        sqlCmd.append("       sum(decode(acc_code,'field_backup_credit_rate',amt,'')) as field_backup_credit_rate ");//放款覆蓋率
        sqlCmd.append("      from a01_operation left join (select * from bn01 where m_year=?)bn01 on a01_operation.bank_code=bn01.bank_no ");
        sqlCmd.append("     where a01_operation.m_year=? ");
        sqlCmd.append("       and m_month=? "); 
        paramList.add(wlx01_m_year);
        paramList.add(s_year);
        paramList.add(s_month);
        sqlCmd.append("     group by bank_code,bank_name ");
        sqlCmd.append("   )c ");  
        sqlCmd.append("where a1.bank_code =b.bank_code ");
        sqlCmd.append("  and a1.bank_code = c.bank_code(+) ");
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                                         "bank_code,bank_name,wr_count,sort1,field_wr1,sort2,field_wr2,"
                                        +"field_over_rate,field_backup_over_rate,field_backup_credit_rate");
        System.out.println("getRemarkList1.size()="+dbData.size()); 
        return dbData;
    }
   	
   //各項警示指標:附表4/附表7[3項警示指標]查詢SQL
   public static List getRemarkList2(String wlx01_m_year,String s_year,String s_month,String unit,String wr_rpt,String table){
       StringBuffer sqlCmd = new StringBuffer();  
       List paramList = new ArrayList();
       String code1="";
       String code2="";
       String code3="";
       int serial1 = 0;
       int serial2 = 0;
       int serial3 = 0;
       sqlCmd.append("select a1.bank_code,a1.bank_name,");
       sqlCmd.append("       b.wr_count,");//警示項目加總
       sqlCmd.append("       sort1,");//第一項指標排名
       sqlCmd.append("       field_wr1,");//第一項指標
       sqlCmd.append("       sort2,");//第二項指標排名
       sqlCmd.append("       field_wr2,");//第二項指標
       sqlCmd.append("       sort3,");//第三項指標排名
       sqlCmd.append("       field_wr3,");//第三項指標      
       sqlCmd.append("       c.field_992710_990230_rate,");//建築放款/上年度信用部決算淨值):附表4使用
       sqlCmd.append("       c.field_over_rate ");//逾放比:附表7使用
       sqlCmd.append(" from ");
       sqlCmd.append("    (select bank_code,bank_name,");
       sqlCmd.append("            decode(sum(sort1),'','',sum(sort1)) as sort1,");
       sqlCmd.append("            sum(decode(field_wr1,null,'',field_wr1)) as field_wr1,");
       sqlCmd.append("            decode(sum(sort2),'','',sum(sort2)) as sort2,");
       sqlCmd.append("            sum(decode(field_wr2,null,'',field_wr2)) as field_wr2,");
       sqlCmd.append("            decode(sum(sort3),'','',sum(sort3)) as sort3,");
       sqlCmd.append("            sum(decode(field_wr3,null,'',field_wr3)) as field_wr3 ");
       sqlCmd.append("       from ");             
       sqlCmd.append("         (select a.bank_code,a.bank_name,");             
       sqlCmd.append("                 sort1,");//第一項指標排名
       sqlCmd.append("                 field_wr1,");//第一項指標
       sqlCmd.append("                 sort2,");//第二項指標排名
       sqlCmd.append("                 field_wr2,");//第二項指標 
       sqlCmd.append("                 sort3,");//第三項指標排名
       sqlCmd.append("                 field_wr3 ");//第三項指標
       sqlCmd.append("            from (select bank_code,bank_name,sort1, field_wr1,sort2, field_wr2,sort3, field_wr3 ");
       sqlCmd.append("                    from (select t.*,rownum as sort1,null as sort2,null as sort3 ");//該警示指標排名     
       sqlCmd.append("                            from (select * ");
       sqlCmd.append("                                    from (select bank_code,");//機構代號         
       sqlCmd.append("                                                 bn01.bank_name,");//機構名稱
       if("4".equals(table)){
           sqlCmd.append("                                             sum(decode(acc_code,'field_992710_rate',decode(wr_range_serial,11,amt,''),'')) as field_wr1,");//建築放款.增加比率前5名:附表4,才需加入
           code1="field_992710_rate";
           serial1=11;
       }else if("7".equals(table)){
           sqlCmd.append("                                             sum(decode(acc_code,'field_diff_992610_cal',decode(wr_range_serial,23,round(amt /?,0),''),'')) as field_wr1,");//應予觀察放款.增加金額前5名:附表7,才需加入
           paramList.add(unit);
           code1="field_diff_992610_cal";
           serial1=23;
       }
       sqlCmd.append("                                                 null as field_wr2,null as field_wr3 ");  
       sqlCmd.append("                                            from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("                                           where wr_operation.m_year=? ");
       sqlCmd.append("                                             and m_month=? ");
       sqlCmd.append("                                             and wr_rpt=? "); //0:與上月比較/1:與上季比較/2:與上年度同期比較
       sqlCmd.append("                                             and acc_code=? ");//依不同表格,變動變數設定
       sqlCmd.append("                                             and wr_range_serial=").append(serial1);//依不同表格,變動變數設定
       sqlCmd.append("                                             and warn_type=? "); 
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code1);
       paramList.add("Y");
       sqlCmd.append("                                           group by bank_code,bank_name ");
       sqlCmd.append("                                          ) ");
       sqlCmd.append("                                    order by field_wr1 desc ");
       sqlCmd.append("                                  )t ");
       sqlCmd.append("                           union all ");
       sqlCmd.append("                          select t1.*,null as sort1,rownum as sort2,null as sort3 ");
       sqlCmd.append("                            from(select * ");
       sqlCmd.append("                      from (select bank_code,");//機構代號
       sqlCmd.append("                               bn01.bank_name,");//機構名稱     
       sqlCmd.append("                               null as field_wr1,");
       if("4".equals(table)){
           sqlCmd.append("                           sum(decode(acc_code,'field_992710_990230_rate',decode(wr_range_serial,13,amt,''),'')) as field_wr2,");//建築放款/上年度信用部決算淨值) 上月沒有.本月有.才顯示>=100%:附表4,才需加入
           code2="field_992710_990230_rate";
           serial2=13;
       }else if("7".equals(table)){
           sqlCmd.append("                           sum(decode(acc_code,'field_diff_992630',decode(wr_range_serial,25,round(amt /?,0),''),'')) as field_wr2,");//應予觀察放款-非會員.增加金額前5名:附表7,才需加入
           paramList.add(unit);
           code2="field_diff_992630";
           serial2=25;
       }
       sqlCmd.append("                               null as field_wr3 "); 
       sqlCmd.append("                          from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("                         where wr_operation.m_year=? ");
       sqlCmd.append("                           and m_month=? ");
       sqlCmd.append("                           and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較
       sqlCmd.append("                           and acc_code=? ");//依不同表格,變動變數設定
       sqlCmd.append("                           and wr_range_serial=").append(serial2);//依不同表格,變動變數設定
       sqlCmd.append("                           and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code2);
       paramList.add("Y");
       sqlCmd.append("                         group by bank_code,bank_name ");
       sqlCmd.append("                         ) ");
       sqlCmd.append("                       order by field_wr2 desc ");
       sqlCmd.append("                     )t1 ");
       sqlCmd.append("                   union all ");
       sqlCmd.append("                  select t2.*,null as sort1,null as sort2,rownum as sort3 ");
       sqlCmd.append("                    from (select * ");
       sqlCmd.append("                            from (select bank_code,");//機構代號
       sqlCmd.append("                             bn01.bank_name,");//機構名稱
       sqlCmd.append("                             null as field_wr1, null as field_wr2,");   
       if("4".equals(table)){
           sqlCmd.append("                         sum(decode(acc_code,'field_diff_992710_990230_rate',decode(wr_range_serial,14,amt,''),'')) as field_wr3 ");//本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)增加百分點前5名:附表4,才需加入
           code3="field_diff_992710_990230_rate";
           serial3=14;
       }else if("7".equals(table)){
           sqlCmd.append("                         sum(decode(acc_code,'field_diff_992730',decode(wr_range_serial,26,round(amt /?,0),''),'')) as field_wr3 ");//應予觀察放款-建築放款.增加金額前5名:附表7,才需加入
           paramList.add(unit);
           code3="field_diff_992730";
           serial3=26;
       }
       sqlCmd.append("                            from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("                           where wr_operation.m_year=? ");
       sqlCmd.append("                             and m_month=? ");
       sqlCmd.append("                             and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較
       sqlCmd.append("                             and acc_code=? ");//依不同表格,變動變數設定
       sqlCmd.append("                             and wr_range_serial=").append(serial3);//依不同表格,變動變數設定
       sqlCmd.append("                             and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code3);
       paramList.add("Y");
       sqlCmd.append("                           group by bank_code,bank_name ");
       sqlCmd.append("                         ) ");
       sqlCmd.append("                       order by field_wr3 desc ");
       sqlCmd.append("                      )t2 "); 
       sqlCmd.append("         )order by bank_code ");
       sqlCmd.append("             )a ");
       sqlCmd.append("       )group by bank_code,bank_name ");
       sqlCmd.append("       )a1,");
       sqlCmd.append("       (select bank_code,");//機構代號
       sqlCmd.append("           bn01.bank_name,");//機構名稱
       sqlCmd.append("           count(*) as wr_count ");
       sqlCmd.append("      from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("     where wr_operation.m_year=?");
       sqlCmd.append("       and m_month=? ");
       sqlCmd.append("       and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較 
       sqlCmd.append("       and warn_type=? "); 
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add("Y");
       sqlCmd.append("         group by bank_code,bank_name ");
       sqlCmd.append("         order by bank_code ");
       sqlCmd.append("       )b,");
       sqlCmd.append("       (select bank_code,");//機構代號
       sqlCmd.append("               bn01.bank_name,");//機構名稱            
       sqlCmd.append("               sum(decode(acc_code,'field_992710_990230_rate',amt,'')) as field_992710_990230_rate,");//建築放款/上年度信用部決算淨值)
       sqlCmd.append("               sum(decode(acc_code,'field_over_rate',amt,'')) as field_over_rate ");//逾放比     
       sqlCmd.append("      from (select m_year,m_month,bank_code,acc_code,amt from wr_operation where m_year=? and m_month=? and wr_rpt=? ");//0:與上月比較/1:與上季比較/2:與上年度同期比較
       sqlCmd.append("             union ");
       sqlCmd.append("            select m_year,m_month,bank_code,acc_code,amt from a01_operation where m_year=? and m_month=? ");
       sqlCmd.append("           )wr_operation ");
       sqlCmd.append("      left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("     group by bank_code,bank_name ");
       sqlCmd.append("       )c "); 
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wlx01_m_year);
       sqlCmd.append(" where a1.bank_code =b.bank_code ");
       sqlCmd.append("  and a1.bank_code = c.bank_code(+) ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,bank_name,wr_count,sort1,field_wr1,sort2,field_wr2,sort3,field_wr3,field_992710_990230_rate,field_over_rate");
       System.out.println("dbData_getRemarkList2.size()="+dbData.size()); 
       return dbData;
   }
   //專案農貸:附表2 [2項警示指標]查詢SQL
   public static List getRemarkList3(String wlx01_m_year,String s_year,String s_month,String unit,String wr_rpt,String table){
       StringBuffer sqlCmd  = new StringBuffer(); 
       List paramList = new ArrayList();
       String code="";
       int serial = 0;
       sqlCmd.append("select a1.bank_code,a1.bank_name,"); 
       sqlCmd.append("       b.wr_count,");//警示項目加總
       sqlCmd.append("       to_number(sort1) sort1,");//第一項指標排名
       sqlCmd.append("       to_number(field_wr1) field_wr1,");//第一項指標
       sqlCmd.append("       to_number(sort2) sort2,");//第二項指標排名
       sqlCmd.append("       to_number(a1.field_wr2) field_wr2 ");//第二項指標                
       sqlCmd.append("from ");
       sqlCmd.append("(select bank_code,bank_name,");
       sqlCmd.append("        decode(sum(sort1),'','',sum(sort1)) as sort1,");
       sqlCmd.append("        sum(decode(field_wr1,null,'',field_wr1)) as field_wr1,");
       sqlCmd.append("        decode(sum(sort2),'','',sum(sort2)) as sort2,");
       sqlCmd.append("        sum(decode(field_wr2,null,'',field_wr2)) as field_wr2 ");
       sqlCmd.append(" from ");
       sqlCmd.append(" (select a.bank_code,a.bank_name,");       
       sqlCmd.append("         sort1,");//警示項目1排名
       sqlCmd.append("         field_wr1,");//警示項目1
       sqlCmd.append("         sort2,");//警示項目2排名
       sqlCmd.append("         field_wr2 ");//警示項目1
       sqlCmd.append("  from ");             
       sqlCmd.append("  ( ");
       sqlCmd.append("  select bank_code,bank_name,sort1,field_wr1,sort2,field_wr2 ");
       sqlCmd.append("  from ");
       sqlCmd.append("  (  ");
       sqlCmd.append("    select t.*,rownum as sort1,null as sort2 ");//--該警示指標排名
       sqlCmd.append("    from(select * from ( ");
       sqlCmd.append("         select bank_code,");//--機構代號
       sqlCmd.append("                bn01.bank_name,");//--機構名稱
       if("2".equals(table)){
           sqlCmd.append("            sum(decode(acc_code,'field_diff_over6m_loan_bal_amt',decode(wr_range_serial,2,round(amt /?,0),''),'')) as field_wr1,");//--專案農貸逾期放款.增加金額前5名: 附表2,才需加入 
           paramList.add(unit);
           code="field_diff_over6m_loan_bal_amt";
           serial=2;
       }
       sqlCmd.append("                null as field_wr2 ");
       sqlCmd.append("           from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("          where wr_operation.m_year=? ");
       sqlCmd.append("            and m_month=? ");
       sqlCmd.append("            and wr_rpt=? ");//3:與上月比較/4:與上季比較/5:與上年度同期比較
       sqlCmd.append("            and acc_code=? ");//依不同表格,變動變數設定
       sqlCmd.append("            and wr_range_serial=").append(serial);//依不同表格,變動變數設定
       sqlCmd.append("            and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code);
       paramList.add("Y");
       sqlCmd.append("          group by bank_code,bank_name ");
       sqlCmd.append("         )order by field_wr1 desc ");
       sqlCmd.append("    )t ");
       sqlCmd.append("    union all ");
       sqlCmd.append("    select t1.*,null,rownum as sort2 ");
       sqlCmd.append("    from(select * from ( ");
       sqlCmd.append("         select bank_code,");//--機構代號
       sqlCmd.append("                bn01.bank_name,");//--機構名稱
       sqlCmd.append("                null as field_wr1,");
       if("2".equals(table)){
           sqlCmd.append("            sum(decode(acc_code,'field_diff_over6m_loan_rate',decode(wr_range_serial,3,amt,''),'')) as field_wr2 ");//--專案農貸逾放比率.增加百分點前5名:附表2,才需加入
           code="field_diff_over6m_loan_rate";
           serial=3;
       }
       sqlCmd.append("           from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("          where wr_operation.m_year=? ");
       sqlCmd.append("            and m_month=? ");
       sqlCmd.append("            and wr_rpt=? ");//--3:與上月比較/4:與上季比較/5:與上年度同期比較
       sqlCmd.append("            and acc_code=? ");//--依不同表格,變動變數設定
       sqlCmd.append("            and wr_range_serial=").append(serial);//--依不同表格,變動變數設定
       sqlCmd.append("            and warn_type=? "); 
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code);
       paramList.add("Y");
       sqlCmd.append("          group by bank_code,bank_name ");
       sqlCmd.append("         )order by field_wr2 desc ");
       sqlCmd.append("    )t1 ");
       sqlCmd.append("  )order by bank_code ");
       sqlCmd.append("  )a ");
       sqlCmd.append(" )group by bank_code,bank_name ");
       sqlCmd.append(")a1,");
       sqlCmd.append("(   ");
       sqlCmd.append("select bank_code,");//--機構代號
       sqlCmd.append("       bn01.bank_name,");//--機構名稱
       sqlCmd.append("       count(*) as wr_count ");
       sqlCmd.append("  from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append(" where wr_operation.m_year=? ");
       sqlCmd.append("   and m_month=? ");
       sqlCmd.append("   and wr_rpt=? ");//--3:與上月比較/4:與上季比較/5:與上年度同期比較
       sqlCmd.append("   and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add("Y");
       sqlCmd.append(" group by bank_code,bank_name ");
       sqlCmd.append(" order by bank_code ");
       sqlCmd.append(")b ");
       sqlCmd.append("where a1.bank_code =b.bank_code ");
        
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,bank_name,wr_count,sort1,field_wr1,sort2,field_wr2");
       System.out.println("dbData_getRemarkList3.size()="+dbData.size()); 
       return dbData;
   }
   //專案農貸:附表1/附表3 [1項警示指標]查詢SQL
   public static List getRemarkList4(String wlx01_m_year,String s_year,String s_month,String unit,String wr_rpt,String table){
       StringBuffer sqlCmd  = new StringBuffer(); 
       List paramList = new ArrayList();
       String code="";
       int serial = 0;
       sqlCmd.append("select a1.bank_code,a1.bank_name,"); 
       sqlCmd.append("       b.wr_count,");//警示項目加總
       sqlCmd.append("       sort1,");//第一項指標排名
       sqlCmd.append("       field_wr1 ");//第一項指標                    
       sqlCmd.append("from ");
       sqlCmd.append(" (select bank_code,bank_name,");       
       sqlCmd.append("         sort1,");//警示項目1排名
       sqlCmd.append("         field_wr1 ");//警示項目1         
       sqlCmd.append("  from ");
       sqlCmd.append("  (    ");
       sqlCmd.append("    select t.*,rownum as sort1 ");//該警示指標排名
       sqlCmd.append("    from(select * from ( ");
       sqlCmd.append("           select bank_code,");//機構代號
       sqlCmd.append("                  bn01.bank_name,");//機構名稱
       if("1".equals(table)){
           sqlCmd.append("              sum(decode(acc_code,'field_diff_loan_bal_amt',decode(wr_range_serial,1,round(amt /?,0),''),'')) as field_wr1 ");//專案農貸放款餘額.增加金額前5名:附表1,才需加入
           paramList.add(unit);
           code="field_diff_loan_bal_amt";
           serial=1;
       }else if("3".equals(table)){
           sqlCmd.append("              sum(decode(acc_code,'field_delay_loan_rate',decode(wr_range_serial,4,amt,''),'')) as field_wr1 ");//專案農貸逾放比率.增加百分點前5名:附表3,才需加入
           code="field_delay_loan_rate";
           serial=4;
       }
       sqlCmd.append("             from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append("            where wr_operation.m_year=? ");
       sqlCmd.append("              and m_month=? ");
       sqlCmd.append("              and wr_rpt=? ");//3:與上月比較/4:與上季比較/5:與上年度同期比較
       sqlCmd.append("              and acc_code=? ");//依不同表格,變動變數設定
       sqlCmd.append("              and wr_range_serial=").append(serial);//依不同表格,變動變數設定
       sqlCmd.append("              and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add(code);
       paramList.add("Y");
       sqlCmd.append("            group by bank_code,bank_name ");
       sqlCmd.append("         )order by field_wr1 desc ");
       sqlCmd.append("    )t ");    
       sqlCmd.append("  )order by bank_code ");
       sqlCmd.append(")a1,");
       sqlCmd.append("(   ");
       sqlCmd.append("select bank_code,");//機構代號
       sqlCmd.append("       bn01.bank_name,");//機構名稱
       sqlCmd.append("       count(*) as wr_count ");//警示項目加總
       sqlCmd.append("  from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
       sqlCmd.append(" where wr_operation.m_year=? ");
       sqlCmd.append("   and m_month=? ");
       sqlCmd.append("   and wr_rpt=? ");//3:與上月比較/4:與上季比較/5:與上年度同期比較
       sqlCmd.append("   and warn_type=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(wr_rpt);
       paramList.add("Y");
       sqlCmd.append(" group by bank_code,bank_name ");
       sqlCmd.append(" order by bank_code ");
       sqlCmd.append(")b ");
       sqlCmd.append("where a1.bank_code =b.bank_code ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,bank_name,wr_count,sort1,field_wr1");
       System.out.println("dbData_getRemarkList4.size()="+dbData.size()); 
       return dbData;
   }
   private static void insertCell(String value,HSSFWorkbook wb,HSSFRow row,int i,short topBorder,short bottomBorder,short leftBorder,short rightBorder, short alignment,boolean warptext){
           HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
           HSSFCellStyle cs1 = wb.createCellStyle();
           //HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
           //設置邊框
           cs1.setBorderTop(topBorder); //上邊框
           cs1.setBorderBottom(bottomBorder); //下邊框
           cs1.setBorderLeft(leftBorder); //左邊框
           cs1.setBorderRight(rightBorder); //右邊框
           cs1.setWrapText(warptext);//自動換行
           cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
           cs1.setAlignment(alignment);//水平置中
           cell.setCellStyle(cs1);
           cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
           cell.setCellValue(value);
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

