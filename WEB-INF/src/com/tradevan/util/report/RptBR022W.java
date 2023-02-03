// 94.04.18 fix 剔除已裁撤機構 by 2295
// 99.04.08 fix 因應縣市合併調整SQL by 2808
//100.06.30 fix 地方主管機關.農/漁會人數重覆計算問題,調整SQL;列印福建省合計  by 2295
//102.07.01 add 操作歷程寫入log by2968
//102.11.25 fix 100年後 1.縣市別省農會更名其他  2.add填表說明  by 2968
//103.02.18 fix 台灣省合計改成臺灣省合計 by 2295
//103.12.24 fix 桃園縣升格調整 by 2968
//109.07.03 add 增加中華民國農會  by 2295
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import javax.servlet.http.HttpServletRequest;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptBR022W {
    public static String createRpt(String datestate,String lguser_id,String ipAddress) {
		String errMsg = "";
		List <DataObject>dbData = null;
		long sumB=0,sum6=0,sum7=0,sumA=0;
		long sumAB=0,sumA6=0,sumA7=0,sumAA=0;		
		try{	
			int m_year = 99  ; //判斷縣市合併用參數
			String simpleNm = "" ;
			String cd01Table = "cd01" ;
			Calendar rightNow = Calendar.getInstance();
            String year = String.valueOf(rightNow.get(Calendar.YEAR)-1911);
            String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);
            String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
			if(Integer.parseInt(year) >=100 ) {
				m_year = 100 ;
				simpleNm = "台閩地區農業金融從業人員統計_100.xls" ;
			}else {
				m_year = 99 ;
				cd01Table = "cd01_99" ;
				simpleNm = "台閩地區農業金融從業人員統計_99.xls" ;
			}
			
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
    		System.out.println("open 範本路徑:"+xlsDir + System.getProperty("file.separator")+ simpleNm) ;
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ simpleNm );
			//System.out.println("Open excel 完成");
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔失敗");} //else System.out.println("open 範本檔成功");
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}  //else System.out.println("open 工作表 成功");
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		if(sheet==null){System.out.println("open sheet 失敗");}//else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )76 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			insertCell(dbData,false,0,"中華民國　"+year+"年　"+month+"月",wb,row,(short)0);
			
			//加上列印日期
			if(datestate.equals("1")){
				row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
				insertCell(dbData,false,0,"列印日期："+year+"年"+month+"月"+day+"日",wb,row,(short)0);
            }
			row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
			insertCell(dbData,false,0,year+"年"+month+"月",wb,row,(short)2);
			
			dbData = (List<DataObject>) getSqlQueryInfo(m_year,cd01Table) ;
			System.out.println("layer 1 dbData.size()="+dbData.size());
			InsertWlXOPERATE_LOG(lguser_id,"BR022W",ipAddress,"P");
			int j=7;
			short top=0,down=0;
			String [] cityLs = {"h","f","A","H","E","b","d","e","W","Z"}; //直轄市列印順序.109.07.03 增加中華民國農會
			HashMap<String,DataObject> map = new HashMap<String,DataObject>() ;
			for(DataObject b : dbData ) {
				//1.列印非直轄市
				if(!"2".equals(Utility.getTrimString(b.getValue("hsien_div")))){
					map.put(Utility.getTrimString(b.getValue("hsien_id")),b);
				}else{
					long numA= b.getValue("total")==null? 0l : Long.parseLong(b.getValue("total").toString()) ;
	  				long numB= b.getValue("type_b")==null? 0l : Long.parseLong(b.getValue("type_b").toString()) ;
	  				long num6= b.getValue("type_6")==null? 0l : Long.parseLong(b.getValue("type_6").toString()) ;
	  				long num7= b.getValue("type_7")==null? 0l : Long.parseLong(b.getValue("type_7").toString()) ;
	  				String hsien_name = Utility.getTrimString(b.getValue("hsien_name"));
	  				sumAA+=numA;
	  				sumAB+=numB;
	  				sumA6+=num6;
	  				sumA7+=num7;
	  				sumA+=numA;
  					sumB+=numB;
  					sum6+=num6;
  					sum7+=num7;
  					row=(sheet.getRow(i+j)==null)? sheet.createRow(i+j) : sheet.getRow(i+j);
  					cell=row.getCell((short)0);
  					cell.setCellValue(String.valueOf(i+1));
  					cell=row.getCell((short)1);
  					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  					cell.setCellValue(hsien_name);
  					insertCell(row,numA,numB,num6,num7);
  					i++ ;
				}
			}
			//台灣省合計===================
			row=(sheet.getRow(i+j)==null)? sheet.createRow(i+j) : sheet.getRow(i+j);
			sheet.addMergedRegion(new Region((i+j), (short) (0),(i+j), (short) (1)));
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("臺灣省合計");
			insertCell(row,sumA,sumB,sum6,sum7);
			i++ ;
			//直轄市統計====================
			for(String hsien_id : cityLs) {
				DataObject b = map.get(hsien_id) ;
				if(b==null) continue ;
				
				long numA= b.getValue("total")==null? 0l : Long.parseLong(b.getValue("total").toString()) ;
  				long numB= b.getValue("type_b")==null? 0l : Long.parseLong(b.getValue("type_b").toString()) ;
  				long num6= b.getValue("type_6")==null? 0l : Long.parseLong(b.getValue("type_6").toString()) ;
  				long num7= b.getValue("type_7")==null? 0l : Long.parseLong(b.getValue("type_7").toString()) ;
  				/*@param srow 指定哪一列, 由 0 開始計算
  			    @param scolumn 指定哪一行, 由 0 開始計算
  			    @param drow 指定哪一列, 由 0 開始計算
  			    @param dcolumn 指定哪一行, 由 0 開始計算*/
  				sumA=0;sumB=0;sum6=0;sum7=0;
				row=(sheet.getRow(i+j)==null)? sheet.createRow(i+j) : sheet.getRow(i+j);
				sheet.addMergedRegion(new Region((i+j), (short) (0),(i+j), (short) (1)));
				cell=row.getCell((short)0);
				HSSFCellStyle cellStyle = cell.getCellStyle();
				cellStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
				cell.setCellStyle(cellStyle) ;
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				if(("h").equals(hsien_id)){//109.07.03 h其他,改成中華民國農會
				  cell.setCellValue("中華民國農會");
				}else if(("W").equals(hsien_id) || ("Z").equals(hsien_id)){//100.06.30 W金門縣.Z連江縣
				   cell.setCellValue(Utility.getTrimString(b.getValue("hsien_name")));
				}else{
				   cell.setCellValue(Utility.getTrimString(b.getValue("hsien_name"))+"合計");
				}
				insertCell(row,numA,numB,num6,num7);
  				i++ ;
			}
			long b1=0l,b2=0l,b3=0l,b4 =0l;
			DataObject bean = null;
			for(int idx=dbData.size()-2;idx<dbData.size();idx++) {
				bean = (DataObject)dbData.get(idx);
				b1 += bean.getValue("total")==null? 0l : Long.parseLong(bean.getValue("total").toString()) ;
	  			b2 += bean.getValue("type_b")==null? 0l : Long.parseLong(bean.getValue("type_b").toString()) ;
	  			b3 += bean.getValue("type_6")==null? 0l : Long.parseLong(bean.getValue("type_6").toString()) ;
	  			b4 += bean.getValue("type_7")==null? 0l : Long.parseLong(bean.getValue("type_7").toString()) ;
			}
			//福建省合計================
			row=(sheet.getRow(i+j)==null)? sheet.createRow(i+j) : sheet.getRow(i+j);
			sheet.addMergedRegion(new Region((i+j), (short) (0),(i+j), (short) (1)));
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("福建省合計");
			insertCell(row,b1,b2,b3,b4);//100.06.30
			i++ ;
			b1=0l;b2=0l;b3=0l;b4 =0l;
			for(DataObject b : dbData ) {
				b1 += b.getValue("total")==null? 0l : Long.parseLong(b.getValue("total").toString()) ;
	  			b2 += b.getValue("type_b")==null? 0l : Long.parseLong(b.getValue("type_b").toString()) ;
	  			b3 += b.getValue("type_6")==null? 0l : Long.parseLong(b.getValue("type_6").toString()) ;
	  			b4 += b.getValue("type_7")==null? 0l : Long.parseLong(b.getValue("type_7").toString()) ;
			}
			//總計======================
			row=(sheet.getRow(i+j)==null)? sheet.createRow(i+j) : sheet.getRow(i+j);
			sheet.addMergedRegion(new Region((i+j), (short) (0),(i+j), (short) (1)));
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("總計");
			insertCell(row,b1,b2,b3,b4);
  			

	       	System.out.println(reportDir + System.getProperty("file.separator")+"台閩地區農業金融從業人員統計.xls");
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"台閩地區農業金融從業人員統計.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
    
	
	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j)
	{
		String insertValue="";
  		if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue= Item;
		System.out.println("insertValue="+insertValue);
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j); 
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	    double value=0;
	    try{
	    	cell.setCellValue(Double.parseDouble(insertValue));
	    }catch(NumberFormatException e){
	    	cell.setCellValue(insertValue);
	    }
	}
	
	public static void insertCell(HSSFRow row,long numA,long numB,long num6,long num7)
	{
		System.out.println("row index="+row.getRowNum()) ;
  		HSSFCell cell=(row.getCell((short)2)==null)? row.createCell((short)2) : row.getCell((short)2); 
  		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
  		cell.setCellValue((double)numA);

  		cell=(row.getCell((short)3)==null)? row.createCell((short)3) : row.getCell((short)3); 
  		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
  		cell.setCellValue((double)numB);

  		cell=(row.getCell((short)4)==null)? row.createCell((short)4) : row.getCell((short)4); 
  		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
  		cell.setCellValue((double)num6);

  		cell=(row.getCell((short)5)==null)? row.createCell((short)5) : row.getCell((short)5); 
  		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
  		cell.setCellValue((double)num7);
	}
	
	/***
	 * 取得查詢資料.
	 * 
	 * @param m_year
	 * @param cd01Table
	 * @return
	 * @throws Exception
	 */
	private static List<DataObject> getSqlQueryInfo(int m_year,String cd01Table ) throws Exception{
		StringBuffer sql = new StringBuffer() ;
		List paramList = new ArrayList() ;
		sql.append("SELECT hsien_name            , ");
		sql.append("       b.hsien_id AS hsien_id, ");
		sql.append("       SUM(DECODE(BANK_TYPE, ");
		sql.append("                  'B',STAFF_NUM, ");
		sql.append("                  '6',STAFF_NUM, ");
		sql.append("                  '7',STAFF_NUM, ");
		sql.append("                  0) ");
		sql.append("       ) AS total, ");
		sql.append("       SUM(DECODE(BANK_TYPE, ");
		sql.append("                  'B',STAFF_NUM, ");
		sql.append("                  0) ");
		sql.append("       ) AS type_b, ");
		sql.append("       SUM(DECODE(BANK_TYPE, ");
		sql.append("                  '6',STAFF_NUM, ");
		sql.append("                  0) ");
		sql.append("       ) AS type_6, ");
		sql.append("       SUM(DECODE(BANK_TYPE, ");
		sql.append("                  '7',STAFF_NUM, ");
		sql.append("                  0) ");
		sql.append("       ) AS type_7, ");
		sql.append("       FR001W_OUTPUT_ORDER as output_order ");
		sql.append(" ,a.HSIEN_DIV ");
		sql.append("FROM   ").append(cd01Table).append(" a ");
		sql.append("       LEFT JOIN ");
		sql.append("              (SELECT  hsien_id , ");
		sql.append("                       staff_num, ");
		sql.append("                       bank_type ");
		sql.append("              FROM     ( ");
		sql.append("                       (SELECT hsien_id  AS hsien_id , ");
		sql.append("                               staff_num AS staff_num, ");
		sql.append("                               bank_type AS bank_type ");
		sql.append("                       FROM    (select * from bn01 where m_year=?)a, ");
		sql.append("                               (select * from wlx01 where m_year=?)b ");
		sql.append("                       WHERE   a.bank_type IN (?, ");
		sql.append("                                               ?) ");
		sql.append("                       AND     a.bank_no = b.bank_no ");
		//sql.append("                       AND     a.m_year  =? ");
		sql.append("                       AND ");
		sql.append("                               ( ");
		sql.append("                                       b.CANCEL_NO      <> ? ");
		sql.append("                               OR      b.CANCEL_NO IS NULL ");
		sql.append("                               ) ");
		sql.append("                       ) ");
		sql.append("                ");
		sql.append("               UNION ALL ");
		sql.append("                         (SELECT hsien_id  AS hsien_id , ");
		sql.append("                                 staff_num AS staff_num, ");
		sql.append("                                 bank_type AS bank_type ");
		sql.append("                         FROM    (select * from ba01 where m_year=?)a, ");
		sql.append("                                 (select * from wlx02 where m_year=?)b ");
		sql.append("                         WHERE   a.bank_type IN (?, ");
		sql.append("                                                 ?) ");
		sql.append("                         AND     a.bank_kind = ? ");
		sql.append("                         AND     a.bank_no   = b.bank_no ");
		//sql.append("                         AND     a.m_year    = ? ");
		sql.append("                         AND ");
		sql.append("                                 ( ");
		sql.append("                                         b.CANCEL_NO      <> ? ");
		sql.append("                                 OR      b.CANCEL_NO IS NULL ");
		sql.append("                                 ) ");
		sql.append("                         ) ");
		sql.append("                  ");
		sql.append("                 UNION ALL ");
		sql.append("                           ( SELECT hsien_id       AS hsien_id , ");
		sql.append("                                   business_person AS staff_num, ");
		sql.append("                                   bank_type       AS bank_type ");
		sql.append("                           FROM    (select * from ba01 where m_year=?)a, ");
		sql.append("                                   (select * from bank_cmml where m_year=?)b ");
		sql.append("                           WHERE   a.bank_type = ? ");
		sql.append("                           AND     a.bank_kind = ? ");
		sql.append("                           AND     a.bank_no   = b.bank_no ");
		//sql.append("                           AND     a.m_year    = ? ");
		sql.append("                           ) ");
		sql.append("                       ) ");
		sql.append("              ) ");
		sql.append("              b ");
		sql.append("     ON       a.hsien_id=b.hsien_id ");
		sql.append("WHERE    a.hsien_id    <> 'Y' ");
		sql.append("GROUP BY hsien_name, ");
		sql.append("         b.hsien_id, ");
		sql.append("         FR001W_OUTPUT_ORDER,a.HSIEN_DIV  ");
		sql.append("ORDER BY FR001W_OUTPUT_ORDER ");
		paramList.add(m_year);
		paramList.add(m_year);
		paramList.add("6");
		paramList.add("7");
		
		paramList.add("Y");
		paramList.add(m_year);
		paramList.add(m_year);
		paramList.add("6");
		paramList.add("7");
		paramList.add("1");
		
		paramList.add("Y");
		paramList.add(m_year);
		paramList.add(m_year);
		paramList.add("B");
		paramList.add("0");
		
		
		return (List<DataObject>) DBManager.QueryDB_SQLParam(sql.toString(),paramList, "total,type_b,type_6,type_7,HSIEN_DIV,hsien_id,hsien_name"); 
	}
	public static String InsertWlXOPERATE_LOG(String lguser_id,String program_id,String ipAddress,String update_type) throws Exception{     
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList() ;
        String errMsg="";
        try {
            sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,update_type)");
            sqlCmd.append("                     VALUES(?,sysdate,?,?,?) ");
            paramList.add(lguser_id);
            paramList.add(program_id);
            paramList.add(ipAddress);
            paramList.add(update_type);//操作類別 I-新增，U-異動，D-刪除，Q-明細，P-列印
            if(updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
                errMsg = errMsg + "相關資料寫入資料庫成功";                    
            }else{
                errMsg = errMsg + "相關資料寫入log失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
            }
        }catch (Exception e){
                System.out.println(e+":"+e.getMessage());
                errMsg = errMsg + "相關資料寫入log失敗";                        
        }   

        return errMsg;
    }
    private static boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
        List updateDBList = new ArrayList();//0:sql 1:data
        List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
        List updateDBDataList = new ArrayList();//儲存參數的List
        
        updateDBDataList.add(paramList);
        updateDBSqlList.add(sql);
        updateDBSqlList.add(updateDBDataList);
        updateDBList.add(updateDBSqlList);
        return DBManager.updateDB_ps(updateDBList) ;
    }
}
