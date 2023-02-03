/*
 94.03.11 Fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
 94.03.15 ADD 月份輸入條件查詢 	by EGG
 94.08.16 fix 更改公式 by 2295
 94.08.30 fix 科目420170取代420700	jwang
 94.09.13 fix 損益科目改本月減上月計算	jwang
 99.09.09 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			  使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 99.09.10 fix 當上月底為12月時,則420100放款利息收入,則不扣除上月底資料
                                 則420170內部融資利息收入.420140/420120統一農(漁)貸利息收入,則不扣除上月底資料 by 2295
 99.09.13 fix 合併RptFR005W_BOAF(申報資料查詢下載) by 2295          
100.10.13 fix 101年.顯示更改後的組織名稱(行政院農業委員會->行政院農業部) by 2295  
103.12.24 fix 桃園縣升格調整的報表     by 2968   
106.08.07 add 農會總表的其他.顯示成中華民國農會 by 2295
108.04.22 add 台灣省合計資料移除中華民國農會 by 2295              
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

public class RptFR005W {
	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type,String rptStyle,String bank_no,String bank_name) {    
		System.out.println("s_month="+s_month);
		String errMsg="";
		String sqlCmd="";
		String unit_name="";
		int rowNum=0;
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String div="2";
		String hsien_id_sum[]={"","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","za","zb"};
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String filename="";
		filename=(bank_type.equals("6"))?"全體農會按縣市別平均利率.xls":"全體漁會按縣市別平均利率.xls";
		StringBuffer field = new StringBuffer();
		StringBuffer fromtable = new StringBuffer();
		StringBuffer sqlCmd_sum_rule = new StringBuffer();//總計/台灣省小計/福建省小計共用sql
		StringBuffer sqlCmd_sum = new StringBuffer();//總計的
		StringBuffer sqlCmd_sumtaiwan = new StringBuffer(); //台灣省小計
		StringBuffer sqlCmd_sumfuchien = new StringBuffer();//福建省小計
		StringBuffer sqlCmd_each = new StringBuffer();  //各縣
		StringBuffer sqlCmd_detail = new StringBuffer();  //各機構明細
		List dbData_sum=null;//總計的
		List dbData_sumtaiwan=null;//台灣省小計
		List dbData_sumfuchien=null;//福建省小計
		List dbData_each=null; //各縣
		List dbData_detail=null;//各機構明細	
		List paramList = new ArrayList();//共同參數
		String cd01_table = "";
        String wlx01_m_year = "";
        //99.09.10 add 查詢年度100年以前.縣市別不同===============================
	    cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	    wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    
		div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
		System.out.println("div="+div);
		if (Integer.parseInt(s_month)==1) {
			s_year_last=String.valueOf(Integer.parseInt(s_year)-1);
			s_month_last="12";
		}else {
			s_year_last=s_year;
			s_month_last=String.valueOf(Integer.parseInt(s_month)-1);
		}
	    System.out.println("s_year="+s_year+":s_month="+s_month+":s_year_last="+s_year_last+":s_month_last="+s_month_last);
		field.append(" round(sum(field_b)/ ?,0) as fieldB , ");//--平均放款總額
		field.append(" round(sum(field_c)/ ?,0) as fieldC , ");//--平均內部融資
		field.append(" round(sum(field_d)/ ?,0) as fieldD , ");//--平均統一農(漁)貸
		field.append(" round(sum(field_e)/ ?,0) as fieldE , ");//--放款利息收入
		field.append(" round(sum(field_f)/ ?,0) as fieldF , ");//--內部融資利息收入
		field.append(" round(sum(field_g)/ ?,0) as fieldG , ");//--統一農(漁)貸利息收入
		field.append(" decode(sum(field_b),0,0,round(sum(field_e) / sum(field_b) *100 ,2))  as fieldH, ");//--平均利率
		field.append(" decode(sum(field_c),0,0,round(sum(field_f) / sum(field_c) *100 ,2))  as fieldI, ");//--內部融資利率
		field.append(" decode(sum(field_d),0,0,round(sum(field_g) / sum(field_d) *100 ,2))  as fieldJ, ");//--統一農貸平均利率
		field.append(" decode(sum(field_r),0,0,round(sum(field_k) / sum(field_r) *100 ,2))  as fieldK, ");//--綜合平均存款利率
		field.append(" decode(sum(field_u),0,0,round(sum(field_l) / sum(field_u) *100 ,2))  as fieldL, ");//--活期存款.平均利率
		field.append(" decode(sum(field_v),0,0,round(sum(field_m) / sum(field_v) *100 ,2))  as fieldM, ");//--活儲存款.平均利率
		field.append(" decode(sum(field_w),0,0,round(sum(field_n) / sum(field_w) *100 ,2))  as fieldN, ");//--員工活期儲蓄存款.平均利率
		field.append(" decode(sum(field_x),0,0,round(sum(field_o) / sum(field_x) *100 ,2))  as fieldO, ");//--定期存款.平均利率
		field.append(" decode(sum(field_y),0,0,round(sum(field_p) / sum(field_y) *100 ,2))  as fieldP, ");//--定期儲蓄存款.平均利率
		field.append(" decode(sum(field_z),0,0,round(sum(field_q) / sum(field_z) *100 ,2))  as fieldQ, ");//--員工定期儲蓄存款.平均利率
		field.append(" round(sum(field_r) / ?,0) as fieldR , ");//--存款總額.合計                                   
		field.append(" round(sum(field_s) / ?,0) as fieldS , ");//--支票存款                                   
		field.append(" round(sum(field_t) / ?,0) as fieldT , ");//--保付支票
		field.append(" round(sum(field_u) / ?,0) as fieldU , ");//--活期存款
		field.append(" round(sum(field_v) / ?,0) as fieldV , ");//--活儲存款
		field.append(" round(sum(field_w) / ?,0) as fieldW , ");//--員工活期儲蓄存款
		field.append(" round(sum(field_x) / ?,0) as fieldX , ");//--定期存款
		field.append(" round(sum(field_y) / ?,0) as fieldY , ");//--定期儲蓄存款
		field.append(" round(sum(field_z) / ?,0) as fieldZ , ");//--員工定期儲蓄存款
		field.append(" round(sum(field_za)/ ?,0) as fieldZA, ");//--公庫存款
		field.append(" round(sum(field_zb)/ ?,0) as fieldZB  ");//--本會支票
		for(int k=1;k<=17;k++){
            paramList.add(unit);
        }
		
		if ("1".equals(rptStyle)) {
		    fromtable.append(" select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
            paramList.add(wlx01_m_year);
            fromtable.append(" left join ( ");
		}
		fromtable.append(" select a01.bank_code,a01.bank_name, ");
		fromtable.append("        field_b,field_c,field_d,field_e,field_f,field_g,field_k,field_l,field_m,field_n,field_o,field_p,");
		fromtable.append("        field_q,field_r,field_s,field_t,field_u,field_v,field_w,field_x,field_y,field_z,field_za,field_zb");
		fromtable.append(" from ( ");
		fromtable.append("        select bank_code,bank_name,");
		fromtable.append("            	 round((sum(decode(acc_code,'120000',amt,'150300',amt,0))- ");
		fromtable.append("               sum(decode(acc_code,'150200',amt,0)))/?,0) as field_b,    ");
		fromtable.append("               round(sum(decode(acc_code,'120700',amt,0))/?,0)  as field_c, ");
		fromtable.append("            	 round((decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(acc_code,'120401',amt,'120402',amt,0)),   ");
		fromtable.append("            		                                            '7',sum(decode(acc_code,'120201',amt,'120202',amt,0)),0),");
		fromtable.append("            		                     '103',sum(decode(acc_code,'120401',amt,'120402',amt,0)),0))/?,0)  as field_d,");
		fromtable.append("               round(sum(decode(acc_code,'220000',amt,0))/?,0) as field_r, ");
		fromtable.append("               round(sum(decode(acc_code,'220100',amt,0))/?,0) as field_s, ");
		fromtable.append("               round(sum(decode(acc_code,'220200',amt,0))/?,0) as field_t, ");
		fromtable.append("               round(sum(decode(acc_code,'220300',amt,0))/?,0) as field_u, ");
		fromtable.append("               round(sum(decode(acc_code,'220400',amt,0))/?,0) as field_v, ");
		fromtable.append("               round(sum(decode(acc_code,'220500',amt,0))/?,0) as field_w, ");
		fromtable.append("               round(sum(decode(acc_code,'220600',amt,0))/?,0) as field_x, ");
		fromtable.append("               round(sum(decode(acc_code,'220700',amt,0))/?,0) as field_y, ");
		fromtable.append("               round(sum(decode(acc_code,'220800',amt,0))/?,0) as field_z, ");
		fromtable.append("               round(sum(decode(acc_code,'220900',amt,0))/?,0) as field_za,");
		fromtable.append("               round(sum(decode(acc_code,'221000',amt,0))/?,0) as field_zb ");
		for(int k=1;k<=14;k++){
            paramList.add(div);
        }
		fromtable.append("        from (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
        fromtable.append("                           WHEN (a01.m_year > 102) THEN '103' ");
        fromtable.append("                           ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 ");
		fromtable.append("        left join (select * from ba01 where m_year=?)ba01 on a01.bank_code = ba01.bank_no ");
		paramList.add(wlx01_m_year);
		fromtable.append("        where (a01.m_year=? and a01.m_month=? or a01.m_year=? and m_month=?) ");
		paramList.add(s_year_last);
		paramList.add(s_month_last);
		paramList.add(s_year);
		paramList.add(s_month);
		fromtable.append("        and acc_code in ('120000','150300','150200','120700','120401','120402','120201','120202',");
		fromtable.append("                         '420100','520100','220000','220100','220200','220300','220400','220500',");
		fromtable.append("                         '220600','220700','220800','220900','221000')");
		fromtable.append("        group by  YEAR_TYPE,bank_type,bank_name,bank_code");
		fromtable.append("        order by bank_code                    ");
		fromtable.append("      )a01,");     
		fromtable.append("      (    ");
		fromtable.append("       select bank_code,bank_name, ");//都是算比率,相加就好不用除2
		fromtable.append("         	    sum(decode(acc_code,'420100',amt,0)) as field_e,");
		fromtable.append("              sum(decode(acc_code,'420170',amt,0)) as field_f,");
		fromtable.append("              decode(bank_type,'6',sum(decode(acc_code,'420140',amt,0)),               ");
		fromtable.append("                               '7',sum(decode(acc_code,'420120',amt,0)),0) as field_g, ");
		fromtable.append("              sum(decode(acc_code,'520100',amt,0)) as field_k,");
		fromtable.append("              sum(decode(acc_code,'520110',amt,0)) as field_l,");
		fromtable.append("              sum(decode(acc_code,'520130',amt,0)) as field_m,");
		fromtable.append("              sum(decode(acc_code,'840640',amt,0)) as field_n,");
		fromtable.append("              sum(decode(acc_code,'520120',amt,0)) as field_o,");
		fromtable.append("              sum(decode(acc_code,'520140',amt,0)) as field_p,");
		fromtable.append("              sum(decode(acc_code,'840670',amt,0)) as field_q ");
		fromtable.append("       from ( ");
		if(s_month_last.equals("12")){//99.09.10 fix 當上月底為12月時,則420100放款利息收入,則不扣除上月底資料
			fromtable.append("             select bank_code,acc_code,decode(acc_code,'420100',0,-amt) as amt ");//前一月份金額為負
		}else{
			fromtable.append("             select bank_code,acc_code,-amt as amt ");//前一月份金額為負	
		}
		fromtable.append("             from a01 ");
		fromtable.append("             where m_year=? and m_month=?");
		paramList.add(s_year_last);
		paramList.add(s_month_last);
		fromtable.append("             and acc_code in ('420100','520100')");
		fromtable.append("             union all ");
		fromtable.append("             select bank_code,acc_code,amt ");//當月份,金額為正
		fromtable.append("             from a01                      ");
		fromtable.append("             where m_year=? and m_month=? ");
		paramList.add(s_year);
		paramList.add(s_month);
		fromtable.append("             and acc_code in ('420100','520100')");
		fromtable.append("             union all                          ");
		if(s_month_last.equals("12")){//99.09.10 fix 當上月底為12月時,則420170內部融資利息收入.420140/420120統一農(漁)貸利息收入,則不扣除上月底資料
			fromtable.append("             select bank_code,acc_code,decode(acc_code,'420170',0,'420120',0,'420140',0,-amt) as amt");
		}else{
			fromtable.append("             select bank_code,acc_code,-amt as amt");
		}
		fromtable.append("             from a03 where m_year=? and m_month=?                          ");
		paramList.add(s_year_last);
		paramList.add(s_month_last);
		fromtable.append("             and acc_code in ('420120','420140','420170','520110','520130','840640','520120','520140','840670')");
		fromtable.append("             union all ");
		fromtable.append("             select bank_code,acc_code,amt ");
		fromtable.append("             from a03 where m_year=? and m_month=? ");
		paramList.add(s_year);
		paramList.add(s_month);
		fromtable.append("             and acc_code in ('420120','420140','420170','520110','520130','840640','520120','520140','840670') ");
		fromtable.append("            )a03 left join (select * from ba01 where m_year=?)ba01 on a03.bank_code = ba01.bank_no ");
		paramList.add(wlx01_m_year);
		fromtable.append("       group by bank_type,bank_name,bank_code ");
		fromtable.append("       order by bank_code ");
		fromtable.append("      )a01_1 ");
		fromtable.append(" where a01.bank_code=a01_1.bank_code ");
	
		
		reportUtil reportUtil=new reportUtil();
		try {
			File xlsDir=new File(Utility.getProperties("xlsDir"));
			File reportDir=new File(Utility.getProperties("reportDir"));
	
			if(!xlsDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
			   		errMsg+=Utility.getProperties("xlsDir")+"目錄新增失敗";
				}
			}
			if(!reportDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
			   		errMsg+=Utility.getProperties("reportDir")+"目錄新增失敗";
				}
			}
			
			String openfile="全體農漁會按縣市別平均利率"+(rptStyle.equals("0")?"_總表"+(Integer.parseInt(s_year)<=99?"_99":(bank_type.equals("6"))?"_農會":"_漁會"):"_明細表")+".xls";
			System.out.println("open file "+openfile);
	
			FileInputStream finput=new FileInputStream(xlsDir+System.getProperty("file.separator")+openfile );
	
		    //設定FileINputStream讀取Excel檔
			POIFSFileSystem fs=new POIFSFileSystem( finput );
			if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
			HSSFWorkbook wb=new HSSFWorkbook(fs);
			if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
			HSSFSheet sheet=wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
			if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
			HSSFPrintSetup ps=sheet.getPrintSetup(); //取得設定
		    //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
		    //sheet.setAutobreaks(true); //自動分頁
	
		    //設定頁面符合列印大小
		    sheet.setAutobreaks( false );
		    //ps.setScale( ( short )65 ); //列印縮放百分比
		    ps.setLandscape( true ); // 設定橫印
		    ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
			
			finput.close();
	
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			//add 參數
			paramList.add(wlx01_m_year);
			paramList.add(bank_type);
            
			//彙整
			if (rptStyle.equals("0")) {
				
				sqlCmd_sum_rule.append(" select");
				sqlCmd_sum_rule.append(field);
				sqlCmd_sum_rule.append(" from ( "); 
				sqlCmd_sum_rule.append(fromtable);				
				sqlCmd_sum_rule.append("      ) d,(select * from v_bank_location where m_year=?) e");				
				sqlCmd_sum_rule.append(" where e.bank_type=? and d.bank_code(+)=e.bank_no");
				
		  		//總計
				sqlCmd_sum.append(sqlCmd_sum_rule);
				sqlCmd_sum.append(" and (e.hsien_id>'Y' or e.hsien_id<'Y')");//--總計
				dbData_sum=DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb");
				System.out.println("總計的dbData_sum.size()="+dbData_sum.size());
				//台灣省小計
				sqlCmd_sumtaiwan.append(sqlCmd_sum_rule);				
				if(Integer.parseInt(s_year) <= 99){
				sqlCmd_sumtaiwan.append(" and e.hsien_id not in ('A','E','Z','W','Y') ");//--台灣省小計,縣市別不含台北/高雄/連江縣/金門縣/其他
				}else{
					sqlCmd_sumtaiwan.append(" and e.hsien_id not in ('f','A','H','b','d','e','Z','W','Y','h') ");//--台灣省小計,縣市別不含新北市/台北市/桃園市/台中市/台南市/高雄/連江縣/金門縣/其他(中華民國農會)	108.04.22 fix
				}
				dbData_sumtaiwan=DBManager.QueryDB_SQLParam(sqlCmd_sumtaiwan.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb");
				System.out.println("台灣省小計的dbData_sumtaiwan.size()="+dbData_sumtaiwan.size());

				//福建省小計
				sqlCmd_sumfuchien.append(sqlCmd_sum_rule);				
				sqlCmd_sumfuchien.append(" and e.hsien_id in ('Z','W')"); //--福建省小計,縣市別為連江縣/金門縣
				dbData_sumfuchien=DBManager.QueryDB_SQLParam(sqlCmd_sumfuchien.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb");
				System.out.println("福建省小計的dbData_sumfuchien.size()="+dbData_sumfuchien.size());
				
				//各縣				
				sqlCmd_each.append(" select distinct g.hsien_id,g.hsien_name,fr001w_output_order,");
				sqlCmd_each.append("        fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,");
				sqlCmd_each.append("        fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,");
				sqlCmd_each.append("        fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb"); 
				sqlCmd_each.append(" from ( ");
				sqlCmd_each.append("      select nvl(e.hsien_id,' ') as hsien_id,nvl(e.hsien_name,'其他') as hsien_name,");
				sqlCmd_each.append(field);
				sqlCmd_each.append("      from (");
                sqlCmd_each.append(fromtable);
                sqlCmd_each.append("           ) d,(select * from v_bank_location where m_year=?) e");                
                sqlCmd_each.append("      where e.bank_type=? and d.bank_code(+)=e.bank_no and (e.hsien_id>'Y' or e.hsien_id<'Y')");               
                sqlCmd_each.append(" group by nvl(e.hsien_id,' '),nvl(e.hsien_name,'其他'),e.fr001w_output_order");
                sqlCmd_each.append(" ) f,(select * from v_bank_location where m_year=?) g"); 
                sqlCmd_each.append(" where f.hsien_id(+)=g.hsien_id"); 
                sqlCmd_each.append(" order by g.fr001w_output_order");
                paramList.add(wlx01_m_year);
                dbData_each=DBManager.QueryDB_SQLParam(sqlCmd_each.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb");
				System.out.println("各縣的dbData_each.size()="+dbData_each.size());
			}else {
				//分支明細
				if(!bank_no.equals("")){//99.09.13 add 單一家信用部明細資料
					sqlCmd_detail.append(" select * ");
					sqlCmd_detail.append(" from ( ");
				}	
				sqlCmd_detail.append(" select e.fr001w_output_order,bank_code,d.bank_name,");
				sqlCmd_detail.append(field);
				sqlCmd_detail.append(" from (select * from ").append(cd01_table).append(" cd01)cd01");
				sqlCmd_detail.append(" left join (");
				sqlCmd_detail.append(fromtable);
				sqlCmd_detail.append(" ) d on d.bank_code=wlx01.bank_no ");
	            sqlCmd_detail.append(" ,(select * from v_bank_location where m_year=?) e ");
	          	sqlCmd_detail.append(" where e.bank_type=?");
	          	sqlCmd_detail.append(" and d.bank_code=e.bank_no");
	          	sqlCmd_detail.append(" group by e.fr001w_output_order,d.bank_code,d.bank_name ");      
	          	sqlCmd_detail.append(" order by e.fr001w_output_order,d.bank_code ");
	          	if(!bank_no.equals("")){//99.09.13 add 單一家信用部明細資料
		            sqlCmd_detail.append(")where bank_code=?");
		            paramList.add(bank_no);
		        }
	          	/*
	          	for(i=0;i<paramList.size();i++){
	            	System.out.println("i=["+i+"]="+(String)paramList.get(i));
	            }
	            */
				dbData_detail=DBManager.QueryDB_SQLParam(sqlCmd_detail.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp,fieldq,fieldr,fields,fieldt,fieldu,fieldv,fieldw,fieldx,fieldy,fieldz,fieldza,fieldzb");
				System.out.println("分支明細的dbData_detail.size()="+dbData_detail.size());
			}
	
			//列印單一家信用部明細表名稱
			if(!bank_no.equals("")){
				row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);
				cell=row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
				cell.setCellValue(bank_name+"按縣市別平均利率");
			}else{//列印農漁會
				row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);
				cell=row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
				cell.setCellValue("全體"+(bank_type.equals("6")?"農會":"漁會")+"按縣市別平均利率");
			}
						
			//列印年度
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			if (s_month.equals("0")) {
				cell.setCellValue("中華民國"+s_year+"年度");
			}else {
			   cell.setCellValue("中華民國"+s_year+"年"+s_month+"月");
			}
			//列印日期
			if (datestate.equals("1")) {
				Calendar rightNow=Calendar.getInstance();
				String year=String.valueOf(rightNow.get(Calendar.YEAR)-1911);
				String month=String.valueOf(rightNow.get(Calendar.MONTH)+1);
				String day=String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
				cell=row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("列印日期："+year+"年"+month+"月"+day+"日");
		    }
	
			//列印金額單位
			cell=row.getCell((short)25);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			unit_name=Utility.getUnitName(unit);
			cell.setCellValue("單位：新台幣"+unit_name+"、％");
			short top=0,down=0;
	
			//彙整
			if (rptStyle.equals("0")) {
				//先印總計
				row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
				if (dbData_sum !=null && dbData_sum.size() !=0) {
					for (i=1;i<=27;i++) {
					    insertCell(dbData_sum,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
					}
		       	}
				//台北市.高雄市
				rowNum=6;
				if (dbData_each !=null && dbData_each.size() !=0) {
				   //台北市(hsien_id=A);dbData_each.get(0)-->台北市
				   for(j=0;j<dbData_each.size();j++){
				       String hsienId = (((DataObject)dbData_each.get(j)).getValue("hsien_id")).toString();
				       //System.out.println("hsienId="+hsienId+":"+((DataObject)dbData_each.get(j)).getValue("hsien_name"));
				       if("99".equals(wlx01_m_year) && "h".equals(hsienId)){
				           row=(sheet.getRow(30)==null)? sheet.createRow(30) : sheet.getRow(30);
				       }else{
				           row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				       }
				       //System.out.println("rowNum="+rowNum);
                       if((j-1)==dbData_each.size()){//最後一筆時加上底線
				       	  down=HSSFCellStyle.BORDER_MEDIUM;
				       }else{
				       	  down=0;
				       }
				       for(i=1;i<=27;i++){
				       	   insertCell(dbData_each,true,j,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		               }
				       if(("99".equals(wlx01_m_year) && "E".equals(hsienId)) || 
				       	  ("100".equals(wlx01_m_year) && "e".equals(hsienId))){//-->高雄市
				       	  System.out.println("j="+j+":"+((DataObject)dbData_each.get(j)).getValue("hsien_name"));
				       	  rowNum++;
				      	  //印臺灣省小計
						  row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						  if(dbData_sumtaiwan !=null && dbData_sumtaiwan.size() !=0){
							for(i=1;i<=27;i++){
			  				    insertCell(dbData_sumtaiwan,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
							}
				       	  }
				       }//end of 高雄市
				       if(("99".equals(wlx01_m_year) && "D".equals(hsienId))/*-->99年度,D:台南市*/ 
				       || ("100".equals(wlx01_m_year) && j==20))/*嘉義市*/{
				       	 System.out.println("j="+j+":"+((DataObject)dbData_each.get(j)).getValue("hsien_name"));				       	  
				       	if("100".equals(wlx01_m_year) && "h".equals(hsienId)){/*h:其他*/
                            rowNum++;
                        }else if("99".equals(wlx01_m_year) && "D".equals(hsienId)){
                            rowNum+=2;
                        }
				       	//System.out.println("rowNum="+rowNum);
				      	  //印福建省小計
						  row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						  if(dbData_sumfuchien !=null && dbData_sumfuchien.size() !=0){
							for(i=1;i<=27;i++){
			  				    insertCell(dbData_sumfuchien,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
							}
				       	  }
				       }//end of 99年台南市.100年嘉義市
				       if(!("99".equals(wlx01_m_year) && "h".equals(hsienId))){
				           rowNum++;
				       }
				   }//end of for dbData_each
				    int tailNum = 0;
	                if("100".equals(wlx01_m_year)){
	                    tailNum = 32;
	                }else if("99".equals(wlx01_m_year)){
	                    tailNum = 35;
	                }
	               
	                /*106.08.07 取消顯示
	                if("6".equals(bank_type)){
    	                row=(sheet.getRow(tailNum)==null)? sheet.createRow(tailNum) : sheet.getRow(tailNum);
    	                insertCell_tail("填表說明 :",wb,row,(short)0,(short)0,(short)0);
                        insertCell_tail("其他係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。",wb,row,(short)1,(short)0,(short)0);
	                }
	                */
				}//end of 列印各個縣市
			}else {
					System.out.println("明細表開始");
					rowNum=5;
					if (dbData_detail !=null && dbData_detail.size() !=0) {
					    for (j=0;j<dbData_detail.size();j++) {
						    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						    if(j==dbData_detail.size()-1) {//最後一筆時加上底線
						       down=HSSFCellStyle.BORDER_MEDIUM;
						       //System.out.println("j.lastdata="+j);
						    }else {
							   down=0;
							   //System.out.println("j="+j);
						    }						    
						    insertCell(dbData_detail,true,j,"bank_name",wb,row,(short)0, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
						    for(i=1;i<=27;i++){
						    	if(i==27){
						    		insertCell_detail(dbData_detail,true,j,"field"+hsien_id_sum[i],wb,row,(short)i,(short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
						    	}else{
						    		insertCell_detail(dbData_detail,true,j,"field"+hsien_id_sum[i],wb,row,(short)i,(short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
						    	}
				            }
						    rowNum++;
					    }
				   	}
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);				
					insertCell_tail("填表",wb,row,(short)0,HSSFCellStyle.BORDER_MEDIUM,(short)0);				
					insertCell_tail("審核",wb,row,(short)2,HSSFCellStyle.BORDER_MEDIUM,(short)0);				
					insertCell_tail("主辦業務人員",wb,row,(short)7,HSSFCellStyle.BORDER_MEDIUM,(short)0);				
					insertCell_tail("機關長官",wb,row,(short)12,HSSFCellStyle.BORDER_MEDIUM,(short)0);
				
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
								
					insertCell_tail("主辦統計人員",wb,row,(short)7,(short)0,(short)0);
				
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				
					insertBlank(15,wb,(short)0,(short)0,row);
				
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
								
					insertCell_tail("資料來源：根據各"+bank_type_name+"信用部資料彙編。",wb,row,(short)0,(short)0,(short)0);
				
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					//100.10.13 fix 更改組織名稱	
					insertCell_tail("填表說明：本表編製一式三份，一份送行政院"+((Integer.parseInt(s_year)>= 101)?"農業部":"農業委員會")+"統計室，一份送本局會計室，一份自存。",wb,row,(short)0,(short)0,(short)0);
					
					System.out.println("明細表end");	
					//95.10.13 增加檢核結果與最後異動日期 
					if(!bank_no.equals("")){
					   sqlCmd_detail = new StringBuffer();  //各機構明細
					   paramList = new ArrayList();//共同參數
					   sqlCmd_detail.append(" select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE");
					   sqlCmd_detail.append(" from WML01");		   		  
					   sqlCmd_detail.append(" where M_YEAR=?");		   		   
					   sqlCmd_detail.append(" and M_MONTH=?");
					   sqlCmd_detail.append(" and BANK_CODE=?");
					   sqlCmd_detail.append(" and REPORT_NO=?");
					   paramList.add(s_year);
					   paramList.add(s_month);
					   paramList.add(bank_no);
					   paramList.add("A03"); 
					   List dbData=DBManager.QueryDB_SQLParam(sqlCmd_detail.toString(),paramList,"");		       	
					   	  
	    		       String UPD_CODE="";
	    		       String UPDATE_DATE="";
	    		       String M_YEAR="";   
	    		       String M_MONTH="";    
	    		       String M_DATE=""; 
	    		       if(dbData.size()>0){
	       		          System.out.println("dbData.size()="+dbData.size()); 
	       		          UPD_CODE = (String)((DataObject)dbData.get(0)).getValue("upd_code");  
	       		          UPDATE_DATE = (String)((DataObject)dbData.get(0)).getValue("update_date");       		   
				          System.out.println("UPD_CODE="+UPD_CODE); 
				          System.out.println("UPDATE_DATE="+UPDATE_DATE); 
				          if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
				          else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
				          else if(UPD_CODE.equals("U")) UPD_CODE="檢核成功";
				          else UPD_CODE="";
				          System.out.println("UPD_CODE="+UPD_CODE);
				          M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
				          M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
				          M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
				          UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
				          System.out.println("UPDATE_DATE="+UPDATE_DATE); 		
				          rowNum=rowNum+3;	      	       
				          row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				          System.out.println("rowNum="+rowNum);		     	
				          insertCell_tail("檢核結果:"+UPD_CODE,wb,row,(short)0,(short)0,(short)0);				  
				          rowNum++;		
				          row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				          System.out.println("rowNum="+rowNum);		
				          insertCell_tail("最後異動日期:"+UPDATE_DATE,wb,row,(short)0,(short)0,(short)0);
				          System.out.println("明細表end");		   
	       		       }else{
	       		       	rowNum=rowNum+3;	      	       
				           row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				       	System.out.println("rowNum="+rowNum);		     	
				       	insertCell_tail("檢核結果:待檢核",wb,row,(short)0,(short)0,(short)0);				  
				       	rowNum++;		
				       	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				       	System.out.println("rowNum="+rowNum);		
				           insertCell_tail("最後異動日期:無",wb,row,(short)0,(short)0,(short)0);
				       	System.out.println("列印檢核結果--明細表end");		      				
	       		       }	
					}
				}
			    
				HSSFFooter footer=sheet.getFooter();
				footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
				footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	
				FileOutputStream fout=new FileOutputStream(reportDir+System.getProperty("file.separator")+filename);
				wb.write(fout);
				//儲存
				fout.close();
				System.out.println("儲存完成");
			}catch(Exception e) {
				System.out.println("createRpt Error:"+e+e.getMessage());
			}
			return errMsg;
	}



	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
	                              short bg,short bordertop,short borderbottom,short borderleft,short borderright) {
		try {
		String insertValue="";
  		if(getstate) insertValue=(((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue=Item;
		//System.out.println("insertValue="+insertValue);
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);


	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	    //double value=0;
	    if(!insertValue.equals("0") && !insertValue.equals("") && insertValue != null){
	    	if(insertValue.indexOf("信用部") != -1){
	    		cell.setCellValue(insertValue);	
	    	}else{
	    	   cell.setCellValue(Utility.setCommaFormat(insertValue));
	    	} 
	    }
	   
		}
			catch(Exception e) {
				System.out.println("insertCell Error:"+e+e.getMessage());
			}
	}



	public static void insertCell_detail(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
	                                     short bg,short bordertop,short borderbottom,short borderleft,short borderright) {
		try {
			String insertValue="";
			if(getstate) insertValue=(((DataObject)dbData.get(index)).getValue(Item)).toString();
			else         insertValue=Item;
			/*
			if(insertValue==null){
				System.out.println(Item+"==null");
			}else{
				System.out.println(j+"insertValue="+insertValue);
			}
			*/
			HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
			HSSFCellStyle cs1=wb.createCellStyle();

			cs1.setBorderTop(bordertop);
			cs1.setBorderBottom(borderbottom);
			cs1.setBorderLeft(borderleft);
			cs1.setBorderRight(borderright);

			if(j==0){
					HSSFFont f=wb.createFont();
					f.setFontHeightInPoints((short) 10);
					cs1.setAlignment(HSSFCellStyle.ALIGN_GENERAL);
					f.setFontName("標楷體");
					cs1.setFont(f);
			}else{
					cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			}

			cell.setCellStyle(cs1);
			cell.setEncoding( HSSFCell.ENCODING_UTF_16 );

			if(!insertValue.equals("0") && !insertValue.equals("") && insertValue !=null){
				if(insertValue.indexOf("信用部") !=-1){
					cell.setCellValue(insertValue);
				}else{
					cell.setCellValue(Utility.setCommaFormat(insertValue));
				}
				//System.out.println("insert test");
			}
		}
			catch(Exception e) {
				System.out.println("insertCell Error:"+e+e.getMessage());
			}
	}



	public static void insertCell_tail(String Item,HSSFWorkbook wb,HSSFRow row,short j,short top,short down) {
		try {			
			//System.out.println(j+"Item="+Item);
			HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
			HSSFCellStyle cs1 = wb.createCellStyle();
			//HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
			cs1.setBorderTop(top);
		    cs1.setBorderBottom(down);
		    cs1.setBorderLeft((short)0);
		    cs1.setBorderRight((short)0);
		  	
			HSSFFont f = wb.createFont();
		    f.setFontHeightInPoints((short) 12);
		    f.setFontName("標楷體");		    
		    cs1.setFont(f);
		    cs1.setAlignment(HSSFCellStyle.ALIGN_GENERAL);
		    cell.setCellStyle(cs1);
			cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
			cell.setCellValue(Item);	

		}catch(Exception e){
			System.out.println("insertCell_tail Error:"+e+e.getMessage());
		}
	}
	public static void insertBlank(int maxlength,HSSFWorkbook wb,short top,short down,HSSFRow row){
	    for(int k=0;k<=maxlength;k++){
   	        insertCell(null,false,0," ",wb,row,(short)k, (short)64,top,(short)0,(short)0,(short)0);	        
        }//end of insert 空值表格
	}
}