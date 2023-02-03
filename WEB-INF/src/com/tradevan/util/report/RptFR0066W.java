/*
 94.03.11 Fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
 94.03.15 ADD 月份輸入條件查詢 	by EGG
 94.08.04 fix a01 改成  ((select * from a01) union (select * from a02)) a01
 94.08.23 add 明細表 by 2295
 94.08.30 fix 放款總額（fieldB）小括號位置	jwang
 94.08.30 add 變數「fromtable2」	jwang 
 95.09.27 fix (一)月底存放比率-->比率用『月底存放比率(A01)_合庫版』
 							    再比對『鄉鎮地區農會不得超過80%，直轄市及省、縣轄市農會不得超過78%』
          fix 非會員存款總額(990310)-1/2(990630)
          fix 倍->% , 10倍=1000% , 1.5倍=150%
          fix 非會員授信總額(990610) -（990511）- (990611)
         	  扣除1.非會員無擔保消費政策性貸款(990511)(94.10月起).
         	     2.對直轄市、縣（市）政府辦理之授信總額(990611) (95.1月起).
          fix 94.10以後.10~13公式.漁會同農會            
          fix 1301.1304.1305.1306金額,以千元為單位                           
          fix (七).(八).(十).(十一).(十二).(十三)引用規定應為"農會漁會信用部各項風險控制比率管理辦法"
                而非"農會漁會信用部對贊助會員及非會員授信及其限額標準"
              (四).(3)加入逾放比低於2%且資本適足率高8%者,得不超過150%               by 2295
 95.10.03 增加檢核結果與最後異動日期 by 2495
 97.01.30 add 990612 非會員政策性農業專案貸款 
          fix 非會員授信總額(990610) -（990511）- (990611) - (990612)[97.2月起] by 2295
 99.01.21 fix 計算上年度信用部決算淨值/全体農(漁)會決算淨值時,需扣除992810投資全國農業金庫尚未攤提之損失[99.1月起] by 2295
 99.03.10 add 加上查詢年度.取得新總機構名稱 by 2295    
 			  使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295  
100.02.16 fix wlx01.bn01區分99/100年度 by 2295
100.10.14 fix (六)計算比率時.非會員存款總額(990620) -1/2(990630) by 2295 	
101.02.06 fix wlx01_m_year by 2295
101.08.23 add 違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目： 
              990812一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
              990813二、因固定資產重估增值
              990814三、因淨值降低 by 2295	 
102.01.17 add 990421已申請符合逾放比率低於百分之一、其資本適足率高於百分之十且備抵保帳覆蓋率高於百分之一百經主管機關同意,農金局回文函號
              990621申請符合逾放比率低於百分之二且資本適足率高於百分之八經主管機關同意,農金局回文函號 by 2295   
102.02.05 fix a01查詢條件 by 2295   
103.01.07 add (二)2.1/2.2及(八)上年度信用部決算淨值為負數時,要顯示違反  by 2295
          add (九)逾新台幣100萬且超過前一年度信用部決算淨值5%,才算違反 by 2295         
104.03.18 add 取消 (六)計算比率時.非會員存款總額(990620)不扣除1/2(990630)公庫存款 by 2295  
104.04.22 add (一)月底存放比不得超過80%之規定,取消(三)限制
          add (四)信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者(990422),不受限制) by 2295          
          add (六)逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者(990622),得不超過 200%) by 2295       
          add (七)調整為50% by 2295
          add (十四)對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額 by 2295   
106.05.22 add 若逾期放款為0,備抵呆帳覆蓋率(field_backup_over_rate=(120800+150300)/99000,因分母為0,則該比率顯示為N/A,調整(四)符合為不受限制及200%/(六)符合200% by 2295
106.10.06 add (五)無擔保消費性貸款增加扣除990511非會員無擔保消費性政策貸款   by 2295
          add (六)非會員授信總額.移除扣除990511非會員無擔保消費性政策貸款=非會員授信總額(990610) - (990611) - (990612)] by 2295
107.04.09 fix (四)不受限制,取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%條件檢核 by 2295
              (六)200%,取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%條件檢核  by 2295
107.05.02 fix (四)/(六)依所勾選的套用適用限額 by 2295 
108.01.08 add (六)辦理自用住宅放款限額,增加定期性公庫存款總額990721/公庫存款以外定期性存款總額990722 by 2295
108.03.27 add 增加報表格式 by 2295           
108.09.10 add 108年10月以後,調整為 6.違反購置住宅放款及房屋修繕放款之餘額不得超過存款總餘額55%之規定 by 2295
110.03.30 add 非會員擔保品種類; 非會員擔保品坐落地; by 6493
		  fix 贊助會員授信總額占贊助會員存款總額之比率; 非會員授信總額占非會員存款總額之比率 by 6493
110.05.13 fix 110/5月套用新格式報表 by 2295
111.05.30 add 16~18項法定比率 by 2295
111.09.27 add 11~14，有關「未達九百萬而在六百萬元以上者，得以九百萬元為最高限額；未達六百萬元者，得以六百萬元為最高限額」，修正為「未達九百萬元者，得以九百萬元為最高限額」。
              14項，有關「依信用部上一年度決算淨值百分之十二.五核算為」，修正為「依信用部上一年度決算淨值百分之十二.五及百分之十五核算分別為」；
                                          「依信用部上一年度決算淨值百分之二.五核算為」，修正為「依信用部上一年度決算淨值百分之二.五及百分之三核算分別為」。		  
 */
package com.tradevan.util.report;

import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.RTFMixer;
import java.text.DecimalFormat;

public class RptFR0066W {
  public static String createRpt(String s_year, String s_month, String bank_code,
                                 String datestate, String bank_type,
                                 String rptStyle,boolean isAP,String printStyle) {
    System.out.println("RptFR0066W.Rtf Start ...");
    String rtf_9901 = "";
    String errMsg = "";
    List dbData = null;
    String sqlCmd = "";
    Properties A02Data = new Properties();
    String acc_code = "";
    String amt = "";
    StringBuffer sb_SQL = new StringBuffer();
    //String div = null;
    String l_year = null;
    //String l_month = null;
    String tmpSQL = null;
    String rtfSrcName = null;
    String rtfTagetName = null;
    String bank_name = null;
    String areaNo = null;
    DecimalFormat dft = new DecimalFormat("#.##");
    DecimalFormat dft_int = new DecimalFormat("#,###");
    double tmp_A = 0;
    double tmp_B = 0;
    List paramList = new ArrayList();
    DataObject bean = null;
    String cd01_table = "";
    String wlx01_m_year = "";
    /*long A_01 = 0;
    long B_01 = 0;
    long C_01 = 0;
    long D_01 = 0;
    long E_01 = 0;*/
     l_year = String.valueOf(Integer.parseInt(s_year)-1);
     
    if(s_month.length()==1){
      s_month = "0"+s_month;
    }
    //100.02.16 add 查詢年度100年以前.縣市別不同===============================
  	cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
  	wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
  	//=====================================================================    
  	printStyle = printStyle.equals("")?"rtf":printStyle;//108.09.17 add
    /*
    if(s_month.equals("01")){
      l_month = "12";
    }else{
      l_month =String.valueOf(Integer.parseInt(s_month)-1);
      if(l_month.length()==1){
        l_month = "0"+l_month;
      }
    }
    if((Integer.parseInt(s_year)>=94)&&(Integer.parseInt(s_month)>6)){
      div="2";
    }else{
      l_month = "";
      l_year = "";
      div="1";
    }
    */
  	
    //99.01.21 add 99.01月份.套用新格式
    if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 9901 && (Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month) < 10401)){
    	rtf_9901 = "_9901";    	
    }	
    if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10401 && (Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month) < 10810)){
        rtf_9901 = "_10401";        
    }
    if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10810 && (Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month) < 11005)){
    	rtf_9901 = "_10810";        
    }
    if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005){
    	rtf_9901 = "_11003";        
    }
    //rtf_9901 = "_10401";//104.04.20測試用
   
    //法定比率第一區塊START
    //法定比率第一區塊SQL
    /*99.01.21
     sb_SQL = sb_SQL.append(" select  bank_code, ");
     sb_SQL = sb_SQL.append(" round(sum(decode(a01.acc_code,'120000',amt,0))/"+div+",0)  \"120000\", ");
     sb_SQL = sb_SQL.append(" round(sum(decode(a01.acc_code,'310000',amt,0)),0)+round(sum(decode(a01.acc_code,'320000',amt,0))/"+div+",0)   \"310000\", ");
     sb_SQL = sb_SQL.append(" round(sum(decode(a01.acc_code,'140000',amt,0))/"+div+",0)  \"140000\", ");
     sb_SQL = sb_SQL.append(" round(sum(decode(a01.acc_code,'220000',amt,0))/"+div+",0)  \"220000\", ");
     sb_SQL = sb_SQL.append(" round(sum(decode(a01.acc_code,'220900',amt,0))/"+div+",0)  \"220900\" ");
     sb_SQL = sb_SQL.append(" from a01,  ba01, wlx01  ");
     if(l_month.equals("01")){
       sb_SQL = sb_SQL.append(" where ((a01.m_year = '" + s_year +
                              "'   and   a01.m_month= '" + s_month +
                              "') or (a01.m_year = '" + l_year +
                              "' and a01.m_month= '" + l_month + "'))  ");
     }else{
       sb_SQL = sb_SQL.append(" where ((a01.m_year = '" + s_year +
                              "'   and   a01.m_month= '" + s_month +
                              "') or (a01.m_year = '" + s_year +
                              "' and a01.m_month= '" + l_month + "'))  ");

     }
     sb_SQL = sb_SQL.append(" and   wlx01.bank_no=ba01.bank_no   ");
     sb_SQL = sb_SQL.append(" and   ba01.bank_type='"+bank_type+"' ");
     sb_SQL = sb_SQL.append(" and   a01.bank_code=ba01.bank_no ");
     sb_SQL = sb_SQL.append(" and   wlx01.bank_no='"+bank_code+"' ");
     sb_SQL = sb_SQL.append(" group by bank_code ");
     tmpSQL = sb_SQL.toString();
     System.out.println("tmpSQL = "+tmpSQL);
     sb_SQL.delete(0,sb_SQL.length());
     dbData = DBManager.QueryDB(tmpSQL, "m_year,m_month,amt,120000,310000,140000,220000,220900,320000");
     System.out.println("dbData.size()= "+dbData.size());
     */
    //String bank_type="";
    //報表路徑
    try{
      if(isAP){
        if (bank_type.equals("6")) {
            rtfSrcName = Utility.getProperties("rtfDir") +System.getProperty("file.separator") +"xx年xx月份xx農會各項法定比率表"+rtf_9901+".rtf";
            rtfTagetName = Utility.getProperties("allReportDir") +System.getProperty("file.separator") + "FR0066W"+bank_code+s_year+s_month+"."+printStyle;
        }else {
          rtfSrcName = Utility.getProperties("rtfDir") +System.getProperty("file.separator") +"xx年xx月份xx漁會各項法定比率表"+rtf_9901+".rtf";
          rtfTagetName = Utility.getProperties("allReportDir") +System.getProperty("file.separator") + "FR0066W"+bank_code+s_year+s_month+"."+printStyle;
        }
      }else{
        if (bank_type.equals("6")) {
          rtfSrcName = Utility.getProperties("rtfDir") +System.getProperty("file.separator") +"xx年xx月份xx農會各項法定比率表"+rtf_9901+".rtf";
          rtfTagetName = Utility.getProperties("reportDir") +System.getProperty("file.separator") + "農會各項法定比率表."+printStyle;
        }else {
          rtfSrcName = Utility.getProperties("rtfDir") +System.getProperty("file.separator") +"xx年xx月份xx漁會各項法定比率表"+rtf_9901+".rtf";
          rtfTagetName = Utility.getProperties("reportDir") + System.getProperty("file.separator") + "漁會各項法定比率表."+printStyle;
        }
      }
    }catch(Exception e){

    }

    System.out.println("rtfSrcName="+rtfSrcName);
    System.out.println("rtfTagetName="+rtfTagetName);
    RTFMixer mix = new RTFMixer(rtfSrcName, rtfTagetName);

    //處理要填上的變數

    //若有中文問題，請轉碼
      Properties props = new Properties();
      //95.09.27 fix (一)月底存放比率-->比率用『月底存放比率(A01)_合庫版』
      /*
      for (int i =0; i < dbData.size(); i++) {
        A_01 = Long.parseLong( ( ( (DataObject) dbData.get(i)).
                                             getValue("120000")).toString());

        B_01 = ( (DataObject) dbData.get(i)).getValue("310000").toString() == null ?
            0 :
            Long.parseLong( ( (DataObject) dbData.get(i)).getValue("310000").
                           toString());
        C_01 = ( (DataObject) dbData.get(i)).getValue("140000") == null ?
            0 :
            Long.parseLong( ( (DataObject) dbData.get(i)).getValue("140000").
                           toString());

        D_01 = ( (DataObject) dbData.get(i)).getValue("220000") == null ?
            0 :
            Long.parseLong( ( (DataObject) dbData.get(i)).getValue("220000").
                           toString());
        E_01 = ( (DataObject) dbData.get(i)).getValue("220900") == null ?
            0 :
            Long.parseLong( ( (DataObject) dbData.get(i)).getValue("220900").
                           toString());

      }
      props.setProperty("&101",dft_int.format(Math.round(A_01/1000)));
      props.setProperty("&102",dft_int.format(Math.round(B_01/1000)));
      props.setProperty("&103",dft_int.format(Math.round(C_01/1000)));
      props.setProperty("&104",dft_int.format(Math.round(D_01/1000)));
      props.setProperty("&105",dft_int.format(Math.round(E_01/1000)));

      if((D_01-E_01/2)!=0){
        if (C_01 > B_01) {
          tmp_A = (double)((double)A_01 / ((double)D_01 - (double)E_01 / 2));
          System.out.println("tmp_A = "+tmp_A);
        }
        else {
          tmp_A = (double)(((double)A_01 - (double)B_01 + (double)C_01) / ((double)D_01 - (double)E_01 / 2));
          System.out.println("tmp_A 1= "+tmp_A);
        }
      }
      props.setProperty("&106",String.valueOf(Math.round(tmp_A*100)));
      */
      //法定比率第一區塊END
      //dbData.clear();
      //其他法定比率START
      /*102.01.17
      sb_SQL = sb_SQL.append( " select a.bank_code,c.bank_name,b.hsien_div_1 ,");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990110',amt,0)) \"990110\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990120',amt,0)) \"990120\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990130',amt,0)) \"990130\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990140',amt,0)) \"990140\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990150',amt,0)) \"990150\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990210',amt,0)) \"990210\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990220',amt,0)) \"990220\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990230',amt,0)) \"990230\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990240',amt,0)) \"990240\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990310',amt,0)) \"990310\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990320',amt,0)) \"990320\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990410',amt,0)) \"990410\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990420',amt,0)) \"990420\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990421',amt,0)) \"990421\", ");//102.01.15 add
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990510',amt,0)) \"990510\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990511',amt,0)) \"990511\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990610',amt,0)) \"990610\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990611',amt,0)) \"990611\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990612',amt,0)) \"990612\", ");//97.01.30
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990620',amt,0)) \"990620\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990621',amt,0)) \"990621\", ");//102.01.15 add
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990630',amt,0)) \"990630\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990710',amt,0)) \"990710\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990720',amt,0)) \"990720\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990810',amt,0)) \"990810\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990812',amt,0)) \"990812\", ");//101.08.23 add
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990813',amt,0)) \"990813\", ");//101.08.23 add
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990814',amt,0)) \"990814\", ");//101.08.23 add
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990910',amt,0)) \"990910\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990920',amt,0)) \"990920\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'990930',amt,0)) \"990930\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991010',amt,0)) \"991010\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991020',amt,0)) \"991020\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991030',amt,0)) \"991030\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991110',amt,0)) \"991110\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991120',amt,0)) \"991120\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991210',amt,0)) \"991210\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991220',amt,0)) \"991220\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991310',amt,0)) \"991310\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'991320',amt,0)) \"991320\", ");
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,'992810',amt,0)) \"992810\", ");//99.01.21 992810投資全國農業金庫尚未攤提之損失[99.1月起]
      sb_SQL = sb_SQL.append( " sum(decode(a.acc_code,null,0,'99141Y',amt,0)) \"99141Y\"  ");
      sb_SQL = sb_SQL.append( " from (select * from A02 ");
      sb_SQL = sb_SQL.append( " where m_year= ? ");
      sb_SQL = sb_SQL.append( " and m_month= ? ");
      sb_SQL = sb_SQL.append( " and bank_code= ?");      
      sb_SQL = sb_SQL.append( " union ");
      sb_SQL = sb_SQL.append( " select m_year,m_month,bank_code,acc_code,amt,'' from A99 ");
      sb_SQL = sb_SQL.append( " where m_year= ? ");
      sb_SQL = sb_SQL.append( " and m_month=? ");
      sb_SQL = sb_SQL.append( " and bank_code=?");
      sb_SQL = sb_SQL.append( " )a,(select * from wlx01 where m_year=?)b,(select * from bn01 where m_year=?)c ");
      sb_SQL = sb_SQL.append( " where a.bank_code = b.bank_no ");
      sb_SQL = sb_SQL.append( " and   a.bank_code = c.bank_no ");
      sb_SQL = sb_SQL.append( " and   c.bank_type = ?");      
      sb_SQL = sb_SQL.append( " and   b.m_year = ? "); 
      sb_SQL = sb_SQL.append( " group by a.bank_code,c.bank_name,b.hsien_div_1 order by bank_code ");
      tmpSQL = sb_SQL.toString();
      //System.out.println("allen tmpSQL = "+tmpSQL);
      sb_SQL.delete(0, sb_SQL.length());
      
      //設定SQL參數
      paramList.add(s_year);
      paramList.add(s_month);
      paramList.add(bank_code);
      paramList.add(s_year);
      paramList.add(s_month);
      paramList.add(bank_code);
      paramList.add(wlx01_m_year);
      paramList.add(wlx01_m_year);
      paramList.add(bank_type);
      paramList.add(wlx01_m_year);//101.02.06 fix wlx01_m_year
      
      
      dbData = DBManager.QueryDB_SQLParam(tmpSQL,paramList, "m_year,m_month,amt,990110,990120,990130,990140,990150,990210,990220,990230,990240,990310,990320,990410,990420,990421,990510,990511,990610,990611,990612,990620,990621,990630,990710,990720,990810,990812,990813,990814,990910,990920,990930,991010,991020,991030,991110,991120,991210,991220,991310,991320,992810,99141Y");
      System.out.println("dbData.size=" + dbData.size());
      */
      
      sb_SQL.append(" select a02.HSIEN_ID, nvl(a02.hsien_div_1,' ') as hsien_div_1,a02.bank_type,a02.bank_code as bank_no,bank_name,");
      sb_SQL.append("      amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,");
      sb_SQL.append("      amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990510,amt990512,amt990610,amt990511,amt990611,amt990612,");
      sb_SQL.append("      amt990620,amt990621,amt990622,amt990623,amt990710,amt990720,amt990721,amt990722,amt990810,amt990860,amt990870,amt990880,");
      sb_SQL.append("      amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt990930,amt991010,amt991020,amt991030,amt991110,");
      sb_SQL.append("      amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,amt990711,amt990712,field_OVER,field_BACKUP,field_DEBIT,");
      sb_SQL.append("      a05.amt as a05bis,");
      sb_SQL.append("      decode(field_CREDIT,0,0,round(field_OVER /  field_CREDIT *100 ,2))  as a01field_over_rate,"); 
      sb_SQL.append("      decode(FIELD_OVER,0,0,round(field_BACKUP /  FIELD_OVER *100 ,2))  as a01field_backup_over_rate,");
      sb_SQL.append("      a01_operation.field_backup_credit_rate as a01field_backup_credit_rate, ");//104.04.16 add 備呆占放款比率
      sb_SQL.append("      a01_operation_sum.field_backup_over_rate as a01field_backup_over_rate_avg, ");//104.04.16 add 全體信用部備呆占狹義逾期放款比率平均值
      sb_SQL.append("	   a01_operation.fieldi_y as a01fieldi_y"); //108.09.10 add 存放比率-存款總餘額	 
      sb_SQL.append(" from  ");
      sb_SQL.append("      (select wlx01.HSIEN_ID,wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name, ");
      sb_SQL.append("              sum(decode(a02.acc_code,'990110',amt,0)) as  amt990110,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990120',amt,0)) as  amt990120,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990130',amt,0)) as  amt990130,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990140',amt,0)) as  amt990140,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990150',amt,0)) as  amt990150,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990210',amt,0)) as  amt990210,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990230',amt,0)) as  amt990230,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990240',amt,0)) as  amt990240,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990220',amt,0)) as  amt990220,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990310',amt,0)) as  amt990310,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990630',amt,0)) as  amt990630,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990320',amt,0)) as  amt990320,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990410',amt,0)) as  amt990410,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990420',amt,0)) as  amt990420,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990421',amt,0)) as  amt990421,");//102.01.15 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990422',amt,0)) as  amt990422,");//104.04.16 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990510',amt,0)) as  amt990510,");//102.01.17 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990512',amt,0)) as  amt990512,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990610',amt,0)) as  amt990610,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990511',amt,0)) as  amt990511,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990611',amt,0)) as  amt990611,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990612',amt,0)) as  amt990612,");//97.01.29 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990620',amt,0)) as  amt990620,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990621',amt,0)) as  amt990621,");//102.01.15 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990622',amt,0)) as  amt990622,");//104.04.16 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990623',amt,0)) as  amt990623,");//110.03.30 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990710',amt,0)) as  amt990710,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990720',amt,0)) as  amt990720,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990721',amt,0)) as  amt990721,");//108.01.08 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990722',amt,0)) as  amt990722,");//108.01.08 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990810',amt,0)) as  amt990810,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990811',amt,0)) as  amt990811,");//101.08.23 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990812',amt,0)) as  amt990812,");//101.08.23 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990813',amt,0)) as  amt990813,");//101.08.23 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990814',amt,0)) as  amt990814,");//101.08.23 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990860',amt,0)) as  amt990860,");//111.05.30 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990870',amt,0)) as  amt990870,");//111.05.30 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990880',amt,0)) as  amt990880,");//111.05.30 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990910',amt,0)) as  amt990910,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990920',amt,0)) as  amt990920,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990930',amt,0)) as  amt990930,");//102.01.17 add
      sb_SQL.append("              sum(decode(a02.acc_code,'991010',amt,0)) as  amt991010,");//102.01.16 add
      sb_SQL.append("              sum(decode(a02.acc_code,'991020',amt,0)) as  amt991020,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991030',amt,0)) as  amt991030,");//102.01.16 add      
      sb_SQL.append("              sum(decode(a02.acc_code,'991110',amt,0)) as  amt991110,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991120',amt,0)) as  amt991120,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991210',amt,0)) as  amt991210,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991220',amt,0)) as  amt991220,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991310',amt,0)) as  amt991310,");
      sb_SQL.append("              sum(decode(a02.acc_code,'991320',amt,0)) as  amt991320,");
      sb_SQL.append("              sum(decode(a02.acc_code,'992810',amt,0)) as  amt992810,");
      sb_SQL.append("              sum(decode(a02.acc_code,'990000',amt,0))  as field_OVER,");
      sb_SQL.append("              sum(decode(a02.acc_code,'120800',amt,'150300',amt,0)) as field_BACKUP,");//102.02.05 add
      sb_SQL.append("              sum(decode(a02.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as field_CREDIT,");//102.02.05 add
      sb_SQL.append("              sum(decode(a02.acc_code,'220000',amt,0)) as field_DEBIT,");//111.05.30 add 
      sb_SQL.append("              sum(decode(a02.acc_code,'996114',amt,0)) as  amt996114,");//104.04.16 add
      sb_SQL.append("              sum(decode(a02.acc_code,'996115',amt,0)) as  amt996115,");//104.04.16 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990711',amt,0)) as  amt990711,");//108.09.10 add
      sb_SQL.append("              sum(decode(a02.acc_code,'990712',amt,0)) as  amt990712 ");//108.09.10 add
      sb_SQL.append("      from (select * from bn01 where m_year=? and bank_type = ?)bn01 left join (select * from a02 where a02.m_year=? and a02.m_month=?");
      paramList.add(wlx01_m_year);
      paramList.add(bank_type);
      paramList.add(s_year);
      paramList.add(s_month);
      sb_SQL.append("                             union");
      sb_SQL.append("                             select  m_year,m_month,bank_code,acc_code,amt,'','','' from a99 where a99.m_year=? and a99.m_month=?");
      paramList.add(s_year);
      paramList.add(s_month);
      //102.02.05 add a01
      sb_SQL.append("                             union");
      sb_SQL.append("                             select m_year,m_month,bank_code,acc_code,amt,'','','' from a01 where a01.m_year=? and a01.m_month=? and acc_code in('990000','120800','150300','120000','220000')");
      paramList.add(s_year);
      paramList.add(s_month);
      sb_SQL.append("                            )a02 on a02.bank_code = bn01.bank_no");      
      sb_SQL.append("      left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.BANK_NO");
      paramList.add(wlx01_m_year);    
      sb_SQL.append("      where bn01.bank_type in (?) and bn01.bank_no in(?) ");
      paramList.add(bank_type);
      paramList.add(bank_code);
      sb_SQL.append("      and wlx01.HSIEN_ID is not null ");
      sb_SQL.append("      group by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name");
      sb_SQL.append("      order by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name)a02");
      sb_SQL.append("   ,(select m_year,m_month,bank_code,acc_code,'4' as type,round(amt/1000,3) as amt,'' as violate from a05 ");
      sb_SQL.append("    where  a05.m_year =? and a05.m_month = ?");
      paramList.add(s_year);
      paramList.add(s_month);
      sb_SQL.append("    and a05.acc_code in ('91060P'))a05");  
      //104.04.16 add===================================================================================
      sb_SQL.append("    ,(select m_year,m_month,bank_code,");       
      sb_SQL.append("            sum(decode(acc_code,'field_over_rate',amt,0)) as  field_over_rate,"); 
      sb_SQL.append("            sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate,");       
      sb_SQL.append("            sum(decode(acc_code,'field_backup_credit_rate',amt,0)) as  field_backup_credit_rate,");//備呆占放款比率 104.04.16 add 
      sb_SQL.append("            sum(decode(acc_code,'fieldi_y',amt,0)) as  fieldi_y "); //存放比率-存款總餘額 108.09.10 add
      sb_SQL.append("      from a01_operation");  
      sb_SQL.append("      where m_year=? and m_month=?");
      paramList.add(s_year);
      paramList.add(s_month);
      sb_SQL.append("      and acc_code in('field_over_rate','field_backup_over_rate','field_backup_credit_rate','fieldi_y')");  
      sb_SQL.append("      group by m_year,m_month,bank_code  ) a01_operation");  
      sb_SQL.append("    ,(select m_year,m_month,");       
      sb_SQL.append("             sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate ");//104.04.16 add備呆占狹義逾期放款比率         
      sb_SQL.append("      from a01_operation");  
      sb_SQL.append("      where m_year=? and m_month=?");
      paramList.add(s_year);
      paramList.add(s_month);
      sb_SQL.append("      and acc_code in('field_backup_over_rate')");  
      sb_SQL.append("      and bank_type='ALL'");  
      sb_SQL.append("      and bank_code='ALL'");  
      sb_SQL.append("      and hsien_id=' '"); 
      sb_SQL.append("      group by m_year,m_month ) a01_operation_sum");
      //=======================================================================================
      sb_SQL.append("  where a02.bank_code = a05.bank_code ");
      sb_SQL.append("  and a02.bank_code   = a01_operation.bank_code");

      dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"m_year,m_month,amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,"
               + "amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990510,amt990512,amt990610,amt990511,amt990611,amt990612,"
               + "amt990620,amt990621,amt990622,amt990623,amt990710,amt990720,amt990721,amt990722,amt990810,amt990811,amt990812,amt990813,amt990814,amt990860,amt990870,amt990880,"
               + "amt990910,amt990920,amt990930,amt991010,amt991020,amt991030,amt991110,amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,amt990711,amt990712,"
               + "a05bis,field_over,field_backup,a01field_over_rate,a01field_backup_over_rate,a01field_backup_credit_rate,a01field_backup_over_rate_avg,a01fieldi_y,field_debit");
      System.out.println("dbData.size()="+dbData.size());
      long amt_990110 = 0;
      long amt_990120 = 0;
      long amt_990130 = 0;
      long amt_990140 = 0;
      long amt_990150 = 0;
      long amt_990210 = 0;
      long amt_990220 = 0;
      long amt_990230 = 0;
      long amt_990240 = 0;
      long amt_990310 = 0;
      long amt_990320 = 0;
      long amt_990410 = 0;
      long amt_990420 = 0;
      long amt_990510 = 0;
      long amt_990511 = 0;                        
      long amt_990610 = 0;      
      long amt_990611 = 0;
      long amt_990612 = 0;                  
      long amt_990620 = 0;
      long amt_990630 = 0;
      long amt_990710 = 0;
      long amt_990720 = 0;
      long amt_990721 = 0;//108.01.08 add
      long amt_990722 = 0;//108.01.08 add
      long amt_990810 = 0;
      long amt_990860 = 0;//111.05.30 add
      long amt_990870 = 0;//111.05.30 add
      long amt_990880 = 0;//111.05.30 add
      long amt_990910 = 0;
      long amt_990920 = 0;
      long amt_990930 = 0;
      long amt_991010 = 0;
      long amt_991020 = 0;
      long amt_991030 = 0;
      long amt_991110 = 0;
      long amt_991120 = 0;
      long amt_991210 = 0;
      long amt_991220 = 0;
      long amt_991310 = 0;
      long amt_991320 = 0;
      long amt_992810 = 0;    
      long amt_990812 = 0;
      long amt_990813 = 0;
      long amt_990814 = 0;
      long amt_990421 = 0;
      long amt_990422 = 0;//104.04.16 add
      long amt_990621 = 0;
      long amt_990622 = 0;//104.04.16 add
      long amt_990623 = 0;//110.03.30 add
      long amt_99141Y = 0;
      long amt_996114=0;//104.04.22 add
      long amt_996115=0;//104.04.22 add
      long amt_990711=0;//108.09.10 add
	  long amt_990712=0;//108.09.10 add 
      double a05bis=0;//102.01.17 add
      long field_over=0;//106.05.22 add 逾期放款
      long field_backup=0;//106.05.22 add 備抵呆帳
      long field_debit=0;//111.05.30 add 存款總額
      double a01field_over_rate=0;//102.01.17 add
      double a01field_backup_over_rate=0;//102.01.17 add備呆占狹義逾期放款比率(備抵呆帳覆蓋率)
      double a01field_backup_credit_rate=0;//104.04.22 add備呆占放款比率
      double a01field_backup_over_rate_avg=0;//104.04.22 add備呆占狹義逾期放款比率平均值
      double a01fieldi_y=0;//108.09.10 add 存放比率-存款總餘額
      String a02_amt_name="";
      
      for(int i=0;i<dbData.size();i++){
    	  System.out.println("test1");
      	  bean = (DataObject) dbData.get(i);
          bank_name = (String)bean.getValue("bank_name");
          props.setProperty("&UNIT",bank_name);         
          areaNo = bean.getValue("hsien_div_1") == null ?"" :(String)bean.getValue("hsien_div_1");
          amt_990110 = Long.parseLong( ( bean.getValue("amt990110")).toString());         
          amt_990120 = Long.parseLong( ( bean.getValue("amt990120")).toString());        
          amt_990130 = Long.parseLong( ( bean.getValue("amt990130")).toString());         
          amt_990140 = Long.parseLong( ( bean.getValue("amt990140")).toString());         
          amt_990150 = Long.parseLong( ( bean.getValue("amt990150")).toString());          
          amt_990210 = Long.parseLong( ( bean.getValue("amt990210")).toString());         
          amt_990220 = Long.parseLong( ( bean.getValue("amt990220")).toString());       
          amt_990230 = Long.parseLong( ( bean.getValue("amt990230")).toString());        
          amt_990240 = Long.parseLong( ( bean.getValue("amt990240")).toString());         
          amt_990310 = Long.parseLong( ( bean.getValue("amt990310")).toString());     
          amt_990320 = Long.parseLong( ( bean.getValue("amt990320")).toString());         
          amt_990410 = Long.parseLong( ( bean.getValue("amt990410")).toString());       
          amt_990420 = Long.parseLong( ( bean.getValue("amt990420")).toString()); 
          amt_990510 = Long.parseLong( ( bean.getValue("amt990510")).toString()); 
          System.out.println("test2");
          //95.09.27 add 
          amt_990511 = Long.parseLong( ( bean.getValue("amt990511")).toString()); 
          amt_990610 = Long.parseLong( ( bean.getValue("amt990610")).toString());
          //95.09.27 add
          amt_990611 = Long.parseLong( ( bean.getValue("amt990611")).toString());  
          //97.01.30 add         
          amt_990612 = Long.parseLong( ( bean.getValue("amt990612")).toString()); 
          amt_990620 = Long.parseLong( ( bean.getValue("amt990620")).toString());  
          amt_990630 = Long.parseLong( ( bean.getValue("amt990630")).toString());        
          amt_990710 = Long.parseLong( ( bean.getValue("amt990710")).toString());         
          amt_990720 = Long.parseLong( ( bean.getValue("amt990720")).toString());          
          amt_990721 = Long.parseLong( ( bean.getValue("amt990721")).toString());//108.01.08 add
          amt_990722 = Long.parseLong( ( bean.getValue("amt990722")).toString());//108.01.08 add
          amt_990810 = Long.parseLong( ( bean.getValue("amt990810")).toString());     
          System.out.println("test2_1");
          amt_990860 = Long.parseLong( ( bean.getValue("amt990860")).toString());//111.05.30 add
          amt_990870 = Long.parseLong( ( bean.getValue("amt990870")).toString());//111.05.30 add
          amt_990880 = Long.parseLong( ( bean.getValue("amt990880")).toString());//111.05.30 add
          System.out.println("test2_2");
          amt_990910 = Long.parseLong( ( bean.getValue("amt990910")).toString());        
          amt_990920 = Long.parseLong( ( bean.getValue("amt990920")).toString());         
          amt_990930 = Long.parseLong( ( bean.getValue("amt990930")).toString());         
          amt_991010 = Long.parseLong( ( bean.getValue("amt991010")).toString());         
          amt_991020 = Long.parseLong( ( bean.getValue("amt991020")).toString());         
          amt_991030 = Long.parseLong( ( bean.getValue("amt991030")).toString());        
          amt_991110 = Long.parseLong( ( bean.getValue("amt991110")).toString());         
          amt_991120 = Long.parseLong( ( bean.getValue("amt991120")).toString());          
          amt_991210 = Long.parseLong( ( bean.getValue("amt991210")).toString());         
          amt_991220 = Long.parseLong( ( bean.getValue("amt991220")).toString());          
          amt_991310 = Long.parseLong( ( bean.getValue("amt991310")).toString());         
          amt_991320 = Long.parseLong( ( bean.getValue("amt991320")).toString());         
          //99.01.21 992810投資全國農業金庫尚未攤提之損失
          amt_992810 = Long.parseLong( ( bean.getValue("amt992810")).toString());
          //101.08.23 990812因購置或汰換安全維護或營業相關設備，經中央主管機關核准/990813因固定資產重估增值/990814因淨值降低
          amt_990812 = Long.parseLong( ( bean.getValue("amt990812")).toString());
          amt_990813 = Long.parseLong( ( bean.getValue("amt990813")).toString());
          amt_990814 = Long.parseLong( ( bean.getValue("amt990814")).toString());
          //102.01.15 990421已申請符合逾放比率低於百分之一、其資本適足率高於百分之十且備抵保帳覆蓋率高於百分之一百經主管機關同意
          //          990621申請符合逾放比率低於百分之二且資本適足率高於百分之八經主管機關同意
          amt_990421 = Long.parseLong( ( bean.getValue("amt990421")).toString());
          amt_990621 = Long.parseLong( ( bean.getValue("amt990621")).toString());
          //104.04.22 990422已申請符合逾放比率低於百分之一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意
          //          990622已申請符合逾放比率低於百分之一、放款覆蓋率高於百分之二、其資本適足率高於百分之十且備抵呆帳覆蓋率高於全體信用部備抵呆帳覆蓋率平均值及不低於百分之一百經主管機關同意
          //          996114對鄉(鎮、市)公所授信未經所隸屬之縣政府保證
          //          996115對直轄市、縣(市)政府投資經營之公營事業，其授信經該直轄市、縣(市)政府保證
          amt_990422 = Long.parseLong( ( bean.getValue("amt990422")).toString());
          amt_990622 = Long.parseLong( ( bean.getValue("amt990622")).toString());
          amt_990623 = Long.parseLong( ( bean.getValue("amt990623")).toString());
          amt_996114 = Long.parseLong( ( bean.getValue("amt996114")).toString());
          amt_996115 = Long.parseLong( ( bean.getValue("amt996115")).toString());
          amt_990711 = Long.parseLong( ( bean.getValue("amt990711")).toString());//108.09.10
          amt_990712 = Long.parseLong( ( bean.getValue("amt990712")).toString());//108.09.10
          field_over = Long.parseLong( ( bean.getValue("field_over")).toString());//106.05.22 add
          field_backup = Long.parseLong( ( bean.getValue("field_backup")).toString());//106.05.22 ad
          System.out.println("test3");
          field_debit = Long.parseLong( ( bean.getValue("field_debit")).toString());//111.05.30 add
          System.out.println("test4");
          System.out.println("990210 ="+amt_990210);
          System.out.println("990220 ="+amt_990220);
          System.out.println("上年度 = "+l_year);
          amt_99141Y = bean.getValue("99141Y") == null ?0 :Long.parseLong( ( bean.getValue("99141Y")).toString());
          a05bis = Double.parseDouble( ( bean.getValue("a05bis")).toString());
          a01field_over_rate = Double.parseDouble( ( bean.getValue("a01field_over_rate")).toString());
          a01field_backup_over_rate = Double.parseDouble( ( bean.getValue("a01field_backup_over_rate")).toString());            
          a01field_backup_credit_rate = Double.parseDouble( ( bean.getValue("a01field_backup_credit_rate")).toString());
          a01field_backup_over_rate_avg = Double.parseDouble( ( bean.getValue("a01field_backup_over_rate_avg")).toString());
          a01fieldi_y = Double.parseDouble( ( bean.getValue("a01fieldi_y")).toString());//108.09.10 add	
          
         //95.09.27 fix (一)月底存放比率-->比率用『月底存放比率(A01)_合庫版』再比對『鄉鎮地區農會不得超過80%，直轄市及省、縣轄市農會不得超過78%』          
         props.setProperty("&101",dft_int.format(Math.round(amt_990110/1000)));//A
         props.setProperty("&102",dft_int.format(Math.round(amt_990120/1000)));//B
         props.setProperty("&103",dft_int.format(Math.round(amt_990130/1000)));//C
         props.setProperty("&104",dft_int.format(Math.round(amt_990140/1000)));//D
         props.setProperty("&105",dft_int.format(Math.round(amt_990150/1000)));//E         
            
         if((amt_990140-amt_990150/2)!=0){
             if(amt_990120 > amt_990130){//IF B>C-->(A-B+C)/(D-E*1/2)
                tmp_A = (double)(((double)amt_990110 - (double)amt_990120 + (double)amt_990130) / ((double)amt_990140 - (double)amt_990150 / 2));
             }else{//IF B<=C-->A/(D-E*1/2)
                tmp_A = (double)((double)amt_990110 / ((double)amt_990140 - (double)amt_990150 / 2));
             }           
         }
         props.setProperty("&106",String.valueOf(Math.round(tmp_A*100)));
          
         //(二)農會信用部對農會經濟事業部門融通資金之限制
         tmp_A = 0;
         tmp_B = 0;
         props.setProperty("&LY1",l_year);
         props.setProperty("&201",dft_int.format(Math.round(amt_990210/1000)));
         props.setProperty("&202",dft_int.format(Math.round(amt_990220/1000)));
         //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&203",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
         if((amt_990230-amt_990240-amt_992810)!=0){
           tmp_A = (double)amt_990210/((double)amt_990230-(double)amt_990240-(double)amt_992810);
           tmp_B = (double)amt_990220/((double)amt_990230-(double)amt_990240-(double)amt_992810);
           props.setProperty("&204",String.valueOf(Math.round(tmp_A*100)));
           props.setProperty("&205",String.valueOf(Math.round(tmp_B*100)));
         }else{
           props.setProperty("&204","0");
           props.setProperty("&205","0");
         }
         //(三)非會員存款之額度限制//104.04.22 取消(三)限制
         tmp_A=0;
         props.setProperty("&LY2",l_year);
         //95.09.27 fix 非會員存款總額(990310)-1/2(990630)
         tmp_A = ((double)amt_990310-((double)amt_990630/2));
         System.out.println("field_990310-(990630)/2="+ tmp_A);
         props.setProperty("&301",dft_int.format(Math.round(tmp_A/1000)));
         //99.01.21 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&302",dft_int.format(Math.round((amt_990320-amt_992810)/1000)));
         if((amt_990320-amt_992810)!=0){
           tmp_A = tmp_A/((double)amt_990320-(double)amt_992810);           
           //95.09.27 fix 倍->% , 10倍=1000%
           System.out.println("field_990310-(990630)/2 / (990320-992810)="+ tmp_A);
           props.setProperty("&303",String.valueOf(Math.round(tmp_A * 100)));
         }else{
           props.setProperty("&303","0");
         }
         //(三)贊助會員授信總額占贊助會員存款總額之比例不得超過100%之規定//104.04.22 原(四)調整為(三)
         //102.01.17 add (信用部逾放比< 2%  且 BIS > 8%者 得不超過 150%)
         //102.01.17 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
         //104.04.22 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制)
         //107.04.09 fix 信用部逾放比< 1% 且BIS > 10 且放款覆蓋率> 2%,已申請經主管機關同意者,不受限制;取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%
         tmp_A=0;
         //long limit4percent = 100;
         String limit4percent = "";
         props.setProperty("&401",dft_int.format(Math.round(amt_990410/1000)));
         props.setProperty("&402",dft_int.format(Math.round(amt_990420/1000)));
         if(amt_990420!=0){
           tmp_A=(double)amt_990410/(double)amt_990420;
           props.setProperty("&403",String.valueOf(Math.round(tmp_A*100)));
         }else{
           props.setProperty("&403","0");
         }
         
         if(Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month) < 11005){        	 
        	 limit4percent = "不得超過100%";
        	 //107.05.02 add 依所勾選的套用適用限額
        	 if(amt_990422 > 0){
        		 limit4percent = "不受限制";
        	 }else if(amt_990421 > 0){
        		 limit4percent = "不得超過200%";
        	 }else{
        		 if(a01field_over_rate < 2 && a05bis > 8){   
        			 limit4percent = "不得超過150%";
        		 } 
        		 if(a01field_over_rate > 2 || a05bis < 8){   
        			 limit4percent = "不得超過100%";
        		 }
        	 }
         }else{
        	 a02_amt_name="";
        	 sb_SQL.delete(0, sb_SQL.length());
        	 paramList.clear();
        	 dbData.clear();
        	 sb_SQL.append(" select acc_code,decode(NVL(amt_name2,''),'','不得超過'|| NVL(amt_name1,'')||'%',amt_name2) as amt_name ");
        	 sb_SQL.append(" from a02 where m_year=? and m_month=? and bank_code=? and acc_code in ('990422') ");
        	 paramList.add(s_year);
             paramList.add(s_month);
        	 paramList.add(bank_code);
        	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
        	 for(int j=0;j<dbData.size();j++){
        		 bean = (DataObject) dbData.get(j);
        		 a02_amt_name = (String)bean.getValue("amt_name");
        	 }
        	 if(amt_990422 > 0){
        		 limit4percent = a02_amt_name;
        	 }else if(amt_990421 > 0){
        		 limit4percent = "不得超過200%";
        	 }else{
        		 limit4percent = "不得超過150%";
        	 }
         }
         /*107.05.02 原檢核條件
         //102.01.17 add 不得超過150% or 200%
         if(a01field_over_rate < 2 && a05bis > 8){   
             limit4percent = "不得超過150%";//106.06.06 add
             //102.01.17 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
             if(a01field_over_rate < 1 && a05bis > 10){                
                if(a01field_backup_over_rate > 100){             
                   if(amt_990421 > 0){//102.01.15 add 990421             
                      limit4percent = "不得超過200%";
                   }                   
                }   
                
                //104.04.22 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且
                //              備抵呆帳覆蓋率高 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制
                //107.04.09 fix 取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%
                //if(a01field_backup_over_rate > a01field_backup_over_rate_avg//107.04.09 取消此項檢核 
                if(a01field_backup_credit_rate > 2 && amt_990422 > 0){//104.04.22 add 990422
                   limit4percent = "不受限制";
                }
                
                //106.05.22 add 若逾期放款為0,備抵呆帳覆蓋率(field_backup_over_rate)因分母為0,則該比率顯示為N/A,調整(四)符合為不受限制及200%
                if(field_over == 0 && a01field_backup_over_rate == 0  && field_backup > 0){
                   //102.09.23 add 若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
                   if(amt_990421 > 0){//102.01.15 add 990421             
                      limit4percent = "不得超過200%";
                   }
                   //104.04.22 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且
                   //              備抵呆帳覆蓋率高 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制                   
                   if(a01field_backup_credit_rate > 2 && amt_990422 > 0){//104.04.22 add 990422
                      limit4percent = "不受限制";
                   }
                }
             }else{
                limit4percent = "不得超過150%";
             }
         }
         */
         props.setProperty("&404",limit4percent);
         
         
         
         
         
         //(四)非會員授信總額占非會員存款總額之比率,不得超過100%之規定//110.03.30 原(六)調整為(四)  
         // 95.09.27 fix 非會員授信總額(990610) -（990511）- (990611)
         //				 扣除1.非會員無擔保消費政策性貸款(990511)(94.10月起).
         //			        2.對直轄市、縣（市）政府辦理之授信總額(990611) (95.1月起).
         // 97.01.30 add 990612 非會員政策性農業專案貸款 
         //          fix 非會員授信總額(990610) -（990511）- (990611) - (990612)[97.2月起]
         //100.10.14 fix 非會員存款總額(990620) -1/2(990630)
         //102.01.17 add 逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150%
         //104.03.18 add 取消非會員存款總額(990620) -1/2(990630)
         //104.04.22 add 逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%)
         //106.10.06 add 非會員授信總額.移除扣除990511非會員無擔保消費性政策貸款
         //          fix 非會員授信總額(990610) - (990611) - (990612)
         //107.04.09 fix 逾放比< 1% 且BIS > 10% 且放款覆蓋率> 2%,已申請經主管機關同意者,得不超過 200%;取消備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%
         tmp_A = 0; 
         tmp_A = (double)amt_990610;
         //100.10.14 fix 非會員存款總額(990620) -1/2(990630)//104.03.18取消
         //tmp_B = ((double)amt_990620-((double)amt_990630/2));//104.03.18取消扣除1/2公庫存款
         tmp_B = (double)amt_990620;
         //System.out.println("field_990620-(990630)/2="+ tmp_B);
         long limit6percent = 100;
         String limit6percentStr = "";
         if(Integer.parseInt(s_year+s_month) >= 9410){
            //tmp_A = tmp_A - amt_990511;//106.10.06 add 移除扣除990511非會員無擔保消費性政策貸款
            if(Integer.parseInt(s_year+s_month) >= 9501){
                tmp_A = tmp_A - amt_990611;
            }
            if(Integer.parseInt(s_year+s_month) >= 9702){//97.01.30 add 增加990612非會員政策性農業專案貸款
                tmp_A = tmp_A - amt_990612;
            }
         }
         props.setProperty("&601",dft_int.format(Math.round(tmp_A/1000)));
         props.setProperty("&602",dft_int.format(Math.round(tmp_B/1000)));
         if(tmp_B!=0){
           tmp_A = tmp_A/tmp_B;
           props.setProperty("&603",String.valueOf(Math.round(tmp_A*100)));
         }else{
           props.setProperty("&603","0");
         }
         
         if(Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month) < 11005){
        	 //107.05.02 add 依所勾選的套用適用限額
        	 if(amt_990622 > 0){
        		 limit6percent = 200;
        	 }else if(amt_990621 > 0){
        		 limit6percent = 150;
        	 }else{
        		 limit6percent = 100;
        	 }
        	 limit6percentStr = String.valueOf(limit6percent);
         }else{
        	 a02_amt_name="";
        	 sb_SQL.delete(0, sb_SQL.length());
        	 paramList.clear();
        	 dbData.clear();
        	 sb_SQL.append(" select acc_code,decode(NVL(amt_name2,''),'','不得超過'|| NVL(amt_name1,'')||'%',amt_name2) as amt_name ");
        	 sb_SQL.append(" from a02 where m_year=? and m_month=? and bank_code=? and acc_code in ('990623') ");
        	 paramList.add(s_year);
             paramList.add(s_month);
        	 paramList.add(bank_code);
        	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
        	 for(int j=0;j<dbData.size();j++){
        		 bean = (DataObject) dbData.get(j);
        		 a02_amt_name = (String)bean.getValue("amt_name");
        	 }
        	 if(amt_990623 > 0){
        		 limit6percentStr = a02_amt_name;
        	 }else if(amt_990622 > 0){
        		 limit6percentStr = "不得超過200%";
        	 }else{
        		 limit6percentStr = "不得超過150%";
        	 }
         }
         /*107.05.02 原檢核條件
         //102.01.17 add 逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150%
         if(a01field_over_rate < 2 && a05bis > 8 ){
            //102.01.17 add 不得超過150%
            if(amt_990621 > 0){
               limit6percent = 150;
            }             
            //104.02.24 add 逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,
            //              已申請經主管機關同意者,得不超過 200%)
            //107.04.09 fix 取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%
            if(a01field_over_rate < 1 && a05bis > 10){
               //if(a01field_backup_over_rate > 100 && a01field_backup_over_rate > a01field_backup_over_rate_avg//107.04.09 取消此條件檢核 
               if(a01field_backup_credit_rate > 2 && amt_990622 > 0){//104.04.22 add 990622
                  limit6percent = 200;
               }
               //106.05.22 add 若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
               if(field_over == 0 && a01field_backup_over_rate == 0 && field_backup > 0 ){                  
                  if(a01field_backup_credit_rate > 2 && amt_990622 > 0){//106.05.19 add
                      limit6percent = 200;
                  }
               }
            }   
         }
         */
         props.setProperty("&604",limit6percentStr);
         
         
         //(五)辦理非會員擔保放款徵提之擔保品種類//110.03.30 add
         a02_amt_name="";
    	 sb_SQL.delete(0, sb_SQL.length());
    	 paramList.clear();
    	 dbData.clear();
    	 sb_SQL.append(" select acc_code,decode(amt,'1','逾放比率低於5%,已申請經主管機關同意者,得以土地或建築物等不動產、動產為擔保品。','2','逾放比率高於5%未達10%,已申請經主管機關同意者, 得以住宅、已取得建築執照或雜項執照之建築基地為擔保品。','3','逾放比率高10%,得以住宅為擔保品。','') as amt_name ");
    	 sb_SQL.append(" from a02 where m_year=? and m_month=? and bank_code=? and acc_code in ('990624') ");
    	 paramList.add(s_year);
         paramList.add(s_month);
    	 paramList.add(bank_code);
    	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
    	 for(int j=0;j<dbData.size();j++){
    		 bean = (DataObject) dbData.get(j);
    		 a02_amt_name = (String)bean.getValue("amt_name");
    	 }
    	 if(a02_amt_name!=null && !"".equals(a02_amt_name)){
    		 props.setProperty("&504",a02_amt_name);
    	 }
         
         //(六)辦理非會員擔保放款徵提之擔保品坐落位置 //110.03.30 add
    	 a02_amt_name="";
    	 sb_SQL.delete(0, sb_SQL.length());
    	 paramList.clear();
    	 dbData.clear();
    	 sb_SQL.append(" select acc_code,decode(amt,'1','符合逾放比率低於1%、放款覆蓋率高於2%、資本適足率高於10%，或逾放比率低於1%、放款覆蓋率高於1.5%且在2%以下、資本適足率高於14%者，與其他直轄市或縣、市無直接接壤者，依所訂「授信政策、徵信及授信處理流程與作業手冊」，得徵提該縣市及'||hsien_name||'之擔保品。','2','符合逾放比率低於1%、放款覆蓋率高於2%、資本適足率高於10%，或逾放比率低於1%、放款覆蓋率高於1.5%且在2%以下、資本適足率高於14%者，與其他直轄市或縣、市有直接接壤者，依所訂「授信政策、徵信及授信處理流程與作業手冊」，得徵提該縣市及下列區域之擔保品：','3','未符合逾放比率低於1%、放款覆蓋率高於2%、資本適足率高於10%，或逾放比率低於1%、放款覆蓋率高於1.5%且在2%以下、資本適足率高於14%者，與其他直轄市或縣、市無直接接壤者，僅得辦理該縣市之擔保品。','4','未符合逾放比率低於1%、放款覆蓋率高於2%、資本適足率高於10%，或逾放比率低於1%、放款覆蓋率高於1.5%且在2%以下、資本適足率高於14%者，與其他直轄市或縣、市有直接接壤者，依所訂「授信政策、徵信及授信處理流程與作業手冊」，得辦理該縣市及下列區域之擔保品：','') as amt_name ");
    	 sb_SQL.append(" from a02 left join (select hsien_id,hsien_name from cd01 order by fr001w_output_order)cd01 on cd01.hsien_id = amt_name ");
    	 sb_SQL.append(" where m_year=? and m_month=? and bank_code=? and acc_code in ('990626') ");
    	 paramList.add(s_year);
         paramList.add(s_month);
    	 paramList.add(bank_code);
    	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
    	 for(int j=0;j<dbData.size();j++){
    		 bean = (DataObject) dbData.get(j);
    		 a02_amt_name = (String)bean.getValue("amt_name");
    	 }
    	 if(a02_amt_name!=null && !"".equals(a02_amt_name)){
    		 props.setProperty("&605",a02_amt_name);
    	 }else{
    		 props.setProperty("&605","");    		 
    	 }
    	 
    	 a02_amt_name="";
    	 sb_SQL.delete(0, sb_SQL.length());
    	 paramList.clear();
    	 dbData.clear();
    	 sb_SQL.append(" select  NVL(cd01.hsien_name||cd02.area_name,'') as amt_name ");
    	 sb_SQL.append(" from a02 left join (select hsien_id,hsien_name from cd01 order by fr001w_output_order)cd01 on cd01.hsien_id =  substr(amt_name1,1,1) ");
    	 sb_SQL.append("          left join (select area_id,area_name,hsien_id from cd02)cd02 on cd02.hsien_id=substr(amt_name1,1,1) and cd02.area_id = substr(amt_name1,3,3) ");
    	 sb_SQL.append(" where m_year=? and m_month=? and bank_code=? and acc_code in ('990626') and amt in (2,4) ");
    	 paramList.add(s_year);
    	 paramList.add(s_month);
    	 paramList.add(bank_code);
    	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
    	 for(int j=0;j<dbData.size();j++){
    		 bean = (DataObject) dbData.get(j);
    		 a02_amt_name = (String)bean.getValue("amt_name");
    	 }
    	 if(a02_amt_name!=null && !"".equals(a02_amt_name)){
    		 props.setProperty("&606",a02_amt_name);
    	 }else{
    		 props.setProperty("&606","");    		 
    	 }
    	 
    	 a02_amt_name="";
    	 sb_SQL.delete(0, sb_SQL.length());
    	 paramList.clear();
    	 dbData.clear();
    	 sb_SQL.append(" select  NVL(cd01.hsien_name||cd02.area_name,'') as amt_name ");
    	 sb_SQL.append(" from a02 left join (select hsien_id,hsien_name from cd01 order by fr001w_output_order)cd01 on cd01.hsien_id =  substr(amt_name2,1,1) ");
    	 sb_SQL.append("          left join (select area_id,area_name,hsien_id from cd02)cd02 on cd02.hsien_id=substr(amt_name2,1,1) and cd02.area_id = substr(amt_name2,3,3) ");
    	 sb_SQL.append(" where m_year=? and m_month=? and bank_code=? and acc_code in ('990626') and amt in (2,4) ");
    	 paramList.add(s_year);
    	 paramList.add(s_month);
    	 paramList.add(bank_code);
    	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
    	 for(int j=0;j<dbData.size();j++){
    		 bean = (DataObject) dbData.get(j);
    		 a02_amt_name = (String)bean.getValue("amt_name");
    	 }
    	 if(a02_amt_name!=null && !"".equals(a02_amt_name)){
    		 props.setProperty("&607",a02_amt_name);
    	 }else{
    		 props.setProperty("&607","");    		 
    	 }
    	 
    	 a02_amt_name="";
    	 sb_SQL.delete(0, sb_SQL.length());
    	 paramList.clear();
    	 dbData.clear();
    	 sb_SQL.append(" select NVL(hsien_name,'') as amt_name ");
    	 sb_SQL.append(" from a02 left join (select hsien_id,hsien_name from cd01 order by fr001w_output_order)cd01 on cd01.hsien_id = amt_name ");
    	 sb_SQL.append(" where m_year=? and m_month=? and bank_code=? and acc_code in ('990626') and amt in (2) ");
    	 paramList.add(s_year);
    	 paramList.add(s_month);
    	 paramList.add(bank_code);
    	 dbData = DBManager.QueryDB_SQLParam(sb_SQL.toString(),paramList,"");
    	 for(int j=0;j<dbData.size();j++){
    		 bean = (DataObject) dbData.get(j);
    		 a02_amt_name = (String)bean.getValue("amt_name");
    	 }
    	 if(a02_amt_name!=null && !"".equals(a02_amt_name)){
    		 props.setProperty("&608",a02_amt_name);
    	 }else{
    		 props.setProperty("&608","");    		 
    	 }
    	 
    	 
    	 //(七)辦理非會員無擔保消費性貸款(1,000千元以下)之限額//110.03.30 原(五)調整為(七)
         //106.10.06 add 無擔保消費性貸款增加扣除990511非會員無擔保消費性政策貸款
     	 //111.09.27 add 農會上年度決算淨值增加扣除990240上次檢查應補提未補提足之備抵呆帳
         tmp_A = 0;
         props.setProperty("&LY3",l_year);
         //106.10.06 add 無擔保消費性貸款增加扣除990511非會員無擔保消費性政策貸款
         props.setProperty("&501",dft_int.format(Math.round((amt_990510-amt_990511)/1000)));//106.10.06 add 增加扣除990511
         //99.01.21 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         //111.09.27 add 增加扣除990240
         props.setProperty("&502",dft_int.format(Math.round((amt_990320-amt_992810-amt_990240)/1000)));
         if((amt_990320-amt_992810-amt_990240)!=0){
           tmp_A = ((double)amt_990510-(double)amt_990511)/((double)amt_990320-(double)amt_992810-(double)-amt_990240);
           props.setProperty("&503",dft_int.format(Math.round(tmp_A*100)));
         }else{
           props.setProperty("&503","0");
         }
    	 
    	 
         //(八)違反購置住宅放款及房屋修繕放款之餘額不得超過存款總餘額55%之規定//110.03.30原(六)調整為(八)
         //108.09.09 add 108年10月以後套用新格式
         if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10810){
        	 tmp_A = 0;
             props.setProperty("&701",dft_int.format(Math.round(amt_990711/1000)));
             props.setProperty("&1405",dft_int.format(Math.round(amt_990712/1000)));            
             props.setProperty("&1406",dft_int.format(Math.round(a01fieldi_y/1000)));//108.01.08 add
             System.out.println("amt_990711="+amt_990711);
             System.out.println("amt_990712="+amt_990712);
             if(a01fieldi_y!=0){
               tmp_A = ((double)amt_990711+(double)amt_990712)/(double)a01fieldi_y;              
               props.setProperty("&703",String.valueOf(Math.round(tmp_A*100)));
             }else{
               props.setProperty("&703","0");
             }
         
         }else{	 
             //(七)辦理自用住宅放款限額
             tmp_A = 0;
             props.setProperty("&701",dft_int.format(Math.round(amt_990710/1000)));
             props.setProperty("&702",dft_int.format(Math.round(amt_990720/1000)));
             props.setProperty("&1405",dft_int.format(Math.round(amt_990721/1000)));//108.01.08 add
             props.setProperty("&1406",dft_int.format(Math.round(amt_990722/1000)));//108.01.08 add
             System.out.println("amt_990721="+amt_990721);
             System.out.println("amt_990722="+amt_990722);
             if(amt_990720!=0){
               tmp_A =(double)amt_990710/(double)amt_990720;
               props.setProperty("&703",String.valueOf(Math.round(tmp_A*100)));
             }else{
               props.setProperty("&703","0");
             }
         }
         //(九)固定資產淨額限制//110.03.30原(八)調整為(九)
         tmp_A = 0;
         props.setProperty("&LY4",l_year);
         props.setProperty("&801",dft_int.format(Math.round(amt_990810/1000)));
         //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&802",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
         if((amt_990230-amt_990240-amt_992810)!=0){
           tmp_A = (double)amt_990810/((double)amt_990230-(double)amt_990240-(double)amt_992810);
           props.setProperty("&803",String.valueOf(Math.round(tmp_A*100)));
         }else{
           props.setProperty("&803","0");
         }
         //101.08.23 add 990812一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
         //              990813二、因固定資產重估增值
         //              990814三、因淨值降低
         //=================================================================
         if(amt_990812 > 0){//一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
            props.setProperty("&804","(V) ");
         }else{
            props.setProperty("&804","   ");
         }
         if(amt_990813 > 0){//二、因固定資產重估增值
             props.setProperty("&805","(V) ");
         }else{
             props.setProperty("&805","   ");
         }
         if(amt_990814 > 0){//三、因淨值降低
             props.setProperty("&806","(V) ");
         }else{
             props.setProperty("&806","   ");
         }
         //================================================================
         //(十)外幣風險之限制//110.03.30原(九)調整為(十)
         //111.09.27 信用部上年度決算淨值增加扣除990240上次檢查應補提未補提足之備抵呆帳
         tmp_A = 0;
         props.setProperty("&LY5",l_year);
         props.setProperty("&901",dft_int.format(Math.round(amt_990910/1000)));
         props.setProperty("&902",dft_int.format(Math.round(amt_990920/1000)));
         //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&903",dft_int.format(Math.round((amt_990230-amt_992810-amt_990240)/1000)));
         if((amt_990230-amt_992810-amt_990240)!=0){
           tmp_A = Math.abs((double)amt_990910-(double)amt_990920)/((double)amt_990230-(double)amt_992810-(double)amt_990240);
           props.setProperty("&904",String.valueOf(Math.round(tmp_A*100)));
         }else{
           props.setProperty("&904","0");
         }
         //95.09.27 fix 94.10以後.10~13公式.漁會同農會=================================
         System.out.println("s_year.s_month="+Integer.parseInt(s_year+s_month));
         if(bank_type.equals("6") || (Integer.parseInt(s_year+s_month) >= 9410)){
           //(十一)對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者為擔保授信限制//110.03.30原(十)調整為(十一)
           //111.09.27 農會上年度決算淨值增加扣除990240上次檢查應補提未補提足之備抵呆帳
           //          修正為「未達九百萬元者，得以九百萬元為最高限額」	 
           tmp_A=0;
           props.setProperty("&LY6",l_year);
           props.setProperty("&LY7",l_year);
           //99.01.21 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           //111.09.27 農會上年度決算淨值增加扣除990240上次檢查應補提未補提足之備抵呆帳
           props.setProperty("&1001",dft_int.format(Math.round((amt_990320-amt_992810-amt_990240)/1000)));
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1002",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
           if((amt_990230-amt_990240-amt_992810)*0.25>9000000){
             props.setProperty("&1003",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           }else{
             props.setProperty("&1003",dft_int.format(Math.round(9000000/1000)));//111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
           }
           props.setProperty("&1004",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           props.setProperty("&1005",dft_int.format(Math.round(amt_991010/1000)));
           props.setProperty("&1006",dft_int.format(Math.round(amt_991020/1000)));
           //95.09.27 fix C/A不得超過1.5倍->991020 / (990320-992810) ,1.5倍=150%
           //99.01.21 fix 99.01開始.上年度全体農/漁會決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           //111.09.27 農會上年度決算淨值增加扣除990240上次檢查應補提未補提足之備抵呆帳
           if((amt_990320-amt_992810-amt_990240)!=0){
             tmp_A = (double)amt_991020/((double)amt_990320-(double)amt_992810-(double)amt_990240);
             props.setProperty("&1007",String.valueOf(Math.round(tmp_A*100)));             
           }else{
             props.setProperty("&1007","0");
           }
           //(十二)對每一會員(含同戶家屬)及同一關係人放款最高限額//110.03.30原(十一)調整為(十二)
           //111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
           props.setProperty("&LY8",l_year);
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1101",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
           if((amt_990230-amt_990240-amt_992810)*0.25>9000000){
             props.setProperty("&1102",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
             props.setProperty("&1103",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           }else{//111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
             props.setProperty("&1102",dft_int.format(Math.round(9000000/1000)));
             props.setProperty("&1103",dft_int.format(Math.round(9000000/1000)));
           }
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1104",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           props.setProperty("&1105",dft_int.format(Math.round(amt_991110/1000)));
           props.setProperty("&1106",dft_int.format(Math.round(amt_991110/1000)));
           if((amt_990230-amt_990240-amt_992810)*0.05>2000000){
             props.setProperty("&1107",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
             props.setProperty("&1108",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
           }else{
             props.setProperty("&1107",dft_int.format(Math.round(2000000/1000)));
             props.setProperty("&1108",dft_int.format(Math.round(2000000/1000)));
           }
           props.setProperty("&1109",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
           props.setProperty("&1110",dft_int.format(Math.round(amt_991120/1000)));
           props.setProperty("&1111",dft_int.format(Math.round(amt_991120/1000)));
           //(十三)對每一贊助會員及同一關係人之授信限額//110.03.30原(十二)調整為(十三)
           //111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
           props.setProperty("&LY9",l_year);
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1201",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
           if((amt_990230-amt_990240-amt_992810)*0.25>9000000){
             props.setProperty("&1202",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
             props.setProperty("&1203",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           }else{//111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
             props.setProperty("&1202",dft_int.format(Math.round(9000000/1000)));
             props.setProperty("&1203",dft_int.format(Math.round(9000000/1000)));
           }
           props.setProperty("&1204",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.25/1000)));
           props.setProperty("&1205",dft_int.format(Math.round(amt_991210/1000)));
           props.setProperty("&1206",dft_int.format(Math.round(amt_991210/1000)));
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           if((amt_990230-amt_990240-amt_992810)*0.05>2000000){
             props.setProperty("&1207",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
             props.setProperty("&1208",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
           }else{
             props.setProperty("&1207",dft_int.format(Math.round(2000000/1000)));
             props.setProperty("&1208",dft_int.format(Math.round(2000000/1000)));
           }
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1209",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.05/1000)));
           props.setProperty("&1210",dft_int.format(Math.round(amt_991220/1000)));
           props.setProperty("&1211",dft_int.format(Math.round(amt_991220/1000)));
           //(十四)對同一非會員及同一關係人之授信限額//110.03.30原(十三)調整為(十四)
           props.setProperty("&LY10",l_year);
           //95.09.27 fix 1301金額,以千元為單位===================================================
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1301",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
           //百分之12.5%
           if((amt_990230-amt_990240-amt_992810)*0.125>9000000){
             props.setProperty("&1302",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.125/1000)));             
           }else{//111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」
             props.setProperty("&1302",dft_int.format(Math.round(9000000/1000)));             
           }
           //111.09.27增加百分之15%
           if((amt_990230-amt_990240-amt_992810)*0.15>9000000){               
               props.setProperty("&1303",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.15/1000)));
           }else{//111.09.27 修正為「未達九百萬元者，得以九百萬元為最高限額」               
              props.setProperty("&1303",dft_int.format(Math.round(9000000/1000)));
           }
           //95.09.27 fix 1304.1305.1306金額,以千元為單位===================================================
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1304",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.125/1000)));
           //111.09.27 增加百分之15%
           props.setProperty("&1312",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.15/1000)));
           props.setProperty("&1305",dft_int.format(Math.round(amt_991310/1000)));
           props.setProperty("&1306",dft_int.format(Math.round(amt_991310/1000)));
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           //百分之2.5
           if((amt_990230-amt_990240-amt_992810)*0.025>2000000){
             props.setProperty("&1307",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.025/1000)));             
           }else{
             props.setProperty("&1307",dft_int.format(Math.round(2000000/1000)));             
           }
           //111.09.27增加百分之3
           if((amt_990230-amt_990240-amt_992810)*0.03>2000000){
               props.setProperty("&1308",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.03/1000)));
           }else{               
               props.setProperty("&1308",dft_int.format(Math.round(2000000/1000)));
           }
           //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
           props.setProperty("&1309",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.025/1000)));
           //111.09.27 增加百分之3%
           props.setProperty("&1314",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)*0.03/1000)));
           props.setProperty("&1310",dft_int.format(Math.round(amt_991320/1000)));
           props.setProperty("&1311",dft_int.format(Math.round(amt_991320/1000)));
           
           //104.04.22 add
           //(十五)對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證，及對直轄市、縣(市)政府投資經營之公營事業，其授信經該直轄市、縣(市)政府保證之限額
           //110.03.30原(十四)調整為(十五)
           props.setProperty("&LY11",l_year);
           props.setProperty("&1401",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
           props.setProperty("&1402",dft_int.format(Math.round(amt_996114/1000)));
           props.setProperty("&1403",dft_int.format(Math.round(amt_996115/1000)));
           tmp_A = 0;
           if((amt_990230-amt_990240-amt_992810)!=0){
               tmp_A = ((double)amt_996114+(double)amt_996115)/((double)amt_990230-(double)amt_990240-(double)amt_992810);
               props.setProperty("&1404",String.valueOf(Math.round(tmp_A*100)));
           }else{
               props.setProperty("&1404","0");
           }
           
         }else{//94.10以前.10~13公式.漁會公式
           //(十)對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者為擔保授信限制
           props.setProperty("&LY6",l_year);
           props.setProperty("&1001",dft_int.format(Math.round(amt_990320/1000)));
           props.setProperty("&1002",dft_int.format(Math.round((amt_990230-amt_990240)/1000)));
           props.setProperty("&1003",dft_int.format(Math.round(amt_991030/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.02>6000000){
             props.setProperty("&1004",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
           }else{
             props.setProperty("&1004",dft_int.format(Math.round(6000000/1000)));
           }
           props.setProperty("&1005",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
           props.setProperty("&1006",dft_int.format(Math.round(amt_991020/1000)));
           if(amt_990320!=0){
             tmp_A = 0;
             tmp_A = (double)amt_991020/(double)amt_990320;
             //String tmp_value = dft.format(tmp_A);
             //95.09.27 fix D/C不得超過1.5倍->991020 / 990320 ,1.5倍=150%
             props.setProperty("&1007",String.valueOf(Math.round(tmp_A*100)));
           }else{
             props.setProperty("&1007","0");
           }
           //(十一)對每一會員(含同戶家屬)及同一關係人放款最高限額
           props.setProperty("&1101",dft_int.format(Math.round((amt_990230-amt_990240)/1000)));
           props.setProperty("&1102",dft_int.format(Math.round(amt_991030/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.02>6000000){
             props.setProperty("&1103",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
             props.setProperty("&1104",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
           }else{
             props.setProperty("&1103",dft_int.format(Math.round(6000000/1000)));
             props.setProperty("&1104",dft_int.format(Math.round(6000000/1000)));
           }
           props.setProperty("&1105",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
           props.setProperty("&1106",dft_int.format(Math.round(amt_991110/1000)));
           props.setProperty("&1107",dft_int.format(Math.round(amt_991110/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.01>2000000){
             props.setProperty("&1108",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
             props.setProperty("&1109",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           }else{
             props.setProperty("&1108",dft_int.format(Math.round(2000000/1000)));
             props.setProperty("&1109",dft_int.format(Math.round(2000000/1000)));
           }
           props.setProperty("&1110",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           props.setProperty("&1111",dft_int.format(Math.round(amt_991120/1000)));
           props.setProperty("&1112",dft_int.format(Math.round(amt_991120/1000)));
           //(十二)對每一贊助會員及同一關係人之授信限額
           props.setProperty("&1201",dft_int.format(Math.round((amt_990230-amt_990240)/1000)));
           props.setProperty("&1202",dft_int.format(Math.round(amt_991030/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.02>6000000){
             props.setProperty("&1203",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
             props.setProperty("&1204",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/1000)));
           }else{
             props.setProperty("&1203",dft_int.format(Math.round(6000000/1000)));
             props.setProperty("&1204",dft_int.format(Math.round(6000000/1000)));
           }
           props.setProperty("&1205",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.02/100)));
           props.setProperty("&1206",dft_int.format(Math.round(amt_991210/1000)));
           props.setProperty("&1207",dft_int.format(Math.round(amt_991210/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.01>2000000){
             props.setProperty("&1208",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
             props.setProperty("&1209",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           }else{
             props.setProperty("&1208",dft_int.format(Math.round(2000000/1000)));
             props.setProperty("&1209",dft_int.format(Math.round(2000000/1000)));
           }
           props.setProperty("&1210",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           props.setProperty("&1211",dft_int.format(Math.round(amt_991220/1000)));
           props.setProperty("&1212",dft_int.format(Math.round(amt_991220/1000)));

           //(十三)對同一非會員及同一關係人之授信限額
           props.setProperty("&1301",dft_int.format(Math.round((amt_990230-amt_990240)/1000)));
           props.setProperty("&1302",dft_int.format(Math.round(amt_991030/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.01>6000000){
             props.setProperty("&1303",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
             props.setProperty("&1304",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           }else{
             props.setProperty("&1303",dft_int.format(Math.round(6000000/1000)));
             props.setProperty("&1304",dft_int.format(Math.round(6000000/1000)));
           }
           props.setProperty("&1305",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.01/1000)));
           props.setProperty("&1306",dft_int.format(Math.round(amt_991310/1000)));
           props.setProperty("&1307",dft_int.format(Math.round(amt_991310/1000)));
           if(((amt_990230-amt_990240)+amt_991030)*0.005>2000000){
             props.setProperty("&1308",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.005/1000)));
             props.setProperty("&1309",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.005/1000)));
           }else{
             props.setProperty("&1308",dft_int.format(Math.round(2000000/1000)));
             props.setProperty("&1309",dft_int.format(Math.round(2000000/1000)));
           }
           props.setProperty("&1310",dft_int.format(Math.round(((amt_990230-amt_990240)+amt_991030)*0.005/1000)));
           props.setProperty("&1311",dft_int.format(Math.round(amt_991320/1000)));
           props.setProperty("&1312",dft_int.format(Math.round(amt_991320/1000)));              		  
         }
         //111.05.30 add 16~18項法定比率
         //(十六) 持有非由政府發行之債券及票券之限額(農漁會信用部各項風險控制比率管理辦法第十一條)
         tmp_A = 0;
         tmp_B = 0;
        
         props.setProperty("&1601",dft_int.format(Math.round(amt_990860/1000)));
         props.setProperty("&1602",dft_int.format(Math.round(field_debit/1000)));         
         if(field_debit!=0){
           tmp_A = (double)amt_990860/(double)field_debit;           
           props.setProperty("&1603",String.valueOf(Math.round(tmp_A*100)));         
         }else{
           props.setProperty("&1603","0");           
         }
         //(十七) 持有單一銀行發行之金融債券及可轉讓定期存單原始取得成本之限額(農會漁會信用部各項風險控制比率管理辦法第十一條)
         tmp_A = 0;
         tmp_B = 0;
         props.setProperty("&LY17",l_year);
         props.setProperty("&1701",dft_int.format(Math.round(amt_990870/1000)));         
         //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&1702",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));
         if((amt_990230-amt_990240-amt_992810)!=0){
           tmp_A = (double)amt_990870/((double)amt_990230-(double)amt_990240-(double)amt_992810);         
           props.setProperty("&1703",String.valueOf(Math.round(tmp_A*100)));         
         }else{
           props.setProperty("&1703","0");         
         }
         
         //(十八) 持有單一企業發行之短期票券及公司債之限額(農會漁會信用部各項風險控制比率管理辦法第十一條)
         tmp_A = 0;
         tmp_B = 0;
         props.setProperty("&LY18",l_year);
         props.setProperty("&1801",dft_int.format(Math.round(amt_990880/1000)));             
         //99.01.21 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
         props.setProperty("&1802",dft_int.format(Math.round((amt_990230-amt_990240-amt_992810)/1000)));        
         if((amt_990230-amt_990240-amt_992810)!=0){
           tmp_A = (double)amt_990880/((double)amt_990230-(double)amt_990240-(double)amt_992810);           
           props.setProperty("&1803",String.valueOf(Math.round(tmp_A*100))); 
         }else{
           props.setProperty("&1803","0");           
         }
       }
      //其他法定比率END
      Calendar rightNow = Calendar.getInstance();
      String year = String.valueOf(rightNow.get(Calendar.YEAR) - 1911);
      String month = String.valueOf(rightNow.get(Calendar.MONTH) + 1);
      String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
      String date = year + "年" +month+"月"+day+"日";      
      props.setProperty("&PRINTDATEYM&",date);      
      props.setProperty("&CY&",s_year);
      props.setProperty("&CM&",s_month);     
      //95.10.03 增加檢核結果與最後異動日期
      dbData = Utility.getWML01UPD_CODE(s_year,s_month,bank_code,"A02");
      String UPD_CODE="";
      String UPDATE_DATE="";
      String M_YEAR="";   
      String M_MONTH="";    
      String M_DATE=""; 
      
      if(dbData.size()>0){
         System.out.println("dbData.size()="+dbData.size()); 
         bean = (DataObject)dbData.get(0);
         UPD_CODE = (String)bean.getValue("upd_code_name");  
         UPDATE_DATE = (String)bean.getValue("update_date");       		   
		 System.out.println("UPD_CODE="+UPD_CODE); 
		 System.out.println("UPDATE_DATE="+UPDATE_DATE);
		 M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
		 M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
		 M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
		 UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
		 UPD_CODE=(String)UPD_CODE;
		 UPDATE_DATE=(String)UPDATE_DATE;
		 System.out.println("UPDATE_DATE="+UPDATE_DATE); 
		 props.setProperty("&Check_Result",UPD_CODE);
         props.setProperty("&Last_Modify_Date",UPDATE_DATE);			   
      }else{
		 UPD_CODE="檢核結果:待檢核"; 
		 UPDATE_DATE="最後異動日期:無";
		 props.setProperty("&Check_Result",UPD_CODE);
         props.setProperty("&Last_Modify_Date",UPDATE_DATE);	      			
     }//end of 列印檢核結果與最後異動日期			
	 

      mix.mix(props);

    return errMsg;
  }
}
