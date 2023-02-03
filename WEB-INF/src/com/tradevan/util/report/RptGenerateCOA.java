/*
   105.10.26~105.11.03 create 產生農委會 OpenData所需報表 by 2295	
   107.02.26 add DS061W~DS064W報表 by 2295
   109.01.16 fix 調整FR0066W報表無法產生 by 2295
   109.09.11 fix 調整因報表格式轉換無法產生報表 by 2295
   111.06.09 fix 調整DS001W報表來源檔 by 2295
*/
package com.tradevan.util.report;



import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.util.*;
import java.text.SimpleDateFormat;
import javax.servlet.http.HttpSession;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.sftp.*;


public class RptGenerateCOA {	
	static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	static SimpleDateFormat filenameformat = new SimpleDateFormat("yyyyMMdd");
	static Date nowlog = new Date();
	static Calendar logcalendar;
	static File xlsDir= null;

	
	public static void main(String args[]) { 
	    	    //RptGenerateCOA a = new RptGenerateCOA();
	  } 
	 
    public static String createRpt(String S_YEAR,String S_MONTH,PrintStream logps,HttpSession session){
		String errMsg = "";
		List hsien_dbData = null; 
		List bank_no_6_dbData = null; 
		List bank_no_7_dbData = null; 
		List bank_no_67_dbData = null; 
		String URLStr ="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		String bank_code="";
		String bank_name = "";
		String bank_type = "";	
		String filename = "";		

		String errRptMsg = "";
		HSSFWorkbook wb = null;
		DataObject bean = null;
		
		/*年報*/
		String YRpt[][] = {{"AN001W","歷年來全體農漁會信用部簡明存款結構比較表.xls"},
		                   {"AN002W","歷年來全體農漁會信用部簡明放款結構比較表.xls"},
		                   {"AN003W","歷年來全體農漁會逾放、本期損益、淨值、資產總額一覽表.xls"},
		                   {"AN004W","多個年度農漁會信用部存款結構及變動表.xls"},
		                   {"AN004WB","多個年度農漁會信用部放款結構及變動表.xls"},
		                   {"AN005W","年底縣市別之全體農漁會信用部存款金額及存款平均餘額表.xls"},
		                   {"AN005WB","年底縣市別之全體農漁會信用部放款金額及放款平均餘額表.xls"},    
		                   {"AN006W","多個年度農漁會信用部簡明資產負債表.xls"},
		                   {"AN007W","多個年度農漁會信用部簡明損益比較表.xls"},
		                   {"AN008W","某年底全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表.xls"},
		                   {"AN009W","某年底各縣市農漁會信用部逾期放款、比率及存放款及比率彙總表.xls"},
		                   {"AN010W","某年底全體農漁會信用部各類財務比率.xls"},
		                   {"AN011W","某年底全體漁會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表.xls"},
		                   {"AN012W","某年底個別農漁會信用部業務經營概況.xls"},
		                   {"MC007W","農漁會信用部簽証會計師明細表.xls"}
		                   };
		/*季報*/                              
		String QRpt[][] = {{"DS014W","農漁會信用部金融機構警示帳戶調查資料.xls"},
		                   {"DS015W","農漁會信用部牌告利率每季申報資料.xls"},
		                   {"DS021W","合格自有資本與風險性資產比率計算表(1-A1).xls"},
		                   {"DS022W","自有資本計算表(1-B).xls"},
		                   {"DS023W","資本扣除項目總表(1-B1).xls"},
		                   {"DS038W","作業風險之資本計提計算表(5-A).xls"},
		                   {"DS051W","市場風險資本扣除項目彙總(6-G).xls"}
		                  };
		/*月報*/
		String MRpt[][] = {
		                   {"BR005W","農漁會分支機構資料表.xls"},
		                   {"BR010W","農漁會信用部通訊錄聯絡簡表.xls"},
		                   {"BR007W","農會分支機構分布表.xls"},
		                   {"BR022W","台閩地區農業金融從業人員統計.xls"},
		                   {"DS001W","農會信用部營運明細資料.xls"},		                             
		                   {"DS002W","農會信用部法定比率資料.xls"},
		                   {"DS003W","農會信用部平均利率資料.xls"},
		                   {"DS004W","農會信用部資產品質分析資料.xls"},
		                   {"DS005W","農會信用部資本適足率資料.xls"},
		                   {"DS006W","農會信用部逾期放款資料.xls"},
		                   {"DS007W","農會信用部在台無住所之外國人新台幣存款表.xls"},
		                   {"DS009W","農會信用部ATM裝設紀錄資料.xls"},
		                   {"DS011W","農會信用部支票存款資料.xls"},
		                   {"DS012W","農會信用部統一農漁貸資料.xls"},
		                   {"DS017W","農會信用部存款帳戶分級差異化管理統計資料.xls"},
		                   {"DS018W","農會信用部基本放款利率計價之舊貸案件資料.xls"},
		                   {"DS053W","保證案件月報表(M201).xls"},		                     
		                   {"DS054W","金融機構別及地區別保證案件分析表(M206).xls"},
		                   {"DS055W","農漁會1年內及目前月份經營統計資料.xls"},
		                   {"DS060W","各經辦機構逾期超過六個月放款金額統計資料.xls"},
		                   {"FR001W","台閩地區農會信用部經營指標.xls"},
		                   {"FR001WA","台閩地區農會信用部經營指標_A.xls"},
		                   {"FR001WE","臺閩地區農漁會信用部經營指標.xls"},
		                   {"FR001WN","N年內及目前月份經營概況趨勢統計分析.xls"},
		                   {"FR002W","全體農會按縣市別經營指標變化表.xls"},
		                   {"FR002WA","各縣市各農會信用部各年月經營指標變化表.xls"},
		                   {"FR003W","農業信用部資產負債表.xls"},
		                   {"FR004W","農業信用部損益表.xls"},
		                   {"FR004W_month","農業信用部損益表.xls"},
		                   {"FR005W","全體農會按縣市別平均利率.xls"},
		                   {"FR0066W","農會各項法定比率表.rtf"},
		                   {"FR006W","信用部違反法定比率規定分析表-農漁會.xls"},
		                   {"FR007W","農(漁)會資產品質分析彙總表.xls"},
		                   {"FR007WA","農會資產品質分析明細表.xls"},
		                   {"FR008W","信用部淨值占風險性資產比率.xls"},
		                   {"FR009W","信用部淨值占風險性資產比率計算表.xls"},
		                   {"FR023W","農會信用部各類對象存放款比率表.xls"},
		                   {"FR024W","全體農會信用部逾放比率超逾5%明細表.xls"},
		                   {"FR025W","全體農會信用部自動化機器彙計.xls"},
		                   {"FR025WA","農漁會信用部ATM裝設台數及異動明細表.xls"},
		                   {"FR026W","在台無住所之外國人新台幣存款表-農會.xls"},
		                   {"FR028W","台灣地區農會信用部放款餘額表.xls"},
		                   {"FR029W","全體農會信用部本期損益或淨值為負單位明細表.xls"},
		                   {"FR030WW_1","台灣區農會信用部支票存款戶數與餘額明細表.xls"},
		                   {"FR030WW_2","台灣區農會信用部支票存款戶數與餘額總表.xls"},
		                   {"FR030WW_3","台灣區農會信用部支票存款戶數與餘額彙計表.xls"},
		                   {"FR031W","農業信用保證基金業務統計(一).xls"},
		                   {"FR032W","農業信用保證基金業務統計(二).xls"},
		                   {"FR034WW","農漁會信用部警示帳戶調查統計總表.xls"},
		                   {"FR036WW_1","全體農會信用部統一農貸資料某二指定期間新增戶數金額比較總表.xls"},
		                   {"FR036WW_2","全體農會信用部統一農貸資料指定單月新增戶數金額比較總表.xls"},
		                   {"FR036WW_3","全體農會信用部統一農貸資料總表.xls"},
		                   {"FR037W","農漁會信用部逾期放款統計表.xls"},
		                   {"FR039W","農會信用部存放款餘額表_總表.xls"},
		                   {"FR040WW","農會存放款彙總表.xls"},
		                   {"FR045W","存款帳戶分級差異化管理統計表_農漁會.xls"},
		                   {"FR046W","全體農漁會信用部基本放款利率計價之舊貸案件總表.xls"},
		                   {"FR054W","全體農漁會信用部各會員別放款金額一覽表.xls"},
		                   {"FR056W","信用部主任及分部主任參訓情形總表.xls"},
		                   {"FR057W","全體農漁會信用部從業人員參加金融相關業務進修情形統計表.xls"},
		                   {"FR058W","全體農漁會信用部主任及分部主任參訓情形統計表.xls"},
		                   {"FR062W","農業信用保證基金業務統計.xls"},
		                   {"FR064W","農漁會信用部聯合貸款案件彙總表.xls"},
		                   {"FR065W","N年內及目前月份逾放比家數統計.xls"},
		                   {"FR067W","N年內及目前月份資本適足率家數統計.xls"},
		                   {"FR068W","逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表.xls"},
		                   {"MC004W","農漁會信用部違反農金局及其子法而遭處分明細表.xls"},
		                   {"DS061W","農業發展基金延期還款案件統計資料.xls"},
		                   {"DS062W","延期還款案件彙計表.xls"},
		                   {"DS063W","農家綜合貸款同戶申貸統計資料.xls"},		                  	                  
		                   {"DS064W","農家綜合貸款一借再借申貸統計資料.xls"}	                   
		                  };
		        
		try{			 
			Utility.printLogTime("RptGenerateCOA begin time");
			String rptIP_coa=Utility.getProperties("rptIP_coa");			
	        String rptID_coa=Utility.getProperties("rptID_coa");			
	        String rptPwd_coa=Utility.getProperties("rptPwd_coa");	
	        xlsDir = new File(Utility.getProperties("xlsDir"));
        	        
    		String SYear = "51";
    		String EYear = "";
    		int Month = 0;
    		String bankTpye="ALL";
    		String Unit="1000";
    		String copyOK="";
    		URLStr = Utility.getProperties("MISWebSite");//http://localhost:81/pages/
    		//URLStr = "http://localhost:96/pages/";
    		//各縣市代碼資料
    		sqlCmd.append("select hsien_id, hsien_name,fr001w_output_order from cd01 where hsien_id not in ('h','Y') order by fr001w_output_order");
    		hsien_dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");
            
    		//產生年報資料=========================================================    
    		for(int i=0;i<YRpt.length;i++){
    		    EYear = (!S_YEAR.equals("")?S_YEAR:Utility.getYear());
    		    Month = (!S_MONTH.equals("")?Integer.parseInt(S_MONTH):Integer.parseInt(Utility.getMonth()));
    		    //EYear="104";
    		    if(Month == 12){//年報為每年1月/30更新
    		       System.out.println(YRpt[i][0]+"create begin");
    		       if(YRpt[i][0].equals("AN001W")){
    		          errRptMsg = RptAN001W.createRpt(SYear,EYear, bankTpye, Integer.parseInt(Unit));
    		       }else if (YRpt[i][0].equals("AN002W")){
    		          errRptMsg = RptAN002W.createRpt(SYear,EYear, bankTpye, Integer.parseInt(Unit));
    		       }else if (YRpt[i][0].equals("AN003W")){
    		           errRptMsg = RptAN003W.createRpt(SYear,EYear, bankTpye, Integer.parseInt(Unit));
    		       }else if (YRpt[i][0].equals("AN004W")){
    		           errRptMsg = RptAN004W.createRpt(String.valueOf(Integer.parseInt(EYear)-5), EYear, bankTpye, Integer.parseInt(Unit));
    		       }else if (YRpt[i][0].equals("AN004WB")){
                       errRptMsg = RptAN004WB.createRpt(String.valueOf(Integer.parseInt(EYear)-5), EYear, bankTpye, Integer.parseInt(Unit));                      
                   }else if (YRpt[i][0].equals("AN005W")){                       
                       errRptMsg = RptAN005W.createRpt(EYear,bankTpye ,"仟元;"+Unit);                      
                   }else if (YRpt[i][0].equals("AN005WB")){    
                       errRptMsg = RptAN005WB.createRpt(EYear,bankTpye ,"仟元;"+Unit);
                   }else if (YRpt[i][0].equals("AN006W")){
                       errRptMsg = RptAN006W.createRpt(String.valueOf(Integer.parseInt(EYear)-5), EYear, bankTpye, Integer.parseInt(Unit));
                   }else if (YRpt[i][0].equals("AN007W")){
                       errRptMsg = RptAN007W.createRpt(EYear, bankTpye, Integer.parseInt(Unit));
                   }else if (YRpt[i][0].equals("AN008W")){    
                       errRptMsg = RptAN008W.createRpt(EYear, bankTpye ,"仟元;"+Unit);
                   }else if (YRpt[i][0].equals("AN009W")){       
                       errRptMsg = RptAN009W.createRpt(EYear, bankTpye ,"仟元;"+Unit);
                   }else if (YRpt[i][0].equals("AN010W")){       
                       errRptMsg = RptAN010W.createRpt(EYear, bankTpye ,"仟元;"+Unit);
                   }else if (YRpt[i][0].equals("AN011W")){       
                       errRptMsg = RptAN011W.createRpt(EYear, bankTpye ,"仟元;"+Unit);
                   }else if (YRpt[i][0].equals("AN012W")){ 
                       if(hsien_dbData.size() != 0){
                           for(int j=0;j<hsien_dbData.size();j++){
                               bean = (DataObject)hsien_dbData.get(j);                               
                               errRptMsg = RptAN012W.createRpt(EYear,bankTpye,(String)bean.getValue("hsien_id"),"仟元;"+Unit);
                               if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+YRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+YRpt[i][0]+"_"+(String)bean.getValue("hsien_id")+".xls");
                           }         
                       }
                   }else if (YRpt[i][0].equals("MC007W")){      
                       errRptMsg = RptFR052W.createRpt("0",EYear,"","");
                   }
    		       
    		       System.out.println(YRpt[i][0]+"create End");
    		       if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
    	           printRptMsg(logps,YRpt[i][0]+YRpt[i][1],errRptMsg);
    	           if(!YRpt[i][0].equals("AN012W")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+YRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+YRpt[i][0]+".xls");
    		    }  
    		}
               		
    		//產生季報資料=========================================================  
            for(int i=0;i<QRpt.length;i++){
                EYear = (!S_YEAR.equals("")?S_YEAR:Utility.getYear());
                //EYear="99";
                Month = (!S_MONTH.equals("")?Integer.parseInt(S_MONTH):Integer.parseInt(Utility.getMonth()));
                //Month=3;
                String Quarter="";
                if(Month == 3 || Month == 6 || Month == 9 || Month == 12){//每季更新為每年 4/7/10/1月的30號
                    System.out.println(QRpt[i][0]+"create begin");
                    session.setAttribute("Unit","1");  
                    session.setAttribute("S_YEAR",EYear);   
                    session.setAttribute("E_YEAR",EYear);
                    session.setAttribute("nowbank_type","ALL");   
                    session.setAttribute("excelaction","download");
                    session.setAttribute("BankList","ALL+全部");
                    session.setAttribute("printStyle","xls"); //109.09.10 add 
                    if(QRpt[i][0].equals("DS014W")){//轉換成季別 
                        session.setAttribute("S_MONTH",((String.valueOf(Month/3)).length()==1?"0":"")+String.valueOf(Month/3));
                        session.setAttribute("E_MONTH",((String.valueOf(Month/3)).length()==1?"0":"")+String.valueOf(Month/3));
                    }else{
                      session.setAttribute("S_MONTH",((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
                      session.setAttribute("E_MONTH",((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
                    }
                    if(QRpt[i][0].equals("DS014W")){ 
                        session.setAttribute("btnFieldList","WLX09_S_WARNING.warnaccount_tbal+警示帳戶(91.11.1~該季末日),WLX09_S_WARNING.warnaccount_refund_apply_amt+警示帳戶內剩餘款項之返還情形-申情退還,WLX09_S_WARNING.warnaccount_refund_amt+警示帳戶內剩餘款項之返還情形-已辦理退還");                        
                    }else if(QRpt[i][0].equals("DS015W")){ 
                        session.setAttribute("btnFieldList","WLX_S_RATE.period_1_fix_rate+定期存款一個月牌告利率,WLX_S_RATE.period_3_fix_rate+定期存款三個月牌告利率,WLX_S_RATE.period_6_fix_rate+定期存款六個月牌告利率,WLX_S_RATE.period_9_fix_rate+定期存款九個月牌告利率,WLX_S_RATE.period_12_fix_rate+定期存款一年牌告利率,WLX_S_RATE.basic_pay_var_rate+基本放款利率(機動),WLX_S_RATE.period_house_var_rate+指數型房貸指標利率,WLX_S_RATE.base_base_rate+基準利率");
                    }else if(QRpt[i][0].equals("DS021W")){                         
                        session.setAttribute("btnFieldList","001000+加權風險性資產-(1)信用風險,002000+加權風險性資產-(2)作業風險,003000+加權風險性資產-(3)市場風險,004000+加權風險性資產-(4)合計,005000+最低資本計提-(5)信用風險,006000+最低資本計提-(6)作業風險,007000+最低資本計提-(7)市場風險,008000+可用資本淨額-(8)第一類資本,009000+可用資本淨額-(9)第二類資本,010000+可用資本淨額-(10)第三類資本,011000+計算所需最低資本-信用風險-(11)第一類資本,012000+計算所需最低資本-信用風險-(12)第二類資本,013000+計算所需最低資本-作業風險-(13)第一類資本,014000+計算所需最低資本-作業風險-(14)第二類資本,015000+計算所需最低資本-市場風險-(15)第一類資本,016000+計算所需最低資本-市場風險-(16)第二類資本,017000+計算所需最低資本-市場風險-(17)第三類資本,018000+合格自有資本淨額-(18)第一類資本,019000+合格自有資本淨額-(19)第二類資本,020000+合格自有資本淨額-(20)第三類資本,021000+合格自有資本淨額-(21)合計,022000+不合格資本-(22)第二類,023000+不合格資本-(23)第三類,024000+合格自有資本與風險性資產比率,025000+第一類資本占風險性資產比率,026000+普通股占風險性資產比率");
                    }else if(QRpt[i][0].equals("DS022W")){     
                        session.setAttribute("btnFieldList","001000+第一類資本-資本普通股,002000+第一類資本-永續非累積特別股,003000+第一類資本-無到期日非累積次順位債券,004000+第一類資本-預收股本,005000+第一類資本-資本公積(固定資產增值公積除外),006000+第一類資本-法定盈餘公積,007000+第一類資本-特別盈餘公積,008000+第一類資本-累積盈餘,009000+第一類資本-少數股權,010000+第一類資本-股東權益其他項目(重估增值及備供出售金融資產未實現利益除外),011000+第一類資本-減:商譽,012000+第一類資本-減:累積虧損,013000+第一類資本-減:資本扣除項目,014000+第一類資本合計(A),015000+第二類資本-永續累積特別股,016000+第二類資本-無到期日累積次順位債券,017000+第二類資本-固定資產增值公積,018000+第二類資本-備供出售金融資產未實現利益之45%,019000+第二類資本-可轉換債券,020000+第二類資本-營業準備及備抵呆帳(不包括特定損失提列者),021000+第二類資本-長期次順位債券,022000+第二類資本-非永續特別股(發行期限五年以上者),023000+第二類資本-永續非累積特別股及無到期日非累積次順位債券合計超出第一類資本總額百分之十五者,024000+第二類資本-減:累積虧損,025000+第二類資本-減:資本扣除項目,026000+第二類資本合計(B),027000+第三類資本-短期次順位債券,028000+第三類資本-非永續特別股(發行期限二年以上者),029000+第三類資本合計(C),030000+自有資本合計(D)=(A)＋(B)＋(C)");
                    }else if(QRpt[i][0].equals("DS023W")){     
                        session.setAttribute("btnFieldList","001000+自第一類資本扣除金額-信用風險標準法,002000+自第一類資本扣除金額-信用風險內部評等法,003000+自第一類資本扣除金額-資產證券化,004000+自第一類資本扣除金額-市場風險,005000+自第一類資本扣除金額-合計,006000+自第二類資本扣除金額-信用風險標準法,007000+自第二類資本扣除金額-信用風險內部評等法,008000+自第二類資本扣除金額-資產證券化,009000+自第二類資本扣除金額-市場風險,010000+自第二類資本扣除金額-合計");
                    }else if(QRpt[i][0].equals("DS038W")){    
                        session.setAttribute("btnFieldList","000000+年度別,001000+利息收入(1),002000+利息費用(2),003000+利息淨收益(3)=(1)-(2),004000+手續費淨收益(4),005000+公平價值變動列入損益之金融資產及負債損益(5),006000+採權益法認列之投資損益(其中投資處分損益不納入)(6),007000+兌換損益(7),008000+其他非利息淨損益(8),009000+利息以外淨收益(9)=(4)＋(5)＋(6)＋(7)＋(8),010000+營業毛利(10)=(3)＋(9),011000+作業風險應計提資本");
                    }else if(QRpt[i][0].equals("DS051W")){         
                        session.setAttribute("btnFieldList","110100+自第一類資本扣除-評價準備提列不足數,110200+自第一類資本扣除-利率風險之資本扣除金額,110300+自第一類資本扣除-權益證券風險之資本扣除項目,111000+自第一類資本扣除-合計,120100+自第二類資本扣除-評價準備提列不足數,120200+自第二類資本扣除-利率風險之資本扣除金額,120300+自第二類資本扣除-權益證券風險之資本扣除項目,121000+自第二類資本扣除-合計");              
                    }

                    URL myURL=new URL(URLStr+QRpt[i][0]+"_Excel.jsp?act=download");
                    URLConnection raoURL;
                    raoURL=myURL.openConnection();
                    raoURL.setRequestProperty("Cookie", "JSESSIONID=" + URLEncoder.encode(session.getId(), "UTF-8"));                        
                    raoURL.setDoInput(true);
                    raoURL.setDoOutput(true);
                    raoURL.setUseCaches(false);
                    
                    BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
                    String raoInputString;            
                    while((raoInputString=infromURL.readLine())!=null){
                          //System.out.println(raoInputString);
                    }
                    infromURL.close();
                    
                    System.out.println(QRpt[i][0]+"create End");
                    if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
                    printRptMsg(logps,QRpt[i][0]+QRpt[i][1],errRptMsg);
                    copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+QRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+QRpt[i][0]+".xls");
                }
            }
           
            
            //產生月報資料(更新顏率為每月30號)=========================================================
            String bank_no_list = "";//6060019+6060019新北市三重區農會信用部,
            sqlCmd.delete(0,sqlCmd.length());
            paramList = new ArrayList();
            sqlCmd.append("select bank_no,bank_name from bn01 where bank_type in (?) and bn_type <> '2' and bank_no != '9999999' and m_year=100 order by bank_no");
            paramList.add("6");
            bank_no_6_dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");//農會
            paramList = new ArrayList();
            paramList.add("7");
            bank_no_7_dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");//漁會
            
            sqlCmd.delete(0,sqlCmd.length());
            paramList = new ArrayList();
            sqlCmd.append("select bank_no,bank_name from bn01 where bank_type in (?,?) and bn_type <> '2' and bank_no != '9999999' and m_year=100 order by bank_no");
            paramList.add("6");
            paramList.add("7");
            bank_no_67_dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");//農漁會
            //動態報表需用的已選取的banklist
            if(bank_no_6_dbData.size() != 0){//所選取的bank_no
                for(int j=0;j<bank_no_6_dbData.size();j++){
                    bean = (DataObject)bank_no_6_dbData.get(j);   
                    bank_no_list += (bank_no_list.length() > 0?",":"")+(String)bean.getValue("bank_no")+"+"+(String)bean.getValue("bank_name");                        
                }         
            }
            
            EYear = (!S_YEAR.equals("")?S_YEAR:Utility.getYear());
            //EYear="100";
            Month = (!S_MONTH.equals("")?Integer.parseInt(S_MONTH):Integer.parseInt(Utility.getMonth()));
            //Month=1;
            session.setAttribute("Unit","1000");  
            session.setAttribute("S_YEAR",EYear);   
            session.setAttribute("E_YEAR",EYear);
            session.setAttribute("S_MONTH",((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
            session.setAttribute("E_MONTH",((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
            session.setAttribute("nowbank_type","ALL");  
            session.setAttribute("bank_type","ALL");
            session.setAttribute("CANCEL_NO","N");
            session.setAttribute("excelaction","download");
            session.setAttribute("BankList","ALL+全部");
            session.setAttribute("printStyle","xls"); //109.09.10 add           
            for(int i=0;i<MRpt.length;i++){          
                System.out.println(MRpt[i][0]+"create begin");
                if (MRpt[i][0].equals("FR001W")){       
                    errRptMsg = RptFR001W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6","0");
                }else if (MRpt[i][0].equals("FR001WA")){    
                    errRptMsg = RptFR001W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6","0");
                }else if (MRpt[i][0].equals("FR001WE")){      
                    errRptMsg = RptFR001WE.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6",null);
                }else if (MRpt[i][0].equals("FR001WN")){          
                    errRptMsg = RptFR001WN.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"ALL","1");
                }else if (MRpt[i][0].equals("FR002W")){
                    errRptMsg = RptFR002W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6");
                }else if (MRpt[i][0].equals("FR002WA") || MRpt[i][0].equals("BR007W")){    
                    session.setAttribute("nowbank_type","6"); 
                    URL myURL = null;
                    if(hsien_dbData.size() != 0){
                        for(int j=0;j<hsien_dbData.size();j++){
                            bean = (DataObject)hsien_dbData.get(j);      
                            if (MRpt[i][0].equals("FR002WA")){
                               errRptMsg = FR002WA_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6",(String)bean.getValue("hsien_id"));
                               if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("hsien_id")+".xls");
                            }else if (MRpt[i][0].equals("BR007W")){                                 
                               myURL= new URL(URLStr+"BR007W_Excel.jsp?act=download&hsien_id="+(String)bean.getValue("hsien_id")+"&cancel_no=N");
                                                
                               URLConnection raoURL;
                               raoURL=myURL.openConnection();
                               raoURL.setRequestProperty("Cookie", "JSESSIONID=" + URLEncoder.encode(session.getId(), "UTF-8"));                        
                               raoURL.setDoInput(true);
                               raoURL.setDoOutput(true);
                               raoURL.setUseCaches(false);                                             
                               BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
                               String raoInputString;            
                               while((raoInputString=infromURL.readLine())!=null){
                                      //System.out.println(raoInputString);
                               }
                               infromURL.close(); 
                               copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("hsien_id")+".xls");
                            }                           
                        }         
                    }
                }else if (MRpt[i][0].equals("FR003W") || MRpt[i][0].equals("FR004W") || MRpt[i][0].equals("FR004W_month") || MRpt[i][0].equals("FR0066W")
                       || MRpt[i][0].equals("FR026W") || MRpt[i][0].equals("FR068W")
                        ){
                    if(bank_no_6_dbData.size() != 0){//農會
                        for(int j=0;j<bank_no_6_dbData.size();j++){
                            bean = (DataObject)bank_no_6_dbData.get(j);
                            if (MRpt[i][0].equals("FR003W")){
                                errRptMsg = FR003W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit); 
                            }else if (MRpt[i][0].equals("FR004W")){          
                                errRptMsg = FR004W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit,"false");
                            }else if (MRpt[i][0].equals("FR004W_month")){
                                errRptMsg = FR004W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit,"true");
                            }else if (MRpt[i][0].equals("FR0066W")){
                                errRptMsg = RptFR0066W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),"1", "6","1",false,"");
                                  }else if ( MRpt[i][0].equals("FR026W")){
                                errRptMsg = FR026W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6",(String)bean.getValue("bank_no"),Unit);
                            }else if ( MRpt[i][0].equals("FR068W")){      
                                errRptMsg = FR068W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),Unit,"6","BOAF000001");
                            }   
                            System.out.println(MRpt[i][0]+".errRptMsg="+errRptMsg);
                            if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+(MRpt[i][0].equals("FR0066W")?".rtf":".xls"));
                            System.out.println(Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+(MRpt[i][0].equals("FR0066W")?".rtf":".xls")+":copyOK="+copyOK);
                        }    
                    } 
                    if(bank_no_7_dbData.size() != 0){//漁會
                        for(int j=0;j<bank_no_7_dbData.size();j++){
                            bean = (DataObject)bank_no_7_dbData.get(j);
                            if (MRpt[i][0].equals("FR003W")){
                               errRptMsg = FR003F_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit);
                               if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+"漁會信用部資產負債表_10301.xls",Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");
                            }else if (MRpt[i][0].equals("FR004W")){ 
                               errRptMsg = FR004F_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit,"false");
                               if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+"漁會信用部損益表_10301.xls",Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");
                            }else if (MRpt[i][0].equals("FR004W_month")){
                               errRptMsg = FR004F_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),(String)bean.getValue("bank_name"),Unit,"true");
                               if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+"漁會信用部損益表_10301.xls",Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");                                
                            }else if (MRpt[i][0].equals("FR0066W")){
                                errRptMsg = RptFR0066W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),"1", "7","1",false,"");
                                if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+"漁會各項法定比率表.rtf",Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".rtf");
                            }else if ( MRpt[i][0].equals("FR026W")){    
                                errRptMsg = FR026W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"7",(String)bean.getValue("bank_no"),Unit);
                                if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+"在台無住所之外國人新台幣存款表-漁會.xls",Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");
                            }else if ( MRpt[i][0].equals("FR068W")){    
                                errRptMsg = FR068W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"),Unit,"7","BOAF000001");
                                if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");
                            }
                        }    
                    }  
                }else if (MRpt[i][0].equals("FR005W")){    
                    errRptMsg = RptFR005W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6","0","","");                
                }else if (MRpt[i][0].equals("FR006W")){         
                    errRptMsg = FR006W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"ALL");
                }else if (MRpt[i][0].equals("FR007W")){     
                    errRptMsg = FR007W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6","1",Unit);                    
                }else if (MRpt[i][0].equals("FR007WA")){    
                    errRptMsg = RptFR007WA.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6");
                }else if (MRpt[i][0].equals("FR008W") || MRpt[i][0].equals("FR045W")){  
                      if(bank_no_67_dbData.size() != 0){//所選取的bank_no
                         for(int j=0;j<bank_no_67_dbData.size();j++){
                              bean = (DataObject)bank_no_67_dbData.get(j);   
                              if (MRpt[i][0].equals("FR008W")){ 
                                  errRptMsg = FR008W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month), (String)bean.getValue("bank_no"),(String)bean.getValue("bank_no")+(String)bean.getValue("bank_name"),"6",Unit);
                              }else if (MRpt[i][0].equals("FR045W")){ 
                                  errRptMsg = RptFR045W_bank_67.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),(String)bean.getValue("bank_no"));
                              }
                              if(errRptMsg.equals("")) copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");                  
                         }    
                      }     
                }else if ( MRpt[i][0].equals("FR009W")){
                    if(bank_no_67_dbData.size() != 0){//所選取的bank_no
                        for(int j=0;j<bank_no_67_dbData.size();j++){
                             bean = (DataObject)bank_no_67_dbData.get(j); 
                             URL myURL= new URL(URLStr+MRpt[i][0]+"_Excel.jsp?act=download&BANK_NO="+(String)bean.getValue("bank_no")+"/"+(String)bean.getValue("bank_no")+"&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&Unit="+Unit+"&coaflag=true");//測試環境
                             URLConnection raoURL;
                             raoURL=myURL.openConnection();
                             raoURL.setRequestProperty("Cookie", "JSESSIONID=" + URLEncoder.encode(session.getId(), "UTF-8"));                        
                             raoURL.setDoInput(true);
                             raoURL.setDoOutput(true);
                             raoURL.setUseCaches(false);
                             
                             BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
                             String raoInputString;            
                             while((raoInputString=infromURL.readLine())!=null){
                                   //System.out.println(raoInputString);
                             }
                             infromURL.close();
                             }
                             copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+"_"+(String)bean.getValue("bank_no")+".xls");
                     } 
                }else if ( MRpt[i][0].equals("FR023W")){
                    errRptMsg = RptFR023W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6");
                }else if ( MRpt[i][0].equals("FR023W")){    
                    errRptMsg = RptFR024W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6");
                }else if ( MRpt[i][0].equals("FR025W")){        
                    errRptMsg = FR025W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6",Unit);                                
                }else if ( MRpt[i][0].equals("FR028W")){        
                    errRptMsg = RptFR028W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6",Unit);
                }else if ( MRpt[i][0].equals("FR029W")){          
                    errRptMsg = RptFR029W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"1","6");                    
                }else if ( MRpt[i][0].equals("FR030WW_2")){     
                    errRptMsg = FR030WB_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6");
                }else if ( MRpt[i][0].equals("FR030WW_3")){     
                    errRptMsg = FR030W_Excel.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6",Unit);
                }else if ( MRpt[i][0].equals("FR031W")){    
                    errRptMsg = RptFR031W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"1");
                }else if ( MRpt[i][0].equals("FR032W")){ 
                    errRptMsg = RptFR032W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"1");
                }else if ( MRpt[i][0].equals("FR034WW")){  
                    errRptMsg = RptFR034WA.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6");
                }else if ( MRpt[i][0].equals("FR030WW_1") || MRpt[i][0].equals("FR036WW_1") || MRpt[i][0].equals("FR036WW_2") || MRpt[i][0].equals("FR036WW_3")){  
                    if ( !MRpt[i][0].equals("FR030WW_1")) session.setAttribute("nowbank_type","6");
                    session.setAttribute("BankList",bank_no_list);
                    URL myURL= null;
                    if ( MRpt[i][0].equals("FR030WW_1")){
                        myURL= new URL(URLStr+"FR030WA_Excel.jsp?act=download&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&bank_type=6");                   
                    }else if ( MRpt[i][0].equals("FR036WW_1")){
                        myURL = new URL(URLStr+"FR036WB_Excel.jsp?act=download&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&E_YEAR="+EYear+"&E_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&Unit="+Unit+"&rptStyle=0");
                    }else if ( MRpt[i][0].equals("FR036WW_2")){
                        myURL = new URL(URLStr+"FR036WA_Excel.jsp?act=download&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&Unit="+Unit+"&rptStyle=0");
                    }else if ( MRpt[i][0].equals("FR036WW_3")){
                        myURL = new URL(URLStr+"FR036W_Excel.jsp?act=download&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month)+"&Unit="+Unit+"&rptStyle=0");
                    }
                    URLConnection raoURL;
                    raoURL=myURL.openConnection();
                    raoURL.setRequestProperty("Cookie", "JSESSIONID=" + URLEncoder.encode(session.getId(), "UTF-8"));                        
                    raoURL.setDoInput(true);
                    raoURL.setDoOutput(true);
                    raoURL.setUseCaches(false);
                    
                    BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
                    String raoInputString;            
                    while((raoInputString=infromURL.readLine())!=null){
                          //System.out.println(raoInputString);
                    }
                    infromURL.close();  
                }else if ( MRpt[i][0].equals("FR037W")){         
                    errRptMsg = RptFR037W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"ALL","6","",Unit);
                }else if ( MRpt[i][0].equals("FR039W")){
                    errRptMsg = RptFR039W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6");
                }else if ( MRpt[i][0].equals("FR040WW")){
                    errRptMsg = RptFR040WW.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"6","0");                  
                }else if ( MRpt[i][0].equals("FR046W")){      
                    errRptMsg = RptFR046W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,bank_type,"0");
                }else if ( MRpt[i][0].equals("FR054W")){      
                    errRptMsg = RptFR054W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"",null);
                }else if ( MRpt[i][0].equals("FR056W")){
                    errRptMsg = RptFR056W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6");
                }else if ( MRpt[i][0].equals("FR057W")){
                    errRptMsg = RptFR057W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"6");
                }else if ( MRpt[i][0].equals("FR058W")){
                    errRptMsg = RptFR058W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
                }else if ( MRpt[i][0].equals("FR062W")){    
                    errRptMsg = RptFR062W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"");
                }else if ( MRpt[i][0].equals("FR064W")){    
                    errRptMsg = RptFR064W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),Unit,"","",null);
                }else if ( MRpt[i][0].equals("FR065W")){        
                    errRptMsg = RptFR065W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"1","2","10","15","30","ALL");
                }else if ( MRpt[i][0].equals("FR067W")){     
                    errRptMsg = RptFR067W.createRpt(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"1","0","6","8");                
                }else if ( MRpt[i][0].equals("MC004W")){
                    errRptMsg = RptFR049W.createRpt("ALL",Utility.getFullDate(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"01"),Utility.getFullDate(EYear,((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month),"31"),"100","");
                }else if(MRpt[i][0].equals("BR005W")){                    
                    session.setAttribute("btnFieldList","BN02.TBANK_NO+總機構代碼,BN01.BANK_NAME+所屬總機構名稱,BN02.BANK_NO+分支機構配賦代號,BN02.BANK_NAME+分支機構名稱,WLX02.CONST_TYPE+分支機構類別,WLX02.SETUP_NO+原始核准設立文號,WLX02.SETUP_NO_DATE+原始核准設立文號日期,WLX02.SETUP_DATE+原始開業日期,WLX02.CHG_LICENSE_DATE+最近換照日期,WLX02.CHG_LICENSE_NO+最近換照文號,WLX02.CHG_LICENSE_REASON+最近換照事由,WLX02.START_DATE+開始營業日,CD01.HSIEN_NAME+縣市別,CD02.AREA_NAME+鄉鎮,WLX02.AREA_ID+郵遞區號,WLX02.ADDR+地址,WLX02.TELNO+電話,WLX02.FAX+傳真,ba01.m2_name_bank_name+所屬之地方主管機關,ba01_1.center_no_name+所屬之資訊共用中心,WLX02.wlx02_m_name+分部主任,WLX02.WEB_SITE+網址,WLX02.EMAIL+e-mail帳號,WLX02.STAFF_NUM+分支機構員工數");                    
                }else if(MRpt[i][0].equals("BR010W")){
                    session.setAttribute("btnFieldList","BN01.BANK_NAME+公司,WLX01_m.firstname+名字,WLX01_m.lastname+姓氏,WLX01.ADDR+商務-街,WLX01.FAX+商務傳真,WLX01.TELNO+商務電話,WLX01.EMAIL+電子郵件地址");                   
                }else if(MRpt[i][0].equals("BR022W")){    
                    errRptMsg = RptBR022W.createRpt("1","BOAF000001","");
                }else if(MRpt[i][0].equals("DS001W")){
                    session.setAttribute("btnFieldList","count_seq+單位數,field_debit+存款總額,field_credit+放款總額,field_dc_rate+存放比率,field_120700+內部融資餘額,field_over+逾放金額,field_over_rate+逾放比率,field_320300+本期損益,field_transfer+轉存全國農業金庫,field_transfer_rate+轉存全國農業金庫比率,field_310000+事業資金及公積,field_net+淨值,field_fixnet_rate+固定資產佔淨值比,field_check_rate+活存比率,field_150200+催收款項,field_backup+備抵呆帳,field_noassure+無擔保放款,field_captial_rate+合格淨值佔風險性資產比率");                    
                }else if(MRpt[i][0].equals("FR025WA") || MRpt[i][0].equals("DS002W") || MRpt[i][0].equals("DS003W") || MRpt[i][0].equals("DS004W") 
                      || MRpt[i][0].equals("DS005W")  || MRpt[i][0].equals("DS006W") || MRpt[i][0].equals("DS007W") || MRpt[i][0].equals("DS009W") 
                      || MRpt[i][0].equals("DS011W")  || MRpt[i][0].equals("DS012W") || MRpt[i][0].equals("DS017W") || MRpt[i][0].equals("DS018W")
                      ){    
                    session.setAttribute("bank_type","6");
                    session.setAttribute("nowbank_type","6");  
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("printStyle","xls"); //109.09.10 add 
                    if(MRpt[i][0].equals("DS002W")){
                        session.setAttribute("btnFieldList","A02_1+1.月底存放比率,A02_2+2.農會信用部對農會經濟事業門融通資金之限制,A02_3+3.非會員存款之額度限制-104.02.13取消(三)限制,A02_4+4.贊助會員授信總額占贊助會員存款總額之比率,A02_5+5.辦理非會員無擔保消費性貸款之限制,A02_6+6.非會員授信總額占非會員存款總額之比率,A02_7+7.辦理自用住宅放款限額,A02_8+8.固定資產淨額限制,A02_8_a+固定資產淨額限制不在此限的原因項目,A02_9+9.外幣風險之限制,A02_10+10.對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者為擔保授信限制,A02_11+11.對每一會員（含同戶家屬）及同一關係人放款最高限額,A02_12+12.對每一贊助會員及同一關係人之授信限額,A02_13+13.對同一非會員及同一關係人之授信限額,A02_14+14.對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額,990110+月平均放款總額,990140+月平均存款總額,990310+非會員存款總額,990510+非會員無擔保消費性貸款,990610+非會員授信總額,990611+對政府部門授信金額,990610-990611+非會員放款扣除對政府部門授信金額,991010+對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者擔保放款最高限額,99141Y+最近決算年度");                   
                    }else if(MRpt[i][0].equals("DS003W")){                        
                        session.setAttribute("btnFieldList","84061P+統一農貸實質平均利率,84068P+內部融資利率,84062P+活期存款實質平均利率,84063P+活期儲蓄存款實質平均利率,84064P+員工活期儲蓄存款實質平均利率,84065P+定期存款實質平均利率,84066P+定期儲蓄存款實質平均利率,84067P+員工定期儲蓄存款實質平均利率,84069P+綜合平均存款利率,420110+一般放款利息收入,420120+貼現利息收入,420130+透支利息收入,420150+專案放款利息收入,420160+農業發展基金利息收入,520100+存款利息支出,520110+活期存款利息,520120+定期存款利息,520130+活期儲蓄存款利息,520140+定期儲蓄存款利息,840640+員工活期儲蓄存款利息,840670+員工定期儲蓄存款利息,field_1200_1502_1503+放款,field_520100_220000+存款-平均利率,field_220000+存款-存款總額");
                    }else if(MRpt[i][0].equals("DS004W")){
                        session.setAttribute("btnFieldList","field_over_a01+A01.逾期放款[990000],field_credit_a01+A01.總放款(放款含催收),field_over_rate_a01+A01.逾放比率,field_over_a04+1.逾期放款[840740],field_credit_a04+2.總放款(放款含催收)[840750],field_over_rate_a04+3.逾放比率,field_840760_a04+4.應予觀察放款金額[840760],field_rate_840760_a04+5.應予觀察放款占總放款比率,field_840760_over_a04+6.逾期放款及應予觀察放款占總放款比率,field_840710_a04_a+放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者840710(A),field_840720_a04_b+中長期分期償債放款，未按期攤還超過三個月至六個月者840720(B),field_a04+A04_逾期巳達列報逾放條件而准免列報者(C)");                    
                    }else if(MRpt[i][0].equals("DS005W")){
                       session.setAttribute("btnFieldList","920101+1.現金,920201+2.對本國中央政府及中央銀行之債權或經其保證之債權,920301+3.以現金、在本會之存款、中央政府或中央銀行債券為擔保之債權,920401+4.對本國中央政府以外各級政府之債權或其保證之債權,920501+5.對本國銀行及其保證之債權,920601+6.住宅用不動產擔保放款上列債權之應收利息,9207X0+7.上列以外規定.風險權數未達100%之資產,920801+8.上列以外之債權及其他資產(以扣除累計折舊後之淨額計算,929901+信用部淨值占風險性資產比率計算表中合計的帳面金額,910101+(1)事業資金,910102+(2)事業公積,910103+(3)法定公積,910104+(4)特別公積,910105+(5)捐贈公積,910106+(6)資產公積,910107+(7)統一農貸公積,910108+(8)累積盈虧(加計「上期損益」),910109+減：備抵呆帳、損失準備及營業準備提列不足之金額,910110+(9)本期損益,910199+第一類資本合計(A),910201+(1)固定資產增值公積,910202+(2)備抵呆帳、損失準備及營業準備,910203+減：特定損失所提列備抵呆帳、損失準備及營業準備,910299+第二類資本合計(B),910300+淨值總額(C)=(A),910401+全國農業金庫股票(D),910402+財金資訊股份有限公司股票(E),910403+合作金庫銀行股票(F),910404+上列以外之聯營出資(F1),910400+合格淨值(G)=(C)-[(D),910500+風險性資產總額(H),91060P+信用部淨值占風險性資產比率(資本適足率)");
                    }else if(MRpt[i][0].equals("DS006W")){    
                       session.setAttribute("btnFieldList","120101+一般放款(無擔保),120102+一般放款(擔保),120301+透支(無擔保),120302+透支(擔保),120700+內部融資(透支),120401+統一農貸(無擔保),120402+統一農貸(擔保),120501+專案放款(無擔保),120502+專案放款(擔保),120601+農發基金-農建放款,120602+農發基金-農機放款,120603+農發基金-購地放款,120604+農發基金-農宅放款,120600+農發基金放款小計,150200+催收款項,970000+合計");                    
                    }else if(MRpt[i][0].equals("DS007W")){    
                       session.setAttribute("btnFieldList","F01_A_1+活期存款-個人,F01_A_2+活期存款-公司、行號、團體,F01_A_3+活期存款-外國專業投資機構,F01_A_4+活期存款-外國銀行,F01_A_5+活期存款-小計,F01_B_1+活期儲蓄存款-個人,F01_B_2+活期儲蓄存款-公司、行號、團體,F01_B_3+活期儲蓄存款-外國專業投資機構,F01_B_4+活期儲蓄存款-外國銀行,F01_B_5+活期儲蓄存款-小計,F01_C_1+定期存款-個人,F01_C_2+定期存款-公司、行號、團體,F01_C_3+定期存款-外國專業投資機構,F01_C_4+定期存款-外國銀行,F01_C_5+定期存款-小計,F01_D_1+支票存款-個人,F01_D_2+支票存款-公司、行號、團體,F01_D_3+支票存款-外國專業投資機構,F01_D_4+支票存款-外國銀行,F01_D_5+支票存款-小計,F01_E_1+合計-個人,F01_E_2+合計-公司、行號、團體,F01_E_3+合計-外國專業投資機構,F01_E_4+合計-外國銀行,F01_E_5+合計-小計");
                    }else if(MRpt[i][0].equals("DS009W")){    
                       session.setAttribute("btnFieldList","WLX05_ATM_SETUP.SITE_ADDR+裝設地點,WLX05_ATM_SETUP.SETUP_DATE+裝設日期,WLX05_ATM_SETUP.CANCEL_DATE+遷移/裁撤日期,WLX05_ATM_SETUP.PROPERTY_NO+財產編號,WLX05_ATM_SETUP.MACHINE_NAME+機器品名,WLX05_ATM_SETUP.COMMENT_M+備註");
                    }else if(MRpt[i][0].equals("DS011W")){        
                       session.setAttribute("btnFieldList","WLX07_M_CHECKBANK.checkbank_bal_tot+總計,WLX07_M_CHECKBANK.checkbank_bal+正會員,WLX07_M_CHECKBANK.checkbank_bal_s+贊助會員,WLX07_M_CHECKBANK.checkbank_bal_n+非會員");
                    }else if(MRpt[i][0].equals("DS012W")){        
                      session.setAttribute("btnFieldList","WLX07_M_CREDIT.creditmonth_amt+本月新增貸放資料,WLX07_M_CREDIT.credityear_amt_acc+本年累計貸放資料,WLX07_M_CREDIT.credit_bal+貸放餘額,WLX07_M_CREDIT.overcreditmonth_amt+本月新增逾放資料,WLX07_M_CREDIT.overcredit_bal+逾放餘額");
                    }else if(MRpt[i][0].equals("DS017W")){
                      session.setAttribute("btnFieldList","warnaccount_cnt+「警示帳戶」戶數(A),limitaccount_cnt+「衍生管制帳戶」戶數(B),erroraccount_cnt+自行篩選有異常並已採資金流出管制措施之存款帳戶戶數(C),otheraccount_cnt+其他帳戶戶數(D),depositaccount_tcnt+存款帳戶總戶數");
                    }else if(MRpt[i][0].equals("DS018W")){ 
                      session.setAttribute("btnFieldList","over_cnt+剩餘件數,over_amt+剩餘金額(A),push_over_amt+剩餘金額-催收款(B),totalamt+全會放出總金額(C),push_totalamt+全會放出總金額-催收款(D),over_total_rate+佔放款總額的比率(A-B)/(C-D)");                                                           
                    }
                }else if(MRpt[i][0].equals("DS053W")){
                    session.setAttribute("btnFieldList","GUARANTEE_CNT1+政策性農業專案貸款-保證件數,GUARANTEE_CNT2+其他政策性貸款-保證件數,GUARANTEE_CNT3+一般農業貸款-保證件數,GUARANTEE_CNT4+統一農貸-保證件數,GUARANTEE_CNT5+統一漁貸-保證件數,GUARANTEE_CNT6+農漁會一般農業貸款-保證件數,GUARANTEE_CNT0+保證件數-合計,LOAN_AMT1+政策性農業專案貸款-貸款金額,LOAN_AMT2+其他政策性貸款-貸款金額,LOAN_AMT3+一般農業貸款-貸款金額,LOAN_AMT4+統一農貸-貸款金額,LOAN_AMT5+統一漁貸-貸款金額,LOAN_AMT6+農漁會一般農業貸款-貸款金額,LOAN_AMT0+貸款金額-合計,GUARANTEE_AMT1+政策性農業專案貸款-保證金額,GUARANTEE_AMT2+其他政策性貸款-保證金額,GUARANTEE_AMT3+一般農業貸款-保證金額,GUARANTEE_AMT4+統一農貸-保證金額,GUARANTEE_AMT5+統一漁貸-保證金額,GUARANTEE_AMT6+農漁會一般農業貸款-保證金額,GUARANTEE_AMT0+保證金額-合計,GUARANTEE_BAL1+政策性農業專案貸款-保證餘額,GUARANTEE_BAL2+其他政策性貸款-保證餘額,GUARANTEE_BAL3+一般農業貸款-保證餘額,GUARANTEE_BAL4+統一農貸-保證餘額,GUARANTEE_BAL5+統一漁貸-保證餘額,GUARANTEE_BAL6+農漁會一般農業貸款-保證餘額,GUARANTEE_BAL0+保證餘額-合計");
                }else if(MRpt[i][0].equals("DS054W")){     
                    session.setAttribute("btnFieldList","GUARANTEE_CNT_YEAR+本年度保證件數,GUARANTEE_AMT_YEAR+本年度保證金額,LOAN_AMT_YEAR+本年度融資金額,GUARANTEE_BAL_YEAR+本年度保證餘額,GUARANTEE_CNT_SUM+總累計保證件數,GUARANTEE_AMT_SUM+總累計保證金額,LOAN_AMT_SUM+總累計融資金額,GUARANTEE_CNT_YEAR0+本年度保證件數-合計,GUARANTEE_AMT_YEAR0+本年度保證金額-合計,LOAN_AMT_YEAR0+本年度融資金額-合計,GUARANTEE_BAL_YEAR0+本年度保證餘額-合計,GUARANTEE_CNT_SUM0+總累計保證件數-合計,GUARANTEE_AMT_SUM0+總累計保證金額-合計,LOAN_AMT_SUM0+總累計融資金額-合計");
                }else if(MRpt[i][0].equals("DS055W")){
                    session.setAttribute("S_YEAR1","1");
                    session.setAttribute("btnFieldList","field_debit+存款總額,field_credit+放款總額,field_dc_rate+存放比率,field_120700+內部融資餘額,field_320300+本期損益,field_320000+盈虧及損益,field_transfer+轉存全國農業金庫,field_transfer_rate+轉存全國農業金庫比率,field_310000+事業資金及公積,field_net+淨值,field_fixnet_rate+固定資產佔淨值比,field_check_rate+活存比率,field_150200+催收款項,field_noassure+無擔保放款,field_990611+對直轄市、縣市府、離保地區鄉公所辦理之授信總額,field_captial_rate +淨值佔風險性資產比率BIS,field_over+狹義逾期放款,field_840740+廣義逾期放款,field_over_rate+狹義逾放比率,field_over_rate_max+狹義逾放比最高,field_over_rate_avg+狹義逾放比平均,field_840740_rate+廣義逾放比率,field_840740_rate_max+廣義逾放比最高,field_840740_rate_avg+廣義逾放比平均,field_backup+備抵呆帳金額,field_backup_over_rate+備抵呆帳占狹義逾期放款比率,field_backup_840740_rate+備抵呆帳占廣義逾期放款比率,field_840760+應予觀察放款金額,field_840760_rate+應予觀察放款金額佔同日放款總餘額比率");
                }else if(MRpt[i][0].equals("DS060W")){
                    session.setAttribute("loan_item","ALL");
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("btnFieldList","field_over6m_loan_bal_cnt+戶數,field_over6m_loan_bal_amt+逾期超過六個月放款餘額,field_over_rate+逾放比率");
                }else if(MRpt[i][0].equals("DS061W")){//107.02.23 add
                    session.setAttribute("loan_item","ALL");
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("btnFieldList","field_delay_loan_cnt+當月-件數(A),field_delay_loan_amt+當月-目前累積餘額,field_delay_loan_cnt_last+上月-件數(B),field_delay_loan_amt_last+上月-目前累積餘額,field_delay_loan_cnt_diff+展延案件數比較(A)-(B),field_delay_loan_cnt_rate+展延案件數增加幅度%(A-B)/B");
                }else if(MRpt[i][0].equals("DS062W")){//107.02.26 add
                    session.setAttribute("loan_item","ALL");
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("btnFieldList","field_delay_loan_cnt+本月份核准延期還款件數(A),field_delay_loan_cnt_sum+當年度累計核准延期還款件數(B),field_loan_bal_cnt+截至本月底尚有餘額專案農貸總件數(C),field_loan_bal_cnt_rate+占尚有餘額專案農貸總件數比率(B)/(C)");
                }else if(MRpt[i][0].equals("DS063W")){//107.02.26 add
                    session.setAttribute("loan_item","08");
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("btnFieldList","field_over3p_loan_amt+同戶三人以上申貸之貸款餘額(A),field_over3p_loan_cnt+家戶數(地址),field_loan_bal_amt+各農漁會貸放餘額(B),field_over3p_loan_rate+與該機構餘額比(A/B),field_over3p_sum_rate+與總餘額比");              
                }else if(MRpt[i][0].equals("DS064W")){//107.02.26 add
                    session.setAttribute("loan_item","08");
                    session.setAttribute("BankList",bank_no_list);
                    session.setAttribute("btnFieldList","field_again_loan_cnt+借款戶一借再借-案件數(A),field_again_loan_bal_amt+借款戶一借再借-貸款餘額,field_again_loan_bal_amt_1th+借款戶一借再借-第一筆農綜貸餘額總計,field_loan_bal_cnt+各農漁會-餘額件數(B),field_loan_bal_amt+各農漁會-貸款餘額,field_loan_bal_cnt_rate+餘額件數比率(A/B),field_loan_bal_amt_rate+與該機構餘額比,field_loan_bal_amt_sum_rate+與總餘額比");             
                } 
                if (MRpt[i][0].startsWith("BR") || MRpt[i][0].startsWith("DS") || MRpt[i][0].equals("FR025WA")){  
                    //URL myURL=new URL(Utility.getProperties("BOAFWebSite")+"DS014W_Excel.jsp?act=download");                   
                    URL myURL= null;
                    if(MRpt[i][0].equals("FR025WA")){
                       myURL= new URL(URLStr+MRpt[i][0]+"_Excel.jsp?act=download&S_YEAR="+EYear+"&S_MONTH="+((String.valueOf(Month)).length()==1?"0":"")+String.valueOf(Month));
                    }else{
                       myURL= new URL(URLStr+MRpt[i][0]+"_Excel.jsp?act=download");
                    }
                
                    URLConnection raoURL;
                    raoURL=myURL.openConnection();
                    raoURL.setRequestProperty("Cookie", "JSESSIONID=" + URLEncoder.encode(session.getId(), "UTF-8"));                        
                    raoURL.setDoInput(true);
                    raoURL.setDoOutput(true);
                    raoURL.setUseCaches(false);
                
                    BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
                    String raoInputString;            
                    while((raoInputString=infromURL.readLine())!=null){
                      //System.out.println(raoInputString);
                    }
                    infromURL.close();
                }
                System.out.println(MRpt[i][0]+".errRptMsg="+errRptMsg);
                System.out.println(MRpt[i][0]+"create End");
                if(errRptMsg.equals("")) errRptMsg = "報表產生完成";
                printRptMsg(logps,MRpt[i][0]+MRpt[i][1],errRptMsg);
                //已區分機構或縣市別的報表檔案不用再搬檔
                if (!MRpt[i][0].equals("FR002WA") && !MRpt[i][0].equals("BR007W") && !MRpt[i][0].equals("FR003W") 
                 && !MRpt[i][0].equals("FR004W") && !MRpt[i][0].equals("FR004W_month") && !MRpt[i][0].equals("FR0066W")  
                 && !MRpt[i][0].equals("FR008W") && !MRpt[i][0].equals("FR009W") && !MRpt[i][0].equals("FR026W") 
                 && !MRpt[i][0].equals("FR045W") && !MRpt[i][0].equals("FR068W")
                ){  
                    copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+MRpt[i][1],Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+".xls");
                    System.out.println(Utility.getProperties("coaxlsDir")+System.getProperty("file.separator")+MRpt[i][0]+".xls"+":copyOK="+copyOK);
                }
            }
                        
	       Utility.printLogTime("RptGenerateCOA end time");	
		}catch(Exception e){
			System.out.println("RptGenerateCOA.createRpt Error:"+e+e.getMessage());			
		}
		return errMsg;
	}
    
    
    public static void printRptMsg(PrintStream logps,String rptKind,String errRptMsg){
    	if(!errRptMsg.equals("")){
	       logcalendar = Calendar.getInstance(); 
		   nowlog = logcalendar.getTime();
	       logps.println(logformat.format(nowlog)+rptKind+":"+errRptMsg);
	       logps.flush();
	    }
    }
    
}
