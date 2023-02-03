/* 修改DL_ListFor_1 變數內容 by Winnin 2004/11/22 
*  DL_ListFor_? structure
*  {"DLId", "DL table name"}
*  95.04.06 增加A06要檢核的rowleng.bank_no.金額index by 2295 
*  96.07.10 add A08存款帳戶分級差異化異資料 by 2295
*  97.01.02 add A09基本放款利率計價之舊貸案件資料 by 2295
*  97.06.13 add A10應予評估資產彙總資料 by 2295
*  104.01.13 add A12 by 2968
*  104.10.12 add A13 by 2295
*  109.06.11 add A15 by 2295
*/

package com.tradevan.util;

public class ListArray {
   
	public static final String[][] DL_ListFor_1 = {								//本國
											{"A01", "營運資料明細檔", "A01"},
											{"A02", "法定比率資料", "A02"},
											{"A03", "平均利率資料", "A03"},
											{"A04", "資產品質分析資料", "A04"},
											{"A05", "資本適足率資料", "A05"},
											{"A06", "逾期放款資料", "A06"},
											{"A08", "存款帳戶分級差異化異資料", "A08"},/*96.07.10 add*/
											{"A09", "基本放款利率計價之舊貸案件資料", "A09"},/*96.12.03 add*/
											{"A99", "法定比率延申資料", "A99"},/*95.05.25 add*/
											{"A10", "應予評估資產彙總資料", "A10"},/*97.06.12 add*/
											{"A12", "基層機構逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表", "A12"},/*104.01.08 add*/
											{"A13", "建築貸款及農地放款統計資料", "A13"},/*104.10.12 add*/
											{"A15", "電子銀行及行動支付業務辦理情形資料", "A15"}/*109.06.11 add*/
											};
	
	public static final String[][] DL_ListFor_3 = {								//外銀
											{"B01", "農業發展基金及農業天然災害基金貸款執行情形表", "B01"},
											{"B02", "農業發展基金及農業天然災害救助基金貸放餘額統計", "B02"},
											{"B03", "農業發展基金貸款有關統計資料表", "B03"}
											};
	public static final String[][] DL_ListFor_4 = {								//農業信用保證基金
											{"M01", "保證案件月報表(按保證項目)", "M01"},
											{"M02", "保證案件月報表(按貸款機構)", "M02"},
											{"M03", "本月份與上月份及去年同期保證情形比較表", "M03"},
											{"M04", "貸款用途分析表", "M04"},
											{"M05", "代位清償案件原因分析表", "M05"},
											{"M06", "地區別保證案件分析表(農會)", "M06"},
											{"M07", "地區別保證案件分析表(漁會)", "M07"},
											{"M08", "保證案件_身份別月報表", "M08"}
											};
	public static final String[][] DL_ListFor_5 = {								//信合社
											{"E01", "營運資料明細檔", "nmpbank6_t"}
											};
	public static final String[][] DL_ListFor_6 = {								//94.11.14 add
											{"F01", "在台無住所外國人新台幣存款資料", "F01"}
											};
	
	public static final String[][] CONST_RULENO = {{"FIX_K", "1000"}};
	//93.11.19回傳報表的欄位長度 modify by 2354 2004.12.14
	public static final String[][] Report_RowLength = {{"A01", "32",""},
													   {"A02", "32",""},	
													   {"A03", "32",""},	
													   {"A04", "32",""},	
													   {"A05", "48",""},	
													   {"A06", "102",""},
													   {"A08", "82",""},/*96.07.10 add*/	
													   {"A09", "96",""},/*97.01.02 add*/	
													   {"A10", "138",""},/*97.06.13 add*/
													   {"A99", "32",""},
													   {"A13", "32",""},/*104.10.12 add*/
													   {"A15", "46",""},/*109.06.11 add*/
													   {"M01", "141",""},
													   {"M02", "141",""},
													   {"M03", "114",""},
													   {"M03_11", "11",""},
													   {"M03_18", "18",""},
													   {"M04", "169",""},
													   {"M05", "153",""},
													   {"M05_11", "11",""},
													   {"M05_18", "18",""},
													   {"M05_25", "25",""},
													   {"M06", "148",""},
													   {"M07", "148",""}
												      };
	//93.11.19回傳多個要檢查的數字index
	public static final String[][] Report_RowIsNumber = {{"A01", "18","32"},
														 {"A02", "18","32"},
														 {"A03", "18","32"},
														 {"A04", "18","32"},
														 {"A05", "18","32"},
														 {"A06", "18","32"},
														 {"A06", "32","46"},
														 {"A06", "46","60"},
														 {"A06", "60","74"},
														 {"A06", "74","88"},
														 {"A06", "88","102"},
														 {"A08", "12","26"},/*96.07.10 add*/
														 {"A08", "26","40"},/*96.07.10 add*/
														 {"A08", "40","54"},/*96.07.10 add*/
														 {"A08", "54","68"},/*96.07.10 add*/
														 {"A08", "68","82"},/*96.07.10 add*/
														 {"A09", "12","26"},/*97.01.02 add*/
														 {"A09", "26","40"},/*97.01.02 add*/
														 {"A09", "40","54"},/*97.01.02 add*/
														 {"A09", "54","68"},/*97.01.02 add*/
														 {"A09", "68","82"},/*97.01.02 add*/
														 {"A09", "82","96"},/*97.01.02 add*/
														 {"A10", "12","26"},/*97.06.13 add*/
														 {"A10", "26","40"},/*97.06.13 add*/
														 {"A10", "40","54"},/*97.06.13 add*/
														 {"A10", "54","68"},/*97.06.13 add*/
														 {"A10", "68","82"},/*97.06.13 add*/
														 {"A10", "82","96"},/*97.06.13 add*/
														 {"A10", "96","110"},/*97.06.13 add*/
														 {"A10", "110","124"},/*97.06.13 add*/
														 {"A10", "124","138"},/*97.06.13 add*/														 
														 {"A99", "18","32"},
														 {"A13", "18","32"},/*104.10.12 add*/
														 {"A15", "18","32"},/*109.06.11 add*/
														 {"A15", "32","46"},/*109.06.11 add*/
														 {"M01", "6","13"},
														 {"M01", "13","27"},
														 {"M01", "27","41"},
														 {"M01", "41","55"},
														 {"M01", "55","69"},
														 {"M01", "69","76"},
														 {"M01", "76","87"},
														 {"M01", "87","94"},
														 {"M01", "94","105"},
														 {"M01", "105","112"},
														 {"M01", "112","123"},
														 {"M01", "123","130"},
														 {"M01", "130","141"},
														 {"M02", "6","13"},
														 {"M02", "13","27"},
														 {"M02", "27","41"},
														 {"M02", "41","55"},
														 {"M02", "55","69"},
														 {"M02", "69","76"},
														 {"M02", "76","87"},
														 {"M02", "87","94"},
														 {"M02", "94","105"},
														 {"M02", "105","112"},
														 {"M02", "112","123"},
														 {"M02", "123","130"},
														 {"M02", "130","141"},
														 {"M03", "2","9"},
														 {"M03", "9","23"},
														 {"M03", "23","37"},
														 {"M03", "37","44"},
														 {"M03", "44","58"},
														 {"M03", "58","72"},
														 {"M03", "72","86"},
														 {"M03", "86","100"},
														 {"M03", "100","114"},
													     {"M03_18", "4","18"},
													     {"M03_11", "4","11"},
														 {"M04", "1", "8"}, 
														 {"M04", "8", "15"}, 
														 {"M04", "15", "29"}, 
														 {"M04", "29", "36"}, 
														 {"M04", "36", "50"}, 
														 {"M04", "50", "57"}, 
														 {"M04", "57", "64"}, 
														 {"M04", "64", "71"}, 
														 {"M04", "71", "85"}, 
														 {"M04", "85", "92"}, 
														 {"M04", "92", "106"}, 
														 {"M04", "106", "113"}, 
														 {"M04", "113", "120"}, 
														 {"M04", "120", "127"}, 
														 {"M04", "127", "141"}, 
														 {"M04", "141", "148"}, 
														 {"M04", "148", "162"}, 
														 {"M04", "162", "169"},														 
														 {"M05","6","13"},
														 {"M05","13","27"},
														 {"M05","27","34"},
														 {"M05","34","48"},
														 {"M05","48","55"},
														 {"M05","55","69"},
														 {"M05","69","76"},
														 {"M05","76","90"},
														 {"M05","90","97"},
														 {"M05","97","111"},
														 {"M05","111","118"},
														 {"M05","118","132"},
														 {"M05","132","139"},
														 {"M05","139","153"},
													     {"M05_18", "4","18"},
													     {"M05_11", "4","11"},
														 {"M06", "1","8"},
														 {"M06", "8","22"},
														 {"M06", "22","36"},
														 {"M06", "36","43"},
														 {"M06", "43","57"},
														 {"M06", "57","71"},
														 {"M06", "71","78"},
														 {"M06", "78","92"},
														 {"M06", "92","106"},
														 {"M06", "106","113"},
														 {"M06", "113","127"},
														 {"M06", "127","134"},
														 {"M06", "134","148"},
														 {"M07", "1","8"},
														 {"M07", "8","22"},
														 {"M07", "22","36"},
														 {"M07", "36","43"},
														 {"M07", "43","57"},
														 {"M07", "57","71"},
														 {"M07", "71","78"},
														 {"M07", "78","92"},
														 {"M07", "92","106"},
														 {"M07", "106","113"},
														 {"M07", "113","127"},
														 {"M07", "127","134"},
														 {"M07", "134","148"}
		      										};
	//93.11.19回傳總機構代號的index
	public static final String[][] Report_BankNo = {{"A01", "5","12"},
													{"A02", "5","12"},
													{"A03", "5","12"},
													{"A04", "5","12"},
													{"A05", "5","12"},													
													{"A06", "5","12"},
													{"A08", "5","12"},/*96.07.10 add*/
													{"A09", "5","12"},/*97.01.02 add*/
													{"A10", "5","12"},/*97.06.13 add*/
													{"A13", "5","12"},/*104.10.12 add*/
													{"A15", "5","12"},/*106.06.11 add*/
													{"A99", "5","12"},
													{"M01", "0","1"},
													{"M02", "0","1"},
													{"M03", "0","2"},
													{"M03_11", "0","4"},
													{"M03_18", "0","4"},
													{"M04", "0","1"},
													{"M05", "0","6"},
													{"M05_11", "0","4"},
													{"M05_18", "0","4"},
													{"M05_25", "0","4"},
													{"M06", "0","1"},
													{"M07", "0","1"}
		      										};

	//93.12.10回傳項目代號的index 
	public static final String[][] Report_BankCode = {{"A01", "12","18"},
													{"A02", "12","18"},
													{"A03", "12","18"},
													{"A04", "12","18"},
													{"A05", "12","18"},
													{"A06", "12","18"},	
													{"A13", "12","18"},/*104.10.12 add*/
													{"A13", "12","18"},/*109.06.11 add*/
													{"A99", "12","18"},
													{"M01", "1","6"},
													{"M02", "1","6"},
													{"M03", "0","2"},
													{"M03_18", "0","4"},
													{"M03_11", "0","4"},
													{"M04", "0","1"},
													{"M05", "0","6"},
													{"M05_25", "0","4"},
													{"M05_18", "0","4"},
													{"M05_11", "0","4"},
													{"M06", "0","1"},
													{"M07", "0","1"}
		      										};


	//93.12.10根據長度回傳需抓取之Report_no的item
	//報表編號,長度,檢查碼,Code起始位置,Code結束位置,Code的編碼原則
	public static final String[][] Report_ReportItem = {{"M03","11", "M03_11","0","4","M03_NOTE"},
												        {"M03","18", "M03_18","0","4","M03_NOTE"},
												        {"M03","114","M03","0","2","M03"},
												        {"M05", "11","M05_11","0","4","M05_NOTE"},
												        {"M05", "18","M05_18","0","4","M05_NOTE"},
												        {"M05", "25","M05_25","0","4","M05_TOTACC"},
												        {"M05", "153","M05","0","6","M05"}
		      									       };

	/****************************************************************
	* Use banktype to get corresponding DL_ListFor_?
	*/
	public static String[][] getListArray(String BankType) {
		String[][] e = {{"", ""}};

		
		if (BankType.equals("1"))	return DL_ListFor_1;
		if (BankType.equals("3"))	return DL_ListFor_3;
		if (BankType.equals("4"))	return DL_ListFor_4;
		if (BankType.equals("5"))	return DL_ListFor_5;

		return e;
	}

	/****************************************************************
	* Use banktype to get corresponding UP_ListFor_?
	*/
	public static String[][] getUpListArray(String BankType) {
		String[][] e = {{"",""}};

		if (BankType.equals("1"))	return DL_ListFor_1;
		if (BankType.equals("3"))	return DL_ListFor_3;
		if (BankType.equals("4"))	return DL_ListFor_4;
		if (BankType.equals("5"))	return DL_ListFor_5;
		if (BankType.equals("6"))	return DL_ListFor_6;

		return e;
	}

	/****************************************************************
	* Use DLId to get corresponding table
	*/
	public static String getTempTableName(String BankType, String DLId) {
		String[][] s = null;
		s = getUpListArray(BankType);
		for (int i = 0; i < s.length; i++) {
			if (s[i][0].equals(DLId))	return s[i][2];
		}
		return "";
	}
	/****************************************************************
	* Use DLId to get DLIdName
	*/
	public static String getDLIdName(String BankType, String DLId) {
		String[][] s = null;
		s = getUpListArray(BankType);
		for (int i = 0; i < s.length; i++) {
			if (s[i][0].equals(DLId))	return s[i][1];
		}
		return "";
	}
	/****************************************************************
	* Use code return corresponding CONST VALUE
	*
	*/
	public static String getCONST_VALUE(String constid) {
		if (constid == null) return "";
		for (int i = 0; i < CONST_RULENO.length; i ++) {
			if (CONST_RULENO[i][0].equals(constid.trim())) {
				return CONST_RULENO[i][1];
			}
		}
		return "";
	}

	/****************************************************************
	* 93.11.19
	* Use Report_no return corresponding rowlength
	* 
	*/
	public static String getRowLength(String Report_no) {
		if (Report_no == null) return "";
		for (int i = 0; i < Report_RowLength.length; i ++) {
			if (Report_RowLength[i][0].equals(Report_no.trim())) {
				return Report_RowLength[i][1];
			}
		}
		return "";
	}
	/****************************************************************
	* 93.11.19
	* Use Report_no return corresponding checkNumberIndexCount
	* 
	*/
	public static int[][] getRowIsNumber(String Report_no) {
		int [][] checkNumber = null;		
		int Count = 0;
		if (Report_no == null) return checkNumber;
		
		for (int i = 0; i < Report_RowIsNumber.length; i ++) {
			if (Report_RowIsNumber[i][0].equals(Report_no.trim())) {
				Count ++;
			}
		}
		checkNumber = new int [Count][2];
		int idxCount = 0;
		for (int i = 0; i < Report_RowIsNumber.length; i ++) {
			if (Report_RowIsNumber[i][0].equals(Report_no.trim())) {				
				checkNumber[idxCount][0]= Integer.parseInt(Report_RowIsNumber[i][1]); 
				checkNumber[idxCount][1]= Integer.parseInt(Report_RowIsNumber[i][2]);
				idxCount ++;
			}
		}
		
		return checkNumber;
	}
	/****************************************************************
	* 93.11.19
	* Use Report_no return corresponding bankno index
	* 
	*/
	public static int[][] getBankNo(String Report_no) {
		int [][] checkNumber = {{0,0}};	
		
		if (Report_no == null) return checkNumber;
		
		forLoop:
		for (int i = 0; i < Report_BankNo.length; i ++) {
			if (Report_BankNo[i][0].equals(Report_no.trim())) {				
				checkNumber[0][0]= Integer.parseInt(Report_BankNo[i][1]); 
				checkNumber[0][1]= Integer.parseInt(Report_BankNo[i][2]);
				break forLoop;
			}
		}
		
		return checkNumber;
	}

	/****************************************************************
	* 93.12.10
	* Use Report_no return corresponding BankCode index
	* 
	*/
	public static int[][] getBankCode(String Report_no) {
		int [][] checkNumber = {{0,0}};	
		
		if (Report_no == null) return checkNumber;
		
		forLoop:
		for (int i = 0; i < Report_BankCode.length; i ++) {
			if (Report_BankCode[i][0].equals(Report_no.trim())) {				
				checkNumber[0][0]= Integer.parseInt(Report_BankCode[i][1]); 
				checkNumber[0][1]= Integer.parseInt(Report_BankCode[i][2]);
				break forLoop;
			}
		}
		
		return checkNumber;
	}
	
	/****************************************************************
	* 93.12.14
	* 根據長度取得應使用之判斷編號
	* 回傳:  檢查碼,Code起始位置,Code結束位置,Code的編碼原則
	*/
	public static String[] getReportItem(String Report_no,int rowLength) {
		String[] reportItem = {"","","",""} ;
		
		forLoop:
		for (int i = 0; i < Report_ReportItem.length; i ++) {
			if (Report_ReportItem[i][0].equals(Report_no.trim()) && Integer.parseInt(Report_ReportItem[i][1])==rowLength) {				
				reportItem[0] = Report_ReportItem[i][2];
				reportItem[1] = Report_ReportItem[i][3];
				reportItem[2] = Report_ReportItem[i][4];
				reportItem[3] = Report_ReportItem[i][5];
				break forLoop;
			}
		}
		return reportItem;
		
	}
} //End of class											
