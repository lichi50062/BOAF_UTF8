package com.tradevan.util;

public class ListArray_FC {

	public static final String[][] FC_BTYPE = {		//子公司別
											{"0", "金融控股公司"} ,
											{"1", "銀行業子公司"} ,
											{"2", "保險業子公司"} ,
											{"3", "證券及證金業子公司"} ,
											{"4", "票券業子公司"} ,
											{"5", "期貨業子公司"} ,
											{"6", "信託業子公司"} ,
											{"7", "信用卡業子公司"} ,
											{"8", "創業投資事業子公司"} ,
											{"9", "外國金融機構子公司"} ,
											{"A", "其他子公司"}
											};

	public static final String[][] BUSINESS_TYPE = {		//行業別
											{"0", "金融控股"},
											{"1", "銀行業"},
											{"2", "保險業"},
											{"3", "證券及證金業"},
											{"4", "票券業"},
											{"5", "期貨業"},
											{"6", "信託業"},
											{"7", "信用卡業"},
											{"8", "創業投資事業"},
											{"9", "外國金融機構"},
											{"A", "其他"},
											};

	public static final String[][] RELATION = {		//利害關係人類別
											{"0", "金融控股公司"},
											{"1", "銀行業子公司"},
											{"2", "保險業子公司"},
											{"3", "證券及證金業子公司"},
											{"4", "票券業子公司"},
											{"5", "期貨業子公司"},
											{"6", "信託業子公司"},
											{"7", "信用卡業子公司"},
											{"8", "創業投資事業子公司"},
											{"9", "外國金融機構子公司"},
											{"A", "其他子公司"},
											{"B", "金控公司及子公司之負責人及大股東"},
											{"C", "金控公司之負責人及大股東經營之企業或團體"},
											{"D", "與金控公司或其子公司半數以上董事相同之公司"},
											{"E", "其他利害關係人"}
											};

	public static final String[][] RELATION_F19 = {		//利害關係人類別
											{"B", "金控公司及子公司之負責人及大股東"},
											{"C", "金控公司之負責人及大股東經營之企業或團體"},
											{"D", "與金控公司或其子公司半數以上董事相同之公司"},
											{"E", "其他利害關係人"},
											{"F", "一般客戶"},
											};

	public static final String[][] RELATION_F20 = {		//利害關係人非屬子公司
											{"B", "金控公司及子公司之負責人及大股東"},
											{"C", "金控公司之負責人及大股東經營之企業或團體"},
											{"D", "與金控公司或其子公司半數以上董事相同之公司"},
											{"E", "其他利害關係人"}
											};

	public static final String[][] RELATION_F07 = {		//同一人、同一關係人、關係企業之關係代號
											{"1", "本人(本公司)"},
											{"2", "二親等血親"},
											{"3", "本人為負責人之企業"},
											{"4", "配偶為負責人之企業"},
											{"5", "前開公司之關係企業（包括控制公司、從屬公司及相互投資公司）"},
											{"6", "他人（他公司）"}
											};

	public static final String[][] TRANS_TYPE = {		//交易性質
											{"01", "存款"},
											{"02", "放款"},
											{"03", "拆放及資金週轉或融通"},
											{"04", "有價證券購買或投資"},
											{"05", "不動產或其他資產買賣"},
											{"06", "資本租賃交易"},
											{"07", "手續費、佣金及服務費收入"},
											{"08", "租金收支"},
											{"09", "存出及存入保證金"},
											{"10", "捐贈"},
											{"11", "保證及背書"},
											{"12", "其他交易"}
											};

	public static final String[][] TRANS_TYPE_F19 = {		//交易性質
											{"02", "放款"},
											{"04", "有價證券購買或投資"},
											{"05", "不動產或其他資產買賣"},
											{"06", "資本租賃交易"},
											{"07", "手續費、佣金及服務費收入"},
											{"08", "租金收支"},
											{"09", "存出及存入保證金"},
											{"10", "捐贈"},
											{"11", "保證及背書"},
											{"12", "其他交易"}
											};

	public static final String[][] CONST_RULENO = {{"FIX_2", "2"}};

	/****************************************************************
	* Use code return corresponding name
	*
	*/
	public static String getCodeName(int idx, String code) {

		String[][] arr = null;
		switch (idx) {
			case 1:						//子公司別
				arr = FC_BTYPE;
				break;
			case 2:
				arr = BUSINESS_TYPE; 	//行業別
				break;
			case 3:
				arr = RELATION;			//利害關係人類別
				break;
			case 4:
				arr = RELATION_F19;		//利害關係人類別
				break;
			case 5:
				arr = RELATION_F07;		//同一人、同一關係人、關係企業之關係代號
				break;
			case 6:
				arr = TRANS_TYPE;		//交易性質
				break;
			case 7:
				arr = TRANS_TYPE_F19;	//交易性質
				break;
			case 8:
				arr = RELATION_F20;		//利害關係人非屬子公司
				break;
			default:
				return "";
		}
		if (code == null)
			return "";

		for (int i = 0; i < arr.length; i++) {
			if (arr[i][0].equals(code.trim())) {
				return 	arr[i][1];
			}
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
} //End of class											