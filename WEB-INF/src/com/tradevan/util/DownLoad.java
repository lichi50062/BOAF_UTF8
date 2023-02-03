//94.06.23 add若為負數時.把負號放在最前面 by 2295

package com.tradevan.util;
public class DownLoad {		//modify by 2354 2004.12.21
	public static String getFileName(String DLId, String bank_code, String sYear, String sMonth) {
		String	tmpBankCode	= null;
		String	tmpYear		= null;
		String	tmpMonth	= null;
		if(DLId.substring(0,1).equals("A")) tmpBankCode = (bank_code.trim() + "0000000").substring(0, 7); 
		else                                           tmpBankCode = "";
		tmpYear = ("000" + sYear.trim()).substring(sYear.trim().length());
		tmpMonth = ("00" + sMonth).substring(sMonth.length());

		return DLId + tmpBankCode + tmpYear + tmpMonth;
	} //End of getFileName

	/****************************************************************
	*	在資料的左方或右方補上指定的值(ex:"0")
	* 	data: Original data
	*	L_R:填充的位置
	*	stuff:填充字元
	*	len:產出資料的長度
	*/
	public static String fillStuff(String data, String L_R, String stuff, int len) {
		StringBuffer	szStuff = new StringBuffer();
		String			newStr 	= null;
		byte[] byte1 =null;	//add by 2354 12.24
		int	datalen	= 0;

		if (data == null)	data = "";

		if (data.length() > len)						//如果超過要取出的長度,截去左邊多出來的位數
			data = data.substring(data.length() - len);

		//datalen = data.trim().length();		//delete by 2354 12.24
		try{
			byte1 = data.trim().getBytes("BIG5");	//add by 2354 12.24
			datalen = byte1.length;					//add by 2354 12.24
		}catch(Exception e){
			datalen = data.trim().length();		//add by 2354 12.24
		}    
		datalen = byte1.length;					//add by 2354 12.24

		if (datalen < len) {
			for (int i = 1; i <= (len - datalen); i++)
				szStuff.append(stuff);
		}

		if (L_R.equals("L")) {							//填在左邊			
			szStuff.append(data.trim());
			newStr = szStuff.toString();
		}

		if (L_R.equals("R"))							//填在右邊
			newStr = data.trim() + szStuff.toString();

		return newStr;

	} //End of fillStuff()

	/****************************************************************
	*	在資料的左方或右方補上指定的值(ex:"0")
	*	data: Original data
	*	L_R:填充的位置
	*	stuff:填充字元
	*	digDec:小數點後要取幾位
	*	len:產出資料的長度
	* 	decimal(16, 0) ->getString();出來不會有小數點,雖然在dbaccess顯示0.0
	*/
	public static String fillStuff(String Data, String L_R, String stuff, int digDec, int len) {
		StringBuffer	szStuff = new StringBuffer();
		String			newStr 	= null;
		String			intData	= null;
		String			decData = null;
		String			newData = null;		
		int				datalen	= 0;
		int				idx		= 0;

		if (Data == null)	Data = "";

		idx = Data.indexOf(".");
		if (idx >= 0) {									//找到小數點
			if (idx == 0)
				intData = "";
			else
				intData = Data.substring(0, idx);
			decData = Data.substring(idx + 1);
			if (digDec == 0)
				decData = "";
			else if (decData.length() > digDec)
				decData = decData.substring(0, digDec);
			else if (decData.length() < digDec) {
					for (int i = 0; i < digDec - decData.length(); i++)
						decData = decData + "0";
			}
			newData = intData + decData;
		}
		else
			newData = Data;

		if (newData.length() > len)						//如果超過要取出的長度,截去左邊多出來的位數
			newData = newData.substring(newData.length() - len);

		datalen = newData.length();

		if (datalen < len) {
			for (int i = 1; i <= (len - datalen); i++)
				szStuff.append(stuff);
		}

		if (L_R.equals("L")) {							//填在左邊
			//System.out.println("szStuff="+szStuff);
			//System.out.println("newData="+newData);
			//94.06.23 add若為負數時.把負號放在最前面 by 2295==============
			if(newData.startsWith("-")){				
				szStuff.insert(0,"-");
				newData=newData.substring(1,newData.length());				
			}
			//========================================================
			szStuff.append(newData);
			newStr = szStuff.toString();
			//System.out.println("newStr="+newStr);
		}

		if (L_R.equals("R"))							//填在右邊
			newStr = newData + szStuff.toString();

		return newStr;

	} //End of fillStuff()

	/****************************************************************
	*	在資料的左方或右方補上指定的值(ex:"0")
	* 	data: Original data
	*	L_R:填充的位置
	*	stuff:填充字元
	*	dt:"D"->to_char(beg_date, '%Y%m%d') 07/12/2000 -> 20000712
	*	len:產出資料的長度
	* 	decimal(16, 0) ->getString();出來不會有小數點,雖然在dbaccess顯示0.0
	*/
	public static String fillStuff(String data, String L_R, String stuff, String dt, int len) {
		StringBuffer	szStuff = new StringBuffer();
		String			newStr 	= null;
		int	datalen	= 0;
		int	yr		= 0;

		if (data == null)
			data = "";
		else {
			yr = Integer.parseInt(data.substring(0, 4)) - 1911;
			data = Integer.toString(yr) + data.substring(4);
		}
		if (data.length() > len)						//如果超過要取出的長度,截去左邊多出來的位數
			data = data.substring(data.length() - len);

		datalen = data.trim().length();

		if (datalen < len) {
			for (int i = 1; i <= (len - datalen); i++)
				szStuff.append(stuff);
		}

		if (L_R.equals("L")) {							//填在左邊
			szStuff.append(data.trim());
			newStr = szStuff.toString();
		}

		if (L_R.equals("R"))							//填在右邊
			newStr = data.trim() + szStuff.toString();

		return newStr;

	} //End of fillStuff()
} //End of class()