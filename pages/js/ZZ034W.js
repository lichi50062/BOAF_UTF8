// 95.08.16 fix 拿掉A06/A04比對 by 2295
// 98.09.01 add A01/A02/A99跨表檢核 by 2295
//102.08.22 add 103.01以後.漁會套用新科目代號.A01/A05與A01/A02/A99取消農漁會共同顯示 by 2295
function changeOption(form){
    var myXML,nodeType,nodeValue, nodeName;

    myXML = document.all("ReportListXML").XMLDocument;
    form.UPD_CODE.length = 0;
    form.BANK_TYPE_List.length = 0;
	nodeType = myXML.getElementsByTagName("fileType");
	nodeValue = myXML.getElementsByTagName("optionValue");
	nodeName = myXML.getElementsByTagName("optionName");
	var oOption;	
	for(var i=0;i<nodeType.length ;i++)
	{		
  		if ((nodeType.item(i).firstChild.nodeValue == form.CheckOption.value))  {
  			oOption = document.createElement("OPTION");
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;
  			form.UPD_CODE.add(oOption);
    	}
    }
     	   
    if(form.CheckOption.value == "2"){//A01.A06         
       document.all.item("div_memo").innerHTML = '<li>◆<font color=#FF0000>A01/A06放款額與逾放額之比對範圍說明如下：</font></li>'
										  	   + '<li>&nbsp; ▲<font color=#0000FF>逾期放款合計：比對(990000)及(970000)與(840740)三個申報的數字是相等</font></li>'
										  	   + '<li>&nbsp; ▲<font color=#0000FF>其他各放款科目的逾期放款合計數<span style="font-size: 12.0pt; font-family: 新細明體">，</span>不可大於各該科目的放款合計數</font><font color=#FF0000>(註:[150200]及[960500]兩個科目不比對)</font></li>'
                          				  	   + '<li>◆<font color=#FF0000>查詢(農漁會信用部)權限說明</font><font color=#0000FF>：</font></li>'
										  	   + '<li>&nbsp;&nbsp;▲「農金局/中央存保及全國農業金庫」：可查詢全部機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>'
                          				  	   + '<li>◆<font color=#FF0000>各欄位金額，如果顯示內容是空白，表示未申報</font></li>';                          				  	  
        oOption = document.createElement("OPTION");
		oOption.text="農會";
  		oOption.value="6";
  		form.BANK_TYPE_List.add(oOption);    
    	oOption = document.createElement("OPTION");
		oOption.text="漁會";
  		oOption.value="7";
  		form.BANK_TYPE_List.add(oOption);    
    }else if(form.CheckOption.value == "4"){//A01.A02.A99   	
       document.all.item("div_memo").innerHTML = '<li>◆<font color="#FF0000">A01/A02/A99比對範圍說明如下：</font></li>'
						  					   + '<li>&nbsp; ▲<font color="#0000FF">A02.990210[內部融資餘額] >= A01.120700[內部融資] </font></li>'
						  					   + '<li>&nbsp; ▲<font color="#0000FF">A02.990210[內部融資餘額] >= A02.990220[內部融資餘額(中、長期)]</font></li>'
						  					   + '<li>&nbsp; ▲<font color="#0000FF">A99.992150(無擔保消費性貸款) >= A02.990510[非會員無擔保消費性貸款]</font></li> '
						  					   + '<li>&nbsp; ▲<font color="#0000FF">A99.992150(無擔保消費性貸款) >= A99.992550(無擔保消費性貸款中之逾期放款)</font></li>'
						  					   + '<li>&nbsp; ▲<font color="#0000FF">A99.992150(無擔保消費性貸款) >= A99.992650(無擔保消費性貸款中之應予觀察放款)</font></li>'
						  					   + '<li>&nbsp; ▲<font color="#0000FF">其他各該科目左式的合計數<span style="font-size: 12.0pt; font-family: 新細明體">，</span>需等於各該科目右式的合計數</font></li>'
       										   + '<li>◆<font color=#FF0000>查詢(農漁會信用部)權限說明：</font><font color=#0000FF>：</font></li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農金局/中央存保及全國農業金庫」：可查詢全部機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>'
                          				  	   + '<li>◆<font color=#FF0000>各欄位金額，如果顯示內容是空白，表示未申報</font></li>';                          				                	
		oOption = document.createElement("OPTION");
		oOption.text="農會";
  		oOption.value="6";
  		form.BANK_TYPE_List.add(oOption);
    
    	oOption = document.createElement("OPTION");
		oOption.text="漁會";
  		oOption.value="7";
  		form.BANK_TYPE_List.add(oOption);
        /*103.01以後.漁會套用新科目代號.取消農漁會共同顯示
    	oOption = document.createElement("OPTION");
		oOption.text="農漁會";
  		oOption.value="ALL";
  		form.BANK_TYPE_List.add(oOption);    
  		*/
    }else{//A01.A05、存保檢核    	
        document.all.item("div_memo").innerHTML = '<li>◆查詢(農漁會信用部)權限說明：</li>'
				  	   + '<li>&nbsp;&nbsp;▲「農金局/中央存保及全國農業金庫」：可查詢全部機構</li>'
				  	   + '<li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>'
				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>'
				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>'
				  	   + '<li>◆各欄位金額，如果顯示內容是空白，表示未申報</li>';                          				                	
		oOption = document.createElement("OPTION");
		oOption.text="農會";
		oOption.value="6";
		form.BANK_TYPE_List.add(oOption);
		
		oOption = document.createElement("OPTION");
		oOption.text="漁會";
		oOption.value="7";
		form.BANK_TYPE_List.add(oOption);
		/*103.01以後.漁會套用新科目代號.取消農漁會共同顯示
		oOption = document.createElement("OPTION");
		oOption.text="農漁會";
		oOption.value="ALL";
		form.BANK_TYPE_List.add(oOption);    
		*/                  				 
    }   
    /*else if(form.CheckOption.value == "3"){//A06.A04         
       document.all.item("div_memo").innerHTML = '<li>◆<font color=#FF0000>A06/A04之比對範圍說明如下：</font></li>'
										  	   + '<li>&nbsp; ▲<font color=#0000FF>A06=&gt;各逾放期數合計之科目【960500】欄位的金額，</font></li>'
										  	   + '<li><font color=#0000FF>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; '
										  	   + '應等於A04.【840731】+ A04.【840732】+A04.【840733】+A04.【840734】+ A04.【840735】的合計申報金額</font></li>'
										  	   + '<li>◆<font color=#FF0000>查詢(農漁會信用部)權限說明</font><font color=#0000FF>：</font></li>'
										  	   + '<li>&nbsp;&nbsp;▲「農金局/中央存保及全國農業金庫」：可查詢全部機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「共用中心」：可查詢加入該「共用中心」之總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會信用部總機構」：可查詢本身總機構</li>'
                          				  	   + '<li>&nbsp;&nbsp;▲「農(漁)會地方主管機關」：可查詢轄管總機構</li>'
                          				  	   + '<li>◆<font color=#FF0000>各欄位金額，如果顯示內容是空白，表示未申報</font></li>';                          				  	  
        oOption = document.createElement("OPTION");
		oOption.text="農會";
  		oOption.value="6";
  		form.BANK_TYPE_List.add(oOption);
    
    	oOption = document.createElement("OPTION");
		oOption.text="漁會";
  		oOption.value="7";
  		form.BANK_TYPE_List.add(oOption);
    
    	oOption = document.createElement("OPTION");
		oOption.text="農漁會";
  		oOption.value="ALL";
  		form.BANK_TYPE_List.add(oOption);                          				  	                                 				  	   
    }*/	   
}
