function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01", "資產負債表", "資產負債表",  null, null);
	menu.addItem("A01Div02", "損益表", "損益表",  null, null);
	    
	
	menu.addSubItem("A01Div01", "流動資產", "流動資產",  "#110000");
	menu.addSubItem("A01Div01", "存放行庫--合作金庫--小計", "存放行庫--合作金庫--小計",  "#110000");
	menu.addSubItem("A01Div01", "存放行庫--全國農業金庫--小計", "存放行庫--全國農業金庫--小計",  "#110310");	
	menu.addSubItem("A01Div01", "存放行庫--合計", "存放行庫--合計",  "#110320");
	
	menu.addSubItem("A01Div01", "放款", "放款",  "#111400");
	menu.addSubItem("A01Div01", "農業發展基基放款--小計", "農業發展基基放款--小計",  "#120502");
	menu.addSubItem("A01Div01", "基金及出資", "基金及出資",  "#150300");
	menu.addSubItem("A01Div01", "固定資產", "固定資產",  "#130300");
	menu.addSubItem("A01Div01", "其他資產", "其他資產",  "#141600");
	menu.addSubItem("A01Div01", "往來", "往來",  "#160500");
	menu.addSubItem("A01Div01", "資產合計", "資產合計",  "#160600");
	menu.addSubItem("A01Div01", "流動負債", "流動負債",  "#100000");
	menu.addSubItem("A01Div01", "透支行庫--小計", "透支行庫--小計",  "#210000");
	 
	menu.addSubItem("A01Div01", "短期借款--小計", "短期借款--小計",  "#210200");
	menu.addSubItem("A01Div01", "存款", "存款",  "#211100");
	menu.addSubItem("A01Div01", "長期負債", "長期負債",  "#221000");
	menu.addSubItem("A01Div01", "長期借款--合計", "長期借款--合計",  "#240000");
	menu.addSubItem("A01Div01", "借入農業發展基金放款資金--借入農建放款資金--小計", "借入農業發展基金放款資金--借入農建放款資金--小計",  "#240100");
	menu.addSubItem("A01Div01", "借入農業發展基金放款資金--合計", "借入農業發展基金放款資金--合計",  "#240210");
	menu.addSubItem("A01Div01", "借入專案放款資金--合計", "借入專案放款資金--合計",  "#240200");
	menu.addSubItem("A01Div01", "其他負債", "其他負債",  "#240400");
	 
	menu.addSubItem("A01Div01", "往來", "往來",  "#260500");
	menu.addSubItem("A01Div01", "負債合計", "負債合計",  "#260600");
	menu.addSubItem("A01Div01", "事業資金及公積", "事業資金及公積",  "#200000");
	menu.addSubItem("A01Div01", "盈虧及損益", "盈虧及損益",  "#310900");
	menu.addSubItem("A01Div01", "淨值合計", "淨值合計",  "#320300");
	menu.addSubItem("A01Div01", "負債及淨值合計", "負債及淨值合計",  "#300000");
	
	
	
	 
	 
    menu.addSubItem("A01Div02", "業務支出", "業務支出",  "#522700");
	menu.addSubItem("A01Div02", "存款利息支出--合計", "存款利息支出--合計",  "#520000");
	menu.addSubItem("A01Div02", "借款利息支出--合計", "借款利息支出--合計",  "#520100");
	menu.addSubItem("A01Div02", "業務外支出", "業務外支出",  "#521900");
	menu.addSubItem("A01Div02", "業務收入", "業務收入",  "#500000");	
	menu.addSubItem("A01Div02", "放款利息收入--合計", "放款利息收入--合計",  "#500000");
	menu.addSubItem("A01Div02", "存儲利息收入--合計", "存儲利息收入--合計",  "#420100");	
	menu.addSubItem("A01Div02", "業務外收入", "業務外收入",  "#420800");	
     
     
	
	menu.showMenu();
}