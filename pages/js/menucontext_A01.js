function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "資產負債表(資產)", "資產負債表(資產)",  null, null);
	menu.addItem("A01Div01right", "資產負債表(負債)", "資產負債表(負債)",  null, null);	
	menu.addItem("A01Div03", "逾期放款金額", "逾期放款金額", "#190000" , "#190000");
	menu.addItem("A01Div02", "損益表", "損益表",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "流動資產", "流動資產",  "#110000");
	menu.addSubItem("A01Div01left", "存放合庫--小計", "存放合庫--小計",  "#110307");
	menu.addSubItem("A01Div01left", "存放農業金庫--小計", "存放農業金庫--小計",  "#110327");	
	menu.addSubItem("A01Div01left", "存放行庫--合計", "存放行庫--合計",  "#110313");
	
	menu.addSubItem("A01Div01left", "放款", "放款",  "#112500");
	menu.addSubItem("A01Div01left", "農業基金放款--小計", "農業基金放款--小計",  "#120604");
	menu.addSubItem("A01Div01left", "基金及出資", "基金及出資",  "#150300");
	menu.addSubItem("A01Div01left", "固定資產", "固定資產",  "#130300");
	menu.addSubItem("A01Div01left", "其他資產", "其他資產",  "#141600");
	menu.addSubItem("A01Div01left", "往來", "往來",  "#152200");
	
	
	menu.addSubItem("A01Div01right", "流動負債", "流動負債",  "#100000");
	menu.addSubItem("A01Div01right", "透支行庫--合計", "透支行庫--合計",  "#210205");
	 
	menu.addSubItem("A01Div01right", "短期借款--合計", "短期借款--合計",  "#210405");
	menu.addSubItem("A01Div01right", "存款", "存款",  "#211100");
	menu.addSubItem("A01Div01right", "長期負債", "長期負債",  "#221000");
	menu.addSubItem("A01Div01right", "長期借款--合計", "長期借款--合計",  "#240105");
	menu.addSubItem("A01Div01right", "借入農發資金-農建-小計", "借入農發資金-農建-小計",  "#240215");
	menu.addSubItem("A01Div01right", "借入農發基金放款資金--合計", "借入農發基金放款資金--合計",  "#240240");
	menu.addSubItem("A01Div01right", "借入專案放款資金--合計", "借入專案放款資金--合計",  "#240305");
	menu.addSubItem("A01Div01right", "其他負債", "其他負債",  "#240400");
	 
	menu.addSubItem("A01Div01right", "往來", "往來",  "#251800");
	menu.addSubItem("A01Div01right", "事業資金及公積", "事業資金及公積",  "#260600");
    menu.addSubItem("A01Div01right", "盈虧及損益", "盈虧及損益",  "#310900");
    menu.addSubItem("A01Div01right", "負債及淨值", "負債及淨值",  "#320300");
	 
    menu.addSubItem("A01Div02", "支出", "支出",  "#522700");		
	menu.addSubItem("A01Div02", "收入", "收入",  "#422700");	
    menu.addSubItem("A01Div02", "合計", "合計",  "#420000");
     
	
	menu.showMenu();
}