//94.10.21 fix 所以link.都移到該item的前一項,才不會被index蓋到 by 2295
function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "資產負債表(資產)", "資產負債表(資產)",  null, null);
	menu.addItem("A01Div01right", "資產負債表(負債)", "資產負債表(負債)",  null, null);	
	menu.addItem("A01Div03", "逾期放款金額", "逾期放款金額", "#100000" , "#100000");
	menu.addItem("A01Div02", "損益表", "損益表",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "流動資產", "流動資產",  "#110000");
	menu.addSubItem("A01Div01left", "存放合庫--小計", "存放合庫--小計",  "#110216");
	menu.addSubItem("A01Div01left", "存放農業金庫--小計", "存放農業金庫--小計",  "#110256");	
	menu.addSubItem("A01Div01left", "存放行庫--合計", "存放行庫--合計",  "#110240");
	
	menu.addSubItem("A01Div01left", "放款", "放款",  "#112400");
	menu.addSubItem("A01Div01left", "農業基金放款--小計", "農業基金放款--小計",  "#120604");
	menu.addSubItem("A01Div01left", "基金及出資", "基金及出資",  "#150300");
	menu.addSubItem("A01Div01left", "固定資產", "固定資產",  "#130300");
	menu.addSubItem("A01Div01left", "其他資產", "其他資產",  "#141200");
	menu.addSubItem("A01Div01left", "往來", "往來",  "#152000");
	
	
	menu.addSubItem("A01Div01right", "流動負債", "流動負債",  "#990000");
	menu.addSubItem("A01Div01right", "透支行庫--合計", "透支行庫--合計",  "#210105");
	 
	menu.addSubItem("A01Div01right", "短期借款--合計", "短期借款--合計",  "#210205");
	menu.addSubItem("A01Div01right", "存款", "存款",  "#211200");
	menu.addSubItem("A01Div01right", "長期負債", "長期負債",  "#221000");
	menu.addSubItem("A01Div01right", "長期借款--合計", "長期借款--合計",  "#240105");
	menu.addSubItem("A01Div01right", "借入專案放款資金--合計", "借入專案放款資金--合計",  "#240205");
	menu.addSubItem("A01Div01right", "借入農發資金-農建-小計", "借入農發資金-農建-小計",  "#240315");
	menu.addSubItem("A01Div01right", "借入農發基金放款資金--合計", "借入農發基金放款資金--合計",  "#240340");
	menu.addSubItem("A01Div01right", "其他負債", "其他負債",  "#240400");
	 
	menu.addSubItem("A01Div01right", "往來", "往來",  "#251300");
	menu.addSubItem("A01Div01right", "負債合計", "負債合計",  "#260700");
	menu.addSubItem("A01Div01right", "事業資金及公積", "事業資金及公積",  "#200000");
    menu.addSubItem("A01Div01right", "盈虧及損益", "盈虧及損益",  "#310900");
    menu.addSubItem("A01Div01right", "淨值合計", "淨值合計",  "#320300");
    menu.addSubItem("A01Div01right", "負債及淨值合計", "負債及淨值合計",  "#300000");
	 
    menu.addSubItem("A01Div02", "業務支出", "業務支出",  "#600000");		
	menu.addSubItem("A01Div02", "業務收入", "業務收入",  "#500000");	
    //menu.addSubItem("A01Div02", "合計", "合計",  "#420000");
     
	
	menu.showMenu();
}