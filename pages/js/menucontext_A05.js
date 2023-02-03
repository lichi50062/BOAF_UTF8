function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "信用部淨值占風險性資產比率計算表", "信用部淨值占風險性資產比率計算表",  null, null);
	menu.addItem("A01Div01right", "信用部淨值佔風險性資料比率", "信用部淨值佔風險性資料比率",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "現金", "現金",  "#A05_div01");
	menu.addSubItem("A01Div01left", "上列以外依規定，風險權數未達100%之資產", "上列以外依規定，風險權數未達100%之資產", "#920601");	
	menu.addSubItem("A01Div01left", "上列以外之債權及其他資產", "上列以外之債權及其他資產",  "#92075P");
	
	menu.addSubItem("A01Div01right", "第一類資本", "第一類資本",  "#A05_div02");	
	menu.addSubItem("A01Div01right", "第二類資本", "第二類資本",  "#910199");	
	menu.addSubItem("A01Div01right", "淨值總額", "淨值總額",  "#910299");	
	menu.addSubItem("A01Div01right", "減項", "減項",  "#910300");	
	menu.addSubItem("A01Div01right", "合格淨值", "合格淨值",  "#910403");	
	menu.addSubItem("A01Div01right", "風險性資產總額", "風險性資產總額",  "#910400");	
	menu.addSubItem("A01Div01right", "信用部淨值占風險性資產比率", "信用部淨值占風險性資產比率",  "#910500");	
	
	menu.showMenu();
}