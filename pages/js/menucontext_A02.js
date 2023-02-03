function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "法定比率資料", "法定比率資料",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "1.月平均放款總額", "1.月平均放款總額 ",  "#990110a");
	menu.addSubItem("A01Div01left", "2.內部融資餘額", "2.內部融資餘額 ", "#990150");
	menu.addSubItem("A01Div01left", "3.非會員存款總額", "3.非會員存款總額", "#990240");
	menu.addSubItem("A01Div01left", "4.贊助會員授信總額", "4.贊助會員授信總額", "#990320");
	menu.addSubItem("A01Div01left", "5.非會員無擔保消費性貸款", "5.非會員無擔保消費性貸款", "#990420");
	menu.addSubItem("A01Div01left", "6.非會員授信總額", "6.非會員授信總額", "#990512");
	menu.addSubItem("A01Div01left", "7.自用住宅放款總額", "7.自用住宅放款總額", "#990630");
	menu.addSubItem("A01Div01left", "8.信用部固定資產淨額", "8.信用部固定資產淨額", "#990720");
	menu.addSubItem("A01Div01left", "9.外幣資產總額(折合新台幣)", "9.外幣資產總額(折合新台幣)", "#990810");
	menu.addSubItem("A01Div01left", "10.對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者擔保放款最高限額", "10.對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者擔保放款最高限額", "#990930");
	menu.addSubItem("A01Div01left", "11.正會員（含同戶家屬）及同一關係人放款之最高限額", "11.正會員（含同戶家屬）及同一關係人放款之最高限額", "#991020");
	menu.addSubItem("A01Div01left", "12.每一贊助會員及同一關係人放款之最高限額", "12.每一贊助會員及同一關係人放款之最高限額", "#991120");
	menu.addSubItem("A01Div01left", "13.對同一非會員及同一關係人放款之最高限額", "13.對同一非會員及同一關係人放款之最高限額", "#991220");
	menu.addSubItem("A01Div01left", "14.最近決算年度", "14.最近決算年度", "#991320");
	menu.showMenu();
}