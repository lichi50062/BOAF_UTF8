function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	menu.addItem("webmasterid", "資產負債表", "資產負債表",  null, null);
	menu.addItem("newsid", "損益表", "損益表",  null, null);
	       
	menu.addSubItem("webmasterid", "業務支出", "業務支出",  "#520000");
	menu.addSubItem("webmasterid", "存款利息支出--合計", "存款利息支出--合計",  "#520100");
	menu.addSubItem("webmasterid", "借款利息支出--合計", "借款利息支出--合計",  "#520200");
	menu.addSubItem("webmasterid", "業務外支出", "業務外支出",  "#522000");
	menu.addSubItem("webmasterid", "業務收入", "業務收入",  "#420000");	
	menu.addSubItem("webmasterid", "放款利息收入--合計", "放款利息收入--合計",  "#420100");
	menu.addSubItem("webmasterid", "存儲利息收入--合計", "存儲利息收入--合計",  "#420300");	
	menu.addSubItem("webmasterid", "業務外收入", "業務外收入",  "#422000");
	
	menu.addSubItem("newsid", "好站連結", "好站連結",  "http://www.tacocity.com.tw/ppkoabcd/good.htm");
	menu.addSubItem("newsid", "免費資源", "免費資源",  "http://www.tacocity.com.tw/ppkoabcd/free.htm");
	menu.addSubItem("newsid", "港都一日遊", "港都一日遊",  "http://www.tacocity.com.tw/ppkoabcd/touristone.htm");
	menu.addSubItem("newsid", "飯店資料", "飯店資料",  "http://www.tacocity.com.tw/ppkoabcd/hotel.htm");
	
	menu.addSubItem("freedownloadid", "台灣網際 NO.1", "台灣網際 NO.1",  "http://ftp.world.net.tw/~pr/home.html");
	
		menu.showMenu();
}