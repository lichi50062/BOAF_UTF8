function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	menu.addItem("webmasterid", "�겣�t�Ū�", "�겣�t�Ū�",  null, null);
	menu.addItem("newsid", "�l�q��", "�l�q��",  null, null);
	       
	menu.addSubItem("webmasterid", "�~�Ȥ�X", "�~�Ȥ�X",  "#520000");
	menu.addSubItem("webmasterid", "�s�ڧQ����X--�X�p", "�s�ڧQ����X--�X�p",  "#520100");
	menu.addSubItem("webmasterid", "�ɴڧQ����X--�X�p", "�ɴڧQ����X--�X�p",  "#520200");
	menu.addSubItem("webmasterid", "�~�ȥ~��X", "�~�ȥ~��X",  "#522000");
	menu.addSubItem("webmasterid", "�~�Ȧ��J", "�~�Ȧ��J",  "#420000");	
	menu.addSubItem("webmasterid", "��ڧQ�����J--�X�p", "��ڧQ�����J--�X�p",  "#420100");
	menu.addSubItem("webmasterid", "�s�x�Q�����J--�X�p", "�s�x�Q�����J--�X�p",  "#420300");	
	menu.addSubItem("webmasterid", "�~�ȥ~���J", "�~�ȥ~���J",  "#422000");
	
	menu.addSubItem("newsid", "�n���s��", "�n���s��",  "http://www.tacocity.com.tw/ppkoabcd/good.htm");
	menu.addSubItem("newsid", "�K�O�귽", "�K�O�귽",  "http://www.tacocity.com.tw/ppkoabcd/free.htm");
	menu.addSubItem("newsid", "�䳣�@��C", "�䳣�@��C",  "http://www.tacocity.com.tw/ppkoabcd/touristone.htm");
	menu.addSubItem("newsid", "�������", "�������",  "http://www.tacocity.com.tw/ppkoabcd/hotel.htm");
	
	menu.addSubItem("freedownloadid", "�x�W���� NO.1", "�x�W���� NO.1",  "http://ftp.world.net.tw/~pr/home.html");
	
		menu.showMenu();
}