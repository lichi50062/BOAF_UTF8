function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "�H�γ��b�ȥe���I�ʸ겣��v�p���", "�H�γ��b�ȥe���I�ʸ겣��v�p���",  null, null);
	menu.addItem("A01Div01right", "�H�γ��b�Ȧ����I�ʸ�Ƥ�v", "�H�γ��b�Ȧ����I�ʸ�Ƥ�v",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "�{��", "�{��",  "#A05_div01");
	menu.addSubItem("A01Div01left", "�W�C�H�~�̳W�w�A���I�v�ƥ��F100%���겣", "�W�C�H�~�̳W�w�A���I�v�ƥ��F100%���겣", "#920601");	
	menu.addSubItem("A01Div01left", "�W�C�H�~�����v�Ψ�L�겣", "�W�C�H�~�����v�Ψ�L�겣",  "#92075P");
	
	menu.addSubItem("A01Div01right", "�Ĥ@���ꥻ", "�Ĥ@���ꥻ",  "#A05_div02");	
	menu.addSubItem("A01Div01right", "�ĤG���ꥻ", "�ĤG���ꥻ",  "#910199");	
	menu.addSubItem("A01Div01right", "�b���`�B", "�b���`�B",  "#910299");	
	menu.addSubItem("A01Div01right", "�", "�",  "#910300");	
	menu.addSubItem("A01Div01right", "�X��b��", "�X��b��",  "#910403");	
	menu.addSubItem("A01Div01right", "���I�ʸ겣�`�B", "���I�ʸ겣�`�B",  "#910400");	
	menu.addSubItem("A01Div01right", "�H�γ��b�ȥe���I�ʸ겣��v", "�H�γ��b�ȥe���I�ʸ겣��v",  "#910500");	
	
	menu.showMenu();
}