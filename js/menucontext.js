function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01", "�겣�t�Ū�", "�겣�t�Ū�",  null, null);
	menu.addItem("A01Div02", "�l�q��", "�l�q��",  null, null);
	    
	
	menu.addSubItem("A01Div01", "�y�ʸ겣", "�y�ʸ겣",  "#110000");
	menu.addSubItem("A01Div01", "�s���w--�X�@���w--�p�p", "�s���w--�X�@���w--�p�p",  "#110000");
	menu.addSubItem("A01Div01", "�s���w--����A�~���w--�p�p", "�s���w--����A�~���w--�p�p",  "#110310");	
	menu.addSubItem("A01Div01", "�s���w--�X�p", "�s���w--�X�p",  "#110320");
	
	menu.addSubItem("A01Div01", "���", "���",  "#111400");
	menu.addSubItem("A01Div01", "�A�~�o�i�����--�p�p", "�A�~�o�i�����--�p�p",  "#120502");
	menu.addSubItem("A01Div01", "����ΥX��", "����ΥX��",  "#150300");
	menu.addSubItem("A01Div01", "�T�w�겣", "�T�w�겣",  "#130300");
	menu.addSubItem("A01Div01", "��L�겣", "��L�겣",  "#141600");
	menu.addSubItem("A01Div01", "����", "����",  "#160500");
	menu.addSubItem("A01Div01", "�겣�X�p", "�겣�X�p",  "#160600");
	menu.addSubItem("A01Div01", "�y�ʭt��", "�y�ʭt��",  "#100000");
	menu.addSubItem("A01Div01", "�z���w--�p�p", "�z���w--�p�p",  "#210000");
	 
	menu.addSubItem("A01Div01", "�u���ɴ�--�p�p", "�u���ɴ�--�p�p",  "#210200");
	menu.addSubItem("A01Div01", "�s��", "�s��",  "#211100");
	menu.addSubItem("A01Div01", "�����t��", "�����t��",  "#221000");
	menu.addSubItem("A01Div01", "�����ɴ�--�X�p", "�����ɴ�--�X�p",  "#240000");
	menu.addSubItem("A01Div01", "�ɤJ�A�~�o�i�����ڸ��--�ɤJ�A�ة�ڸ��--�p�p", "�ɤJ�A�~�o�i�����ڸ��--�ɤJ�A�ة�ڸ��--�p�p",  "#240100");
	menu.addSubItem("A01Div01", "�ɤJ�A�~�o�i�����ڸ��--�X�p", "�ɤJ�A�~�o�i�����ڸ��--�X�p",  "#240210");
	menu.addSubItem("A01Div01", "�ɤJ�M�ש�ڸ��--�X�p", "�ɤJ�M�ש�ڸ��--�X�p",  "#240200");
	menu.addSubItem("A01Div01", "��L�t��", "��L�t��",  "#240400");
	 
	menu.addSubItem("A01Div01", "����", "����",  "#260500");
	menu.addSubItem("A01Div01", "�t�ŦX�p", "�t�ŦX�p",  "#260600");
	menu.addSubItem("A01Div01", "�Ʒ~����Τ��n", "�Ʒ~����Τ��n",  "#200000");
	menu.addSubItem("A01Div01", "�����ηl�q", "�����ηl�q",  "#310900");
	menu.addSubItem("A01Div01", "�b�ȦX�p", "�b�ȦX�p",  "#320300");
	menu.addSubItem("A01Div01", "�t�Ťβb�ȦX�p", "�t�Ťβb�ȦX�p",  "#300000");
	
	
	
	 
	 
    menu.addSubItem("A01Div02", "�~�Ȥ�X", "�~�Ȥ�X",  "#522700");
	menu.addSubItem("A01Div02", "�s�ڧQ����X--�X�p", "�s�ڧQ����X--�X�p",  "#520000");
	menu.addSubItem("A01Div02", "�ɴڧQ����X--�X�p", "�ɴڧQ����X--�X�p",  "#520100");
	menu.addSubItem("A01Div02", "�~�ȥ~��X", "�~�ȥ~��X",  "#521900");
	menu.addSubItem("A01Div02", "�~�Ȧ��J", "�~�Ȧ��J",  "#500000");	
	menu.addSubItem("A01Div02", "��ڧQ�����J--�X�p", "��ڧQ�����J--�X�p",  "#500000");
	menu.addSubItem("A01Div02", "�s�x�Q�����J--�X�p", "�s�x�Q�����J--�X�p",  "#420100");	
	menu.addSubItem("A01Div02", "�~�ȥ~���J", "�~�ȥ~���J",  "#420800");	
     
     
	
	menu.showMenu();
}