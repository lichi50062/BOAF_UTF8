function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "�겣�t�Ū�(�겣)", "�겣�t�Ū�(�겣)",  null, null);
	menu.addItem("A01Div01right", "�겣�t�Ū�(�t��)", "�겣�t�Ū�(�t��)",  null, null);	
	menu.addItem("A01Div03", "�O����ڪ��B", "�O����ڪ��B", "#190000" , "#190000");
	menu.addItem("A01Div02", "�l�q��", "�l�q��",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "�y�ʸ겣", "�y�ʸ겣",  "#110000");
	menu.addSubItem("A01Div01left", "�s��X�w--�p�p", "�s��X�w--�p�p",  "#110307");
	menu.addSubItem("A01Div01left", "�s��A�~���w--�p�p", "�s��A�~���w--�p�p",  "#110327");	
	menu.addSubItem("A01Div01left", "�s���w--�X�p", "�s���w--�X�p",  "#110313");
	
	menu.addSubItem("A01Div01left", "���", "���",  "#112500");
	menu.addSubItem("A01Div01left", "�A�~������--�p�p", "�A�~������--�p�p",  "#120604");
	menu.addSubItem("A01Div01left", "����ΥX��", "����ΥX��",  "#150300");
	menu.addSubItem("A01Div01left", "�T�w�겣", "�T�w�겣",  "#130300");
	menu.addSubItem("A01Div01left", "��L�겣", "��L�겣",  "#141600");
	menu.addSubItem("A01Div01left", "����", "����",  "#152200");
	
	
	menu.addSubItem("A01Div01right", "�y�ʭt��", "�y�ʭt��",  "#100000");
	menu.addSubItem("A01Div01right", "�z���w--�X�p", "�z���w--�X�p",  "#210205");
	 
	menu.addSubItem("A01Div01right", "�u���ɴ�--�X�p", "�u���ɴ�--�X�p",  "#210405");
	menu.addSubItem("A01Div01right", "�s��", "�s��",  "#211100");
	menu.addSubItem("A01Div01right", "�����t��", "�����t��",  "#221000");
	menu.addSubItem("A01Div01right", "�����ɴ�--�X�p", "�����ɴ�--�X�p",  "#240105");
	menu.addSubItem("A01Div01right", "�ɤJ�A�o���-�A��-�p�p", "�ɤJ�A�o���-�A��-�p�p",  "#240215");
	menu.addSubItem("A01Div01right", "�ɤJ�A�o�����ڸ��--�X�p", "�ɤJ�A�o�����ڸ��--�X�p",  "#240240");
	menu.addSubItem("A01Div01right", "�ɤJ�M�ש�ڸ��--�X�p", "�ɤJ�M�ש�ڸ��--�X�p",  "#240305");
	menu.addSubItem("A01Div01right", "��L�t��", "��L�t��",  "#240400");
	 
	menu.addSubItem("A01Div01right", "����", "����",  "#251800");
	menu.addSubItem("A01Div01right", "�Ʒ~����Τ��n", "�Ʒ~����Τ��n",  "#260600");
    menu.addSubItem("A01Div01right", "�����ηl�q", "�����ηl�q",  "#310900");
    menu.addSubItem("A01Div01right", "�t�Ťβb��", "�t�Ťβb��",  "#320300");
	 
    menu.addSubItem("A01Div02", "�~�Ȥ�X", "�~�Ȥ�X",  "#522700");		
	menu.addSubItem("A01Div02", "�~�Ȧ��J", "�~�Ȧ��J",  "#422700");	
    menu.addSubItem("A01Div02", "�X�p", "�X�p",  "#420000");
     
	
	menu.showMenu();
}