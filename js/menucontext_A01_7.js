//94.10.21 fix �ҥHlink.�������item���e�@��,�~���|�Qindex�\�� by 2295
function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "�겣�t�Ū�(�겣)", "�겣�t�Ū�(�겣)",  null, null);
	menu.addItem("A01Div01right", "�겣�t�Ū�(�t��)", "�겣�t�Ū�(�t��)",  null, null);	
	menu.addItem("A01Div03", "�O����ڪ��B", "�O����ڪ��B", "#100000" , "#100000");
	menu.addItem("A01Div02", "�l�q��", "�l�q��",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "�y�ʸ겣", "�y�ʸ겣",  "#110000");
	menu.addSubItem("A01Div01left", "�s��X�w--�p�p", "�s��X�w--�p�p",  "#110216");
	menu.addSubItem("A01Div01left", "�s��A�~���w--�p�p", "�s��A�~���w--�p�p",  "#110256");	
	menu.addSubItem("A01Div01left", "�s���w--�X�p", "�s���w--�X�p",  "#110240");
	
	menu.addSubItem("A01Div01left", "���", "���",  "#112400");
	menu.addSubItem("A01Div01left", "�A�~������--�p�p", "�A�~������--�p�p",  "#120604");
	menu.addSubItem("A01Div01left", "����ΥX��", "����ΥX��",  "#150300");
	menu.addSubItem("A01Div01left", "�T�w�겣", "�T�w�겣",  "#130300");
	menu.addSubItem("A01Div01left", "��L�겣", "��L�겣",  "#141200");
	menu.addSubItem("A01Div01left", "����", "����",  "#152000");
	
	
	menu.addSubItem("A01Div01right", "�y�ʭt��", "�y�ʭt��",  "#990000");
	menu.addSubItem("A01Div01right", "�z���w--�X�p", "�z���w--�X�p",  "#210105");
	 
	menu.addSubItem("A01Div01right", "�u���ɴ�--�X�p", "�u���ɴ�--�X�p",  "#210205");
	menu.addSubItem("A01Div01right", "�s��", "�s��",  "#211200");
	menu.addSubItem("A01Div01right", "�����t��", "�����t��",  "#221000");
	menu.addSubItem("A01Div01right", "�����ɴ�--�X�p", "�����ɴ�--�X�p",  "#240105");
	menu.addSubItem("A01Div01right", "�ɤJ�M�ש�ڸ��--�X�p", "�ɤJ�M�ש�ڸ��--�X�p",  "#240205");
	menu.addSubItem("A01Div01right", "�ɤJ�A�o���-�A��-�p�p", "�ɤJ�A�o���-�A��-�p�p",  "#240315");
	menu.addSubItem("A01Div01right", "�ɤJ�A�o�����ڸ��--�X�p", "�ɤJ�A�o�����ڸ��--�X�p",  "#240340");
	menu.addSubItem("A01Div01right", "��L�t��", "��L�t��",  "#240400");
	 
	menu.addSubItem("A01Div01right", "����", "����",  "#251300");
	menu.addSubItem("A01Div01right", "�t�ŦX�p", "�t�ŦX�p",  "#260700");
	menu.addSubItem("A01Div01right", "�Ʒ~����Τ��n", "�Ʒ~����Τ��n",  "#200000");
    menu.addSubItem("A01Div01right", "�����ηl�q", "�����ηl�q",  "#310900");
    menu.addSubItem("A01Div01right", "�b�ȦX�p", "�b�ȦX�p",  "#320300");
    menu.addSubItem("A01Div01right", "�t�Ťβb�ȦX�p", "�t�Ťβb�ȦX�p",  "#300000");
	 
    menu.addSubItem("A01Div02", "�~�Ȥ�X", "�~�Ȥ�X",  "#600000");		
	menu.addSubItem("A01Div02", "�~�Ȧ��J", "�~�Ȧ��J",  "#500000");	
    //menu.addSubItem("A01Div02", "�X�p", "�X�p",  "#420000");
     
	
	menu.showMenu();
}