function showToolbar()
{
// AddItem(id, text, hint, location, alternativeLocation);
// AddSubItem(idParent, text, hint, location);

	menu = new Menu();
	
	menu.addItem("A01Div01left", "�k�w��v���", "�k�w��v���",  null, null);
	    
	
	menu.addSubItem("A01Div01left", "1.�륭������`�B", "1.�륭������`�B ",  "#990110a");
	menu.addSubItem("A01Div01left", "2.�����ĸ�l�B", "2.�����ĸ�l�B ", "#990150");
	menu.addSubItem("A01Div01left", "3.�D�|���s���`�B", "3.�D�|���s���`�B", "#990240");
	menu.addSubItem("A01Div01left", "4.�٧U�|���«H�`�B", "4.�٧U�|���«H�`�B", "#990320");
	menu.addSubItem("A01Div01left", "5.�D�|���L��O���O�ʶU��", "5.�D�|���L��O���O�ʶU��", "#990420");
	menu.addSubItem("A01Div01left", "6.�D�|���«H�`�B", "6.�D�|���«H�`�B", "#990512");
	menu.addSubItem("A01Div01left", "7.�ۥΦ�v����`�B", "7.�ۥΦ�v����`�B", "#990630");
	menu.addSubItem("A01Div01left", "8.�H�γ��T�w�겣�b�B", "8.�H�γ��T�w�겣�b�B", "#990720");
	menu.addSubItem("A01Div01left", "9.�~���겣�`�B(��X�s�x��)", "9.�~���겣�`�B(��X�s�x��)", "#990810");
	menu.addSubItem("A01Div01left", "10.��t�d�H�B�U�������u�λP��t�d�H�ο�z�«H��¾�����Q�`���Y�̾�O��ڳ̰����B", "10.��t�d�H�B�U�������u�λP��t�d�H�ο�z�«H��¾�����Q�`���Y�̾�O��ڳ̰����B", "#990930");
	menu.addSubItem("A01Div01left", "11.���|���]�t�P��a�ݡ^�ΦP�@���Y�H��ڤ��̰����B", "11.���|���]�t�P��a�ݡ^�ΦP�@���Y�H��ڤ��̰����B", "#991020");
	menu.addSubItem("A01Div01left", "12.�C�@�٧U�|���ΦP�@���Y�H��ڤ��̰����B", "12.�C�@�٧U�|���ΦP�@���Y�H��ڤ��̰����B", "#991120");
	menu.addSubItem("A01Div01left", "13.��P�@�D�|���ΦP�@���Y�H��ڤ��̰����B", "13.��P�@�D�|���ΦP�@���Y�H��ڤ��̰����B", "#991220");
	menu.addSubItem("A01Div01left", "14.�̪�M��~��", "14.�̪�M��~��", "#991320");
	menu.showMenu();
}