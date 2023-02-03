function doSubmitWM005W(form) {

	if (!checkSingleYM(form.M_YEAR, form.M_MONTH))
		return false;

	form.action = 'WM005W_ShowList';
	form.submit();

}
