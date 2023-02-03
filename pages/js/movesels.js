//移動選取的項目
function movesel(source,target)
{
    for(i=0;i<source.length;i++){
    	if(source.options[i] != null && source.options[i].selected==true){
    		if(target.options[0]!=null && target.options[0].value=="null")
    			target.options[0] = null;
    		target.options[target.options.length] = new Option(source.options[i].text,source.options[i].value);
    		source.options[i] = null;
    		if(source.options[0]!=null && source.options[0].value=="null")
    			source.options[0] = null;
    		i-=1;
    	}
    }
}

//移動全部
function moveallsel(source,target)
{
    for(i=0;i<source.length;i++){
    	source.options[i].selected=true;
    }
    movesel(source,target);
}
//該項目上移
function moveup(sel){
	for(i=0;i<sel.length;i++){
		if(sel.options[i].selected==true && i-1!=-1){
			swapopt(sel,i,i-1);
		}
	}
}
//該項目下移
function movedown(sel){
	for(i=sel.length-1;i>=0;i--){
		if(sel.options[i].selected==true && (i+1 <= sel.length-1)){
			swapopt(sel,i,i+1);
		}
	}
}
	
function swapopt(sel,preindex,lasindex){
	text1=sel.options[preindex].text;
	value1=sel.options[preindex].value;
	sel.options[preindex].text=sel.options[lasindex].text;
	sel.options[preindex].value=sel.options[lasindex].value;
	sel.options[lasindex].text=text1;
	sel.options[lasindex].value=value1;
	sel.options[preindex].selected=false;
	sel.options[lasindex].selected=true;
}