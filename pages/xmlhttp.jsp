<%!
    public boolean isNaN(String O) {
        if (O==null || O.equals("null")) {
            return true;
        } else {
            return false;
        }
    }
%>
<%
    String a = request.getParameter("a");
    String b = request.getParameter("b");
    String total;
    if (isNaN(a) || isNaN(b)) {
        a="";
        b="";
        total="";
    } else {
        total=Integer.toString((Integer.parseInt(a)+Integer.parseInt(b)));
    }

    String acc = request.getHeader("Accept");
    System.out.println("acc=" + acc);
    if (acc.indexOf("message/x-jl-formresult")!=-1) {
        out.println(total);
        //out.println("<table border=1><tr><td>Hello</td></tr></table>");
    } else {
%>
<script src="js/xmlhttp.js" type="text/javascript"></script>
<script>
function calc() {
    frm=document.forms[0];
    url="xml.jsp?a="+frm.elements['a'].value+"&b="+frm.elements['b'].value;
    xmlhttp.open("POST",url,true);
    xmlhttp.onreadystatechange=function() {
                if (xmlhttp.readyState==4) {
                    document.forms[0].elements['total'].value=xmlhttp.responseText;
                    msg.innerHTML=xmlhttp.responseText;
                }
            }
    xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
    xmlhttp.send(null);
    return false;
}
</script>
<form action="/pages/xmlhttp.jsp" method="get" onsubmit="return calc();">
<input type=text name="a" value="<%=a%>"> + <input type=text name="b" value="<%=b%>">
 = <input type=text name="total" value="<%=total%>">
<input type="submit" value="Calculate">
</form>
<div id="msg">
Hello World
</div>
<%
    }
%>