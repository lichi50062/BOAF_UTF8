//106.10.23 add reLogin.jsp by 2295
//108.06.13 add 過濾字元union、select、delete by 2295
//108.10.28 add 針對printStyle過濾換行符號(\n、%0a) by 2295
//110.09.09 add printStyle過濾(\r) by 2295
//111.02.15 add zz005w by 2295
package com.tradevan.util;

import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpServletRequestWrapper;


public class encodingFilter implements Filter {

	/* (non-Javadoc)
	 * @see javax.servlet.Filter#doFilter(javax.servlet.ServletRequest, javax.servlet.ServletResponse, javax.servlet.FilterChain)
	 */
	String encoding="UTF-8";
	
	public void doFilter(
		ServletRequest arg0,
		ServletResponse arg1,
		FilterChain arg2)
		throws IOException, ServletException {
		// TODO Auto-generated method stub
		//System.out.println("----------------filter xxxx-------------");
		arg0.setCharacterEncoding(encoding);
		//System.out.println("CharacterEncoding="+arg0.getCharacterEncoding());
		//arg2.doFilter(arg0, arg1);
		HttpServletRequest uRequest = (HttpServletRequest)arg0;
        HttpServletResponse uResponse = (HttpServletResponse)arg1;
        String sId = uRequest.getSession().getId();        
        uResponse.setHeader("SET-COOKIE","JSESSIONID="+sId+"; httpOnly; secure");//105.04.06 add
        uResponse.setHeader("Strict-Transport-Security", "max-age=31536000");//110.09.16 add
        uResponse.addHeader("X-FRAME-OPTIONS", "SAMEORIGIN");//106.05.03 add
        //uResponse.setHeader("Content-Security-Policy", "default-src 'self'");//110.11.03 add無法新增,IE可登入,但Chrome/Edge無法登入
        
		   arg2.doFilter(new XSSRequestWarpper((HttpServletRequest)arg0), arg1);
        
	}

	/* (non-Javadoc)
	 * @see javax.servlet.Filter#getFilterConfig()
	 */
	public FilterConfig getFilterConfig() {
		// TODO Auto-generated method stub
		System.out.println("-------------filter yyyyy---------");
		return null;
	}

	/* (non-Javadoc)
	 * @see javax.servlet.Filter#setFilterConfig(javax.servlet.FilterConfig)
	 */
	public void setFilterConfig(FilterConfig arg0) {
		// TODO Auto-generated method stub
		//System.out.println("-------------filter zzzz---------");
		encoding=arg0.getInitParameter("encoding");
	}
	public void init(FilterConfig arg0){
		//System.out.println("-------------filter oooo---------");
		encoding=arg0.getInitParameter("encoding");
	}
	public void destroy(){		
	}
}
class XSSRequestWarpper extends HttpServletRequestWrapper {
	private String newPath;

	public XSSRequestWarpper(HttpServletRequest uesrequestt) {
		super(uesrequestt);
		//System.out.println("XSSRequestWarpper");		
		// TODO Auto-generated constructor stub
	}

	public String[] getParameterValues(String parameter) {

		String[] values = super.getParameterValues(parameter);
		if (values == null) {
			return null;
		}
		int count = values.length;
		String[] encodedValues = new String[count];
		for (int i = 0; i < count; i++) {
			encodedValues[i] = cleanXSS(values[i]);
		}
		return encodedValues;
	}

	public String getParameter(String parameter) {
		String value = super.getParameter(parameter);
		if (value == null) {
			return null;
		}
		HttpServletRequest uRequest = (HttpServletRequest)super.getRequest();
       	
		String sourceIp=uRequest.getRemoteAddr();
        //System.out.println("sourceIp : "+sourceIp  );
        //System.out.println("RemoteHost : "+uRequest.getRemoteHost()  );
        String path = uRequest.getPathInfo();
        if (path == null) {
            path = uRequest.getServletPath();
        }
        //System.out.println("path :"+path);
        
        if (path.startsWith("/pages/ZZ042W.jsp") || path.startsWith("/pages/reLogin.jsp") || path.startsWith("/pages/ZZ005W.jsp")
        ||(path.startsWith("/pages/debug/getdb_data.jsp") ||path.startsWith("/pages/debug/updatedb_data.jsp")                
        && (sourceIp.startsWith("127.0.0.1") || sourceIp.startsWith("202.173.49") || sourceIp.startsWith("202.173.43")))) 
        {
           return value;
        }else{
           //return value;  
           //System.out.println("parameter="+parameter); 
           if(parameter.equals("printStyle")){//針對printStyle過濾換行符號(\n、%0a)
        	   return printStyle_cleanXSS(value);
           }
		   return cleanXSS(value);
        }
	}
	
	private String cleanXSS(String value) {  
	   
	    value=value.replaceAll("<", "& lt;");
        value=value.replaceAll(">", "& gt;");
        value=value.replaceAll("//", "& gt;");
        value=value.replaceAll("'", "& 39;");//單引號
        //value = value.replaceAll("\\(", "& 40;").replaceAll("\\)", "& 41;");//局內資料有  
        value=value.replaceAll("eval\\((.*)\\)", "");
        value=value.replaceAll("\\?", "？"); //106.04.17 add       
        value=value.replaceAll("[\\\"\\\'][\\s]*javascript:(.*)[\\\"\\\']", "\"\"");  
        //value = value.replaceAll("\"","&quot;") ;//雙引號 //106.04.17 add//111.01.17 fix
        //value=value.replaceAll("+", "＋");//106.04.17 add //局內資料有  
        //value=value.replaceAll(";", "；");//106.04.17 add //局內資料有  
        //value=value.replaceAll("-", "－");  //106.04.17 add //局內資料有  
        //value = value.replaceAll("(", "（");//局內資料有  
        //value = value.replaceAll(")", "）");//局內資料有  
        value = value.replaceAll("script", "");        
        
        value = value.replace("..", "");//107.05.23 add//..如果寫成replaceAll所有字串會都清成空白//暫時拿掉
        value = value.replaceAll("%2", "& gt;");//107.05.23 add       
        
        value = value.replaceAll("union", "& gt;");//108.06.13 add     
        value = value.replaceAll("select", "& gt;");//108.06.13 add
        value = value.replaceAll("delete", "& gt;");//108.06.13 add
              
        return value;  
    }
	private String printStyle_cleanXSS(String value) {
		//針對printStyle過濾換行符號(\n、%0a)//108.10.28 add
		value=value.replaceAll("\n", "");
		value=value.replaceAll("\r", "");//110.09.09 add
	    value=value.replaceAll("%0a", "");
		return  value;
	}
}