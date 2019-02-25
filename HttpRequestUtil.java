package com.sf.channelexpand.util;

import javax.servlet.http.HttpServletRequest;

/**
 * Title: HttpRequestUtil.java  
 * Description:  
 * Copyright: Copyright (c) 2018
 * @author Kang Chen  
 * @date 2018年9月29日 下午3:28:46
 * @version 1.0  
 */
public class HttpRequestUtil {

	/**
	 * 获取当前请求的IP地址
	 * @param request
	 * @return
	 */
	public static String getIpAddr(HttpServletRequest request) {
		String ip="";
		if (request.getHeader("x-forwarded-for") == null) {
			ip = request.getRemoteAddr();  
	    }else{
	    	ip = request.getHeader("x-forwarded-for");  
	    }
	    if(ip.length()>15){ //"***.***.***.***".length() = 15  
            if(ip.indexOf(",")>0){  
            	ip = ip.substring(0,ip.indexOf(","));  
            }
        }
	return ip;
	}

}
