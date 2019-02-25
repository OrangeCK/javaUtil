package com.sf.channelexpand.util;

import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Date;

import org.apache.commons.lang.StringUtils;

/**
 * Title: ApiUtil.java  
 * Description: 对外接口工具
 * Copyright: Copyright (c) 2018
 * @author Kang Chen  
 * @date 2018年12月5日 下午3:43:16
 * @version 1.0  
 */
public class ApiUtil {
	/**
	 * 时间戳误差范围
	 */
	private static final long timeErrorRange = 5*60*1000;
	/**
	 * 加密为校验码
	 * @param appCode 接入编码
	 * @param appKey 接入秘钥
	 * @param timestamp 接入时间戳
	 * @return 返回校验码
	 */
	public static String encryptToCheckWord(String appCode, String appKey, String timestamp){
		long apiTimestamp = Long.parseLong(timestamp);
		// 获取当前时间戳
		long currTimestamp = new Date().getTime();
//		System.out.println("当前时间戳" + currTimestamp + "时间戳误差范围" + Math.abs(currTimestamp - apiTimestamp));
		// 判断时间戳是否在误差范围之内
		timestamp = (Math.abs(currTimestamp - apiTimestamp) <= timeErrorRange) ? timestamp : null;
		if(StringUtils.isNotEmpty(timestamp)){
			String checkWord = appCode + appKey + timestamp;
			return MD5Util.encryptToMD5(checkWord);
		}else{
			return timestamp;
		}
	}
	
	/**
	 * 调用http连接
	 * @param connectUrl 链接地址
	 * @param param 传入的参数
	 * @return 返回的数据
	 */
	public static String connectUrlByHttp(String connectUrl, String param){
		String result = "";
        DataOutputStream dataout = null;
        BufferedReader br = null;
		try {
			URL url = new URL(connectUrl);
			// 将url 以 open方法返回的urlConnection  连接强转为HttpURLConnection连接  (标识一个url所引用的远程对象连接)
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            // 设置连接输出流为true,默认false (post 请求是以流的方式隐式的传递参数)
            connection.setDoOutput(true);
            // 设置连接输入流为true
            connection.setDoInput(true);
            // 设置请求方式为post
            connection.setRequestMethod("POST");
            // post请求缓存设为false
            connection.setUseCaches(false);
            // 设置该HttpURLConnection实例是否自动执行重定向
            connection.setInstanceFollowRedirects(true);
            // 设置请求头里面的各个属性 (以下为设置内容的类型,设置为经过urlEncoded编码过的form参数) application/x-www-form-urlencoded->表单数据
            connection.setRequestProperty("Content-Type", "text/xml;charset=utf-8");
            // 建立连接 (请求未开始,直到connection.getInputStream()方法调用时才发起,以上各个参数设置需在此方法之前进行)
            connection.connect();
            // 创建输入输出流,用于往连接里面输出携带的参数
            dataout = new DataOutputStream(connection.getOutputStream());
            // 将参数输出到连接
            dataout.writeBytes(param);
            // 输出完成后刷新并关闭流
            dataout.flush();
            // 连接发起请求,处理服务器响应  (从连接获取到输入流并包装为bufferedReader)
            br = new BufferedReader(new InputStreamReader(connection.getInputStream(), "UTF-8"));
            String line;
            StringBuilder sb = new StringBuilder();
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            // 销毁连接
            connection.disconnect();
            result = sb.toString();
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if(dataout != null){
					dataout.close();	
				}
				if(br != null){
					br.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return result;
	}
	
	/**
	 * 返回错误的xml报文
	 * @param msg
	 * @return
	 */
	public static String returnErrorXml(String msg){
		StringBuilder sb = new StringBuilder();
		sb.append("<?xml version='1.0' encoding='UTF-8'?>");
        sb.append("<Response service='OrderService'>");
        sb.append("<Head>ERR</Head>");
        sb.append("<ERROR code='CCOP'>" + msg + "</ERROR>");
        sb.append("</Response>");
        return sb.toString();		
	}
	
	public static void main(String[] args) {
		String appCode = "CK";
		String appKey = "CkLimin91013102";
		String timestamp = "1545042712000";
		String encryptWord = encryptToCheckWord(appCode, appKey, timestamp);
		System.out.println("校验码是" + encryptWord);
	}
	
}
