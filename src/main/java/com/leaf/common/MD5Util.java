package com.leaf.common;

import sun.misc.BASE64Encoder;

import java.security.MessageDigest;

public class MD5Util {
	
	/**
	 * MD5+BASE64
	 * @param str
	 * @return
	 * @throws Exception
	 */
	public static String encoderByMd5(String str){
		String newstr = null;
		try {
			MessageDigest md5=MessageDigest.getInstance("MD5");
	        BASE64Encoder base64en = new BASE64Encoder();
	        newstr = base64en.encode(md5.digest(str.getBytes("utf-8")));
	        return newstr;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return newstr;
        
    }
	
	/**
	 * 验证
	 * @param newpasswd
	 * @param oldpasswd
	 * @return
	 * @throws Exception
	 */
	public static boolean checkpassword(String newpasswd,String oldpasswd){
        if(encoderByMd5(newpasswd).equals(oldpasswd))
            return true;
        else
            return false;
    }
	
	
}
