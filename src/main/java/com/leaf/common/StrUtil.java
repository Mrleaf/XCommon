package com.leaf.common;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

/**
 * 作者 leaf
 * 时间 2017年5月15日下午7:45:43
 */
public class StrUtil {



	public static Object nullStringConverNull(Object obj)throws Exception{
		Class<?> objClass = obj.getClass();
//		Object bean = objClass.newInstance();
		Field[] fs = objClass.getDeclaredFields();
		for(int i = 0 ; i < fs.length; i++){
			Field f = fs[i];
			f.setAccessible(true); //设置些属性是可以访问的
			String type = f.getType().getName();//得到此属性的类型
			Object val = f.get(obj);//得到此属性的值
//			System.out.println("name:"+f.getName()+"\t value = "+val +"\t type = "+type);
			if(type.endsWith("String")&&ValidUtil.isNullOrEmpty((String)val)){
				f.set(obj,null);
			}
		}
		return obj;
	}

	private static String change(String src) {
		if (src != null) {
			StringBuffer sb = new StringBuffer(src);
			sb.setCharAt(0, Character.toUpperCase(sb.charAt(0)));
			return sb.toString();
		} else {
			return null;
		}
	}

	/**
	 * 字符转list
	 * @param str
	 * @param regex
	 * @return
	 */
	public static List<String> str2List(String str,String regex){
		if(ValidUtil.isNullOrEmpty(str))
			return null;
        String[] arrayStr = str.split(regex);
        List<String> list = Arrays.asList(arrayStr);
		List arrayList = new ArrayList(list);
        return arrayList;
	}

	/**
	 * 数组转字符串
	 * @param list
	 * @return
	 */
	public static String list2Str(List<String> list){
		String s = "";
		for(String str :list){
			if(ValidUtil.isNullOrEmpty(s))
				s += str;
			else
				s += ","+str;
		}
		return s;
	}
	/**
	 * 是否是整数
	 * @param b
	 * @return
	 */
	public static boolean isInteger(BigDecimal b){
		if(new BigDecimal(b.intValue()).compareTo(b)==0){
			return true;
		}else{
			return false;
		}
	}

	/**
	 * 随机字母
	 * ASCII 65大写
	 * ASCII 97小写
	 */
	public static String randomChar(int x){
		Random random = new Random();
		return String.valueOf((char)(x + random.nextInt(26)));
	}

	/**
	 * 获取字母
	 * @param str
	 * @return
	 */
    public static String getChar(String str){
		String c = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
		String charStr = "";
		for(int x=0;x<c.length();x++){
			if(ValidUtil.isNullOrEmpty(str)){
				charStr = String.valueOf(c.charAt(x));
				break;
			}else if(!str.contains(String.valueOf(c.charAt(x)))){
				charStr = String.valueOf(c.charAt(x));
				break;
			}
		}
		return charStr;
	}
    
    /**
	 * 字符串是否为空串，包括本身字符串为null以及""
	 *
	 * @author 查畅
	 * @param target
	 *            目标串
	 * @return 目标串是否为空串
	 * @author huangzhen
	 */
	public static boolean isNullString(String target) {
		if (target == null || "".equals(target.trim())) {
			return true;
		}
		return false;
	}
	
	/***
	 * 根据指定的数字字符串,获取自增1后,指定位数lenth长度的字符串(位数不够的用0补充)
	 * 如果自增后的数字大于需要返回的位数,则返回自增后的数字字符窜
	 * 
	 * @author 查畅
	 * @param numStr
	 *            数字字符串
	 * @param lenth
	 *            需要的返回的字符串长度
	 * @return
	 */
	public static String getStrForNumStr(String numStr, int lenth) {
		if (numStr == null || numStr.equals("")) {
			return "";
		}
		Integer num = Integer.parseInt(numStr) + 1;
		String str = num + "";
		if (str.length() >= lenth) {// 如果数字长度大于等于需要返回的长度,则直接返回该数字
			return str;
		} else {// 如果小于需要返回的长度,则用零来补充
			for (int i = str.length(); i < lenth; i++) {
				str = "0" + str;
			}
		}
		return str;
	}
}
