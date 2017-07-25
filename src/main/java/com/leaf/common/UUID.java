package com.leaf.common;

public class UUID {
	
	public static String getUUID(){
		return java.util.UUID.randomUUID().toString().replaceAll("-", "").toUpperCase();
	}


}
