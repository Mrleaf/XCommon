package com.leaf.common;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * 作者 leaf
 * 时间 2017年5月15日上午11:45:42
 * 校验工具
 */
public class ValidUtil {
	/**
	 * 对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmpty(Object value) {
		return value == null;
	}

	/**
	 * 字符串对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmpty(String value) {
		return value == null || value.isEmpty();
	}

	/**
	 * 字符串对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmptyStr(String value) {
		return value == null || value.isEmpty() || value.trim().equals("") || value.equals(" ") || value.equals("[]") || value.equals("null") || value.equals("undefined");
	}

	/**
	 * 整形对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmpty(Integer value) {
		return value == null;
	}

	/**
	 * 长整形对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmpty(Long value) {
		return value == null || value == 0;
	}

	public static boolean isNullOrEmpty(Short value) {
		return value == null;
	}
	/**
	 * 布尔对象的非空校验
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNullOrEmpty(Boolean value) {
		return value == null || !value;
	}

	/**
	 * 数组的非空校验
	 *
	 * @param array
	 * @return
	 */
	public static <T extends Object> boolean isNullOrEmpty(T[] array) {
		return array == null || array.length == 0;
	}

	/**
	 * 对象集合的非空校验
	 *
	 * @param list
	 * @return
	 */
	public static boolean isNullOrEmpty(List<? extends Object> list) {
		return list == null || list.isEmpty();
	}

	/**
	 * 对象键值对集合的非空校验
	 *
	 * @param map
	 * @return
	 */
	public static boolean isNullOrEmpty(Map<? extends Object, ? extends Object> map) {
		return map == null || map.size() == 0;
	}

	/**
	 * 验证字符串是否可转换为数字(整型)
	 *
	 * @param string
	 * @return boolean
	 */

	public static boolean isAIntString(String string) {
		try {
			Integer.parseInt(string);
			return true;
		} catch (NumberFormatException e) {
			return false;
		}
	}

	/**
	 * 校验字段是否为空
	 * @param obj
	 * @param fields
	 * @return
	 */
	public static boolean isNotNull(Object obj, String ... fields){
		for(String field : fields){
			try {
				Field f = obj.getClass().getDeclaredField(field);
				f.setAccessible(true);
				Object value = f.get(obj);
				if(value == null){
					return false;
				}
			} catch (NoSuchFieldException e) {
				throw new RuntimeException("NoSuchFieldException: "+field, e);
			} catch (IllegalAccessException e) {
				throw new RuntimeException("IllegalAccessException: "+field, e);
			}
		}
		return true;
	}

}
