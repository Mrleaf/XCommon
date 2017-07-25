package com.leaf.common;

import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

/**
 * Created by leaf
 * 时间 2017/5/19 0019 12:04
 */
public class DateUtil {

    private final static String DATE_PATTERN_CN = "yyyy-MM-dd HH:mm:ss";

    private final static String DATE_PATTERN_ZH_CN = "yyyy年MM月dd日 HH时mm分ss秒";

    private final static String DATE_PATTERN_ZH_CN_SHORT = "yyyy年MM月dd日";

    private final static String DATE_PATTERN_CN_SHORT = "yyyy-MM-dd";

    private final static String DATE_PATTERN_CN_TIME = "HH:mm:ss";


    /**
     * 获取当前日期时间字符串
     *
     * @return 返回日期时间字符串(yyyy-MM-dd HH:mm:ss)
     */
    public static String getDateStr() {
        return getDateStr(DATE_PATTERN_CN);
    }

    /**
     * 获取当前日期字符串
     *
     * @return 返回日期字符串(yyyy-MM-dd)
     */
    public static String getShortDateStr() {
        return getDateStr(DATE_PATTERN_CN_SHORT);
    }

    /**
     * 获取当前时间字符串
     *
     * @return 返回时间字符串(HH:mm:ss)
     */
    public static String getTimeStr() {
        return getDateStr(DATE_PATTERN_CN_TIME);
    }

    /**
     * 获取指定格式的日期字符串
     *
     * @param format 日期字符串格式，例如：yyyy-MM-dd HH:mm:ss
     * @return
     */
    public static String getDateStr(String format) {
        SimpleDateFormat formatter = new SimpleDateFormat(format);
        String dateStr = formatter.format(new Date());
        return dateStr;
    }

    /**
     * 转换为日期时间字符串(yyyy-MM-dd HH:mm:ss)
     *
     * @param date 日期对象
     * @return
     */
    public static String toDateStr(Date date) {
        SimpleDateFormat formatter = new SimpleDateFormat(DATE_PATTERN_CN);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * 转换为日期时间字符串(依据format)
     *
     * @param date 日期对象
     * @return
     */
    public static String toDateStr(Date date, String format) {
        SimpleDateFormat formatter = new SimpleDateFormat(format);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * 转换为日期字符串(yyyy-MM-dd)
     *
     * @param date 日期对象
     * @return
     */
    public static String toShortDateStr(Date date) {
        SimpleDateFormat formatter = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * 转换为日期字符串(yyyy年MM月dd日 HH时mm分ss秒)
     *
     * @param date 日期对象
     * @return
     */
    public static String toCnDateStr(Date date) {
        SimpleDateFormat formatter = new SimpleDateFormat(DATE_PATTERN_ZH_CN);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * 转换为日期字符串(yyyy年MM月dd日)
     *
     * @param date 日期对象
     * @return
     */
    public static String toCnShortDateStr(Date date) {
        SimpleDateFormat formatter = new SimpleDateFormat(DATE_PATTERN_ZH_CN_SHORT);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * 转换为日期对象
     *
     * @param date 长时间格式字符串(yyyy-MM-dd HH:mm:ss)
     * @return
     */
    public static Date toDate(String date) {
        return toDate(date, DATE_PATTERN_CN);
    }

    /**
     * 转换为日期对象
     *
     * @param date 短时间格式字符串(yyyy-MM-dd)
     * @return
     */
    public static Date toShortDate(String date) {
        return toDate(date, DATE_PATTERN_CN_SHORT);
    }

    /**
     * 转换为日期对象
     *
     * @param date
     * @param format
     * @return
     */
    public static Date toDate(String date, String format) {
        SimpleDateFormat formatter = new SimpleDateFormat(format);
        ParsePosition pos = new ParsePosition(0);
        return formatter.parse(date, pos);
    }

    /**
     * 计算指定天数间隔的日期字符串
     */
    public static String getDateStrByDays(String nowdate, String days) {
        try {
            SimpleDateFormat format = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
            String mdate = "";
            Date d = toDate(nowdate);
            long myTime = (d.getTime() / 1000) + Integer.parseInt(days) * 24 * 60 * 60L;
            d.setTime(myTime * 1000);
            mdate = format.format(d);
            return mdate;
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 计算指定分钟间隔的日期字符串
     *
     * @param dateStr
     * @return
     */
    public static String getDateStrByMinutes(String dateStr, int minutes) {
        SimpleDateFormat format = new SimpleDateFormat(DATE_PATTERN_CN);
        try {
            Date date = format.parse(dateStr);
            long seconds = (date.getTime() / 1000) + minutes * 60L;
            date.setTime(seconds * 1000);
            return format.format(date);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 计算指定分钟间隔的日期
     *
     * @param date
     * @param minutes
     * @return
     */
    public static Date getDateByMinutes(Date date, int minutes) {
        try {
            long seconds = (date.getTime() / 1000) + minutes * 60L;
            date.setTime(seconds * 1000);
        } catch (Exception e) {
            return null;
        }
        return date;
    }

    /**
     * 两个时间之间的天数
     *
     * @param date1
     * @param date2
     * @return
     */
    public static long getDays(String date1, String date2) {
        if (date1 == null || date1.equals(""))
            return 0;
        if (date2 == null || date2.equals(""))
            return 0;
        // 转换为标准时间
        SimpleDateFormat myFormatter = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
        Date date = null;
        Date mydate = null;
        try {
            date = myFormatter.parse(date1);
            mydate = myFormatter.parse(date2);
        } catch (Exception e) {
        }
        long day = (date.getTime() - mydate.getTime()) / (24 * 60 * 60 * 1000L);
        return day;
    }

    /**
     * 两个时间之间的月差值
     *
     * @param dateStr1
     * @param dateStr2
     * @return
     */
    public static long getMonth(String dateStr1, String dateStr2) {
        if (dateStr1 == null || dateStr1.equals("")) {
            return 0;
        }
        if (dateStr2 == null || dateStr2.equals("")) {
            return 0;
        }
        // 转换为标准时间
        SimpleDateFormat myFormatter = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
        Date date1 = null;
        Date date2 = null;
        try {
            date1 = myFormatter.parse(dateStr1);
            date2 = myFormatter.parse(dateStr2);
        } catch (Exception e) {
        }
        long month = (date2.getTime() - date1.getTime()) / (24 * 30 * 60 * 60 * 1000L);
        return month;
    }

    /**
     * 获取指定日期月份的最后一天
     *
     * @param date 至少包含yyyyMMdd
     * @return
     */
    public static String getLastDateOfMonth(String date) {
        String str = date.substring(0, 6);
        String month = date.substring(4, 6);
        int mon = Integer.parseInt(month);
        if (mon == 1 || mon == 3 || mon == 5 || mon == 7 || mon == 8 || mon == 10 || mon == 12) {
            str += "31";
        } else if (mon == 4 || mon == 6 || mon == 9 || mon == 11) {
            str += "30";
        } else {
            if (isLeapYear(date)) {
                str += "29";
            } else {
                str += "28";
            }
        }
        str += "235959";
        return str;
    }
    /**
     * 获取指定日期月份的第一天
     *
     * @param date yyyy-MM-dd HH:mm:ss
     * @return
     */
    public static String getFirstDateOfMonth(String date) {
        String str = date.substring(0, 8);
        str+= "01 00:00:00";
        return str;
    }


    /**
     * 获取指定日期月份的最后一天
     *
     * @param date 至少包含yyyy-MM-dd
     * @return
     */
    public static String getLastDayOfMonth(String date) {
        String str = date.substring(0, 8);
        String month = date.substring(5, 7);
        int mon = Integer.parseInt(month);
        if (mon == 1 || mon == 3 || mon == 5 || mon == 7 || mon == 8 || mon == 10 || mon == 12) {
            str += "31";
        } else if (mon == 4 || mon == 6 || mon == 9 || mon == 11) {
            str += "30";
        } else {
            if (isLeapYear(date)) {
                str += "29";
            } else {
                str += "28";
            }
        }
        return str;
    }
    /**
     * 获取指定日期月份的第一天
     *
     * @param date yyyy-MM-dd HH:mm:ss
     * @return
     */
    public static String getFirstDayOfMonth(String date) {
        String str = date.substring(0, 8);
        str+= "01";
        return str;
    }
    /**
     * 判断是否闰年
     *
     * @param dateStr
     * @return
     */
    public static boolean isLeapYear(String dateStr) {
        /**
         * 详细设计： 1.被400整除是闰年，否则： 2.不能被4整除则不是闰年 3.能被4整除同时不能被100整除则是闰年
         * 3.能被4整除同时能被100整除则不是闰年
         */
        Date date = toDate(dateStr,"yyyyMMdd");
        GregorianCalendar gc = (GregorianCalendar) Calendar.getInstance();
        gc.setTime(date);
        int year = gc.get(Calendar.YEAR);
        if ((year % 400) == 0)
            return true;
        else if ((year % 4) == 0) {
            if ((year % 100) == 0)
                return false;
            else
                return true;
        } else
            return false;
    }

    /**
     * 判断二个时间是否在同一个周
     *
     * @param date1
     * @param date2
     * @return
     */
    public static boolean isSameWeekDates(Date date1, Date date2) {
        Calendar cal1 = Calendar.getInstance();
        Calendar cal2 = Calendar.getInstance();
        cal1.setTime(date1);
        cal2.setTime(date2);
        int subYear = cal1.get(Calendar.YEAR) - cal2.get(Calendar.YEAR);
        if (0 == subYear) {
            if (cal1.get(Calendar.WEEK_OF_YEAR) == cal2.get(Calendar.WEEK_OF_YEAR))
                return true;
        } else if (1 == subYear && 11 == cal2.get(Calendar.MONTH)) {
            // 如果12月的最后一周横跨来年第一周的话则最后一周即算做来年的第一周
            if (cal1.get(Calendar.WEEK_OF_YEAR) == cal2.get(Calendar.WEEK_OF_YEAR))
                return true;
        } else if (-1 == subYear && 11 == cal1.get(Calendar.MONTH)) {
            if (cal1.get(Calendar.WEEK_OF_YEAR) == cal2.get(Calendar.WEEK_OF_YEAR))
                return true;
        }
        return false;
    }

    /**
     * 校验传入日期字符串格式是否正确
     *
     * @param date
     */
    public static boolean isDate(String date) {
        SimpleDateFormat formatter;
        if (date == null)
            return false;
        if (date.length() > 10) {
            formatter = new SimpleDateFormat(DATE_PATTERN_CN);
        } else {
            formatter = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
        }
        try {
            formatter.parse(date);
            return true;
        } catch (ParseException pe) {
            return false;
        }
    }

    /**
     * 判断传入日期是否为同一天
     *
     * @param date1
     * @param date2
     * @return
     */
    public static boolean isSameDay(Date date1, Date date2) {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        return formatter.format(date1).equals(formatter.format(date2));
    }

    /**
     * 比较两个日期是否相等
     *
     * @param d1
     * @param d2
     * @return boolean
     * @author yangyc
     */
    public static boolean compareDate(Date d1, Date d2) {
        SimpleDateFormat sdf = new SimpleDateFormat(DATE_PATTERN_CN);
        String s1 = sdf.format(d1);
        String s2 = sdf.format(d2);
        return s1.equals(s2) ? true : false;
    }

    /**
     * 比较两个日期是否相等
     *
     * @param d1
     * @param d2
     * @return boolean
     * @author yangyc
     */
    public static boolean compareDateTime(Date d1, Date d2) {
        if(d1!=null&&d2!=null){
            String s1 = String.valueOf(d1.getTime());
            String s2 = String.valueOf(d2.getTime());
            return s1.equals(s2) ? true : false;
        }else if(d1==null&&d2==null){
            return true;
        }else{
            return false;
        }
    }

    /**
     * 时间比较  d1是否大于等于d2
     * @param d1
     * @param d2
     * @return
     */
    public static boolean compareTime(Date d1, Date d2) {
        if(d1!=null&&d2!=null){
            return d1.getTime()>=d2.getTime()?true : false;
        }else{
            return false;
        }
    }

    /**
     * 指定日期加天数
     * @param date
     * @param days
     * @return
     * @throws ParseException
     */
    public static Date addDay(Date date, int days) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat(DATE_PATTERN_CN_SHORT);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.DAY_OF_MONTH,calendar.get(Calendar.DAY_OF_MONTH) + days);
        return sdf.parse(sdf.format(calendar.getTime()));
    }

    /**
     * 方法概要:增加日期
     *
     * @param date
     * @param days
     * @return
     */
    public static Date addDays(Date date, int days) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.DAY_OF_MONTH, calendar.get(Calendar.DAY_OF_MONTH)
                + days);// 让日期加1
        date = calendar.getTime();
        return date;
    }

    /**
     * 根据所给日期返回两日期相差的秒数
     *
     * @param d1
     * @param d2
     * @return 返回两个日期间隔的毫秒数
     */

    public static long getSecond(Date d1, Date d2) {
        long a1 = d1.getTime();
        long a2 = d2.getTime();
        long a3 = (a1 - a2) / 1000;

        return a3;
    }

    /**
     * 根据所秒数,计算相差的时间并以**天**时返回
     *
     * @param
     * @param
     * @return
     */
    public static String getBeapartDayHour(long m) {
        String beapartdate = "";
        int nDay = (int) m / (24 * 60 * 60);
        int nHour = (int) (m - nDay * 24 * 60 * 60) / (60 * 60);
        beapartdate = nDay + "天" + nHour + "小时";
        return beapartdate;
    }

    /**
     * 时间搓
     * @return
     */
    public static String getTimestamp(){
        return String.valueOf(new Date().getTime());
    }

    public static String getTime(long time){
        SimpleDateFormat sdf = new SimpleDateFormat(DATE_PATTERN_CN);
        String date = sdf.format(new Date(time*1000L));
        return date;
    }

    /**
	 * 获取某年/月/周第一天日期
	 * 
	 * @return Date
	 */
	public static String getDateFirstDay(String date, String type) {
		int year, month = 0;
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");// 设置日期格式
		if (type.equals("year")) {
			// 获取年份数字
			year = Integer.parseInt(date.substring(0, 4));

			Calendar calendar = Calendar.getInstance();
			calendar.clear();
			calendar.set(Calendar.YEAR, year);
			Date currYearFirst = calendar.getTime();
			date = df.format(currYearFirst);
		} else if (type.equals("month")) {
			// 获取年份、月份数字
			year = Integer.parseInt(date.substring(0, 4));
			month = Integer.parseInt(date.substring(5, 7));

			Calendar cal = Calendar.getInstance();
			cal.set(Calendar.YEAR, year);
			cal.set(Calendar.MONTH, month - 1);
			cal.set(Calendar.DAY_OF_MONTH, cal.getMinimum(Calendar.DATE));
			date = df.format(cal.getTime());
		} else if (type.equals("week")) {
			Calendar calendar = new GregorianCalendar();
			calendar.setFirstDayOfWeek(Calendar.MONDAY);

			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			try {
				Date newDate = sdf.parse(date);
				calendar.setTime(newDate);
			} catch (ParseException e) {
				e.printStackTrace();
			}
			calendar.set(Calendar.DAY_OF_WEEK, calendar.getFirstDayOfWeek());
			Date currWeekFirst = calendar.getTime();
			date = df.format(currWeekFirst);
		}
		return date;
	}

	/**
	 * 获取某年/月/周最后一天日期
	 *
	 * @return Date
	 */
	public static String getDateLastDay(String date, String type) {
		int year, month = 0;
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");// 设置日期格式
		if (type.equals("year")) {
			// 获取年份数字
			year = Integer.parseInt(date.substring(0, 4));

			Calendar calendar = Calendar.getInstance();
			calendar.clear();
			calendar.set(Calendar.YEAR, year);
			calendar.roll(Calendar.DAY_OF_YEAR, -1);
			Date currYearLast = calendar.getTime();
			date = df.format(currYearLast);
		} else if (type.equals("month")) {
			// 获取年份、月份数字
			year = Integer.parseInt(date.substring(0, 4));
			month = Integer.parseInt(date.substring(5, 7));

			Calendar cal = Calendar.getInstance();
			cal.set(Calendar.YEAR, year);
			cal.set(Calendar.MONTH, month - 1);
			cal.set(Calendar.DAY_OF_MONTH, 1);
			int value = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
			cal.set(Calendar.DAY_OF_MONTH, value);
			date = df.format(cal.getTime());
		} else if (type.equals("week")) {
			Calendar calendar = new GregorianCalendar();
			calendar.setFirstDayOfWeek(Calendar.MONDAY);
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			try {
				Date newDate = sdf.parse(date);
				calendar.setTime(newDate);
			} catch (ParseException e) {
				e.printStackTrace();
			}
			calendar.set(Calendar.DAY_OF_WEEK, calendar.getFirstDayOfWeek() + 6);
			Date currWeekLast = calendar.getTime();
			date = df.format(currWeekLast);
		}
		return date;
	}
	
	/**
     * 返回指定年季的季的第一天
     *
     * @param year
     * @param quarter
     * @return
     */
	public static String getFirstDayOfQuarter(Integer year, Integer quarter) {
		Calendar calendar = Calendar.getInstance();
		String month = "";
		if (quarter == 1) {
			month = "01";
		} else if (quarter == 2) {
			month = "04";
		} else if (quarter == 3) {
			month = "07";
		} else if (quarter == 4) {
			month = "10";
		} else {
			month = calendar.get(Calendar.MONTH) + "";
		}
		return getDateFirstDay(year + "-" + month, "month");
	}

    /**
     * 返回指定年季的季的最后一天
     *
     * @param year
     * @param quarter
     * @return
     */
	public static String getLastDayOfQuarter(Integer year, Integer quarter) {
		Calendar calendar = Calendar.getInstance();
		String month = "";
		if (quarter == 1) {
			month = "03";
		} else if (quarter == 2) {
			month = "06";
		} else if (quarter == 3) {
			month = "09";
		} else if (quarter == 4) {
			month = "12";
		} else {
			month = calendar.get(Calendar.MONTH) + "";
		}
		return getDateLastDay(year + "-" + month, "month");
	}

}
