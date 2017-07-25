package com.leaf.model;

import java.io.Serializable;

/**
 * Created by leaf
 * 时间 2017/5/23 0023 20:41
 */
public class ExcelInfo implements Serializable {
    protected String columnsCode;//字段名
    protected String columnsName;
    protected String type;//类型 String Date Number
    protected int length;//长度 数据库字段长度
    protected String value; //单元格内容
    protected boolean isMoney;//
    protected boolean isError;  //单元格校验是否有错
    public ExcelInfo(){}
    public ExcelInfo(String columnsCode, String columnsName){
        this.columnsCode = columnsCode;
        this.columnsName = columnsName;
    }
    public ExcelInfo(String columnsCode, String columnsName, boolean isMoney){
        this.columnsCode = columnsCode;
        this.columnsName = columnsName;
        this.isMoney = isMoney;
    }
    public ExcelInfo(String columnsCode, String columnsName, String type, int length){
        this.columnsCode = columnsCode;
        this.columnsName = columnsName;
        this.type = type;
        this.length = length;
    }
    public ExcelInfo(String columnsCode, String columnsName, String type, int length, boolean isMoney){
        this.columnsCode = columnsCode;
        this.columnsName = columnsName;
        this.type = type;
        this.length = length;
        this.isMoney = isMoney;
    }
    public boolean isMoney() {
        return isMoney;
    }

    public void setMoney(boolean money) {
        isMoney = money;
    }
    public String getColumnsCode() {
        return columnsCode;
    }

    public void setColumnsCode(String columnsCode) {
        this.columnsCode = columnsCode;
    }

    public String getColumnsName() {
        return columnsName;
    }

    public void setColumnsName(String columnsName) {
        this.columnsName = columnsName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public int getLength() {
        return length;
    }

    public void setLength(int length) {
        this.length = length;
    }

    public String getValue() {
        return value!=null?value:"";
    }

    public void setValue(String value) {
        this.value = value;
    }

    public boolean isError() {
        return isError;
    }

    public void setError(boolean error) {
        isError = error;
    }
}

