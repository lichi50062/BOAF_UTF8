package com.tradevan.util.sql;

public class SQLParameter {
    private Object value;
    private int index;
    private int sqlType=0;
    
    public SQLParameter(int index, Object value, int sqlType) {
        setIndex(index);
        setValue(value);
        setSQLType(sqlType);
    }
    
    public SQLParameter(int index, Object value) {
        setIndex(index);
        setValue(value);
    }
    
    public int getIndex() {
        return this.index;
    }
    public void setIndex(int index) {
        this.index = index;
    }
    public int getSQLType() {
        return this.sqlType;
    }
    public void setSQLType(int sqltype) {
        sqlType = sqltype;
    }
    public Object getValue() {
        return this.value;
    }
    public void setValue(Object value) {
        this.value = value;
    }
    
    

}
