package com.tradevan.util.dao;

import java.util.*;
import java.io.*;

public class DataObject implements Serializable {
    private Map values;

    public DataObject(Map values) {
    	this.values = values;
    }
    public DataObject() {
    	this.values = new HashMap();
    }

    public void setValue(String key, Object value) {
        values.put(key, value);
    }
    public Object getValue(String key) {
        return values.get(key);
    }
    public Map getValues() {
        return values;
    }
    public Object [] getKeys() {
    	return values.keySet().toArray();
    }
}
