/*
 * Created on 2010/1/28
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package com.tradevan.util.dao;

import java.sql.Connection;

import com.tradevan.util.sql.DBConnectFactory;

public class DBFactory implements DBConnectFactory {
     
    public Connection getDBConnection() {
    	Connection connection = null;
    	if(connection == null){
            connection = Factory.getConnection();
     	}   
    	return connection;
    }
    
    
}
