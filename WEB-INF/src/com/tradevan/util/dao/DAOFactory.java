package com.tradevan.util.dao;

public class DAOFactory {
    public DAOFactory() {
    }

	public static RdbCommonDao getRdbCommonDao(String poolname) {
	      System.out.println("test1060214");
		  RdbCommonDao cld = new RdbCommonDao(poolname);
		  return cld;
	}
	
}
