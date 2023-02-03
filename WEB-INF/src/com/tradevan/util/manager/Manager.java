package com.tradevan.util.manager;
import java.sql.Connection;
public interface Manager {
  public Object prepareRepList(Object model, Connection con) throws Exception;

}
