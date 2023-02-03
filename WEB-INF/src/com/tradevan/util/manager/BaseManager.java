package com.tradevan.util.manager;

import java.sql.Connection;

public abstract class BaseManager
    implements Manager {
  public static Object getManager() {
    return new Object();
  }

  public abstract Object prepareRepList(Object model, Connection con) throws Exception;



}
