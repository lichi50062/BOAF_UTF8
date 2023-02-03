package com.tradevan.util.sql;

import java.sql.*;

public class SQLTemplateException extends SQLException{
    public SQLTemplateException() {
      super();
    }

    public SQLTemplateException(String msg) {
        super("SQLTemplateException : "+msg);
    }


}