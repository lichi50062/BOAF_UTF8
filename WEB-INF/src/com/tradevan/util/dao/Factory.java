package com.tradevan.util.dao;

import java.sql.Connection;
import java.sql.SQLException;

import com.tradevan.commons.cdao.*;
import com.tradevan.util.sql.*;

public class Factory {
    
    private static com.tradevan.commons.cdao.DAOFactory factory = new com.tradevan.commons.cdao.DAOFactory();
    public static com.tradevan.commons.cdao.DAOFactory getDaoFactory(){
        return factory;
    }
   
    public static Connection getConnection(){
        DaoConnection daoConn=factory.getDaoConnection(null);
        try {
            daoConn.openConnection();
        } catch (SQLException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return daoConn.getConnection();
    }    
}
