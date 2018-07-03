package com.app.poi.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import org.springframework.stereotype.Service;

@Service
public class DB {

	public Connection getDBConnection() throws SQLException, ClassNotFoundException {
		Connection con = null;
		Class.forName("oracle.jdbc.driver.OracleDriver");  
		con = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe","system","admin");
		
		if(!con.isClosed()) {
			System.out.println("connected to hsql db");
		} else {
			System.out.println(" not able to connect to hsqldb");
		}
		
		return con;
	}
}
