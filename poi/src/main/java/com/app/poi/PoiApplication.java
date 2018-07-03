package com.app.poi;

import java.io.IOException;
import java.sql.SQLException;

import javax.annotation.PostConstruct;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.app.poi.db.DB;
import com.app.poi.service.ReadExeclService;

@SpringBootApplication
public class PoiApplication {
	
	@Autowired
	private ReadExeclService execlService;
	
	@Autowired
	private DB db;
	

	public static void main(String[] args) {
		SpringApplication.run(PoiApplication.class, args);
	}
	
	@PostConstruct
	public void init() throws IOException {
		
//	execlService.readEcexcelFile();
//	execlService.readExeclData();
		try {
			execlService.generateDdatafromDB(db.getDBConnection());
			execlService.readExeclData(db.getDBConnection());
			
		} catch (ClassNotFoundException | SQLException e) {
			e.printStackTrace();
		}
		
		
	}
}
