package com.app.poi.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import oracle.net.aso.p;

@Service
public class ReadExeclService {
	private static final String FILE_NAME = "./stu.xlsx";

	public void readEcexcelFile() {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
		Object[][] datatypes = { { "Datatype", "Type", "Size(in bytes)" }, { "int", "Primitive", 2 },
				{ "float", "Primitive", 4 }, { "double", "Primitive", 8 }, { "char", "Primitive", 1 },
				{ "String", "Non-Primitive", "No fixed size" } };

		int rowNum = 0;
		System.out.println("Creating excel");

		for (Object[] datatype : datatypes) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (Object field : datatype) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}
		}

		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Done");
	}

	public void generateDdatafromDB(Connection dbConnection) throws SQLException {
		Statement statement = dbConnection.createStatement();
		ResultSet rs = statement.executeQuery("SELECT * FROM student");
		List<Student> student = new ArrayList<>();
		
		while (rs.next()) {
			student.add(new Student(rs.getInt("usn"), rs.getString("name")));
		}
		
		rs.close();
		statement.close();
		dbConnection.close();

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student");

		int rowNum = 0;
		System.out.println("Creating student excel");

		int coloum=1 ;
		Row row = null;
		Cell cell = null;

		for (Student stu : student) {
			if(coloum==1) {
			row = sheet.createRow(rowNum++);
			int colNum = 0;
			cell = row.createCell(colNum++);
			cell.setCellValue("USN");
			row.createCell(colNum++).setCellValue("Name");
			coloum++;
			}
			row = sheet.createRow(rowNum++);
			int colNum = 0;
			cell = row.createCell(colNum++);
			cell.setCellValue(stu.getUsn());
			row.createCell(colNum++).setCellValue(stu.getName());
			coloum++;
			
		}

		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Done");
	}


	public void readExeclData(Connection connection) throws IOException, SQLException {

		FileInputStream excelFile = new FileInputStream(new File("./input.xlsx"));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet datatypeSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = datatypeSheet.iterator();

		List<Student> stu = new ArrayList<>();

		while (iterator.hasNext()) {

			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();

			while (cellIterator.hasNext()) {

				Cell currentCell1 = cellIterator.next();
				Cell currentCell2 = cellIterator.next();
				Student s = new Student();
				s.setUsn((int) currentCell1.getNumericCellValue());
				s.setName(currentCell2.getStringCellValue());
				stu.add(s);
			}
		}
		
		System.out.println("write data to DB");
		
		PreparedStatement prepareStatement = connection.prepareStatement("insert into student values(?,?)");
		for (Student student : stu) {
			prepareStatement.setInt(1, student.getUsn());
			prepareStatement.setString(2, student.getName());
			
			prepareStatement.executeUpdate();
		}
		
		System.out.println("write data to DB Done");
		
		
	}
}
