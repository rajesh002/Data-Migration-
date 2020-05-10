package controller;

import java.io.*;
import java.sql.Connection;

import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
public class Main {
	 public static void main(String[] args) {
	        String jdbcURL = "jdbc:oracle:thin:@localhost:1521:orcl";
	        String username = "system";
	        String password = "Rajesh123";
	 
	        String excelFilePath = "Students.xlsx";
	 
	        int batchSize = 20;
	 
	        Connection connection = null;
	 
	        try {
	            long start = System.currentTimeMillis();
	             
	            FileInputStream inputStream = new FileInputStream(excelFilePath);
	 
	            Workbook workbook = new XSSFWorkbook(inputStream);
	 
	            Sheet firstSheet = workbook.getSheetAt(0);
	            Iterator<Row> rowIterator = firstSheet.iterator();
	 
	            connection = DriverManager.getConnection(jdbcURL, username, password);
	            connection.setAutoCommit(false);
	  
	            String sql = "INSERT INTO students (name, marks, progress) VALUES (?, ?, ?)";
	            PreparedStatement statement = connection.prepareStatement(sql);    
	             
	            int count = 0;
	             
	            rowIterator.next(); 
	             
	            while (rowIterator.hasNext()) {
	                Row nextRow = rowIterator.next();
	                Iterator<Cell> cellIterator = nextRow.cellIterator();
	 
	                while (cellIterator.hasNext()) {
	                    Cell nextCell = cellIterator.next();
	 
	                    int columnIndex = nextCell.getColumnIndex();
	 
	                    switch (columnIndex) {
	                    case 0:
	                        String name = nextCell.getStringCellValue();
	                        statement.setString(1, name);
	                        break;
	                    case 1:
	                    	int marks = (int) nextCell.getNumericCellValue();
	                        statement.setInt(2, marks);
	                    case 2:
	                        int progress = (int) nextCell.getNumericCellValue();
	                        statement.setInt(3, progress);
	                    }
	 
	                }
	                 
	                statement.addBatch();
	                 
	                if (count % batchSize == 0) {
	                    statement.executeBatch();
	                }              
	 
	            }
	 
	            workbook.close();
	             
	            // execute the remaining queries
	            statement.executeBatch();
	  
	            connection.commit();
	            connection.close();
	             
	            long end = System.currentTimeMillis();
	            System.out.printf("Import done in %d ms\n", (end - start));
	             
	        } catch (IOException ex1) {
	            System.out.println("Error reading file");
	            ex1.printStackTrace();
	        } catch (SQLException ex2) {
	            System.out.println("Database error");
	            ex2.printStackTrace();
	        }
	 
	    }
}
