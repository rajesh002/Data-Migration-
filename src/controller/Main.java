package controller;

import java.io.*;
import java.sql.Connection;

import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
public class Main {
	 public static void main(String[] args) {
	        String jdbcURL = "jdbc:oracle:thin:@localhost:1521:orcl";
	        String username = "system";
	        String password = "Rajesh123";
	        String excelFilePath = "Students.xlsx";
	 
	        Connection connection = null;
	 
	        try { 
	            FileInputStream inputStream = new FileInputStream(excelFilePath);
	 
	            Workbook workbook = new XSSFWorkbook(inputStream);
	 
	            Sheet firstSheet = workbook.getSheetAt(0);
	            Iterator<Row> rowIterator = firstSheet.iterator();
	            Iterator<Row> rowIterator1 = firstSheet.iterator();
	 
	            connection = DriverManager.getConnection(jdbcURL, username, password);
	            connection.setAutoCommit(false);
	  
	                
	            String create = "CREATE TABLE student(";
	            StringBuilder createstmt = new StringBuilder(create);
	            int count=0;
	            
	            String insert = "INSERT INTO student VALUES (";
	            
	            
	            
	            
	           if(rowIterator.hasNext()){
	        	   Row secondRow = firstSheet.getRow(1);
	        	   Row firsttRow = rowIterator.next();
	        	   
	        	   Iterator<Cell> headIterator = firsttRow.cellIterator();
	               Iterator<Cell> cellIterator = secondRow.cellIterator();
	               
	                while (cellIterator.hasNext()) {
	                	count+=1;
	                    Cell nextCell = cellIterator.next();
	                    Cell headNextCell = headIterator.next();
	                    CellType type = nextCell.getCellType();
	                    if (type == CellType.STRING) 
	                    	createstmt.append(headNextCell+" varchar(30),");
	                    else if(type == CellType.NUMERIC)
	                    	createstmt.append(headNextCell+" number,");
	                    }
	                
	                createstmt.setCharAt(createstmt.length()-1,')');      
	                }
	           
	           
	           try {
	           Statement stmt = connection.createStatement();
	           String temp=createstmt.toString();
	           stmt.executeUpdate(temp);
				} catch (SQLException e) {
					
				}
	             
	          
	           
	           for(int i=1;i<=count;i++) {
	        	   if(i!=count)
	        		   insert+="?,";
	        	   else
	        		   insert+="?";
	           }
	           insert+=")";
	           
	            
	           PreparedStatement statement = connection.prepareStatement(insert);
	           
	           
	           rowIterator1.next();
	           int index;
	            while (rowIterator1.hasNext()) {
	            	index=1;
	                Row nextRow = rowIterator1.next();
	                Iterator<Cell> cellIterator = nextRow.cellIterator();
	 
	                while (cellIterator.hasNext()) {
	                    Cell nextCell = cellIterator.next();
	                    CellType type = nextCell.getCellType();
	                    if (type == CellType.STRING) {
	                        String name = nextCell.getStringCellValue();
	                        statement.setString(index, name);
	                    }
	                    else {
	                        int progress = (int)nextCell.getNumericCellValue();
	                        statement.setInt(index, progress);
	                    }
	                    index++;
	                }
	                statement.addBatch();
	                    
	                }
	             
	            System.out.println(createstmt);
	            System.out.println(insert);
	            
	            
	            workbook.close();
	            
	            statement.executeBatch();
	  
	            connection.commit();
	            connection.close();
	             
	            
	            System.out.printf("sucess");
	             
	        }
	        catch(Exception e)  
			{  
				e.printStackTrace();  
			}  
	 
	    }
}
