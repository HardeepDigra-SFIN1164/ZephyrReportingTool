package com.report;


import  java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

import  org.apache.poi.hssf.usermodel.HSSFSheet;  
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFRow;  

public class WriteExcel  
{  
	static SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm");  
    static Date date = new Date();  
    static String FileName = "RESULT";//formatter.format(date);
 
    
	public static void clearExcelData()
	{
		try
		{
			String userDirectory = System.getProperty("user.dir");
			FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");
		
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			/*Row row = sheet.getRow(rowNo);
			
			for (int i = sheet.getLastRowNum(); i >= 1; i--) {
				  sheet.removeRow(sheet.getRow(i));
				}*/
			int row_value = report_CycleName.folderCount;
			for(int i = row_value; i >= 1; i--)
			{
			//System.out.println("Excel row: "+row_value);
//			  Row row = sheet.getRow(i);
//			   sheet.removeRow(row);
				   Cell cell0 = sheet.getRow(1).getCell(0);
					cell0.setCellValue(" ");
					
					Cell cell1 = sheet.getRow(1).getCell(1);
					cell1.setCellValue(" ");
					
					Cell cell2 = sheet.getRow(1).getCell(2);
					cell2.setCellValue(" ");
					
					Cell cell3 = sheet.getRow(1).getCell(3);
					cell3.setCellValue(" ");
					
					Cell cell4 = sheet.getRow(i).getCell(4);
					cell4.setCellValue(" ");
					
					Cell cell5 = sheet.getRow(i).getCell(5);
					cell5.setCellValue(" ");
					
					Cell cell6 = sheet.getRow(i).getCell(6);
					cell6.setCellValue(" ");
					
					Cell cell7 = sheet.getRow(i).getCell(7);
					cell7.setCellValue(" ");
					
					Cell cell8 = sheet.getRow(i).getCell(8);
					cell8.setCellValue(" ");
					
					Cell cell9 = sheet.getRow(i).getCell(9);
					cell9.setCellValue(" ");
					
					Cell cell10 = sheet.getRow(i).getCell(10);
					cell10.setCellValue(" ");
					
					Cell cell11 = sheet.getRow(i).getCell(11);
					cell11.setCellValue(" ");
					
					Cell cell12 = sheet.getRow(i).getCell(12);
					cell12.setCellValue(" ");
					
					
					Cell cell13 = sheet.getRow(i).getCell(13);
					cell13.setCellValue(" ");
					

					
			}
			
			  
			FileOutputStream output = new FileOutputStream(userDirectory + "/" +"Report.xlsx");
			workbook.write(output);
			output.close();
		
		workbook.close();
	//	System.out.println("Excel sheet is cleared successfully");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}
	public static void main(String[]args)   
	{  
		
	}
	public void writeExcel(int RowNumber , String count , int Status) {
	try   
	{  	
		   int rowNo = RowNumber;
		
		String userDirectory = System.getProperty("user.dir");

		String filename = userDirectory + "/" +"Report.xlsx" ;  
		
		FileInputStream fis = new FileInputStream(filename); 
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Cell cell = sheet.getRow(rowNo).getCell(Status);
		
		cell.setCellValue(count);

		fis.close();

		FileOutputStream outFile = new FileOutputStream((filename));
		workbook.write(outFile);
		outFile.close();
		//filename = userDirectory + "/"+FileName+"- Report.xlsx";
		
		
		
		
	//System.out.println("Excel file has been generated successfully.");  
	//afterExecution();

	}   

	catch (Exception e)   
	{  
		e.printStackTrace();  
	}  
	}	  
	
	/*public static void afterExecution() throws IOException
	{

		String userDirectory = System.getProperty("user.dir");

		String filename = userDirectory + "/" +"Report.xlsx" ;  
		
		FileInputStream fis = new FileInputStream(filename);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		for(int i = 5 ; i<=12 ; i++)
		{
			if(sheet.getRow(rowNo).getCell(i).getStringCellValue() == " ")
			{
			Cell cell = sheet.getRow(rowNo).getCell(i);
			
			cell.setCellValue("0");
			}
		}
		

		fis.close();

		FileOutputStream outFile = new FileOutputStream((filename));
		workbook.write(outFile);
		outFile.close();
	}*/
	
	public static void writeExcel2(int RowNumber ,int ColNumber, String value ) {
	try   
	{  	
		   int rowNo = RowNumber+1;
		
		String userDirectory = System.getProperty("user.dir");

		String filename = userDirectory + "/" +"Report.xlsx" ;  
		
		FileInputStream fis = new FileInputStream(filename); 
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(1);
		
		Cell cell = sheet.getRow(rowNo).getCell(ColNumber);
		
		cell.setCellValue(value);

		fis.close();

		FileOutputStream outFile = new FileOutputStream((filename));
		workbook.write(outFile);
		outFile.close();
		//filename = userDirectory + "/"+FileName+"- Report.xlsx";
		
		
		
		
	//System.out.println("Excel file has been generated successfully.");  
	//afterExecution();

	}   

	catch (Exception e)   
	{  
		e.printStackTrace();  
	}  
	}
	
	public static void clearExcelData2()
	{
		try
		{
			
			String userDirectory = System.getProperty("user.dir");

			String filename = userDirectory + "/" +"Report.xlsx" ;  
			XSSFWorkbook workbook = new XSSFWorkbook(filename);
			XSSFSheet sheet = workbook.getSheetAt(1);
			
			/*Row row = sheet.getRow(rowNo);
			
			for (int i = sheet.getLastRowNum(); i >= 1; i--) {
				  sheet.removeRow(sheet.getRow(i));
				}*/
			
			for(int i = report_CycleName.defects_rowsize; i >= 1; i--)
			{
	//			System.out.println("Excel row: "+i);
//			  Row row = sheet.getRow(i);
//			   sheet.removeRow(row);
			   Cell cell = sheet.getRow(i).getCell(0);
				cell.setCellValue(" ");
				Cell cell1 = sheet.getRow(i).getCell(1);
				cell1.setCellValue(" ");
				Cell cell2 = sheet.getRow(i).getCell(2);
				cell2.setCellValue(" ");
				report_CycleName.masterSummaryList.clear();
				report_CycleName.masterDefectList.clear();
			}
			
			  
			FileOutputStream output = new FileOutputStream(filename);
			workbook.write(output);
			output.close();
		
		workbook.close();
	//	System.out.println("Excel sheet is cleared successfully");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}
}  