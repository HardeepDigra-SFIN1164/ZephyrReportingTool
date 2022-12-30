package com.report;

import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
{
	public static int rowNo = 1;
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
			
			for(int i = sheet.getLastRowNum(); i >= 1; i--)
			{
				//System.out.println("Excel row: "+i);
			  Row row = sheet.getRow(i);
			   sheet.removeRow(row);
			  
			}
			
		
		workbook.close();
		//System.out.println("Excel sheet is cleared successfully");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}
	@SuppressWarnings("resource")
	public static String readExcel(String data){
		try {
		String userDirectory = System.getProperty("user.dir");
		FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");
	
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
//		int rowCount = sheet.getLastRowNum();
//		System.out.println("Number of Test cases: "+rowCount);
//		
		//Row row = 1 ;
		Row row = sheet.getRow(rowNo);
	//	int  totalRows = 3;
	//	System.out.println("Number of rows: "+totalRows);
	//	while(totalRows!=0)
		{ 
			//row= sheet.getRow(1);
			if(data.equals(null))
			{
				System.out.println("projectId, cycleId and folderId is null");
			}
			else if(data.equals("projectId"))
				{
				Cell Idcellpid = row.getCell(1);
				int pid =  (int) Idcellpid.getNumericCellValue();
			//	System.out.println("pid: "+pid);
							
			      String pcode = Integer.toString(pid);
			    return pcode;	
				}
				
				
				else if(data.equals("cycleId"))
				{
					Cell Idcellcid = row.getCell(3);
					String cid = Idcellcid.getStringCellValue();
			//		System.out.println("cid: "+cid);
					return cid;
				}
			
				
				else if(data.equals("folderId"))
				{
					Cell Idcellfid = row.getCell(5);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
			
				else if(data.equals("cycleName"))
				{
					Cell Idcellfid = row.getCell(2);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
				else if(data.equals("folderName"))
				{
					Cell Idcellfid = row.getCell(4);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
		
			//totalRows--;
			
		}
		
		
		
		
		
	
	
//		
//		testCaseIDs.add(testCaseId); 
//
//		Cell statCell = row.getCell(9);
//		String Excelstatus = statCell.getStringCellValue();
//		testCaseStatus.add(Excelstatus);
		
		/*
//		int count = 0;
//		for (int i = 1; i <= rowCount; i++)
//		{
//			Row row = sheet.getRow(i);
//		if(row.toString()!= "")
//		{
//			count++;
//		}
//		}
//		
//		System.out.println("Number of non empty Rows: "+count);
//		
		int i =0;
		int rowCount_ = sheet.getLastRowNum();
		for (i = 0; i < rowCount_; i++) {
		    boolean rowEmpty = true;
		    String currentRow = "";
		    for (int j = 0; j < sheet.getColumns(); j++) {
		        Cell cell = sheet.getCell(j, i);
		        String con=cell.getContents();
		        if(con !=null && con.length()!=0){
		            rowEmpty = false;
		        }
		        currentRow += con + "|";
		    }
		    if(!rowEmpty) {
		        System.out.println(currentRow);
		    }
		}

//		for (int i = 1; i <= rowCount; i++) {
//			Row row = sheet.getRow(i);
//			Cell Idcell = row.getCell(0);
//			String testCaseId = Idcell.getStringCellValue();
//			testCaseIDs.add(testCaseId); 
//
//			Cell statCell = row.getCell(9);
//			String Excelstatus = statCell.getStringCellValue();
//			testCaseStatus.add(Excelstatus);
//
//			if (testCaseId == null) {
//				break;
//			}
//		}*/
		
		
		
		workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
		
	}


}
