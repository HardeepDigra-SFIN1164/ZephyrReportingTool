package com.report;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Properties;

import org.apache.commons.text.CaseUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.report.report_CycleName;
public class ReadData {
	public static FileInputStream fis;
	@SuppressWarnings("finally")
	public static void main (String[] args) throws InterruptedException, IOException
	{
		ExcelToHtml();
	}
	public static void ExcelToHtml() throws IOException
	{
	File obj1= new File((System.getProperty("user.dir") +"/Report.xlsx"));
	FileInputStream creds = new FileInputStream(obj1);
	XSSFWorkbook workbook = new XSSFWorkbook(creds);
	XSSFSheet sheet = workbook.getSheetAt(0);
	int rowCount= sheet.getPhysicalNumberOfRows();
	int ActualCount=0;
	int a=report_CycleName.last_rowCount;
//	System.out.println(a);
	for (int i=1; i<=a;i++)
	{
		try {
		String total1=sheet.getRow(i).getCell(11).getStringCellValue().trim();
		if(total1.equals("") || total1.equals(" "))
		continue;
		else
			ActualCount++;
		}
		finally{
			continue;
		}
	}
	//reading header values
	String header2 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(2).getStringCellValue()),true,' ');
	String header4 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(4).getStringCellValue()),true,' ');
	String header6 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(6).getStringCellValue()),true,' ');
	String header7 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(7).getStringCellValue()),true,' ');
	String header8 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(8).getStringCellValue()),true,' ');
	String header9 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(9).getStringCellValue()),true,' ');
	String header10 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(10).getStringCellValue()),true,' ');
	String header11 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(11).getStringCellValue()),true,' ');
	String header12 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(12).getStringCellValue()),true,' ');
	String header13 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(13).getStringCellValue()),true,' ');
	//displaying header values
	String html = "<html><head><title>Test Result</title>"
			+ "<style>"
			+ "     	table, td, th {padding: 5px; border: 1px solid black;}"
			+ "			table {border-collapse: collapse;}"
			+ "</style>"
			+ "</head>"
			+ "<body><b style=color:Black;>Hello All,<b>"
			+ "<b><p style=color:Black;> Please find the attached report below: <p></b>"
			+ "<div id=container style=width: 400px; height: 200px; margin: 0 auto></div>"
			+ "<table id = tb1>"
			+ "<tbody><Caption style=text-align:left;color:Black;text-decoration:underline;font-weight:bold>Test Summary:</Caption>"
			+ "<tr bgcolor=#0B5394>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header2+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;border-top-color:black;border-right-color:black;border-bottom-color:black;>"+header4+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header11+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;border-top-color:black;border-right-color:black;border-bottom-color:black;>"+header6+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;border-top-color:black;border-right-color:black;border-bottom-color:black;>"+header7+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;border-top-color:black;border-right-color:black;border-bottom-color:black;>"+header8+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header9+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header10+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header12+"</td>"
			+ "<td style=color:White;text-align:center;font-weight:bold;>"+header13+"*</td>"
			+ "</tr>";
	String defects="",wip="",blocked="",unexecuted="", fail="", pass="", total="";
	int total1=0, total2=0, pass1=0, pass2=0, fail2=0, fail1=0, unexecuted1=0, unexecuted2 = 0, wip1=0, wip2=0, blocked1=0, blocked2=0, defects1=0, defects2= 0;
	float passPercent=0, failPercent=0, unexePercent=0, wipPercent=0, blockedPercent=0;
	for (int i=1; i<=ActualCount;i++)
	{
	total=sheet.getRow(i).getCell(11).getStringCellValue().trim();
	if(total.equals(""))
	{
		total1 = 0; 
	}
	else
	{
		total1=Integer.parseInt(sheet.getRow(i).getCell(11).getStringCellValue().trim());
	}
	total2  = total2 + total1;
	pass=sheet.getRow(i).getCell(6).getStringCellValue().trim();
	if(pass.equals(""))
	{
		pass1 = 0; 
	}
	else
	{
		pass1=Integer.parseInt(sheet.getRow(i).getCell(6).getStringCellValue().trim());
	}
	pass2=pass2 + pass1;
	passPercent=(float)(((int)(((pass2*100)/total2) *100.0))/100.0);
	fail=sheet.getRow(i).getCell(7).getStringCellValue().trim();
	if(fail.equals(""))
	{
		fail1 = 0; 
	}
	else
	{
		fail1=Integer.parseInt(sheet.getRow(i).getCell(7).getStringCellValue().trim());
	}	
	fail2=fail2 + fail1;
	failPercent=(float)(((int)(((fail2*100)/total2) *100.0))/100.0);
	unexecuted=sheet.getRow(i).getCell(8).getStringCellValue().trim();
	if(unexecuted.equals(""))
	{
		unexecuted1 = 0; 
	}
	else
	{
		unexecuted1=Integer.parseInt(sheet.getRow(i).getCell(8).getStringCellValue().trim());
	}
	unexecuted2=unexecuted2 + unexecuted1;
	unexePercent=(float)(((int)(((unexecuted2*100)/total2) *100.0))/100.0);
	wip=sheet.getRow(i).getCell(9).getStringCellValue().trim();
	if(wip.equals(""))
	{
		wip1 = 0; 
	}
	else
	{
		wip1=Integer.parseInt(sheet.getRow(i).getCell(9).getStringCellValue().trim());
	}
	wip2=wip2+wip1;
	wipPercent=(float)(((int)(((wip2*100)/total2) *100.0))/100.0);
	blocked = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	if(blocked.equals(""))
	{
		blocked1 = 0; 
	}
	else
	{
		blocked1=Integer.parseInt(sheet.getRow(i).getCell(10).getStringCellValue().trim());
	}
	blocked2=blocked2 + blocked1;
	blockedPercent=(float)(((int)(((blocked2*100)/total2) *100.0))/100.0);
	defects = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	if(defects.equals(""))
	{
		defects1 = 0; 
	}
	else
	{
		defects1=Integer.parseInt(sheet.getRow(i).getCell(12).getStringCellValue());
	}
	defects2= defects2 + defects1;
	}
	//to read data till the rowCount
	for(int i =1; i <=ActualCount; i++ )
	{
		XSSFRow row=sheet.getRow(i);
		if(row!=null)
		{
		String data2 = sheet.getRow(i).getCell(2).getStringCellValue();
		String data4 = sheet.getRow(i).getCell(4).getStringCellValue();
		String data6 = sheet.getRow(i).getCell(6).getStringCellValue();
		String data7 = sheet.getRow(i).getCell(7).getStringCellValue();
		String data8 = sheet.getRow(i).getCell(8).getStringCellValue();
		String data9 = sheet.getRow(i).getCell(9).getStringCellValue();
		String data10 = sheet.getRow(i).getCell(10).getStringCellValue();
		String data11 = sheet.getRow(i).getCell(11).getStringCellValue();
		String data12 = sheet.getRow(i).getCell(12).getStringCellValue();
		String data13= sheet.getRow(i).getCell(13).getStringCellValue();

		html = html+"<tr>"
				+ "<td style=color:Black;text-align:center;>"+data2+"</td>"
				+ "<td style=color:Black;text-align:center;border-right-color:black;border-bottom-color:black;>"+data4+"</td>"
				+ "<td style=color:Blue;text-align:center;>"+data11+"</td>"
				+ "<td style=color:Green;text-align:center;border-right-color:black;border-bottom-color:black;>"+data6+"</td>"
				+ "<td style=color:Red;text-align:center;border-right-color:black;border-bottom-color:black;>"+data7+"</td>"
				+ "<td style=color:#D16002;text-align:center;border-right-color:black;border-bottom-color:black;>"+data8+"</td>"
				+ "<td style=color:Black;text-align:center;>"+data9+"</td>"
				+ "<td style=color:Black;text-align:center;>"+data10+"</td>"
				+ "<td style=color:Red;text-align:center;>"+data12+"</td>"
				+ "<td style=color:Red;text-align:center;>"+data13+"</td>"

				+ "</tr>";
		}
	}
	//total count row
	html=html+"<tr bgcolor=#CFE2F3>"
			+ "<td style=color:Black;text-align:center; colspan=2;font-weight:bold><b>TotalCount</b></td>"
	+ "<td style=color:Blue;text-align:center;font-weight:bold>"+total2+"</td>"
	+ "<td style=color:Green;text-align:center;border-right-color:black;border-bottom-color:black;font-weight:bold>"+pass2+"</td>"
	+ "<td style=color:Red;text-align:center;font-weight:bold;font-weight:bold>"+fail2+"</td>"
	+ "<td style=color:#D16002;text-align:center;border-right-color:black;border-bottom-color:black;font-weight:bold>"+unexecuted2+"</td>"
	+ "<td style=color:Black;text-align:center;border-right-color:black;border-bottom-color:black;font-weight:bold>"+wip2+"</td>"
	+ "<td style=color:Black;text-align:center;border-right-color:black;border-bottom-color:black;font-weight:bold>"+blocked2+"</td>"
	+ "<td style=color:Red;text-align:center;font-weight:bold;>"+defects2+"</td>"
	+ "<td style=color:Red;text-align:center;font-weight:bold;> </td>"
	+ "</tr>"
	+ "</tbody></table>"
	+ "<p style=color:Red;>*Open defects are shown in defects name column</p><br>";
	//Defects table
	if(defects2!=0) 
	{
		File obj2= new File((System.getProperty("user.dir") +"/Report.xlsx"));
		FileInputStream creds2 = new FileInputStream(obj2);
		XSSFWorkbook workbook2 = new XSSFWorkbook(creds2);
		XSSFSheet sheet2 = workbook2.getSheetAt(1);
		String dheader1 = CaseUtils.toCamelCase((sheet2.getRow(0).getCell(0).getStringCellValue()),true,' ');
		String dheader2 = CaseUtils.toCamelCase((sheet2.getRow(0).getCell(1).getStringCellValue()),true,' ');
		String dheader3 = CaseUtils.toCamelCase((sheet2.getRow(0).getCell(2).getStringCellValue()),true,' ');
		int ActualCount1=0;
		int a1=report_CycleName.defects_rowsize;
		System.out.println("Defects Count: "+a1);
		for (int i=1; i<=a1;i++)
		{
			try {
			String id= sheet2.getRow(i).getCell(0).getStringCellValue().trim();
			if(id.equals("") || id.equals(" "))
			continue;
			else
				ActualCount1++;
			}
			finally{
				continue;
			}
		}
	html=html+"<table id = tb2>"
	+ "<tbody><Caption style=text-align:left;color:Black;text-decoration:underline;font-weight:bold>Defect Summary:</Caption>"
	+ "<tr bgcolor=#0B5394>"
	+ "<td style=color:White;text-align:center;font-weight:bold>"+dheader1+"</td>"
	+ "<td style=color:White;text-align:center;font-weight:bold>"+dheader2+"</td>"
	+ "<td style=color:White;text-align:center;font-weight:bold>"+dheader3+"</td>"
	+ "</tr>";
	for(int i =1; i <=ActualCount1; i++ )
	{
		XSSFRow row=sheet2.getRow(i);
		if(row!=null)
		{
		String defdata1 = sheet2.getRow(i).getCell(0).getStringCellValue();
		String defdata2 = sheet2.getRow(i).getCell(1).getStringCellValue();
		String defdata3 = sheet2.getRow(i).getCell(2).getStringCellValue();
		html = html+"<tr>"
						+ "<td style=color:Black;text-align:center;>"+defdata1+"</td>"
						+ "<td style=color:Black;text-align:left;>"+defdata2+"</td>"
						+ "<td style=color:Black;text-align:center;>"+defdata3+"</td>"
						+ "</tr>"
						+ "</tbody></table><br>";
		}
		}	
	}
	html= html+ "</body></html>";
	File fw = new File ("./result.html");
	BufferedWriter bw= new BufferedWriter(new FileWriter(fw));
	bw.write(html);
	bw.close();	
	workbook.close();
	}
}

