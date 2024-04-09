package com.training.excelsheet;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXL {
	 static XSSFWorkbook excelWorkbook;
	static XSSFSheet excelsheet;

	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub
		//get the filepath
		String filepath="C:\\Users\\shiga\\OneDrive\\Documents\\ReadExcel.xlsx";
		//add this file in to inputstream
		FileInputStream excelfileinput=new FileInputStream(filepath);
		//file inputstream convert this in to a excelworkbook
		//two important class needed for this,XSSFWorkbook and XSSFSheet, 
		//for this classes I need 2 dependencies,Apache poi dependency and Apache poi-ooxml dependency
		//first load the workbook then load the sheets//the workbook from that i am taking sheet
		excelWorkbook= new XSSFWorkbook(excelfileinput);
		excelsheet =excelWorkbook.getSheet("test1");
	System.out.println(	excelsheet.getRow(0).getCell(0));
	int totalRows=excelsheet.getLastRowNum();
	for(int i=0;i<=totalRows;i++) {//for rows
		for(int j=0;j<2;j++) {//for coloumn hardcoding
		
			System.out.print(excelsheet.getRow(i).getCell(j));
			System.out.print("\t");
	}
		System.out.println();
	}

	}}
