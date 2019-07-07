package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelData 
{
   public void writeExcel(String filename,String sheetname,int rownum,int colnum,String val) throws IOException
   {
	   FileInputStream fis=new FileInputStream(filename);
	   XSSFWorkbook wb= new XSSFWorkbook(fis);
	   XSSFSheet sheet=wb.getSheet(sheetname);
	   XSSFRow row=sheet.createRow(rownum);
	   XSSFCell cell=row.createCell(colnum);
	   cell.setCellValue(val);
	   FileOutputStream fileout=new FileOutputStream(filename);
	   wb.write(fileout);
	   
   }
}
