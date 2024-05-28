package org.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class base2 {
 
	   public static void createNewExcelFile(int rowNum, int cellNum,String writeData) throws IOException {
		   File f = new File("C:\\Users\\PARIMALA DEVI\\eclipse-workspace\\DataDriven\\Excel\\newfile.xlsx");
		   Workbook wb = new XSSFWorkbook();
		   Sheet newsheet = wb.createSheet("Datas");
		   Row newrow = newsheet.createRow(rowNum);
		   Cell newcell = newrow.createCell(cellNum);
		   newcell.setCellValue(writeData);
		   FileOutputStream fos = new FileOutputStream(f);
		   wb.write(fos);
	   }
		   public static void createCell(int rowNum,int cellNum,String newData ) throws IOException {
			   File f = new File("C:\\Users\\PARIMALA DEVI\\eclipse-workspace\\DataDriven\\Excel\\newfile.xlsx");
			   FileInputStream fis = new FileInputStream(f);
			   Workbook wb = new XSSFWorkbook(fis);
			   Sheet s = wb.getSheet("Datas");
			   Row r = s.getRow(rowNum);
			   Cell c = r.createCell(cellNum);
			   c.setCellValue(newData);
			   FileOutputStream fos = new FileOutputStream(f);
			   wb.write(fos);
	 }
		   public static void createRow(int creRow,int creCell,String newData ) throws IOException {
			   File f = new File("C:\\Users\\PARIMALA DEVI\\eclipse-workspace\\DataDriven\\Excel\\newfile.xlsx");
			   FileInputStream fis = new FileInputStream(f);
			   Workbook wb = new XSSFWorkbook(fis);
			   Sheet s = wb.getSheet("Datas");
			   Row r = s.createRow(creRow);
			   Cell c = r.createCell(creCell);
			   c.setCellValue(newData);
			   FileOutputStream fos = new FileOutputStream(f);
			   wb.write(fos);
		   
		   }
		

	}
	


