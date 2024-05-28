package org.data;

import java.io.IOException;

public class WriteData extends base2 {
	public static void main(String[] args) throws IOException {
		
		createNewExcelFile(0, 0, "selenim");
		createCell(0, 1, "java");
		createCell(0, 2, "DataDriven");
		createCell(0, 3, "POM");

       createRow(1, 0, "appium");
       createCell(1, 1, "cucumber");
       createCell(1, 2, "junit");
       createCell(1, 3, "TestNG");
		
		
	}

}
