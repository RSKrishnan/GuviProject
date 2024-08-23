package org.JATTask8;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingAndReadingUsingApachaPOI {

	public static void main(String[] args) throws IOException {

		File file = new File("Utils\\Task1.xlsx");
		// Creating a workbook
		XSSFWorkbook wbk = new XSSFWorkbook();
		// creating a sheet with name as Sheet1
		XSSFSheet createSheet = wbk.createSheet("Sheet1");
		// create the row
		XSSFRow row0 = createSheet.createRow(0);
		// Value had been set in cell '0'
		row0.createCell(0).setCellValue("Name");
		// Value had been set in cell '1'
		row0.createCell(1).setCellValue("Age");
		// Value had been set in cell '2'
		row0.createCell(2).setCellValue("Email");

		XSSFRow row1 = createSheet.createRow(1);
		row1.createCell(0).setCellValue("John Doe");
		row1.createCell(1).setCellValue("30");
		row1.createCell(2).setCellValue("John@test.com");

		XSSFRow row2 = createSheet.createRow(2);
		row2.createCell(0).setCellValue("Jane Doe");
		row2.createCell(1).setCellValue("28");
		row2.createCell(2).setCellValue("John@test.com");

		XSSFRow row3 = createSheet.createRow(3);
		row3.createCell(0).setCellValue("Bob Smith");
		row3.createCell(1).setCellValue("35");
		row3.createCell(2).setCellValue("jacky@example.com");

		XSSFRow row4 = createSheet.createRow(4);
		row4.createCell(0).setCellValue("Swapnil");
		row4.createCell(1).setCellValue("37");
		row4.createCell(2).setCellValue("swapnil@example.com");
		// Setting up output byte stream
		FileOutputStream fileout = new FileOutputStream(file);
		// writing the data in excel sheet
		wbk.write(fileout);
		//workbook is closed
		wbk.close();
		//calling the method of readexcel
		ReadExcel();
	}
	public static void ReadExcel() throws IOException {
		// Referring to the file for reading and writing
		File file = new File("Utils\\Task1.xlsx");
		// To read the data from a file in the form of sequence of bytes or images
		FileInputStream input = new FileInputStream(file);
		// Opening the workbook
		XSSFWorkbook wbk = new XSSFWorkbook(input);
		// Opening the sheet with sheetname
		XSSFSheet sheetAt = wbk.getSheet("Sheet1");
		// Gives the tot no of rows from the sheet
		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < physicalNumberOfRows; i++) {
			// opening the row
			XSSFRow row = sheetAt.getRow(i);
			// reading no of cells in row
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < physicalNumberOfCells; j++) {
				// opening the cell to read
				XSSFCell cell = row.getCell(j);
				System.out.print("\t");
				//doing the conversion of cell data to string from numeric, boolean
				switch (cell.getCellType()) {
				case NUMERIC:
					System.out.print(String.valueOf(cell.getNumericCellValue()));
				case BOOLEAN:
					System.out.print(String.valueOf(cell.getBooleanCellValue()));
				default:
					System.out.print(cell.getStringCellValue());
				}
			}
			System.out.println(" ");
		}
		wbk.close();
	}
}
