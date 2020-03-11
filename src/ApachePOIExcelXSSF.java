/** 
 * @author Jake Jeong
 * Much of this code was taken from:
 * https://mkyong.com/java/apache-poi-reading-and-writing-excel-file-in-java/
 * */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.ss.usermodel.*;
    
public class ApachePOIExcelXSSF {
	
	private static final String FILE_NAME = "studentListing.xlsx";
	
	/**
	 * Creates a simple Excel workbook file populated with test values.
	 */
	public static void write() {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Students");
		
		Object[][] studentLists = {
				{"Wightman", "INFO1103", 98},
				{"Astley", "MUS1009", 69},
				{"Riordan", "ENG2083", 87}
		};
		
		int rowNum = 0, colNum = 0;
		System.out.println("Creating Excel file...");
		
		for (Object[] studentList : studentLists) {
			Row row = sheet.createRow(rowNum++);
			
			colNum = 0;
			for (Object field : studentList) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String)
					cell.setCellValue((String)field);
				if (field instanceof Integer)
					cell.setCellValue((Integer)field);
			}
		}
		
	    try {
	        FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
	        workbook.write(outputStream);
	        workbook.close();
	    } catch (FileNotFoundException e) {
	        e.printStackTrace();
	    } catch (IOException e) {
	        e.printStackTrace();
	    }

	    System.out.println("Done!\n");
	}

	/**
	 * Reads from the created Excel workbook file and prints to console output.
	 */
	public static void read() {
		System.out.println("> Reading studentListing.xlsx:");
		try {
			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet studentListSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = studentListSheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellType() == CellType.STRING)
						System.out.print(currentCell.getStringCellValue() + " | ");
					else if (currentCell.getCellType() == CellType.NUMERIC)
						System.out.print(currentCell.getNumericCellValue() + " | ");
				}
				System.out.println();
			}
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
