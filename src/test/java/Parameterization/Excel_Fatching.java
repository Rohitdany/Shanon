package Parameterization;

import java.io.FileInputStream;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel_Fatching {
	public static void main(String[] args) throws Throwable {
		System.out.Println("I am updated from Github");
		// open file in java readable format
		FileInputStream fis = new FileInputStream(".\\src\\main\\resources\\ExcelFile.xlsx");
		// create workbookfactory
		Workbook wb = WorkbookFactory.create(fis);
		// open the sheet
		org.apache.poi.ss.usermodel.Sheet sh = wb.getSheet("rainbow");
		//		// open the row
		 Row rw = sh.getRow(1);
		// open coloumn
		Cell ce = rw.getCell(0);
		String data = ce.getStringCellValue();
		System.out.println(data);
	}
}
