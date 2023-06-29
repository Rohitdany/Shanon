package Parameterization;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class XLWriting {
public static void main(String[] args) throws Throwable {
	// open file in java readable format
				FileInputStream fis = new FileInputStream(".\\\\src\\\\main\\\\resources\\\\ExcelFile.xlsx");
				// create workbookfactory
				Workbook wb = WorkbookFactory.create(fis);
				// open the sheet
				Sheet sh = wb.getSheet("rainbow");
				// open the row
				Row rw = sh.createRow(4);
				// open coloumn
				Cell ce = rw.createCell(8);
				ce.setCellValue("ReVISION");
				FileOutputStream fos=new FileOutputStream(".\\\\src\\\\main\\\\resources\\\\ExcelFile.xlsx");
				//write data in xl
				wb.write(fos);
				System.out.println("Data is written in xl");
				wb.close();
}
}
