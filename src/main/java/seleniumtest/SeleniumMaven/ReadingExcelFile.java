package seleniumtest.SeleniumMaven;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelFile {
	String Filepath = "C:\\Users\\HP\\Desktop\\sk.xlsx";
	static List<String> Arraylist= new ArrayList<String>();

	public void readExcelfile() throws IOException {
	FileInputStream FS = new FileInputStream(Filepath);
	XSSFWorkbook workbook = new XSSFWorkbook(FS);
Sheet sheet = workbook.getSheet("Class");
		int totalRows = sheet.getPhysicalNumberOfRows();
		for (int i=0;i<totalRows;i++) {
			Row inputRow = sheet.getRow(i);
			int totalColumn = inputRow.getLastCellNum();
			for(int j=0;j<totalColumn;j++) {
				Cell cellValue = inputRow.getCell(j);
				String actualValue = cellValue.getStringCellValue();
				Arraylist.add(actualValue);
			 }}
		for(String List :Arraylist) {
				System.out.print(List);
				System.out.print(" ");
				}}
public void WriteExcel(List<String> Arraylist) throws IOException {
		String Filepath1 = "C:\\Users\\HP\\Desktop\\skinput1.xlsx";
		FileOutputStream FOS = new FileOutputStream(Filepath1);
		XSSFWorkbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("login");
		int totalrow = Arraylist.size();

		for(int i=0;i<totalrow;i++) {
			Row inputrow = sheet.createRow(i);
			int inputcolumn = Arraylist.size();
			for(int j=0;j<inputcolumn;j++) {
				Cell cellvalue = inputrow.createCell(0);
				cellvalue.setCellValue(Arraylist.get(i)); 
			}
			
	}
		workbook.write(FOS);
		workbook.close();
		FOS.close();
		System.out.println("done");
		
	}

	public static void main(String[] args) throws IOException {
		ReadingExcelFile obj = new ReadingExcelFile();
		obj.readExcelfile();
		obj.WriteExcel(Arraylist);
	
	
		

	}

}
