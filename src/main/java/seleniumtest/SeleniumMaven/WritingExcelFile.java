package seleniumtest.SeleniumMaven;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelFile {
	String Filepath = "C:\\Users\\HP\\Desktop\\skinput.xlsx";
	public void WriteExcel() {
		FileOutputStream FOS = new FileOutputStream(Filepath);
		XSSFWorkbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("login");
		
	}

	public static void main(String[] args) {
		

	}

}
