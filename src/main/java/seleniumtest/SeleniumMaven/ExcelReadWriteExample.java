package seleniumtest.SeleniumMaven;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReadWriteExample {
    public static void main(String[] args) {
        String inputFile = "C:\\Users\\HP\\Desktop\\sk.xlsx"; // Path to the input Excel file
        String outputFile = "C:\\Users\\HP\\Desktop\\skinput.xlsx"; // Path to the output Excel file

        ArrayList<ArrayList<String>> data = readExcel(inputFile);

        if (data != null) {
            System.out.println("Data read from Excel:");
            for (ArrayList<String> row : data) {
                System.out.println(row);
            }

            writeExcel(outputFile, data);
            System.out.println("New Excel file created successfully.");
        }
    }

    public static ArrayList<ArrayList<String>> readExcel(String inputFile) {
        ArrayList<ArrayList<String>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            for (Row row : sheet) {
                ArrayList<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(cell.toString());
                }
                data.add(rowData);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return data;
    }

    public static void writeExcel(String outputFile, ArrayList<ArrayList<String>> data) {
        try (Workbook workbook = WorkbookFactory.create(true);
             FileOutputStream fos = new FileOutputStream(outputFile)) {

            Sheet sheet = workbook.createSheet("Sheet1");

            int rownum = 0;
            for (ArrayList<String> rowData : data) {
                Row row = sheet.createRow(rownum++);
                int cellnum = 0;
                for (String cellData : rowData) {
                    Cell cell = row.createCell(cellnum++);
                    cell.setCellValue(cellData);
                }
            }

            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

