
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

  private static Workbook workbook;
  private static Sheet sheet;
  private static Cell cell;
  private static FileInputStream inputStream;
  private static FileOutputStream outputStream;
  private static String filePath = "C:\\Users\\azeem\\eclipse-workspace\\java-read-write-excel\\java-read-write-excel-file-using-apache-poi-master\\sample-xlsx-file.xlsx";

  
  public static void main(String[] args) throws IOException  {
	  ExcelUtils.readData(filePath, "Employee", 1, 4);	  
  }
  


  //Method to read data from an Excel file
  public static String readData(String filePath, String sheetName, int rowNum, int colNum) throws IOException {
    inputStream = new FileInputStream(filePath);
    workbook = new XSSFWorkbook(inputStream);
    sheet = workbook.getSheet(sheetName);
    cell = sheet.getRow(rowNum).getCell(colNum);
    inputStream.close();
    System.out.println(cell.toString());
    return cell.toString();

  }

  //Method to write data to an Excel file
  public static void writeData(String filePath, String sheetName, int rowNum, int colNum, String data) throws IOException {
    inputStream = new FileInputStream(filePath);
    workbook = new XSSFWorkbook(inputStream);
    sheet = workbook.getSheet(sheetName);
    cell = sheet.getRow(rowNum).getCell(colNum);
    cell.setCellValue(data);
    outputStream = new FileOutputStream(filePath);
    workbook.write(outputStream);
    outputStream.close();
    inputStream.close();
  }
}
