package Excel;


import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

static String FileLocation="./src/main/java/Files/Data.xlsx";
static String sheetName = "Sheet1";

public static void updateCellData(int rowNum, int colNum, String newData) throws Exception {

FileInputStream fis = new FileInputStream(FileLocation);
XSSFWorkbook workbook = new XSSFWorkbook(fis);
XSSFSheet sheet = workbook.getSheet(sheetName);
Row row = sheet.getRow(rowNum);
Cell cell = row.createCell(colNum);
// cell = row.getCell(colNum);
cell.setCellValue(newData);

FileOutputStream fos = new FileOutputStream(FileLocation);
workbook.write(fos);

fos.close();

}
}

