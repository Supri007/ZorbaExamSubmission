import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WriteExcelFile {

    public static void main(String[] args) throws Exception {
        File file = new File("C:\\SQl Exam\\July 21\\Employee.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        System.out.println(xssfSheet.getLastRowNum() + 1);
        int newRowCount = xssfSheet.getLastRowNum() + 1;
        Row row = xssfSheet.getRow(1);
        int cellCount = row.getLastCellNum();
        int totalCellCount = cellCount + 3;
        //Create new row
        Row newRow = xssfSheet.createRow(cellCount);
        Cell cell = null;

        for (int i = cellCount; i < totalCellCount; i++)
        {
            String colName = xssfSheet.getRow(0).getCell(i).getStringCellValue();
            System.out.println(colName);
            System.out.println(i);
            for (int j = i; j < 3; j++){
                switch(j){
                    case 6:
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Null");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1001);
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1004);
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1004);
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1001);
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1005);
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue(1001);
                        break;
                    case 7:
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("R&D");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("R&D");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = newRow.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        break;
                    case 8:
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(60);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(20);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(30);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(40);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(20);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(15);
                        cell = newRow.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(25);
                        break;
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(file);
        xssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Successfully Written new row to excel file..");
    }
}
