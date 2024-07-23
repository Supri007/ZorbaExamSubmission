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
        Row row1 = xssfSheet.getRow(1);
        Row row2 = xssfSheet.getRow(2);
        Row row3 = xssfSheet.getRow(3);
        Row row4 = xssfSheet.getRow(4);
        Row row5 = xssfSheet.getRow(5);
        Row row6 = xssfSheet.getRow(6);
        Row row7 = xssfSheet.getRow(7);
        int cellCount = row1.getLastCellNum();
        System.out.println("Cell count: " + cellCount);
        int totalCellCount = cellCount + 3;

        Cell cell = null;

        for (int i = cellCount; i < totalCellCount; i++)
        {
            String colName = xssfSheet.getRow(0).getCell(i).getStringCellValue();
            System.out.println(colName);
            System.out.println(i);
            for (int j = i; j < i + 3; j++){
                switch(j){
                    case 5:
                        cell = row1.createCell(j, CellType.STRING);
                        cell.setCellValue("Null");
                        cell = row2.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1001);
                        cell = row3.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1004);
                        cell = row4.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1004);
                        cell = row5.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1001);
                        cell = row6.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1005);
                        cell = row7.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(1001);
                        break;
                    case 6:
                        cell = row1.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = row2.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = row3.createCell(j, CellType.STRING);
                        cell.setCellValue("R&D");
                        cell = row4.createCell(j, CellType.STRING);
                        cell.setCellValue("R&D");
                        cell = row5.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = row6.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        cell = row7.createCell(j, CellType.STRING);
                        cell.setCellValue("Finance");
                        break;
                    case 7:
                        cell = row1.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(60);
                        cell = row2.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(20);
                        cell = row3.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(30);
                        cell = row4.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(40);
                        cell = row5.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(20);
                        cell = row6.createCell(j, CellType.NUMERIC);
                        cell.setCellValue(15);
                        cell = row7.createCell(j, CellType.NUMERIC);
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
