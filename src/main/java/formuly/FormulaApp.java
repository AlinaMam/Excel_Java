package formuly;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FormulaApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Формулы");
        Row row0 = sheet.createRow(0);

        /*Cell cell0 = row0.createCell(0);
        cell0.setCellValue(2);

        Cell cell1 = row0.createCell(1);
        cell1.setCellValue(7);

        Cell cell2 = row0.createCell(2);
        cell2.setCellFormula("A1*B1");*/

        Row row1 = sheet.createRow(3);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue(1);

        Row row2 = sheet.createRow(4);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(2);

        Row row3 = sheet.createRow(5);
        Cell cell3 = row3.createCell(0);
        cell3.setCellValue(3);

        Row row4 = sheet.createRow(6);
        Cell cell4 = row4.createCell(0);
        cell4.setCellValue(4);

        Row row5 = sheet.createRow(7);
        Cell cell5 = row5.createCell(0);
        cell5.setCellFormula("SUM(A3:A7)");

        FileOutputStream fos = new FileOutputStream("my1.xls");
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
