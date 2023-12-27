package Size_cells;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class SizeApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Лист_01");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Новая ячейка");

//        sheet.setColumnWidth(0, 5000);//изменить размер ячейки
//        sheet.autoSizeColumn(0);//автоподгон ширины

        //с высотой работаем через row
        row.setHeightInPoints(15);

        //объединить ячейки
        sheet.addMergedRegion(new CellRangeAddress(0, 5, 0, 2));


        FileOutputStream fos = new FileOutputStream("my4.xls");
        wb.write(fos);
        fos.close();
        wb.close();


    }
}
