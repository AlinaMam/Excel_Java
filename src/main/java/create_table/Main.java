package create_table;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();//объект книги
        Sheet sheet1 = wb.createSheet("Издатели");//лист книги

        Row row1 = sheet1.createRow(3);//создаем строку 3
        Cell cell1 = row1.createCell(4);//создаем ячуйку 4 в строке 3 (0, 1, 2, 3, 4)
        cell1.setCellValue("O'Reilly");//создаем значение

        Sheet sheet2 = wb.createSheet("Книги");
        Row row2 = sheet2.createRow(0);//создаем строку 3
        Cell cell2 = row2.createCell(0);//создаем ячуйку 4 в строке 3 (0, 1, 2, 3, 4)

        Row row3 = sheet2.createRow(1);//создаем строку 3
        Cell cell3 = row3.createCell(3);
        cell2.setCellValue("Война и Мир");//создаем значение
        cell3.setCellValue("Евгений Онегин");

        Sheet sheet3 = wb.createSheet("Авторы");
        Sheet sheet4 = wb.createSheet(WorkbookUtil.createSafeSheetName("вапр56!"));//если имя листа состоит из каки=то спец.символов

        try (FileOutputStream fos = new FileOutputStream("my.xls")) {//создаем файл, куда будем все это записывать
            wb.write(fos);//записываем книгу в поток
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}