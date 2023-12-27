package styles_and_fonts;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.TextAlign;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.apache.poi.xssf.usermodel.TextAlign.CENTER;

public class Main {
    public static void main(String[] args) throws IOException {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Лист 01");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Привет");

        CellStyle style = wb.createCellStyle();//создаем style
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);//цвет ячейки, сплошной цвет, это тип заливки
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());// цвет ячейки
        style.setAlignment(HorizontalAlignment.CENTER);//выравнивание по горизонтали
        style.setVerticalAlignment(VerticalAlignment.TOP);//выравнивание по верхнему краю
        style.setBorderBottom(BorderStyle.DASH_DOT_DOT);//нижняя граница ячейки
        style.setBottomBorderColor(IndexedColors.GREEN.getIndex());//цвет нижней границы ячейки

        cell.setCellStyle(style); //применим стиль к ячейке

        //работает со шрифтом
        Font font = wb.createFont();
        font.setFontName("Courier New");//имя шрифта
        font.setFontHeightInPoints((short) 15);//размер шрифта
        font.setBold(true);//;жирный шрифт
        font.setStrikeout(true);//текст зачеркнутый
        font.setUnderline(Font.U_SINGLE);//одинарное подчеркивание
        font.setColor(IndexedColors.RED.getIndex());//цвет шрифта

        //применис font к style
        style.setFont(font);

        FileOutputStream fos = new FileOutputStream("my2.xls");
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
