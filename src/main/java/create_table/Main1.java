package create_table;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class Main1 {
    //чтение данных из Excel файла
    public static SimpleDateFormat format = new SimpleDateFormat("yyyy.MM.dd");
    public static void main(String[] args) {
        try (Workbook wb = new XSSFWorkbook(new FileInputStream("/Users/alina/Desktop/Java/для примера.xlsx"))) {
          /*  String result = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
            System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(1)));
            System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(2)));
            System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(3)));*/
            for (Row row:wb.getSheetAt(0)) {
                for (Cell cell:row) {
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                    System.out.print(cellRef.formatAsString());
                    System.out.print(" - ");
                    System.out.println(getCellText(cell));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getCellText(Cell cell) {
        String result = null;
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = format.format(cell.getDateCellValue());
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                System.out.println("Dont't know this type, ggole it");
        }
        return result;
    }
}

