package grafic;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class LineChart {
    public static void main(String[] args) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) { //создаем книгу
            XSSFSheet sheet = wb.createSheet("linechart");//создаем лист
            final int NUM_OF_ROWS = 3; //колв-о строк 3
            final int NUM_OF_COLUMNS = 10; //кол-во колонок 10

            // Create a row and put some cells in it. Rows are 0 based.
            Row row; //строка
            Cell cell; //ячейка
            for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++) { //цикл по строкам
                row = sheet.createRow((short) rowIndex);
                for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
                    cell = row.createCell((short) colIndex);
                    cell.setCellValue(colIndex * (rowIndex + 1.0));
                }
            }
            //создаем фигуру
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            //область расположения графика, размечаем область
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 20, 40);

            //рисуем график
            XSSFChart chart = drawing.createChart(anchor);
            //создаем легенду
            XDDFChartLegend legend = chart.getOrAddLegend();
            //создаем позицию легенды
            legend.setPosition(LegendPosition.TOP_RIGHT);

            // Use a category axis for the bottom axis.
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM); //ось х
            bottomAxis.setTitle("x");
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT); //ось y
            leftAxis.setTitle("f(x)");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            //значение по x
            XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            //значение по y
            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xs, ys1); //рисуем графики
            series1.setTitle("2x", null);
            series1.setSmooth(false);
            series1.setMarkerStyle(MarkerStyle.STAR);
            XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) data.addSeries(xs, ys2);
            series2.setTitle("3x", null);
            series2.setSmooth(true);
            series2.setMarkerSize((short) 6);
            series2.setMarkerStyle(MarkerStyle.TRIANGLE);
            chart.plot(data);

            // if your series have missing values
            // chart.displayBlanksAs(DisplayBlanks.GAP);

            solidLineSeries(data, 0, PresetColor.CHARTREUSE);
            solidLineSeries(data, 1, PresetColor.TURQUOISE);

            // temporary workaround for https://bz.apache.org/bugzilla/show_bug.cgi?id=67510
            if (bottomAxis.hasNumberFormat()) bottomAxis.setNumberFormat("@");
            if (leftAxis.hasNumberFormat()) leftAxis.setNumberFormat("#,##0.00");

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("ooxml-line-chart.xlsx")) {
                wb.write(fileOut);
            }
        }
    }

    private static void solidLineSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(fill);
        XDDFChartData.Series series = data.getSeries(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setLineProperties(line);
        series.setShapeProperties(properties);
    }
}

