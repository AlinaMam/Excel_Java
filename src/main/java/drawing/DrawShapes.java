package drawing;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class DrawShapes {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Картинка");

        //создаем художника
        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
        //область рисования фигуры, указываем диапазон, 2 координаты (левый верхний угол и правый нижний)
        HSSFClientAnchor anchor = new HSSFClientAnchor();
        anchor.setCol1(2);
        anchor.setRow1(2);
        anchor.setCol2(10);
        anchor.setRow2(10);

        //создадим фигуры
        HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);
       /* shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);

        shape.setLineStyleColor(255,0,0);//цвет линии
        shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT*3);//толщина линии
        shape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHDOTGEL);//тип линии*/

        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);

        shape.setLineStyleColor(255,0,0);
        shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT*3);
        shape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHDOTGEL);
        shape.setFillColor(0,0, 255);

        FileOutputStream fos = new FileOutputStream("abc.xls");
        wb.write(fos);
        fos.close();
    }
}
