import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;

public class Practics {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook(); // создание документа в памяяти
        Sheet sheet = wb.createSheet("Лист 1"); // создание лимста в файле
        Row row = sheet.createRow(2); // выбор записи в 3 столбец
        Cell cell = row.createCell(2); // выбор записи в 3 ячейку
        cell.setCellValue("Текст"); // запись текста в выбранную ранее чейку
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Courier New"); // выбор стиля текста в выбранной ранее ячейке
        font.setBold(true); // выбор жирного шрифта в выбранной ранее ячейке
        style.setFont(font);
        cell.setCellStyle(style);
        System.out.println(wb.getSheetAt(0).getRow(2).getCell(2).getStringCellValue()); // вывод текста в редактируемой ранее ячейке
        FileOutputStream fos = new FileOutputStream("Practics.xls"); // запись документа в файл
        wb.write(fos);
        fos.close();
    }
}
