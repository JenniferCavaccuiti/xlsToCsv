import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

public class Main {

    public static void main(String[] args) throws IOException {
        File file = new File("xls/test.xlsx");
        FileInputStream inputStream = new FileInputStream(file);
        Workbook wb = new XSSFWorkbook(inputStream);

        Sheet sheet = wb.getSheetAt(0);
        removeRow(sheet);
        removeColumn(sheet);
        changePriceFormat(sheet);
        createCsv(wb);
    }

    private static void changePriceFormat(Sheet sheet) {
        for (Row row : sheet) {
            Cell cell = row.getCell(12);
            row.getCell(12).setCellValue(String.valueOf(cell));
        }
    }



    public static void removeColumn(Sheet sheet) {
        for (Row row : sheet) {
            int i = 4;
            while (i <= 13) {
                if (i != 12) {
                    Cell cell = row.getCell(i);
                    row.removeCell(cell);
                }
                i++;
            }
        }
    }

    public static void removeRow(Sheet sheet) {

        Cell cell;
        Iterator<Row> rowIterator = sheet.iterator();
        int lastLine = 0;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            cell = row.getCell(0);
            if (cell.getStringCellValue().equals("")) {
                lastLine = row.getRowNum();
            }
            if (row.getRowNum() == 0 || lastLine != 0) {
                rowIterator.remove();
            }
        }
    }

    public static void createCsv(Workbook wb) throws FileNotFoundException {
        DataFormatter formatter = new DataFormatter();
        PrintStream out = new PrintStream(new FileOutputStream("xls/test.csv"),
                true, StandardCharsets.UTF_8);
        for (Sheet sheet : wb) {
            for (Row row : sheet) {
                boolean firstCell = true;
                for (Cell cell : row) {
                    if (!firstCell) out.print(',');
                    String text = formatter.formatCellValue(cell);
                    out.print(text);
                    firstCell = false;
                }
                out.println();
            }
        }
    }
}