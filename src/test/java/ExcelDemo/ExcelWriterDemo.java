package ExcelDemo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class ExcelWriterDemo {
    @Test
    public void testWriteDataUsingArrayList() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Student2 Info");
        List<Object []> stdList = Arrays.asList(new Object[]{"Name", "Age", "Country", "Subject"},
                new Object[]{"Miriam", 35, "Uk", "Java"},
                new Object[]{"Esther", 25, "Uk", "Python"}
        );
        int rowCount = 0;

        for (Object student[] : stdList) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Object value : student) {
                Cell cell = row.createCell(columnCount++);
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                }

            }
        }
        String filePath = System.getProperty("user.dir") + "\\Datasource\\Student1.xlsx";
        FileOutputStream fos = null;

        fos= new FileOutputStream(filePath);
        workbook.write(fos);
        fos.close();
        System.out.println("Student info written to file successfully");

    }
}
