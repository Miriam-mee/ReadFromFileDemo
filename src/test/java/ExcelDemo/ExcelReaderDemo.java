package ExcelDemo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class ExcelReaderDemo {
    @Test
    public void testReadExcelUsingForLoop() throws IOException {
        String filePath = System.getProperty("user.dir") + "\\Datasource\\Student.xlsx";
        FileInputStream fis = new FileInputStream(filePath);
        Workbook book = new XSSFWorkbook(fis);
        Sheet sheet = book.getSheet("Student Info");
        int lastRowNum = sheet.getLastRowNum();
        int lastCellNum = sheet.getRow(0).getLastCellNum();
        System.out.println("Row count " + lastRowNum);
        System.out.println("Column count " + lastCellNum);
        for (int r = 0; r < lastRowNum; r++) {
            Row row = sheet.getRow(r);
            for (int col = 0; col < lastCellNum; col++) {
                Cell cell = row.getCell(col);
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case FORMULA:
                        System.out.println(cell.getNumericCellValue());
                }
                System.out.print("   |    ");
            }
            System.out.println();
        }

    }

    @Test
    public void testReadExcelUsingIterator() throws IOException {
        String filePath = System.getProperty("user.dir") + "\\Datasource\\Student.xlsx";
        FileInputStream fis = new FileInputStream(filePath);
        Workbook book =null;
        try{
            fis = new FileInputStream(filePath);
            book= new XSSFWorkbook(fis);
        }
        catch (Exception e){
            e.printStackTrace();
        }
        finally {
            if(fis !=null)
                fis.close();
        }
        Sheet sheet = book.getSheet("Student Info");
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            Iterator<Cell> cellIterator = row.iterator();
            while (cellIterator.hasNext()) ;
            Cell cell = cellIterator.next();
            switch (cell.getCellType()) {
                case STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    System.out.print(cell.getNumericCellValue());

            }
            System.out.print("   |    ");
        }
        System.out.println();
    }




    @Test
    public void testWriteDemoUsingForEachLoop() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employee Info");
        Object EmployeeInfo[][] = {
                {"Name", "Age", "County", "Subject"},
                {"Mimi", 40, "Nigeria", "Java"},
                {"Esther", 30, "Nigeria", "Python"},
                {"Abas", 25, "UK", "Java"}
        };
        int rowCount = 0;

        for (Object student[] : EmployeeInfo) {
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
        String filePath = System.getProperty("user.dir") + "\\Datasource\\Student2.xlsx";
        FileOutputStream fos = null;

        fos = new FileOutputStream(filePath);
        workbook.write(fos);
        fos.close();
        System.out.println("Student info written to file successfully");


    }

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
        String filePath = System.getProperty("user.dir") + "\\Datasource\\Student.xlsx";
        FileOutputStream fos = null;

        fos = new FileOutputStream(filePath);
        workbook.write(fos);
        fos.close();
        System.out.println("Student info written to file successfully");

    }
}