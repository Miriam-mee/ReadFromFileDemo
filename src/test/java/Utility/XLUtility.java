package Utility;

import lombok.Builder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class XLUtility {
private FileInputStream fi;
private FileOutputStream fo;
private XSSFWorkbook workbook;
private XSSFSheet sheet;
private XSSFRow row;
private XSSFCell cell;
private DataFormatter dataFormatter;
private FillPatternType fillPatternType;
private IndexedColors indexedColors;
private CellStyle cellStyle;
private String path;

public XLUtility(String path){
    this.path= path;
}

public int getRowCount(String sheetname) throws IOException {
    fi=new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet=workbook.getSheet(sheetname);
    int rowCount= sheet.getPhysicalNumberOfRows();
    workbook.close();
    fi.close();
    return rowCount;
}

public int getCellCount(String sheetname, int rownum) throws IOException {
    fi= new FileInputStream(path);
    workbook= new XSSFWorkbook(fi);
    sheet=workbook.getSheet(sheetname);
    row=sheet.getRow(rownum);
    int CellCount=row.getLastCellNum();
    workbook.close();
    fi.close();
    return CellCount;
}

public String getCellData(String sheetname, int rownum, int column) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetname);
    row = sheet.getRow(rownum);
    cell = row.getCell(column);


    DataFormatter formatter = new DataFormatter();
    String data = null;
    try {
        data = formatter.formatCellValue(cell);
//        Returns the formatted value of the cell as a string
    } catch (Exception e) {

        data = " ";
    }

    workbook.close();
    fi.close();
    return data;

}

public void setCellDate(String sheetname, int rownum, int column,String data) throws IOException {
    File xlfile = new File(path);
    if (!xlfile.exists()) {
        //       if file does not exist, it creates a new one
        workbook = new XSSFWorkbook();
                fo=new FileOutputStream(path);
                workbook.write(fo);
    }
    fi=new FileInputStream(path);
    workbook=new XSSFWorkbook(fi);

    if(workbook.getSheetIndex(sheetname)==1)
//        if sheet does not exist, it creates new one
        workbook.createSheet(sheetname);

    sheet=workbook.getSheet(sheetname);

    if(sheet.getRow(rownum)==null)
        //        if row does not exist, it creates new one
        sheet.createRow(rownum);
}

public void fillGreenColor(String sheetname, int rownum, int column) throws IOException {
    fi = new FileInputStream(path);
    workbook = new XSSFWorkbook(fi);
    sheet = workbook.getSheet(sheetname);
    row=sheet.getRow(rownum);
    cell= row.getCell(column);

    cellStyle = workbook.createCellStyle();
    cellStyle.setFillForegroundColor(indexedColors.GREEN.getIndex());
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    cell.setCellStyle(cellStyle);
    workbook.write(fo);
    workbook.close();
    fi.close();
    fo.close();
}
}
