import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XSSFExample {

    public static void main(String[] args) throws IOException {
        XSSFExample excelObject = new XSSFExample();
        //excelObject.writeToExcel();
        excelObject.readFromExcel();
    }
    public void writeToExcel() throws IOException {
        //create an Excel workbook
        XSSFWorkbook wBook = new XSSFWorkbook();
        //create a sheet
        XSSFSheet wSheet = wBook.createSheet("sheet1");
        //create the first row
        XSSFRow wRow = wSheet.createRow(0);
        //create the first column
        XSSFCell wCell = wRow.createCell(0);
        //write sth to the first cell
        wCell.setCellValue("ID");
        //save workbook to file
        wBook.write(new FileOutputStream("SimpleXSSF.xslx"));
        //close the workbook
        wBook.close();
    }

    public void readFromExcel() throws IOException {
        //create an object of file input stream
        FileInputStream fis = new FileInputStream("simpleXSSF.xslx");
        //create a workbook object
        XSSFWorkbook wBook = new XSSFWorkbook(fis);
        //refer to the worksheet
        XSSFSheet wSheet = wBook.getSheet("sheet1");
        //get the first row
        XSSFRow wRow = wSheet.getRow(0);
        //get the first column
        XSSFCell wCell = wRow.getCell(0);
        //get the value
        String strValue = wCell.getStringCellValue();
        //print the value
        System.out.println(strValue);
        //close workbook
        wBook.close();
    }
}
