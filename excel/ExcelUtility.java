package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelUtility {

	public static Object[][] readExcel() throws IOException {
		String location="C:\\Users\\Subiksha\\Desktop\\day10\\excel.xlsx";
		FileInputStream fis=new FileInputStream(location);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0);
		int total=sheet.getPhysicalNumberOfRows();
		System.out.println("Total number of rows: "+total);
        int column=sheet.getRow(0).getLastCellNum();
        System.out.println("Total columns: "+column);
        XSSFRow row=sheet.getRow(1);
        XSSFCell cell=row.getCell(0);
        Object[][]data=new Object[total-1][column];
        for(int i=1;i<total;i++)
        {
        	XSSFRow row1=sheet.getRow(i);
        	for(int j=0;j<column;j++){
        		XSSFCell cell1=row1.getCell(j);
        		data[i-1][j]=cell.getNumericCellValue();
        		System.out.println(cell1.getNumericCellValue());
        	}fis.close();
        	workbook.close();
        }
		return data;
        } 
	public static void updateExcel() throws IOException {
		String location="C:\\Users\\Subiksha\\Desktop\\day10\\excel.xlsx";
		FileInputStream fis=new FileInputStream(location);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0);
		XSSFRow newrow=sheet.createRow(1);
		XSSFCell newcell=newrow.createCell(1);
		newcell.setCellValue("updated");
		System.out.println("document updated successfully");
		FileOutputStream fos=new FileOutputStream(location);
		workbook.write(fos);
		workbook.close();
		fos.close();
	}
        


}

