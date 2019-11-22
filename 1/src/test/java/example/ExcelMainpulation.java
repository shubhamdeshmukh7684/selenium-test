package example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMainpulation {

	public void readExcel()
	{
		File filepath=new File("C:\\Users\\Sush\\Desktop\\Test.xlsx");
		try {
			FileInputStream in = new FileInputStream(filepath);
			
			XSSFWorkbook wb = new XSSFWorkbook(in);
			XSSFSheet sheet= wb.getSheetAt(0);
			
			System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
			System.out.println(sheet.getRow(0).getCell(1).getStringCellValue());
			
			System.out.println("File Successfully Read..!!");
			Row r1 = sheet.createRow(1);
			Cell c1= r1.createCell(0);
			c1.setCellValue("Shubham");
			
			sheet.createRow(2).createCell(0).setCellValue("Direct Created");
			
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Sush\\Desktop\\Test.xlsx"));
            wb.write(out);
            System.out.println("File Successfully Written..!!");
            out.close();
            wb.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
	}
	
	public static void main(String args[])
	{
		ExcelMainpulation e = new ExcelMainpulation();
		e.readExcel();
	}
}
