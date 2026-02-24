package excel_methods;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel_files {

	public static void main(String[] args) throws Throwable {
		// read excel path
		FileInputStream fi = new FileInputStream("C:\\Users\\ginee\\OneDrive\\Documents\\Book1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		XSSFSheet ws = wb.getSheet("EMployee");
		int Row = ws.getLastRowNum();
		
		//XSSFRow wr = ws.getRow(Row);
		System.out.println(Row);
		
		int cells = ws.getRow(0).getLastCellNum();
		System.out.println(cells);
		
		for(int i=1;i<=Row;i++) {
		
		String fname = ws.getRow(i).getCell(0).getStringCellValue();
		String mname = ws.getRow(i).getCell(1).getStringCellValue();
		String lname = ws.getRow(i).getCell(2).getStringCellValue();
		int emp = (int) ws.getRow(i).getCell(3).getNumericCellValue();
		
		System.out.println(fname+"  "+mname+"    "+lname+"    "+emp);
		ws.getRow(i).createCell(4).setCellValue("pass");
		
		XSSFCellStyle style = wb.createCellStyle();
		XSSFFont font = wb.createFont();
		style.setFont(font);
		
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		ws.getRow(i).getCell(4).setCellStyle(style);
		}
		
		FileOutputStream fo = new FileOutputStream("D:/myres.xlsx");
		wb.write(fo);
		fo.close();
		wb.close();
		
		

		
		
	
		

	}

}
