package ru.vaszol.javaexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelApp {

	public static void main(String[] args) throws IOException{
		Workbook wb = new HSSFWorkbook(); //объект excel книги
		Sheet sheet=wb.createSheet("MySheet"); //создаёт страницу в книге
		
		FileOutputStream fos = new FileOutputStream("my.xls"); //
		
		wb.write(fos);
		fos.close();
	}
}
