package ru.vaszol.javaexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelApp {

	public static void main(String[] args) throws IOException{
		Workbook wb = new HSSFWorkbook(); //объект excel книги
		Sheet sheet0=wb.createSheet("Издатели"); //создаёт страницу в книге
		Sheet sheet1=wb.createSheet("Книги"); //создаёт страницу в книге
		Sheet sheet2=wb.createSheet("Авторы"); //создаёт страницу в книге
		Sheet sheet3=wb.createSheet(WorkbookUtil.createSafeSheetName("валоыврало?*Г:?**№?*?*")); //создаёт страницу в книге с нестандартным именем

		FileOutputStream fos = new FileOutputStream("my.xls"); //
		
		wb.write(fos);
		fos.close();
	}
}
