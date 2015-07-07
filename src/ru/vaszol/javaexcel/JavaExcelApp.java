package ru.vaszol.javaexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelApp {

	public static void main(String[] args) throws IOException{
		Workbook wb = new HSSFWorkbook(); //объект excel книги
		Sheet sheet0=wb.createSheet("Издатели"); //создаёт страницу в книге

        Row row0 = sheet0.createRow(3); //создали строку в листе
        Cell cell0=row0.createCell(4); //создали ячейку в строке
        cell0.setCellValue("O'Reilly"); //записываем данные в ячейку

		Sheet sheet1=wb.createSheet("Произведения"); //создаёт страницу в книге

        Row row1 = sheet1.createRow(0); //создали строку в листе
        Cell cell1 = row1.createCell(0); //создали ячейку в строке
        cell1.setCellValue("Война и мир"); //записываем данные в ячейку

        Row row2 = sheet1.createRow(1); //создали строку в листе
        Cell cell2 = row2.createCell(3); //создали ячейку в строке
        cell2.setCellValue("Евгений онегин"); //записываем данные в ячейку

		Sheet sheet2=wb.createSheet("Авторы"); //создаёт страницу в книге
		Sheet sheet3=wb.createSheet(WorkbookUtil.createSafeSheetName("валоыврало?*Г:?**№?*?*")); //создаёт страницу в книге с нестандартным именем

		FileOutputStream fos = new FileOutputStream("my.xls"); //
		
		wb.write(fos);
		fos.close();
	}
}
