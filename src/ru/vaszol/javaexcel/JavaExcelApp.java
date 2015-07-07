package ru.vaszol.javaexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class JavaExcelApp {

    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy,MM,dd"); //формат даты

	public static void main(String[] args) throws IOException{
        //запись/создание
        Workbook wb0 = new HSSFWorkbook();
        Sheet sheet = wb0.createSheet("Формулы");
        Row row0 = sheet.createRow(0);

        Cell cell0 = row0.createCell(0);
        cell0.setCellValue(2);

        Cell cell1 = row0.createCell(1);
        cell1.setCellValue(7);

        Cell cell2 = row0.createCell(2);
        cell2.setCellFormula("A1+B1");


        Row row3 = sheet.createRow(3);
        Cell cell3 = row3.createCell(0);
        cell3.setCellValue(1);

        Row row4 = sheet.createRow(4);
        Cell cell4 = row4.createCell(0);
        cell4.setCellValue(2);

        Row row5 = sheet.createRow(5);
        Cell cell5 = row5.createCell(0);
        cell5.setCellValue(3);

        Row row6 = sheet.createRow(6);
        Cell cell6 = row6.createCell(0);
        cell6.setCellValue(4);

        Row row7 = sheet.createRow(7);
        Cell cell7 = row7.createCell(0);
        cell7.setCellFormula("SUM(A4:A7)");


        FileOutputStream fos=new FileOutputStream("c:/Users/vas/Documents/readXLS.xls");
        wb0.write(fos);
        fos.close();
        wb0.close();

        //чтение
        FileInputStream fis = new FileInputStream("c:/Users/vas/Documents/readXLS.xls");
		Workbook wb = new HSSFWorkbook(fis); //
        for(Row row:wb.getSheetAt(0)){
            for(Cell cell:row){
                CellReference celRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.println(celRef.formatAsString());
                System.out.println(" - ");
                System.out.println(getCellText(cell));
            }
        }
        /**
        System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(0)));
        System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(1)));
        System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(2)));
        System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(3)));
         */
        fis.close();
	}

    public static String getCellText(Cell cell){
        String result="";
        /** определяем формат ячейки*/
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_STRING:
                result=cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)){
                    result=sdf.format(cell.getDateCellValue());
                }else {
                    result=Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result=Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result=cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }
}
