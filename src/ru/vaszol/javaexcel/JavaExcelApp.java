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
