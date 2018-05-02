package com.example.demo4.poiExperiments;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

//My tools for different excell stuff
public class ExcelNormalizer {

    //remove O from unitNumbers at all sheets of workbook.
    public static void doPretty(){
        try {
            FileInputStream fis = new FileInputStream("G:\\excelFiles\\xx7.xlsx");
            Workbook wb = new XSSFWorkbook(fis);

            for (int sheetIdx = 2; sheetIdx <wb.getNumberOfSheets(); sheetIdx++) {
                Sheet sheet  = wb.getSheetAt(sheetIdx);
                Cell  cell = sheet.getRow(0).getCell(0);
                List<Cell> cellList = new ArrayList<>();
                deleteOu(cell);
                for (int rowIter = 1; rowIter < 29; rowIter++) {
                    deleteOu(sheet.getRow(rowIter).getCell(1));
                                  }
                for (int rowIter = 1; rowIter < 35; rowIter++) {
                    deleteOu(sheet.getRow(rowIter).getCell(5));
                }
                for (int rowIter = 1; rowIter < 13; rowIter++) {
                    deleteOu(sheet.getRow(rowIter).getCell(9));
                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream("G:\\excelFiles\\xx8.xlsx");
            wb.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    //удаляет букву О которую женя пишет в комплектность from Cell
    public static void deleteOu(Cell cell){
            if (cell.getCellTypeEnum() != CellType.NUMERIC) {
                String cellValue = cell.getStringCellValue();
                cellValue = cellValue.replace("О", "0"); //это русское О
                cell.setCellValue(cellValue);
            }


    }

    public static void disableWarnings() {
        System.err.close();
        System.setErr(System.out);
    }
}
