package com.example.demo4.poiExperiments;

import com.example.demo4.model.Unit;
import com.example.demo4.model.unitTools.UnitParams;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcellWorkerR1 {

    public static void disableWarnings() {
        System.err.close();
        System.setErr(System.out);
    }

    public static void main(String[] args) throws IOException {
        disableWarnings();

        Unit unit = new Unit("0", "0", LocalDate.now(), "0");
        List<Unit> unitList = new ArrayList<>();
        InputStream fis = null;
        fis = new FileInputStream("G:\\excelFiles\\xx7.xlsx");
        Workbook wb = new XSSFWorkbook(fis);
        UnitParams paramsForCol1 = new UnitParams(1, 29, 0, 4);
        UnitParams paramsForCol2 = new UnitParams(1, 35, 4, 8);
        UnitParams paramsForCol3 = new UnitParams(1, 13, 8, 12);
        Sheet sheet = wb.getSheetAt(3);
       /* for (int rowIdx = paramsForCol2.getStartAtRow(); rowIdx < paramsForCol1.getStopAtRow(); rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            unit = readUnitFromRow(row, paramsForCol2);
            unitList.add(unit);
        }*/
        unitList = readOneColumn(sheet, paramsForCol3);
        unitList.stream().forEach(x -> System.out.println(x));

    }


    public static List<Unit> readOneColumn(Sheet sheet, UnitParams unitParams){
        List<Unit> unitList = new ArrayList<>();
        for (int rowIdx = unitParams.getStartAtRow(); rowIdx < unitParams.getStopAtRow(); rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            unitList.add(readUnitFromRow(row,unitParams));
        }
        return unitList;
    }


    private static Unit readUnitFromRow(Row row, UnitParams params) {

        Unit unit = new Unit();
        Object object = "";
        for (int i = params.getStartAtCell(); i < params.getStopAtCell(); i++) {
            Cell cell = row.getCell(i);
            System.out.println(row.getRowNum() + " " + i + " " + cell);
            if (cell != null) {
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        object = cell.getNumericCellValue();
                        if (DateUtil.isCellDateFormatted(cell))
                            object = cell.getDateCellValue();
                        break;
                    case FORMULA:
                        object = cell.getNumericCellValue();
                        break;
                    case STRING:
                        object = cell.getStringCellValue();
                }
            } else {
                cell.setCellValue("");
            }

            int a = 0;
            Double d = 0.0;
            LocalDate ld = LocalDate.of(1975, 2, 2);
            if ((i == 0) || (i == 4) || (i == 8)) {
                unit.setUnitType((String) object);
            }
            if ((i == 1) || (i == 5) || (i == 9)) {
                if (object instanceof String)
                    unit.setUnitNumber((String) object);
                if (object instanceof Double) {
                    d = (Double) object;
                    a = d.intValue();
                    unit.setUnitNumber(Integer.toString(a));
                }
            }
            if ((i == 2) || (i == 6) || (i == 10)) {
                if (object instanceof Date) {
                    ld = ((Date) object).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                } else {

                }
                unit.setUnitReleaseDate(ld);
            }
            if ((i == 3) || (i == 7) || (i == 11)) {
                if ((!(object instanceof Date)) && (!(object instanceof Double)))
                    unit.setUnitNote((String) object);
                else
                    unit.setUnitNote("");
            }
        }
        return unit;
    }


    public static void writeToWorkBook(List<Unit> unitList, String workBookName) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("s1");
        for (int i = 0; i < unitList.size(); i++) {
            Row row = sheet.createRow(i);
            writeRow(wb, row, unitList.get(i));
        }
        saveWb(wb, workBookName);
    }

    public static void saveWb(Workbook wb, String workBookName) {
        try (FileOutputStream fout = new FileOutputStream("G:\\excelFiles\\" + workBookName + ".xlsx")) {
            wb.write(fout);
            fout.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void writeRow(Workbook wb, Row row, Unit unit) {

        CreationHelper creationHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        LocalDate localDate = unit.getUnitReleaseDate();

        row.createCell(0).setCellValue(unit.getUnitType());
        row.createCell(1).setCellValue(unit.getUnitNumber());

        Cell cellThree = row.createCell(2);
        cellThree.setCellValue(Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant()));
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/mm/yyyy"));
        cellThree.setCellStyle(cellStyle);

        row.createCell(3).setCellValue(unit.getUnitNote());
    }
}


