package com.myexcel.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcellWorker {

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
        Sheet sheet = wb.getSheetAt(3);
        for (int i = 0; i < 25; i++) {
            Row row = sheet.getRow(i);
            unit = readUnitFromRow(row);
            System.out.println(unit);
            unitList.add(unit);
        }

        writeToWorkBook(unitList, "tempExcel");
    }


    public static Unit readUnitFromRow(Row row) {
        Unit unit = new Unit();
        Object object = "";
        for (int i = 0; i < 4; i++) {

            Cell cell = row.getCell(i);
            System.out.println("читаю ячейку " + i);

            if (cell != null) {
                System.out.print(cell + " ");
                System.out.println(cell.getCellTypeEnum());
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
            switch (i) {
                case 0:
                    unit.setUnitType((String) object);
                    break;
                case 1:
                    if (object instanceof String)
                        unit.setUnitNumber((String) object);
                    if (object instanceof Double) {
                        d = (Double) object;
                        a = d.intValue();
                        unit.setUnitNumber(Integer.toString(a));
                    }
                    break;
                case 2:
                    if (object instanceof Date) {
                        ld = ((Date) object).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

                    } else {

                    }
                    unit.setUnitReleaseDate(ld);
                    break;
                case 3:
                    if ((!(object instanceof Date)) && (!(object instanceof Double)))
                        unit.setUnitNote((String) object);
                    else
                        unit.setUnitNote("");
                    break;
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

        try (FileOutputStream fout = new FileOutputStream("G:\\excelFiles\\"+workBookName+".xlsx")) {

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


    public static void someMethod() {   //copy from main
        InputStream fis = null;
        try {
            fis = new FileInputStream("G:\\excelFiles\\xx7.xlsx");
            Workbook wb = new XSSFWorkbook(fis);
            Sheet sheet = wb.getSheetAt(3);
            System.out.println(" lastRowNum is " + sheet.getLastRowNum());

            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Unit unit = new Unit();
                if (row != null) {
                    for (int j = 0; j < 4; j++) {
                        Object cellObj = "";

                        Cell cell = row.getCell(j);
                        if (cell == null) {
                            System.out.print(" Cell type is blank ");
                        } else {
                            if (cell.toString().equals("б\\п") || (cell.toString().equals("б/п"))) {
                                cellObj = LocalDate.of(0, 1, 1);
                            } else {
                                System.out.print(" " + i + " ; " + j + " " + "lastCellNum is " + row.getLastCellNum() + " " + cell.getCellTypeEnum() + "=" + cell);
                                switch (cell.getCellTypeEnum()) {
                                    case NUMERIC:
                                        System.out.print(" " + cell.getNumericCellValue());
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                            cellObj = (cell.getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDate(); //localDate from date
                                            System.out.println(" " + cellObj + " is date formatted; ");
                                        }
                                        System.out.println();
                                        break;
                                    case STRING:
                                        cellObj = cell.getStringCellValue();
                                        System.out.println("s" + cellObj);
                                        break;
                                    case BLANK:
                                        System.out.println("blank cell type");
                                        break;
                                }
                                if (j == 0)
                                    unit.setUnitType(cellObj.toString());
                                if (j == 1)
                                    unit.setUnitNumber(cellObj.toString());
                                if (j == 2)
                                    unit.setUnitReleaseDate((LocalDate) cellObj);
                                if (j == 3)
                                    unit.setUnitNote(cellObj.toString());
                            }


                        }
                    }
                } else {
                    System.out.println("row = null");
                }
                System.out.println(unit);
                System.out.println();
            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}


