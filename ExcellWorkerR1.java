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

/**
 * Класс для работы с файлом комплектности
 *
 * @Author aaa4
 * @Author Gurzhy Alex
 * @Version 1.1
 */
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
        writeToWorkBook(unitList, "exp3",3);
        unitList.stream().forEach(x -> System.out.println(x));

    }

    /**
     * Метод, читающий 4 ячейки в каждой строке в зависимости от заданных unitParams
     *
     * @param sheet      - лист из книги Workbook
     * @param unitParams - экземпляр класса, где хранятся начальные и конечные значения для перечисления
     *                   строк (Row) и стробцов (cell)
     * @return список типа Unit, содержащие данные одного из разделов комплектности
     */
    public static List<Unit> readOneColumn(Sheet sheet, UnitParams unitParams) {
        List<Unit> unitList = new ArrayList<>();
        for (int rowIdx = unitParams.getStartAtRow(); rowIdx < unitParams.getStopAtRow(); rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            unitList.add(readUnitFromRow(row, unitParams));
        }
        return unitList;
    }

    /**
     * Метод, считывающий одну строчку(вернее 4 ячейки в ней) из листа книги эксель. Начало и конец отсчета
     * берет из params и возвращающей экземпляр класса Unit
     *
     * @param row    - строка в листе книги
     * @param params - экземпляр класса, хранящего начальное и конечное значение cell для перебора по ячейкам в строке
     * @return возвращает экземпляр Unit с заполненными полями из строки комплектности
     */
    private static Unit readUnitFromRow(Row row, UnitParams params) {
        Unit unit = new Unit();
        Object object = "";      //создать объект
        for (int i = params.getStartAtCell(); i < params.getStopAtCell(); i++) {    //перебор по четырем ячейкам
            Cell cell = row.getCell(i);        //получить ячейку из строки
            if (cell != null) {                //вообще может быть нуль, но лучше избавиться от него тут
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:               //число. Дата тоже NUMERIC чаще всего идет
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
                cell.setCellValue("");         //для нуля в ячейке задает ей тип String и значение = ""
            }

            int a = 0;          //вспомогательная переменная для читабельности перевода из дабла в стринг
            Double d = 0.0;     //то же самое
            LocalDate ld = LocalDate.of(1975, 2, 2); //эта дата записывается в те ячейки, которые б/п
            if ((i == 0) || (i == 4) || (i == 8)) {     //стартовые ячейки 1, 2, 3 колонок комплектности
                unit.setUnitType((String) object);      //UnitType
            }                                           //скобки тут лежат, т.к. в без них не всегда мне очевидно что где
            if ((i == 1) || (i == 5) || (i == 9)) {
                if (object instanceof String)
                    unit.setUnitNumber((String) object); //UnitNumber
                if (object instanceof Double) {          //если номер в комплектности идет через нумерик, а не текст
                    d = (Double) object;
                    a = d.intValue();
                    unit.setUnitNumber(Integer.toString(a));
                }
            }
            if ((i == 2) || (i == 6) || (i == 10)) {
                if (object instanceof Date) {
                    ld = ((Date) object).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();//каст из Date в LocalDate
                } else {

                }
                unit.setUnitReleaseDate(ld);  //UnitReleaseDate
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


    /**
     * Метод записывает в лист книги эксель (.xsls) список блоков
     *
     * @param unitList       - список классов типа Unit
     * @param workBookName   - имя книги
     * @param numberOfColumn - номер колонки
     */
    public static void writeToWorkBook(List<Unit> unitList, String workBookName, int numberOfColumn) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("s1");
        for (int i = 0; i < unitList.size(); i++) {
            Row row = sheet.createRow(i);
            writeRow(wb, row, unitList.get(i), numberOfColumn);
        }
        saveWb(wb, workBookName);
    }


    /**
     * Метод сохраняет книгу эксель
     *
     * @param wb           - книга эксель типа XSSFWorkbook
     * @param workBookName - имя книги в виде строки
     */
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


    /**
     * Метод записывает поля из параметра unit в ячейки строки Row, книги wb
     *
     * @param wb             - книга эксель типа XSSFWorkbook
     * @param row            - строка книги org.apache.poi.ss.usermodel.Row
     * @param unit           - сущность для хранения параметров блоков
     * @param numberOfColumn - 1, 2, 3 номер колонки
     */
    public static void writeRow(Workbook wb, Row row, Unit unit, int numberOfColumn) {

        CreationHelper creationHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        LocalDate localDate = unit.getUnitReleaseDate();
        if (numberOfColumn == 1)
            numberOfColumn = 0;
        else {
            if (numberOfColumn == 2)
                numberOfColumn =4;
            else
                numberOfColumn = 8;
        }
        row.createCell(0 + numberOfColumn).setCellValue(unit.getUnitType());
        row.createCell(1 + numberOfColumn).setCellValue(unit.getUnitNumber());

        Cell cellThree = row.createCell(2 + numberOfColumn);
        cellThree.setCellValue(Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant()));
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/mm/yyyy"));
        cellThree.setCellStyle(cellStyle);

        row.createCell(3 + numberOfColumn).setCellValue(unit.getUnitNote());


    }
}


