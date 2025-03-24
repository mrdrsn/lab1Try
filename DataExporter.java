package com.mycompany.exceltry;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class DataExporter {
    
    public static String[] statisticNeeds = {"Среднее геометрическое", "Среднее арифметическое", "Оценка стандратного отклонения",
                                            "Размах",
                                            "Количество элементов","Коэффициент вариации", "Доверительный этап мат. ожидания",
                                            "Оценка дисперсии",
                                            "Максимум", "Минимум"};
//    public String[] sampleName;
    public static DataStorage storage = new DataStorage();
    
    public static Workbook createEmptyExportBook(String source, int listIndex, String dest) throws FileNotFoundException, IOException{
        
        DataExporter.storage = DataImporter.importFromExcel(source, listIndex);
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Расчеты для выборок");
        
        Row row = sheet.createRow(0);
        Cell emptyCell = row.createCell(0);
        
        //создаю шапку таблицы (обозначения выборок)
        for(int i = 0; i < storage.getSampleNames().length; i++){
            Cell sampleName = row.createCell(i+1);
            sampleName.setCellValue(storage.getSampleNames()[i]);
        }
        
        //создаю столбец с обозначением стат показателя
        for(int i = 0; i < statisticNeeds.length; i++){
            Row r = sheet.createRow(i+1);
            Cell statName = r.createCell(0);
            statName.setCellValue(statisticNeeds[i]);
        }
        
        //создаю матрицу ковариации(?)
        Row covRow = sheet.getRow(0);
        Cell covTitle = covRow.createCell(2 + storage.getSampleNames().length); //5 должно вычисляться по формуле
        covTitle.setCellValue("Матрица ковариации");
        
        for (int i = 0; i < storage.getSampleNames().length; i++) {
            Row rowI = sheet.getRow(0); // Используем существующие строки
            Cell sampleName = rowI.createCell(i + 1 + covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + storage.getSampleNames()[i]);
        }
        
        for (int j = 0; j < storage.getSampleNames().length; j++) {
            Row rowJ = sheet.getRow(j+1); // Используем существующие строки
            Cell sampleName = rowJ.createCell(covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + storage.getSampleNames()[j]);
        }
        // Автоматическое изменение размера столбцов
        for (int i = 0; i <= 2+2*storage.getSampleNames().length; i++) {
            sheet.autoSizeColumn(i);
            if(i >= 1 && i <= 5){
                sheet.setColumnWidth(i, 256*10);
            }
        }
        wb.write(new FileOutputStream(dest));
        return wb;
    }
    
    public static void createCalculationBook(String source, int listIndex, String file) throws IOException {
        Workbook wb = createEmptyExportBook(source, listIndex, file);
        Sheet sheet = wb.getSheet("Расчеты для выборок");
        Row headerRow = sheet.getRow(0); // Строка с названиями выборок

        Map<String, Double> geometricMeanMap = DataCalculations.calculateGeometricMean(storage);
        Map<String, Double> arithmeticMeanMap = DataCalculations.calculateArithmeticMean(storage);
        Map<String, Double> standardDeviationMap = DataCalculations.calculateStandardDeviation(storage);
        Map<String, Double> sampleRangeMap = DataCalculations.calculateSampleRange(storage);
        Map<String, Double> sizeMap = DataCalculations.calculateSize(storage);
        Map<String, Double> minMap = DataCalculations.calculateMin(storage);
        Map<String, Double> maxMap = DataCalculations.calculateMax(storage);
        Map<String, Double> varianceMap = DataCalculations.calculateVariance(storage);
        Map<String, Double> varCoeffMap = DataCalculations.calculateVarCoeff(storage);
        
        for (Map.Entry<String, Double> entry : geometricMeanMap.entrySet()) {
            String columnName = entry.getKey();
            double calculatedValue = entry.getValue();

            // Поиск столбца с соответствующим названием
            Iterator<Cell> cellIter = headerRow.cellIterator();
            while (cellIter.hasNext()) {
                Cell currentCell = cellIter.next();
                if (currentCell.getCellType() == CellType.STRING && 
                    currentCell.getStringCellValue().equals(columnName)) {
                    int columnIndex = currentCell.getColumnIndex();

                    // Обновление значения в строке данных
                    Row dataRow = sheet.getRow(1); // Предполагается, что данные находятся во второй строке
                    Cell valueCell = dataRow.getCell(columnIndex);
                    if (valueCell == null) {
                        valueCell = dataRow.createCell(columnIndex);
                    }
                    valueCell.setCellValue(calculatedValue);
                    break; // Выходим из цикла после обработки ячейки
                }
            }
        }

        // Сохранение изменений в файл
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            wb.write(fileOut);
        } finally {
            wb.close();
        }
    }
}
