/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.exceltry;

//import static DataImporter.importFromExcel;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 *
 * @author nsoko
 */
public class ExcelTry {

    public static void main(String[] args) throws IOException {
        String file = "C:\\Users\\nsoko\\Downloads\\Лаба_1 образцы данных.xlsx";
        DataExporter.createCalculationBook(file, 1, "C:\\Users\\nsoko\\OneDrive\\Desktop\\testEmpty6.xlsx");
//        DataStorage ds = new DataStorage();
//        DataStorage ds = DataImporter.importFromExcel(file, 2, ds);
//        System.out.println(data.keySet());
//        System.out.println(data);
//        Map<String,Double> gm = DataCalculations.calculateGeometricMean(ds);
//        List<Double> am = DataCalculations.calculateArithmeticMean(ds);
//        System.out.println("Средние арифметические: " + am);
//        System.out.println("Средние геометрические: " + gm);
    }
}