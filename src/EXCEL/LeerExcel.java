/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package EXCEL;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author luisg
 */
public class LeerExcel {
    
    public static String obtenerRuta(){
        String a = "PreguntasCMMIEtapas(1).xlsx";
        String ruta = System.getProperty("user.dir");
        String concat = ruta+"\\"+a;
        return concat;
    }
    

    public LeerExcel() {
        
    }
    
    public static void main(String[] args) {
//        System.out.println(obtenerRuta());
        leerPreguntasAseguramientoDeLaCalidad();
    }
    
    public static void crearExcel(){
        Workbook libro = new XSSFWorkbook();
        Sheet hoja = libro.createSheet("hoja 1");
        try {
            FileOutputStream archivo = new FileOutputStream("excel.xlsx");
            libro.write(archivo);
            archivo.close();
            
        } catch (Exception e) {
        }
    }
    
    public static String[] leerPreguntasGestionDeRequisitos(){
        String[] preguntas1 = new String[14];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas1.length; i++) {
                Row fila = hoja.getRow(i+16);
                Cell celda = fila.getCell(1);
                preguntas1[i] = celda.getStringCellValue();
            }
            return preguntas1;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasPlanificacionDeProyectos(){
        String[] preguntas = new String[25];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+33);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
                
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasAseguramientoDeLaCalidad(){
        String[] preguntas = new String[14];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+61);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasDesarrolloDeRequisitos(){
        String[] preguntas = new String[22];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+78);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasVerificacion(){
        String[] preguntas = new String[20];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+103);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasGestionDeRiesgo(){
        String[] preguntas = new String[19];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+126);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasProcesosOrganizativos(){
        String[] preguntas = new String[17];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+148);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
    public static String[] leerPreguntasGestionCuantitativaDeProyectos(){
        String[] preguntas = new String[19];
        try {
            FileInputStream archivo = new FileInputStream(new File(obtenerRuta()));
            XSSFWorkbook libro = new XSSFWorkbook(archivo);
            XSSFSheet hoja = libro.getSheetAt(0);
            for (int i = 0; i < preguntas.length; i++) {
                Row fila = hoja.getRow(i+168);
                Cell celda = fila.getCell(1);
                preguntas[i] = celda.getStringCellValue();
            }
            return preguntas;
        } catch (Exception e) {
            return null;
        }
    }
    
}
