/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generararchivoxmlmegacable;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author jorgepercastegui
 */
public class GenerarArchivoXMLMegaCable {

    /**
     * @param args the command line arguments
     */
    public void readExcelFile(File excelFile){
    InputStream excelStream = null;
        try {
            excelStream = new FileInputStream(excelFile);
            //Rpresentación del más alto nivel de la hoja de excel.
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(excelStream);
            //Elegimos la hoja que se pasa por parámetro.
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            //Objeto que nos permite leer una fila de la hoja, y de de aquí extraer el contenido de las celdas.
            HSSFRow hssfRow;
            //Inicializo el objeto que leerá el valor de la celda
            HSSFCell cell;
            //Obtengo el numero de filas ocupadas en la hoja
            int rows = hssfSheet.getLastRowNum();
            //Obtengo el numero de columnas ocupadas en la hoja
            int cols = 0;
            //Cadena que usamos para almacenar la lectura de la celda
            String cellValue;
            //Para este ejemplo vamos a recorrer las filas obteniendo los datos que queremos
            for(int r = 0; r< rows; r++){
            hssfRow = hssfSheet.getRow(r);
                if(hssfRow == null){
                    break;
                }else{
                    System.out.print("Row: " + r + " -> ");
                    for(int c =0; c< (cols = hssfRow.getLastCellNum()); c++){
                    
                            cellValue = hssfRow.getCell(c) == null?"":
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING)?hssfRow.getCell(c).getStringCellValue():
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC)?"" + hssfRow.getCell(c).getNumericCellValue():
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN)?"" + hssfRow.getCell(c).getBooleanCellValue():
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK)?"BLANK":
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA)?"FORMULA":
                                    (hssfRow.getCell(c).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR)?"ERROR":"";
                                    System.out.print("[Column " + c + ": "+cellValue + "] ");
                    }
                    System.out.println();
                }
            }
            
         
            
        } catch (FileNotFoundException fileNotFoundException) {
             System.out.println("The file not exists (No se encontró el fichero): " + fileNotFoundException);
        } catch (IOException ex) {
            System.out.println("Error in file procesing (Error al procesar el fichero): " + ex);
        }finally {
            try {
                excelStream.close();
            } catch (IOException ex) {
                System.out.println("Error in file processing after close it (Error al procesar el fichero después de cerrarlo): " + ex);
            }
        }
    
    }
    
    public static void main(String[] args) {
        GenerarArchivoXMLMegaCable megacablexml = new GenerarArchivoXMLMegaCable();
        megacablexml.readExcelFile(new File("PaisesIdiomasMonedas.xls"));
        
    }
    
}
