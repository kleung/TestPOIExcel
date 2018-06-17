package com.test.TestApachiPOI;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	OPCPackage pkg = null;
    	try {
    		pkg = OPCPackage.open("testExcelFile.xlsx");
    		XSSFWorkbook wb = new XSSFWorkbook(pkg);
    		
    		int sheetCount = wb.getNumberOfSheets();
    		System.out.println("Worksheet count: " + sheetCount);
    		
    		for(int workSheetCounter = 0; workSheetCounter < sheetCount; workSheetCounter++) {
    			System.out.println("Current sheet: " + workSheetCounter);
    			XSSFSheet sheet = wb.getSheetAt(workSheetCounter);
    			
    			int lastRowNum = sheet.getLastRowNum();
    			System.out.println("Current sheet row count: " + (lastRowNum + 1));
    			
    			for(int rowCounter = 0; rowCounter <= lastRowNum; rowCounter++) {
    				System.out.println("Current row: " + rowCounter);
    				XSSFRow row = sheet.getRow(rowCounter);
    				
    				short rowSize = row.getLastCellNum();
    				System.out.println("Current row column count: " + rowSize);
    				for(int columnCounter = 0; columnCounter < rowSize; columnCounter++) {
    					System.out.println("Current column: " + columnCounter);
    					XSSFCell cell = row.getCell(columnCounter);
    					
    					Object cellValue = null;
    					
    					CellType cellType = cell.getCellTypeEnum();
    					
    					switch(cellType) {
    						case _NONE : 
    						case BLANK : cellValue = null;break;
    						case BOOLEAN : cellValue = cell.getBooleanCellValue();break;
    						case FORMULA : {
    							String formula = cell.getCellFormula();
    							System.out.println("This cell contains a formula, which is: " + formula);
    								
    							FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
    							CellValue evaluatedValue = evaluator.evaluate(cell);
    							
    							switch (evaluatedValue.getCellTypeEnum()) {
    								case _NONE : 
    								case BLANK : cellValue = null;break;
    								case BOOLEAN : cellValue = evaluatedValue.getBooleanValue(); break;
    								case NUMERIC : cellValue = evaluatedValue.getNumberValue(); break;
    								case STRING : cellValue = evaluatedValue.getStringValue();break;
    							}
    						}; break;
    						case NUMERIC : cellValue = cell.getNumericCellValue();break;
    						case STRING : cellValue = cell.getStringCellValue();break;
    					}
    					
    					System.out.println("Current cell value: " + cellValue);
    				}
    			}
    			
    		}
    	} catch (Exception e) {
    		e.printStackTrace();
    	} finally {
    		if(pkg != null) {
    			try {
    				pkg.close();
    			} catch (Exception e) {
    				e.printStackTrace();
    			}
    		}
    	}
    	
    }
}
