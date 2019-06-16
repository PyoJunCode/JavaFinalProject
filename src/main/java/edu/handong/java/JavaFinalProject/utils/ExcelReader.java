package edu.handong.java.JavaFinalProject.utils;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {
	
	ArrayList<String> datas = new ArrayList<String>();
	
	public ArrayList<String> getDatas(){
		return datas;
	}

	
	public ExcelReader(InputStream is) {
		ArrayList<String> values = new ArrayList<String>();
		
		try (InputStream inp = is) {
		    //InputStream inp = new FileInputStream("workbook.xlsx");
		    
		        Workbook wb = WorkbookFactory.create(inp);
		        Sheet sheet = wb.getSheetAt(0);
		        Iterator<Row> iterator = sheet.iterator();
		        iterator.next();
		        while(iterator.hasNext()) { //row 
		        	Row currentRow = iterator.next();
		        	Iterator<Cell> cellIt= currentRow.cellIterator();
		        	String concatData = "";
		        	
		        	while(cellIt.hasNext()) { // cell
		        		Cell currentCell = cellIt.next();
		        		if(currentCell.getCellType() == CellType.STRING) {
		        			//System.out.println(currentCell.getStringCellValue());
		        			concatData += currentCell.getStringCellValue() + "#";
		        		}
		        		else if(currentCell.getCellType() == CellType.NUMERIC) {
		        			//System.out.println(currentCell.getNumericCellValue());
		        			concatData += currentCell.getNumericCellValue() + "#";
		        		}
		        		else if(currentCell.getCellType() == CellType.BLANK && cellIt.hasNext()) {
		        			
		        			concatData += "" + "#";
		        		}
		        	}
		        	//System.out.println("data : " + concatData);
		        	datas.add(concatData);
		        }
		        
		    } catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
		
	}
}