package edu.handong.java.JavaFinalProject.utils;

import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.ArrayList;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;

public class ZipReader extends Thread {
	
	String filePath;
	String fileName;
	private ArrayList<ExcelReader> eachEx = new ArrayList<ExcelReader>() ; 
	private ArrayList<ExcelReader> eachEx2 = new ArrayList<ExcelReader>() ;

	
	public ZipReader(String path, String name) {
		filePath = path;
		fileName = name;
		readFileInZip(filePath);
		
	}
	

	public String getFileName() {
		return fileName;
	}

	public void readFileInZip(String path) {
		
		ZipFile zipFile;
		try {
			zipFile = new ZipFile(path, "EUC-KR");
			Enumeration<? extends ZipArchiveEntry> entries = zipFile.getEntries();

		    while(entries.hasMoreElements()){
		    	ZipArchiveEntry entry = entries.nextElement();
		    	
		    	if(entry.getName().substring(entry.getName().lastIndexOf(".")+1).equals("xlsx")) { //if not, error.
		    		
		    		 InputStream stream = zipFile.getInputStream(entry);
		    		 
		    		
		    		//System.out.println(entry.getName());
		    		
		    		if(entry.getName().contains("요약문")) {
		    			ExcelReader myReader = new ExcelReader(stream);
		    			eachEx.add(myReader); //why fileType should be first?
		    		}
		    		else if(entry.getName().contains("표")) {
		    			ExcelReader myReader = new ExcelReader(stream);
		    			eachEx2.add(myReader); //why fileType should be first?
		    		}
		    		
		       
		    		
		    
		       
		    }
		    	
		    }
		    if(eachEx.isEmpty()) System.out.println("is Empty");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public ArrayList<ExcelReader> getExcelReader(){
		return eachEx;
	}
	public ArrayList<ExcelReader> getExcelReader2(){
		return eachEx2;
	}
}