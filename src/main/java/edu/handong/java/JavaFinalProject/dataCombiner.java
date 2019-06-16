package edu.handong.java.JavaFinalProject;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import edu.handong.java.JavaFinalProject.utils.ExcelReader;
import edu.handong.java.JavaFinalProject.utils.ZipReader;
import edu.handong.java.JavaFinalProject.utils.customException;



public class dataCombiner {
	
	String dataPath;
	String resultPath;
	String[] finalName;
	
	ArrayList<ZipReader> inDirFiles = new ArrayList<ZipReader>(); 
	
	public void run(String[] args) {
		
		Options options = createOptions();
		
		parseOptions(options, args);
		
		
		finalName = resultPath.split("\\.");
		
		
		//extension = resultPath.split(".")[1];
		
		writer();

	}
	
	
	
	private Options createOptions() {
		Options options = new Options();

		// add options by using OptionBuilder
		options.addOption(Option.builder("i").longOpt("input")
				.desc("Set an input file path")
				.hasArg()
				.argName("Input path")
				.required()
				.build());

		// add options by using OptionBuilder
		options.addOption(Option.builder("o").longOpt("output")
				.desc("Set an output file path")
				.hasArg()
				.argName("Output path")
				.required()
				.build());



		return options;
	}
	
	private void printHelp(Options options) {
		// automatically generate the help statement
		HelpFormatter formatter = new HelpFormatter();
		String header = "HGU Data Combiner";
		String footer ="";
		formatter.printHelp("HGU Data Combiner", header, options, footer, true);
	}
	
	
	private void parseOptions(Options options, String[] args) {
		CommandLineParser parser = new DefaultParser();

		String fullPath = null;
		
		try {

			CommandLine cmd = parser.parse(options , args);
			

			dataPath = cmd.getOptionValue("i");
			resultPath = cmd.getOptionValue("o");
			
			

				
				File filetest = new File(dataPath);
				
				if(filetest.isDirectory()) {
					fullPath = filetest.getCanonicalPath(); // convert to absolute path for convenient.
				}
				else if(filetest.isFile()) {
					throw new customException("Please unzip mother zip first.");
				}
				else {
					throw new customException("Error : unexpected data. Please check error.csv");
				}
				
				
				
			}
			
			catch (customException e) {
				System.out.println(e.getMessage());
				System.exit(-1);
			}
			catch (IOException e) {
				System.out.println(e.getMessage());
				System.exit(-1);

			}
			catch (Exception e) {
			printHelp(options);
			System.exit(-1);
			}
			
			
		
			File abFile = new File(fullPath);
			
			File[] inDir = abFile.listFiles();
			
			for(File files : inDir) {
				
				if(files.getName().substring(files.getName().lastIndexOf(".")+1).equals("zip")) { //if not , error occurred.
					
				try {
					inDirFiles.add(new ZipReader(files.getCanonicalPath(), files.getName()));
					
				}
				catch (IOException e) {
					e.printStackTrace();
				}
				}
			}
			//https://manorgass.tistory.com/60 
			
			 inDirFiles.sort(new Comparator<ZipReader>() {
	              @Override
	              public int compare(ZipReader path1, ZipReader path2) {
	            	  String one = path1.getFileName();
	            	  String two = path2.getFileName();
	                     // TODO Auto-generated method stub
	            	  
	            	  return one.compareTo(two);
	                    
	              }
	       });

			
			
			ArrayList<Thread> threads = new ArrayList<Thread>();
			
			for(ZipReader adder : inDirFiles) {
				Thread thread = new Thread(adder);
				thread.start();
				threads.add(thread);
			}
			try {
				for(Thread adder: threads) {
					adder.join(); //waiting
				}
			}catch (InterruptedException e) {
				System.out.println(e.getMessage());
				System.exit(-1);
			}
			

		
		}
	
	public void writer() {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("combined");
		int rowNum = 0;
		
		
		File wFile = new File("./"+finalName[0] + "1." + finalName[1]);
		File wFile2 = new File("./"+finalName[0] + "2." + finalName[1]);
		wFile.getParentFile().mkdirs();
		wFile2.getParentFile().mkdirs();
		//else if (!wFile.exists()) wFile.createNewFile();
		
		
		try {

			
			
			BufferedWriter bw = new BufferedWriter(new FileWriter(wFile));
			wFile.createNewFile();
			BufferedWriter bw2 = new BufferedWriter(new FileWriter(wFile2));
			wFile2.createNewFile();
			
			CSVPrinter csvPrinter = new CSVPrinter(bw, CSVFormat.DEFAULT.withHeader("Num","제목","요약문 (300자 내외)", "핵심어\n" + 
					"(keyword,쉽표로 구분)", "조회날짜", "실제자료조회\n" + 
							"출처 (웹자료링크)", "원출처 (기관명 등)", "제작자\n" + 
									"(Copyright 소유처)" ));
			CSVPrinter csvPrinter2 = new CSVPrinter(bw2, CSVFormat.DEFAULT.withHeader("Num", "제목(반드시 요약문 양식에 입력한 제목과 같아야 함.)", "표/그림 일련번호", "자료유형(표,그림,…)", "자료에 나온 표나 그림 설명(캡션)", "자료가 나온 쪽번호"));
			int count = 0;
			for(ZipReader zip : inDirFiles) {
				for(ExcelReader list : zip.getExcelReader()) {
					for(String str : list.getDatas() ) {
						if(!str.equals("")) {
							//System.out.println("Title : " + count++ + str);
							String[] splited = str.split("#");
							
							if(splited.length == 7)
							csvPrinter.printRecord(zip.getFileName().substring(1,4), splited[0],splited[1],splited[2],splited[3],splited[4],splited[5],splited[6]);
						
						}
					}
				}
				
				for(ExcelReader list : zip.getExcelReader2()) {
					for(String str : list.getDatas() ) {
						//System.out.println(str);
						if(!str.equals("")) {
							//System.out.println("Title : " + count++ + str);
							String[] splited = str.split("#");
							
							if(splited.length == 5)
							csvPrinter2.printRecord(zip.getFileName().substring(1,4), splited[0],splited[1],splited[2],splited[3],splited[4]);
						
						}
					}
				}
				
				
			}
			
			System.out.println("Write Done");
			
			csvPrinter.flush();
			csvPrinter.close();
			csvPrinter2.flush();
			csvPrinter2.close();

		}
		
		
		 catch (Exception e) {
			 System.out.println(e.getMessage());
			 System.exit(0);
			 
		 }
		
		
	}
	
	
}
