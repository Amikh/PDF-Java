package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ConvertPDF {
	private static final Logger LOG = LoggerFactory.getLogger(ConvertPDF.class);
	
	public static String FILE_NAME = "C:\\Users\\internet\\Desktop\\excel-test\\excel.xlsx"; // the path excel file
	 static Workbook workbook = null;
	 static String isFileName = null;
	 public static List<String> filesList= new ArrayList<String>();
	 public static int count = 0;
	
	public static void main(String[] args) throws IOException {	
		
		isListFileForRename();  // first part - get all files from excel and put them to array list 
		isRenameFile();         //second  part - begin to run from array list and make text file from pdf  
	}
	
	   //FIRST PART
		public static void isListFileForRename() throws IOException {	
	        try {
	            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
	            workbook = new XSSFWorkbook(excelFile);
	            Sheet datatypeSheet = workbook.getSheetAt(0);
	            Iterator<Row> iterator = datatypeSheet.iterator();
	            while (iterator.hasNext()) {
	                Row currentRow = iterator.next();
	                Iterator<Cell> cellIterator = currentRow.iterator();
	                while (cellIterator.hasNext()) {
	                    Cell currentCell = cellIterator.next();
	                    if (currentCell.getCellType() == CellType.STRING) {
	                    	isFileName= currentCell.getStringCellValue();   // here if the String  
	                       LOG.info("The file name is : "+isFileName);
	                    } else if (currentCell.getCellType() == CellType.NUMERIC) { // here if the number 
	                    	 double tmp = currentCell.getNumericCellValue();
	                    	 isFileName = Double.toString(tmp);                 // convert number to string 
	                    	 LOG.info("The file name is : "+isFileName);
	                    }   
	                          filesList.add(isFileName); // Here adding our links to array list 
	                }       
	            }
	        } catch (FileNotFoundException e) {
	        	LOG.error(e.getMessage());
	            e.printStackTrace();
	        } catch (IOException e) {
	        	LOG.error(e.getMessage());
	            e.printStackTrace();
	        }finally {
				workbook.close();
			}
		}	    

	   
	//SECOND PART
	public static void isRenameFile()throws IOException {
		for(String str :filesList) {
			String FileString = str.toString();
			LOG.info("File for convert to text : "+ FileString);
			  isPDFParser(FileString);
			  count++; // the count for adding to files 
		}
		   	 
	}
	public static void isPDFParser(String fileStr) throws IOException {
		  //Loading an existing document
	      File file = new File(fileStr);
	      PDDocument document = PDDocument.load(file);
	      //Instantiate PDFTextStripper class
	      PDFTextStripper pdfStripper = new PDFTextStripper();
	      //Retrieving text from PDF document
	      String text = pdfStripper.getText(document);
	      	isCreateTextFiel(text);
	      //Closing the document
	      document.close();
	   }
	public static void isCreateTextFiel(String text) throws IOException {
		FileWriter writer = null;
		try {
			writer = new FileWriter("C:\\Users\\internet\\Desktop\\excel-test\\MyFile_"+count+".txt", true);
			writer.write(text);
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			writer.close();
		}
	}
	
}
