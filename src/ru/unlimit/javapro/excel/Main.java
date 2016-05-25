package ru.unlimit.javapro.excel; 
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

public class Main {
    public static void main(String[] args) throws IOException {
    	//int sheetnum = 0;
        //doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Nagruzka_2013_denna.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/test.xls", 4, sheetnum++);
    	doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data1.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/navch_robota.xls", 3, "Навчальна робота");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота (практика).xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data2.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/navch_robota_pr.xls", 3, "Навчальна робота (практика)");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота (інші види).xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data3.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/navch_robota_iv.xls", 3, "Навчальна робота (інші види)");
        //doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data1.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/Індивідуальний план.xls", 3, sheetnum++);
        //doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота (практика).xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data2.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/Індивідуальний план.xls", 3, sheetnum++);
        //doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчальна робота (інші види).xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data3.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/Індивідуальний план.xls", 3, sheetnum++);
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Навчально-виховна і навчально-методична робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data4.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/nv_i_nm_robota.xls", 3, "Навчально-виховна і навчально-методична робота");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Методична робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data5.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/metod_robota.xls", 1, "Методична робота");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Наукова робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data6.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/nauk_robota.xls", 1, "Наукова робота");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Організаційна робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data7.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/org_robota.xls", 1, "Організаційна робота");
        doAllWork("C:/Users/Наталия/workspace/javaexel/testfiles/Виховна робота.xls", "C:/Users/Наталия/workspace/javaexel/testfiles/data8.txt", "C:/Users/Наталия/workspace/javaexel/testfiles/vih_robota.xls", 1, "Виховна робота");
          //readFromFile("C:/Users/Наталия/workspace/javaexel/testfiles/data1.txt");
    }
    
    
    public static void doAllWork(String inputExcelName, String dataFileName, String outputExcelName, int rownum, String Name) throws IOException{
    	try {
    	Workbook workbook = readExcelAndGetWorkbook(inputExcelName, rownum);
    		// FileInputStream file = new FileInputStream(new File(inputExcelName));
    	ArrayList<ArrayList<Object>> list = readFromFile(dataFileName);
    	Sheet sheet = workbook.getSheetAt(0);
    	
    	for (ArrayList<Object> secondList: list) {
    		 Row row = sheet.createRow(rownum++);
    		 int cellnum = 0;
    		 for (Object object: secondList) {
    			 Cell cell = row.createCell(cellnum++);
    			 cell.setCellValue(object.toString());
    		 }
    	}
    	workbook.setSheetName(workbook.getSheetIndex(sheet), Name);
    	writeWorkbook(workbook, outputExcelName);
    	} catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public static ArrayList<ArrayList<Object>> readFromFile(String dataFileName) throws IOException
    {
    	ArrayList<ArrayList<Object>>  list  = new ArrayList<ArrayList<Object>> ();
    	BufferedReader br = new BufferedReader(new FileReader(dataFileName));
        String strLine;
        ArrayList<Object> row=new  ArrayList<Object>();
        while ((strLine = br.readLine()) != null){
        	row = new ArrayList<Object>();
            for (String cellString: strLine.split(", ")) {
            	try {
                    int val =Integer.parseInt(cellString);
                    row.add(val);
                	} catch(NumberFormatException e) {
                		row.add(cellString);
                }
            }
          list.add(row);  
          System.out.println(row);
        }
        br.close();
 	return list;
    	
    }

    /* public static void doAllWork(String inputExcelName, String dataFileName, String outputExcelName, int rownum) {
        FileInputStream fstream = null;
        try {
            fstream = new FileInputStream(dataFileName);
            Workbook workbook = readExcelAndGetWorkbook(inputExcelName, rownum);
            Sheet sheet = workbook.getSheetAt(0);
          BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
            String strLine;
            //int rownum = 3;
            while ((strLine = br.readLine()) != null){
                Row row = sheet.createRow(rownum++);
                int cellnum = 0;
                for (String cellString: strLine.split(", ")) {
                    Cell cell = row.createCell(cellnum++);
                    try {
                        cell.setCellValue(Integer.parseInt(cellString));
                    } catch(NumberFormatException e) {
                        cell.setCellValue(cellString);
                    }
                }
            }
            br.close();
            writeWorkbook(workbook, outputExcelName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/

    public static Workbook readExcelAndGetWorkbook(String fileName, int rownum) throws IOException {
        FileInputStream file = new FileInputStream(new File(fileName));
        Workbook workbook = new HSSFWorkbook(file);
        //remove first two sheets
        //workbook.removeSheetAt(1);
        //workbook.removeSheetAt(0);
        Sheet sheet = workbook.getSheetAt(0);
       while (sheet.getPhysicalNumberOfRows() > rownum) {
          sheet.removeRow(sheet.getRow(sheet.getLastRowNum()));
       }

        return workbook;
    }

    public static void writeWorkbook(Workbook workbook, String fileName) throws IOException {
        FileOutputStream out = new FileOutputStream(new File(fileName));
        workbook.write(out);
        out.close();
    }
	
   
}
