package com.hm.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReading {

    public static void writeSheetAsCsv(Sheet sheet, String csvfilePath, String outputFileName) throws IOException {
        Row row = null;
        StringBuilder csvSheet = new StringBuilder();
        FileWriter csvfile;
        // OpenCSV writer object to create CSV file
        if(!csvfilePath.endsWith("/")) {
        	csvfilePath += "/";
        }
        csvfile = new FileWriter(csvfilePath + outputFileName + 
        		"_" + new SimpleDateFormat("ddMMyyyy").format(new Date()) + ".csv");
        int batchSize = 100;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            if(isEmptyRow(row)) {
            	continue;
            }
            StringBuilder csvRow = new StringBuilder();
            if (row != null && row.getLastCellNum() > 0) {
	            for (int j = 0; j < row.getLastCellNum(); j++) {
	            	Cell cell = row.getCell(j);
	            	if(null == cell) {
	            		csvRow.append("");
	            	} else {
	            		String cellValue = "";
	            		if(cell.toString().contains(",")) {
	            			cellValue = "\"" + cell + "\"";
	            		} else {
	            			cellValue = "" + getCellValue(cell);
	            		}
	            		csvRow.append(cellValue);
	            	}
	                if (j+1 < row.getLastCellNum()) {
	                	csvRow.append(",");
	                } else {
	                	csvRow.append(System.lineSeparator());
	                }
	            }
	            csvSheet.append(csvRow);
            }
            
            if(i % batchSize == 0) {
            	csvfile.append(csvSheet.toString());
            	csvSheet.setLength(0);
            } 
        }
        csvfile.append(csvSheet.toString());
		csvfile.close();
    }
    
    private static boolean isEmptyRow(Row row) {
        if (row == null || row.getLastCellNum() <= 0) {
            return true;
        }
        for(int j = 0; j < row.getLastCellNum(); j++) {
	        Cell cell = row.getCell(j);
	        if (cell != null && !"".equals((cell+"").trim())) {
	            return false;
	        }
        }
        return true;
    }
    
    public static Object getCellValue(Cell cell) {
    	switch (cell.getCellType()) {
		case NUMERIC:
			Double d = cell.getNumericCellValue();
			if (d % 1 == 0) {
				return d.longValue();
			}
			return d;
		default:
			return cell;
		}
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
    	System.out.println("excel2csv tool ...");
    	System.out.println("Usage: excel2csv sample.xls outputpath [commaSeparated_sheetsNamesToConvert [commaSeparated_outputFilesNames]]");
    	System.out.println();
//    	System.out.println("ex1: excel2csv sample.xlsx /home");
//    	System.out.println("ex2: excel2csv sample.xls C:\\home sheet1,sheet2,sheet3");
//		System.out.println("ex3: excel2csv sample.xlsx /home sheet1,sheet2,sheet3 AA,BB,CC");
    	if(args.length < 2 || args.length > 4) {
    		System.out.println("wrong number of arguments. Required two arguments: " + 
    				System.lineSeparator() + "1.excel_file_to_covert " + System.lineSeparator() + "2.generated_files_path"
    				 + System.lineSeparator() + "3.OPTIONAL commaSeparated_sheetsNamesToConvert"
    				 + System.lineSeparator() + "4.OPTIONAL commaSeparated_outputFilesNames");
    		System.exit(0);
    	}
    	String outputPath = args[1];
        InputStream inp = null;
        try {
            
        	inp = new FileInputStream(args[0]);
            Workbook wb = WorkbookFactory.create(inp);

            List<String> sheetNames = null;
            if(args.length > 2 && null != args[2]) {
            	sheetNames = Arrays.asList(args[2].split(","));
            }
            List<String> outputFilesNames = null;
            if(args.length > 3 && null != args[3]) {
            	outputFilesNames = Arrays.asList(args[3].split(","));
            }
            
            for(int i = 0; i < wb.getNumberOfSheets(); i++) {
            	Sheet sheet = wb.getSheetAt(i);
            	if(sheetNames == null) {
            		writeSheetAsCsv(sheet, outputPath, sheet.getSheetName());
            	} else {
            		if(sheetNames.contains(sheet.getSheetName())) {
            			if(null == outputFilesNames) {
            				writeSheetAsCsv(sheet, outputPath, sheet.getSheetName());
            			} else {
            				writeSheetAsCsv(wb.getSheetAt(i), outputPath, outputFilesNames.get(sheetNames.indexOf(sheet.getSheetName())));
            			}
            		}
            	}
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("XXException " + ex.getMessage());
            return;
        } catch (IOException ex) {
            Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("XXException " + ex.getMessage());
            return;
        } finally {
            try {
                inp.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
                System.out.println("XXException " + ex.getMessage());
            }
        }
        
        System.out.println();
        System.out.println("Task Completed");
    }
}
