package com.rl.pptgenerator.util;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

@Component
public class ExcelDataReader {

	private static final Logger logger = LogManager.getLogger(ExcelDataReader.class);
	
	@Value("${report.column.0}")
	private String reportColumn0;

	@Value("${report.column.1}")
	private String reportColumn1;

	@Value("${report.column.2}")
	private String reportColumn2;

	@Value("${report.column.3}")
	private String reportColumn3;

	@Value("${report.column.4}")
	private String reportColumn4;
	
	@Value("${report.column.5}")
	private String reportColumn5;

	@Value("${report.column.6}")
	private String reportColumn6;

	@Value("${report.column.7}")
	private String reportColumn7;

	@Value("${report.column.8}")
	private String reportColumn8;

	@Value("${report.column.9}")
	private String reportColumn9;
	
	@Value("${report.column.10}")
	private String reportColumn10;

	@Value("${report.column.11}")
	private String reportColumn11;

	@Value("${report.column.12}")
	private String reportColumn12;

	@Value("${report.column.13}")
	private String reportColumn13;

	@Value("${report.column.14}")
	private String reportColumn14;	
	
	public List readExcelData(String location, String sourceName) 
	{
		List results = null;

		try
		{
			File file = new File(location+sourceName);
			Workbook workbook = new XSSFWorkbook(file);
			XSSFSheet projectsheet = (org.apache.poi.xssf.usermodel.XSSFSheet)workbook.getSheetAt(0);
			Iterator<Row> rowIterator = projectsheet.iterator();
			HashMap project = null;
			String cellValue = null;
			int rowCount=0; int colCount=0;
			
			results = new ArrayList();
			
			while (rowIterator.hasNext()) {
				Row row = null;
				/**
				 * We are skipping the first row in the excel which is the header.
				 * Note: After extracting the excel from PlanView Portfolios, please review the rows 2 & 3 after the header
				 * which do not contain project information
				 */
				while(rowCount<1)
				{
					rowIterator.next();
					rowCount++;
				}
				/**
				 * We start reading the data from the 2nd row onwards. 
				 * That is, when the counter is at 1 onwards.
				 */
				if(rowCount>=1) 
				{
					row = rowIterator.next(); 
					/**
					 * Since we have already incremented the rowCount in the preceding while loop upto 1,
					 * we will NOT increment it when it is equal to 1, but only restart incrementing it
					 * when it is 2 or more.
					 */
					if(rowCount>=2)
					{
						rowCount++;
					}
					colCount=0;
					 
					project = new HashMap();
					
					Iterator<Cell> cellIterator = row.cellIterator();
	                while (cellIterator.hasNext()) {
	                	Cell cell = cellIterator.next();
	                	cellValue = cell.toString();
                		cellValue=cellValue.trim();
	                	switch(colCount) {
	                		case 0:
	                			project.put(reportColumn0, cellValue);
	                			break;
	                		case 1:
	                			project.put(reportColumn1, cellValue);
	                			break;
	                		case 2:
	                			project.put(reportColumn2, cellValue);
	                			break;
	                		case 3:
	                			project.put(reportColumn3, cellValue);
	                			break;
	                		case 4:
	                			project.put(reportColumn4, cellValue);
	                			break;
	                		case 5:
	                			project.put(reportColumn5, cellValue);
	                			break;
	                		case 6:
	                			project.put(reportColumn6, cellValue);
	                			break;
	                		case 7:
	                			project.put(reportColumn7, cellValue);
	                			break;
	                		case 8:
	                			project.put(reportColumn8, cellValue);
	                			break;
	                		case 9:
	                			project.put(reportColumn9, cellValue);
	                			break;
	                		case 10:
	                			project.put(reportColumn10, cellValue);
	                			break;
	                		case 11:
	                			project.put(reportColumn11, cellValue);
	                			break;
	                		case 12:
	                			project.put(reportColumn12, cellValue);
	                			break;
	                		case 13:
	                			project.put(reportColumn13, cellValue);
	                			break;
	                	}
	                    colCount++;
	                }
	                
	                results.add(project);
	                String tempProjectName = null;
	                for(int i=0; i<results.size();i++)
	                {
	                	HashMap projectRow = (HashMap)results.get(i);
	                	tempProjectName = (String)projectRow.get(reportColumn0);
	                	if(tempProjectName==null || tempProjectName.trim().equals("Yes"))
	                	{
	                		results.remove(i);
	                	}
	                }
				}
			}
			
			for(int row=0;row<results.size(); row++)
			{
				HashMap projectRow = (HashMap)results.get(row);
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
		return results;
	}


}
