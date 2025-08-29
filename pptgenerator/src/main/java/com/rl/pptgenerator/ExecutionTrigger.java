package com.rl.pptgenerator;
import java.util.HashMap;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.springframework.stereotype.Component;

import com.rl.pptgenerator.util.ExcelDataReader;

@Component
public class ExecutionTrigger {

	private static final Logger logger = LogManager.getLogger(ExecutionTrigger.class);
	
	@Autowired
	ApplicationContext ctx;
	
	@Autowired
	private ExcelDataReader excelReader;	
	
	@Value("${source.location}")
	private String sourceLocation;	

	@Value("${source.file.name}")
	private String sourceFileName;
	
	public ExecutionTrigger()
	{
		//empty constructor
	}
	
	public void startExecution()
	{
		List<HashMap> results = excelReader.readExcelData(sourceLocation, sourceFileName);		
	}
}
