package com.rl.pptgenerator;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;

@Component
public class ExecutionTrigger {

	private static final Logger logger = LogManager.getLogger(ExecutionTrigger.class);
	
	@Autowired
	ApplicationContext ctx;
	
	
}
