package com.rl.pptgenerator;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import com.rl.pptgenerator.spring.ApplicationConfig;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class Initiator {

	private static final Logger logger = LogManager.getLogger(Initiator.class);

	public static void main(String[] args)
	{
		ApplicationContext ctx = new AnnotationConfigApplicationContext(ApplicationConfig.class);
	}
}
