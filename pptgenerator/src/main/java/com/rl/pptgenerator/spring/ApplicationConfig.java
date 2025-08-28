package com.rl.pptgenerator.spring;

import java.lang.reflect.Constructor;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.context.annotation.PropertySources;
import org.springframework.context.support.PropertySourcesPlaceholderConfigurer;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

@Configuration

@PropertySources({
    @PropertySource("classpath:application.properties")
})
@ComponentScan(basePackages = {"com.rl.pptgenerator, com.rl.pptgenerator.util, com.rl.pptgenerator.spring"})
public class ApplicationConfig {

}
