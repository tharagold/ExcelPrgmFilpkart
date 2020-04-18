package com.stepdefinition;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.resources.FunctionalLibrary;

import cucumber.api.java.After;
import cucumber.api.java.Before;

public class Hook  {
	
	FunctionalLibrary f=FunctionalLibrary.getInstance();
	
@Before
		public void beforeScenario() {
		
        f.launch("http://demo.guru99.com/telecom/index.html");		
	     FunctionalLibrary.driver.manage().window().maximize();
	}
	
	@After
	public void afterScenario() {	
		
   FunctionalLibrary.driver.quit();
	}
	
}