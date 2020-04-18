package com.stepdefinition;

import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.objectrepository.AddCustomerPage;
import com.objectrepository.HomePage;
import com.resources.FunctionalLibrary;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.cucumber.datatable.DataTable;

public class AddCustomerSteps  {
	
	FunctionalLibrary f= FunctionalLibrary.getInstance();
	
	HomePage hp=HomePage.getHomeInstance();
	AddCustomerPage acp=AddCustomerPage.getCustomerInstance();

	@Given("User click on add customer button")
	public void user_click_on_add_customer_button() throws Throwable {
		try {
		handleFrame();
		}catch(Exception e) {
		}
	   f.button(hp.getAddCustomerBtn());
	   
	}

	@When("User enters all the fields")
	public void user_enters_all_the_fields() throws Throwable {
		try {
			handleFrame();
		}catch(Exception e) {
		}
		 f.button(acp.getDoneBtn());
		    f.insertText(acp.getFname(), "thara");
		    f.insertText(acp.getLname(), "gold");
		    f.insertText(acp.getEmail(), "thara123@gmail.com");
		    f.insertText(acp.getAddress(), "pudukkottai");
		    f.insertText(acp.getTelephno(), "65286637231");
	 }
	@When("User enters all the fields with one dimensional")
	public void user_enters_all_the_fields_with_one_dimensional(DataTable customerData) throws Throwable {
		List<String> data =customerData.asList(String.class);
		
		try {
			handleFrame();
		}catch(Exception e) {
		}
		 f.button(acp.getDoneBtn());
		   f.insertText(acp.getFname(), data.get(0));
		    f.insertText(acp.getLname(), data.get(1));
		    f.insertText(acp.getEmail(), data.get(2));
		    f.insertText(acp.getAddress(), data.get(3));
		    f.insertText(acp.getTelephno(), data.get(4));	
		
		
/*	   driver.findElement(By.xpath("(//label[@for='done'])[1]")).click();
	   driver.findElement(By.id("fname")).sendKeys(data.get(0));
	   driver.findElement(By.id("lname")).sendKeys(data.get(1));
	   driver.findElement(By.id("email")).sendKeys(data.get(2));
	   driver.findElement(By.name("addr")).sendKeys(data.get(3));
	   driver.findElement(By.id("telephoneno")).sendKeys(data.get(4));*/
	 }
	@When("User enters all the fields with one dimensional map")
	public void user_enters_all_the_fields_with_one_dimensional_map(DataTable customerData) throws Throwable {
		 Map<String,String> data =customerData.asMap(String.class,String.class);
		 try {
				handleFrame();
			}catch(Exception e) {
			}
		  FunctionalLibrary.driver.findElement(By.xpath("(//label[@for='done'])[1]")).click();
		  FunctionalLibrary.driver.findElement(By.id("fname")).sendKeys(data.get("Fname"));
		  FunctionalLibrary. driver.findElement(By.id("lname")).sendKeys(data.get("Lname"));
		  FunctionalLibrary.driver.findElement(By.id("email")).sendKeys(data.get("Email"));
		  FunctionalLibrary.driver.findElement(By.name("addr")).sendKeys(data.get("Address"));
		  FunctionalLibrary. driver.findElement(By.id("telephoneno")).sendKeys(data.get("Phno"));
		
	}
	


	@When("User click on submit button")
	public void user_click_on_submit_button() throws Exception {
		
		
    f.button(acp.getSubmitBtn());		  
	}

	@Then("User should be displayed cutomer ID is generated or no")
	public void user_should_be_displayed_cutomer_ID_is_generated_or_no() throws Throwable {
	    
	
		try {
			handleFrame();
		}catch(Exception e) {
		}
	Assert.assertTrue(FunctionalLibrary.driver.findElement(By.xpath("(//td[@align='center'])[2]")).isDisplayed());
	String text = FunctionalLibrary.driver.findElement(By.xpath("(//td[@align='center'])[2]")).getText();
	   System.out.println(text);
	}
	
	public void handleFrame() throws Throwable {
	 Thread.sleep(5000);
	 FunctionalLibrary.driver.switchTo().frame(FunctionalLibrary.driver.findElement(By.xpath("//*[@id=\"flow_close_btn_iframe\"]")));
	 FunctionalLibrary.driver.findElement(By.id("closeBtn")).click();
	 FunctionalLibrary.driver.switchTo().parentFrame();
	}
	

}
