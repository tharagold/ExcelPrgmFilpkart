package com.stepdefinition;

import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.objectrepository.AddTariffPlanPage;
import com.objectrepository.HomePage;
import com.resources.FunctionalLibrary;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.cucumber.datatable.DataTable;

public class AddTariffPlanSteps {
	
	 FunctionalLibrary f=FunctionalLibrary.getInstance();
	 HomePage hp= HomePage.getHomeInstance();
	 AddTariffPlanPage atp=AddTariffPlanPage.getTariffPlanPageInstance();	
	 
	@Given("User click on Add tariff plan button")
	public void user_click_on_Add_tariff_plan_button() throws Throwable {
		try {
			handleFrame();
			}catch(Exception e) {
			}
		f.button(hp.getAddTariffBtn());
	}
	

	@When("User Enters all the field")
	public void user_Enters_all_the_field() throws Throwable {
		try {
			handleFrame();
		}catch(Exception e) {
		}
		
		 f.button(atp.getDoneBtn());
		    f.insertText(atp.getMR(), "110");
		    f.insertText(atp.getFLM(), "120");
		    f.insertText(atp.getFIM(), "130");
		    f.insertText(atp.getFSP(), "140");
		    f.insertText(atp.getLPMC(), "11");
		    f.insertText(atp.getIPMC(), "22");
		    f.insertText(atp.getSPC(), "44");
		    
	  
	}
	@When("User Enters all the field with two dim list")
	public void user_Enters_all_the_field_with_two_dim_list(DataTable TariffPlan) throws Throwable {
		List<List<String>> data=TariffPlan.asLists(String.class);
		try {
			handleFrame();
		}catch(Exception e) {
		
		}
		
		 f.button(atp.getDoneBtn());
		   f.insertText(atp.getMR(), data.get(3).get(0));
		    f.insertText(atp.getFLM(), data.get(3).get(1));
		    f.insertText(atp.getFLM(), data.get(3).get(2));
		    f.insertText(atp.getFSP(), data.get(3).get(3));
		    f.insertText(atp.getLPMC(), data.get(2).get(4));	
		    f.insertText(atp.getIPMC(), data.get(3).get(5));	
		    f.insertText(atp.getSPC(), data.get(3).get(6));	
		
	/*	driver.findElement(By.id("rental1")).sendKeys(data.get(3).get(0));
		 driver.findElement(By.id("local_minutes")).sendKeys(data.get(3).get(1));
		 driver.findElement(By.id("inter_minutes")).sendKeys(data.get(3).get(2));
		 driver.findElement(By.name("sms_pack")).sendKeys(data.get(3).get(3));
		 driver.findElement(By.id("minutes_charges")).sendKeys(data.get(2).get(4));
		 driver.findElement(By.name("inter_charges")).sendKeys(data.get(3).get(5));
		 driver.findElement(By.id("sms_charges")).sendKeys(data.get(3).get(0));*/
		    
	  }
		
	@When("User Enters all the field with two dim map")
	public void user_Enters_all_the_field_with_two_dim_map(DataTable TariffPlan) throws Throwable {
		List<Map<String, String>> data=TariffPlan.asMaps(String.class,String.class);
		try {
			handleFrame();
		}catch(Exception e) {
		
		}
		
		FunctionalLibrary.driver.findElement(By.id("rental1")).sendKeys(data.get(3).get("MR"));
		FunctionalLibrary.driver.findElement(By.id("local_minutes")).sendKeys(data.get(3).get("FLM"));
		FunctionalLibrary.driver.findElement(By.id("inter_minutes")).sendKeys(data.get(3).get("FIM"));
		FunctionalLibrary. driver.findElement(By.name("sms_pack")).sendKeys(data.get(3).get("FSP"));
		FunctionalLibrary.driver.findElement(By.id("minutes_charges")).sendKeys(data.get(2).get("LPMC"));
		FunctionalLibrary.driver.findElement(By.name("inter_charges")).sendKeys(data.get(3).get("IPMC"));
		FunctionalLibrary.driver.findElement(By.id("sms_charges")).sendKeys(data.get(3).get("SPC"));
		
	}
	
	@When("User Enters all the field  {string},{string},{string},{string},{string},{string},{string}")
	public void user_Enters_all_the_field(String string, String string2, String FIM, String string4, String string5, String string6, String SPC) {
		FunctionalLibrary.driver.findElement(By.id("rental1")).sendKeys(string);
		FunctionalLibrary.driver.findElement(By.id("local_minutes")).sendKeys(string2);
		FunctionalLibrary.driver.findElement(By.id("inter_minutes")).sendKeys(FIM);
		FunctionalLibrary.driver.findElement(By.name("sms_pack")).sendKeys(string4);
		FunctionalLibrary.driver.findElement(By.id("minutes_charges")).sendKeys(string5);
		FunctionalLibrary. driver.findElement(By.name("inter_charges")).sendKeys(string6);
		FunctionalLibrary.driver.findElement(By.id("sms_charges")).sendKeys(SPC);
		
		
	}
	
	@When("User click on  tariff plan submit button")
	public void user_click_on_tariff_plan_submit_button() {
		
		f.button(atp.getSubmitBtn());
		
		 //driver.findElement(By.xpath("(//input[@value='submit'])")).click();
	}
	

	@Then("User should be displayed the congratulation you add tariff plan")
	public void user_should_be_displayed_the_congratulation_you_add_tariff_plan() throws Throwable {
		try {
			handleFrame();
		}catch(Exception e) {
		}
		Assert.assertEquals( FunctionalLibrary.driver.findElement(By.xpath("//h2[text()='Congratulation you add Tariff Plan']")).getText(), "Congratulation you add Tariff Plan");

	}
	public void handleFrame() throws Throwable {
		 Thread.sleep(3000);
		 FunctionalLibrary.driver.switchTo().frame( FunctionalLibrary.driver.findElement(By.xpath("//iframe[@id='flow_close_btn_iframe']")));
		 FunctionalLibrary.driver.findElement(By.id("closeBtn")).click();
		 FunctionalLibrary.driver.switchTo().parentFrame();
	}
}

