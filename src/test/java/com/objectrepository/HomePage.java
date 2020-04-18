package com.objectrepository;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.resources.FunctionalLibrary;

public class HomePage {
	
	public static HomePage hp;
	public HomePage() {
		FunctionalLibrary f = FunctionalLibrary.getInstance();
	
       PageFactory.initElements(FunctionalLibrary.driver, this);
	
	}
	public static HomePage getHomeInstance() {
		if(hp==null) {
			hp = new HomePage();
		}
		return hp;
	}
	
	@FindBy(xpath = "(//a[text()='Add Customer'])[1]")
	private  WebElement addCustomerBtn;

	@FindBy(xpath = "(//a[text()='Add Tariff Plan'])")
	private WebElement addTariffBtn;

	public WebElement getAddCustomerBtn() {
		return addCustomerBtn;
	}

	public WebElement getAddTariffBtn() {
		return addTariffBtn;
	}
	
	

}
