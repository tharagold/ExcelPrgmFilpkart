package com.objectrepository;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.resources.FunctionalLibrary;

public class AddCustomerPage {
	
	public static AddCustomerPage acp;
	FunctionalLibrary f = FunctionalLibrary.getInstance();
	  
	private AddCustomerPage() {
      PageFactory.initElements(FunctionalLibrary.driver, this);	
     }
	public static AddCustomerPage getCustomerInstance() {
		if(acp==null) {
			acp = new AddCustomerPage();
		}
		return acp;
	}
	
    @FindBy(xpath = "(//label[@for='done'])[1]")
    private WebElement doneBtn;
    
    @FindBy(id = "fname")
    private WebElement fname;
    
    @FindBy(id = "lname")
    private WebElement lname;
    
    @FindBy(id = "email")
    private WebElement email;
    
    @FindBy(id = "addr")
    private WebElement address;
    
    @FindBy(id = "telephoneno")
    private WebElement telephno;
    
    @FindBy(xpath = "(//input[@value='submit'])[1]")
    private WebElement submitBtn;

	public WebElement getDoneBtn() {
		return doneBtn;
	}

	public WebElement getFname() {
		return fname;
	}

	public WebElement getLname() {
		return lname;
	}

	public WebElement getEmail() {
		return email;
	}

	public WebElement getAddress() {
		return address;
	}

	public WebElement getTelephno() {
		return telephno;
	}

	public WebElement getSubmitBtn() {
		return submitBtn;
	}
    
    
    
    
}
