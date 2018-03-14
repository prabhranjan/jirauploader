package com.pack.objects;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class testCaseUpload {

	private WebDriver driver;
	
	// Variables
	/*public String noOfTestCases;
	public int numberOfTestCasesToCreate;*/
	public String testCasePageURL = "https://jira.gfs.com/jira/secure/CreateIssue!default.jspa";
	public String issueTypeValueForTestCase = "Test Case Spec";
	public String descriptionBold = "*Description*";
	public String preConditionBold = "*Pre Conditions*";
	public String stepsBold = "*Steps*";
	public String testTypeManualValue = "10541";
	public String jiraTestCaseKey;
	public String jiraTestCaseID;
	public String jiraTestCaseSummary;
	public String testCaseStepsCountString;
	
	// Locators
	public By username = By.id("login-form-username");
	public By password = By.id("login-form-password");
	public By login = By.id("login");
	public By issueType = By.id("issuetype-field");
	public By testCaseLogo = By.xpath("html/body/div[1]/section/div/div/section/form/div[1]/div[3]/div[2]/img");
	public By issueTypeSubmit = By.id("issue-create-submit");
	public By testCaseID = By.id("customfield_10124");
	public By testCaseSummary = By.id("summary");
	public By textMode = By.id("aui-uid-1");
	public By description = By.id("description");
	public By testType = By.id("customfield_10150");
	public By component = By.id("components-textarea");
	public By dueDate = By.id("duedate");
	public By affectsVersion = By.id("versions-textarea");
	public By tester = By.id("customfield_10221");
	public By labels = By.id("labels-textarea");
	public By timeTracking = By.id("timetracking");
	public By submit = By.id("issue-create-submit");
	public By jiraID = By.id("key-val");
	public By summary = By.id("summary-val");
	public By testCaseIDCreated = By.id("customfield_10124-val");
	
	public testCaseUpload(WebDriver driver){
		PageFactory.initElements(driver, this);
		this.driver = driver;
	}
	
	public boolean loginToJira(List<HashMap<String,String>> map) throws Exception{
		try{
			WebElement usernameField = driver.findElement(username);
			
			if (usernameField.isDisplayed()|| usernameField.isEnabled())
				usernameField.sendKeys(map.get(0).get("username"));
			else System.out.println("Username field is not present on screen");
			
			WebElement passwordField = driver.findElement(password);
			
			if (passwordField.isDisplayed()|| passwordField.isEnabled())
				passwordField.sendKeys(map.get(0).get("password"));
			else System.out.println("Password field is not present on screen");
			
			WebElement loginButton = driver.findElement(login);
			
			if (loginButton.isDisplayed()|| loginButton.isEnabled())
				loginButton.click();
			else System.out.println("Login button is not present on screen");
			
			try {
				WebDriverWait wait = new WebDriverWait(driver, 20);
				wait.until(ExpectedConditions.alertIsPresent());
				Alert alert = driver.switchTo().alert();
				System.out.println(alert.getText());
				alert.dismiss();
				Assert.assertTrue(alert.getText().contains("Session"));
			} catch (Exception e) {
				// exception handling
			}
						
			return true;
		}
		catch (Exception e)
		{
		e.printStackTrace();
		}
			{
			return false;
			}
	}
	
	/*public boolean catchAlertMessage(){
		try {
			WebDriverWait wait = new WebDriverWait(driver, 20);
			wait.until(ExpectedConditions.alertIsPresent());
			Alert alert = driver.switchTo().alert();
			System.out.println(alert.getText());
			alert.dismiss();
			Assert.assertTrue(alert.getText().contains("Session"));
			return true;
		} catch (Exception e) {
			return false;
		}
	}*/

	public boolean readAndInsertDataFromDataFiles(List<HashMap<String,String>> map) throws IOException, InterruptedException{
		try {
		FileInputStream testCaseDataFile = new FileInputStream(new File("./Data/TCData2.xlsx"));
		@SuppressWarnings("resource")
		XSSFWorkbook workbookTestCaseDataFields = new XSSFWorkbook(testCaseDataFile);
		XSSFSheet sheetTestCaseData = workbookTestCaseDataFields.getSheetAt(0);
		XSSFSheet sheetTestCaseStepsData = workbookTestCaseDataFields.getSheetAt(1);
		
		String numberOfTestCasesToBeCreated = map.get(0).get("TC Count");
		
		Integer numberOfTestCasesToCreate = Integer.parseInt(numberOfTestCasesToBeCreated);
		
		System.out.println(numberOfTestCasesToCreate);
		
		Integer testCaseStepsStartRow = 1;
		Integer testCaseStepsEndRow = 0;
		
		for (int i = 1; i <= numberOfTestCasesToCreate; i++) {
			
			XSSFRow rowWiseValues = sheetTestCaseData.getRow(i);
			
			// Read test case ID
			XSSFCell testCaseIDValue = rowWiseValues.getCell(3);
			String testCaseIDString = testCaseIDValue.toString();
			
			// Read Summary
			
			XSSFCell testCaseSummaryValue = rowWiseValues.getCell(1);
			String testCaseSummaryString = testCaseSummaryValue.toString();
				
			// Read Number of steps
			XSSFCell testCaseStepsCountValue = rowWiseValues.getCell(4);
			String testCaseStepsCount = testCaseStepsCountValue.toString();
			
			if (testCaseStepsCount.contains(".0"))
			{
			testCaseStepsCountString = testCaseStepsCount.substring(0, testCaseStepsCount.indexOf("."));
			}
			else 
			{
			testCaseStepsCountString = testCaseStepsCount;	
			}
			
			int testCaseStepsCountInteger = Integer.parseInt(testCaseStepsCountString);
			int testCaseStepsCountWithHeaderInteger = testCaseStepsCountInteger +1;
						
			// Read precondition data
			XSSFCell testCasePreConditionValue = rowWiseValues.getCell(0);
			String testCasePreConditionString = testCasePreConditionValue.toString();
						
			// Read pre-requisite data
			XSSFCell testCasePreRequisiteValue = rowWiseValues.getCell(2);
			String testCasePreRequisiteString = testCasePreRequisiteValue.toString();
						
			driver.get(testCasePageURL);
			
			WebElement issueTypeDropDown = driver.findElement(issueType);
			
			if (issueTypeDropDown.isDisplayed() || issueTypeDropDown.isEnabled())
				issueTypeDropDown.sendKeys(issueTypeValueForTestCase + Keys.ENTER);
			else System.out.println("Issue Type dropdown is not present on screen");
			
			WebDriverWait waitForTestCaseLogo = new WebDriverWait(driver, 400);
			@SuppressWarnings("unused")
			WebElement element = waitForTestCaseLogo.until(ExpectedConditions.visibilityOfElementLocated(testCaseLogo));
			
			WebElement issueTypeSubmitButton = driver.findElement(issueTypeSubmit);
			
			if (issueTypeSubmitButton.isDisplayed()|| issueTypeSubmitButton.isEnabled()){
				issueTypeSubmitButton.click();
			Thread.sleep(3500);}
			else System.out.println("Issue Type Submit button is not present on screen");
			
			Thread.sleep(3500);
			
			WebElement testCaseIdFiled = driver.findElement(testCaseID);
			
			if (testCaseIdFiled.isDisplayed()|| testCaseIdFiled.isEnabled())
				testCaseIdFiled.sendKeys(testCaseIDString);
			else System.out.println("Test CaseID field is not present on screen");
			
			WebElement testCaseSummaryFiled = driver.findElement(testCaseSummary);
			
			if (testCaseSummaryFiled.isDisplayed()|| testCaseSummaryFiled.isEnabled())
				testCaseSummaryFiled.sendKeys(testCaseSummaryString);
			else System.out.println("Test Case Summary field is not present on screen");
			
			testCaseStepsEndRow += testCaseStepsCountWithHeaderInteger;
			
			System.out.println("End Step row number for test case number " + i + "is : " + testCaseStepsEndRow );
			
			String textModeStatus = driver.findElement(textMode).getAttribute("aria-selected");
			
			System.out.println(textModeStatus);
			
			if (textModeStatus.contentEquals("true")) {
				
				Thread.sleep(2500);
				
				WebElement testCaseSummaryFiled2 = driver.findElement(testCaseSummary);
				
				if (testCaseSummaryFiled2.isDisplayed()|| testCaseSummaryFiled2.isEnabled())
					testCaseSummaryFiled2.sendKeys(Keys.TAB);
				else System.out.println("Test Case Summary field is not present on screen");

				Thread.sleep(2000);
				
				WebElement descriptionFieldStaticTextEnter = driver.findElement(description);
				
				if (descriptionFieldStaticTextEnter.isDisplayed()|| descriptionFieldStaticTextEnter.isEnabled())
					descriptionFieldStaticTextEnter.sendKeys(descriptionBold + Keys.ENTER + Keys.ENTER);
				else System.out.println("Description field is not present on screen");
				
				WebElement descriptionFieldPreConditionEnter = driver.findElement(description);
				
				if (descriptionFieldPreConditionEnter.isDisplayed()|| descriptionFieldPreConditionEnter.isEnabled())
					descriptionFieldPreConditionEnter.sendKeys(testCasePreConditionString + Keys.ENTER + Keys.ENTER);
				else System.out.println("Description field is not present on screen");
				
				WebElement descriptionFieldPreRequisiteEnter = driver.findElement(description);
				
				if (descriptionFieldPreRequisiteEnter.isDisplayed()|| descriptionFieldPreRequisiteEnter.isEnabled())
					descriptionFieldPreRequisiteEnter.sendKeys(preConditionBold + Keys.ENTER + Keys.ENTER + testCasePreRequisiteString + Keys.ENTER );
				else System.out.println("Description field is not present on screen");
				
				WebElement descriptionFieldStaticTextStepsEnter = driver.findElement(description);
				
				if (descriptionFieldStaticTextStepsEnter.isDisplayed()|| descriptionFieldStaticTextStepsEnter.isEnabled())
					descriptionFieldStaticTextStepsEnter.sendKeys(stepsBold + Keys.ENTER + Keys.ENTER);
				else System.out.println("Description field is not present on screen");
				
			}
			
			else if (textModeStatus.contentEquals("false")) {

				driver.close();

				System.out.println("Select Text mode instead of visual for description field manually and start again.\n Steps: \n 1. Open Jira \n 2. Go to Create Test Case page \n 3. Navigate to Description Field \n 4. Select Text Mode By clicking on Text button.");
			}
			
			System.out.println("Start Step row number for test case number " + i+1 + "is : " + testCaseStepsStartRow );
			
			
			for (int j = testCaseStepsStartRow; j <= testCaseStepsEndRow; j++) {
				
				XSSFRow rowWiseStepsValues = sheetTestCaseStepsData.getRow(j);
				
				XSSFCell testCaseStepsSerialValue = rowWiseStepsValues.getCell(0);
				String testCaseStepsSerialString = testCaseStepsSerialValue.toString();
				
				XSSFCell testCaseStepsInstructionsValue = rowWiseStepsValues.getCell(1);
				String testCaseStepsInstructionsString = testCaseStepsInstructionsValue.toString();
				
				XSSFCell testCaseStepsVerificationValue = rowWiseStepsValues.getCell(2);
				String testCaseStepsVerificationString = testCaseStepsVerificationValue.toString();
				
				WebElement descriptionFieldStepsTableEnter = driver.findElement(description);
				
				if (descriptionFieldStepsTableEnter.isDisplayed()|| descriptionFieldStepsTableEnter.isEnabled())
				descriptionFieldStepsTableEnter.sendKeys("||"+testCaseStepsSerialString + "|" + testCaseStepsInstructionsString + "|" + testCaseStepsVerificationString + "|" + Keys.ENTER);
				else System.out.println("Description field is not present on screen");
				
			}
			
			testCaseStepsStartRow += testCaseStepsCountWithHeaderInteger;
			
			WebElement testTypeDropDown = driver.findElement(testType);

			if (testTypeDropDown.isDisplayed()|| testTypeDropDown.isEnabled()){
				
			Select chooseTestType = new Select(testTypeDropDown);
			chooseTestType.selectByValue(testTypeManualValue);
			
			}
				
			else System.out.println("Description field is not present on screen");
			
			WebElement componentField = driver.findElement(component);
			
			if (componentField.isDisplayed()||componentField.isEnabled()){
				componentField.sendKeys(map.get(0).get("Component"));
			}
			else System.out.println("Component field is not present on screen");
			
			WebElement dueDateField = driver.findElement(dueDate);
			
			if (dueDateField.isDisplayed()||dueDateField.isEnabled()){
				dueDateField.click();
			}
			else System.out.println("Due Date field is not present on screen");
			
			
			WebElement affectsVersionField = driver.findElement(affectsVersion);
			
			if (affectsVersionField.isDisplayed()||affectsVersionField.isEnabled()){
				affectsVersionField.sendKeys(map.get(0).get("AffVer"));
			}
			else System.out.println("Affects version field is not present on screen");
			
			WebElement dueDateField2 = driver.findElement(dueDate);
			
			if (dueDateField2.isDisplayed()||dueDateField2.isEnabled()){
				dueDateField2.click();
			}
			else System.out.println("Due Date field is not present on screen");
			
			WebElement testerField = driver.findElement(tester);
			
			if (testerField.isDisplayed()||testerField.isEnabled()){
				testerField.sendKeys(map.get(0).get("Tester"));
			}
			else System.out.println("Tester field is not present on screen");
			
			WebElement labelOneEnterField = driver.findElement(labels);
			
			if (labelOneEnterField.isDisplayed()||labelOneEnterField.isEnabled()){
				labelOneEnterField.sendKeys(map.get(0).get("label1"));
			}
			else System.out.println("Label field is not present on screen");
			
			WebElement timeTrackingField = driver.findElement(timeTracking);
			
			if (timeTrackingField.isDisplayed()||timeTrackingField.isEnabled()){
				timeTrackingField.click();
			}
			else System.out.println("Time Tracking field is not present on screen");
			
			WebElement labelTwoEnterField = driver.findElement(labels);
			
			if (labelTwoEnterField.isDisplayed()||labelTwoEnterField.isEnabled()){
				labelTwoEnterField.click();
				labelTwoEnterField.sendKeys(map.get(0).get("label2"));
			}
			else System.out.println("Label field is not present on screen");
			
			WebElement timeTrackingField2 = driver.findElement(timeTracking);
			
			if (timeTrackingField2.isDisplayed()||timeTrackingField2.isEnabled()){
				timeTrackingField2.click();
			}
			else System.out.println("Time Tracking field is not present on screen");
			
			WebElement SubmitButton = driver.findElement(submit);
			
			if (SubmitButton.isDisplayed()||SubmitButton.isEnabled()){
				SubmitButton.click();
			}
			else System.out.println("Submit button is not present on screen.");
			
			// write results to excel
			
			WebElement jiraIDLink = driver.findElement(jiraID);
			
			if (jiraIDLink.isDisplayed()|| jiraIDLink.isEnabled()){
			
			jiraTestCaseKey = jiraIDLink.getText();}
				
			else System.out.println("Jira Test Case ID field is not present on screen");
			
			WebElement jiraTestCaseIDField = driver.findElement(testCaseIDCreated);
			
			if (jiraTestCaseIDField.isDisplayed()|| jiraTestCaseIDField.isEnabled()){
			
			jiraTestCaseID = jiraTestCaseIDField.getText();}
				
			else System.out.println("Jira Test Case ID field is not present on screen.");
			
			WebElement jiraTestCaseSummaryField = driver.findElement(summary);
			
			if (jiraTestCaseSummaryField.isDisplayed()|| jiraTestCaseSummaryField.isEnabled()){
			
			jiraTestCaseSummary = jiraTestCaseSummaryField.getText();}
				
			else System.out.println("Jira Test Case ID field is not present on screen");
			
			@SuppressWarnings("resource")
			XSSFWorkbook workBookResults = new XSSFWorkbook();
			XSSFSheet sheetResults = workBookResults.createSheet("Results");
			XSSFRow headerRow = sheetResults.createRow((short) 0);
			headerRow.createCell((short) 0).setCellValue("Test Case Key");
			headerRow.createCell((short) 1).setCellValue("Test Case ID");
			headerRow.createCell((short) 2).setCellValue("Summary");
			
			 XSSFRow rows = sheetResults.createRow((short) i);
	            rows.createCell((short) 0).setCellValue(jiraTestCaseKey);
	            rows.createCell((short) 1).setCellValue(jiraTestCaseID);
	            rows.createCell((short) 2).setCellValue(jiraTestCaseSummary);
		
	            FileOutputStream fileOut = new FileOutputStream("./uploadResult.xlsx");
	            workBookResults.write(fileOut);
	            fileOut.close();
		}
		return true;
		
		} catch (Exception e) {
			return false;
		}
	}
	
	
	
}
