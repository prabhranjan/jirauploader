package com.pack.runner;

import java.util.HashMap;
import java.util.List;

import org.testng.annotations.Test;

import com.pack.base.base;
import com.pack.excel.dataHelper;
import com.pack.objects.testCaseUpload;

public class testCaseUploadRunner extends base{
	//public WebDriver driver;
	public List<HashMap<String, String>> datamap;
	
	public testCaseUploadRunner() {
		datamap = dataHelper.data();
	}
	
	@Test
	
	public void uploadTestCase() throws Exception{
		
		testCaseUpload TestCaseUpload = new testCaseUpload(driver);
		
		
		TestCaseUpload.loginToJira(datamap);
		TestCaseUpload.readAndInsertDataFromDataFiles(datamap);
		
	}
	
	
}
