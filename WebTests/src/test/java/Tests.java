import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonObject;

import io.restassured.RestAssured;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class Tests {

	@Test
	public void TestGet() {
		RestAssured.useRelaxedHTTPSValidation();
		RestAssured.baseURI = "https://localhost:44345";
		RequestSpecification httpRequest = RestAssured.given();
		
        Response response = httpRequest.get("/api/ToDo");
        
        System.out.println("Response Body is =>  " + response.asString());
	}
	
	@Test
	public void TestPost() {
		RestAssured.useRelaxedHTTPSValidation();
		RestAssured.baseURI = "https://localhost:44345";
		RequestSpecification httpRequest = RestAssured.given();
		httpRequest.header("Content-Type", "application/json");
		JsonObject requestParams = new JsonObject();
		requestParams.addProperty("name", "walk cat");
		requestParams.addProperty("isComplete", true);
		httpRequest.body(requestParams.toString());
        Response response = httpRequest.post("/api/ToDo");
        
        System.out.println("Response Body is =>  " + response.asString());
	}
	
	@Test
	public void BuildJSON() throws Exception {
		Workbook oExcelFile = WorkbookFactory.create(new FileInputStream(new File("C:\\Users\\Adam\\Desktop\\testfile.xlsx")));
		Sheet oSheet = oExcelFile.getSheetAt(0);
		HashMap<String,Object> fields = new HashMap<String,Object>();
		Pattern pattern = Pattern.compile("\\d+");
		Matcher matcher;
		int fieldLength = 0;
		boolean isNumeric;
		
		// Read in all rows to hash map
		Iterator<Row> rows = oSheet.rowIterator();
		while (rows.hasNext()) {
			Row row = (Row) rows.next();
			if (row.getRowNum() != 0) {
				String s = row.getCell(1).getStringCellValue();
				matcher = pattern.matcher(s);
				matcher.find();
				fieldLength = Integer.parseInt(matcher.group());
				isNumeric = row.getCell(2).getStringCellValue().contains("Character") ? false : true;
				fields.put(row.getCell(0).getStringCellValue(), getFieldValue(isNumeric, fieldLength));
			}
		}
		oExcelFile.close();
		
		// Build JSON payload
		Gson payload = new GsonBuilder().setPrettyPrinting().create();
		String jsonPayload = payload.toJson(fields);
		System.out.println("jsonPayload = \n" + jsonPayload);
		
		// Save to file
		FileUtils.writeStringToFile(new File("C:\\Users\\Adam\\Desktop\\test.json"), jsonPayload, "UTF-8");
		
	}
	
	private static Object getFieldValue(boolean isNumber, int fieldLength) {
		String strValue;
		if (isNumber) {
			strValue = StringUtils.repeat("9", fieldLength);
			return Integer.parseInt(strValue);
		}
		else {
			strValue = StringUtils.repeat("A", fieldLength);
			return strValue;
		}
		
	}
}
