package practiceKM;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excelUtility {
	
public String getDataFromExcel(String sheetName, int rowNum, int celNum) throws Throwable {
		
		FileInputStream fis = new FileInputStream("./TestData/TestScriptData.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheet(sheetName);
		Row row = sheet.getRow(rowNum);
		Cell cel = row.getCell(celNum);
		String data = cel.getStringCellValue();
		wb.close();	
		return data;

}
}
