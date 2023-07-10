package read.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
public class ReadSpecificData {
	
	public static String path = "E:\\New folder\\DataDriven.xlsx";
	WebDriver driver;

	public static void main(String[] args) throws IOException {
		
		// Step 1 – To locate the location of file.
		
		
		File file = new File(path);
		
		//Step 2 – Instantiate FileInputStream to read from the file specified.
		
		
		
			FileInputStream fis = new FileInputStream(file);
			
			// Step 3 – Create object of XSSFWorkbook class
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			
			// Step 4 – To read excel sheet by sheet name	
			XSSFSheet sheet = wb.getSheet("test steps");
			
			/*To access data from the XLSX file, use of  the following methods:
				
				getRow(int rownum)
				getCell(int cellnum)
				getStringCellValue()
				getNumericCellValue() */
			
			//Find number of rows in excel file
	        int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();      
	        System.out.println("row count:"+rowCount);
	        
	      //iterate over all the row to print the data present in each cell.
	        for(int i=0;i<=rowCount;i++){
	             
	            //get cell count in a row
	            int cellcount=sheet.getRow(i).getLastCellNum();       
	            
	          //iterate over each cell to print its value       
	            for(int j=0;j<cellcount;j++){
	                System.out.print(sheet.getRow(i).getCell(j).getStringCellValue().toString() +"||");
	            }
	            System.out.println();
	        }
	}

}
