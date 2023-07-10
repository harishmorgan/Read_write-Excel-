package read.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ReadWriteExcelFile {

	public static void main(String[] args) throws IOException {
		
		// Step 1 – Create a blank workbook.

		XSSFWorkbook wb = new XSSFWorkbook();

		//Step 2 – Create a sheet and pass name of the sheet.
		
		XSSFSheet sheet = wb.createSheet("Write_TestData");
		
		ArrayList<Object[]> data = new ArrayList<Object[]>();
        data.add(new String[] { "Name", "Id", "Salary" });
        data.add(new Object[] { "Jim", "001A", 10000 });
        data.add(new Object[] { "Jack", "1001B", 40000 });
        data.add(new Object[] { "Tim", "2001C", 20000 });
        data.add(new Object[] { "Gina", "1004S", 30000 });
		
		// Step 3 – Create a Row. A spreadsheet consists of rows and cells. It has a grid layout.
        int rownum = 0;
        for (Object[] employeeDetails : data) { 
		XSSFRow row = sheet.createRow(rownum++);
		
		//Step 4 – Create cells in a row. 
		//A row is a collection of cells. 
		//When you enter data in the sheet, it is always stored in the cell.
		 int cellnum = 0;
         for (Object obj : employeeDetails) {
		
		XSSFCell cell = row.createCell(cellnum++);
		
		  // Set value to cell
        if (obj instanceof String)
            cell.setCellValue((String) obj);
        else if (obj instanceof Double)
            cell.setCellValue((Double) obj);
        else if (obj instanceof Integer)
            cell.setCellValue((Integer) obj);
    }
}
		
		//Step 5 – Write data to an OutputStream.
        try {
        	 
            // Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("EmployeeDetails.xlsx"));
           wb.write(out);
            out.close();
            System.out.println("EmployeeDetails.xlsx has been created successfully");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            wb.close();
        }
    }
 
}