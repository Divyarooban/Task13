package Task13;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Filewrite {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		{  
			try   
			{  
			//declare file name to be create   
			String filename = "D:\\Divya\\Guvigeeks.xls";  
			//creating an instance of HSSFWorkbook class  
			HSSFWorkbook workbook = new HSSFWorkbook();  
			//invoking creatSheet() method and passing the name of the sheet to be created   
			HSSFSheet sheet = workbook.createSheet("Sheetone");   
			//creating the 0th row using the createRow() method  
			HSSFRow rowhead = sheet.createRow((short)0);  
			//creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method  
			rowhead.createCell(0).setCellValue("Name");  
			rowhead.createCell(1).setCellValue("Age");  
			rowhead.createCell(3).setCellValue("E-mail");    
			//creating the 1st row  
			HSSFRow row = sheet.createRow((short)1);  
			//inserting data in the first row  
			row.createCell(0).setCellValue("John Deo");  
			row.createCell(1).setCellValue("30");  
			row.createCell(2).setCellValue("john@test.com");   
			//creating the 2nd row  
			HSSFRow row1 = sheet.createRow((short)2);  
			//inserting data in the second row  
			row1.createCell(0).setCellValue("Jane Deo");  
			row1.createCell(1).setCellValue("28");  
			row1.createCell(2).setCellValue("john@test.com");
			//creating the 3rd row
			HSSFRow row2 = sheet.createRow((short)3);  
			//inserting data in the third row
			row2.createCell(0).setCellValue("Bob Smith");  
			row2.createCell(1).setCellValue("35");  
			row2.createCell(2).setCellValue("jacky@example.com");
			//creating the 3rd row
			HSSFRow row3 = sheet.createRow((short)4);  
			row3.createCell(0).setCellValue("Swapnil");  
			row3.createCell(1).setCellValue("37");  
			row3.createCell(2).setCellValue("joe@example.com");
			//inserting data in the Fourth row
			FileOutputStream fileOut = new FileOutputStream(filename);  
			workbook.write(fileOut);  
			//closing the Stream  
			fileOut.close();  
			//closing the workbook  
			workbook.close();  
			//prints the message on the console  
			System.out.println("Excel file has been generated successfully.");  
			}   
			catch (Exception e)   
			{  
			e.printStackTrace();  
			}  
			}  

}
}
