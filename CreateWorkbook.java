package Task13;
import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;   
import org.apache.poi.ss.usermodel.Workbook;

public class CreateWorkbook {

	public static void main(String[] args) throws FileNotFoundException, IOException{
		// TODO Auto-generated method stub
		Workbook wb = new HSSFWorkbook();   
		//creates an excel file at the specified location  
		OutputStream fileOut = new FileOutputStream("D:\\Divya\\GuviGeeks.xls"); 
		System.out.println("Excel File has been created successfully.");   
		wb.write(fileOut);    
	}

}
