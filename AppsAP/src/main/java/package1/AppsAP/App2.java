/*copy display name column in a separate sheet
remove duplicates
save it with name AP.xlsx and sheet name with Sheet1
copy and paste in files folder */
 
package package1.AppsAP;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App2 
{
    public static void main( String[] args ) throws IOException 
    {
  		//reading AP.xlsx
		FileInputStream fis = new FileInputStream(".\\files\\AP.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet= workbook.getSheet("Sheet1");
		int rows = sheet.getLastRowNum(); //119
		
		//creating appsNames.xlsx
		String filename="appsNames.xlsx";
		XSSFWorkbook workbook1 = new XSSFWorkbook();
		XSSFSheet sheet1 = workbook1.createSheet("sheet1");
		
		//0th Row
		Row headerRow = sheet1.createRow(0); //0th Row
		Cell heading1 = headerRow.createCell(0); //0 col
		heading1.setCellValue("Application");
		Cell heading2 = headerRow.createCell(1); //1 col
		heading2.setCellValue("Access Profile");
		
		int k=1;
		for(int i=1;i<=rows;i++) { 
		String displayName = sheet.getRow(i).getCell(0).getStringCellValue();
		if(displayName.startsWith("Profile")||displayName.startsWith("PROFILE")) {
			Row row = sheet1.createRow(k);
			String[] ap = displayName.split(" "); 
			String appName = ap[1];
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Azure")) {
			Row row = sheet1.createRow(k);
			String appName = "Azure";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("MSSQL")) {
			Row row = sheet1.createRow(k);
			String appName = "MSSQL";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("OracleDB")) {
			Row row = sheet1.createRow(k);
			String appName = "OracleDB";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("PAM")) {
			Row row = sheet1.createRow(k);
			String appName = "PAM";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("CIFS")) {
			Row row = sheet1.createRow(k);
			String appName = "CIFS";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("TMFT")) {
			Row row = sheet1.createRow(k);
			String appName = "TMFT";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Plutora")) {
			Row row = sheet1.createRow(k);
			String appName = "Plutora";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("SCCM")) {
			Row row = sheet1.createRow(k);
			String appName = "SCCM";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Cisco ISE")) {
			Row row = sheet1.createRow(k);
			String appName = "Cisco ISE";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("VRA")) {
			Row row = sheet1.createRow(k);
			String appName = "VRA";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Splunk")) {
			Row row = sheet1.createRow(k);
			String appName = "Splunk";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("AN")) {
			Row row = sheet1.createRow(k);
			String[] ap = displayName.split(" "); 
			String appName = "AN"+" "+ap[1];
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("AWS")) {
			Row row = sheet1.createRow(k);
			String appName = "AWS";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Account NBNAN")) {
			Row row = sheet1.createRow(k);
			String appName = "Active Network - NBNAN";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.startsWith("Account NBNCO")) {
			Row row = sheet1.createRow(k);
			String appName = "Active Directory";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.contains("Unify")) {
			Row row = sheet1.createRow(k);
			String appName = "Unify";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.contains("Maximo_HWM")) {
			Row row = sheet1.createRow(k);
			String appName = "Maximo_HWM";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else if(displayName.contains("Maximo_WWM")) {
			Row row = sheet1.createRow(k);
			String appName = "Maximo_WWM";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		else {
			Row row = sheet1.createRow(k);
			String appName = "#NA";
			Cell c = row.createCell(0);
			c.setCellValue(appName);
			Cell c1 = row.createCell(1);
			c1.setCellValue(displayName);
		k++;
		}
		}
				
		FileOutputStream fileOut = new FileOutputStream(filename);  
		workbook1.write(fileOut);   
    	fileOut.close();
    	
    	System.out.println("done");
    }
}