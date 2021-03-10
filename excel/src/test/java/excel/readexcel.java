package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class readexcel {
	public static void main (String[] args) throws FileNotFoundException, IOException {
		
		//Reading the contents of the file and storing in a hashmap. File is read using scanner.
		File fil1 = new File("C:\\Users\\Vasanth\\Desktop\\parse.txt");
		HashMap<String, String> map = new HashMap<String, String>();
		String[] parts;
		Scanner myReader = new Scanner(fil1);
		while (myReader.hasNextLine())
		{
			String data = myReader.nextLine();
			if (!data.contains("record"))
			{
				
				parts = data.split(" +", 2);			
				String key = parts[0].trim();
				String value = parts[1].trim();
				map.put(key, value);
			}
			else
			{
				System.out.println("Ignoring as it doesnt hold a key value pair");
			}
		}
		
		//Writing Hashmap to excel & csv
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sh = wb.createSheet("Test Data");
		int rowno = 0;	
		for(Map.Entry entry:map.entrySet())
		{
			XSSFRow row = sh.createRow(rowno++);
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());
		}		
		wb.write(new FileOutputStream(".\\Datafiles\\hashtoexcel.xlsx"));
		wb.write(new FileOutputStream(".\\Datafiles\\hashexcel.csv")); 
		wb.close();
		}
		
	}


