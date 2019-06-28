package sample.testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelsheet {
	public static void main(String[] args) throws IOException {


		
		try {
			File loc=new File("C:\\Users\\Dineshkumar\\eclipse-workspace\\testing\\excel\\data.xlsx");
			FileInputStream stream=new FileInputStream(loc);
			Workbook w= new XSSFWorkbook(stream);
			Sheet s=w.getSheet("Sheet1");
			Row r = s.getRow(1);
			Cell c = r.getCell(2);
			System.out.println(c);
			
			
			
		} catch (FileNotFoundException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

}
