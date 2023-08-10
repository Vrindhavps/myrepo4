package excelPackage1;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode
{
	static FileInputStream f;//class
	static XSSFWorkbook w;//class
    static XSSFSheet sh;//class
    public static String readStringData(int i,int j) throws IOException
    {
    	f=new FileInputStream("C:\\Users\\vrind\\Desktop\\ExcelSample.xlsx");// add throws fileNotFound exception
    	w=new XSSFWorkbook(f);// add throws IO Exception
    	sh=w.getSheet("Sheet1");
    	XSSFRow r=sh.getRow(i); //row no
    	XSSFCell c=r.getCell(j);
    	return c.getStringCellValue();
}
    public static String readIntegerData(int i,int j) throws IOException
    {
    	f=new FileInputStream("C:\\Users\\vrind\\Desktop\\ExcelSample.xlsx");// add throws fileNotFound exception
    	w=new XSSFWorkbook(f);//add throws IO Exception
    	sh=w.getSheet("Sheet1");
    	XSSFRow r=sh.getRow(i);
    	XSSFCell c=r.getCell(j);
    	int a=(int) c.getNumericCellValue();//to convert string to integer use valueOf method
    	return String.valueOf(a);
    }

	}
