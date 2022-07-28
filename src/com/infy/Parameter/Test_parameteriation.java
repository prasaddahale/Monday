package com.infy.Parameter;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test_parameteriation 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException
	{
		String path = "D:\\workplace\\Parameteriation\\Data\\prasad_parameter_test1.xlsx";
		FileInputStream f = new FileInputStream(path);
		
//		String r1 = WorkbookFactory.create(f).getSheet("Sheet1").getRow(0).getCell(1).getStringCellValue();
//        System.out.println(r1);
        
//        double r2 = WorkbookFactory.create(f).getSheet("Sheet1").getRow(0).getCell(1).getNumericCellValue();
//        System.out.println(r2);
		
//		CellType type = WorkbookFactory.create(f).getSheet("sheet1").getRow(0).getCell(1).getCellType();
//		System.out.println(type);
		
		Workbook wbook = WorkbookFactory.create(f);
		String r1 = wbook.getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
		 System.out.println(r1);
	        
       double r2 = wbook.getSheet("Sheet1").getRow(0).getCell(1).getNumericCellValue();
       System.out.println(r2);
       //get row count
       //get column count
       //nested for loop
       //datatype not stable
       
      int p = wbook.getSheet("sheet1").getLastRowNum();
    short p1 = wbook.getSheet("sheet1").getRow(0).getLastCellNum();
    System.out.println("row count = "+p+"coloum count = "+p1);
    
//    for(int i=0;i<=p;i++)
//    {
//    	for(int j=0;j<p1;j++)
//    	{
//    		Object result = wbook.getSheet("sheet1").getRow(i).getCell(j);
//    		System.out.print(result +"  ||  ");
//    	}
//    	System.out.println(" ");
//    }
    
    for(int i=0;i<=p;i++)
    {
    	for(int j=0;j<p1;j++)
    	{
    		Cell type = wbook.getSheet("sheet1").getRow(i).getCell(j);
    		switch(type.getCellType())
    		{
    		case STRING : System.out.println(type.getStringCellValue());break;
    		case NUMERIC : System.out.println(type.getNumericCellValue());break;
    		}
    		System.out.print(" || ");
    		
    	}
    	 System.out.println(" ");
    }
   
	}

}
