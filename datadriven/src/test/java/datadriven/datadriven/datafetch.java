package datadriven.datadriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCalcCell;

public class datafetch {

	public static void main(String[] args) throws IOException {
		String filePath= ".\\datafiles\\population.xlsx";
		FileInputStream inputd= new FileInputStream(filePath);
		XSSFWorkbook workbook= new XSSFWorkbook(inputd);
		XSSFSheet sheet= workbook.getSheetAt(0);
		
		/*
		// Using For Loop
		
		int rows= sheet.getLastRowNum();   // finding total count of rows  
		int column= sheet.getRow(1).getLastCellNum();   // finding the total no. of cells in the 1st row
		
		for(int r=0;r<=rows;r++)
		{
                XSSFRow row= sheet.getRow(r); // Fetching the row values stored in variable 'r'
                
                  for(int c=0;c<column;c++)
                  {
                	  XSSFCell cell= row.getCell(c); // getting the cell value but dont know about the type of the cell value 
                	  switch(cell.getCellType()) // FOR CHECKING THE TYPE OF THE VALUE IN THE CELL
                	  {
                	  case STRING: System.out.print(cell.getStringCellValue());
                	               break;
                	  case NUMERIC:System.out.print(cell.getNumericCellValue());
                	  			    break;
                	  case BOOLEAN: System.out.print(cell.getBooleanCellValue());
                	                break;
                	  }
                	  System.out.print(" | ");
                  }            
            	  System.out.println();		
		
		}
		
	*/
		
		
		// USING ITERATOR
		
		Iterator sheetIterator= sheet.iterator();  /// for traversing through different sheet
		
		   while(sheetIterator.hasNext())  // pointing to the each sheet till the last sheet
		   {
			   XSSFRow row=(XSSFRow) sheetIterator.next();  // first value of the sheet will be[ row value] assigning to row variable
			   Iterator cellIterator = row.cellIterator();  // for iterating cell value inside the row
			      
			       while(cellIterator.hasNext())
			       {
			    	   
			    	   XSSFCell cell= (XSSFCell) cellIterator.next();//////Copying the current cell to cell variable **
			    	   
			    	   switch(cell.getCellType())
			    	   
			    	   {
			    	   case STRING: System.out.print(cell.getStringCellValue()); 
			    	                break;
			    	   case NUMERIC: System.out.print(cell.getNumericCellValue());
			    	                break;
			    	   case BOOLEAN: System.out.print(cell.getBooleanCellValue());
			    	                break;
			    	   }
			    	   System.out.print(" | ");
			       }
			   System.out.println();
		   }
		}

}
