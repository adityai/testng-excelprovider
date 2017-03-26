package com.iaditya.testng.excel_provider;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

/**
 * Sample test class to demonstrate use of excel file as the data source of the test data.
 * 
 * @author adityai
 *
 */
public class NewTest {

    private Workbook workbook = null;
    private static Object[][] testData = null;

    public NewTest(){
    	initExcelReader();
    }
	/**
	 * Sample test case that pulls test data from excel file
	 * 
	 * @param a
	 * @param b
	 */
  @Test(dataProvider="excelDataProvider")
  public void testCase001(String a, String b) {
	  Assert.assertEquals(a, "testData_001_1");
	  Assert.assertEquals(b, "testData_001_2");
  }

  /**
   * Sample test case that pulls test data from excel database
   * 
   * @param a
   * @param b
   */
  @Test(dataProvider="excelDataProvider")
  public void testCase002(String a, String b) {
	  Assert.assertEquals(a, "testData_002_1");
	  Assert.assertEquals(b, "testData_002_2");
  }
  

  /**
   * Load data from excel into an object array
   */
  private void initExcelReader() {
	  int i = 0;
	  int j = 0;
	  int rows = 0;
	  int columns = 0;
	  
      try {
          FileInputStream excelFile = new FileInputStream(new File("src/test/resources/ExcelTestData.xlsx"));
          workbook = new XSSFWorkbook(excelFile);
          Sheet sheet = workbook.getSheetAt(0);
          rows = sheet.getPhysicalNumberOfRows();
          columns = sheet.getRow(0).getPhysicalNumberOfCells();
          
          Iterator<Row> iterator = sheet.iterator();
          testData = new String[rows][columns];

          while (iterator.hasNext()) {
              Row currentRow = iterator.next();
              Iterator<Cell> cellIterator = currentRow.iterator();
        	  j = 0;
              while (cellIterator.hasNext()) {
                  Cell currentCell = cellIterator.next();
                  if (currentCell.getCellTypeEnum() == CellType.STRING) {
                      System.out.print(currentCell.getStringCellValue() + "--");
                      testData[i][j] = currentCell.getStringCellValue();
                  } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                      System.out.print(currentCell.getNumericCellValue() + "--");
                      testData[i][j] = currentCell.getNumericCellValue();
                  }
                  j++;
              }
              i++;
              System.out.println();
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      } catch (IOException e) {
          e.printStackTrace();
      }
  }

  /**
   * AfterSuite method to close the Excel file
   */
  @AfterSuite
  private void closeExcel() {
	  try {
		workbook.close();
	} catch (IOException e) {
		System.out.println("ERROR: While closing workbook");
		e.printStackTrace();
	}
	  
  }

  /**
   * DataProvider loaded from test data present in the Excel file
   * @param method
   * @return
   */
  @DataProvider(name="excelDataProvider")
  private Object[][] getData(Method method) {
	  Object[][] data = null;
	  System.out.println("**** " + testData.length + " " + testData[1][1]);
	  for (int i = 1; i < testData.length; i++) {
		  System.out.println(testData[i][1]);
		  if (testData[i][1].equals(method.getName())) {
			  data = new Object[][] {{ testData[i][2], testData[i][3] }};
		  }
	  }
	  return data;
  }
  
}
  
