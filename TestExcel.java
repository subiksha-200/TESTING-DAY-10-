package excel;

import java.io.IOException;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class TestExcel {
  @Test(dataProvider="testData")
  public void addMethod(double num1,double num2)throws IOException {
	  double result=num1+num2;
	  ExcelUtility.updateExcel();
  }
  @DataProvider
  public Object[][] testData()throws IOException{
	  Object[][]data=ExcelUtility.readExcel();
	  return data;
  }
}
