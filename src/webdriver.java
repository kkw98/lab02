import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class webdriver {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub		
		
		 Workbook wb =null;
	        Sheet sheet = null;
	        Row row = null;
	        InputStream is = new FileInputStream("input.xlsx");
	        wb = new XSSFWorkbook(is);
	        if(wb != null){
	            //获取第一个sheet
	            sheet = wb.getSheetAt(0);
	            //获取最大行数
	            int rownum = sheet.getPhysicalNumberOfRows();
	            for (int i = 0; i<rownum; i++) {
	                row = sheet.getRow(i);
	                if(row !=null){
	                	
	                	String stu_number =  (String) getCellFormatValue(row.getCell(0));
	                	BigDecimal bd = new BigDecimal(stu_number);  
	                	stu_number = bd.toPlainString();
	                	if(stu_number.equals("3015218150")){ 
	                		System.out.println("3015218150用户信息不存在");
	                		continue;
	                    }
                    	String stu_git =  (String) getCellFormatValue(row.getCell(1));
                    	String pwd = stu_number.substring(4, 10);	
                    	
                    	//打开火狐浏览器
                		WebDriver driver = new FirefoxDriver();
                		driver.get("https://psych.liebes.top/st/");   
                		//输入用户名
                		WebElement input_number = driver.findElement(By.id("username"));
                		input_number.clear();
                		input_number.sendKeys(stu_number);
                		//输入密码
                		WebElement input_pwd = driver.findElement(By.id("password"));
                		input_pwd.clear();
                		input_pwd.sendKeys(pwd);
                		//登录
                		WebElement btn = driver.findElement(By.id("submitButton"));
                		btn.click();
                		//登录成功之后，获取URL信息
                		String git_web = driver.findElement(By.xpath("/html/body/div/div[2]/a/p")).getText();
                		//比较查询信息   
                		stu_git = stu_git.trim();
                		
                		if(stu_git.equals(git_web)||stu_git.equals(git_web+"/"))
                		{
                			System.out.println(stu_number+"用户信息一致");
                		}
                		else
                		{
                			System.out.println(stu_number+"用户信息不一致");
                		}
                		
                		driver.close();
                    	
	                }else  break;
	               
	            }
	            wb.close();            
	        }
	}	
	
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:{
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                //判断cell是否为日期格式
                if(DateUtil.isCellDateFormatted(cell)){
                    cellValue = cell.getDateCellValue();
                }else{
                    //数字
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            }
            case Cell.CELL_TYPE_STRING:{
                cellValue = cell.getRichStringCellValue().getString();
                break;
            }
            default:
                cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }

}
