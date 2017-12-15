import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//Read a basic excel file without any header content & write the same into another excel file
public class XLSXReaderWriter4 {

	private static final File excel = 
			new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sample.xls");
	private static final File outExcel = 
			new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sampleOutput.xls");
	
	public static void main(String[] args) throws IOException{
	
		List<ExcelFile> contentList = getContentListfromExcel();
		writeContentListToExcel(contentList);
		System.out.println(contentList);
	}

	private static void writeContentListToExcel(List<ExcelFile> contentList) throws IOException {

		HSSFWorkbook outbook = new HSSFWorkbook();
		HSSFSheet sheet = outbook.createSheet("Sheet1");
		
		HSSFRow row = sheet.createRow(0);
		int rowIndex = 0;
		
		for(ExcelFile content : contentList){
			row = sheet.createRow(rowIndex++);
			int cellIndex = 0;
			//first place in row is ID
			row.createCell(cellIndex++).setCellValue(content.getID());
			//second place in row is name
			row.createCell(cellIndex++).setCellValue(content.getName());
			//third place in row is marks in Salary
			row.createCell(cellIndex++).setCellValue(content.getSalary());
			//fourth place in row is marks in Department
			row.createCell(cellIndex++).setCellValue(content.getDepartment());
			//fifth place in row is marks in Manager
			row.createCell(cellIndex++).setCellValue(content.getManager());
			}
		
			 FileOutputStream fos = new FileOutputStream(outExcel);
			 outbook.write(fos);
			 System.out.println(outExcel + " is successfully written");
			 
			 outbook.close();
			 fos.close();
			 
		
	}

	private static List<ExcelFile> getContentListfromExcel() throws IOException{
		
		List<ExcelFile> contentList = new ArrayList<>();
		
		FileInputStream file = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(file);
		HSSFSheet sheet = book.getSheetAt(0);
		HSSFRow row = sheet.getRow(0);
			
		//Iterate through each rows one by one
        for(int i = 0 ; i < row.getLastCellNum(); i++){
           	HSSFRow nextRow = sheet.getRow(i);
           	ExcelFile content = new ExcelFile();

           	for(int j = 0 ; j < nextRow.getLastCellNum() ; j++){
           		HSSFCell nextCell = nextRow.getCell(j);
                	
               	if(nextCell.getColumnIndex() == 0)
                	content.setID(nextCell.getNumericCellValue());
                else if(nextCell.getColumnIndex() == 1)
                	content.setName(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 2)
                	content.setSalary(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 3)
                	content.setDepartment(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 4)
                	content.setManager(nextCell.getStringCellValue());
                	
                }//end of for loop of cells
               contentList.add(content);
              }//End of for loop of rows
			book.close();
			file.close();
			
		return contentList;
	}
}
