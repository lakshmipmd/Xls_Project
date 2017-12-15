import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


//Read a excel file without header & Write hard-coded contents into the same excel file
public class XLSXReaderWriter {

	@SuppressWarnings("deprecation")
	public static void main(String[] args) {
		
		try {
		File excel = new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sample.xls");
		FileInputStream fis = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(fis);
		HSSFSheet sheet = book.getSheetAt(0);
		
		Iterator<Row> itr = sheet.iterator();
		
		//Iterating over the excel file in java
		while(itr.hasNext()){
			Row row = itr.next();
			
			//Iterating over each column of excel file
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()){
				
				Cell cell = cellIterator.next();

				switch(cell.getCellType()){
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue() );
					break;
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue() );
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.println(cell.getBooleanCellValue() );
					break;
				default:
				}
				
			}
			System.out.println(" ");
			
		}
		
		//writing data into xls file
		Map<String, Object[]> newData= new HashMap<String, Object[]>();
		newData.put("7", new Object[] {7d, "Sonya","75k", "SALES", "Rupert"});
		newData.put("8", new Object[] {8d, "Kris","85k", "SALES", "Rupert"});
		newData.put("9", new Object[] {9d, "Dave","90k", "SALES", "Rupert"});
		
		Set<String> newRows = newData.keySet();
		int rownum = sheet.getLastRowNum();
		
		for(String key : newRows){
			Row row = sheet.createRow(rownum++);
			Object[] objArr = newData.get(key);
			int cellNum = 0;
			
			for(Object obj : objArr){
				Cell cell = row.createCell(cellNum++);
				if(obj instanceof String){
					cell.setCellValue((String)obj);
				}else if(obj instanceof Boolean){
					cell.setCellValue((Boolean)obj);
				}else if(obj instanceof Date){
					cell.setCellValue((Date)obj);
				}else if(obj instanceof Double){
					cell.setCellValue((Double)obj);
				}
			}
		}
		
		//open an output stream to save written data into excel file
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		System.out.println("Writing an excel file finished....");
		
		//close workbook, output stream & excel file to prevent leak
		os.close();
		book.close();
	    fis.close();
		
		} catch (FileNotFoundException fe){
			fe.printStackTrace();
		}
		catch ( IOException e) {
			e.printStackTrace();
		}
		
	}

}
