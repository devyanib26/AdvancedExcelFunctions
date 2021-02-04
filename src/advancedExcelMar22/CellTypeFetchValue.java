package advancedExcelMar22;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CellTypeFetchValue {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file = new File("E:\\Workspace\\MapToExcel.xls");
		FileInputStream input = new FileInputStream(file);
		
		HSSFWorkbook workbook = new HSSFWorkbook(input);
		HSSFSheet sheet = workbook.getSheet("MapValues");
		
		int maxRow = sheet.getLastRowNum();
		for (int i=0; i< maxRow; i++) {
			HSSFRow row = sheet.getRow(i);
			int maxCell = row.getLastCellNum();
			for (int j=0; j< maxCell; j++) {
				HSSFCell cell = row.getCell(j);
				
				if(cell == null) {
					System.out.println("Cell is Null");
				}
				
				else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
					System.out.println("Cell is BLANK.");
				}
				
				else if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
					System.out.println("String Value: " + cell.getStringCellValue());
				}
				
				else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
					System.out.println("Numeric Value: "+ cell.getNumericCellValue());
				}
				
				else if(cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
					System.out.println("Boolean value: "+ cell.getBooleanCellValue());
				}
				
				else {
					System.out.println("Invalid Cell");
				}
					
			}
		}
		
	}

}
