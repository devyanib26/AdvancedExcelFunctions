package advancedExcelMar22;

import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteListToExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		ArrayList<String> list = new ArrayList<>();
		list.add("Pune");
		list.add("Mumbai");
		list.add("Jaipur");
		list.add("Nasik");
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("List");
		HSSFRow row = sheet.createRow(0);

		for (int i=0; i< list.size();i++) {
			HSSFCell cell = row.createCell(i);
			cell.setCellValue(list.get(i));
		}
		
		FileOutputStream output = new FileOutputStream("E:\\Workspace\\ListToExcel.xls");
		workbook.write(output);
		output.close();
		
	}

}
