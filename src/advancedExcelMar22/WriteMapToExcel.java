package advancedExcelMar22;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteMapToExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		ArrayList<String> list1 = new ArrayList<>();
		list1.add("Pune");
		list1.add("Mumbai");

		ArrayList<String> list2 = new ArrayList<>();
		list2.add("Jaipur");
		list2.add("Nasik");
		
		ArrayList<String> list3 = new ArrayList<>();
		list3.add("Hubli");
		list3.add("Indore");
		
		HashMap<Integer, ArrayList<String>> map = new HashMap<>();
		map.put(0, list1);
		map.put(1, list2);
		map.put(2, list3);
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("MapValues");
		
		int i = 0;
		for (Integer key : map.keySet()) {
			HSSFRow row = sheet.createRow(i);
			ArrayList<String> list = map.get(key);
			for (int j=0; j< list.size();j++) {
				HSSFCell cell = row.createCell(j);
				cell.setCellValue(list.get(j));
			}
			i++;
		} 
		
		
		FileOutputStream output = new FileOutputStream("E:\\Workspace\\MapToExcel.xls");
		workbook.write(output);
		output.close();
		
	}

}
