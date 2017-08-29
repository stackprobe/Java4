package bodewig_test.tools.poi;

import bodewig.tools.poi.ExcelData;
import charlotte.tools.FileTools;

public class ExcelDataTest {
	public static void main(String[] args) {
		try {
			main2();
		}
		catch(Throwable e) {
			e.printStackTrace();
		}
		System.exit(0);
	}

	private static void main2() throws Exception {
		FileTools.copyFile(
				"C:/var/Tests/ExcelData_0001.xls",
				"C:/temp/1.xls"
				);
		ExcelData ed = new ExcelData(
				"C:/temp/1.xls"
				);
		ed.load();
		ed.setCellString(0, 1, 1, "ahe-ahe");
		ed.save();
	}
}
