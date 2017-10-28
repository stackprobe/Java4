package bodewig.kintai;

import bodewig.tools.poi.ExcelData;
import charlotte.tools.CsvData;
import charlotte.tools.FileTools;
import charlotte.tools.WorkDir;

public class Main {
	public static void main(String[] args) {
		try {
			new Main().main2();
		}
		catch(Throwable e) {
			e.printStackTrace();
		}
	}

	private void main2() throws Exception {
		try(WorkDir wd = new WorkDir()) {
			String templateFile = wd.makeSubPath() + ".xls";

			FileTools.writeAllBytes(
					templateFile,
					FileTools.readToEnd(Main.class.getResource("res/Template.xls"))
					);

			ExcelData ed = new ExcelData(templateFile);

			// test
			// test
			// test

			ed.load();

			ed.setCellString(0, 2, 2, ed.getCellString(0, 8, 2));

			ed.save();

			FileTools.copyFile(templateFile, "C:/temp/1.xls");

			CsvData.Stream reader = new CsvData.Stream("C:/var/dat/kinmujisseki_20171027.csv");
			try {
				reader.readOpen();

				System.out.println("" + reader.readCell());
				System.out.println("" + reader.readCell());
				System.out.println("" + reader.readCell());
			}
			finally {
				reader.readClose();
			}
		}
	}
}
