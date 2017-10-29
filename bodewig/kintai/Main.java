package bodewig.kintai;

import java.io.File;
import java.util.List;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import bodewig.tools.poi.ExcelData;
import charlotte.tools.AutoTable;
import charlotte.tools.CsvData;
import charlotte.tools.FileTools;
import charlotte.tools.MapTools;
import charlotte.tools.QueueData;
import charlotte.tools.StringTools;
import charlotte.tools.SwingTools;
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
		String ioHomeDir = FileTools.combine(System.getenv("USERPROFILE"), "Desktop");

		System.out.println("ioHomeDir: " + ioHomeDir); // test

		try(WorkDir wd = new WorkDir()) {
			String xlsFile = wd.makeSubPath() + ".xls";

			FileTools.writeAllBytes(
					xlsFile,
					FileTools.readToEnd(Main.class.getResource("res/Template.xls"))
					);

			ExcelData ed = new ExcelData(xlsFile);
			ed.load();

			String csvFile;

			{
				JFileChooser fc = new JFileChooser(ioHomeDir);

				fc.setFileFilter(new FileNameExtensionFilter("CSV ファイル (*.csv)", "csv"));
				fc.setSelectedFile(new File(ioHomeDir, "kinmujisseki.csv"));
				fc.setDialogTitle("勤務表(csv)を選択してね");

				int ret = fc.showOpenDialog(null);

				if(ret == JFileChooser.CANCEL_OPTION) {
					return;
				}
				if(ret != JFileChooser.APPROVE_OPTION) {
					throw new Exception("JFileChooser error!");
				}
				csvFile = fc.getSelectedFile().getCanonicalPath();
			}

			//Map<String, QueueData<String>> vars = MapTools.<QueueData<String>>create();
			Map<String, QueueData<String>> vars = getVars(csvFile);

			final int rowcnt = 100;
			final int colcnt = 30;

			for(int rowidx = 0; rowidx < rowcnt; rowidx++) {
				for(int colidx = 0; colidx < colcnt; colidx++) {
					String cell = ed.getCellString(0, rowidx, colidx);
					QueueData<String> qVar = vars.get(cell);

					if(qVar != null) {
						String cellNew = qVar.poll();

						if(cellNew == null) {
							cellNew = "";
						}
						ed.setCellString(0, rowidx, colidx, cellNew);
					}
				}
			}
			ed.save();

			String wFile;

			{
				JFileChooser fc = new JFileChooser(ioHomeDir);

				fc.setFileFilter(new FileNameExtensionFilter("XLS ファイル (*.xls)", "xls"));
				fc.setSelectedFile(new File(ioHomeDir, "勤務表.xls"));
				fc.setDialogTitle("出力ファイル(xls)を入力してね");

				int ret = SwingTools.showSaveDialogConfirmOverwrite(fc, null);

				if(ret == JFileChooser.CANCEL_OPTION) {
					return;
				}
				if(ret != JFileChooser.APPROVE_OPTION) {
					throw new Exception("JFileChooser error!");
				}
				wFile = fc.getSelectedFile().getCanonicalPath();
			}

			FileTools.copyFile(xlsFile, wFile);
		}
	}

	private Map<String, QueueData<String>> _vars;

	private void addVar(String key, String value) {
		QueueData<String> qVar = _vars.get(key);

		if(qVar == null) {
			qVar = new QueueData<String>();
			_vars.put(key, qVar);
		}
		qVar.add(value);
	}

	private Map<String, QueueData<String>> getVars(String csvFile) throws Exception {
		_vars = MapTools.<QueueData<String>>create();

		CsvData csv = new CsvData();
		csv.readFile(csvFile);
		AutoTable<String> tbl = csv.getTable();

		{
			String cell = tbl.get(3, 0);
			List<String> tYM = StringTools.numericTokenize(cell);
			String sY = tYM.get(0);
			String sM = tYM.get(1);

			addVar("$YM", sY + " 年 " + sM + " 月");
		}

		{
			String grp = tbl.get(6, 1);
			String grpNo = tbl.get(3, 1);

			addVar("$G", grp + " (" + grpNo + ")");
		}

		{
			String myName = tbl.get(2, 2);
			String myNo = tbl.get(5, 2);

			addVar("$N", myName + " (" + myNo + ")");
		}

		int mTotalK = 0;

		for(int y = 5; ; y++) {
			String sDate = tbl.get(2, y);

			if(sDate.equals("")) {
				break;
			}
			List<String> tDate = StringTools.numericTokenize(sDate);
			String sM = tDate.get(0);
			String sD = tDate.get(1);

			sDate = sM + " 月 " + sD + " 日";

			String sTimeSt = tbl.get(6, y);
			String sTimeEd = tbl.get(7, y);

			String sTimeY;
			String sTimeK;

			if(sTimeSt.equals("")) {
				sTimeSt = "";
				sTimeEd = "";
				sTimeY = "";
				sTimeK = "";
			}
			else {
				int mSt = sTimeToMinute(sTimeSt);
				int mEd = sTimeToMinute(sTimeEd);

				int mY = 60; // 固定
				int mK = mEd - mSt - mY;

				sTimeSt = minuteToSTime(mSt);
				sTimeEd = minuteToSTime(mEd);
				sTimeY = minuteToSTime(mY);
				sTimeK = minuteToSTime(mK);

				mTotalK += mK;
			}

			addVar("$D", sDate);
			addVar("$ST", sTimeSt);
			addVar("$ET", sTimeEd);
			addVar("$YT", sTimeY);
			addVar("$KT", sTimeK);
		}

		addVar("$TKT", minuteToSTime(mTotalK));

		return _vars;
	}

	private int sTimeToMinute(String sTime) {
		List<String> tTime = StringTools.numericTokenize(sTime);
		String sH = tTime.get(0);
		String sM = tTime.get(1);

		int h = Integer.parseInt(sH);
		int m = Integer.parseInt(sM);

		return h * 60 + m;
	}

	private String minuteToSTime(int mm) {
		int h = mm / 60;
		int m = mm % 60;

		return StringTools.zPad(h, 2) + ":" + StringTools.zPad(m, 2);
	}
}
