package bodewig.tools.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import charlotte.tools.FileTools;

public class HSSFExcelData {
	private String _excelFile;
	private HSSFWorkbook _wb = null;

	public HSSFExcelData(String excelFile) {
		_excelFile = excelFile;
	}

	public void load() throws Exception {
		InputStream is = new FileInputStream(_excelFile);
		try {
			_wb = new HSSFWorkbook(new POIFSFileSystem(is));
		}
		finally {
			FileTools.close(is);
		}
	}

	public void save() throws Exception {
		OutputStream os = new FileOutputStream(_excelFile);
		try {
			_wb.write(os);
		}
		finally {
			FileTools.close(os);
		}
		_wb = null;
	}

	/**
	 *
	 * @param sheetIndex 0 から
	 * @return
	 */
	public HSSFSheet getSheet(int sheetIndex) {
		while(_wb.getNumberOfSheets() <= sheetIndex) {
			_wb.createSheet();
		}
		HSSFSheet sheet = _wb.getSheetAt(sheetIndex);

		if(sheet == null) {
			throw new RuntimeException();
		}
		return sheet;
	}

	public void setImage(int sheetIndex, int rowIndex, int colIndex, byte[] pngImageData, double scale) {
		HSSFSheet sheet = getSheet(sheetIndex);
		int picIndex = _wb.addPicture(pngImageData, HSSFWorkbook.PICTURE_TYPE_PNG);

		HSSFCreationHelper helper = (HSSFCreationHelper)_wb.getCreationHelper();
		HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
		HSSFClientAnchor anchor = helper.createClientAnchor();
		anchor.setRow1(rowIndex);
		anchor.setCol1(colIndex);

		HSSFPicture picture = patriarch.createPicture(anchor, picIndex);
		picture.resize(scale);
	}
}
