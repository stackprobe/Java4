package bodewig.tools.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import charlotte.tools.FileTools;

public class XSSFExcelData {
	private String _excelFile;
	private XSSFWorkbook _wb = null;

	public XSSFExcelData(String excelFile) {
		_excelFile = excelFile;
	}

	public void load() throws Exception {
		InputStream is = new FileInputStream(_excelFile);
		try {
			_wb = new XSSFWorkbook(is);
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
	public XSSFSheet getSheet(int sheetIndex) {
		while(_wb.getNumberOfSheets() <= sheetIndex) {
			_wb.createSheet();
		}
		XSSFSheet sheet = _wb.getSheetAt(sheetIndex);

		if(sheet == null) {
			throw new RuntimeException();
		}
		return sheet;
	}

	public void setImage(
			int sheetIndex,
			int rowIndex,
			int colIndex,
			int rowPix,
			int colPix,
			int rowCount,
			int colCount,
			int rowPix2,
			int colPix2,
			byte[] pngImageData
			) {
		XSSFSheet sheet = getSheet(sheetIndex);
		int picIndex = _wb.addPicture(pngImageData, XSSFWorkbook.PICTURE_TYPE_PNG);

		XSSFDrawing drawing = sheet.createDrawingPatriarch();
		XSSFClientAnchor anchor = drawing.createAnchor(
				XSSFShape.EMU_PER_PIXEL * colPix,
				XSSFShape.EMU_PER_PIXEL * rowPix,
				XSSFShape.EMU_PER_PIXEL * colPix2,
				XSSFShape.EMU_PER_PIXEL * rowPix2,
				colIndex,
				rowIndex,
				colIndex + colCount - 1,
				rowIndex + rowCount - 1
				);

		//anchor.setAnchorType(2);

		XSSFPicture picture = drawing.createPicture(anchor, picIndex);
	}
}
