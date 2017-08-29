package bodewig.tools.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import charlotte.tools.FileTools;
import charlotte.tools.StringTools;
import charlotte.tools.TimeData;

public class ExcelData {
	private String _excelFile;
	private Workbook _wb = null;

	public ExcelData(String excelFile) {
		_excelFile = excelFile;
	}

	public void load() throws Exception {
		InputStream is = new FileInputStream(_excelFile);
		try {
			_wb = WorkbookFactory.create(is);
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
	public Sheet getSheet(int sheetIndex) {
		while(_wb.getNumberOfSheets() <= sheetIndex) {
			//System.out.println("ExcelData_create_sheet"); // test
			_wb.createSheet();
		}
		Sheet sheet = _wb.getSheetAt(sheetIndex);

		if(sheet == null) {
			throw new RuntimeException("sheet is null");
		}
		return sheet;
	}

	/**
	 *
	 * @param sheet
	 * @param rowIndex 0 から
	 * @return
	 */
	public Row getRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);

		if(row == null) {
			//System.out.println("ExcelData_create_row: " + rowIndex); // test
			row = sheet.createRow(rowIndex);
		}
		return row;
	}

	/**
	 *
	 * @param row
	 * @param colIndex 0 から
	 * @return
	 */
	public Cell getCell(Row row, int colIndex) {
		Cell cell = row.getCell(colIndex);

		if(cell == null) {
			//System.out.println("ExcelData_create_cell: " + colIndex); // test
			cell = row.createCell(colIndex);
		}
		return cell;
	}

	public Cell getCell(int sheetIndex, int rowIndex, int colIndex) {
		return getCell(getRow(getSheet(sheetIndex), rowIndex), colIndex);
	}

	public String getCellString(int sheetIndex, int rowIndex, int colIndex) {
		Cell cell = getCell(sheetIndex, rowIndex, colIndex);

		switch(cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();

		case Cell.CELL_TYPE_NUMERIC:
			if(DateUtil.isCellDateFormatted(cell)) {
				return new TimeData(cell.getDateCellValue()).getString();
			}
			return "" + cell.getNumericCellValue();

		case Cell.CELL_TYPE_BOOLEAN:
			return StringTools.toString(cell.getBooleanCellValue());

		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();

		case Cell.CELL_TYPE_ERROR:
			return "" + cell.getErrorCellValue();

		case Cell.CELL_TYPE_BLANK:
			return "";
		}
		return "";
	}

	public void setCellString(int sheetIndex, int rowIndex, int colIndex, String value) {
		Cell cell = getCell(sheetIndex, rowIndex, colIndex);

		cell.setCellValue(value); // XXX
	}
}
