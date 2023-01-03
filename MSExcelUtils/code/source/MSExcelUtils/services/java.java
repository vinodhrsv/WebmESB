package MSExcelUtils.services;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
// --- <<IS-END-IMPORTS>> ---

public final class java

{
	// ---( internal utility methods )---

	final static java _instance = new java();

	static java _newInstance() { return new java(); }

	static java _cast(Object o) { return (java)o; }

	// ---( server methods )---




	public static final void XlsToXlsx (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(XlsToXlsx)>> ---
		// @sigtype java 3.5
		// [i] field:0:required fileNameXls
		// [o] field:0:required fileNameXlsx
			
		IDataCursor pipelineCursor = pipeline.getCursor();
		String fileNameXls = IDataUtil.getString(pipelineCursor, "fileNameXls");
		String convertedFileName = null;
		try { 
			convertedFileName = convertXls( pipeline,fileNameXls);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
				
		IDataUtil.put(pipelineCursor, "fileNameXlsx", convertedFileName);
		pipelineCursor.destroy();
		// --- <<IS-END>> ---

                
	}



	public static final void XlsxToXls (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(XlsxToXls)>> ---
		// @sigtype java 3.5
		// [i] field:0:required fileNameXlsx
		// [o] field:0:required fileNameXls
		IDataCursor pipelineCursor = pipeline.getCursor();
		String fileNameXlsx = IDataUtil.getString(pipelineCursor, "fileNameXlsx");
		String convertedFileName = null;
		try {
			convertedFileName = convertXlsx( pipeline,fileNameXlsx);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
				
		IDataUtil.put(pipelineCursor, "fileNameXls", convertedFileName);
		pipelineCursor.destroy();
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	
	private static String convertXlsx(IData pipeline, String fileName)throws Exception{
		
		IData data = IDataFactory.create();
		IDataCursor outCursor = pipeline.getCursor();
	
		return convertXLSX2XLS(fileName); 
	}
	
	private static String convertXls(IData pipeline, String fileName)throws Exception{
		
		IData data = IDataFactory.create();
		IDataCursor outCursor = pipeline.getCursor();
	
		return convertXLS2XLSX(fileName); 
	}
	
	
	private static String convertXLSX2XLS(String xlsxFilePath) {
		Map cellStyleMap = new HashMap();
		String xlsFilePath = null;
		XSSFWorkbook workbookIn = null;
		File xlsxFile = null;
		HSSFWorkbook workbookOut = null;
		OutputStream out = null;
		String XLS = ".xls";
		try {
			InputStream inputStream = new FileInputStream(xlsxFilePath);
			xlsFilePath = xlsxFilePath.substring(0, xlsxFilePath.lastIndexOf('.')) + XLS;
			workbookIn = new XSSFWorkbook(inputStream);
			xlsxFile = new File(xlsFilePath);
			if (xlsxFile.exists())
				xlsxFile.delete();
			workbookOut = new HSSFWorkbook();
			int sheetCnt = workbookIn.getNumberOfSheets();
	
			for (int i = 0; i < sheetCnt; i++) {
				Sheet sheetIn = workbookIn.getSheetAt(i);
				Sheet sheetOut = workbookOut.createSheet(sheetIn.getSheetName());
				Iterator<Row> rowIt = sheetIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = rowIt.next();
					Row rowOut = sheetOut.createRow(rowIn.getRowNum());
					copyRowPropertiesXlsx(rowOut, rowIn,cellStyleMap);
				}
			}
			out = new BufferedOutputStream((OutputStream)new FileOutputStream(xlsxFile));
			workbookOut.write(out);
		} catch (Exception ex) {
			System.err.println("Exception Occured inside transFormXLS2XLSX :: file Name :: " + xlsFilePath
					+ ":: reason ::" + ex.getMessage());
			ex.printStackTrace();
			xlsxFilePath = null;
		} finally {
			try {
				if (workbookOut != null)
					//(workbookOut).close();
				if (workbookIn != null)
					//workbookIn.close();
				if (out != null)
					out.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		return xlsFilePath;
	}
	
	private static void copyRowPropertiesXlsx(Row rowOut, Row rowIn, Map cellStyleMap) {
		rowOut.setRowNum(rowIn.getRowNum());
		rowOut.setHeight(rowIn.getHeight());
		rowOut.setHeightInPoints(rowIn.getHeightInPoints());
		rowOut.setZeroHeight(rowIn.getZeroHeight());
		Iterator<Cell> cellIt = rowIn.cellIterator();
		while (cellIt.hasNext()) {
			Cell cellIn = cellIt.next();
			Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());
			rowOut.getSheet().setColumnWidth(cellOut.getColumnIndex(),
					rowIn.getSheet().getColumnWidth(cellIn.getColumnIndex()));
			copyCellPropertiesXlsx(cellOut, cellIn, cellStyleMap);
		}
	
	}
	
	private static void copyCellPropertiesXlsx(Cell cellOut, Cell cellIn, Map cellStyleMap) {
	
		Workbook wbOut = cellOut.getSheet().getWorkbook();		
		switch (cellIn.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;
	
		case Cell.CELL_TYPE_BOOLEAN:
			cellOut.setCellValue(cellIn.getBooleanCellValue());
			break;
	
		case Cell.CELL_TYPE_ERROR:
			cellOut.setCellValue(cellIn.getErrorCellValue());
			break;
	
		case Cell.CELL_TYPE_FORMULA:
			cellOut.setCellFormula(cellIn.getCellFormula());
			break;
	
		case Cell.CELL_TYPE_NUMERIC:
			cellOut.setCellValue(cellIn.getNumericCellValue());
			break;
	
		case Cell.CELL_TYPE_STRING:
			cellOut.setCellValue(cellIn.getStringCellValue());
			break;
		}
		XSSFCellStyle styleIn = (XSSFCellStyle) cellIn.getCellStyle();
		HSSFCellStyle styleOut = null;
		if (cellStyleMap.get(styleIn.getIndex()) != null) {
			styleOut = (HSSFCellStyle) cellStyleMap.get(styleIn.getIndex());
		} else {
			styleOut = (HSSFCellStyle) wbOut.createCellStyle();
			styleOut.setAlignment(styleIn.getAlignment());
			DataFormat format = wbOut.createDataFormat();
			styleOut.setDataFormat(format.getFormat(styleIn.getDataFormatString()));
			XSSFColor forgroundColor = styleIn.getFillForegroundColorColor();
			if (forgroundColor != null) {				
				styleOut.setFillPattern(styleIn.getFillPattern());
			}
			styleOut.setFillPattern(styleIn.getFillPattern());
			styleOut.setBorderBottom(styleIn.getBorderBottom());
			styleOut.setBorderLeft(styleIn.getBorderLeft());
			styleOut.setBorderRight(styleIn.getBorderRight());
			styleOut.setBorderTop(styleIn.getBorderTop());			
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setHidden(styleIn.getHidden());
			styleOut.setIndention(styleIn.getIndention());
			styleOut.setLocked(styleIn.getLocked());
			styleOut.setRotation(styleIn.getRotation());			
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setWrapText(styleIn.getWrapText());
			cellOut.setCellComment(cellIn.getCellComment());
			cellStyleMap.put(styleIn.getIndex(), styleOut);
		}
		cellOut.setCellStyle(styleOut);
	}
	
	private static  String convertXLS2XLSX(String xlsFilePath) {
		Map cellStyleMap = new HashMap();
		String xlsxFilePath = null;
		HSSFWorkbook workbookIn = null;
		File xlsxFile = null;
		XSSFWorkbook workbookOut = null;
		OutputStream out = null;
		String XLSX = ".xlsx";
		try {
			InputStream inputStream = new FileInputStream(xlsFilePath);
			xlsxFilePath = xlsFilePath.substring(0, xlsFilePath.lastIndexOf('.')) + XLSX;
			workbookIn = new HSSFWorkbook(inputStream);
			xlsxFile = new File(xlsxFilePath);
			if (xlsxFile.exists())
				xlsxFile.delete();
			workbookOut = new XSSFWorkbook();
			int sheetCnt = workbookIn.getNumberOfSheets();
	
			for (int i = 0; i < sheetCnt; i++) {
				Sheet sheetIn = workbookIn.getSheetAt(i);
				Sheet sheetOut = workbookOut.createSheet(sheetIn.getSheetName());
				Iterator<Row> rowIt = sheetIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = rowIt.next();
					Row rowOut = sheetOut.createRow(rowIn.getRowNum());
					copyRowPropertiesXls(rowOut, rowIn,cellStyleMap);
				}
			}
			out = new BufferedOutputStream((OutputStream)new FileOutputStream(xlsxFile));
			workbookOut.write(out);
		} catch (Exception ex) {
			System.err.println("Exception Occured inside transFormXLS2XLSX :: file Name :: " + xlsFilePath
					+ ":: reason ::" + ex.getMessage());
			ex.printStackTrace();
			xlsxFilePath = null;
		} finally {
			try {
				if (workbookOut != null)
					//(workbookOut).close();
				if (workbookIn != null)
					//workbookIn.close();
				if (out != null)
					out.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		return xlsxFilePath;
	}
	
	private static void copyRowPropertiesXls(Row rowOut, Row rowIn, Map cellStyleMap) {
		rowOut.setRowNum(rowIn.getRowNum());
		rowOut.setHeight(rowIn.getHeight());
		rowOut.setHeightInPoints(rowIn.getHeightInPoints());
		rowOut.setZeroHeight(rowIn.getZeroHeight());
		Iterator<Cell> cellIt = rowIn.cellIterator();
		while (cellIt.hasNext()) {
			Cell cellIn = cellIt.next();
			Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());
			rowOut.getSheet().setColumnWidth(cellOut.getColumnIndex(),
					rowIn.getSheet().getColumnWidth(cellIn.getColumnIndex()));
			copyCellPropertiesXls(cellOut, cellIn, cellStyleMap);
		}
	
	}
	
	private static void copyCellPropertiesXls(Cell cellOut, Cell cellIn, Map cellStyleMap) {
	
		Workbook wbOut = cellOut.getSheet().getWorkbook();
		//HSSFPalette hssfPalette = cellIn.getSheet().getWorkbook().getCustomPalette();
		switch (cellIn.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;
	
		case Cell.CELL_TYPE_BOOLEAN:
			cellOut.setCellValue(cellIn.getBooleanCellValue());
			break;
	
		case Cell.CELL_TYPE_ERROR:
			cellOut.setCellValue(cellIn.getErrorCellValue());
			break;
	
		case Cell.CELL_TYPE_FORMULA:
			cellOut.setCellFormula(cellIn.getCellFormula());
			break;
	
		case Cell.CELL_TYPE_NUMERIC:
			cellOut.setCellValue(cellIn.getNumericCellValue());
			break;
	
		case Cell.CELL_TYPE_STRING:
			cellOut.setCellValue(cellIn.getStringCellValue());
			break;
		}
		HSSFCellStyle styleIn = (HSSFCellStyle) cellIn.getCellStyle();
		XSSFCellStyle styleOut = null;
		if (cellStyleMap.get(styleIn.getIndex()) != null) {
			styleOut = (XSSFCellStyle) cellStyleMap.get(styleIn.getIndex());
		} else {
			styleOut = (XSSFCellStyle) wbOut.createCellStyle();
			styleOut.setAlignment(styleIn.getAlignment());
			DataFormat format = wbOut.createDataFormat();
			styleOut.setDataFormat(format.getFormat(styleIn.getDataFormatString()));
			/*HSSFColor forgroundColor = styleIn.geth;
			if (forgroundColor != null) {
				short[] foregroundColorValues = forgroundColor.getTriplet();
				styleOut.setFillForegroundColor(new XSSFColor(new java.awt.Color(foregroundColorValues[0],
						foregroundColorValues[1], foregroundColorValues[2])));
				styleOut.setFillPattern(styleIn.getFillPattern());
			}*/
			styleOut.setFillPattern(styleIn.getFillPattern());
			styleOut.setBorderBottom(styleIn.getBorderBottom());
			styleOut.setBorderLeft(styleIn.getBorderLeft());
			styleOut.setBorderRight(styleIn.getBorderRight());
			styleOut.setBorderTop(styleIn.getBorderTop());
			/*HSSFColor bottom = hssfPalette.getColor(styleIn.getBottomBorderColor());
			if (bottom != null) {
				short[] bottomColorArray = bottom.getTriplet();
				styleOut.setBottomBorderColor(new XSSFColor(new java.awt.Color(bottomColorArray[0],
						bottomColorArray[1], bottomColorArray[2])));
			}
			HSSFColor top = hssfPalette.getColor(styleIn.getTopBorderColor());
			if (top != null) {
				short[] topColorArray = top.getTriplet();
				styleOut.setTopBorderColor(new XSSFColor(new java.awt.Color(topColorArray[0], topColorArray[1],
						topColorArray[2])));
			}
			HSSFColor left = hssfPalette.getColor(styleIn.getLeftBorderColor());
			if (left != null) {
				short[] leftColorArray = left.getTriplet();
				styleOut.setLeftBorderColor(new XSSFColor(new java.awt.Color(leftColorArray[0], leftColorArray[1],
						leftColorArray[2])));
			}
			HSSFColor right = hssfPalette.getColor(styleIn.getRightBorderColor());
			if (right != null) {
				short[] rightColorArray = right.getTriplet();
				styleOut.setRightBorderColor(new XSSFColor(new java.awt.Color(rightColorArray[0], rightColorArray[1],
						rightColorArray[2])));
			}*/
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setHidden(styleIn.getHidden());
			styleOut.setIndention(styleIn.getIndention());
			styleOut.setLocked(styleIn.getLocked());
			styleOut.setRotation(styleIn.getRotation());
			//styleOut.setShrinkToFit(styleIn.getShrinkToFit());
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setWrapText(styleIn.getWrapText());
			cellOut.setCellComment(cellIn.getCellComment());
			cellStyleMap.put(styleIn.getIndex(), styleOut);
		}
		cellOut.setCellStyle(styleOut);
	}
	
	
		
	// --- <<IS-END-SHARED>> ---
}

