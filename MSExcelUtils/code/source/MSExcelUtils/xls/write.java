package MSExcelUtils.xls;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
// --- <<IS-END-IMPORTS>> ---

public final class write

{
	// ---( internal utility methods )---

	final static write _instance = new write();

	static write _newInstance() { return new write(); }

	static write _cast(Object o) { return (write)o; }

	// ---( server methods )---




	public static final void writeToExcel (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(writeToExcel)>> ---
		// @sigtype java 3.5
		// [i] record:0:required inDoc
		// [i] - record:0:optional excelHeader
		// [i] -- field:0:required headerColumn1
		// [i] -- field:0:required headerColumn2
		// [i] -- field:0:required headerColumn3
		// [i] -- field:0:required headerColumn4
		// [i] -- field:0:required headerColumn5
		// [i] -- field:0:required headerColumn6
		// [i] -- field:0:required headerColumn7
		// [i] -- field:0:required headerColumn8
		// [i] -- field:0:required headerColumn9
		// [i] -- field:0:required headerColumn10
		// [i] - record:1:optional excelData
		// [i] -- field:0:required columnData1
		// [i] -- field:0:required columnData2
		// [i] -- field:0:required columnData3
		// [i] -- field:0:required columnData4
		// [i] -- field:0:required columnData5
		// [i] -- field:0:required columnData6
		// [i] -- field:0:required columnData7
		// [i] -- field:0:required columnData8
		// [i] -- field:0:required columnData9
		// [i] -- field:0:required columnData10
		// [i] - field:0:optional fileName
		// [i] - field:0:required option
		// [o] record:0:required outDoc
		// [o] - field:0:required file
		// [o] - object:0:required bytes
		// [o] - object:0:required stream
		// [o] - field:0:required errorMessage
		// [o] - field:0:required status
		// Initialize parameters for header data
		String headerColumn1 = null; 
		String headerColumn2 = null;
		String headerColumn3 = null;
		String headerColumn4 = null;
		String headerColumn5 = null;
		String headerColumn6 = null;
		String headerColumn7 = null;
		String headerColumn8 = null;
		String headerColumn9 = null;
		String headerColumn10 = null;
		 
		IData	excelHeader = null;
		IData[]	excelData = null;
		String	fileName = null;
		String	option = null;
				
				
		// Initialize parameters for excel data
		
		String	columnData1 = null;
		String	columnData2 = null;
		String	columnData3 = null;
		String	columnData4 = null;
		String	columnData5 = null;
		String	columnData6 = null;
		String	columnData7 = null;
		String	columnData8 = null;
		String	columnData9 = null;
		String	columnData10 = null;
		
		
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		IData inDoc = IDataUtil.getIData( pipelineCursor, "inDoc" );
		if(inDoc != null){
			IDataCursor inDocCursor = inDoc.getCursor();
		
			excelHeader = IDataUtil.getIData( inDocCursor, "excelHeader" );
			excelData = IDataUtil.getIDataArray( inDocCursor, "excelData" );
			fileName = IDataUtil.getString( inDocCursor, "fileName" );
			option = IDataUtil.getString( inDocCursor, "option" );
			inDocCursor.destroy();
		}	
		
		// excelHeader
					
					if ( excelHeader != null)
					{
						IDataCursor excelHeaderCursor = excelHeader.getCursor();
							headerColumn1 = IDataUtil.getString( excelHeaderCursor, "headerColumn1" );
							headerColumn2 = IDataUtil.getString( excelHeaderCursor, "headerColumn2" );
							headerColumn3 = IDataUtil.getString( excelHeaderCursor, "headerColumn3" );
							headerColumn4 = IDataUtil.getString( excelHeaderCursor, "headerColumn4" );
							headerColumn5 = IDataUtil.getString( excelHeaderCursor, "headerColumn5" );
							headerColumn6 = IDataUtil.getString( excelHeaderCursor, "headerColumn6" );
							headerColumn7 = IDataUtil.getString( excelHeaderCursor, "headerColumn7" );
							headerColumn8 = IDataUtil.getString( excelHeaderCursor, "headerColumn8" );
							headerColumn9 = IDataUtil.getString( excelHeaderCursor, "headerColumn9" );
							headerColumn10 = IDataUtil.getString( excelHeaderCursor, "headerColumn10" );
						excelHeaderCursor.destroy();
					
					}
				
				
					//XSSFWorkbook workbook = new XSSFWorkbook();
					HSSFWorkbook workbook = new HSSFWorkbook();
					Sheet sheet = workbook.createSheet("Sheet1");
					
					int sheetRowCount = 0;
					
					
					Object[][] excelDataHeading = {
							{headerColumn1,headerColumn2,headerColumn3,headerColumn4,headerColumn5,headerColumn6,headerColumn7,headerColumn8,headerColumn9,headerColumn10}
					};
					
					for (Object[] datatype : excelDataHeading){
						Row row = sheet.createRow(sheetRowCount++);
						int colNum = 0;
						
						for(Object field: datatype){
							Cell cell = row.createCell(colNum++);
							
							if(field instanceof String){
								cell.setCellValue((String) field);
							} else if (field instanceof Integer){
								cell.setCellValue((Integer) field);
							}
						}
							
					}
					
					// excelData
					//excelData = IDataUtil.getIDataArray( pipelineCursor, "excelData" );
					if ( excelData != null)
					{
						for ( int i = 0; i < excelData.length; i++ )
						{
							IDataCursor excelDataCursor = excelData[i].getCursor();
								columnData1 = IDataUtil.getString( excelDataCursor, "columnData1" );
								columnData2 = IDataUtil.getString( excelDataCursor, "columnData2" );
								columnData3 = IDataUtil.getString( excelDataCursor, "columnData3" );
								columnData4 = IDataUtil.getString( excelDataCursor, "columnData4" );
								columnData5 = IDataUtil.getString( excelDataCursor, "columnData5" );
								columnData6 = IDataUtil.getString( excelDataCursor, "columnData6" );
								columnData7 = IDataUtil.getString( excelDataCursor, "columnData7" );
								columnData8 = IDataUtil.getString( excelDataCursor, "columnData8" );
								columnData9 = IDataUtil.getString( excelDataCursor, "columnData9" );
								columnData10 = IDataUtil.getString( excelDataCursor, "columnData10" );
							excelDataCursor.destroy();
							
							Row row = sheet.createRow(sheetRowCount++);
							for(int k = 0; k < 10; k++){
								Cell cell = row.createCell(k);
								if (k==0)
									cell.setCellValue(columnData1);
								if (k==1)
									cell.setCellValue(columnData2);
								if (k==2)
									cell.setCellValue(columnData3);						
								if (k==3)
									cell.setCellValue(columnData4);
								if (k==4)
									cell.setCellValue(columnData5);
								if (k==5)
									cell.setCellValue(columnData6);
								if (k==6)
									cell.setCellValue(columnData7);
								if (k==7)
									cell.setCellValue(columnData8);						
								if (k==8)
									cell.setCellValue(columnData9);
								if (k==9)
									cell.setCellValue(columnData10);
							}
						}
					}
				
					
				
					try{
						if(option.equals(file)){						
							FileOutputStream outputStream = new FileOutputStream(fileName+".xls");
							workbook.write(outputStream);
							outputStream.close();
							workbook.close();
							
							IData outDocMSExcel = IDataFactory.create();
							IDataCursor outDocMSExcelCursor = outDocMSExcel.getCursor();
							IDataUtil.put( outDocMSExcelCursor, "status", "success" );
							IDataUtil.put( outDocMSExcelCursor, "file", fileName+".xls" );							
							outDocMSExcelCursor.destroy();
							IDataUtil.put(pipelineCursor, "outDoc", outDocMSExcel);		
							pipelineCursor.destroy();
							
							
						}else if(option.equals(bytes)){							
							ByteArrayOutputStream bytestream = new ByteArrayOutputStream();
							workbook.write(bytestream);
							bytestream.close();
							workbook.close();
							
							IData outDocMSExcel = IDataFactory.create();
							IDataCursor outDocMSExcelCursor = outDocMSExcel.getCursor();
							IDataUtil.put( outDocMSExcelCursor, "status", "success" );
							IDataUtil.put( outDocMSExcelCursor, "bytes", bytestream.toByteArray() );						
							outDocMSExcelCursor.destroy();
							IDataUtil.put(pipelineCursor, "outDoc", outDocMSExcel);		
							pipelineCursor.destroy();
							
							
						}else if(option.equals(stream)){							
							ByteArrayOutputStream stream = new ByteArrayOutputStream();
							workbook.write(stream);
							stream.close();
							workbook.close();
							
							IData outDocMSExcel = IDataFactory.create();
							IDataCursor outDocMSExcelCursor = outDocMSExcel.getCursor();
							IDataUtil.put( outDocMSExcelCursor, "status", "success" );
						    IDataUtil.put( outDocMSExcelCursor, "stream", stream );						
							outDocMSExcelCursor.destroy();
							IDataUtil.put(pipelineCursor, "outDoc", outDocMSExcel);		
							pipelineCursor.destroy();
							
							
						}
						
					} catch(Exception e){
						
						IData outDocMSExcel = IDataFactory.create();
						IDataCursor outDocMSExcelCursor = outDocMSExcel.getCursor();
						IDataUtil.put(outDocMSExcelCursor, "errorMessage", e.toString());
						IDataUtil.put( outDocMSExcelCursor, "status", "failed" );					
						outDocMSExcelCursor.destroy();
						IDataUtil.put(pipelineCursor, "outDoc", outDocMSExcel);		
						pipelineCursor.destroy();
						}
					
					pipelineCursor.destroy();
			
		
			
			
			
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	private static String file="file";
	private static String bytes="bytes";
	private static String stream="stream";
	// --- <<IS-END-SHARED>> ---
}

