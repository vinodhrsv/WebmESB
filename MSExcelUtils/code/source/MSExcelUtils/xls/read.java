package MSExcelUtils.xls;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import com.wm.app.b2b.server.Server;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Properties;
import java.util.StringTokenizer;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import com.ibm.icu.text.SimpleDateFormat;
import com.ibm.icu.util.Calendar;
// --- <<IS-END-IMPORTS>> ---

public final class read

{
	// ---( internal utility methods )---

	final static read _instance = new read();

	static read _newInstance() { return new read(); }

	static read _cast(Object o) { return (read)o; }

	// ---( server methods )---




	public static final void readFromExcel (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(readFromExcel)>> ---
		// @sigtype java 3.5
		// [i] record:0:optional inDoc
		// [i] - field:0:required fileName
		// [i] - object:0:required fileStream
		// [i] - object:0:required byteArrayStream
		// [i] - field:0:required fileData
		// [o] record:0:optional WorkBook
		// [o] - record:1:required WorkSheet
		// [o] field:0:optional isValid
		IDataCursor pipelineCursor = pipeline.getCursor();
		String	fileName = null;
		BufferedInputStream	fileStream = null;
		InputStream    byteArrayStream = null;
		String	fileData = null; 
		 
		// data 
					IData	data = IDataUtil.getIData( pipelineCursor, "inDoc" );
					if ( data != null)
					{
						IDataCursor dataCursor = data.getCursor();
						fileName = IDataUtil.getString( dataCursor, "fileName" );
						fileStream = (BufferedInputStream)IDataUtil.get( dataCursor, "fileStream" );
						byteArrayStream = (InputStream)IDataUtil.get( dataCursor, "byteArrayStream" );
						fileData = IDataUtil.getString( dataCursor, "fileData" );
						dataCursor.destroy();
					}
		
					byte bin_array[] = (byte[])IDataUtil.get( pipelineCursor, "binData" );
					String xlsData = IDataUtil.getString( pipelineCursor, "xlsData" );	
					String s_encoding = IDataUtil.getString( pipelineCursor, "encoding" );
					String s_validate = IDataUtil.getString( pipelineCursor, "validate" );
					String s_returnErrors = IDataUtil.getString( pipelineCursor, "returnErrors" );
					
					pipelineCursor.destroy();
					
					if (xlsData != null)
						fileData = xlsData;
		
		IDataCursor pipelineCurosr_1 = pipeline.getCursor();
		IData[] work_sheet_list = null;
		IData[] row_list = null;
		boolean is_valid = false;
		 
		try{
			HSSFWorkbook wb = null;
			// Handle Inputs here
			if (fileName != null)
			{
			    if (fileName.length() > 0)
			    { 				
		    		wb = new HSSFWorkbook(new FileInputStream(fileName));
		            }
			}
			else if (fileStream != null)
			{		
				wb = new HSSFWorkbook(fileStream);
			}
			else if (fileData != null)
			{
				if (fileData.length() > 0)
				    wb = new HSSFWorkbook(new ByteArrayInputStream(fileData.getBytes()));
			}
			else if (bin_array != null)
			{		
				wb = new HSSFWorkbook(new ByteArrayInputStream(bin_array));
			}
			else if (byteArrayStream != null)
			{
				wb = new HSSFWorkbook(new BufferedInputStream(byteArrayStream));
			}			
						
			HSSFSheet sheet = null;
			HSSFRow row = null;
			HSSFCell cell = null;
			
			String cl="";
			double icl=0;
			boolean bcl=false;
			
			work_sheet_list = new IData[wb.getNumberOfSheets()];
			IDataCursor idc_sheet_node = null;
			for(int ws=0; ws<wb.getNumberOfSheets();ws++){
			sheet = wb.getSheetAt(ws);
			work_sheet_list[ws] = IDataFactory.create();
			idc_sheet_node = work_sheet_list[ws].getCursor();
			
			
			//Read Excel Data
			
			int row_cnt = sheet.getPhysicalNumberOfRows();
			row_list = new IData[row_cnt];
			
			for (int i = 0; i < row_cnt; i++){
				
				row = sheet.getRow(i);
				
				IDataCursor idc_row_node = null;
				row_list[i] = IDataFactory.create();
				idc_row_node = row_list[i].getCursor();
				
				if(row!=null){
					int phys_cell_num = row.getPhysicalNumberOfCells();
					int last_cell_num = row.getLastCellNum();
					int total_cell = phys_cell_num;
					 
					if (phys_cell_num < last_cell_num)
						total_cell = last_cell_num;
						for(int j=0; j<total_cell;j++){
							cell = row.getCell((short)j);
							if(cell!=null){
								switch (cell.getCellType()){
								case STRING:
									cl = cell.getStringCellValue();
									IDataUtil.put(idc_row_node, "C"+Integer.toString(j), cl);
									break;
									
								case NUMERIC:
									icl = cell.getNumericCellValue();
									if(isCellDateFormatted(cell)){
										Calendar cal = Calendar.getInstance();
										cal.setTime(getJavaDate(icl,false));
										String pattern = getCellDateFormat(cell);
										SimpleDateFormat df = new SimpleDateFormat(pattern);
										String dateStr = df.format(cal.getTime());
										
										IDataUtil.put(idc_row_node, "C"+Integer.toString(j), dateStr);
										
									}
									else
										IDataUtil.put(idc_row_node, "C"+Integer.toString(j), Double.toString(icl));
									break;
									
								case BOOLEAN:
									bcl = cell.getBooleanCellValue();
									if(bcl)
										IDataUtil.put(idc_row_node, "C"+Integer.toString(j), "true");
									else
										IDataUtil.put(idc_row_node, "C"+Integer.toString(j), "false");
									
									break;
									
								case FORMULA:
									
									try {
										icl = cell.getNumericCellValue();
										if(!Double.isNaN(icl))
											IDataUtil.put(idc_row_node, "C"+Integer.toString(j), Double.toString(icl));
										else
										{
											cl = cell.getStringCellValue();
											IDataUtil.put(idc_row_node, "C"+Integer.toString(j), cl);
										}
									} catch (Exception fe) {
										cl = cell.getCellFormula();
										IDataUtil.put(idc_row_node, "C"+Integer.toString(j), cl);
		
									}
										
									break;
									
									
								case BLANK:
									IDataUtil.put(idc_row_node, "C"+Integer.toString(j), "");
									break;
		
								case ERROR:
									IDataUtil.put(idc_row_node, "C"+Integer.toString(j), "");
									break;
								} //End Of Switch
							}
							else
							{
								IDataUtil.put(idc_row_node, "C"+Integer.toString(j), "");
							}
						} //end of for loop j
					
					
				}//end of row!=null if
			}//end of for loop i
			
			//Setup worksheet record
			
			IDataUtil.put(idc_sheet_node, "row", row_list);
			
			} //End of ws for loop
			is_valid = true;
			wb.close();
		}
		catch(Exception e){
			is_valid = false;
			e.printStackTrace();
			throw new ServiceException(e.getMessage());
		}
		pipelineCurosr_1.destroy();
		
		IDataCursor pipelineCursor_2 = pipeline.getCursor();
		
		IData recordMSExcel = IDataFactory.create();
		IDataCursor recordMSExcelCursor = recordMSExcel.getCursor();
		
		IDataUtil.put(recordMSExcelCursor, "Worksheet", work_sheet_list);
		recordMSExcelCursor.destroy();
		
		IDataUtil.put(pipelineCursor, "WorkBook", recordMSExcel);
		if(is_valid)
			IDataUtil.put(pipelineCursor, "isValid", "true");
		else 
			IDataUtil.put(pipelineCursor, "isValid", "false");
		
		pipelineCursor.destroy();
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	private static final long   DAY_MILLISECONDS  = 24 * 60 * 60 * 1000;
	
	
	
	
	////////////////////////////////////
	// Debug method     
	protected static Properties _props;
	
	private static boolean DEBUG = false;
	private static void printdbg(String msg)
	{
	       String str = getProperty("enableDebug", "false");
	       
	       Boolean dbg = Boolean.valueOf(str); 
	       if (dbg.booleanValue())
	              System.out.println(msg);
	}
	////////////////////////////////////
	
	  protected static Properties getProps()
	  {
	  if(_props == null)
	      try
	      {
	          File cfgfn = new File(Server.getResources().getPackageConfigDir("SimpleExcel"), "excel.cnf");
	          if(cfgfn.exists())
	          {
	              Properties tmp = new Properties();
	              FileInputStream fin = new FileInputStream(cfgfn);
	              tmp.load(fin);
	              fin.close();
	              _props = tmp;
	          }
	      }
	      catch(IOException io) { }
	  return _props;
	  }
	
	  public static String getProperty(String propertyName, String defValue)
	  {
	  Properties props = getProps();
	  String retval = null;
	  if(props != null)
	      retval = props.getProperty(propertyName, defValue);
	  return retval;
	  }
	
	public static String build_name(String str)
	{
	String name = "";
	StringTokenizer strtok = new StringTokenizer(str," ");
	String temp = "";
	int count = 0;
	while (strtok.hasMoreElements())
	{
	  temp = (String)strtok.nextElement();
	  if (count == 0)
	  name = temp;
	  else name += "_"+temp;
	
	  count++;
	}
	return name;
	} 
	/**
	 * Given a double, checks if it is a valid Excel date.
	 *
	 * @return true if valid
	 * @param  value the double value
	 */
	public static boolean isValidExcelDate(double value)
	{
	    return (value > -Double.MIN_VALUE);
	}
	
	///////////////////////////////////////////////////////////////
	// Method returns Java date pattern mapped from Excel date
	public static String getCellDateFormat(HSSFCell cell)
	{
	  String dt = "dd/MM/yyyy";
	HSSFCellStyle style = cell.getCellStyle();
	int i = style.getDataFormat();
	switch(i) 
	{
	  // Internal Date Formats as described on page 427 in Microsoft Excel Dev's Kit...
	  case 0x0e: //m/d/yyyy
	       dt = "MM/dd/yyyy";
	       break;
	  case 0x0f: //d-mmm
	       dt = "dd-MMM";
	       break;
	  case 0x10: //d-mmm-yy
	       dt = "dd/MM/yyyy";
	       break;
	  case 0x11: //mmm-yy
	       dt = "MMMM-yy";
	       break;
	  case 0x12: //h:mmAM/PM
	       dt = "hh:mm aa";
	       break;              
	  case 0x13: //h:mm:ssAM/PM
	       dt = "hh:mm:ss aa";
	       break;
	  case 0x14: //h:mm
	       dt = "hh:mm";
	       break;
	  case 0x15: //h:mm:ss
	       dt = "hh:mm:ss";
	       break;
	  case 0x16: //m/d/yyyy h:mm
	       dt = "MM/dd/yyyy hh:mm";
	       break;
	  case 0x2d: //mm:ss
	       dt = "mm:ss";
	       break;
	  case 0x2e: //[h]:mm:ss
	       dt = "hh:mm:ss";
	       break;
	  case 0x2f: //mm:ss.0
	      dt = "mm:ss.SSSS";
	  break;
	  
	  default:
	       dt = "dd-MMM-yy";
	  break;
	}  
	  
	  return dt;
	}
	
	//////////////////////////////////////////////////////////////////
	// method to determine if the cell is a date, versus a number...
	public static boolean isCellDateFormatted(HSSFCell cell) 
	{
	  boolean bDate = false; 
	
	  double d = cell.getNumericCellValue();
	  if ( isValidExcelDate(d) ) {
	HSSFCellStyle style = cell.getCellStyle();
	int i = style.getDataFormat();
	switch(i) { 
	  // Internal Date Formats as described on page 427 in Microsoft Excel Dev's Kit...
	  case 0x0e:
	  case 0x0f:
	  case 0x10:
	  case 0x11:
	  case 0x12: 
	  case 0x13:
	  case 0x14:
	  case 0x15:
	  case 0x16:
	  case 0x2d:
	  case 0x2e:
	  case 0x2f:
	   bDate = true;
	  break;
	
	  default:
	   bDate = false;
	  break;
	}
	  }
	  return bDate;
	}
	
	/**
	 * Given a Calendar, return the number of days since 1600/12/31.
	 *
	 * @return days number of days since 1600/12/31
	 * @param  cal the Calendar
	 * @exception IllegalArgumentException if date is invalid
	 */
	
	private static int absoluteDay(Calendar cal)
	{
	    return cal.get(Calendar.DAY_OF_YEAR)
	           + daysInPriorYears(cal.get(Calendar.YEAR));
	}
	
	/**
	 * Return the number of days in prior years since 1601
	 *
	 * @return    days  number of days in years prior to yr.
	 * @param     yr    a year (1600 < yr < 4000)
	 * @exception IllegalArgumentException if year is outside of range.
	 */
	
	private static int daysInPriorYears(int yr)
	{
	    if (yr < 1601)
	    {
	        throw new IllegalArgumentException(
	            "'year' must be 1601 or greater");
	    }
	    int y    = yr - 1601;
	    int days = 365 * y      // days in prior years
	               + y / 4      // plus julian leap days in prior years
	               - y / 100    // minus prior century years
	               + y / 400;   // plus years divisible by 400
	
	    return days;
	}
	
	/**
	 *  Given an Excel date with either 1900 or 1904 date windowing,
	 *  converts it to a java.util.Date.
	 *
	 *  @param date  The Excel date.
	 *  @param use1904windowing  true if date uses 1904 windowing,
	 *   or false if using 1900 date windowing.
	 *  @return Java representation of the date, or null if date is not a valid Excel date
	 */
	public static Date getJavaDate(double date, boolean use1904windowing) {
	    if (isValidExcelDate(date)) {
	        int startYear = 1900;
	        int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
	        int wholeDays = (int)Math.floor(date);
	        if (use1904windowing) {
	            startYear = 1904;
	            dayAdjust = 1; // 1904 date windowing uses 1/2/1904 as the first day
	        }
	        else if (wholeDays < 61) {
	            // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
	            // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
	            dayAdjust = 0;
	        }
	        GregorianCalendar calendar = new GregorianCalendar(startYear,0, wholeDays + dayAdjust);
	        int millisecondsInDay = (int)((date - Math.floor(date)) * (double) DAY_MILLISECONDS + 0.5);
	        calendar.set(GregorianCalendar.MILLISECOND, millisecondsInDay);
	        return calendar.getTime();
	    }
	    else {
	        return null;
	    }
	}
		
	// --- <<IS-END-SHARED>> ---
}

