package org.unisens.unisens2excel;

import java.io.File;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.unisens.Entry;
import org.unisens.Event;
import org.unisens.EventEntry;
import org.unisens.MeasurementEntry;
import org.unisens.SignalEntry;
import org.unisens.TimedEntry;
import org.unisens.Unisens;
import org.unisens.UnisensFactory;
import org.unisens.UnisensFactoryBuilder;
import org.unisens.UnisensParseException;
import org.unisens.Value;
import org.unisens.ValuesEntry;

import com.fasterxml.jackson.databind.ObjectMapper;

public class Unisens2Excel
{

    Unisens unisens;

    double baseSampleRate;

    File excelOutputFile;
    
    List<Entry> outputEntries;

    int nColumns = 0;

    
    Map<ValuesEntry, Value> queuedValues;
    Event queuedMarker;
    
    final static String OUTPUT_FILE_NAME = "Results.xlsx";
    
    final static String MARKER_ENTRY_ID = "marker.csv";

    final static String[] DATE_TIME_HEADERS = { "Time rel", "Day rel", "Time rel", "Date abs", "Time abs"};
    final static String[] DATE_TIME_UNITS = { "[s]", "[d]", "[hh:mm:ss]", "[yyyy-mm-dd]", "[hh:mm:ss]"};
    final static String[] DATE_TIME_DESCRIPTIONS = { "Relative time from start of measurements in seconds", "Number of days from start of measurement", "Relative time from start if measurement", "Absolute date", "Absolute time"};

    MarkerFormat markerFormat = MarkerFormat.SIMPLE;
    
    CellStyle headline;
    Sheet resultSheet;
    Row resultSheetRow;
    Sheet parameterSheet;
    
    
    public Unisens2Excel(String unisensPath, double baseSamplerate, String excelPathAndFile) throws UnisensParseException, FileNotFoundException
    {
        this(unisensPath, baseSamplerate, new File(excelPathAndFile));

    }
    
    public Unisens2Excel(String unisensPath, double baseSamplerate) throws UnisensParseException, FileNotFoundException
    {
        this(unisensPath, baseSamplerate, new File(unisensPath+"\\"+OUTPUT_FILE_NAME));
    }

    public Unisens2Excel(String unisensPath, double baseSamplerate, File excelPathAndFile) throws UnisensParseException, FileNotFoundException
    {
        UnisensFactory unisensFactory = UnisensFactoryBuilder.createFactory();
        this.unisens = unisensFactory.createUnisens(unisensPath);
        this.excelOutputFile = excelPathAndFile;
        this.baseSampleRate = baseSamplerate;
        this.outputEntries = new ArrayList<Entry>();
        this.queuedValues = new HashMap<ValuesEntry, Value>();


        List<Entry> allEntries = unisens.getEntries();
        Iterator<Entry> iterator = allEntries.iterator();
        Entry entry;
        Entry markerEntry = null;
        
        while (iterator.hasNext())
        {
            entry = (Entry) iterator.next();
            if ((entry instanceof org.unisens.SignalEntry) || (entry instanceof org.unisens.ValuesEntry))
            {
                if (((TimedEntry) entry).getSampleRate() == baseSampleRate)
                {
                    outputEntries.add(entry);
                    this.nColumns += ((MeasurementEntry) entry).getChannelCount();
                }
            }
            if ((entry instanceof org.unisens.EventEntry) && (entry.getId().equals(MARKER_ENTRY_ID)) && (((TimedEntry)entry).getSampleRate() > baseSamplerate) )
            {
            	markerEntry = entry;
            }
        }

        Collections.sort(outputEntries, new UnisensEntryComparer());
        if (markerEntry != null)
        {
        	this.outputEntries.add(0, markerEntry);
        }
    }

    public void setMarkerFormat(MarkerFormat markerFormat) {
    	this.markerFormat = markerFormat;
    }
    
    public void renderXLS() throws IOException
    {

        Calendar calender = Calendar.getInstance();

        if (this.nColumns > 0)
        {

            // keep 100 rows in memory, exceeding rows will be flushed to disk
            SXSSFWorkbook wb = new SXSSFWorkbook(100);

            headline = wb.createCellStyle();
            Font f = wb.createFont();
            f.setBoldweight(Font.BOLDWEIGHT_BOLD);
            headline.setFont(f);
            headline.setBorderBottom(CellStyle.BORDER_THIN);

            CellStyle number = wb.createCellStyle();
            DataFormat df = wb.createDataFormat();
            number.setDataFormat(df.getFormat("#,##0.0"));

            CreationHelper createHelper = wb.getCreationHelper();
            CellStyle hhmmss = wb.createCellStyle();
            hhmmss.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));

            createHelper = wb.getCreationHelper();
            CellStyle yyyymmdd = wb.createCellStyle();
            yyyymmdd.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));

            List<Marker> markerList = new ArrayList<Marker>();
            ObjectMapper objectMapper= new ObjectMapper();
            objectMapper.setTimeZone(TimeZone.getDefault());
            
            // create sheet for results
            resultSheet = wb.createSheet("movisens DataAnalyzer Results");

            // create sheet for column descriptions
            parameterSheet = wb.createSheet("movisens DataAnalyzer Parameter Decriptions");
            parameterSheet.setColumnWidth(0, 7000);
            parameterSheet.setColumnWidth(1, 7000);
            parameterSheet.setColumnWidth(2, 30000);

            int columnCount = 0;
            int parameterRowCount = 0;
            Row resultSheetRow = resultSheet.createRow(0);

            //create header for date and time
            for (int i = 0; i < DATE_TIME_HEADERS.length; i++)
            {
            	addColumnHeaderRow(resultSheetRow, columnCount, DATE_TIME_HEADERS[i] + " " + DATE_TIME_UNITS[i]);
                columnCount++;

                addParameter(parameterRowCount, DATE_TIME_HEADERS[i], DATE_TIME_UNITS[i], DATE_TIME_DESCRIPTIONS[i]);
                parameterRowCount++;
            }
            
            //create header for entries
            for (int i = 0; i < outputEntries.size(); i++)
            {
            	Entry entry = outputEntries.get(i);
            	
            	//for now we only have marker entry
            	if (entry instanceof EventEntry)
            	{
            		EventEntry eventEntry = (EventEntry) entry;
            		
            		if (markerFormat == MarkerFormat.JSON)
            		{
            			addColumnHeaderRow(resultSheetRow, columnCount, "Marker in JSON format");
            			addParameter(parameterRowCount, "Marker", "in JSON format", "Markers from marker.csv set by sensor and in UnisensViewer");
            		}
            		else
            		{
            			addColumnHeaderRow(resultSheetRow, columnCount, "Marker (Time; Type; Comment)");
            			addParameter(parameterRowCount, "Marker", "(Time; Type; Comment)", "Markers from marker.csv set by sensor and in UnisensViewer");
            		}
            		resultSheet.setColumnWidth(columnCount, 10000);

            		columnCount++;
                    parameterRowCount++;
            		
            	}
            	
            	if (entry instanceof MeasurementEntry)
            	{	
            		MeasurementEntry measurementEntry = (MeasurementEntry) entry;

	                for (int j = 0; j < measurementEntry.getChannelCount(); j++)
	                {
	                	addColumnHeaderRow(resultSheetRow, columnCount, measurementEntry.getChannelNames()[j] + " [" + measurementEntry.getUnit() + "]");
	                    columnCount++;
	
	            		addParameter(parameterRowCount, measurementEntry.getChannelNames()[j], "[" + measurementEntry.getUnit() + "]", measurementEntry.getComment());
	                    parameterRowCount++;
	                }
            	}
            }
            
            XSSFSheet sheet = wb.getXSSFWorkbook().getSheet(parameterSheet.getSheetName());
			updateDimensionRef(sheet, 3, parameterRowCount);

            int rowNumber = 0;
            int maxColNumber = 0;

            while (true)
            {
                resultSheetRow = resultSheet.createRow(rowNumber + 1);

                int nEntriesWithData = 0;
                int cellnum = 0;
                Cell cell;

                // add timing cells
                int tRelSeconds = (int) (rowNumber / this.baseSampleRate);
                cell = resultSheetRow.createCell(cellnum);
                cell.setCellValue(tRelSeconds);
                cellnum++;

                int tRelDay = (int) Math.floor(tRelSeconds / 24.0 / 60 / 60);
                cell = resultSheetRow.createCell(cellnum);
                cell.setCellValue(tRelDay);
                cellnum++;

                calender.setTimeInMillis(0);
                calender.set(Calendar.MINUTE, 0);
                calender.set(Calendar.HOUR, 0);
                calender.set(Calendar.SECOND, 0);
                calender.add(Calendar.SECOND, tRelSeconds);
                cell = resultSheetRow.createCell(cellnum);
                cell.setCellValue(calender.getTime());
                cell.setCellStyle(hhmmss);
                cellnum++;

                Date tAbsDate = new Date(this.unisens.getTimestampStart().getTime() + (long)tRelSeconds * 1000L);
                cell = resultSheetRow.createCell(cellnum);
                cell.setCellValue(tAbsDate);
                cell.setCellStyle(yyyymmdd);
                cellnum++;

                cell = resultSheetRow.createCell(cellnum);
                cell.setCellValue(tAbsDate);
                cell.setCellStyle(hhmmss);
                cellnum++;

                // add data cells
                for (int i = 0; i < outputEntries.size(); i++)
                {
                    Entry entry = outputEntries.get(i);

                    
                    if (entry instanceof EventEntry)
                    {
                    	EventEntry eventEntry = (EventEntry) entry;
                 	
                    	double tStartSec = (double)rowNumber / this.baseSampleRate;
                    	double tEndSec = tStartSec + 1.0 / this.baseSampleRate;
                    	
                    	Event marker = queuedMarker;
                    	
                    	while (true)
                    	{
                    		if (queuedMarker == null)
                    		{
                                List<Event> events = eventEntry.read(1);
                                if (events.size() > 0)
                                {
                                    marker = events.get(0);
                                } 
                                else
                                {
                                	marker = null;
                                }
                    		} 
                    		else
                    		{
                    			marker = queuedMarker;
                    		}
                    		
                    		if (marker != null)
                    		{	
                    			double tSec = marker.getSampleStamp() / eventEntry.getSampleRate();
                    			if ( (tSec >= tStartSec) && (tSec < tEndSec) )
                    			{
                           			markerList.add(new Marker(this.unisens.getTimestampStart(), marker, eventEntry.getSampleRate()));
                    				queuedMarker = null;
                    			}
                    			else
                    			{
                    				queuedMarker = marker;
                    				break;
                    			}
                    		}
                    		else
                    		{
                    			queuedMarker = null;
                    			break;
                    		}
                    	}
                    	
                		if (markerList.size() > 0)
                		{
                			cell = resultSheetRow.createCell(cellnum);
                    		if (markerFormat == MarkerFormat.JSON)
                    		{
                    			cell.setCellValue(objectMapper.writeValueAsString(markerList));
                    		}
                    		else
                    		{
                    			cell.setCellValue(markerList2String(markerList));
                    		}
                    		markerList.clear();
                		}
                		cellnum++;
                	
                    }
                    
                    if (entry instanceof SignalEntry)
                    {
                        SignalEntry signalEntry = (SignalEntry) entry;
                        double[][] dataArray = (double[][]) signalEntry.readScaled(1);

                        if (dataArray.length > 0)
                        {
                        	nEntriesWithData++;
                        	double[] data = dataArray[0];
                            // write data in cells
                            for (int j = 0; j < signalEntry.getChannelCount(); j++)
                            {
                                cell = resultSheetRow.createCell(cellnum);
                                cellnum++;
                                cell.setCellValue(data[j]);
                            }
                        }
                        else
                        {
                        	cellnum = cellnum + signalEntry.getChannelCount();
                        }
                    }
                    else if (entry instanceof ValuesEntry)
                    {
                        ValuesEntry valuesEntry = (ValuesEntry) entry;
                        Value value = this.queuedValues.get(valuesEntry);
                        if ((value == null) || (value.getSampleStamp() < rowNumber))
                        {
                            Value[] values = valuesEntry.readScaled(1);
                            if (values.length > 0)
                            {
                                value = values[0];
                                this.queuedValues.put(valuesEntry, value);
                            }
                            else
                            {
                                value = null;
                            }
                        }
                        if (value != null)
                        {
                        	nEntriesWithData++;
                            // write data into cells
                            if (value.getSampleStamp() == rowNumber)
                            {
                                for (int j = 0; j < valuesEntry.getChannelCount(); j++)
                                {
                                    cell = resultSheetRow.createCell(cellnum);
                                    cellnum++;
                                    cell.setCellValue(((double[]) value.getData())[j]);
                                }
                            }
                            else
                            {
                                // insert empty cells
                                cellnum += valuesEntry.getChannelCount();
                            }
                        }
                        else
                        {
                            // insert empty cells
                            cellnum += valuesEntry.getChannelCount();
                        }
                    }
                }
                
                if(cellnum > maxColNumber) {
					maxColNumber = cellnum;
				}
                
                if (nEntriesWithData==0)
                {
                	resultSheet.removeRow(resultSheetRow);
                	break;
                }

                rowNumber++;
                
            }
            
            XSSFSheet sheet2 = wb.getXSSFWorkbook().getSheet(resultSheet.getSheetName());
			updateDimensionRef(sheet2, maxColNumber, rowNumber+1);

			//delete Results.xsls if it already exists
			if (excelOutputFile.exists())
			{
				excelOutputFile.delete();
			}
				
            FileOutputStream excelOutputStream = new FileOutputStream(excelOutputFile);
            wb.write(excelOutputStream);
            excelOutputStream.close();

            // dispose of temporary files backing this workbook on disk
            wb.dispose();
        }
        this.unisens.closeAll();
    }
    
    
    void addColumnHeaderRow(Row row, int columnCount, String header)
    {
    	//result sheet
        Cell cell = row.createCell(columnCount);
        cell.setCellValue(header);
        cell.setCellStyle(headline);
        resultSheet.setColumnWidth(columnCount, 7000);
        columnCount++;
    }
    
    void addParameter(int rowCount, String name, String unit, String comment)
    {
        Row pRow = parameterSheet.createRow(rowCount);
        Cell cell = pRow.createCell(0);
        cell.setCellValue(name);
        cell = pRow.createCell(1);
        cell.setCellValue(unit);
        cell = pRow.createCell(2);
        cell.setCellValue(comment);
        rowCount++;
    }
    
    String markerList2String(List<Marker> markerList)
    {
    	String str="";
    	for (Marker marker : markerList) {
            str = str + marker.toString();
        }
		return str;
    }

    /**
	 * Takes in a sheet and 1-based base-10 column and row numbers and corrects
	 * the sheet dimensions Fixes Bug:
	 * https://issues.apache.org/bugzilla/show_bug.cgi?id=53611
	 */
	public static void updateDimensionRef(Sheet sheet, int colNumber,
			int rowNumber) {
		((XSSFSheet) sheet)
				.getCTWorksheet()
				.getDimension()
				.setRef("A1:"
						+ CellReference.convertNumToColString(colNumber - 1)
						+ rowNumber);
	}
	
	public static void main(String[] args) throws UnisensParseException, IOException 
	{
		
		batchProcess(args[0], 1.0/Double.parseDouble(args[1]));
	}
	
	public static void batchProcess(String path, double sampleRate) throws UnisensParseException, IOException 
	{
		Path basePath = Paths.get(path);
		ArrayList<Path> allUnisensPaths = getAllUnisensPaths(basePath, new ArrayList<Path>());
		for (Path unisensPath : allUnisensPaths)
		{
			System.out.print(unisensPath.toAbsolutePath().toString()+"...");
			Unisens2Excel unisens2Excel = new Unisens2Excel(unisensPath.toAbsolutePath().toString(), sampleRate);
			unisens2Excel.renderXLS();
			System.out.println("...done.");
		}
		
	}
	
	private static ArrayList<Path> getAllUnisensPaths(Path path, ArrayList<Path> list)
	{
	    try 
	    {
	    	DirectoryStream<Path> stream = Files.newDirectoryStream(path);
	        for (Path entry : stream) {
	            if (Files.isDirectory(entry)) {
	            	getAllUnisensPaths(entry, list);
	            }
	            else
	            {
	            	if (entry.toFile().getName().toLowerCase().equals("unisens.xml"))
	            	{
	            		list.add(entry.getParent());
	            	}
	            }
	        }
	        stream.close();
	    }
	    catch (IOException e)
	    {
	    	e.printStackTrace();
	    }

		return list; 
		
	}

	
	
}
