package org.unisens.unisens2excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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
import org.unisens.MeasurementEntry;
import org.unisens.SignalEntry;
import org.unisens.TimedEntry;
import org.unisens.Unisens;
import org.unisens.UnisensFactory;
import org.unisens.UnisensFactoryBuilder;
import org.unisens.UnisensParseException;
import org.unisens.Value;
import org.unisens.ValuesEntry;

public class Unisens2Excel
{

    Unisens unisens;

    double baseSampleRate;

    int nColumns = 0;

    List<Entry> outputEntries;

    File excelOutputFile;

    Map<ValuesEntry, Value> currentValues;

    final static String[] dateTimeHeaders = { "Time rel", "Day rel", "Time rel", "Date abs", "Time abs"};

    final static String[] dateTimeUnits = { "[s]", "[d]", "[hh:mm:ss]", "[yyyy-mm-dd]", "[hh:mm:ss]"};

    final static String[] dateTimeDescriptions = { "Relative time from start of measurements in seconds", "Number of days from start of measurement", "Relative time from start if measurement", "Absolute date", "Absolute time"};

    public Unisens2Excel(String unisensPath, double baseSamplerate, String excelPathAndFile) throws UnisensParseException, FileNotFoundException
    {
        this(unisensPath, baseSamplerate, new File(excelPathAndFile));

    }

    public Unisens2Excel(String unisensPath, double baseSamplerate, File excelPathAndFile) throws UnisensParseException, FileNotFoundException
    {
        UnisensFactory unisensFactory = UnisensFactoryBuilder.createFactory();
        this.unisens = unisensFactory.createUnisens(unisensPath);
        this.excelOutputFile = excelPathAndFile;
        this.baseSampleRate = baseSamplerate;
        this.outputEntries = new ArrayList<Entry>();
        this.currentValues = new HashMap<ValuesEntry, Value>();

        List<Entry> allEntries = unisens.getEntries();
        Iterator<Entry> iterator = allEntries.iterator();
        Entry entry;
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
        }

        Collections.sort(outputEntries, new UnisensEntryComparer());

    }

    public void renderXLS() throws IOException
    {

        Calendar calender = Calendar.getInstance();

        if (this.nColumns > 0)
        {

            // keep 100 rows in memory, exceeding rows will be flushed to disk
            SXSSFWorkbook wb = new SXSSFWorkbook(100);

            CellStyle headline = wb.createCellStyle();
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

            // create sheet for results
            Sheet resultSheet = wb.createSheet("movisens DataAnalyzer Results");

            // create sheet for column descriptions
            Sheet parameterSheet = wb.createSheet("movisens DataAnalyzer Parameter Decriptions");
            parameterSheet.setColumnWidth(0, 7000);
            parameterSheet.setColumnWidth(1, 3000);
            parameterSheet.setColumnWidth(2, 30000);

            int rowCount = 0;
            Row row = resultSheet.createRow(0);

            for (int i = 0; i < dateTimeHeaders.length; i++)
            {
                Cell cell = row.createCell(i);
                cell.setCellValue(dateTimeHeaders[i] + " " + dateTimeUnits[i]);
                cell.setCellStyle(headline);
                resultSheet.setColumnWidth(i, 7000);

                Row pRow = parameterSheet.createRow(i);
                cell = pRow.createCell(0);
                cell.setCellValue(dateTimeHeaders[i]);
                cell = pRow.createCell(1);
                cell.setCellValue(dateTimeUnits[i]);
                cell = pRow.createCell(2);
                cell.setCellValue(dateTimeDescriptions[i]);
                rowCount++;
            }

            for (int i = 0; i < outputEntries.size(); i++)
            {
                MeasurementEntry measurementEntry = (MeasurementEntry) outputEntries.get(i);

                for (int j = 0; j < measurementEntry.getChannelCount(); j++)
                {
                    Cell cell = row.createCell(i + j + dateTimeHeaders.length);
                    cell.setCellValue(measurementEntry.getChannelNames()[j] + " [" + measurementEntry.getUnit() + "]");
                    cell.setCellStyle(headline);
                    resultSheet.setColumnWidth(i + j + dateTimeHeaders.length, 7000);

                    Row pRow = parameterSheet.createRow(i + j + dateTimeHeaders.length);
                    cell = pRow.createCell(0);
                    cell.setCellValue(measurementEntry.getChannelNames()[j]);
                    cell = pRow.createCell(1);
                    cell.setCellValue("[" + measurementEntry.getUnit() + "]");
                    cell = pRow.createCell(2);
                    cell.setCellValue(measurementEntry.getComment());
                    rowCount++;
                }
            }
            
            XSSFSheet sheet = wb.getXSSFWorkbook().getSheet(parameterSheet.getSheetName());
			updateDimensionRef(sheet, 3, rowCount);

            int rowNumber = 0;
            int maxColNumber = 0;
            boolean go = true;
            while (go)
            {
                row = resultSheet.createRow(rowNumber + 1);
                int cellnum = 0;
                Cell cell;

                // add timing cells
                int tRelSeconds = (int) (rowNumber / this.baseSampleRate);
                cell = row.createCell(cellnum);
                cell.setCellValue(tRelSeconds);
                cellnum++;

                int tRelDay = (int) Math.floor(tRelSeconds / 24.0 / 60 / 60);
                cell = row.createCell(cellnum);
                cell.setCellValue(tRelDay);
                cellnum++;

                calender.setTimeInMillis(0);
                calender.set(Calendar.MINUTE, 0);
                calender.set(Calendar.HOUR, 0);
                calender.set(Calendar.SECOND, 0);
                calender.add(Calendar.SECOND, tRelSeconds);
                cell = row.createCell(cellnum);
                cell.setCellValue(calender.getTime());
                cell.setCellStyle(hhmmss);
                cellnum++;

                Date tAbsDate = new Date(this.unisens.getTimestampStart().getTime() + tRelSeconds * 1000);
                cell = row.createCell(cellnum);
                cell.setCellValue(tAbsDate);
                cell.setCellStyle(yyyymmdd);
                cellnum++;

                cell = row.createCell(cellnum);
                cell.setCellValue(tAbsDate);
                cell.setCellStyle(hhmmss);
                cellnum++;

                // add data cells
                for (int i = 0; i < outputEntries.size(); i++)
                {
                    Entry entry = outputEntries.get(i);
                    if (entry instanceof SignalEntry)
                    {
                        SignalEntry signalEntry = (SignalEntry) entry;
                        double[][] dataArray = (double[][]) signalEntry.readScaled(1);

                        if (dataArray.length < 1)
                        {
                            resultSheet.removeRow(row);
                            go = false;
                            break;
                        }
                        else
                        {
                            double[] data = dataArray[0];
                            // write data in cells
                            for (int j = 0; j < signalEntry.getChannelCount(); j++)
                            {
                                cell = row.createCell(cellnum);
                                cellnum++;
                                cell.setCellValue(data[j]);
                            }
                        }
                    }
                    else if (entry instanceof ValuesEntry)
                    {
                        ValuesEntry valuesEntry = (ValuesEntry) entry;
                        Value value = this.currentValues.get(valuesEntry);
                        if ((value == null) || (value.getSampleStamp() < rowNumber))
                        {
                            Value[] values = valuesEntry.readScaled(1);
                            if (values.length > 0)
                            {
                                value = values[0];
                                this.currentValues.put(valuesEntry, value);
                            }
                            else
                            {
                                value = null;
                            }
                        }
                        if (value == null)
                        {
                            resultSheet.removeRow(row);
                            go = false;
                            break;
                        }
                        else
                        {
                            // write data in cells
                            if (value.getSampleStamp() == rowNumber)
                            {
                                for (int j = 0; j < valuesEntry.getChannelCount(); j++)
                                {
                                    cell = row.createCell(cellnum);
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

                    }
                }
                
                if(cellnum > maxColNumber) {
					maxColNumber = cellnum;
				}
                rowNumber++;
            }
            
            XSSFSheet sheet2 = wb.getXSSFWorkbook().getSheet(resultSheet.getSheetName());
			updateDimensionRef(sheet2, maxColNumber, rowNumber+1);

            FileOutputStream excelOutputStream = new FileOutputStream(excelOutputFile);
            wb.write(excelOutputStream);
            excelOutputStream.close();

            // dispose of temporary files backing this workbook on disk
            wb.dispose();
        }
        this.unisens.closeAll();
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
}
