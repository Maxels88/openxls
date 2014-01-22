/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 * 
 * This file is part of OpenXLS.
 * 
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 * 
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package docs.samples.Charts;

/** Demonstration of Chart handling
 *

 This Class Demonstrates the basic functionality of of ExtenXLS Chart Handling.

 */

import com.extentech.ExtenXLS.ChartHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.ChartNotFoundException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.toolkit.Logger;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

public class TestCharts
{

	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Charts/";

	public static void main( String[] args )
	{
		TestCharts t = new TestCharts();
		t.testCharts();
	}

	/**
	 * Test the operation of the ChartHandle functions
	 */

	public ChartHandle ct = null;
	private String finpath = "", sheetname = "", foutpath = "";

	/**
	 * read in an XLS file, work with Sheets
	 */
	public void testCharts()
	{
		// these are the files to test
		String[] testfiles = {
				"testChart1.xls", "Sheet1",
		};
		for( int i = 0; i < testfiles.length; i++ )
		{
			if( false )
			{
				java.io.File logfile = new java.io.File( "extenxls" + i + ".log" );
				try
				{
					BufferedOutputStream sysout = new BufferedOutputStream( new FileOutputStream( logfile ) );
					System.setOut( new PrintStream( sysout ) );
				}
				catch( IOException e )
				{
					System.out.println( "IOE: " + e );
				}
			}
			finpath = testfiles[i];
			sheetname = testfiles[++i];
			System.out.println( "testReadWrite" + String.valueOf( i ) + " =====================> " + finpath + ":" + sheetname );
			doit( finpath, sheetname );
			testNewChart();
			createBubbleChart();
			System.out.println( "testReadWrite DONE" );
		}
	}

	/**
	 * create bubble chart from scratch
	 * <p/>
	 * ------------------------------------------------------------
	 */
	private void createBubbleChart()
	{

		WorkBookHandle book = new WorkBookHandle();
		book.removeAllWorkSheets();
		WorkSheetHandle sheet = book.createWorkSheet( "Bubble" );

		sheet.add( "Personal Care", "A1" );
		sheet.add( new Double( 0.20 ), "A2" );
		sheet.add( new Double( 1.40 ), "A3" );
		sheet.add( Integer.valueOf( 89000 ), "A4" );

		sheet.add( "Health Care", "B1" );
		sheet.add( new Double( 0.15 ), "B2" );
		sheet.add( new Double( 1.30 ), "B3" );
		sheet.add( Integer.valueOf( 60000 ), "B4" );

		sheet.add( "Dry Grocery", "C1" );
		sheet.add( new Double( 0.45 ), "C2" );
		sheet.add( new Double( .60 ), "C3" );
		sheet.add( Integer.valueOf( 52000 ), "C4" );

		sheet.add( "Frozen Products", "D1" );
		sheet.add( new Double( 0.32 ), "D2" );
		sheet.add( new Double( 0.65 ), "D3" );
		sheet.add( Integer.valueOf( 13000 ), "D4" );

		sheet.add( "Personal Care 2", "A6" );
		sheet.add( new Double( 0.20 ), "A7" );
		sheet.add( new Double( 1.40 ), "A8" );
		sheet.add( Integer.valueOf( 89000 ), "A9" );

		sheet.add( "Health Care 2", "B6" );
		sheet.add( new Double( 0.15 ), "B7" );
		sheet.add( new Double( 1.30 ), "B8" );
		sheet.add( Integer.valueOf( 60000 ), "B9" );

		sheet.add( "Dry Grocery 2", "C6" );
		sheet.add( new Double( 0.45 ), "C7" );
		sheet.add( new Double( .60 ), "C8" );
		sheet.add( Integer.valueOf( 52000 ), "C9" );

		sheet.add( "Frozen Products 2", "D6" );
		sheet.add( new Double( 0.32 ), "D7" );
		sheet.add( new Double( 0.65 ), "D8" );
		sheet.add( Integer.valueOf( 13000 ), "D9" );

		try
		{
			ChartHandle chart = book.createChart( "Bubble 1", sheet );
			chart.setChartType( ChartHandle.BUBBLECHART );

			chart.addSeriesRange( "Bubble!A1", "Bubble!A2", "Bubble!A3", "Bubble!A4" );
			chart.addSeriesRange( "Bubble!B1", "Bubble!B2", "Bubble!B3", "Bubble!B4" );
			chart.addSeriesRange( "Bubble!C1", "Bubble!C2", "Bubble!C3", "Bubble!C4" );
			chart.addSeriesRange( "Bubble!D1", "Bubble!D2", "Bubble!D3", "Bubble!D4" );
			ChartHandle chart2 = book.createChart( "Bubble 2", sheet );
			chart2.setChartType( ChartHandle.BUBBLECHART );
			chart2.addSeriesRange( "Bubble!A1", "Bubble!A2", "Bubble!A3", "Bubble!A4" );
			chart2.addSeriesRange( "Bubble!B1", "Bubble!B2", "Bubble!B3", "Bubble!B4" );
			chart2.addSeriesRange( "Bubble!C1", "Bubble!C2", "Bubble!C3", "Bubble!C4" );
			chart2.addSeriesRange( "Bubble!D1", "Bubble!D2", "Bubble!D3", "Bubble!D4" );

			this.testWrite( book, "NewBubbleChartOut.xls" );
		}
		catch( Exception e )
		{
			Logger.logErr( "CreateBubbleChart Failed " + e.toString() );
		}
	}

	/**
	 * Creates a new Chart in a new WorkBook.
	 * <p><p>
	 */
	public void testNewChart()
	{
		WorkBookHandle book = new WorkBookHandle();
		// attempt to create a new Chart
		WorkSheetHandle sheet = null;
		try
		{
			sheet = book.getWorkSheet( "Sheet1" );
		}
		catch( WorkSheetNotFoundException e )
		{
			Logger.logErr( e );
		}
		sheet.add( Integer.valueOf( 212 ), "A1" );
		sheet.add( Integer.valueOf( 54 ), "A2" );
		sheet.add( Integer.valueOf( 212 ), "B1" );
		sheet.add( Integer.valueOf( 54 ), "B2" );
		book.createChart( "NEWCHART", sheet );

		try
		{
			ChartHandle cht = book.getChart( "NEWCHART" );
			// Set ChartType
			/*
			int[] chartTypes= {
					ChartHandle.AREA,	// 0
					ChartHandle.BAR,	// 1
					ChartHandle.COL,	// 2
					ChartHandle.DOUGHNUT,	// 3
					ChartHandle.LINE,	// 4
					ChartHandle.PIE,	// 5
					ChartHandle.RADAR,	// 6
					ChartHandle.RADARAREA,	// 7
					ChartHandle.SCATTER,	// 8
					ChartHandle.SURFACE,	// 9
					ChartHandle.BUBBLE,		// 10
					ChartHandle.PYRAMID,	// 11
					ChartHandle.CONE,		// 12
					ChartHandle.CYLINDER	// 13
					ChartHandle.PYRAMINDBAR,
					ChartHandle.CONEBAR,
					ChartHandle.CYLINDERBAR
					};		
			*/

			cht.setChartType( ChartHandle.BUBBLECHART );
		}
		catch( Exception ex )
		{
			Logger.logErr( "Problem accessing new chart.", ex );
		}
		this.testWrite( book, "NewChartOut.xls" );

	}

	public void doit( String finp, String sheetnm )
	{
		WorkBookHandle book = new WorkBookHandle( finp );
		WorkSheetHandle sheet = null;
		try
		{
			sheet = book.getWorkSheet( sheetnm );
			WorkSheetHandle[] handles = book.getWorkSheets();
			ChartHandle[] charts = book.getCharts();
			for( int i = 0; i < charts.length; i++ )
			{
				System.out.println( "Found Chart: " + charts[i] );
			}
			sheet.add( Integer.valueOf( 4 ), "A4" );
			sheet.add( Integer.valueOf( 5 ), "A5" );
			sheet.add( Integer.valueOf( 14 ), "B4" );
			sheet.add( Integer.valueOf( 15 ), "B5" );
			sheet.add( Integer.valueOf( 24 ), "C4" );
			sheet.add( Integer.valueOf( 25 ), "C5" );

			try
			{
				ct = book.getChart( "Test Chart" );
			}
			catch( ChartNotFoundException e )
			{
				System.out.println( e );
			}
			sheet.add( new Float( 103.256 ), "F23" );
			sheet.add( Integer.valueOf( 125 ), "G23" );

			if( ct.changeSeriesRange( "Sheet1!C23:E23", "Sheet1!C23:G23" ) )
			{
				;
			}
			System.out.println( "Successfully Changed Series Range!" );

			sheet.add( "D", "F24" );
			sheet.add( "E", "G24" );
			if( ct.changeCategoryRange( "Sheet1!C24:E24", "Sheet1!C24:G24" ) )
			{
				;
			}
			System.out.println( "Successfully Changed Categories Range!" );

			if( ct.changeTextValue( "Category X", "Widget Class" ) )
			{
				;
			}
			System.out.println( "Successfully Changed Categories Label!" );

			if( ct.changeTextValue( "Value Y", "Sales Totals" ) )
			{
				;
			}
			System.out.println( "Successfully Changed Values Label!" );

			ct.setTitle( "New Chart Title!" );
			System.out.println( "Chart Name: " + ct );

			foutpath = finp + "output.xls";

			// Copy the worksheet
			book.copyWorkSheet( "Sheet2", "SheetNEW" );

			// Copy the Chart
			try
			{
				book.copyChartToSheet( "New Chart Title!", "SheetNEW" );
			}
			catch( Exception e )
			{
				System.out.println( e );
			}

			testWrite( book, foutpath );
			WorkBookHandle b2 = new WorkBookHandle( foutpath );
		}
		catch( WorkSheetNotFoundException e )
		{
			System.out.println( e );
		}
	}

	public void testWrite( WorkBookHandle b, String fout )
	{
		try
		{
			java.io.File f = new java.io.File( wd + fout );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			b.write( bbout );
			bbout.flush();
			fos.close();
		}
		catch( java.io.IOException e )
		{
			Logger.logInfo( "IOException in Tester.  " + e );
		}
	}

}
