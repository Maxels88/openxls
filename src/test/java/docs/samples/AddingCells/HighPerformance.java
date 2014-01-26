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
package docs.samples.AddingCells;

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.DateConverter;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.openxls.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

/**
 * This example shows ExtenXLS with all the high performance settings
 * enabled, and optimized for adding
 * <p/>
 * <p/>
 * this test uses a bit of memory, be sure to set the Xmx setting on the Java command line
 * <p/>
 * ie: -Xms16M -Xmx1032M
 * <p/>
 * ------------------------------------------------------------
 */
public class HighPerformance
{
	private static final Logger log = LoggerFactory.getLogger( HighPerformance.class );
	public static DateFormat in_format = new SimpleDateFormat( "yyyy-MM-dd HH:mm:ss" );

	public static void main( String[] args )
	{

		String wd = System.getProperty( "user.dir" ) + "/docs/samples/AddingCells/";

		for( int z = 24000; z < 25000; z += 1000 )
		{
			WorkBookHandle bookHandle = new WorkBookHandle();

			// bookHandle.setDupeStringMode(WorkBookHandle.SHAREDUPES);
			bookHandle.setStringEncodingMode( WorkBookHandle.STRING_ENCODING_COMPRESSED ); // Change to UNICODE if you have eastern strings

			// System.setProperty("org.openxls.ExtenXLS.cacheCellHandles","true");

			log.info( "ExtenXLS Version: " + WorkBookHandle.getVersion() );
			BufferedReader fileReader = null;
			log.info( "Begin test." );

			try
			{
				int recordNum = 0;
				int sheetNum = 0;
				int row = 0;

				WorkSheetHandle sheetHandle;
				sheetHandle = addWorkSheet( bookHandle, sheetNum );
				//log.info(bookHandle.getStats());

				// See the valid list of format patterns in the
				// API docs for FormatHandle
				WorkSheetHandle formatSheet = bookHandle.getWorkSheet( "Sheet2" );
				formatSheet.setHidden( true );

				// These cells are added to create a format in the workbook...
				//
				CellHandle currencycell = formatSheet.add( new Double( 123.234 ), "A1" );
				currencycell.setFormatPattern( "($#,##0);($#,##0)" );

				CellHandle numericcell = formatSheet.add( new Double( 123.234 ), "A2" );
				numericcell.setFormatPattern( "0.00" );

				CellHandle datecell = formatSheet.add( new java.util.Date( System.currentTimeMillis() ), "A3" );

				int CurrencyFormat = formatSheet.getCell( "A1" ).getFormatId();
				int NumericFormat = formatSheet.getCell( "A2" ).getFormatId();
				int DateFormat = formatSheet.getCell( "A3" ).getFormatId();

				log.info( "Starting adding cells." );
				String line = "1234	3			4 ZZZZZZZZZZZZ	640	2	6			2005-01-28 00:00:00	8	9	7477747	QA01898388			2005-01-28 00:00:00	2005-01-28 00:00:00		0	0	0	0	0	0	1805000	1805000		1805000		2	0				NL	8	7	SOME ACCOUNT INC	293881	72	AKZO ZZZZZZZZZZZZ				28-Jan-05	783321	802778	99999	1294092184	640	1857520	A\r\n";

				/**
				 *  NOTE:This is a very important setting for performance
				 *  eliminates the lookup and return of a CellHandle for
				 *  each new Cell added
				 */
				sheetHandle.setFastCellAdds( true );

				for( int t = 0; t < z; t++ )
				{

					String[] tokens = StringTool.getTokensUsingDelim( line, "\t" );
					if( tokens != null )
					{
						for( int i = 0; i < tokens.length; i++ )
						{
							if( tokens[i] != null && recordNum != 0 )
							{
								switch( i )
								{

									case 10:
										;
									case 17:
										;
									case 18:
										;
									case 34:

										String dtsr = tokens[i];
										if( !dtsr.equals( "" ) )
										{
											java.sql.Date dtx = new java.sql.Date( in_format.parse( dtsr ).getTime() );

											// change the date pattern here...
											double dr = DateConverter.getXLSDateVal( dtx );
											sheetHandle.add( new Double( dr ), row, i );
										}
										break;
									case 16:
									case 20:
									case 21:
									case 22:
									case 23:
									case 24:
									case 25:
									case 26:
									case 27:
									case 29:
									case 31:
										if( !tokens[i].equals( "" ) )
										{
											// allows you to store numbers as numbers in XLS
											Object ob = null;
											try
											{ // try to get it as a number
												ob = new Double( tokens[i] + t );
											}
											catch( Exception ex )
											{
												try
												{ // try to get it as a number
													ob = Integer.valueOf( tokens[i] + t );
												}
												catch( Exception ext )
												{
													ob = tokens[i] + t;
												}
											}
											sheetHandle.add( ob, row, i );
										}
										break;

									default:
										if( !tokens[i].equals( "" ) )
										{
											// allows you to store numbers as numbers in XLS
											Object ob = tokens[i] + t;
											try
											{ // try to get it as a number
												ob = new Double( ob.toString() );
											}
											catch( Exception ex )
											{
												try
												{ // try to get it as a number
													ob = Integer.valueOf( ob.toString() );
												}
												catch( Exception ext )
												{
													;
												}
											}
											sheetHandle.add( ob, row, i );
										}
								}
							}
							else if( tokens[i] != null )
							{

								sheetHandle.add( tokens[i], row, i );
							}

						}
					}

					recordNum++;
					row++;
					if( recordNum % 1000 == 0 )
					{
						log.info( recordNum + " Rows Added" );
					}
					if( recordNum % 65000 == 0 )
					{
						row = 0;
						sheetNum++;
						sheetHandle = addWorkSheet( bookHandle, sheetNum );
						sheetHandle.setFastCellAdds( true );
					}
				}
			}
			catch( Exception e )
			{
				log.error("", e );
			}
			finally
			{
				File oFile = new File( wd + "fastAddOut_" + z + ".xls" );

				try
				{
					log.info( "Begin writing XLS file..." );
					// MUST use a buffered out for writing performance
					BufferedOutputStream bout = new BufferedOutputStream( new FileOutputStream( oFile ) );
					bookHandle.write( bout );
					bout.flush();
					bout.close();
					log.info( "Done writing XLS file." );
				}
				catch( Exception e1 )
				{
					log.error( "",e1 );
				}
				log.info( "Start reading XLS file." );
				WorkBookHandle wbh = new WorkBookHandle( wd + "fastAddOut_" + z + ".xls" );
				log.info( "Done reading XLS file." );
				wbh = null;
				bookHandle = null;
				System.gc();
			}
			log.info( "End test." );
		}
	}

	/**
	 * @param bookHandle
	 * @param k
	 * @param sheetHandle
	 * @return
	 * @throws WorkSheetNotFoundException
	 */
	private static WorkSheetHandle addWorkSheet( WorkBookHandle bookHandle, int k ) throws WorkSheetNotFoundException
	{
		WorkSheetHandle sheetHandle = null;
		try
		{
			sheetHandle = bookHandle.getWorkSheet( k );
		}
		catch( WorkSheetNotFoundException e1 )
		{
			sheetHandle = bookHandle.createWorkSheet( "sheet" + k );

		}
		return sheetHandle;
	}
}
