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
package docs.samples.ExcelWriterRefugee;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.RowHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;

public class WriteExcelFromResultset
{
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/ExcelWriterRefugee/";

	public static void main( String[] args )
	{
		new WriteExcelFromResultset( args );

	}

	public WriteExcelFromResultset( String[] args )
	{

		WorkBookHandle book = new WorkBookHandle( wd + "ExcelTemplate.xls" );
		book.setDupeStringMode( WorkBookHandle.ALLOWDUPES );

		HashMap globalVariableList = new HashMap();
		String COMPANY_NAME = "ACME Corporation";
		String COMPANY_CITY = "Boston";
		String COMPANY_STATE = "MA";

		try
		{
			ResultSet ReportRs;   /*this is the result set that holds the actual data that will be pushed to the template*/
			ReportRs = getReportResultSet();  /*the call to populate the result set*/

			setDataSource( book, globalVariableList, ReportRs, "SHEET1" );
			setCellDataSource( "Company Name:        " + COMPANY_NAME, "Company_Name", book, globalVariableList );
			setCellDataSource( "Company City:           " + COMPANY_CITY, "Company_City", book, globalVariableList );
			setCellDataSource( "Company State:         " + COMPANY_STATE, "Company_State", book, globalVariableList );

			// write new Excel file to filesystem or outputstream
			testWrite( book, wd + "output.xls" );
		}
		catch( Exception e )
		{
			e.printStackTrace();
		}
		finally
		{
			book.close();
		}
	}

	public void testWrite( WorkBookHandle b, String nm )
	{
		try
		{
			java.io.File f = new java.io.File( nm );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			b.write( bbout );
			bbout.flush();
			bbout.close();
			fos.flush();
			fos.close();
			System.gc();
		}
		catch( java.io.IOException e )
		{
			System.out.println( "IOException in Tester.  " + e );
		}
	}

	private ResultSet getReportResultSet()
	{
		Connection m_sqlCon = null;
		Statement m_stmt = null;
		ResultSet rs = null;

		String custid = "102";

		String strQuery = "SELECT customers.customName, " +
				"orders.orderID, orders.status, " +
				"items.itemcode, items.description, items.pricequote, items.quantity " +
				"FROM customers, orders, items " +
				"WHERE customers.custID=orders.custID AND items.orderID=orders.orderID ";
		//"AND customers.custID=" + custid;
		String strDriver = "org.gjt.mm.mysql.Driver";
		String strUrl = "jdbc:mysql://db.extentech.com/test";
		String strUser = "test";
		String strPassword = "test";

		try
		{
			java.sql.Driver d = (java.sql.Driver) Class.forName( strDriver ).newInstance();
			java.sql.DriverManager.registerDriver( d );
			m_sqlCon = java.sql.DriverManager.getConnection( strUrl, strUser, strPassword );

			m_stmt = m_sqlCon.createStatement();
			rs = m_stmt.executeQuery( strQuery );

		}
		catch( Exception e )
		{
			System.err.println( "ERROR: executing Resultset faild: " + e.toString() );
			e.printStackTrace();
		}
		return rs;
	}

	private void setDataSource( WorkBookHandle _book, HashMap variableList, ResultSet _rs, String whichSheet )
	{
		String cellData = "";
		HashMap fieldList = null;

		try
		{
			fieldList = new HashMap();

			// Loop through all worksheets
			CellHandle[] cells = _book.getCells();
			for( int t = 0; t < cells.length; t++ )
			{
				cellData = cells[t].getStringVal();
				// Look for cells that start with '%%='
				if( cellData.startsWith( "%%=" ) )
				{
					HashMap cellLocation = null;
					// Check to see if it is a variable and store location
					if( cellData.charAt( 3 ) == (char) '$' )
					{
						cellLocation = new HashMap();
						cellLocation.put( "format", cells[t].getFormatHandle() );
						cellLocation.put( "sheet", Integer.valueOf( cells[t].getSheetNum() ) );
						cellLocation.put( "row", Integer.valueOf( cells[t].getRowNum() ) );
						cellLocation.put( "col", Integer.valueOf( cells[t].getColNum() ) );
						variableList.put( cellData.substring( 4, cellData.length() ), cellLocation );
						// Otherwise store location of query field if it matches query string passed in
					}
					else if( cellData.substring( 3, cellData.indexOf( "." ) ).equalsIgnoreCase( whichSheet ) )
					{
						cellLocation = new HashMap();
						cellLocation.put( "format", cells[t].getFormatHandle() );
						cellLocation.put( "sheet", Integer.valueOf( cells[t].getSheetNum() ) );
						cellLocation.put( "row", Integer.valueOf( cells[t].getRowNum() ) );
						cellLocation.put( "col", Integer.valueOf( cells[t].getColNum() ) );
						fieldList.put( cellData.substring( cellData.indexOf( "." ) + 1, cellData.length() ), cellLocation );
					}
				}
			}
			int rowpos = 0, row = 0;
			// Loop through Resultset
			while( _rs.next() )
			{
				// Only ask the resultset for fields that were in the template
				Set set = fieldList.keySet();
				Iterator iter = set.iterator();
				WorkSheetHandle sheet = null;
				// loop through the columns
				while( iter.hasNext() )
				{
					// Find field in stored hashmap to determine spreadsheet location
					String key = (String) iter.next();
					HashMap _c = (HashMap) fieldList.get( key );
					if( sheet == null )
					{ // insert a row
						sheet = _book.getWorkSheet( ((Integer) _c.get( "sheet" )).intValue() );
						row = ((Integer) _c.get( "row" )).intValue();
						RowHandle copyRow = sheet.getRow( row );

						// only insert after 1st row
						if( rowpos > 0 )
						{
							sheet.insertRow( row + rowpos, copyRow, WorkSheetHandle.ROW_INSERT_ONCE, true );
						}

						rowpos++;
					}
					int col = ((Integer) _c.get( "col" )).intValue();
					// insert data into the cell
					if( rowpos == 1 )
					{
						//sheet.getCell((row+rowpos)-1,col).remove(true);
						sheet.add( _rs.getObject( key ), (row + rowpos) - 1, col );
					}
					else
					{
						CellHandle newcell = sheet.getCell( (row + rowpos) - 1, col );
						newcell.setVal( _rs.getObject( key ) );
					}
				}

				// alternate the background color of rows
				if( rowpos % 2 == 0 )
				{
					RowHandle rw = sheet.getRow( row + rowpos - 1 );
					rw.setBackgroundColor( _book.getColorTable()[FormatHandle.COLOR_AQUA] ); // you can use any color in the colortable
				}
			}
		}
		catch( Exception e )
		{
			e.printStackTrace();
		}
	}

	private void setCellDataSource( String cellText, String whichCell, WorkBookHandle _book, HashMap variableList )
	{
		HashMap _c = (HashMap) variableList.get( whichCell );
		int sheet = ((Integer) _c.get( "sheet" )).intValue();
		int row = ((Integer) _c.get( "row" )).intValue();
		int col = ((Integer) _c.get( "col" )).intValue();
		try
		{
			_book.getWorkSheet( sheet ).getCell( row, col ).setVal( cellText );
		}
		catch( Exception e )
		{
			e.printStackTrace();
		}
	}
}