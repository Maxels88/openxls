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
package docs.samples.Compare2Spreadsheets;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Iterator;
import java.util.List;
import java.util.Vector;

/**
 * ------------------------------------------------------------
 * This test demonstrates how to read in 2 spreadsheets and compare the contents.
 * <p/>
 * This is just a beginning as there are further comparisons that can be performed which include:
 * <p/>
 * - compare the output from WorkBookHandle.getXML();
 * - compare missing and new sheets
 * - compare formatting and layout
 */
public class Compare2Spreadsheets
{
	private static final Logger log = LoggerFactory.getLogger( Compare2Spreadsheets.class );
	public static void main( String[] args )
	{
		Compare2Spreadsheets t1 = new Compare2Spreadsheets();

		WorkBookHandle bk1 = new WorkBookHandle( "test_diff.xlsx" );
		WorkBookHandle bk2 = new WorkBookHandle( "test_diff2.xlsx" );

		try
		{
			t1.testCompare2Spreadsheets( bk1, bk2 );
		}
		catch( Exception e )
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * a simple test of comparing 2 files
	 * <p/>
	 * ignores formatting, only compares values
	 *
	 * @param bk1 the original file
	 * @param bk2 the new file to compare with
	 */
	public void testCompare2Spreadsheets( WorkBookHandle bk1, WorkBookHandle bk2 ) throws Exception
	{

		WorkSheetHandle[] s1 = bk1.getWorkSheets();
		WorkSheetHandle[] s2 = bk2.getWorkSheets();

		if( s1.length != s2.length )
		{
			log.info( "These Workbooks do not have the same sheet count.  Original: " + s1.length + " vs. Compard:" + s2.length );
		}

		List matchedSheets = new Vector();

		// iterate the original and match sheets
		for( int x = 0; x < s1.length; x++ )
		{
			try
			{
				if( s1[x].getSheetName().equals( s2[x].getSheetName() ) )
				{
					matchedSheets.add( s1[x] );
				}
			}
			catch( Exception xe )
			{
				;
			}
		}

		Iterator it = matchedSheets.iterator();
		while( it.hasNext() )
		{
			WorkSheetHandle s1x = (WorkSheetHandle) it.next();
			WorkSheetHandle s2x = bk2.getWorkSheet( s1x.getSheetName() );

			// get all the cells, and compare them
			CellHandle[] c1x = s1x.getCells();
			CellHandle[] c2x = s2x.getCells();

			System.out.println( getCellText( c1x ) );

			System.out.println( getCellText( c2x ) );

		}

	}

	/**
	 * simple helper method to convert a cell array to comparison strings
	 *
	 * @param cx
	 * @return
	 */
	static String getCellText( CellHandle[] cx )
	{
		StringBuffer sbx = new StringBuffer();

		for( int t = 0; t < cx.length; t++ )
		{
			sbx.append( cx[t].toString() );
			sbx.append( "	" );
		}
		return sbx.toString();

	}
}