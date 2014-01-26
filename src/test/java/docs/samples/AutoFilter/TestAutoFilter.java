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
package docs.samples.AutoFilter;

import org.openxls.ExtenXLS.AutoFilterHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * This Class Demonstrates the AutoFilter Functionality of ExtenXLS
 */
public class TestAutoFilter
{
	private static final Logger log = LoggerFactory.getLogger( TestAutoFilter.class );
	public WorkBookHandle book;
	public WorkSheetHandle sheet;
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Excel2007/";
	String od = System.getProperty( "user.dir" );

	/**
	 * Test ExtenXLS handling of AutoFilters
	 * ------------------------------------------------------------
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{
		TestAutoFilter test = new TestAutoFilter();
		test.TestReadExistingAFs();        // illustrate reading existing autofilters
		test.TestAlterExistingAF();        // illustrate modifying existing autofitlers
		test.TestAddAF();                // illustrate adding a variety of types of autofilters
		test.TestRemoveAF();            // illustrate removing autofilters
	}

	/**
	 * illustrate reading and interpreting of existing af's
	 */
	public void TestReadExistingAFs()
	{
		book = new WorkBookHandle( wd + "testAutoFilter.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			AutoFilterHandle[] afs = sheet.getAutoFilterHandles();    // get a handle to the existing AutoFilters
			// assert the autofilters in the file
			// 1st autofilter column 1 (B)
			log.info( "Column " + afs[0].getCol() );    //1
			log.info( "Condition Column 1 (B): " + afs[0].toString() );    //=MegaStore
			// 2nd Column C
			log.info( "Column " + afs[1].getCol() );    //2
			log.info( "Condition Column 2 (C): " + afs[1].toString() );    //=A-M
			// 3rd Column E
			log.info( "Column " + afs[2].getCol() );    //4 (E)
			log.info( "Condition Column 3 (E): " + afs[2].toString() );    //>2.98
			// 4th Column
			log.info( "Column " + afs[3].getCol() );
			log.info( "Condition Column 4: " + afs[3].toString() );    // Top 50 Items
			log.info( "Is it a Top 10 Filter? " + afs[3].isTop10() );
		}
		catch( Exception e )
		{
			log.error( e.toString() );
/**/
		}
	}

	/**
	 * Alter existing Af's
	 */
	public void TestAlterExistingAF()
	{
		// file has 7 columns, 2 of which have autofilters set
		book = new WorkBookHandle( wd + "testAutoFilter-test1.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			AutoFilterHandle[] afs = sheet.getAutoFilterHandles();    // get handle to existing AutoFilters
			// initial asserts
			// should be two afs on sheet 0 (AutoFilter Data):
			if( (afs.length != 2) ||
					(!afs[0].toString().equals( "=A-M" )) ||
					(!afs[1].toString().equals( ">2.98" )) )
			{
				log.error( "Incorrect Input File" );
			}
			// alter 1st condition
			afs[0].setTop10( 7, true, true ); // set to top 50%
			// alter second af condition
			// Test setting >, >=, <, <= to a value (custom)
			afs[1].setVal( new Double( 2.99 ), ">=" );
			afs[1].setVal2( new Double( 5.99 ), "<=", true );
			// write out changes
			sheet.evaluateAutoFilters();
			book.write( od + "TestAlterExistingAF-OUT.xls" );
		}
		catch( Exception e )
		{
			log.error( e.toString() );
		}
	}

	/**
	 * illustrate turning off all AutoFilters
	 */
	public void TestRemoveAF()
	{
		// file has 7 columns, 2 of which have autofilters set
		book = new WorkBookHandle( wd + "testAutoFilter-test1.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			AutoFilterHandle[] afs = sheet.getAutoFilterHandles();    // get handles to existing auto filters
			if( afs.length != 2 )
			{
				log.error( "Incorrect Input File" );                // two initially
			}
			sheet.removeAutoFilters();                                // remove auto filters
			book.write( od + "testRemoveAF-OUT.xls" );
		}
		catch( Exception e )
		{
			log.error( e.toString() );
		}
	}

	/**
	 * Test Adding AFs to a sheet that initially has none
	 * <p/>
	 * tests:  	simple top N,
	 * between 2 and 10,
	 * not equal and between
	 * bottom N
	 * blanks
	 * non-blanks
	 */
	public void TestAddAF()
	{
		// source is a workbook with two columns of simple numeric data, no autofilters
		addAF( "Top 5 Items", 0 );        // -OUT-1
		addAF( ">=2.0 AND <=10.0", 1 );    // -OUT-2
		addAF( "<>5.0~>=2.0 AND <=10.0", 0 );    // -0UT-3
		addAF( "Bottom 5 Items", 1 );        // -OUT-4
		addAF( "Top 10%", 1 );            // -OUT-5
		addAF( "Bottom 3%", 1 );            // -OUT-6
		addAF( "Blanks", 0 );                // -OUT-7
		addAF( "Non Blanks", 1 );            // -OUT-8
	}

	static int nAdditions = 0;

	private void addAF( String condition, int col )
	{
		// source is a workbook with two columns of simple numeric data, no autofilters
		book = new WorkBookHandle( wd + "testAutoFilter-test3.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );

			String[] split = StringTool.splitString( condition, "~" );    // for cases of more than 1 condition (AND or OR)
			for( int i = 0; i < split.length; i++ )
			{
				String cond = split[i];
				AutoFilterHandle af = sheet.addAutoFilter( col + i );        // add a new AutoFilter to column col
				if( cond.equals( "Top 5 Items" ) )    // set a "Top n" items
				{
					af.setTop10( 5, false, true );
				}
				else if( cond.equals( ">=2.0 AND <=10.0" ) )
				{        // set a "Betweens" condition
					af.setVal( Integer.valueOf( 2 ), ">=" );
					af.setVal2( Integer.valueOf( 10 ), "<=", true );
				}
				else if( cond.equals( "Bottom 5 Items" ) )        // set a "Bottom n" condition
				{
					af.setTop10( 5, false, false );    // show bottom 5 items
				}
				else if( cond.equals( "Top 10%" ) )                // set a "Top n %" condition
				{
					af.setTop10( 10, true, true );    // show top 10 percent
				}
				else if( cond.equals( "<>5.0" ) )                    // set a "Not equals" condition
				{
					af.setVal( Integer.valueOf( 5 ), "<>" );
				}
				else if( cond.equals( "Bottom 3%" ) )                // set a "Bottom n %" condition
				{
					af.setTop10( 3, true, false );    // show bottom 5 percent
				}
				else if( cond.equals( "Blanks" ) )                    // set filter on Blanks
				{
					af.setFilterBlanks();
				}
				else if( cond.equals( "Non Blanks" ) )                // set filter on Not Blanks
				{
					af.setFilterNonBlanks();
				}
			}

			// NOTE: must explicitly evalutateAutoFilters to activate changes
			sheet.evaluateAutoFilters();
			book.write( od + "TestAddAF-OUT-" + nAdditions + ".xls" );
		}
		catch( Exception e )
		{
			log.error( e.toString() );
		}
		nAdditions++;
	}
}


       