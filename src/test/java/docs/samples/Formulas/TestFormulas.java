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
package docs.samples.Formulas;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.FormulaHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.toolkit.Logger;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * This Class Demonstrates the functionality of of ExtenXLS Formula manipulation.
 */

public class TestFormulas
{

	public static void main( String[] args )
	{
		testformula t = new testformula();
		t.testFormula();
		t.testHandlerFunctions();
		t.testMultiChange();
		t.changeSingleCellLoc();
	}
}

/**
 * Test the manipulation of Formulas within a worksheet.
 */
class testformula
{
	WorkBookHandle book = null;
	WorkSheetHandle sheet = null;
	String sheetname = "Sheet1";
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Formulas/";
	String finpath = wd + "testFormula.xls";

	/**
	 * thrash multiple changes to formula references and recalc
	 * ------------------------------------------------------------
	 */
	public void testMultiChange()
	{
		try
		{
			Logger.logInfo( "Testing multiple changes to formula references and recalc" );
			WorkBookHandle wbx = new WorkBookHandle();
			WorkSheetHandle sheet1 = wbx.getWorkSheet( 0 );
			sheet1.add( new Double( 100.123 ), "A1" );
			sheet1.add( new Double( 200.123 ), "A2" );
			CellHandle cx = sheet1.add( "=sum(A1*A2)", "A3" );
			Logger.logInfo( String.valueOf( cx ) );
			Logger.logInfo( "start setting 100k vals" );
			for( int t = 0; t < 100000; t++ )
			{
				sheet1.getCell( "A1" ).setVal( Math.random() * 10000 );
				sheet1.getCell( "A2" ).setVal( Math.random() * 10000 );
				Object calced = cx.getVal();
				Logger.logInfo( calced.toString() );
			}
			Logger.logInfo( "done setting 100k vals" );
			wbx.write( wd + "testFormulas_out.xls" );

		}
		catch( Exception ex )
		{
			Logger.logErr( "testFormulas.testMultiChange: " + ex.toString() );
		}
	}

	/**
	 * Demonstrates Dynamic Formula Calculation
	 */
	public void testCalculation()
	{
		try
		{
			this.openSheet( finpath, sheetname );
			// c4 + d4 = f4
			CellHandle mycell1 = sheet.getCell( "C4" );
			CellHandle mycell2 = sheet.getCell( "D4" );
			CellHandle myformulacell = sheet.getCell( "F4" );

			// output the calculated values
			FormulaHandle form = myformulacell.getFormulaHandle();
			System.out.println( form.calculate().toString() );

			// change the values then recalc			
			mycell1.setVal( 99 );
			mycell2.setVal( 420 );
			System.out.println( form.calculate().toString() );

			testWrite( "testCalculation_out.xls" );
		}
		catch( CellNotFoundException e )
		{
			System.out.println( "cell not found" + e );
		}
		catch( FormulaNotFoundException e )
		{
			System.out.println( "No formula to change" + e );
		}
		catch( Exception e )
		{
			Logger.logErr( "TestFormulas failed.", e );
		}
	}

	/**
	 * Move a Cell Reference within a Formula
	 */
	public void changeSingleCellLoc()
	{
		try
		{
			this.openSheet( finpath, sheetname );
			CellHandle mycell = sheet.getCell( "A10" );
			FormulaHandle form = mycell.getFormulaHandle();
			form.changeFormulaLocation( "A3", "G10" );
			testWrite( "testChangeSingleCellLoc_out.xls" );
		}
		catch( CellNotFoundException e )
		{
			System.out.println( "cell not found" + e );
		}
		catch( FormulaNotFoundException e )
		{
			System.out.println( "No formula to change" + e );
		}
	}

	/**
	 * Move a Cell range reference within a Formula
	 */
	public void testHandlerFunctions()
	{
		try
		{
			this.openSheet( finpath, sheetname );
			CellHandle mycell = sheet.getCell( "E8" );
			FormulaHandle myhandle = mycell.getFormulaHandle();
			boolean b = myhandle.changeFormulaLocation( "A1:B2", "D1:D28" );
			testWrite( "testHandlerFunctions_out.xls" );
		}
		catch( CellNotFoundException e )
		{
			System.out.println( "cell not found" + e );
		}
		catch( FormulaNotFoundException e )
		{
			System.out.println( "No formula to change" + e );
		}
	}

	/**
	 * Add a cell to a Cell range reference within a Formula
	 */
	public void testCellHandlerFunctions()
	{
		try
		{
			this.openSheet( finpath, sheetname );
			CellHandle mycell = sheet.getCell( "E8" );
			CellHandle secondcell = sheet.getCell( "D19" );
			FormulaHandle myhandle = mycell.getFormulaHandle();
			boolean b = myhandle.addCellToRange( "A1:B2", secondcell );
			testWrite( "testCellHandlerFunctions_out.xls" );
		}
		catch( CellNotFoundException e )
		{
			System.out.println( "cell not found" + e );
		}
		catch( FormulaNotFoundException e )
		{
			System.out.println( "No formula to change" + e );
		}
	}

	/**
	 * Run tests
	 */
	public void testFormula()
	{
		try
		{
			String finpath = wd + "testFormula.xls";
			String sheetname = "Sheet1";
			this.openSheet( finpath, sheetname );
			sheet.removeRow( 2, true );
			testWrite( "testFormula_out.xls" );
		}
		catch( Exception e )
		{
			System.out.println( "Exception in testFORMULA.testFormulaSeries(): " + e );
		}
	}

	WorkSheetHandle sht = null;

	/**
	 * Demonstrates calculation of formulas
	 * <p/>
	 * Jan 19, 2010
	 *
	 * @param fs
	 * @param sh
	 */
	void testFormulaCalc( String fs, String sh )
	{
		WorkBookHandle book = new WorkBookHandle( fs );
		sheetname = sh;
		try
		{
			sht = book.getWorkSheet( sheetname );
		}
		catch( Exception e )
		{
			Logger.logErr( "TestFormulas failed.", e );
		}

		FormulaHandle f = null;
		Double i = null;

		/************************************
		 * Formula Parse test
		 **************************************/
		if( sheetname.equalsIgnoreCase( "Sheet1" ) )
		{
			try
			{

				// one ref & ptgadd
				sht.add( null, "A1" );
				CellHandle c = sht.getCell( "A1" );
				c.setFormula( "b1+5" );
				f = c.getFormulaHandle();
				i = (Double) f.calculate();

				// two refs & ptgadd
				sht.add( null, "A2" );
				c = sht.getCell( "A2" );
				c.setFormula( "B1+ A1" );
				f = c.getFormulaHandle();
				i = (Double) f.calculate();

				// ptgsub
				f.setFormula( "B1 - 5" );
				i = (Double) f.calculate();

				// ptgmul
				f.setFormula( "D1 * F1" );
				i = (Double) f.calculate();

				// ptgdiv
				f.setFormula( "E1 / F1" );
				i = (Double) f.calculate();

				// ptgpower
				f.setFormula( "E1 ^ F1" );
				i = (Double) f.calculate();

				f.setFormula( "E1 > F1" );
				Boolean b = (Boolean) f.calculate();

				f.setFormula( "E1 >= F1" );
				b = (Boolean) f.calculate();

				f.setFormula( "E1 < F1" );
				b = (Boolean) f.calculate();

				f.setFormula( "E1 <= F1" );
				b = (Boolean) f.calculate();

				f.setFormula( "Pi()" );
				i = (Double) f.calculate();

				f.setFormula( "LOG(10,2)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "ROUND(32.443,1)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "MOD(45,6)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "DATE(1998,2,4)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "SUM(1998,2,4)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "IF(TRUE,1,0)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "ISERR(\"test\")" );
				b = (Boolean) f.calculate();
				System.out.println( b.toString() );

				// many operand ptgfuncvar
				f.setFormula( "SUM(12,3,2,4,5,1)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				// test with a sub-calc
				f.setFormula( "IF((1<2),1,0)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "IF((1<2),MOD(45,6),1)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "IF((1<2),if((true),8,1),1)" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

				f.setFormula( "IF((SUM(23,2,3,4)<12),if((true),8,1),DATE(1998,2,4))" );
				i = (Double) f.calculate();
				System.out.println( i.toString() );

			}
			catch( CellNotFoundException e )
			{
				Logger.logErr( "TestFormulas failed.", e );
			}
			catch( FunctionNotSupportedException e )
			{
				Logger.logErr( "TestFormulas failed.", e );
			}
			catch( Exception e )
			{
				Logger.logErr( "TestFormulas failed.", e );
			}
			testWrite( "testCalcFormulas_out.xls" );
		}
	}

	public void openSheet( String finp, String sheetnm )
	{
		book = new WorkBookHandle( finp );
		try
		{
			sheet = book.getWorkSheet( sheetnm );
		}
		catch( WorkSheetNotFoundException e )
		{
			System.out.println( "couldn't find worksheet" + e );
		}

	}

	public void testWrite( String fname )
	{
		try
		{
			java.io.File f = new java.io.File( wd + fname );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			book.write( bbout );
			bbout.flush();
			fos.close();
		}
		catch( java.io.IOException e )
		{
			Logger.logInfo( "IOException in Tester.  " + e );
		}
	}

}