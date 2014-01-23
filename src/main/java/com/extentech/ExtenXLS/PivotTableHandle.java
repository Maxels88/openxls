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
package com.extentech.ExtenXLS;

import com.extentech.formats.XLS.ColumnNotFoundException;
import com.extentech.formats.XLS.SxStreamID;
import com.extentech.formats.XLS.Sxview;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/** PivotTable Handle allows for manipulation of PivotTables within a WorkBook.

 @see    WorkBookHandle
 @see    WorkSheetHandle
 @see    CellHandle
 */
/**
 *
 */
/**
 *
 */

/**
 *
 */
public class PivotTableHandle
{
	private static final Logger log = LoggerFactory.getLogger( PivotTableHandle.class );

	public static int AUTO_FORMAT_Report1 = 1;
	public static int AUTO_FORMAT_Report2 = 2;
	public static int AUTO_FORMAT_Report3 = 3;
	public static int AUTO_FORMAT_Report4 = 4;
	public static int AUTO_FORMAT_Report5 = 5;
	public static int AUTO_FORMAT_Report6 = 6;
	public static int AUTO_FORMAT_Report7 = 7;
	public static int AUTO_FORMAT_Report8 = 8;
	public static int AUTO_FORMAT_Report9 = 9;
	public static int AUTO_FORMAT_Report10 = 10;
	public static int AUTO_FORMAT_Table1 = 11;
	public static int AUTO_FORMAT_Table2 = 12;
	public static int AUTO_FORMAT_Table3 = 13;
	public static int AUTO_FORMAT_Table4 = 14;
	public static int AUTO_FORMAT_Table5 = 15;
	public static int AUTO_FORMAT_Table6 = 16;
	public static int AUTO_FORMAT_Table7 = 17;
	public static int AUTO_FORMAT_Table8 = 18;
	public static int AUTO_FORMAT_Table9 = 19;
	public static int AUTO_FORMAT_Table10 = 20;
	public static int AUTO_FORMAT_Classic = 30;
	private Sxview pt;
	private WorkBookHandle book;
	private WorkSheetHandle worksheet;

	/**
	 * Create a new PivotTableHandle from an Excel PivotTable
	 *
	 * @param PivotTable - the PivotTable to create a handle for.
	 */
	public PivotTableHandle( Sxview f, WorkBookHandle bk )
	{
		pt = f;
		book = bk;
		//SxStreamID sxid= bk.getWorkBook().getPivotStrean(pt.getICache());
		//cellRange= sxid.getCellRange();
	}

	public WorkSheetHandle getWorkSheetHandle()
	{
		return worksheet;
	}

	/**
	 * @return Returns the cellRange.
	 */
	public CellRange getDataSourceRange()
	{
		SxStreamID sxid = book.getWorkBook().getPivotStream( pt.getICache() );
		return sxid.getCellRange();
	}

	/**
	 * Sets the Pivot Table Range to represent the Data to analyse
	 * <br>NOTE: any existing data will be replaced
	 *
	 * @param cellRange The cellRange to set.
	 */
	public void setSourceDataRange( CellRange cellRange )
	{
		SxStreamID sxid = book.getWorkBook().getPivotStream( pt.getICache() );
		sxid.setCellRange( cellRange );
		try
		{
			pt.setNPivotFields( (short) cellRange.getCols().length );
		}
		catch( ColumnNotFoundException e )
		{

		}

	}

	/**
	 * Sets the Pivot Table Range to represent the Data to analyse
	 * <br>NOTE: any existing data will be replaced
	 * <br>If the cell range does not contain sheet information, the sheet that the pivot table is located will be used
	 *
	 * @param cellRange
	 */
	public void setSourceDataRange( String range )
	{
		int[] rc = ExcelTools.getRangeCoords( range );
		if( range.indexOf( "!" ) == -1 )
		{
			range = this.getWorkSheetHandle() + "!" + range;
		}
		SxStreamID sxid = book.getWorkBook().getPivotStream( pt.getICache() );
		sxid.setCellRange( range );

		pt.setNPivotFields( (short) ((rc[3] - rc[1]) + 1) );
	}

	/**
	 * Sets the Pivot Table data source from a named range
	 *
	 * @param namedrange Named Range
	 */
	public void setSource( String namedrange )
	{
// TODO: finish; update DCONNAME					
	}

	/**
	 * get the Name of the PivotTable
	 *
	 * @return String - value of the PivotTable if stored as a String.
	 */
	public String getTableName()
	{
		return pt.getTableName();
	}

	/**
	 * set the Name of the PivotTable
	 *
	 * @param String - value of the PivotTable if stored as a String.
	 */
	public void setTableName( String tx )
	{
		pt.setTableName( tx );
	}

	/**
	 * returns the name of the data field
	 *
	 * @return
	 */
	public String getDataName()
	{
		return pt.getDataName();
	}

	/**
	 * sets the name of the data field for this pivot table
	 *
	 * @param name
	 */
	public void setDataName( String name )
	{
		pt.setDataName( name );
	}

	/**
	 * returns whether a given row is contained in this PivotTable
	 *
	 * @param int the row number
	 * @return boolean whether the row is in the table
	 */
	public boolean containsRow( int x )
	{
		if( (x <= pt.getRwLast()) && (x >= pt.getRwFirst()) )
		{
			return true;
		}
		return false;
	}

	/**
	 * return a more friendly
	 *
	 * @see java.lang.Object#toString()
	 */
	public String toString()
	{
		return this.getTableName();
	}

	/**
	 * returns whether a given col is contained in this PivotTable
	 *
	 * @param int the column number
	 * @return boolean whether the col is in the table
	 */
	public boolean containsCol( int x )
	{
		if( (x <= pt.getColLast()) && (x >= pt.getColFirst()) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Takes a string as a current PivotTable location, and changes that pointer
	 * in the PivotTable to the new string that is sent. This can take single
	 * cells "A5" and cell ranges, "A3:d4" Returns true if the cell range
	 * specified in PivotTableLoc exists & can be changed else false. This also
	 * cannot change a cell pointer to a cell range or vice versa.
	 *
	 * @param String
	 *            - range of Cells within PivotTable to modify
	 * @param String
	 *            - new range of Cells within PivotTable
	 *
	 *
	 *            public boolean changePivotTableLocation(String PivotTableLoc,
	 *            String newLoc) throws PivotTableNotFoundException{
	 *
	 *            Logger.logInfo("Changing: " + PivotTableLoc + " to: " +
	 *            newLoc);
	 *
	 *            return false;
	 *
	 *            }
	 */

	/**
	 * set whether to display row grand totals
	 *
	 * @param boolean whether to display row grand totals
	 */
	public void setRowsHaveGrandTotals( boolean b )
	{
		pt.setFRwGrand( b );
	}

	/**
	 * get whether table displays row grand totals
	 *
	 * @return boolean whether table displays row grand total
	 */
	public boolean getRowsHaveGrandTotals()
	{
		return pt.getFRwGrand();
	}

	/**
	 * set whether to display a column grand total
	 *
	 * @param boolean whether to display column grand total
	 */
	public void setColsHaveGrandTotals( boolean b )
	{
		pt.setColGrand( b );
	}

	/**
	 * get whether table displays a column grand total
	 *
	 * @return boolean whether table displays column grand total
	 */
	public boolean getColsHaveGrandTotals()
	{
		return pt.getFColGrand();
	}

	/**
	 * set the auto format for the Table
	 * <p/>
	 * <pre>
	 *         The valid formats are:
	 *
	 *             PivotTableHandle.AUTO_FORMAT_Report1
	 *             PivotTableHandle.AUTO_FORMAT_Report2
	 *             PivotTableHandle.AUTO_FORMAT_Report3
	 *             PivotTableHandle.AUTO_FORMAT_Report4
	 *             PivotTableHandle.AUTO_FORMAT_Report5
	 *             PivotTableHandle.AUTO_FORMAT_Report6
	 *             PivotTableHandle.AUTO_FORMAT_Report7
	 *             PivotTableHandle.AUTO_FORMAT_Report8
	 *             PivotTableHandle.AUTO_FORMAT_Report9
	 *             PivotTableHandle.AUTO_FORMAT_Report10
	 *             PivotTableHandle.AUTO_FORMAT_Table1
	 *             PivotTableHandle.AUTO_FORMAT_Table2
	 *             PivotTableHandle.AUTO_FORMAT_Table3
	 *             PivotTableHandle.AUTO_FORMAT_Table4
	 *             PivotTableHandle.AUTO_FORMAT_Table5
	 *             PivotTableHandle.AUTO_FORMAT_Table6
	 *             PivotTableHandle.AUTO_FORMAT_Table7
	 *             PivotTableHandle.AUTO_FORMAT_Table8
	 *             PivotTableHandle.AUTO_FORMAT_Table9
	 *             PivotTableHandle.AUTO_FORMAT_Table10
	 * </pre>
	 *
	 * @param int the auto format Id for the table
	 */
	public void setAutoFormatId( int b )
	{
		pt.setItblAutoFmt( (short) b );
	}

	/**
	 * get the auto format for the Table
	 *
	 * @return int the auto format Id for the table
	 */
	public int getAutoFormatId()
	{
		return pt.getItblAutoFmt();
	}

	/**
	 * set whether to auto format the Table
	 *
	 * @param boolean whether to auto format the table
	 */
	public void setUsesAutoFormat( boolean b )
	{
		pt.setFAutoFormat( b );
	}

	/**
	 * get whether table has auto format applied
	 *
	 * @param boolean whether table has auto format applied
	 */
	public boolean getUsesAutoFormat()
	{
		return pt.getFAutoFormat();
	}

	/**
	 * set Width/Height Autoformat is applied
	 *
	 * @param boolean whether to apply the Width/Height Autoformat
	 */
	public void setAutoWidthHeight( boolean b )
	{
		pt.setFWH( b );
	}

	/**
	 * get whether Width/Height Autoformat is applied
	 *
	 * @return boolean whether the Width/Height Autoformat is applied
	 */
	public boolean getAutoWidthHeight()
	{
		return pt.getFWH();
	}

	/**
	 * set whether Font Autoformat is applied
	 *
	 * @param boolean whether to apply the Font Autoformat
	 */
	public void setAutoFont( boolean b )
	{
		pt.setFFont( b );
	}

	/**
	 * get whether Font Autoformat is applied
	 *
	 * @return boolean whether the Font Autoformat is applied
	 */
	public boolean getAutoFont()
	{
		return pt.getFFont();
	}

	public void removeArtifacts()
	{
		int[] coords = {
				pt.getRwFirst() - 2, pt.getColFirst(), pt.getRwLast(), pt.getColLast()
		};
		try
		{
			CellRange newr = new CellRange( worksheet, coords, true );
			CellHandle[] ch = newr.getCells();
			for( CellHandle aCh : ch )
			{
				if( aCh != null )
				{
					aCh.remove( true );
				}
			}
		}
		catch( Exception e )
		{
			log.error( "could not remove artifacts in PivotTableHandle: " + e );
		}
	}

	/**
	 * set whether Alignment Autoformat is applied
	 *
	 * @param boolean whether to apply the Alignment Autoformat
	 */
	public void setAutoAlign( boolean b )
	{
		pt.setFAlign( b );
	}

	/**
	 * get whether Alignment Autoformat is applied
	 *
	 * @return boolean whether the Alignment Autoformat is applied
	 */
	public boolean getAutoAlign()
	{
		return pt.getFAlign();
	}

	/**
	 * set whether Border Autoformat is applied
	 *
	 * @param boolean whether to apply the Border Autoformat
	 */
	public void setAutoBorder( boolean b )
	{
		pt.setFBorder( b );
	}

	/**
	 * get whether Border Autoformat is applied
	 *
	 * @return boolean whether the Border Autoformat is applied
	 */
	public boolean getAutoBorder()
	{
		return pt.getFBorder();
	}

	/**
	 * set whether Pattern Autoformat is applied
	 *
	 * @param boolean whether to apply the Pattern Autoformat
	 */
	public void setAutoPattern( boolean b )
	{
		pt.setFPattern( b );
	}

	/**
	 * get whether Pattern Autoformat is applied
	 *
	 * @return boolean whether the Pattern Autoformat is applied
	 */
	public boolean getAutoPattern()
	{
		return pt.getFPattern();
	}

	/**
	 * set whether Number Autoformat is applied
	 *
	 * @param boolean whether to apply the Number Autoformat
	 */
	public void setAutoNumber( boolean b )
	{
		pt.setFNumber( b );
	}

	/**
	 * get whether Number Autoformat is applied
	 *
	 * @return boolean whether the Number Autoformat is applied
	 */
	public boolean getAutoNumber()
	{
		return pt.getFNumber();
	}

	/**
	 * set the first row in the PivotTable
	 */
	public void setRowFirst( int s )
	{
		// s--; // these are zero-based rows
		pt.setRwFirst( (short) s );
	}

	/**
	 * get the first row in the PivotTable
	 */
	public int getRowFirst()
	{
		return pt.getRwFirst() + 1;
	}

	/**
	 * set the last row in the PivotTable
	 */
	public void setRowLast( int s )
	{
		s--;
		pt.setRwLast( (short) s );
	}

	/**
	 * get the last row in the PivotTable
	 */
	public int getRowLast()
	{
		return pt.getRwLast() + 1;
	}

	/**
	 * set the first Column in the PivotTable
	 */
	public void setColFirst( int s )
	{
		pt.setColFirst( (short) s );
	}

	/**
	 * get the first Column in the PivotTable
	 */
	public int getColFirst()
	{
		return pt.getColFirst();
	}

	/**
	 * set the last Column in the PivotTable
	 */
	public void setColLast( int s )
	{
		pt.setColLast( (short) s );
	}

	/**
	 * get the last Column in the PivotTable
	 */
	public int getColLast()
	{
		return pt.getColLast();
	}

	/**
	 * set the first header row
	 */
	public void setRowFirstHead( int s )
	{
		s--;
		pt.setRwFirstHead( (short) s );
	}

	/**
	 * get the first header row
	 */
	public int getRowFirstHead()
	{
		return pt.getRwFirstHead() + 1;
	}

	/**
	 * set the first Row containing data
	 */
	public void setRowFirstData( int s )
	{
		s--; // zero-based rows
		pt.setRwFirstData( (short) s );
	}

	/**
	 * get the first Row containing data
	 */
	public int getRowFirstData()
	{
		return pt.getRwFirstData() + 1;
	}

	/**
	 * set the first Column containing data
	 */
	public void setColFirstData( int s )
	{
		pt.setColFirstData( (short) s );
	}

	/**
	 * get the first Column containing data
	 */
	public int getColFirstData()
	{
		return pt.getColFirstData();
	}

	/**
	 * returns the JSON representation of this PivotTable
	 *
	 * @return JSON Pivot Table representation
	 */
	public String getJSON()
	{
		// copy all methods to this sucker
		JSONObject thePivot = new JSONObject();

		try
		{
			thePivot.put( "title", this.getTableName() );
//			thePivot.put("cellrange", this.getCellRange().getRange());

		}
		catch( Exception e )
		{
			throw new WorkBookException( "PivotTableHandle.getJSON failed:" + e, WorkBookException.RUNTIME_ERROR );
		}
		return thePivot.toString();
	}

}