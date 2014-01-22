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

import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONArray;
import org.json.JSONObject;

import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELLS;

/**
 * The NameHandle provides access to a Named Range and its Cells.<br>
 * <br>
 * Use the NameHandle to work with individual Named Ranges in an XLS file.<br>
 * <br>  <br>
 * With a NameHandle you can:
 * <br><br>
 * <blockquote>
 * get a handle to the Cells in a Name<br>
 * set the default formatting for a Name
 * </blockquote>
 * <br>
 * <br>
 *
 * @see WorkBookHandle
 * @see WorkSheetHandle
 * @see FormulaHandle
 */
public class NameHandle
{

	private Name myName;
	private WorkBook mybook;
	private int DEBUGLEVEL = -1;
	private boolean createblanks = false;
	private CellRange initialRange = null;

	/**
	 * Creates a new Named Range from a CellRange
	 *
	 * @param Name of the Range, the CellRange referenced by the Name
	 */
	public NameHandle( String namestr, CellRange cr )
	{
		mybook = cr.getWorkBook();
		initialRange = cr; // cache
		myName = new Name( mybook.getWorkBook(), namestr );
		this.setName( namestr );
		this.setLocation( cr.toString() );
	}

	/**
	 * Create a NameHandle from an internal Name record
	 *
	 * @param c
	 * @param myb
	 */
	protected NameHandle( Name c, WorkBookHandle myb )
	{
		myName = c;
		mybook = myb;
	}

	/**
	 * Returns a handle to the object (either workbook or sheet) that is scoped
	 * to the name record
	 * <p/>
	 * Default scope is a WorkBookHandle, else the WorkSheetHandle is returned.
	 *
	 * @return the scope of the name
	 */
	public Handle getScope() throws WorkSheetNotFoundException
	{
		int itab = myName.getItab();
		if( itab == 0 )
		{
			return mybook;
		}
		WorkSheetHandle sheet = mybook.getWorkSheet( itab - 1 );
		return sheet;
	}

	/**
	 * Set the scope of this name to that of the handle passed in.
	 * <p/>
	 * This can either be a WorkbookHandle or a WorksheetHandle
	 * <p/>
	 * note: this will only be functional for the workbook that the name is contained in,
	 * you cannot change the scope to a different Document with this method.
	 *
	 * @param scope Workbookhandle or WorksheetHandle
	 */
	public void setScope( Handle scope )
	{
		int newitab = 0;
		if( scope instanceof WorkSheetHandle )
		{
			newitab = (((WorkSheetHandle) scope).getSheetNum() + 1);
		}
		try
		{
			myName.setNewScope( newitab );
		}
		catch( WorkSheetNotFoundException e )
		{
			// this really shouldnt happen unless you are passing a scope in from a different workbook
			Logger.logErr( "ERROR: setting new scope on name: " + e );
		}
	}

	/**
	 * Create a new named range in the workbook
	 * note that this does not set the actual range, just the workbook and name,
	 * follow up with setLocation
	 *
	 * @param namestr
	 * @param myb
	 * @deprecated
	 */
	public NameHandle( String namestr, WorkBookHandle myb )
	{
		mybook = myb;
		myName = new Name( mybook.getWorkBook(), namestr );
		this.setName( namestr );
	}

	/**
	 * Create a new named range in the workbook
	 *
	 * @param name     name that should be used to reference this named range
	 * @param location rangeDef Range of the cells for this named range, in excel syntax including sheet name, ie "Sheet1!A1:D1"
	 * @param book     WorkBookHandle to insert this named range into
	 */
	public NameHandle( String name, String location, WorkBookHandle book )
	{
		mybook = book;
		myName = new Name( mybook.getWorkBook(), name );
		this.setName( name );
		this.setLocation( location );
		mybook.getWorkBook().associateDereferencedNames( myName );
	}

	/**
	 * Return an XML representation of this name record
	 */
	public String getXML()
	{
		StringBuffer retXML = new StringBuffer();
		String nmv = getExpressionString();
		if( nmv.indexOf( "=!" ) > -1 )
		{ // add the sheetname
			Boundsheet bs = myName.getSheet();
			if( bs != null )
			{
				nmv = bs.getSheetName() + "!" + nmv;
			}
			else
			{ // TODO: why no sheet defined for name???
				nmv = StringTool.replaceChars( "!", nmv, "" );
			}
		}
		// it's possible that name expression can have quotes and other non-compliant characters
		nmv = com.extentech.toolkit.StringTool.convertXMLChars( nmv );

		String nmx = getName();
		nmx = com.extentech.toolkit.StringTool.convertXMLChars( nmx );

		if( nmx.length() == 1 )
		{ // deal with snafu character names
			if( Character.isLetterOrDigit( nmx.charAt( 0 ) ) )
			{
				Logger.logInfo( "NameHandle getting XML for name: " + nmx );
			}
			else
			{
				nmx = "#NAME";
			}
		}

		if( nmv.startsWith( "=" ) )
		{
			nmv = nmv.substring( 1 );
		}
		retXML.append( "		<NamedRange Name=\"" );
		retXML.append( nmx );
		retXML.append( "\" RefersTo=\"" );
		retXML.append( nmv );
		if( myName.getItab() != 0 )
		{ // if not workbook-scoped, add sheet scope for later retrieval
			retXML.append( "\" Scope=\"" );
			retXML.append( myName.getItab() );
		}
		retXML.append( "\"/>" );
		return retXML.toString();
	}

	/**
	 * @return String XML rep of all the cells referenced by this range
	 */
	public String getExpandedXML()
	{
		String nmv = getExpressionString();
		if( nmv.indexOf( "=!" ) > -1 )
		{ // add the sheetname
			Boundsheet bs = myName.getSheet();
			if( bs != null )
			{
				nmv = bs.getSheetName() + "!" + nmv;
			}
			else
			{ // TODO: why no sheet defined for name???
				nmv = StringTool.replaceChars( "!", nmv, "" );
			}
		}
		String nmx = getName();
		nmx = com.extentech.toolkit.StringTool.convertXMLChars( nmx );

		if( nmx.length() == 1 )
		{ // deal with snafu character names
			if( Character.isLetterOrDigit( nmx.charAt( 0 ) ) )
			{
				Logger.logInfo( "NameHandle getting XML for name: " + nmx );
			}
			else
			{
				nmx = "#NAME";
			}
		}

		if( nmv.startsWith( "=" ) )
		{
			nmv = nmv.substring( 1 );
		}
		StringBuffer retXML = new StringBuffer();
		retXML.append( "\t<NamedRange Name=\"" );
		retXML.append( nmx );
		retXML.append( "\" RefersTo=\"" );
		retXML.append( nmv );
		retXML.append( "\">\n" );
		try
		{
			CellHandle[] cells = getCells();

			for( int i = 0; i < cells.length; i++ )
			{
				retXML.append( "\t\t" + cells[i].getXML() + "\n" );
			}
		}
		catch( CellNotFoundException ex )
		{
			Logger.logErr( "NameHandle.getExpandedXML failed: ", ex );
		}
		retXML.append( "</NamedRange>\n" );

		return retXML.toString();
	}

	/**
	 * sets the default format id for the Name's Cells
	 *
	 * @param int Format Id for all Cells in Name
	 */
	public void setFormatId( int i )
	{
		// myName.setXFRecord(i);
		// TODO: why doesn't this work with myName?
		try
		{
			CellHandle[] cs = this.getCells();
			for( int t = 0; t < cs.length; t++ )
			{
				cs[t].setFormatId( i );
			}
		}
		catch( CellNotFoundException ex )
		{
			Logger.logErr( "NameHandle.setFormatId failed: ", ex );
		}
	}

	/**
	 * deletes a row of cells in this named range, shifts subsequent
	 * rows up.
	 *
	 * @param the column to use as unique index
	 */
	public void deleteRow( int idxcol ) throws Exception
	{
		CellRange[] rngs = getCellRanges();
		if( rngs.length > 1 )
		{
			throw new WorkBookException( "NamedRange.updateRow Object array failed: too many CellRanges.",
			                             WorkBookException.RUNTIME_ERROR );
		}
		if( rngs.length == 0 )
		{
			throw new WorkBookException( "NamedRange.updateRow Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR );
		}

		//
		WorkSheetHandle shtx = rngs[0].getSheet();
		int x = rngs[0].firstcellrow;

		RowHandle[] rx = rngs[0].getRows();
		boolean found = false;
		// iterate, find, update
		for( int t = 0; t < rx.length; t++ )
		{

		}
	}

	/**
	 * update a row of cells in this named range
	 *
	 * @param an  array of Objects to update existing
	 * @param the column to use as unique index
	 */
	public void updateRow( Object[] objarr, int idxcol ) throws Exception
	{
		CellRange[] rngs = getCellRanges();
		if( rngs.length > 1 )
		{
			throw new WorkBookException( "NamedRange.updateRow Object array failed: too many CellRanges.",
			                             WorkBookException.RUNTIME_ERROR );
		}
		if( rngs.length == 0 )
		{
			throw new WorkBookException( "NamedRange.updateRow Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR );
		}

		//
		WorkSheetHandle shtx = rngs[0].getSheet();
		int x = rngs[0].firstcellrow;

		RowHandle[] rx = rngs[0].getRows();
		boolean found = false;
		// iterate, find, update
		for( int t = 0; t < rx.length; t++ )
		{

			CellHandle[] cx = rx[t].getCells();
			if( cx.length > idxcol )
			{
				if( cx[idxcol].getStringVal().equalsIgnoreCase( objarr[idxcol].toString() ) )
				{
					found = true;
					for( int z = 0; z < cx.length; z++ )
					{
						cx[z].setVal( objarr[z] );
					}
				}
			}
		}

		if( !found )
		{
			this.addRow( objarr );
		}
	}

	/**
	 * add a row of cells to this named range
	 *
	 * @param an array of Objects to insert at last rown
	 */
	public void addRow( Object[] objarr ) throws Exception
	{
		CellRange[] rngs = getCellRanges();
		if( rngs.length > 1 )
		{
			throw new WorkBookException( "NamedRange.add Object array failed: too many CellRanges.", WorkBookException.RUNTIME_ERROR );
		}
		if( rngs.length == 0 )
		{
			throw new WorkBookException( "NamedRange.add Object array failed: zero CellRanges", WorkBookException.RUNTIME_ERROR );
		}

		//
		WorkSheetHandle shtx = rngs[0].getSheet();

		int x = rngs[0].firstcellrow;
		CellHandle[] cxx = shtx.insertRow( x, objarr, true );

		//for(int t=0;t<cxx.length;t++)
		//	if(cxx[t]!=null)rngs[0].addCellToRange(cxx[t]);
		rngs[0] = new CellRange( this.myName.getLocation(), mybook, createblanks );
	}

	/**
	 * add a cell to this named range
	 *
	 * @param cx
	 */
	public void addCell( CellHandle cx ) throws Exception
	{
		CellRange[] rngs = getCellRanges();
		if( rngs.length > 1 )
		{
			throw new WorkBookException(
					"NamedRange.addCell failed -- more than one cell range defined in this NameHandle, cannot determine where to add cell.",
					WorkBookException.RUNTIME_ERROR );
		}
		rngs[0].addCellToRange( cx );
		setLocation( rngs[0].getRange() );
	}

	/**
	 * set the referenced cells for the named range
	 * <p/>
	 * this reference should be in the standard excel syntax including sheetname,
	 * for instance "Sheet1!A1:Z255"
	 */
	public void setLocation( String strloc )
	{

		// if there is no sheetname, and no sheet, then this is an unusable name and cannot be added
		if( (strloc.indexOf( "!" ) == -1) && (strloc.indexOf( "=" ) == -1) )
		{
			if( this.get2DSheetName() == null )
			{
				this.remove();
				throw new IllegalArgumentException( "Named Range References must include a Sheet name." );
			}
			strloc = this.get2DSheetName() + "!" + strloc;
		}
		try
		{
			myName.setLocation( strloc );
		}
		catch( FunctionNotSupportedException e )
		{
			Logger.logWarn( "NameHandle.setLocation :" + strloc + " failed: " + e.toString() );
		}
	}

	/**
	 * get the referenced named cell range as string in standard excel syntax including sheetname,
	 * <p/>
	 * for instance "Sheet1!A1:Z255"
	 */
	public String getLocation()
	{
		try
		{
			String loc = myName.getLocation();
			return loc;
		}
		catch( Exception e )
		{
			Logger.logErr( "Error getting named range location" + e );
			return null;
		}
	}

	/**
	 * set the name String for the range definition
	 *
	 * @param String definition name
	 */
	public void setName( String newname )
	{
		myName.setName( newname );
	}

	/**
	 * returns the name String for the range definition
	 *
	 * @return String definition name
	 */
	public String getName()
	{
		return myName.getName();
	}

	public String toString()
	{
		return getName();
	}

	/**
	 * returns the name's formula String for the range definition
	 *
	 * @return String the expression string
	 */
	public String getExpressionString()
	{
		try
		{
			return myName.getExpressionString();
		}
		catch( Exception e )
		{
			if( DEBUGLEVEL > -1 )
			{
				Logger.logWarn( "Could not parse expression string for name: " + this.getName() );
			}
		}
		return "#ERR";
	}

	/**
	 * removes this Named Range from the WorkBook.
	 *
	 * @return whether the removal was a success
	 */
	public boolean remove()
	{
		boolean success = false;
		try
		{
			success = this.myName.getWorkBook().removeName( myName );
		}
		catch( Exception e )
		{
			return false;
		}
		return success;
	}

	/**
	 * set whether the CellRanges referenced by the NameHandle
	 * will add blank records to the WorkBook for any missing Cells
	 * contained within the range.
	 *
	 * @param b set whether to create blank records for missing Cells
	 */
	public void setCreateBlanks( boolean b )
	{
		createblanks = b;
	}

	/**
	 * gets the array of Cells in this Name
	 * <p/>
	 * NOTE: this method variation also returns the Sheetname for the name record if not null.
	 * <p/>
	 * Thus this method is limited to use with 2D ranges.
	 *
	 * @param fragment whether to enclose result in NameHandle tag
	 * @return Cell[] all Cells defined in this Name
	 */
	public String getCellRangeXML( boolean fragment )
	{
		StringBuffer sbx = new StringBuffer();
		if( !fragment )
		{
			sbx.append( "<?xml version=\"1\" encoding=\"utf-8\"?>" );
		}
		sbx.append( "<NameHandle Name=\"" + this.getName() + "\">" );
		sbx.append( getCellRangeXML() );
		sbx.append( "</NameHandle>" );
		return sbx.toString();
	}

	/**
	 * gets the array of Cells in this Name
	 *
	 * @return Cell[] all Cells defined in this Name
	 */
	public String getCellRangeXML()
	{
		StringBuffer sbx = new StringBuffer();
		try
		{
			CellHandle[] celx = this.getCells();
			RowHandle rowhold = null;
			for( int x = 0; x < celx.length; x++ )
			{
				RowHandle rx = celx[x].getRow();
				if( x == 0 )
				{
					rowhold = rx;
					sbx.append( "<Row Number='" + rowhold.getRowNumber() + "'>" );
					sbx.append( celx[x].getXML() ); // ignores merged ranges here
				}
				else if( rowhold.getRowNumber() == rx.getRowNumber() )
				{
					sbx.append( celx[x].getXML() ); // ignores merged ranges here
				}
				else
				{
					sbx.append( "</Row>" );
					rowhold = rx;
					sbx.append( "<Row Number='" + rowhold.getRowNumber() + "'>" );
					sbx.append( celx[x].getXML() ); // ignores merged ranges here
				}
			}
			sbx.append( "</Row>" );
		}
		catch( CellNotFoundException ex )
		{
			Logger.logErr( "NameHandle.getCellRangeXML failed: ", ex );
		}
		//	sbx.append("</Row>");
		return sbx.toString();
	}

	/**
	 * gets the array of Cells in this Name
	 *
	 * @return Cell[] all Cells defined in this Name
	 */
	public CellHandle[] getCells() throws CellNotFoundException
	{
		if( false )
		{
			return null;
		}
		// first get the ptgArea3d from the parsed expression
		try
		{
			CompatibleVector cellhandles = new CompatibleVector();
			CellRange[] rngz = this.getCellRanges();
			if( rngz != null )
			{
				for( int t = 0; t < rngz.length; t++ )
				{
					try
					{
						CellHandle[] cells = rngz[t].getCells();
						// for(int b = (cells.length-1);b>=0;b--)cellhandles.add(cells[b]);
						for( int b = 0; b < cells.length; b++ )
						{
							if( cells[b] != null )
							{
								cellhandles.add( cells[b] );
							}
						}
					}
					catch( Exception ex )
					{
						Logger.logWarn( "Could not get cells for range: " + rngz[t] );
					}
				}
			}
			else
			{
				return null;
			}
			CellHandle[] ret = new CellHandle[cellhandles.size()];
			return (CellHandle[]) cellhandles.toArray( ret );
		}
		catch( Exception e )
		{
			if( e instanceof CellNotFoundException )
			{
				throw (CellNotFoundException) e;
			}
			throw new CellNotFoundException( e.toString() );
		}
	}

	/**
	 * Get an Array of CellRanges, one per referenced WorkSheet.
	 * <p/>
	 * If this method throws CellNotFoundExceptions, then you are
	 * addressing a sparsely populated CellRange.
	 * <p/>
	 * Use 'setCreateBlanks(true)' to populate these Cells and avoid
	 * this error.
	 *
	 * @return
	 * @throws Exception
	 */
	public CellRange[] getCellRanges() throws Exception
	{
		if( initialRange != null )
		{
			CellRange[] ret = { initialRange };
			return ret;
		}
		// 20100217 KSC: try a better way (that can handle 3D refs and complex cell ranges)
		String loc = myName.getLocation();    // may contain one or more ranges, separated by ","'s if complex
		WorkSheetHandle[] sheets = this.getReferencedSheets();
		// handle commas within quoted sheet names (TestExtenXLSEngine.TestQuotedSheetsWithCommansInNRs)
		// NOTE that sheetnames must be properly qualified for the split to work below 
		String[] nranges = loc.split( ",(?=([^'|\"]*'[^'|\"]*'|\")*[^'|\"]*$)" );
		// below cannot handle commas embedded within quotes
//		String[] nranges= StringTool.splitString(loc, ",");	// TODO: can be another delimeter?
		CellRange[] ranges = new CellRange[nranges.length];
		for( int i = 0; i < nranges.length; i++ )
		{
			String r = nranges[i];
			if( r.indexOf( "!" ) == -1 )
			{ // no sheetname
				r = sheets[i] + "!" + r;
			}
			ranges[i] = new CellRange( r, mybook, createblanks );
			ranges[i].setParent( this.myName );
		}

		return ranges;
	}

	/**
	 * Get WorkSheetHandles for all of the Boundsheets referenced in
	 * this NameHandle.
	 *
	 * @return an array of WorkSheetHandles referenced in this Name
	 */
	public WorkSheetHandle[] getReferencedSheets() throws WorkSheetNotFoundException
	{
		//first get the ptgArea3d from the parsed expression
		Boundsheet[] bs = myName.getBoundSheets();
		if( bs == null )
		{
			throw new WorkSheetNotFoundException( "Worksheet for Named Range: " + this.toString() + ":" + this.myName.getExpressionString() );
		}
		if( bs[0] == null )
		{
			throw new WorkSheetNotFoundException( "Worksheet for Named Range: " + this.toString() + ":" + this.myName.getExpressionString() );
		}
		WorkSheetHandle[] ret = new WorkSheetHandle[bs.length];
		for( int x = 0; x < ret.length; x++ )
		{
			ret[x] = mybook.getWorkSheet( bs[x].toString() );
		}
		return ret;
	}

	/**
	 * return the calculated value of this Name
	 * if it contains a parsed Expression (Formula)
	 *
	 * @return
	 * @throws FunctionNotSupportedException
	 */
	public Object getCalculatedValue() throws FunctionNotSupportedException
	{
		return myName.getCalculatedValue();
	}

	/**
	 * Sets the location lock on the Cell Reference at the
	 * specified  location
	 * <p/>
	 * Used to prevent updating of the Cell Reference when
	 * Cells are moved.
	 *
	 * @param location of the Cell Reference to be locked/unlocked
	 * @param lock     status setting
	 * @return boolean whether the Cell Reference was found and modified
	 */
	public boolean setLocationLocked( String loc, boolean l )
	{
		int x = Ptg.PTG_LOCATION_POLICY_UNLOCKED;
		if( l )
		{
			x = Ptg.PTG_LOCATION_POLICY_LOCKED;
		}
		return myName.setLocationPolicy( loc, x );
	}

	/**
	 * Return a JSON object representing this name Handle.
	 * name:'nameOfRange'
	 * cellrange:'Sheet1!A1:B1'
	 *
	 * @return
	 */
	public String getJSON()
	{
		return getJSON( false );
	}

	/**
	 * Return a JSON object representing this name Handle.
	 * name:'nameOfRange'
	 * cellrange:'Sheet1!A1:B1'
	 * cells:celldata
	 *
	 * @param whether to return cell data
	 * @return
	 */
	public String getJSON( boolean celldata )
	{
		JSONObject theNameHandle = new JSONObject();
		try
		{
			theNameHandle.put( "name", this.getName() );
			theNameHandle.put( "cellrange", myName.getLocation() );

			if( celldata )
			{
				StringBuffer ret = new StringBuffer();
				CellHandle[] cx1 = getCells();
				int p = cx1[0].getRowNum();
				for( int x = 0; x < cx1.length; x++ )
				{
					if( cx1[x] != null )
					{
						if( cx1[x].getRowNum() != p )
						{
							ret.append( "\r\n" );
							p = cx1[x].getRowNum();
						}
						ret.append( "'" );
						try
						{
							ret.append( cx1[x].getStringVal() );
						}
						catch( Exception ex )
						{ // handles empties
							;//
						}
						if( x != (cx1.length - 1) )
						{
							ret.append( "'," );
						}
						else
						{
							ret.append( "'" );
						}
					}
				}

				theNameHandle.put( "celldata", ret.toString() );
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error creating JSON name handle: " + e );
		}
		return theNameHandle.toString();
	}

	/**
	 * Returns a JSON object in the same format as {@link CellRange#getJSON}.
	 */
	public JSONObject getJSONCellRange()
	{
		JSONObject theRange = new JSONObject();
		JSONArray cells = new JSONArray();
		try
		{
			CellHandle[] chandles = getCells();
			for( int i = 0; i < chandles.length; i++ )
			{
				CellHandle thisCell = chandles[i];
				JSONObject result = new JSONObject();

				result.put( JSON_CELL, thisCell.getJSONObject() );
				cells.put( result );
			}
			theRange.put( JSON_CELLS, cells );
		}
		catch( Exception e )
		{
			Logger.logErr( "Error getting NamedRange JSON: " + e );
		}
		return theRange;
	}

	/**
	 * return the sheetname for a 2D named range
	 * <p/>
	 * NOTE: Does not work for 3D ranges
	 *
	 * @return the Sheet name if this is a 2d range
	 */
	public String get2DSheetName()
	{
		try
		{
			return this.myName.getBoundSheets()[0].getSheetName();
		}
		catch( Exception e )
		{
			try
			{
				return this.myName.getSheet().getSheetName();
			}
			catch( Exception ex )
			{
				return null;
			}
		}
	}

}