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
package com.extentech.formats.XLS.charts;

import com.extentech.formats.XLS.ExpressionParser;
import com.extentech.formats.XLS.Externsheet;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.formulas.GenericPtg;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgArea3d;
import com.extentech.formats.XLS.formulas.PtgMemFunc;
import com.extentech.formats.XLS.formulas.PtgMystery;
import com.extentech.formats.XLS.formulas.PtgParen;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.formats.XLS.formulas.PtgRef3d;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;

import java.util.Iterator;
import java.util.List;
import java.util.Stack;

/**
 * <b>Ai: Linked Chart Data (1051h)</b><br>
 * <p/>
 * This record specifies linked series data or text.
 * <p/>
 * Offset           Name    Size    Contents
 * --
 * 4               id      1       index id: (0=title or text, 1=series vals, 2=series cats + (3=bubbles)
 * 5               rt      1       reference type(0=default,1=text in formula bar, 2=worksheet, 4=error)
 * 6               grbit   2       flags
 * 8               ifmt    2       Index to number format (used if fCustomIfmt= true)
 * 10              cce     2       size of rgce (in bytes)
 * 12              rgce    var     Parsed formula of link
 * <p/>
 * The grbit field contains the following option flags.
 * <p/>
 * 0	0	0x1		fCustomIfmt		TRUE if this object has a custom number format; FALSE if number format is linked to data source
 * a	1	0x2		(reserved)		Reserved; must be zero
 * 0	5-2	0x3C	st				Source type (always zero)
 * a	7-6	0xCO	(reserved)		Reserved; must be zero
 * 1	7-0	0xFF	(reserved)		Reserved; must be zero
 * <p/>
 * </pre>
 *
 * @see Formula
 * @see Chart
 */

public final class Ai extends GenericChartObject implements ChartObject
{
	/**
	 *
	 */
	private static final long serialVersionUID = -6647823755603289012L;
	private Stack expression;
	protected int id = -1, ifmt = -1, cce = -1;
	private short grbit = -1, rt = -1;
	private boolean fCustomIfmt = false;
	//private CompatibleVector xlsrecs = new CompatibleVector();
	private SeriesText st = null;
	// define the type of Ai record
	public static final int TYPE_TEXT = 0;
	public static final int TYPE_VALS = 1;
	public static final int TYPE_CATEGORIES = 2;
	public static final int TYPE_BUBBLES = 3;

	/**
	 * This is for storage of the boundsheet name when the ai record is being moved from
	 * one workbook to another.  In these cases, populate this value, and pull it back out
	 * to locate the cxti and reset the reerences
	 *
	 * @return the name of the original boundsheet this record is associated with.
	 */
	private String boundName = "";
	private String origSheetName = "";
	private int boundXti = -1;

	/**
	 * returns the bound sheet name - must be called after populateForTransfer
	 *
	 * @return String
	 */    // 20080116 KSC: Changed name for clarity
	public String getBoundName()
	{
		return boundName;
	}

	/**
	 * returns the sheet reference index - must call after populateForTransfer
	 *
	 * @return
	 */
	public int getBoundXti()
	{
		return boundXti;
	}

	/**
	 * Sets the boundsheet name for the data referenced in this AI.
	 * + sets the xti of the boundsheet reference (necessary upon sheet copy/move)
	 * Does not currently support multi-boundsheet references.
	 *
	 * @see updateSheetRef
	 */
	public void populateForTransfer( String origSheetName )
	{
		if( "".equals( boundName ) )
		{
			this.origSheetName = origSheetName;    // 20080708 KSC: trap original sheet name for updateSheetRefs comparison
			for( Object anExpression : expression )
			{
				Ptg p = (Ptg) anExpression;
				if( p instanceof PtgArea3d )
				{
					PtgArea3d pt = (PtgArea3d) p;
					try
					{
						boundName = pt.getSheetName();        //original boundname + xti sheet ref
						boundXti = pt.getIxti();
					}
					catch( Exception e )
					{
						Logger.logErr( "Ai.populateForTransfer: Chart contains links to other data sources" );
					}
				}
				else if( p instanceof PtgRef3d )
				{
					PtgRef3d pt = (PtgRef3d) p;
					try
					{
						boundName = pt.getSheetName();       //original boundname + xti sheet ref
						boundXti = pt.getIxti();
					}
					catch( Exception e )
					{
						Logger.logErr( "Ai.populateForTransfer: Chart contains links to other data sources" );
					}
				}
				else if( p instanceof PtgMemFunc )
				{    // 20091015 KSC: for non-contiguous ranges
					Ptg ptg = ((PtgMemFunc) p).getFirstloc();
					if( ptg instanceof PtgRef3d )
					{
						PtgRef3d pr = (PtgRef3d) ptg;
						try
						{
							boundName = pr.getSheetName();       //original boundname + xti sheet ref
							boundXti = pr.getIxti();
						}
						catch( Exception e )
						{
							Logger.logErr( "Ai.populateForTransfer: Chart contains links to other data sources" );
						}
					}
					else
					{    // should be a PtgArea3d
						PtgArea3d pr = (PtgArea3d) ptg;
						try
						{
							boundName = pr.getSheetName();       //original boundname + xti sheet ref
							boundXti = pr.getIxti();
						}
						catch( Exception e )
						{
							Logger.logErr( "Ai.populateForTransfer: Chart contains links to other data sources" );
						}
					}
				}
			}
		}
	}

	public void setLegend( String newLegend )
	{
		this.setRt( 1 );
		st.setText( newLegend );
		expression = null;    // if setting to a string, no expression!
	}

	/**
	 * Refer to the associated Series Text.
	 *
	 * @param txt
	 * @return
	 */
	public boolean setText( String txt )
	{
		try
		{
//			XLSRecord x = getRecord(0);
//			((SeriesText)x ).setText(txt);
			st.setText( txt );
			return true;
		}
		catch( ClassCastException e )
		{
			// Logger.logInfo("Error getting Chart String value: " + e);
			return false;
		}
	}

	public String getText()
	{
		try
		{
//			XLSRecord x = getRecord(0);
//			if(x instanceof SeriesText)return  ((SeriesText)x ).toString();
			if( st != null )
			{
				return st.toString();
			}
		}
		catch( Exception e )
		{
			if( DEBUGLEVEL > 0 )
			{
				Logger.logWarn( "Error getting Chart String value: " + e );
			}
			return getDefinition();
		}
		//TODO: figure out why this doesn't find the title -- see "reportS01Template.xls"
		return "undefined";
	}

	public String toString()
	{
		switch( id )
		{
			case Ai.TYPE_TEXT:
				return getText();
			case Ai.TYPE_VALS:
				return getDefinition();
			case Ai.TYPE_CATEGORIES:
				return getDefinition();
			case Ai.TYPE_BUBBLES:
				return getDefinition();
		}
		return super.toString();
	}

	//	public void addRecord(BiffRec rec){
//		xlsrecs.add(rec);	 
//	}
	public void setSeriesText( SeriesText s )
	{
		st = s;
	}

//	protected XLSRecord getRecord(int i){
//		return (XLSRecord) xlsrecs.get(i);
//	}

	/**
	 * set the Externsheet reference
	 * for any associated PtgArea3d's
	 */
	public void setExternsheetRef( int x ) throws WorkSheetNotFoundException
	{
		byte[] dt = this.getData();
		int pos = 8;

		for( Object anExpression : expression )
		{
			Ptg p = (Ptg) anExpression;
			if( p instanceof PtgArea3d )
			{
				PtgArea3d pt = (PtgArea3d) p;
				pt.setIxti( (short) x );
				pt.addToRefTracker();
				if( DEBUGLEVEL > 3 )
				{
					Logger.logInfo( "Setting sheet reference for: " + pt.toString() + "  in Ai record." );
				}
				// register with the Externsheet reference
				this.getWorkBook().getExternSheet().addPtgListener( pt );
				updateRecord();
			}
			else if( p instanceof PtgRef3d )
			{    // 20091015 KSC: Added
				PtgRef3d pr = (PtgRef3d) p;
				pr.setIxti( (short) x );
				pr.addToRefTracker();
				if( DEBUGLEVEL > 3 )
				{
					Logger.logInfo( "Setting sheet reference for: " + pr.toString() + "  in Ai record." );
				}
				// register with the Externsheet reference
				this.getWorkBook().getExternSheet().addPtgListener( pr );
				updateRecord();
			}
			else if( p instanceof PtgMemFunc )
			{    // 20091015 KSC: Added
				Ptg pr = ((PtgMemFunc) p).getFirstloc();
				if( pr instanceof PtgRef3d )
				{
					((PtgRef3d) pr).setIxti( (short) x );
					((PtgRef3d) pr).addToRefTracker();
					this.getWorkBook().getExternSheet().addPtgListener( (PtgRef3d) pr );
				}
				else
				{    // should be a PtgArea3d
					((PtgArea3d) pr).setIxti( (short) x );
					((PtgArea3d) pr).addToRefTracker();
					this.getWorkBook().getExternSheet().addPtgListener( (PtgArea3d) pr );
				}
				if( DEBUGLEVEL > 3 )
				{
					Logger.logInfo( "Setting sheet reference for: " + pr.toString() + "  in Ai record." );
				}
				// register with the Externsheet reference
				updateRecord();
			}
			else
			{
				Logger.logInfo( "Ai.setExternsheetRef: unknown Ptg" );
			}
		}
	}

	/**
	 * set the Externsheet reference
	 * for any associated PtgArea3d's that match the old reference.
	 * <p/>
	 * invaluble for modifying only one sheets worth of references (ie a move sheet situation)
	 */
	public void setExternsheetRef( int oldRef, int newRef ) throws WorkSheetNotFoundException
	{
		for( Object anExpression : expression )
		{
			Ptg p = (Ptg) anExpression;
			if( p instanceof PtgArea3d )
			{
				PtgArea3d pt = (PtgArea3d) p;
				int oRef = pt.getIxti();
				if( oRef == oldRef )
				{    // got the one to update
					pt.removeFromRefTracker();    // 20100506 KSC: added
					pt.setSheetName( this.getSheet().getSheetName() );    // 20100415 KSC: added
					pt.setIxti( (short) newRef );
					pt.addToRefTracker();    // 20080709 KSC
					if( DEBUGLEVEL > 3 )
					{
						Logger.logInfo( "Setting sheet reference for: " + pt.toString() + "  in Ai record." );
					}
					// register with the Externsheet reference
					this.getWorkBook().getExternSheet().addPtgListener( pt );
					updateRecord();
				}
			}
			else if( p instanceof PtgRef3d )
			{
				PtgRef3d pr = (PtgRef3d) p;
				int oRef = pr.getIxti();
				if( oRef == oldRef )
				{
					pr.removeFromRefTracker();
					pr.setSheetName( this.getSheet().getSheetName() );    // 20100415 KSC: added
					pr.setIxti( (short) newRef );
					if( !pr.getIsRefErr() )
					{
						pr.addToRefTracker();
					}
					if( DEBUGLEVEL > 3 )
					{
						Logger.logInfo( "Setting sheet reference for: " + pr.toString() + "  in Ai record." );
					}
					// register with the Externsheet reference
					this.getWorkBook().getExternSheet().addPtgListener( pr );
					updateRecord();
				}
			}
			else if( p instanceof PtgMemFunc )
			{    // 20091015 KSC: Added
				Ptg pr = ((PtgMemFunc) p).getFirstloc();
				if( pr instanceof PtgRef3d )
				{
					int oRef = ((PtgRef3d) pr).getIxti();
					if( oRef == oldRef )
					{
						((PtgRef3d) pr).removeFromRefTracker();    // 20100506 KSC: added
						((PtgArea3d) pr).setSheetName( this.getSheet().getSheetName() );    // 20100415 KSC: added
						((PtgRef3d) pr).setIxti( (short) newRef );
						((PtgRef3d) pr).addToRefTracker();
						this.getWorkBook().getExternSheet().addPtgListener( (PtgRef3d) pr );
					}
				}
				else
				{    // should be a PtgArea3d
					int oRef = ((PtgRef3d) pr).getIxti();
					if( oRef == oldRef )
					{
						((PtgRef3d) pr).removeFromRefTracker();    // 20100506 KSC: added
						((PtgArea3d) pr).setSheetName( this.getSheet().getSheetName() );    // 20100415 KSC: added
						((PtgArea3d) pr).setIxti( (short) newRef );
						((PtgArea3d) pr).addToRefTracker();
						this.getWorkBook().getExternSheet().addPtgListener( (PtgArea3d) pr );
					}
				}
				if( DEBUGLEVEL > 3 )
				{
					Logger.logInfo( "Setting sheet reference for: " + pr.toString() + "  in Ai record." );
				}
				// register with the Externsheet reference
				updateRecord();
			}
			else if( p instanceof PtgMystery )
			{
				// TODO: do what???
			}
		}
	}

	/**
	 * take the original boundSheet + boundSheet xti reference and update it to
	 * the sheet reference in the new workbook (or same workbook but different ixti reference)
	 */
	public void updateSheetRef( String newSheetName, String origWorkBookName )
	{
		try
		{
			// 20080630/20080708 KSC: Fixes for [BugTracker 1799] + [BugTracker 1434]
			if( boundXti > -1 )
			{    // has populate for transfer been called - should!
				int newSheetNum = -1;
				try
				{
					if( !boundName.equalsIgnoreCase( origSheetName ) )
					{    // Ai reference is on a dfferent sheet, see if it exists in new workbook
						newSheetNum = this.getWorkBook().getWorkSheetByName( boundName ).getSheetNum();
					}
					else    // Ai reference is on same sheet, point now to new sheet
					{
						newSheetNum = this.getWorkBook().getWorkSheetByName( newSheetName ).getSheetNum();
					}
				}
				catch( Exception e )
				{ // 20080123 KSC: if links arent there, fix == try to make an External ref
					for( Object anExpression : expression )
					{
						if( anExpression instanceof PtgArea3d )
						{
							PtgArea3d p = (PtgArea3d) anExpression;
							Logger.logWarn( "External References are unsupported: External reference found in Chart: " + p.getSheetName() );
							p.setSheetName( boundName );    // set external reference to original boundsheet name
							p.setExternalReference( origWorkBookName );
							this.setExternsheetRef( p.getIxti() );
						}
						else
						{
							Logger.logInfo( "Ai.updateSheetRef:" );
						}
					}
				}
				if( newSheetNum != -1 )
				{
					this.setSheet( this.getWorkBook().getWorkSheetByName( newSheetName ) ); // 20100415 KSC: set Ai sheet ref to new sheet
					Externsheet xsht = this.getWorkBook().getExternSheet( true );    // create if necessary
					int newXRef = xsht.insertLocation( newSheetNum, newSheetNum );
					this.setExternsheetRef( boundXti, newXRef );
					boundXti = newXRef;    // 20100506 KSC: reset
					boundName = newSheetName;    // ""
				}
			}
			else
			{// debugging 20100415
//				Logger.logErr("Ai.updateSheetRef: boundxti is -1 for AI " + this.toString());
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Ai.updateSheetRef: " + e.toString() );
		}
	}

	/**
	 * get the display name
	 */
	String getName()
	{
		return "Chart Ai";
	}

	/**
	 * get the definition text
	 * the definition is stored
	 * in Excel parsed format
	 */
	public String getDefinition()
	{
		StringBuffer sb = new StringBuffer();
		Ptg[] ep = new Ptg[expression.size()];
		ep = (Ptg[]) expression.toArray( ep );
		for( Ptg anEp : ep )
		{
			if( !(anEp instanceof PtgParen) )    // 20091019 KSC: if complex series, will have a PtgParen after PtgMemFunc;
/*				if (ep[t] instanceof PtgMemFunc)
					sb.append("(" + ep[t].getString() + ")");
				else*/
			{
				sb.append( anEp.getString() );
			}
		}
		return sb.toString();
	}

	/*  Returns an array of ptgs that represent any BiffRec ranges in the formula.  
		Ranges can either be in the format "C5" or "Sheet1!C4:D9"      
	*/
	public Ptg[] getCellRangePtgs() throws FormulaNotFoundException
	{
		return ExpressionParser.getCellRangePtgs( expression );
	}

	/**
	 * Return the type (ID) of this Ai record
	 *
	 * @return int rt
	 */
	public int getType()
	{
		return id;
	}

	/**
	 * return the custom number format for this AI (category, series or bubble)
	 * <br>if 0, use default number format as specified by Axis or Chart
	 *
	 * @return
	 */
	public int getIfmt()
	{
		return ifmt;
	}

	/*  Returns the ptg that matches the string location sent to it.
		this can either be in the format "C5" or a range, such as "C4:D9"      
	*/
	public List getPtgsByLocation( String loc )
	{
		try
		{
			return ExpressionParser.getPtgsByLocation( loc, expression );
		}
		catch( FormulaNotFoundException e )
		{
			Logger.logWarn( "failed to update Chart Series Location: " + e );
		}
		return null;
	}

	/**
	 * locks the Ptg at the specified location
	 */

	public boolean setLocationPolicy( String loc, int l )
	{
		List dx = this.getPtgsByLocation( loc );
		Iterator lx = dx.iterator();
		while( lx.hasNext() )
		{
			Ptg d = (Ptg) lx.next();
			if( d != null )
			{
				d.setLocationPolicy( l );
			}
		}
		return true;
	}

	/**
	 * simplified version of changeAiLocation which locates current Ptg and updates expression
	 * NOTE: newLoc is expected to be a valid reference
	 *
	 * @param p
	 * @param newLoc
	 * @return
	 */
	public boolean changeAiLocation( Ptg p, String newLoc )
	{
		String[] aiLocs = StringTool.splitString( newLoc, "," );
		for( String aiLoc : aiLocs )
		{
			try
			{    // NOTE: Ai has only 1 expression OR 2: 1= PtgMemFunc 2=PtgParen
				// looking up via getExpressionLocByPtg errors if loc is the same but Ptg's are different objects eg. PtgRef3d vs PtgArea3d ...
				if( expression.get( 0 ) instanceof PtgMemFunc )
				{
					try
					{    // must find particular ptg in PtgMemFunc's subexression to reset
						// APPARENTLY WE DO NOT NEED TO UDPATE PTGMEMFUNC SUBEXPRESSION EXPLICITLY
						//int z= ExpressionParser.getExpressionLocByPtg(p, ((PtgMemFunc) expression.get(0)).getSubExpression());
//						Stack subexp= ((PtgMemFunc) expression.get(0)).getSubExpression();
//						for (int z= 0; z < subexp.size(); z++) {
//							if (p.equals(subexp.get(z))) {
						p.setLocation( aiLoc );    // updates ref. tracker
//								((PtgMemFunc) expression.get(0)).getSubExpression().set(z, p);	// update expression with new Ptg
//							}
//						}
					}
					catch( Exception ex )
					{
						Logger.logErr( "Ai.changeAiLocation: Error updating Location in non-contiguous range " + ex.toString() );
					}
				}
				else
				{
					p.setLocation( aiLoc );    // updates ref. tracker
					expression.set( 0, p );    // update expression with new Ptg
					if( this.getType() == Ai.TYPE_TEXT )
					{    // must reset text for SeriesText as well
						try
						{
							Object o = p.getValue();
							this.setText( o.toString() );
						}
						catch( Exception e )
						{
							;
						}
					}
				}
			}
			catch( Exception e )
			{
				Logger.logErr( "Ai.changeAiLocation: Error updating Location to " + newLoc + ":" + e.toString() );
				return false;
			}
		}
		updateRecord();
		return true;
	}

	/**
	 * Takes a string as a current formula location, and changes
	 * that pointer in the formula to the new string that is sent.
	 * This can take single cells "A5" and cell ranges, "A3:d4"
	 * Returns true if the cell range specified in formulaLoc exists & can be changed
	 * else false.  This also cannot change a cell pointer to a cell range or vice
	 * versa.
	 *
	 * @param String - range of Cells within Formula to modify
	 * @param String - new range of Cells within Formula
	 */
	public boolean changeAiLocation( String loc, String newLoc )
	{
		// TODO: Implement formula policy!! -jm
		Ptg ptg = null;
		int z = -1;
		try
		{
			if( expression.size() > 0 )
			{
				z = ExpressionParser.getExpressionLocByLocation( loc, expression );
				ptg = (Ptg) expression.get( z );    // 20090917 KSC: since creating new ptgs below, must remove original from reftracker "by hand"
			}
			if( ptg != null )
			{
				((PtgRef) ptg).removeFromRefTracker();
			}
		}
		catch( Exception e )
		{
		}
		if( (z == -1) && newLoc.equals( "" ) )
		{// no reference -- happens on legends, category ai's ...
			this.getData()[1] = 1;    // text reference rather than worksheet reference
			return false;
		}
		ptg = PtgRef.createPtgRefFromString( newLoc, this );
		if( z != -1 )    // then must change original
		{
			expression.set( z, ptg );    // update expression with new Ptg
		}
		else
		{
			expression.add( ptg );
		}
		updateRecord();
		return true;
	}

	/*  Update the record byte array with the modified ptg records
	*/
	public void updateRecord()
	{
		int offy = 8; // the start of the parsed expression
		byte[] rkdata = this.getData();
		byte[] updated = new byte[rkdata.length];
		System.arraycopy( rkdata, 0, updated, 0, offy );
		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			Ptg ptg = (Ptg) o;
			byte[] b = ptg.getRecord();
			//	 must inc. size if Ptgs have inc.'d ... see changeAiLocation 
			int len = b.length;
			if( (updated.length - offy) < len )
			{
				byte[] newArr = new byte[offy + len];
				System.arraycopy( updated, 0, newArr, 0, updated.length );
				// update cce in array as well ...
				cce += (newArr.length - updated.length);
				updated = newArr;
				byte[] ix = ByteTools.shortToLEBytes( (short) cce );
				System.arraycopy( ix, 0, updated, 6, 2 );
			}
			System.arraycopy( b, 0, updated, offy, len );
			offy = offy + len;
		}
		this.setData( updated );
	}

	@Override
	public void init()
	{
		super.init();
		id = this.getByteAt( 0 );
		// index id: (0=title or text, 1=series vals, 2=series cats, 3= bubbles
		rt = this.getByteAt( 1 );
		// reference type(0=default,1=text in formula bar, 2=worksheet, 4=error)
		grbit = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		fCustomIfmt = (grbit & 0x1) == 0x1;
		// flags
		ifmt = (int) ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		// Index to number format
		cce = (int) ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		// size of rgce (in bytes)
		int pos = 8;
		//  Parsed formula of link

		// 	get the parsed expression
		byte[] expressionbytes = this.getBytesAt( pos, cce );
		expression = ExpressionParser.parseExpression( expressionbytes, this );
		pos += cce;
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( this.getName() + ":" + this.getDefinition() );
		}
	}

	/**
	 * set reference type(0=default,1=text in formula bar, 2=worksheet, 4=error)
	 *
	 * @param i
	 */
	public void setRt( int i )
	{
		rt = (short) i;
		this.getData()[1] = (byte) rt;
	}

	/**
	 * Get a prototype with the specified ai types.
	 * <p/>
	 * 0 = Legend AI
	 * 1=  Series Value Ai
	 * 2 = Category Ai
	 * 3 = Unknown, undocumented, but neccesarry AI
	 * 4 = Blank Legend AI with no reference.
	 */
	public static ChartObject getPrototype( byte[] aiType )
	{
		Ai ai = new Ai();
		ai.setOpcode( AI );
		ai.setData( aiType );
		ai.init();
		return ai;
	}

	// 20070801 KSC: since changeAiLocation now allows addition of new expression bytes, alter
	// default prototype bytes here to not include any expression bytes ..
	protected static byte[] AI_TYPE_LEGEND = new byte[]{ 0, 2, 0, 0, 0, 0, 0, 0 }; //, 7, 0, 58, 0, 0, 0, 0, 0, 0};
	protected static byte[] AI_TYPE_SERIES = new byte[]{ 1, 2, 0, 0, 0, 0, 0, 0 }; //, 11, 0, 59, 0, 0, 1, 0, 1, 0, 1, 0, 3, 0};
	protected static byte[] AI_TYPE_CATEGORY = new byte[]{ 2, 2, 0, 0, 0, 0, 0, 0 }; // 11, 0, 59, 0, 0, 0, 0, 0, 0, 1, 0, 3, 0};
	protected static byte[] AI_TYPE_BUBBLE = new byte[]{ 3, 1, 0, 0, 0, 0, 0, 0 };
	protected static byte[] AI_TYPE_NULL_LEGEND = new byte[]{ 0, 1, 0, 0, 0, 0, 0, 0 };

	public Stack getExpression()
	{
		return expression;
	}

	@Override
	public void close()
	{
		if( expression != null )
		{
			while( !expression.isEmpty() )
			{
				GenericPtg p = (GenericPtg) expression.pop();
				if( p instanceof PtgRef )
				{
					p.close();
				}
				else
				{
					p.close();
				}
				p = null;
			}
		}
		super.close();
	}

	@Override
	protected void finalize()
	{
		this.close();
	}
}