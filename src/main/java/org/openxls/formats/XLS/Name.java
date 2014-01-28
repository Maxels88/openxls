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
package org.openxls.formats.XLS;

import org.openxls.formats.XLS.formulas.FormulaCalculator;
import org.openxls.formats.XLS.formulas.FormulaParser;
import org.openxls.formats.XLS.formulas.GenericPtg;
import org.openxls.formats.XLS.formulas.IlblListener;
import org.openxls.formats.XLS.formulas.IxtiListener;
import org.openxls.formats.XLS.formulas.Ptg;
import org.openxls.formats.XLS.formulas.PtgArea3d;
import org.openxls.formats.XLS.formulas.PtgArray;
import org.openxls.formats.XLS.formulas.PtgMemFunc;
import org.openxls.formats.XLS.formulas.PtgMystery;
import org.openxls.formats.XLS.formulas.PtgRef;
import org.openxls.formats.XLS.formulas.PtgRef3d;
import org.openxls.formats.XLS.formulas.PtgRefErr3d;
import org.openxls.formats.XLS.formulas.PtgStr;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Stack;

/**
 * <b>Name: Defined Name (218h)</b><br>
 * <p/>
 * Name records describe a name in the workbook
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       grbit           2       option flags
 * 6       chKey           1       Keyboard Shortcut
 * 7       cch             1       length of name text
 * 8       cce             2       length of name definition *stored in Excel parsed format
 * 10      ixals           2       index to sheet containing name
 * 12      itab            2       NAME SCOPE -- 0= workbook, 1+= sheet
 * 14      cchCustMenu     1       length of custom menu text
 * 15      cchDescript     1       length of description text
 * 16      cchHelpTopic    1       length of help topic text
 * 17      cchStatusText   1       length of status bar text
 * 18      rgch            var     name text
 * var     rgce            var     name definition
 * var     rcchCustMenu    var     cust menu text
 * var     rgchDescr       var     description text
 * var     rgchHelpTopic   var     help text
 * var     rgchStatusText  var     status bar text
 * <p/>
 * </p></pre>
 */
public final class Name extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Name.class );
	private static final long serialVersionUID = -7868028144327389601L;
	short grbit = -1;
	boolean builtIn = false;
	String rgch = "";            //     name text
	String rgce = "";            //     name definition
	String rcchCustMenu = "";    //     cust menu text
	String rgchDescr = "";       //     description text
	String rgchHelpTopic = "";   //     help text
	String rgchStatusText = "";  //     status bar text
	byte chKey = -1;
	byte cch = -1;
	short cce = -1;   // 2
	short ixals = -1;   // 2
	short itab = -1;   // 2
	byte cchCustMenu = -1;
	byte cchDescript = -1;
	byte cchHelpTopic = -1;
	byte cchStatusText = -1;
	public byte builtInType = -1;
	private Stack expression = null;
	Externsheet externsheet;
	private Ptg ptga;
	private ArrayList ilblListeners = new ArrayList();
	private String cachedOOXMLExpression = null;
	private byte[] expressionbytes = null;        // for deferred Name expression init

	protected static final byte CONSOLIDATE_AREA = 0x0;
	protected static final byte AUTO_OPEN = 0x1;
	protected static final byte AUTO_CLOSE = 0x2;
	protected static final byte EXTRACT = 0x3;
	protected static final byte DATABASE = 0x4;
	protected static final byte CRITERIA = 0x5;
	protected static final byte PRINT_AREA = 0x6;
	protected static final byte PRINT_TITLES = 0x7;
	protected static final byte RECORDER = 0x8;
	protected static final byte DATA_FORM = 0x9;
	protected static final byte AUTO_ACTIVATE = 0xA;
	protected static final byte AUTO_DEACTIVATE = 0xB;
	protected static final byte SHEET_TITLE = 0xC;
	protected static final byte _FILTER_DATABASE = 0xD;

	/* this byte array differs from the prototype in that it has no description field.  It is 
		used by formulas inserting names that don't exist within the workbook.  The name will 
		show up in the formula, but nowhere else on the excel file
	**/
	private byte[] FILLER_NAME_BYTES = { 0x0, 0x0, 0x0, 0x1, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0 };

	// 20090811 KSC: this prototype contains a range; clear out	private byte[] PROTOTYPE_NAME_BYTES = {0x0,0x0,0x0,0x6,0xb,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x61,0x64,0x66,0x61,0x64,0x66,0x3b,0x1,0x0,0x21,0x0,0x21,0x0,0x1,0x0,0x3,0x0};
	private byte[] PROTOTYPE_NAME_BYTES = { 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0 };

	public Name()
	{
		// default constructor
	}

	public Name( WorkBook bk, String namestr )
	{
		byte[] bl = PROTOTYPE_NAME_BYTES;
		setData( bl );
		setOpcode( NAME );
		setLength( (short) (bl.length) );
		setWorkBook( bk );
		try
		{
			setExternsheet( bk.getExternSheet() );
		}
		catch( WorkSheetNotFoundException x )
		{
			log.warn( "Name could not reference WorkBook Externsheet." + x.toString() );
		}

		init( false );
		setName( namestr );
		bk.insertName( this );
	}

	/**
	 * Used for default name entry in formulas.  The name will not exist in the workbook outside
	 * of the formula it is used in.  It has no location.
	 *
	 * @param bk Workbook containing name
	 * @param b  Use whatever, just a flag to use this different constructor
	 */
	public Name( WorkBook bk, boolean b )
	{
		byte[] bl = FILLER_NAME_BYTES;
		setData( bl );
		setOpcode( NAME );
		setLength( (short) (bl.length) );
		setWorkBook( bk );
		try
		{
			setExternsheet( bk.getExternSheet() );
		}
		catch( WorkSheetNotFoundException x )
		{
			log.warn( "Name could not reference WorkBook Externsheet." + x.toString() );
		}
		init();
		bk.insertName( this );
	}

	/**
	 * Store ptgName references to this Name record
	 * so they can be accessed
	 */
	public void addIlblListener( IlblListener ptgname )
	{
		ilblListeners.add( ptgname );
	}

	public void removeIlblListener( IlblListener ptgname )
	{
		ilblListeners.remove( ptgname );
	}

	public ArrayList getIlblListeners()
	{
		return ilblListeners;
	}

	public void updateIlblListeners()
	{
		short ilbl = (short) getWorkBook().getNameNumber( getName() );
		Iterator i = ilblListeners.iterator();
		while( i.hasNext() )
		{
			((IlblListener) i.next()).setIlbl( ilbl );
		}
	}

	/**
	 * Initialize the Name record
	 *
	 * @see XLSRecord#init()
	 */
	@Override
	public void init()
	{
		init( true );        // default= init Expression
	}

	/**
	 * init Name record
	 *
	 * @param initExpression true if should parse formula/ref expression (will be false on wb load)
	 */
	public void init( boolean initExpression )
	{
		super.init();
		getData();
		//  Logger.logInfo("[" + ByteTools.getByteString(data, false) + "]");

		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		chKey = getByteAt( 2 );
		cch = getByteAt( 3 );
		cce = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		ixals = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		itab = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );

		cchCustMenu = getByteAt( 10 );
		cchDescript = getByteAt( 11 );
		cchHelpTopic = getByteAt( 12 );
		cchStatusText = getByteAt( 13 );

		if( (grbit & 0x20) == 0x20 )
		{
			builtIn = true;
		}

		int pos = 15;
		if( getByteAt( 14 ) == 0x1 )
		{
			cch *= 2;
		}// rich byte;
		// get the Name
		try
		{
			byte[] namebytes = getBytesAt( pos, cch );
			if( getByteAt( 14 ) == 0x1 )
			{
				rgch = new String( namebytes, UNICODEENCODING );
			}
			else if( builtIn )
			{
				builtInType = namebytes[0];
				switch( builtInType )
				{
					case CONSOLIDATE_AREA:
						rgch = "Built-in: CONSOLIDATE_AREA";
						break;

					case AUTO_OPEN:
						rgch = "Built-in: AUTO_OPEN";
						break;

					case AUTO_CLOSE:
						rgch = "Built-in: AUTO_CLOSE";
						break;

					case EXTRACT:
						rgch = "Built-in: EXTRACT";
						break;

					case DATABASE:
						rgch = "Built-in: DATABASE";
						break;

					case CRITERIA:
						rgch = "Built-in: CRITERIA";
						break;

					case PRINT_AREA:
						rgch = "Built-in: PRINT_AREA";
						break;

					case PRINT_TITLES:
						rgch = "Built-in: PRINT_TITLES";
						break;

					case RECORDER:
						rgch = "Built-in: RECORDER";
						break;

					case DATA_FORM:
						rgch = "Built-in: DATA_FORM";
						break;

					case AUTO_ACTIVATE:
						rgch = "Built-in: AUTO_ACTIVATE";
						break;

					case AUTO_DEACTIVATE:
						rgch = "Built-in: AUTO_DEACTIVATE";
						break;

					case SHEET_TITLE:
						rgch = "Built-in: SHEET_TITLE";
						break;

					case _FILTER_DATABASE:
						rgch = "Built-in: _FILTER_DATABASE";
						break;
				}
			}
			else
			{
				rgch = new String( namebytes );
			}
				log.debug( getName() );
			pos += cch;

			// get the parsed expression
		        /*byte[] */
			expressionbytes = getBytesAt( pos, cce );
			if( initExpression )
			{
				parseExpression();
			}

		}
		catch( Exception e )
		{
				log.warn( "problem reading Name record expression for Name:" + getName() + " " + e );
		}
	}

	/**
	 * parse Expression separately from init
	 */
	public void parseExpression()
	{
		if( (expressionbytes != null) && (expression == null) )
		{
			expression = ExpressionParser.parseExpression( expressionbytes, this );
			if( expression == null )
			{
				PtgMystery gpg = new PtgMystery();
				gpg.init( expressionbytes );
				expression = new Stack();
				expression.push( gpg );
			}
			expressionbytes = null;
		}
	}

	public Boundsheet[] getBoundSheets()
	{
		if( ptga == null )
		{
			try
			{
				initPtga();
			}
			catch( Exception e )
			{
				;
			}
		}
		if( ptga instanceof PtgRef3d )
		{
			PtgRef3d p3d = (PtgRef3d) ptga;
			Boundsheet b = p3d.getSheet( getWorkBook() );
			Boundsheet[] ret = new Boundsheet[1];
			ret[0] = b;
			return ret;
		}
		if( ptga instanceof PtgArea3d )
		{
			PtgArea3d p3d = (PtgArea3d) ptga;
			return p3d.getSheets( getWorkBook() );
		}
		if( ptga instanceof PtgMemFunc )
		{
			PtgMemFunc p = (PtgMemFunc) ptga;
			return p.getSheets( getWorkBook() );
		}
		return null;
	}

	/**
	 * Return the name which identifies this Name record.  If this is a built in record
	 * it will return a generic version of what the built-in is doing.
	 *
	 * @see XLSRecord#toString()
	 */
	public String toString()
	{
		return getName();
	}

	/**
	 * Return the expression string for this name record.
	 *
	 * @return
	 */
	public String getExpressionString()
	{
		if( ((expression == null) || (expression.size() == 0)) && (getCachedOOXMLExpression() != null) )
		{
			return "=" + getCachedOOXMLExpression();
		}
		if( expression != null )
		{
			return FormulaParser.getExpressionString( expression );
		}
		return "=";
	}

	/**
	 * Get the expression for this Name record
	 *
	 * @return
	 */
	public Stack getExpression()
	{
		return expression;
	}

	/**
	 * set the expression for this Name record
	 */
	public void setExpression( Stack x )
	{
		expression = x;
		updatePtgs();
	}

	/**
	 * Return the location of the Name record.  This seems to be slightly wrong as it only
	 * returns the location of PTGA.  It's possible to have more complex records than this.
	 *
	 * @return
	 * @throws Exception
	 */
	public String getLocation() throws Exception
	{
		if( ptga == null )
		{
			try
			{
				initPtga();
			}
			catch( Exception e )
			{
				log.warn( "Name.getLocation() failed: " + e.toString() );
			}
		}
		if( ptga == null )
		{
			return null; // it's a NameX or some other non-Cell Name
		}
		if( ptga instanceof PtgRefErr3d )        // 20080228 KSC: return Exception rather than null upon Deleted Named Range
		{
			throw new CellNotFoundException( "Named Range " + getName() + " has been deleted or it's referenced cell is invalid" );    // JM - why not return loc? 'Cause it's deleted!!! :) // 20071203 KSC
		}
		if( ptga instanceof PtgArea3d )
		{
			return ptga.getLocation();    //20080214 KSC: returns correct string
		}
		if( ptga instanceof PtgRef3d )
		{
			return ptga.getLocation();// +":"+ ptga.getLocation();; // BAD -- returns a 2d range for a 1d ref
		}
		return ptga.toString();    // PtgMemFunc ...
	}

	/**
	 * Is this a string referencing Name?
	 *
	 * @return
	 */
	public boolean isStringReference()
	{
		if( (expression.size() == 1) && (expression.get( 0 ) instanceof PtgStr) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Essentially a wrapped setLocation call that is handled for initialization of names
	 * from OOXML files.   Our parsing is not up to snuff for some more complex name records,
	 * and this is a workaround until we are able to handle those expressions
	 * <p/>
	 * TODO: parse complex expressions better and get rid of this method
	 *
	 * @param xpression
	 */
	public void initializeExpression( String xpression )
	{
		try
		{
			setLocation( xpression, false );
		}
		catch( FunctionNotSupportedException e )
		{
			log.warn( "Unable to parse Name record expression: " + xpression );
			cachedOOXMLExpression = xpression;
		}
	}

	public void setLocation( String newloc )
	{
		setLocation( newloc, true );
	}

	/**
	 * Set the location for the first ptg in the expression
	 *
	 * @param newloc
	 * @param clearAffectedCells true if should clear ptga formula cached vals
	 * @throws FunctionNotSupportedException TODO
	 */
	public void setLocation( String newloc, boolean clearAffectedCells ) throws FunctionNotSupportedException
	{
		if( newloc.indexOf( "{" ) > -1 )
		{
			throw new FunctionNotSupportedException( "Unable to parse string expression for name record " + newloc );
		}
		if( ptga == null )
		{
			initPtga();
		}
		if( ptga != null )
		{
			((PtgRef) ptga).removeFromRefTracker();
		}
//	    ptga = new PtgArea3d(false);
		try
		{// can be a named value constant
			Formula f = new Formula();
			f.setWorkBook( wkbook );
			Stack ptgs = FormulaParser.getPtgsFromFormulaString( f, newloc );
			if( ptgs.size() == 1 )
			{    // usual case of 1 ref or 1 PtgMemFunc (complex expression)
				ptga = (Ptg) ptgs.pop();
				//  KSC: memory usage changes: now parent rec nec. for reference tracking; if change must update
				if( (ptga instanceof PtgRef) && ((PtgRef) ptga).getUseReferenceTracker() )
				{
					((PtgRef) ptga).updateInRefTracker( this );
				}
				ptga.setParentRec( this );
			}
			else
			{    // less common case of a formula expression
				expression = ptgs;
				ptga = null;    // otherwise will overwrite expression - see below for handling
			}
		}
		catch( Exception e )
		{    // usually some #REF! error
			log.warn( "Name.setLocation: Error processing location " + e.toString() );
		}
		if( ptga instanceof PtgRef )
		{    // ensure that references are absolute
			((PtgRef) ptga).setColRel( false );
			((PtgRef) ptga).setRowRel( false );
			// TODO: get PtgMemFunc's components and ensure references are absolute
			// clear affected cells so that formulas which reference the named range get recalced with the new value(s)
			if( clearAffectedCells )
			{    // will only be false when initializing an OOXML workbook, which does not need to clear affected cells since it's initializing
				try
				{
					BiffRec[] b = ((PtgRef) ptga).getRefCells();
					for( BiffRec aB : b )
					{
						if( aB != null )
						{
							getWorkBook().getRefTracker().clearAffectedFormulaCells( aB );
						}
					}
				}
				catch( NullPointerException e )
				{
				}    // if cells aren't present ...
			}
		}
		if( ptga != null )
		{
			expression = new Stack();
			expression.add( ptga );    // update expression with new Ptg -- assume there's only 1 ptg!!
			try
			{
				externsheet.addPtgListener( (IxtiListener) ptga );
			}
			catch( Exception e )
			{
				; // ptg constants will fail here, ptgNames ...
			}
		}
		else
		{
			initPtga();    // will calculate and set ptga to
		}
		updatePtgs();
	}

	// remove the name
	// why the boolean - NR
	@Override
	public boolean remove( boolean b )
	{
		boolean ret = super.remove( true );
		wkbook.removeName( this );
		return ret;
	}

	@Override
	public void setWorkBook( WorkBook b )
	{
		super.setWorkBook( b );
	}

	/**
	 * set the Externsheet rec
	 */
	void setExternsheet( Externsheet e ) throws WorkSheetNotFoundException
	{
		externsheet = e;
		if( e == null )
		{
			externsheet = wkbook.getExternSheet( true );
		}
	}

	/**
	 * Initializes ptga.
	 * <p/>
	 * Seems to be an init method to make this
	 */
	void initPtga()
	{
		if( expression == null )
		{
			init();
		}
		if( (expression == null) || (expression.size() == 0) )
		{
			return;
		}
		Ptg p;
		//if the usual case of 1 reference-type (area, ref, memfunc...) ptg:
		if( expression.size() == 1 )
		{
			p = (Ptg) expression.get( 0 );
			if( p.getIsReference() )
			{
				ptga = p;
			}
		}
		else
		{ // otherwise it's a formula expression
			p = FormulaCalculator.calculateFormulaPtg( expression );
			if( p.getIsReference() )
			{
				ptga = p;
			}
		}
/*		
		if (expression != null && expression.size() > 0){
			// this may be an invalid rec
			for(int t=0;t<expression.size();t++){
				Ptg p = (Ptg) expression.get(t);   
				p.setParentRec(this);
				if (p.getIsReference())	 
					ptga= p;
				// otherwise it's a constant named range expression
			}        
		}
*/
	}

	/**
	 * set the Externsheet reference
	 * for any associated PtgArea3d's
	 */
	public void setExternsheetRef( int x )
	{
		// TODO: this doesn't account for formula expressions ...
		for( Object anExpression : expression )
		{
			Ptg p = (Ptg) anExpression;
			if( p instanceof PtgArea3d )
			{
					log.debug( "PtgArea3d encountered in Ai record." );
				PtgArea3d ptg3d = (PtgArea3d) p;
				ptg3d.setIxti( (short) x );
				ptga = ptg3d;
			}
			if( p instanceof PtgRef3d )
			{
					log.debug( "PtgRef3d encountered in Ai record." );
				PtgRef3d ptg3d = (PtgRef3d) p;
				ptg3d.setIxti( (short) x );
				ptga = ptg3d;
			}
		}
	}

	/**
	 * update Ptga ixti for moved/copied worksheets
	 */
	public void updateSheetRefs( String origWorkBookName )
	{
		if( ptga == null )
		{
			initPtga();
		}
		if( ptga instanceof PtgArea3d )
		{    // PtgRef3d,etc
			PtgArea3d p = (PtgArea3d) ptga;
			try
			{
				getWorkBook().getWorkSheetByName( p.getSheetName() );
				ptga.setLocation( ptga.toString() );
			}
			catch( WorkSheetNotFoundException we )
			{
				log.warn( "External References Not Supported:  UpdateSheetReferences: External Worksheet Reference Found: " + p.getSheetName() );
				p.setExternalReference( origWorkBookName );
			}
		}
	}

	/**
	 * set the display name
	 * <p/>
	 * Affects the following byte values:
	 * 7       cch             1       length of name text
	 * 18      rgch            var     name text
	 */
	public void setName( String newname )
	{
		try
		{
			int modnamelen = 0;
			byte[] dta = getData();
			byte[] namebytes;
			boolean isuni = false;
			if( getByteAt( 14 ) == 0x1 )
			{    // 20100604 KSC: added handling of unicode
				namebytes = newname.getBytes( UNICODEENCODING );
				isuni = true;
			}
			else
			{
				namebytes = newname.getBytes( XLSConstants.DEFAULTENCODING );
			}
			modnamelen = namebytes.length;
			int bodlen = dta.length;
			bodlen -= cch;
			bodlen += modnamelen;
			byte[] newbytes = new byte[bodlen];
			System.arraycopy( dta, 0, newbytes, 0, 15 );
			System.arraycopy( namebytes, 0, newbytes, 15, modnamelen );
			System.arraycopy( dta, cch + 15, newbytes, modnamelen + 15, (dta.length - (cch + 15)) );
			cch = (byte) modnamelen;
			if( !isuni )
			{
				newbytes[3] = cch;
			}
			else
			{
				newbytes[3] = (byte) (cch / 2);
			}
			setData( newbytes );
			rgch = newname;
			// search for dreferenced names to rehook up
		}
		catch( UnsupportedEncodingException e )
		{
			log.warn( "UnsupportedEncodingException in setting NamedRange name: " + e );
		}
	}

	/**
	 * get the display name
	 */
	public String getName()
	{
		return rgch;
	}

	/**
	 * return the case-insensitive version of the display name
	 *
	 * @return
	 */
	public String getNameA()
	{
		return rgch.toUpperCase();    // Case-insensitive
	}

	int calc_id = 1;

	/**
	 * return the calculated value of this Name
	 * if it contains a parsed Expression (Formula)
	 *
	 * @return
	 * @throws FunctionNotSupportedException
	 */
	public Object getCalculatedValue() throws FunctionNotSupportedException
	{
		return FormulaCalculator.calculateFormula( expression );

	}

	/**
	 * get the definition text
	 * the definition is stored
	 * in Excel parsed format
	 */
	String getDefinition()
	{
		if( ptga == null )
		{
			initPtga();
		}
		return ptga.getString();
		/*
		StringBuffer sb = new StringBuffer();
		Ptg[] ep = new Ptg[expression.size()];
		ep = (Ptg[]) expression.toArray(ep);
		for(int t = 0;t<ep.length;t++){
			sb.append(ep[t].getString());   
		}
		return sb.toString();*/
	}

	/**
	 * get the descriptive text
	 */
	String getDescription()
	{
		return rgchDescr;
	}

	/**
	 * do any pre-streaming processing such as expensive
	 * index updates or other deferrable processing.
	 */
	@Override
	public void preStream()
	{
		try
		{
			updatePtgs();
		}
		catch( Exception e )
		{
			log.warn( "problem updating Name record expression for Name:" + getName() );
		}

	}

	/**
	 * Update the record byte array with the modified ptg records
	 */
	public void updatePtgs()
	{
		if( expression == null ) // happens upon init
		{
			return;
		}
		byte[] rkdata = getData();
		int offset = 15 + cch; // the start of the parsed expression
		int sz = offset;
		int sz2 = rkdata.length - (offset + cce);
		cce = 0;
		// add up the size of the expressions
		for( int i = 0; i < expression.size(); i++ )
		{
			Ptg ptg = (Ptg) expression.elementAt( i );
			cce += ptg.getLength();
		}
		sz += cce;
		sz += sz2;
		byte[] updated = new byte[sz];
		System.arraycopy( rkdata, 0, updated, 0, offset );
		byte[] cbytes = ByteTools.shortToLEBytes( cce );
		updated[4] = cbytes[0];
		updated[5] = cbytes[1];
		// 20090317 KSC: added handling for PtgArrays
		boolean hasArray = false;
		byte[] arraybytes = new byte[0];
		for( int i = 0; i < expression.size(); i++ )
		{
			Ptg ptg = (Ptg) expression.elementAt( i );
			byte[] b;
			// 20090317 KSC: added handling for PtgArrays
			if( ptg instanceof PtgArray )
			{
				PtgArray pa = (PtgArray) ptg;
				b = pa.getPreRecord();
				arraybytes = ByteTools.append( pa.getPostRecord(), arraybytes );
				hasArray = true;
			}
			else
			{
				b = ptg.getRecord();
			}
			try
			{
				System.arraycopy( b, 0, updated, offset, b.length/*20071206 KSC: Not necessarily the same ptg.getLength()*/ );
			}
			catch( Exception e )
			{
				log.warn( "setting ExternalSheetValue in Name rec: value: " + ptg.getOpcode() + ": " + e );
			}
			offset = offset + ptg.getLength();
		}
		// 20090317 KSC: added handling for PtgArrays
		if( hasArray )
		{
			updated = ByteTools.append( arraybytes, updated );
		}

		// add the rest if any
		if( sz2 > 0 )
		{
			System.arraycopy( rkdata, (rkdata.length - sz2), updated, offset, sz2 );
		}
		setData( updated );
	}

	@Override
	public String getCellAddress()
	{
		if( getSheet() != null )
		{
			return getSheet() + "!" + getName();
		}
		//try{
		//	return this.getLocation();
		//}catch(Exception e){
		return getName(); // ok
		//}
	}

	/**
	 * Returns an array of ptgs that represent any BiffRec ranges in the formula.
	 * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
	 */
	public Ptg[] getCellRangePtgs() throws FormulaNotFoundException
	{
		if( ptga == null )
		{
			initPtga();
		}
		Ptg[] p = ptga.getComponents();
		if( p == null )    // a single ref
		{
			return new Ptg[]{ ptga };
		}
		return p;
//		return ExpressionParser.getCellRangePtgs(expression);
	}

	/**
	 * Returns the ptg that matches the string location sent to it.
	 * this can either be in the format "C5" or a range, such as "C4:D9"
	 */
	public List getPtgsByLocation( String loc )
	{
		try
		{
			return ExpressionParser.getPtgsByLocation( loc, expression );
		}
		catch( FormulaNotFoundException e )
		{
			log.warn( "updating Chart Series Location failed: " + e );
		}
		return null;
	}

	/**
	 * Set all ptg3ds to the new sheet
	 * <br>Used when copying worksheets ..
	 *
	 * @param newSheet
	 */
	public void updateSheetReferences( Boundsheet newSheet )
	{
		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			if( o instanceof PtgMemFunc )
			{
				PtgMemFunc pmf = (PtgMemFunc) o;
				Stack s = pmf.getSubExpression();
				for( int x = 0; x < s.size(); x++ )
				{
					Object ox = s.elementAt( x );
					if( ox instanceof PtgArea3d )
					{
						PtgArea3d p3d = (PtgArea3d) ox;
						p3d.setReferencedSheet( newSheet );
					}// do we have other types we need to handle here? (within memfunc)
				}
			}
			else if( o instanceof PtgArea3d )
			{// do we have other types we need to handle here? (outside memfunc);
				((PtgArea3d) o).setReferencedSheet( newSheet );
			}
			// do nothing
		}
		updatePtgs();
	}

	/**
	 * Return an array of ptgs that make up this Name record

	 *
	 * @return
	 */
/*	public Ptg[] getComponents(){
		if (ptga==null) 
			initPtga();
		return ptga.getComponents();
/*		
//		if (ptga!=null) { // then this is a location-type ptg (ref or memfunc ...)
			Ptg[] p = new Ptg[expression.size()];
			for (int i=0;i<expression.size();i++){
				p[i] = (Ptg)expression.elementAt(i);
			}
			return p;
//		} else {	// calculate formula - assume ends up as a reference
//			Ptg p= FormulaCalculator.calculateFormulaPtg(this.expression);
//			return new Ptg[] {p};
//		}
 * */
//	}

	/**
	 * locks the Ptg at the specified location
	 */

	public boolean setLocationPolicy( String loc, int l )
	{
		List dx = getPtgsByLocation( loc );
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

	public Externsheet getExternsheet()
	{
		return externsheet;
	}

	public Ptg getPtga()
	{
		if( ptga == null )
		{
			initPtga();
		}
		return ptga;
	}

	public boolean isBuiltIn()
	{
		return builtIn;
	}

	public byte getBuiltInType()
	{
		return builtInType;
	}

	/**
	 * Set this name record as a built in type.  Use the static bytes
	 * to set the built in type.
	 *
	 * @param builtinType
	 */
	public void setBuiltIn( byte builtinType )
	{
		// 20100215 KSC: redo
		grbit |= 0x20;    // set built-in bt
		if( builtinType == _FILTER_DATABASE )    // TODO: is this the only one?
		{
			grbit |= 0x1;    // set hidden bit
		}
		byte[] grbytes = ByteTools.shortToLEBytes( grbit );
		byte[] newData = new byte[16];
		newData[0] = grbytes[0];
		newData[1] = grbytes[1];
		newData[3] = 0x1;            // cch
		newData[15] = builtinType;
		setData( newData );
		init();
	}

	public short getIxals()
	{
		return ixals;
	}

	public void setIxals( short ixals )
	{
		this.ixals = ixals;
		byte[] b = ByteTools.shortToLEBytes( ixals );
		byte[] rkdata = getData();
		rkdata[6] = b[0];
		rkdata[7] = b[1];
		setData( rkdata );
	}

	/**
	 * Return the named range scope (0= workbook, 1 or more= sheet)
	 *
	 * @return Named Range Scope
	 */
	public short getItab()
	{
		return itab;
	}

	public void setItab( short itab )
	{
		this.itab = itab;
		byte[] b = ByteTools.shortToLEBytes( itab );
		getData()[8] = b[0];
		getData()[9] = b[1];
	}

	/**
	 * sets a namestring to a constant (non-reference) value
	 *
	 * @param bk
	 * @param namestr String name
	 * @param value   The expression statement
	 * @param scope   1 based reference to sheet scope, 0 is for workbook (default)
	 */
	public Name( WorkBook bk, String namestr, String value, int scope )
	{
		byte[] bl = PROTOTYPE_NAME_BYTES;
		setData( bl );
		setOpcode( NAME );
		setLength( (short) (bl.length) );
		setItab( (short) scope );
		setWorkBook( bk );
		try
		{
			setExternsheet( bk.getExternSheet() );
		}
		catch( WorkSheetNotFoundException x )
		{
			log.warn( "Name could not reference WorkBook Externsheet." + x.toString() );
		}
		init();
		setName( namestr );
		bk.insertName( this );    // calls addRecord which calls addName
		initializeExpression( value );
	}

	/**
	 * @return Returns the cachedOOXMLExpression.
	 */
	public String getCachedOOXMLExpression()
	{
		return cachedOOXMLExpression;
	}

	/**
	 * @param cachedOOXMLExpression The cachedOOXMLExpression to set.
	 */
	public void setCachedOOXMLExpression( String cachedOOXMLExpression )
	{
		this.cachedOOXMLExpression = cachedOOXMLExpression;
	}

	/**
	 * Set the scope (itab) of this name
	 *
	 * @param newitab
	 * @throws WorkSheetNotFoundException
	 */
	public void setNewScope( int newitab ) throws WorkSheetNotFoundException
	{
		if( itab == 0 )
		{
			getWorkBook().removeLocalName( this );
		}
		else
		{
			getWorkBook().getWorkSheetByNumber( itab - 1 ).removeLocalName( this );
			;
		}
		if( newitab == 0 )
		{
			getWorkBook().addLocalName( this );
		}
		else
		{
			getWorkBook().getWorkSheetByNumber( newitab - 1 ).addLocalName( this );
			;
		}
		setItab( (short) newitab );
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		externsheet = null;
		if( ptga != null )
		{
			if( ptga instanceof PtgRef )
			{
				ptga.close();
			}
			else
			{
				ptga.close();
			}
			ptga = null;
		}
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
		close();
	}
}