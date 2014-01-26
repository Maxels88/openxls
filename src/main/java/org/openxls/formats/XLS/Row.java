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

import org.openxls.ExtenXLS.FormatHandle;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

/**
 * <b>ROW 0x208: Describes a single row on a MS Excel Sheet.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---------------------------------------------------------------
 * 0       rw          2       Row number
 * 2       colMic      2       First defined column in the row
 * 4       colMac      2       Last defined column in the row plus 1
 * 6       miyRw      	2       Row Height in twips (1/20 of a printer's point, or 1/1440 of an inch)
 * 14-0 7FFFH Height of the row, in twips
 * 15 8000H 0 = Row has custom height; 1 = Row has default height
 * 8      irwMac      	2       Optimizing, set to 0
 * 10      reserved    2
 * 12      grBit       2       Option Flags
 * 14      ixfe        2       Index to XF record for row
 * <p/>
 * grib options
 * offset	Bits	mask	name					contents
 * ---------------------------------------------------------------
 * 2-0 	07H 	outlineLevel			Outline level of the row
 * 4 		10H 	fCollapsed 				1 = Outline group starts or ends here (depending on where
 * the outline buttons are located, see WSBOOL record and is fCollapsed
 * 5 		20H 	fHidden					1 = Row is fHidden (manually, or by a filter or outline group)
 * 6 		40H 	altered					1 = Row height and default font height do not match
 * 7 		80H 	fFormatted	1 = Row has explicit default format (fl)
 * 8 		100H 							Always 1
 * 27-16 	0FFF0000H 							If fl = 1: Index to default XF record
 * 28 	10000000H 							1 = Additional space above the row. This flag is set, if the
 * upper border of at least one cell in this row or if the lower
 * border of at least one cell in the row above is formatted with
 * a thick line style. Thin and medium line styles are not taken
 * into account.
 * 29  20000000H 							1 = Additional space below the row. This flag is set, if the
 * lower border of at least one cell in this row or if the upper
 * border of at least one cell in the row below is formatted with
 * a medium or thick line style. Thin line styles are not taken
 * into account.
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 * @see BOUNDSHEET
 * @see INDEX
 * @see Dbcell
 * @see ROW
 * @see Cell
 * @see XLSRecord
 */

public final class Row extends XLSRecord
{
	static final int defaultsize = 16;
	private static final Logger log = LoggerFactory.getLogger( Row.class );
	private static final long serialVersionUID = 6848429681761792740L;
	private short colMic;
	private short colMac;
	private int miyRw;
	private Dbcell dbc;
	private BiffRec firstcell;
	private BiffRec lastcell;
	private boolean fCollapsed;
	private boolean fHidden;
	private boolean fUnsynced;
	private boolean fGhostDirty;
	private boolean fBorderTop;
	private boolean fBorderBottom;
	private boolean fPhonetic;
	private int outlineLevel = 0;

	public Row()
	{
		super();
	}

	public Row( int rowNum, WorkBook book )
	{
		setWorkBook( book );
		setLength( (short) defaultsize );
		setOpcode( ROW );
		byte[] dta = new byte[defaultsize];
		dta[6] = (byte) 0xff;
		dta[13] = 0x1;
		dta[14] = 0xf;
		setData( dta );
		originalsize = defaultsize;
		init();
		setRowNumber( rowNum );
	}

	/**
	 * Get the height of a row
	 */
	public int getRowHeight()
	{
		//	15th bit set= row size is not default
		if( miyRw < 0 )    // not 100% sure of this ...
		{
			miyRw = (miyRw + 1) * -1;
		}
		return miyRw;
	}

	/**
	 * Set the height of a row in twips (1/20th of a point)
	 */
	public void setRowHeight( int x )
	{
		log.debug( "Updating Row Height: " + getRowNumber() + " to: " + x );
		fUnsynced = true;  // set bit 6 = row height and default font DO NOT MATCH
		updateGrbit();
		// 10      miyRw       2       Row Height
		byte[] rw = ByteTools.shortToLEBytes( (short) x );
		System.arraycopy( rw, 0, getData(), 6, 2 );
		miyRw = x;
	}

	/**
	 * Get a cell from the row
	 *
	 * @throws CellNotFoundException
	 */
	public BiffRec getCell( short d ) throws CellNotFoundException
	{
		return getSheet().getCell( getRowNumber(), d );
	}

	/**
	 * get a collection of cells in column-based order	*
	 */
	public Collection<BiffRec> getCells()
	{
		try
		{
			return getSheet().getCellsByRow( getRowNumber() );
		}
		catch( CellNotFoundException e )
		{
			// no cells in this row
		}
		return new ArrayList<>();
	}

	/**
	 * set the position of the ROW on the Worksheet
	 */
	@Override
	public void setRowNumber( int n )
	{
		log.debug( "Updating Row Number: " + getRowNumber() + " to: " + n );
		rw = n;
		byte[] rwb = ByteTools.shortToLEBytes( (short) rw );
		System.arraycopy( rwb, 0, getData(), 0, 2 );
	}

	/**
	 * get the cells as an array.  Needed when
	 * operations will be used on the cell array causing concurrentModificationException
	 * problems on the TreeMap collection
	 *
	 * @return
	 */
	public Object[] getCellArray()
	{
		Collection<BiffRec> cells = getCells();
		Object[] br = new Object[cells.size()];
		br = cells.toArray();
		return br;
	}

	/**
	 * get the position of the ROW on the Worksheet
	 */
	@Override
	public int getRowNumber()
	{
		if( rw < 0 )
		{
			int rowi = rw * -1;
			rw = MAXROWS - rowi;
		}
		return rw;
	}

	public int getNumberOfCells()
	{
		return getCells().size();
	}

	/**
	 * Return an ordered array of the BiffRecs associated with this row.
	 * <p/>
	 * This includes child records, and other non-cell-associated records
	 * that should be in the row block (such as Formula Shrfmlas and Arrays)
	 *
	 * @int outputId = random id passed in that is specific to a worksheets output.  Allows tracking what
	 * internal records (such as shared formulas) have been written already
	 */
	public List getValRecs( int outputId )
	{
		ArrayList v = new ArrayList();
		Collection cx = getCells();
		Iterator it = cx.iterator();
		while( it.hasNext() )
		{
			BiffRec br = (BiffRec) it.next();
			v.add( br );
			if( br instanceof Formula )
			{
				br.preStream();    // must do now so can ensure internal records are properly set
				Collection itx = ((Formula) br).getInternalRecords();
				BiffRec[] brints = (BiffRec[]) itx.toArray( new BiffRec[itx.size()] );
				for( BiffRec brint : brints )
				{
					//Don't allow dupes!
					if( !v.contains( brint ) )
					{
						v.add( brint );
					}
				}
			}
			if( br.getHyperlink() != null )
			{
				v.add( br.getHyperlink() );
			}
		}
		return v;
	}

	/**
	 * iterate and get the record index of the last val record
	 */
	public int getLastRecIndex()
	{
		if( lastcell != null )
		{
			return lastcell.getRecordIndex();
		}
		return getRecordIndex(); // empty row
	}

	/**
	 * sets or clears the Unsynced flag
	 * <br>The Unsynched flag is true if the row height is manually set
	 * <br>If false, the row height should auto adjust when necessary
	 *
	 * @param bUnsynced
	 */
	public void setUnsynched( boolean bUnsynced )
	{
		fUnsynced = bUnsynced;
		updateGrbit();
	}

	/**
	 * This flag determines if the row has been formatted.
	 * If this flag is not set, the XF reference will not affect the row.
	 * However, if it's true then the row will be formatted according to
	 * the XF ref.
	 *
	 * @return
	 */
	public boolean getExplicitFormatSet()
	{
		return fGhostDirty;
	}

	/**
	 * return the min/max column for this row
	 *
	 * @return
	 */
	public int[] getColDimensions()
	{
		return new int[]{ colMic, colMac };
	}

	/**
	 * update the col indexes
	 */
	public void updateColDimensions( short col )
	{
		if( col > MAXCOLS )
		{
			return;
		}
		byte[] cl = null;
		if( col < colMic )
		{
			colMic = col;
			cl = ByteTools.shortToLEBytes( colMic );
			System.arraycopy( cl, 0, getData(), 2, 2 );
		}
		if( col > colMac )
		{
			colMac = col;
			colMac = ++col;
			cl = ByteTools.shortToLEBytes( colMac );
			System.arraycopy( cl, 0, getData(), 4, 2 );
		}
	}

	public String toString()
	{
		StringBuffer celladdrs = new StringBuffer();
		Collection cx = getCells();
		Iterator it = cx.iterator();
		while( it.hasNext() )
		{
			celladdrs.append( "{" );
			celladdrs.append( it.next().toString() );
			celladdrs.append( "}" );
		}
		return String.valueOf( getRowNumber() + celladdrs.toString() );
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		dbc = null;
		firstcell = null;
		lastcell = null;
		setWorkBook( null );
		setSheet( null );
	}

	public void setHeight( int twips )
	{
		if( (twips < 2) || (twips > 8192) )
		{
			throw new IllegalArgumentException( "twips value " + String.valueOf( twips ) + " is out of range, must be between 2 and 8192 inclusive" );
		}

		miyRw = twips;
		fUnsynced = true;
	}

	public void clearHeight()
	{
		fUnsynced = false;
	}

	/**
	 * Returns the Outline level (depth) of the row
	 *
	 * @return
	 */
	public int getOutlineLevel()
	{
		return outlineLevel;

	}

	/**
	 * Set the Outline level (depth) of the row
	 *
	 * @param x
	 */
	public void setOutlineLevel( int x )
	{
		outlineLevel = x;
		getSheet().getGuts().setRowGutterSize( 10 + (10 * x) );
		getSheet().getGuts().setMaxRowLevel( x + 1 );
		updateGrbit();
//		implement bit masking set on grbit
	}

	/**
	 * Returns whether the row is collapsed
	 *
	 * @return
	 */
	public boolean isCollapsed()
	{
		return fCollapsed;
	}

	/**
	 * Set whether the row is fCollapsed
	 * hides all contiguous rows with the same outline level
	 *
	 * @param b
	 */
	public void setCollapsed( boolean b )
	{
		fCollapsed = b;
		fHidden = b;
		boolean keepgoing = true;
		int counter = 1;
		while( keepgoing )
		{
			Row r = getSheet().getRowByNumber( getRowNumber() + counter );
			if( (r != null) && (r.getOutlineLevel() == getOutlineLevel()) )
			{
				r.setHidden( b );
			}
			else
			{
				keepgoing = false;
			}
			counter++;
		}
		counter = 1;
		keepgoing = true;
		while( keepgoing )
		{
			Row r = getSheet().getRowByNumber( getRowNumber() - counter );
			if( (r != null) && (r.getOutlineLevel() == getOutlineLevel()) )
			{
				r.setHidden( b );
			}
			else
			{
				keepgoing = false;
			}
			counter++;
		}
		updateGrbit();
		// implement bit masking set on grbit
	}

	@Override
	public void init()
	{
		super.init();
		getData();

		// get the number of the row
		rw = ByteTools.readUnsignedShort( getByteAt( 0 ), getByteAt( 1 ) );
		colMic = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		colMac = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		miyRw = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );

		// bytes 8 - 11 are reserved		
		byte byte12 = getByteAt( 12 );

		/**
		 *     A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row.
		 B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
		 C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
		 D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
		 E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
		 F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
		 */
		outlineLevel = (byte12 & 0x7);
		fCollapsed = (byte12 & 0x10) != 0;    // ?? 0x8 ??
		fHidden = (byte12 & 0x20) != 0;
		fUnsynced = (byte12 & 0x40) != 0;
		fGhostDirty = (byte12 & 0x80) != 0;
		/**/
		// byte 13 is reserved
		byte byte15 = getByteAt( 15 );
		byte byte14 = getByteAt( 14 );
		if( fGhostDirty )
		{ // then explicit ixfe set
			// The low-order byte is sbyte 14. The low-order nybble of the
			// high-order byte is stored in the high-order nybble of byte 15.
			ixfe = ((byte14 & 0xFF) | (((byte15 & 0xFF) << 8) & 0xFFF)); // 12 bits
			if( (ixfe < 0) || (ixfe > getWorkBook().getNumXfs()) )
			{    // KSC: TODO: ixfe calc is wrong ...?
				ixfe = 15;//this.getWorkBook().getDefaultIxfe();
				fGhostDirty = false;
			}
		}
		else
		{
			ixfe = 15;//this.getWorkBook().getDefaultIxfe();
		}

		/** f
		 * from excel documentation:
		 fBorderTop=
		 G - fExAsc (1 bit): A bit that specifies whether any cell in the row has a thick top border, or any cell in the row directly above the current row has a thick bottom border.
		 Thick borders are specified by the following enumeration values from BorderStyle: THICK and DOUBLE.
		 fBorderBottom=
		 H - fExDes (1 bit): A bit that specifies whether any cell in the row has a medium or thick bottom border, or any cell in the row directly below the current row has a medium or thick top border.
		 Thick borders are previously specified. Medium borders are specified by the following enumeration values from BorderStyle: MEDIUM, MEDIUMDASHED, MEDIUMDASHDOT, MEDIUMDASHDOTDOT, and SLANTDASHDOT.
		 */
		fBorderTop = (byte15 & 0x10) != 0;
		fBorderBottom = (byte15 & 0x20) != 0;
		fPhonetic = (byte15 & 0x40) != 0;
	}

	/**
	 * Returns whether the row is hidden
	 * TODO:  same issue as setHidden above!
	 *
	 * @return
	 */
	public boolean isHidden()
	{
		return fHidden;
	}

	@Override
	public void preStream()
	{
		if( getSheet() != null )
		{
			updateColDimensions( colMac );
			updateColDimensions( colMic );
		}
		else
		{
			log.warn( "Missing Boundsheet in Row.prestream for Row: " + getRowNumber() + getCellAddress() );
		}

		byte[] data = new byte[16];
		data[0] = (byte) (rw & 0x00FF);
		data[1] = (byte) ((rw & 0xFF00) >>> 8);
		data[2] = (byte) (colMic & 0x00FF);
		data[3] = (byte) ((colMic & 0xFF00) >>> 8);
		data[4] = (byte) (colMac & 0x00FF);
		data[5] = (byte) ((colMac & 0xFF00) >>> 8);
		data[6] = (byte) (miyRw & 0x00FF);
		data[7] = (byte) ((miyRw & 0xFF00) >>> 8);
		// bytes 8 - 11 are reserved

		/**
		 *     A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row.
		 B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
		 C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
		 D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
		 E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
		 F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
		 */
		if( outlineLevel != 0 )
		{
			data[12] |= outlineLevel;
		}
		if( fCollapsed )
		{
			data[12] |= 0x10;    // 0x8 ???
		}
		if( fHidden )
		{
			data[12] |= 0x20;
		}
		if( fUnsynced )
		{
			data[12] |= 0x40;
		}
		if( fGhostDirty )
		{
			data[12] |= 0x80;
		}
		/**/
		// byte 13 is reserved
		data[13] = 1;

		// The low-order byte is byte 14. The low-order nybble of the
		// high-order byte is stored in the high-order nybble of byte 15.
		data[14] = (byte) (ixfe & 0x00FF);
		data[15] = (byte) ((ixfe >> 8));    //& 0x0F00) >>> 4);
		if( fBorderTop )
		{
			data[15] |= 0x10;
		}
		if( fBorderBottom )
		{
			data[15] |= 0x20;
		}
		if( fPhonetic )
		{
			data[15] |= 0x40;
		}
		// byte 15 bit 0x01 is reserved

		setData( data );
	}

	/**
	 * Set whether the row is fHidden
	 *
	 * @param b
	 */
	public void setHidden( boolean b )
	{
		fHidden = b;
		updateGrbit();
//		implement bit masking set on grbit
	}

	/**
	 * true if row height has been altered from default
	 * i.e. set manually
	 *
	 * @return
	 */
	public boolean isAlteredHeight()
	{
		return fUnsynced;
	}

	/**
	 * Removes the format currently applied to this row, if any.
	 */
	public void clearIxfe()
	{
		ixfe = 0;
		fGhostDirty = false;
	}

	/**
	 * Returns whether a format has been set for this row.
	 */
	public boolean hasIxfe()
	{
		return fGhostDirty;
	}

	/**
	 * returns true if there is a thick bottom border set on the row
	 */
	public boolean getHasThickTopBorder()
	{
		if( !fGhostDirty )
		{
			return false;
		}
		if( fBorderTop )
		{
			try
			{
				int bs = getXfRec().getTopBorderLineStyle();
				return ((bs == FormatHandle.BORDER_DOUBLE) || (bs == FormatHandle.BORDER_THICK));
			}
			catch( Exception e )
			{
				;
			}
		}
		return false;
	}

	/**
	 * sets this row to have a thick top border
	 */
	public void setHasThickTopBorder( boolean hasBorder )
	{
		fBorderTop = hasBorder;
		if( hasBorder )
		{
			FormatHandle fh = new FormatHandle( null, getXfRec() );
			fh.setTopBorderLineStyle( FormatHandle.BORDER_THICK );
			ixfe = fh.getFormatId();
			myxf = null;    // reset
		}
		fGhostDirty = true;
	}

	/**
	 * Additional space above the row. This flag is set, if the
	 * upper border of at least one cell in this row or if the lower
	 * border of at least one cell in the row above is formatted with
	 * a thick line style. Thin and medium line styles are not taken
	 * into account.
	 */
	public boolean getHasAnyThickTopBorder()
	{
		return fBorderTop;
	}

	/**
	 * flags this row to have at least one cell that has a thick top border
	 * <p/>
	 * For internal use only
	 *
	 * @param hasBorder
	 */
	public void setHasAnyThickTopBorder( boolean hasBorder )
	{
		fBorderTop = hasBorder;
	}

	/**
	 * returns true if there is a thick bottom border set on the row
	 */
	public boolean getHasThickBottomBorder()
	{
		if( !fGhostDirty )
		{
			return false;
		}
		if( fBorderBottom )
		{
			try
			{
				int bs = getXfRec().getBottomBorderLineStyle();
				return ((bs == FormatHandle.BORDER_DOUBLE) || (bs == FormatHandle.BORDER_THICK));
			}
			catch( Exception e )
			{
				;
			}
		}
		return fBorderBottom;
	}

	/**
	 * sets this row to have a thick bottom border
	 */
	public void setHasThickBottomBorder( boolean hasBorder )
	{
		fBorderBottom = hasBorder;
		if( hasBorder )
		{
			FormatHandle fh = new FormatHandle( null, getXfRec() );
			fh.setBottomBorderLineStyle( FormatHandle.BORDER_THICK );
			ixfe = fh.getFormatId();
			myxf = null; // reset
		}
		fGhostDirty = true;
	}

	/**
	 * Additional space below the row. This flag is set, if the
	 * lower border of at least one cell in this row or if the upper
	 * border of at least one cell in the row below is formatted with
	 * a medium or thick line style. Thin line styles are not taken
	 * into account.
	 */
	public boolean getHasAnyBottomBorder()
	{
		return fBorderBottom;
	}

	/**
	 * flags this row to have at least one cell that has a thick bottom border
	 * <p/>
	 * For internal use only
	 *
	 * @param hasBorder
	 */
	public void setHasAnyThickBottomBorder( boolean hasBorder )
	{
		fBorderBottom = hasBorder;
	}

	/**
	 * get the Dbcell record which contains the
	 * cell offsets for this row.
	 * <p/>
	 * needed in computing new INDEX offset values
	 */
	Dbcell getDBCell()
	{
		return dbc;
	}

	/**
	 * set the Dbcell record which contains the
	 * cell offsets for this row.
	 * <p/>
	 * needed in computing new INDEX offset values
	 */
	public void setDBCell( Dbcell d )
	{
		dbc = d;
	}

	/**
	 * add a cell to the Row.  Instead of using the full
	 * cell address as the treemap identifier, just use the column.
	 * this allows the natural ordering of the treemap to work to our
	 * advantage on output, ordering cells from lowest to highest col.
	 */
	void addCell( BiffRec c )
	{
		c.setRow( this );

		/*
		 * I'm not clear on this operation, it seems as if we add blank records, this just
		 * applies a format to the cell rather than actually replacing the valrec?

		if (c.getOpcode()!=MULBLANK) {	// KSC: Added
			BiffRec existing = (BiffRec)cells.get(Short.valueOf(cellCol));
			if( existing != null){
	            if (this.getWorkBook().getFactory().iscompleted()) {
	    		    if((c instanceof Blank)) {
	    		    	existing.setIxfe(c.getIxfe());
	    		        return;
	    		    }else {
	    		        cells.remove(Short.valueOf(cellCol));
	                    c.setRow(this);
	                    cells.put(Short.valueOf(cellCol), c);
	                    this.lastcell = c;
	    		    }
	            }
			}else {
	    		c.setRow(this);
	    		cells.put(Short.valueOf(cellCol), c);
	    		this.lastcell = c;
	        }
		} else { // expand mulblanks to each referenced cell
		 */

		/* We should be able to handle this with cellAddressible, hopefully.
			short colFirst= ((Mulblank)c).colFirst;
			short colLast= ((Mulblank) c).colLast;
			for (short i= colFirst; i <= colLast; i++) {
                cells.put(Short.valueOf(i), c);
                this.lastcell = c;
			}
		}
		 */
	}

	void removeCell( BiffRec c )
	{
		getSheet().removeCell( c );
	}

	/**
	 * Applies the format with the given ID to this row.
	 *
	 * @param ixfe the format ID. Must be between 0x0 and 0xFFF.
	 * @throws IllegalArgumentException if the given format ID cannot be
	 *                                  encoded in the 1.5 byte wide field provided for it
	 */
	@Override
	public void setIxfe( int ixfe )
	{
		if( (ixfe & ~0xFFF) != 0 )
		{
			throw new IllegalArgumentException( "ixfe value 0x" + Integer.toHexString( ixfe ) + " out of range, must be between 0x0 and 0xfff" );
		}
		this.ixfe = ixfe;
		if( ixfe != getWorkBook().getDefaultIxfe() )
		{
			fGhostDirty = true;
		}
	}

	/**
	 * remove cell via column number
	 */
	void removeCell( short c )
	{
		getSheet().removeCell( getRowNumber(), c );
	}

	/**
	 * Get the real max col
	 */
	int getRealMaxCol()
	{
		int collast = 0;
		Iterator cs = getCells().iterator();
		while( cs.hasNext() )
		{
			BiffRec c = (BiffRec) cs.next();
			if( c.getColNumber() > collast )
			{
				collast = c.getColNumber();
			}
		}
		return collast;
	}

	/**
	 * Gets the ID of the format currently applied to this row.
	 *
	 * @return the ID of the current format,
	 * or the default format ID if no format has been applied
	 */
	@Override
	public int getIxfe()
	{
		return fGhostDirty ? ixfe : getWorkBook().getDefaultIxfe();
	}

	int getMaxCol()
	{
		preStream();
		return colMac;
	}

	int getMinCol()
	{
		preStream();
		return colMic;
	}

	/**
	 * Update the internal Grbit based
	 * on values existant in the row
	 */
	private void updateGrbit()
	{
		preStream();
	}

}