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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.formats.LEO.BlockByteConsumer;
import com.extentech.formats.LEO.BlockByteReader;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * <pre>
 * The XLS byte stream is composed of records delimited by a header with type
 * and data length information, followed by a record Body which contains
 *
 * the byte[] data for the record.
 *
 * XLSRecord subclasses provide specific functionality.
 *
 * </pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Index
 * @see Dbcell
 * @see Row
 * @see Cell
 */

public class XLSRecord implements BiffRec, BlockByteConsumer, Serializable, XLSConstants
{
	public boolean isString;
	public boolean isBoolean;
	public boolean isBlank;
	public int offset = 0; // byte stream offset of rec
	public transient Xf myxf;
	public Hlink hyperlink = null;
	protected boolean isContinueMerged = false;
	protected int rw = -1;
	protected short col;
	protected transient Index idx;
	protected Sheet worksheet;
	protected transient WorkBook wkbook;
	protected transient ByteStreamer streamer;
	protected List<Continue> continues;
	int reclen;
	byte[] data;
	boolean isValueForCell;
	boolean isFPNumber;
	boolean isDoubleNumber = false;
	boolean isIntNumber;
	boolean isFormula;
	boolean isReadOnly = false;
	int originalsize = -1;
	int originalOffset = 0;
	/**
	 * ixfe - an unsigned integer that specifies a 0-based index of a cell XF record in the collection of XF records in the globals substream
	 */
	int ixfe = -1;
	Row myrow = null;
	CellRange mergeRange;
	private static final Logger log = LoggerFactory.getLogger( XLSRecord.class );
	private static final long serialVersionUID = -106915096753184441L;
	private short opcode;
	private transient BlockByteReader databuf;
	private transient BlockByteReader encryptedDatabuf;
	private int firstblock;
	private int lastblock;

	public XLSRecord()
	{
	}

	/**
	 * FOR UNIT TESTING ONLY
	 * @param row
	 * @param col
	 */
	XLSRecord( int row, int col )
	{
		rw = row;
		this.col = (short) col;
	}

	@Override
	public short getOpcode()
	{
		return opcode;
	}

	/**
	 * get a new, generic instance of a Record.
	 */
	protected static XLSRecord getPrototype()
	{
		log.warn( "Attempt to get prototype XLSRecord failed.  There is no prototype record defined for this record type." );
		return null;
	}

	@Override
	public void setOpcode( short op )
	{
		opcode = op;
	}

	public boolean shouldEncrypt()
	{
		return true;
	}

	@Override
	public void setHyperlink( Hlink hl )
	{
		hyperlink = hl;
	}

	/**
	 * get the int val of the type for the valrec
	 */
	public int getCellType()
	{
		if( isBlank )
		{
			return TYPE_BLANK;
		}
		if( isString )
		{
			return TYPE_STRING;
		}
		if( isDoubleNumber )
		{
			return TYPE_DOUBLE;
		}
		if( isFPNumber )
		{
			return TYPE_FP;
		}
		if( isIntNumber )
		{
			return TYPE_INT;
		}
		if( isFormula )
		{
			return TYPE_FORMULA;
		}
		if( isBoolean )
		{
			return TYPE_BOOLEAN;
		}
		return -1;
	}

	@Override
	public Formula getFormulaRec()
	{
		if( this instanceof Formula )
		{
			((Formula) this).populateExpression();
			return (Formula) this;
		}
		return null;
	}

	public String getRecDesc()
	{
		String ret = "";
		String name = getClass().getSimpleName();

		// record name
		ret += (name.equals( "XLSRecord" ) ? "unknown" : name.toUpperCase());
		// hex record number
		ret += " (" + Integer.toHexString( opcode ).toUpperCase() + "h)";
		// stream offset
		ret += " at " + Integer.toHexString( offset ).toUpperCase() + "h";
		// size
		ret += " length " + Integer.toHexString( reclen ).toUpperCase() + "h";
		// file offset
		ret += " file " + ((databuf == null) ? "no file" : databuf.getFileOffsetString( offset, reclen ));
		// cell address, if applicable
		if( isValueForCell() )
		{
			ret += " cell " + getCellAddress();
		}

		return ret;
	}

	/**
	 * clone a record
	 */
	@Override
	public Object clone()
	{
		try
		{
			String cn = getClass().getName();
			XLSRecord rec = (XLSRecord) Class.forName( cn ).newInstance();
			byte[] inb = getBytes();
			rec.setData( inb );
			rec.streamer = streamer;
			rec.setWorkBook( getWorkBook() );
			rec.setOpcode( getOpcode() );
			rec.setLength( getLength() );
			rec.setSheet( getSheet() );    //20081120 KSC: otherwise may set sheet incorrectly in init
			rec.init();
			return rec;
		}
		catch( Exception e )
		{
			log.info( "cloning XLSRecord " + getCellAddress() + " failed: " + e );
		}
		return null;
	}

	@Override
	public void setRow( Row r )
	{
		myrow = r;
	}

	public String toString()
	{
		return getRecDesc();
	}

	/**
	 * get the row of this cell
	 */
	@Override
	public Row getRow()
	{
		if( myrow == null )
		{
			myrow = worksheet.getRowByNumber( rw );
		}
		return myrow;
	}

	/**
	 * return whether this is a numeric type
	 */
	public boolean isNumber()
	{
		if( opcode == RK )
		{
			return true;
		}
		if( opcode == NUMBER )
		{
			return true;
		}
		return false;
	}

	@Override
	public Xf getXfRec()
	{
		if( myxf == null )
		{
			if( (ixfe > -1) && (ixfe < wkbook.getNumXfs()) )
			{
				myxf = wkbook.getXf( ixfe );
			}
		}
		return myxf;
	}

	/**
	 * return whether this is a formula record
	 */
	public boolean isFormula()
	{
		return isFormula;
	}

	/**
	 * returns the existing font record
	 * for this Cell
	 */
	@Override
	public Font getFont()
	{
		WorkBook b = getWorkBook();
		if( b == null )
		{
			return null;
		}
		getXfRec();
		if( myxf == null )
		{
			return null;
		}
		return myxf.getFont();
	}

	/**
	 * return the real (not just boundsheet) record index of this object
	 */
	public int getRealRecordIndex()
	{
		return streamer.getRealRecordIndex( this );
	}

	/**
	 * @return
	 */
	@Override
	public CellRange getMergeRange()
	{
		return mergeRange;
	}

	/**
	 * returns the cell address in int[] {row, col} format
	 */
	public int[] getIntLocation()
	{
		return new int[]{ rw, col };
	}

	/**
	 * @param range
	 */
	@Override
	public void setMergeRange( CellRange range )
	{
		mergeRange = range;
	}

	/**
	 * return the cell address with sheet reference
	 * eg Sheet!A12
	 *
	 * @return String
	 */
	public String getCellAddressWithSheet()
	{
		if( getSheet() != null )
		{
			return getSheet().getSheetName() + "!" + getCellAddress();
		}
		return getCellAddress();
	}

	/**
	 * Removes this BiffRec from the WorkSheet
	 *
	 * @param whether to nullify this Cell
	 */
	@Override
	public boolean remove( boolean nullme )
	{
		if( (worksheet != null) && isValueForCell )
		{
			getSheet().removeCell( this );
		}

		if( streamer != null )
		{
			streamer.removeRecord( this );
		}

		worksheet = null;
		return true;
	}

	/**
	 * get the int val of the type for the valrec
	 */
	@Override
	public Object getInternalVal()
	{
		try
		{
			switch( getCellType() )
			{
				case TYPE_BLANK:
					return getStringVal(); //essentially return "";

				case TYPE_STRING:
					return getStringVal();

				case TYPE_FP: // always use Doubles to avoid loss of precision... see:
					// details http://stackoverflow.com/questions/916081/convert-float-to-double-without-losing-precision
					return getDblVal();

				case TYPE_DOUBLE:
					return getDblVal();

				case TYPE_INT:
					return getIntVal();

				case TYPE_FORMULA:
					// OK this is broken, obviously we need to return a calced Object
					Object obx = ((Formula) this).calculateFormula();
					return obx;
//					return getStringVal();

				case TYPE_BOOLEAN:
					return getBooleanVal();

				default:
					return null;
			}
		}
		catch( Exception e )
		{
			return null;
		} // should never happen here...
	}

	/**
	 * set the Formatting for this BiffRec from the pattern
	 * match.
	 * <p/>
	 * case insensitive pattern match is performed...
	 */
	@Override
	public String getFormatPattern()
	{
		if( myxf == null )
		{
			return "";
		}
		return myxf.getFormatPattern();
	}

	/**
	 * resets the cache bytes so they do not take up space
	 */
	public void resetCacheBytes()
	{
		//  data = null;
	}

	// Methods from BlockByteConsumer //

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		if( databuf != null )
		{
			databuf.clear();
			databuf = null;
		}
		idx = null;
		worksheet = null;
		myxf = null;
		wkbook = null;
		streamer = null;
		if( continues != null )
		{
			for( Continue aContinue : continues )
			{
				aContinue.close();
			}
			continues.clear();
		}
		if( hyperlink != null )
		{
			hyperlink.close();
			hyperlink = null;
		}
		mergeRange = null;
		if( myrow != null )
		{
			myrow = null;
		}
		mergeRange = null;

	}

	/**
	 * Set the relative position within the data
	 * underlying the block vector represented by
	 * the BlockByteReader.
	 * <p/>
	 * In other words, this is the relative position
	 * used by the BlockByteReader to offset the Consumer's
	 * read position within the collection of Data Blocks.
	 * <p/>
	 * This may be an offset relative to the data in a file,
	 * or within a Storage contained in a file.
	 * <p/>
	 * The Workbook Storage for example will contain a non-contiguous
	 * collection of Blocks containing data from any number of
	 * positions in a file.
	 * <p/>
	 * This collection forms a contiguous span of bytes comprising
	 * an XLS Workbook.  The XLSRecords within this span of bytes will
	 * set their relative position within this 'virtual' array.  Thus
	 * the XLSRecord positions are relative to the order of bytes contained
	 * in the Block collection.  The BOF record then is at offset 0 within the
	 * data of the first Block, even though the underlying data of this
	 * first Block may be anywhere on disk.
	 *
	 * @param pos
	 */
	@Override
	public void setOffset( int pos )
	{
		if( originalOffset < 1 )
		{
			originalOffset = pos;
		}
		offset = pos;
	}

	/**
	 * Get the color table for the associated workbook.  If the workbook is null
	 * then the default colortable will be returned.
	 */
	public java.awt.Color[] getColorTable()
	{
		try
		{
			return getWorkBook().getColorTable();
		}
		catch( Exception e )
		{
			return FormatHandle.COLORTABLE;
		}
	}

	/**
	 * Get the relative position within the data
	 * underlying the block vector represented by
	 * the BlockByteReader.
	 *
	 * @return relative position
	 */
	@Override
	public int getOffset()
	{
		if( data == null )
		{
			return originalOffset;
		}
		return offset;
	}

	/** Get the blocks containing this Consumer's data
	 *
	 * @return
	 */
/*unused: public Block[] getBlocks() {
		return myblocks;
	}*/

	/**
	 * Set the blocks containing this Consumer's data
	 *
	 * @param myblocks
	 */
/* unused: 	public void setBlocks(Block[] myb) {
		myblocks = myb;
	}*/
	protected void initRowCol()
	{
		int pos = 0;
		byte[] bt = getBytesAt( pos, 2 );
		rw = ByteTools.readUnsignedShort( bt[0], bt[1] );

		pos += 2;
		col = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );
	}

	/**
	 * Sets the index of the first block
	 *
	 * @return
	 */
	@Override
	public void setFirstBlock( int i )
	{
		firstblock = i;
	}

	/**
	 * Merge continue data in to the record data array
	 */
	protected void mergeContinues()
	{
		List cx = getContinueVect();
		if( cx != null )
		{
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			// get the main bytes first!!!
			try
			{
				out.write( data );
			}
			catch( IOException e )
			{
				log.warn( "ERROR: parsing record continues failed: " + toString() + ": " + e );
			}
			Iterator it = cx.iterator();
			while( it.hasNext() )
			{
				Continue c = (Continue) it.next();
				byte[] nb = c.getData();
				if( c.getHasGrbit() )
				{    // remove it - happens in continued StringRec ... a rare case
					byte[] newnb = new byte[nb.length - 1];
					System.arraycopy( nb, 1, newnb, 0, newnb.length );
					nb = newnb;
				}
				try
				{
					out.write( nb );
				}
				catch( IOException a )
				{
					log.warn( "ERROR: parsing record continues failed: " + toString() + ": " + a );
				}
			}
			data = out.toByteArray();
		}
		isContinueMerged = true;
	}

	/**
	 * Sets the index of the last block
	 *
	 * @return
	 */
	@Override
	public void setLastBlock( int i )
	{
		lastblock = i;
	}

	/**
	 * Get the value of the record as an Object.
	 * To use the Object, cast it to the native
	 * type for the record.  Ie: a String value
	 * would need to be cast to a String, an Integer
	 * to an Integer, etc.
	 */
	Object getVal()
	{
		return null;
	}

	/**
	 * Returns the index of the first block
	 *
	 * @return
	 */
	@Override
	public int getFirstBlock()
	{
		return firstblock;
	}

	/**
	 * Returns the index of the last block
	 *
	 * @return
	 */
	@Override
	public int getLastBlock()
	{
		return lastblock;
	}

	/**
	 * Dumps this record as a human-readable string.
	 */
	@Override
	public String toHexDump()
	{
		return getRecDesc() + "\n" + ByteTools.getByteDump( getData(), 0 );
	}

	/**
	 * Copy all formatting info from source biffrec
	 *
	 * @see com.extentech.formats.XLS.BiffRec#copyFormat(com.extentech.formats.XLS.BiffRec)
	 */
	@Override
	public void copyFormat( BiffRec source )
	{
		Xf clone = (Xf) source.getXfRec().clone();
		Font fontClone = (Font) source.getXfRec().getFont().clone();

		log.info( source + ":" + source.getXfRec() + ":" + clone );
		int fid = -1;
		int xid = -1;
		// see if we have an equivalent Xf/Font combo
		if( getWorkBook().getFontRecs().contains( fontClone ) )
		{
			fid = getWorkBook().getFontRecs().indexOf( fontClone );
		}
		if( getWorkBook().getXfrecs().contains( clone ) )
		{
			xid = getWorkBook().getXfrecs().indexOf( clone );
		}

		// add the xf/font and set the ixfe
		getWorkBook().addRecord( clone, false );
		getWorkBook().addRecord( fontClone, false );
		clone.setFont( fontClone.getIdx() );
		setXFRecord( clone.getIdx() );

	}

	/** methods from CompatibleVectorHints

	 protected transient int recordIdx = -1;
	 */
	/** provide a hint to the CompatibleVector
	 about this objects likely position.

	 public int getRecordIndexHint(){return recordIdx;}
	 */

	/**
	 * return the record index of this object
	 */
	@Override
	public int getRecordIndex()
	{
		if( streamer == null )
		{
			if( getSheet() != null )    // KSC: Added
			{
				return getSheet().getSheetRecs().indexOf( this );
			}
			return -1;
		}
		return streamer.getRecordIndex( this );
	}

	/**
	 * adds a CONTINUE record to the array of CONTINUE records
	 * for this record, containing all data
	 * beyond the 8224 byte record size limit.
	 */
	@Override
	public void addContinue( Continue c )
	{
		if( continues == null )
		{
			continues = new ArrayList<>();
		}
		continues.add( c );
	}

	/**
	 * remove all Continue records
	 */
	@Override
	public void removeContinues()
	{
		if( continues != null )
		{
			continues.clear();
		}
	}

	@Override
	public List getContinueVect()
	{
		if( continues != null )
		{
			return continues;
		}
		continues = new CompatibleVector();
		return continues;
	}

	/**
	 * returns whether this record has a CONTINUE
	 * record containing data beyond the 8228 byte
	 * record size limit.
	 * <p/>
	 * XLSRecords can have 0 or more CONTINUE records.
	 */
	@Override
	public boolean hasContinues()
	{
		if( continues == null )
		{
			return false;
		}
		if( continues.size() > 0 )
		{
			return true;
		}
		return false;
	}

	/**
	 * set whether this record contains the  value
	 * of the Cell.
	 */
	@Override
	public void setIsValueForCell( boolean b )
	{
		isValueForCell = b;
	}

	/**
	 * associate this record with its Index record
	 */
	@Override
	public void setIndex( Index id )
	{
		idx = id;
	}

	/**
	 * Associate this record with a worksheet.
	 * First checks to see if there is already
	 * a cell with this address.
	 */
	@Override
	public void setSheet( Sheet b )
	{
		worksheet = b;
	}

	/**
	 * get the WorkSheet for this record.
	 */
	@Override
	public Boundsheet getSheet()
	{
		return (Boundsheet) worksheet;
	}

	@Override
	public void setWorkBook( WorkBook wk )
	{
		wkbook = wk;
	}

	@Override
	public WorkBook getWorkBook()
	{
		if( (wkbook == null) && (worksheet != null) )
		{
			wkbook = worksheet.getWorkBook();
		}

		return wkbook;
	}

	/**
	 * set the column
	 */
	@Override
	public void setCol( short i )
	{
		byte[] c = ByteTools.shortToLEBytes( i );
		System.arraycopy( c, 0, getData(), 2, 2 );
		col = i;
	}

	@Override
	public void setRowCol( int[] x )
	{
		setRowNumber( x[0] );
		setCol( (short) x[1] );
	}

	/**
	 * set the row
	 */
	@Override
	public void setRowNumber( int i )
	{
		byte[] r = ByteTools.cLongToLEBytes( i );
		System.arraycopy( r, 0, getData(), 0, 2 );
		rw = i;
	}

	@Override
	public short getColNumber()
	{
		return col;
	}

	@Override
	public int getRowNumber()
	{
		if( rw < 0 )
		{
			int rowi = rw * -1;
			if( wkbook.getIsExcel2007() )
			{
				rw = MAXROWS - rowi;
			}
			else
			{
				rw = MAXCOLS_BIFF8 - rowi;
			}
		}

		return rw;
	}

	/**
	 * get a string address for the
	 * cell based on row and col ie: "H22"
	 */
	@Override
	public String getCellAddress()
	{
		int rownum = rw + 1;
		if( (rownum < 0) && (col >= 0) )
		{ // > 32k and the rows go negative... !
			rownum = MAXROWS + rownum;
		}
		else if( (rownum == 0) && (col >= 0) )
		{ // the very last row...MAXROWS_BIFF8
			rownum = MAXROWS;
		}
		if( (rownum > MAXROWS) || (col < 0) )
		{
			log.warn( "XLSRecord.getCellAddress() Row/Col info incorrect for Cell:" + ExcelTools.getAlphaVal( col ) + String.valueOf( rownum ) );
			return "";
		}
		return ExcelTools.getAlphaVal( col ) + rownum;
	}

	/**
	 * perform record initialization
	 */
	@Override
	public void init()
	{
		if( originalsize == 0 )
		{
			originalsize = reclen;
		}
	}

	/**
	 * get a default "empty" data value for this record
	 */
	@Override
	public Object getDefaultVal()
	{
		if( isDoubleNumber )
		{
			return 0.0;
		}
		if( isFPNumber )
		{
			return 0.0f;
		}
		if( isBoolean )
		{
			return false;
		}
		if( isIntNumber )
		{
			return 0;
		}
		if( isString )
		{
			return "";
		}
		if( isFormula )
		{
			return getFormulaRec().getFormulaString();
		}
		if( isBlank )
		{
			return "";
		}
		return null;
	}

	/**
	 * get the data type name for this record
	 */
	@Override
	public String getDataType()
	{
		if( isValueForCell )
		{
			if( isBlank )
			{
				return "Blank";
			}
			if( isDoubleNumber )
			{
				return "Double";
			}
			if( isFPNumber )
			{
				return "Float";
			}
			if( isBoolean )
			{
				return "Boolean";
			}
			if( isIntNumber )
			{
				return "Integer";
			}
			if( isString )
			{
				return "String";
			}
			if( isFormula )
			{
				return "Formula";
			}
		}
		return null;
	}

	/**
	 * Get the value of the record as a Boolean.
	 * Value must be parseable as a Boolean.
	 */
	@Override
	public boolean getBooleanVal()
	{
		return false;
	}

	/**
	 * Get the value of the record as an Integer.
	 * Value must be parseable as an Integer or it
	 * will throw a NumberFormatException.
	 */
	@Override
	public int getIntVal()
	{
		return (int) Float.NaN;
	}

	/**
	 * Get the value of the record as a Double.
	 * Value must be parseable as an Double or it
	 * will throw a NumberFormatException.
	 */
	@Override
	public double getDblVal()
	{
		return Float.NaN;
	}

	/**
	 * Get the value of the record as a Float.
	 * Value must be parseable as an Float or it
	 * will throw a NumberFormatException.
	 */
	@Override
	public float getFloatVal()
	{
		return Float.NaN;
	}

	/**
	 * Get the value of the record as a String.
	 */
	@Override
	public String getStringVal()
	{
		return null;
	}

	/**
	 * Get the value of the record as a String.
	 */
	@Override
	public String getStringVal( String encoding )
	{
		return null;
	}

	@Override
	public void setStringVal( String v )
	{
		log.error( "Setting String Val on generic XLSRecord, value not held" );
	}

	@Override
	public void setBooleanVal( boolean b )
	{
		log.error( "Setting Boolean Val on generic XLSRecord, value not held" );
	}

	@Override
	public void setIntVal( int v )
	{
		log.error( "Setting int Val on generic XLSRecord, value not held" );
	}

	public void setFloatVal( float v )
	{
		log.error( "Setting float Val on generic XLSRecord, value not held" );
	}

	@Override
	public void setDoubleVal( double v )
	{
		log.error( "Setting Double Val on generic XLSRecord, value not held" );
	}

	/**
	 * do any pre-streaming processing such as expensive
	 * index updates or other deferrable processing.
	 */
	@Override
	public void preStream()
	{
		// override in sub-classes
	}

	/**
	 * set the XF (format) record for this rec
	 */
	@Override
	public void setXFRecord()
	{
		if( wkbook == null )
		{
			return;
		}
		if( (ixfe > -1) && (ixfe < wkbook.getNumXfs()) )
		{
			if( (myxf == null) || (myxf.tableidx != ixfe) )
			{
				myxf = wkbook.getXf( ixfe );
				myxf.incUseCount();
			}
		}
	}

	/**
	 * set the XF (format) record for this rec
	 */
	@Override
	public void setXFRecord( int i )
	{
		if( (i != ixfe) || (myxf == null) )
		{
			setIxfe( i );
			setXFRecord();
		}
	}

	/**
	 * set the XF (format) record for this rec
	 */
	@Override
	public void setIxfe( int i )
	{
		ixfe = i;
		byte[] newxfe = ByteTools.cLongToLEBytes( i );
		byte[] b = getData();
		if( b != null )
		{
			System.arraycopy( newxfe, 0, b, 4, 2 );
		}
		setData( b );
	}

	/**
	 * get the ixfe
	 */
	@Override
	public int getIxfe()
	{
		return ixfe;
	}

	@Override
	public void setByteReader( BlockByteReader db )
	{
		databuf = db;
		data = null;
	}

	@Override
	public BlockByteReader getByteReader()
	{
		return databuf;
	}

	/**
	 * Hold onto the original encrypted bytes so we can do a look ahead on records
	 */
	@Override
	public void setEncryptedByteReader( BlockByteReader db )
	{
		encryptedDatabuf = db;

	}

	@Override
	public BlockByteReader getEncryptedByteReader()
	{
		return encryptedDatabuf;
	}

	@Override
	public void setData( byte[] b )
	{
		data = b;
		databuf = null;
	}

	/**
	 * gets the record data merging any Continue record
	 * data.
	 */
	@Override
	public byte[] getData()
	{
		int len = getLength();
		if( len == 0 )
		{
			return new byte[]{ };
		}

		if( data != null )
		{
			return data;
		}

		if( len > MAXRECLEN )
		{
			setData( getBytesAt( 0, MAXRECLEN ) );
		}
		else
		{
			setData( getBytesAt( 0, len - 4 ) );
		}

		if( (opcode == SST) || (opcode == SXLI) )
		{
			return data;
		}

		if( !isContinueMerged && hasContinues() )
		{
			mergeContinues();
		}
		return data;
	}

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	@Override
	public byte[] getBytes()
	{
		return getBytesAt( 0, getLength() );
	}

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	@Override
	public byte[] getBytesAt( int off, int len )
	{
		if( data != null )
		{
			if( (len + off) > data.length )
			{
				len = data.length - off; // deal with bad requests
			}
			byte[] ret = new byte[len];
			System.arraycopy( data, off, ret, 0, len );
			return ret;
		}
		if( databuf == null )
		{
			return null;
		}
		return databuf.get( this, off, len );
	}

	/** Sets a subset of temporary bytes for fast access during init methods...
	 *

	 public void initCacheBytes(int start, int len) {
	 data = this.getBytesAt(start, len);
	 } */

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	@Override
	public byte getByteAt( int off )
	{
		if( data != null )
		{
			return data[off];
		}
		if( databuf == null )
		{
			throw new InvalidRecordException( "XLSRecord has no data buffer." + getCellAddress() );
		}
		return databuf.get( this, off );
	}

	@Override
	public void setLength( int len )
	{
		if( originalsize <= 0 )
		{
			originalsize = len;  // returns the original len always
		}
		reclen = len; // returns updated lengths
	}

	/**
	 * Returns the length of this
	 * record, including the 4 header bytes
	 */
	@Override
	public int getLength()
	{
		if( data != null )
		{
			return data.length + 4;
		}
		if( databuf == null ) // a new rec
		{
			return -1;
		}
		if( (hasContinues()) &&
				(!isContinueMerged) &&
				!(opcode == SST) &&
				!(opcode == SXLI) )
		{
			if( reclen > MAXRECLEN )
			{
				setData( getBytesAt( 0, MAXRECLEN ) );
			}
			else
			{
				setData( getBytesAt( 0, reclen ) );
			}
			mergeContinues();
			return data.length + 4;
		}
		return reclen + 4;
	}

	/**
	 * @return Returns the isValueForCell.
	 */
	@Override
	public boolean isValueForCell()
	{
		return isValueForCell;
	}

	/**
	 * @param isVal true if this is a cell-type record
	 */
	public void setValueForCell( boolean isVal )
	{
		isValueForCell = isVal;
	}

	/**
	 * @return Returns the isReadOnly.
	 */
	@Override
	public boolean isReadOnly()
	{
		return isReadOnly;
	}

	/**
	 * @return Returns the streamer.
	 */
	@Override
	public ByteStreamer getStreamer()
	{
		return streamer;
	}

	/**
	 * @param streamer The streamer to set.
	 */
	@Override
	public void setStreamer( ByteStreamer str )
	{
		streamer = str;
	}

	/**
	 * @return Returns the hyperlink.
	 */
	@Override
	public Hlink getHyperlink()
	{
		return hyperlink;
	}

	@Override
	public void postStream()
	{
		// nothing here -- use to blow out data
	}

}