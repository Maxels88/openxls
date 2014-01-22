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
import com.extentech.toolkit.Logger;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.AbstractList;
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

	private static final long serialVersionUID = -106915096753184441L;

	private short opcode;
	int reclen;
	byte[] data;
	private transient BlockByteReader databuf;
	private transient BlockByteReader encryptedDatabuf;
	protected boolean isContinueMerged = false;
	protected transient int DEBUGLEVEL = 0;
	boolean isValueForCell;
	boolean isFPNumber;
	boolean isDoubleNumber = false;
	boolean isIntNumber;
	public boolean isString;
	public boolean isBoolean;
	boolean isFormula;
	public boolean isBlank;
	boolean isReadOnly = false;
	protected int rw = -1;
	protected short col;
	public int offset = 0; // byte stream offset of rec
	protected transient Index idx;
	protected Sheet worksheet;
	public transient Xf myxf;
	protected transient WorkBook wkbook;
	protected transient ByteStreamer streamer;
	protected AbstractList continues;
	int originalsize = -1, originalIndex = 0, originalOffset = 0, ixfe = -1;
	public Hlink hyperlink = null;
	Row myrow = null;
	CellRange mergeRange;
	private int firstblock, lastblock;

	public short getOpcode()
	{
		return opcode;
	}

	public void setOpcode( short op )
	{
		this.opcode = op;
	}

	public void setHyperlink( Hlink hl )
	{
		this.hyperlink = hl;
	}

	public Formula getFormulaRec()
	{
		if( this instanceof Formula )
		{
			((Formula) this).populateExpression();
			return (Formula) this;
		}
		return null;
	}

	public boolean shouldEncrypt()
	{
		return true;
	}

	public void setRow( Row r )
	{
		myrow = r;
	}

	/**
	 * get the row of this cell
	 */
	public Row getRow()
	{
		if( myrow == null )
		{
			myrow = this.worksheet.getRowByNumber( rw );
		}
		return myrow;
	}

	public Xf getXfRec()
	{
		if( myxf == null )
		{
			if( (ixfe > -1) && (ixfe < wkbook.getNumXfs()) )
			{
				this.myxf = wkbook.getXf( this.ixfe );
			}
		}
		return myxf;
	}

	/**
	 * returns the existing font record
	 * for this Cell
	 */
	public Font getFont()
	{
		WorkBook b = this.getWorkBook();
		if( b == null )
		{
			return null;
		}
		this.getXfRec();
		if( myxf == null )
		{
			return null;
		}
		return myxf.getFont();
	}

	/**
	 * @return
	 */
	public CellRange getMergeRange()
	{
		return mergeRange;
	}

	/**
	 * @param range
	 */
	public void setMergeRange( CellRange range )
	{
		mergeRange = range;
	}

	/**
	 * Removes this BiffRec from the WorkSheet
	 *
	 * @param whether to nullify this Cell
	 */
	public boolean remove( boolean nullme )
	{
		boolean success = false;
		if( worksheet != null && isValueForCell )
		{
			getSheet().removeCell( this );
		}

		if( streamer != null )
		{
			streamer.removeRecord( this );
		}

		if( nullme )
		{
			try
			{
				this.finalize();
			}
			catch( Throwable t )
			{
				;
			}
		}
		this.worksheet = null;
		return true;
	}

	/**
	 * set the Formatting for this BiffRec from the pattern
	 * match.
	 * <p/>
	 * case insensitive pattern match is performed...
	 */
	public String getFormatPattern()
	{
		if( myxf == null )
		{
			return "";
		}
		return myxf.getFormatPattern();
	}

	/**
	 * get the int val of the type for the valrec
	 */
	public int getCellType()
	{
		if( this.isBlank )
		{
			return TYPE_BLANK;
		}
		if( this.isString )
		{
			return TYPE_STRING;
		}
		if( this.isDoubleNumber )
		{
			return TYPE_DOUBLE;
		}
		if( this.isFPNumber )
		{
			return TYPE_FP;
		}
		if( this.isIntNumber )
		{
			return TYPE_INT;
		}
		if( this.isFormula )
		{
			return TYPE_FORMULA;
		}
		if( this.isBoolean )
		{
			return TYPE_BOOLEAN;
		}
		return -1;
	}

	// Methods from BlockByteConsumer //

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
	public void setOffset( int pos )
	{
		if( originalOffset < 1 )
		{
			originalOffset = pos;
		}
		offset = pos;
	}

	/**
	 * Get the relative position within the data
	 * underlying the block vector represented by
	 * the BlockByteReader.
	 *
	 * @return relative position
	 */
	public int getOffset()
	{
		if( this.data == null )
		{
			return this.originalOffset;
		}
		return (int) offset;
	}

	/** Get the blocks containing this Consumer's data
	 *
	 * @return
	 */
/*unused: public Block[] getBlocks() {
		return myblocks;
	}*/

	/** Set the blocks containing this Consumer's data
	 *
	 * @param myblocks
	 */
/* unused: 	public void setBlocks(Block[] myb) {
		myblocks = myb;
	}*/

	/**
	 * Sets the index of the first block
	 *
	 * @return
	 */
	public void setFirstBlock( int i )
	{
		firstblock = i;
	}

	/**
	 * Sets the index of the last block
	 *
	 * @return
	 */
	public void setLastBlock( int i )
	{
		lastblock = i;
	}

	/**
	 * Returns the index of the first block
	 *
	 * @return
	 */
	public int getFirstBlock()
	{
		return firstblock;
	}

	/**
	 * Returns the index of the last block
	 *
	 * @return
	 */
	public int getLastBlock()
	{
		return lastblock;
	}

	protected void initRowCol()
	{
		int pos = 0;

		byte[] bt = this.getBytesAt( pos, 2 );
		rw = ByteTools.readUnsignedShort( bt[0], bt[1] );
		pos += 2;
		col = ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
	}

	public String toString()
	{
		return getRecDesc();
	}

	public String getRecDesc()
	{
		String ret = "";
		String name = this.getClass().getSimpleName();

		// record name
		ret += (name.equals( "XLSRecord" ) ? "unknown" : name.toUpperCase());
		// hex record number
		ret += " (" + Integer.toHexString( opcode ).toUpperCase() + "h)";
		// stream offset
		ret += " at " + Integer.toHexString( offset ).toUpperCase() + "h";
		// size
		ret += " length " + Integer.toHexString( reclen ).toUpperCase() + "h";
		// file offset
		ret += " file " + (databuf == null ? "no file"/*originalFileOffset*/ : databuf.getFileOffsetString( offset, reclen ));
		// cell address, if applicable
		if( this.isValueForCell() )
		{
			ret += " cell " + this.getCellAddress();
		}

		return ret;
	}

	/**
	 * Dumps this record as a human-readable string.
	 */
	public String toHexDump()
	{
		return getRecDesc() + "\n" + ByteTools.getByteDump( this.getData(), 0 );
	}

	/**
	 * Copy all formatting info from source biffrec
	 *
	 * @see com.extentech.formats.XLS.BiffRec#copyFormat(com.extentech.formats.XLS.BiffRec)
	 */
	public void copyFormat( BiffRec source )
	{
		Xf clone = (Xf) source.getXfRec().clone();
		Font fontClone = (Font) source.getXfRec().getFont().clone();

		Logger.logInfo( source + ":" + source.getXfRec() + ":" + clone );
		int fid = -1;
		int xid = -1;
		// see if we have an equivalent Xf/Font combo
		if( this.getWorkBook().getFontRecs().contains( fontClone ) )
		{
			fid = this.getWorkBook().getFontRecs().indexOf( fontClone );
		}
		if( this.getWorkBook().getXfrecs().contains( clone ) )
		{
			xid = this.getWorkBook().getXfrecs().indexOf( clone );
		}

		// add the xf/font and set the ixfe
		this.getWorkBook().addRecord( clone, false );
		this.getWorkBook().addRecord( fontClone, false );
		clone.setFont( fontClone.getIdx() );
		this.setXFRecord( clone.getIdx() );

	}

	/**
	 * clone a record
	 */
	public Object clone()
	{
		try
		{
			String cn = getClass().getName();
			XLSRecord rec = (XLSRecord) Class.forName( cn ).newInstance();
			byte[] inb = getBytes();
			rec.setData( inb );
			rec.streamer = this.streamer;
			rec.setWorkBook( getWorkBook() );
			rec.setOpcode( getOpcode() );
			rec.setLength( getLength() );
			rec.setSheet( getSheet() );    //20081120 KSC: otherwise may set sheet incorrectly in init
			rec.init();
			return rec;
		}
		catch( Exception e )
		{
			Logger.logInfo( "cloning XLSRecord " + this.getCellAddress() + " failed: " + e );
		}
		return null;
	}

	/**
	 * return whether this is a numeric type
	 */
	public boolean isNumber()
	{
		if( this.opcode == RK )
		{
			return true;
		}
		if( this.opcode == NUMBER )
		{
			return true;
		}
		return false;
	}

	/**
	 * return whether this is a formula record
	 */
	public boolean isFormula()
	{
		return this.isFormula;
	}

	/** methods from CompatibleVectorHints

	 protected transient int recordIdx = -1;
	 */
	/** provide a hint to the CompatibleVector
	 about this objects likely position.

	 public int getRecordIndexHint(){return recordIdx;}
	 */

	/** set index information about this
	 objects likely position.

	 public void setRecordIndexHint(int i){
	 lastidx = i;
	 recordIdx = i;
	 }
	 */
	/**
	 * set the DEBUG level
	 */
	public void setDebugLevel( int b )
	{
		DEBUGLEVEL = b;
	}

	public int lastidx = -1;

	/**
	 * return the real (not just boundsheet) record index of this object
	 */
	public int getRealRecordIndex()
	{
		return streamer.getRealRecordIndex( this );
	}

	/**
	 * return the record index of this object
	 */
	public int getRecordIndex()
	{
		if( streamer == null )
		{
			if( this.getSheet() != null )    // KSC: Added
			{
				return this.getSheet().getSheetRecs().indexOf( this );
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
	public void addContinue( Continue c )
	{
		if( continues == null )
		{
			continues = new ArrayList();
		}
		continues.add( c );
	}

	/**
	 * remove all Continue records
	 */
	public void removeContinues()
	{
		if( continues != null )
		{
			continues.clear();
		}
	}

	public List getContinueVect()
	{
		if( continues != null )
		{
			return this.continues;
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
	public boolean hasContinues()
	{
		if( continues == null )
		{
			return false;
		}
		if( this.continues.size() > 0 )
		{
			return true;
		}
		return false;
	}

	/**
	 * set whether this record contains the  value
	 * of the Cell.
	 */
	public void setIsValueForCell( boolean b )
	{
		isValueForCell = b;
	}

	/**
	 * associate this record with its Index record
	 */
	public void setIndex( Index id )
	{
		idx = id;
	}

	/**
	 * Associate this record with a worksheet.
	 * First checks to see if there is already
	 * a cell with this address.
	 */
	public void setSheet( Sheet b )
	{
		this.worksheet = b;
	}

	/**
	 * get the WorkSheet for this record.
	 */
	public Boundsheet getSheet()
	{
		return (Boundsheet) worksheet;
	}

	public void setWorkBook( WorkBook wk )
	{
		wkbook = (WorkBook) wk;
	}

	public WorkBook getWorkBook()
	{
		if( (wkbook == null) && (worksheet != null) )
		{
			wkbook = worksheet.getWorkBook();
		}

		return (WorkBook) wkbook;
	}

	/**
	 * set the column
	 */
	public void setCol( short i )
	{
		byte[] c = ByteTools.shortToLEBytes( (short) i );
		System.arraycopy( c, 0, getData(), 2, 2 );
		col = i;
	}

	public void setRowCol( int[] x )
	{
		this.setRowNumber( x[0] );
		this.setCol( (short) x[1] );
	}

	/**
	 * set the row
	 */
	public void setRowNumber( int i )
	{
		byte[] r = ByteTools.cLongToLEBytes( i );
		System.arraycopy( r, 0, getData(), 0, 2 );
		rw = i;
	}

	public short getColNumber()
	{
		return col;
	}

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
			if( DEBUGLEVEL > -1 )
			{
				Logger.logWarn( "XLSRecord.getCellAddress() Row/Col info incorrect for Cell:" + ExcelTools.getAlphaVal( col ) + String.valueOf(
						rownum ) );
			}
			return "";
		}
		return ExcelTools.getAlphaVal( col ) + rownum;
	}

	/**
	 * returns the cell address in int[] {row, col} format
	 */
	public int[] getIntLocation()
	{
		return new int[]{ rw, col };
	}

	/**
	 * return the cell address with sheet reference
	 * eg Sheet!A12
	 *
	 * @return String
	 */
	public String getCellAddressWithSheet()
	{
		if( this.getSheet() != null )
		{
			return this.getSheet().getSheetName() + "!" + this.getCellAddress();
		}
		return this.getCellAddress();
	}

	/**
	 * perform record initialization
	 */
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
	public Object getDefaultVal()
	{
		if( this.isDoubleNumber )
		{
			return new Double( 0.0 );
		}
		if( this.isFPNumber )
		{
			return new Float( 0.0f );
		}
		if( this.isBoolean )
		{
			return Boolean.valueOf( false );
		}
		if( this.isIntNumber )
		{
			return Integer.valueOf( 0 );
		}
		if( this.isString )
		{
			return "";
		}
		if( this.isFormula )
		{
			return this.getFormulaRec().getFormulaString();
		}
		if( this.isBlank )
		{
			return "";
		}
		return null;
	}

	/**
	 * get the data type name for this record
	 */
	public String getDataType()
	{
		if( this.isValueForCell )
		{
			if( this.isBlank )
			{
				return "Blank";
			}
			if( this.isDoubleNumber )
			{
				return "Double";
			}
			if( this.isFPNumber )
			{
				return "Float";
			}
			if( this.isBoolean )
			{
				return "Boolean";
			}
			if( this.isIntNumber )
			{
				return "Integer";
			}
			if( this.isString )
			{
				return "String";
			}
			if( this.isFormula )
			{
				return "Formula";
			}
		}
		return null;
	}

	/**
	 * get the int val of the type for the valrec
	 */
	public Object getInternalVal()
	{
		try
		{
			switch( this.getCellType() )
			{
				case TYPE_BLANK:
					return getStringVal(); //essentially return "";

				case TYPE_STRING:
					return getStringVal();

				case TYPE_FP: // always use Doubles to avoid loss of precision... see: 
					// details http://stackoverflow.com/questions/916081/convert-float-to-double-without-losing-precision
					return new Double( getDblVal() );

				case TYPE_DOUBLE:
					return new Double( getDblVal() );

				case TYPE_INT:
					return Integer.valueOf( getIntVal() );

				case TYPE_FORMULA:
					// OK this is broken, obviously we need to return a calced Object
					Object obx = ((Formula) this).calculateFormula();
					return obx;
//					return getStringVal();

				case TYPE_BOOLEAN:
					return Boolean.valueOf( getBooleanVal() );

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
	 * Get the value of the record as a Boolean.
	 * Value must be parseable as a Boolean.
	 */
	public boolean getBooleanVal()
	{
		return false;
	}

	/**
	 * Get the value of the record as an Integer.
	 * Value must be parseable as an Integer or it
	 * will throw a NumberFormatException.
	 */
	public int getIntVal()
	{
		return (int) Float.NaN;
	}

	/**
	 * Get the value of the record as a Double.
	 * Value must be parseable as an Double or it
	 * will throw a NumberFormatException.
	 */
	public double getDblVal()
	{
		return (double) Float.NaN;
	}

	/**
	 * Get the value of the record as a Float.
	 * Value must be parseable as an Float or it
	 * will throw a NumberFormatException.
	 */
	public float getFloatVal()
	{
		return Float.NaN;
	}

	/**
	 * Get the value of the record as a String.
	 */
	public String getStringVal()
	{
		return null;
	}

	/**
	 * Get the value of the record as a String.
	 */
	public String getStringVal( String encoding )
	{
		return null;
	}

	public void setStringVal( String v )
	{
		Logger.logErr( "Setting String Val on generic XLSRecord, value not held" );
	}

	public void setBooleanVal( boolean b )
	{
		Logger.logErr( "Setting Boolean Val on generic XLSRecord, value not held" );
	}

	public void setIntVal( int v )
	{
		Logger.logErr( "Setting int Val on generic XLSRecord, value not held" );
	}

	public void setFloatVal( float v )
	{
		Logger.logErr( "Setting float Val on generic XLSRecord, value not held" );
	}

	public void setDoubleVal( double v )
	{
		Logger.logErr( "Setting Double Val on generic XLSRecord, value not held" );
	}

	/**
	 * do any pre-streaming processing such as expensive
	 * index updates or other deferrable processing.
	 */
	public void preStream()
	{
		// override in sub-classes
	}

	/**
	 * set the XF (format) record for this rec
	 */
	public void setXFRecord()
	{
		if( wkbook == null )
		{
			return;
		}
		if( (ixfe > -1) && (ixfe < wkbook.getNumXfs()) )
		{
			if( myxf == null || myxf.tableidx != ixfe )
			{
				this.myxf = wkbook.getXf( this.ixfe );
				this.myxf.incUseCount();
			}
		}
	}

	/**
	 * set the XF (format) record for this rec
	 */
	public void setXFRecord( int i )
	{
		if( i != ixfe || myxf == null )
		{
			this.setIxfe( i );
			this.setXFRecord();
		}
	}

	/**
	 * set the XF (format) record for this rec
	 */
	public void setIxfe( int i )
	{
		this.ixfe = i;
		byte[] newxfe = ByteTools.cLongToLEBytes( i );
		byte[] b = this.getData();
		if( b != null )
		{
			System.arraycopy( newxfe, 0, b, 4, 2 );
		}
		this.setData( b );
	}

	/**
	 * get the ixfe
	 */
	public int getIxfe()
	{
		return this.ixfe;
	}

	/**
	 * get a new, generic instance of a Record.
	 */
	protected static XLSRecord getPrototype()
	{
		Logger.logWarn( "Attempt to get prototype XLSRecord failed.  There is no prototype record defined for this record type." );
		return null;
	}

	public void setByteReader( BlockByteReader db )
	{
		databuf = db;
		data = null;
	}

	public BlockByteReader getByteReader()
	{
		return databuf;
	}

	/**
	 * Hold onto the original encrypted bytes so we can do a look ahead on records
	 */
	public void setEncryptedByteReader( BlockByteReader db )
	{
		encryptedDatabuf = db;

	}

	public BlockByteReader getEncryptedByteReader()
	{
		return encryptedDatabuf;
	}

	public void setData( byte[] b )
	{
		data = b;
		this.databuf = null;
	}

	/**
	 * gets the record data merging any Continue record
	 * data.
	 */
	public byte[] getData()
	{
		int len = 0;
		if( (len = this.getLength()) == 0 )
		{
			return new byte[]{
			};
		}
		if( data != null )
		{
			return data;
		}
		if( len > MAXRECLEN )
		{
			setData( this.getBytesAt( 0, MAXRECLEN ) );
		}
		else
		{
			setData( this.getBytesAt( 0, len - 4 ) );
		}

		if( this.opcode == SST || this.opcode == SXLI )
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
	 * Merge continue data in to the record data array
	 */
	protected void mergeContinues()
	{
		List cx = this.getContinueVect();
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
				Logger.logWarn( "ERROR: parsing record continues failed: " + this.toString() + ": " + e );
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
					Logger.logWarn( "ERROR: parsing record continues failed: " + this.toString() + ": " + a );
				}
			}
			this.data = out.toByteArray();
		}
		isContinueMerged = true;
	}

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	public byte[] getBytes()
	{
		return this.getBytesAt( 0, this.getLength() );
	}

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	public byte[] getBytesAt( int off, int len )
	{
		if( this.data != null )
		{
			if( len + off > data.length )
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
		return this.databuf.get( this, off, len );
	}

	/** Sets a subset of temporary bytes for fast access during init methods...
	 *

	 public void initCacheBytes(int start, int len) {
	 data = this.getBytesAt(start, len);
	 } */

	/**
	 * resets the cache bytes so they do not take up space
	 */
	public void resetCacheBytes()
	{
		//  data = null;
	}

	/**
	 * Gets the byte from the specified position in the
	 * record byte array.
	 *
	 * @param off
	 * @return
	 */
	public byte getByteAt( int off )
	{
		if( this.data != null )
		{
			return data[off];
		}
		if( this.databuf == null )
		{
			throw new InvalidRecordException( "XLSRecord has no data buffer." + this.getCellAddress() );
		}
		return this.databuf.get( this, off );
	}

	public void setLength( int len )
	{
		if( this.originalsize <= 0 )
		{
			this.originalsize = len;  // returns the original len always
		}
		this.reclen = len; // returns updated lengths
	}

	/**
	 * Returns the length of this
	 * record, including the 4 header bytes
	 */
	public int getLength()
	{
		if( data != null )
		{
			return data.length + 4;
		}
		else if( this.databuf == null ) // a new rec
		{
			return -1;
		}
		if( (hasContinues()) &&
				(!isContinueMerged) &&
				!(this.opcode == SST) &&
				!(this.opcode == SXLI) )
		{
			if( this.reclen > MAXRECLEN )
			{
				setData( this.getBytesAt( 0, MAXRECLEN ) );
			}
			else
			{
				setData( this.getBytesAt( 0, this.reclen ) );
			}
			mergeContinues();
			return data.length + 4;
		}
		return this.reclen + 4;
	}

	/**
	 * @return Returns the isValueForCell.
	 */
	public boolean isValueForCell()
	{
		return isValueForCell;
	}

	/**
	 * @param isVal true if this is a cell-type record
	 */
	public void setValueForCell( boolean isVal )
	{
		this.isValueForCell = isVal;
	}

	/**
	 * @return Returns the isReadOnly.
	 */
	public boolean isReadOnly()
	{
		return isReadOnly;
	}

	/**
	 * @return Returns the streamer.
	 */
	public ByteStreamer getStreamer()
	{
		return streamer;
	}

	/**
	 * @param streamer The streamer to set.
	 */
	public void setStreamer( ByteStreamer str )
	{
		streamer = str;
	}

	/**
	 * @return Returns the hyperlink.
	 */
	public Hlink getHyperlink()
	{
		return hyperlink;
	}

	public void postStream()
	{
		// nothing here -- use to blow out data
	}

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
			for( int i = 0; i < continues.size(); i++ )
			{
				((XLSRecord) continues.get( i )).close();
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
	 * Get the color table for the associated workbook.  If the workbook is null
	 * then the default colortable will be returned.
	 */
	public java.awt.Color[] getColorTable()
	{
		try
		{
			return this.getWorkBook().getColorTable();
		}
		catch( Exception e )
		{
			return FormatHandle.COLORTABLE;
		}
	}

}