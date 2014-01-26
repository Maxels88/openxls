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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.NumberFormat;

/**
 * <b>RK: RK Number (7Eh)</b><br>
 * This record stores an internal numeric type.  Stores data in one of four
 * RK 'types' which determine whether it is an integer or an IEEE floating point
 * equivalent.
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      rk          4       RK number
 * </p></pre>
 *
 * @see MULRK
 * @see NUMBER
 */

public final class Rk extends XLSCellRecord implements Mulled
{
	private static final Logger log = LoggerFactory.getLogger( Rk.class );
	private static final long serialVersionUID = -3027662614434608240L;
	int rkType;
	double Rkdouble;
	int RKint;
	Mulrk mymul;

	public static final int RK_FP = 0;
	public static final int RK_FP_100 = 1;
	public static final int RK_INT = 2;
	public static final int RK_INT_100 = 3;

	public boolean DEBUG = false;

	@Override
	public void setMyMul( Mul m )
	{
		mymul = (Mulrk) m;
	}

	@Override
	public Mul getMyMul()
	{
		return mymul;
	}

	@Override
	public void setNoMul()
	{
		mymul = null;
	}

	/**
	 * default constructor
	 */
	public Rk()
	{
		super();
	}

	/**
	 * Provide constructor which automatically
	 * sets the body data and header info.  This
	 * is needed by MULRK which creates the RKs without
	 * the benefit of WorkBookFactory.parseRecord().
	 */
	Rk( byte[] b, int r, int c )
	{
		rw = r;
		col = (short) c;
		setData( b );
		setOpcode( RK );
		setLength( (short) 10 );
		init( b );
	}

	/**
	 * This init method
	 * is needed by MULRK which creates the RKs without
	 * the benefit of WorkBookFactory.parseRecord().
	 */
	// called by Mulrk.init
	void init( byte[] b, int r, int c )
	{
		rw = r;
		col = (short) c;
		byte[] rwbt = ByteTools.shortToLEBytes( (short) r );
		byte[] colbt = ByteTools.shortToLEBytes( (short) c );
		byte[] newData = new byte[10];
		newData[0] = rwbt[0];
		newData[1] = rwbt[1];
		newData[2] = colbt[0];
		newData[3] = colbt[1];
		System.arraycopy( b, 0, newData, 4, b.length );
		setData( newData );
		setOpcode( RK );
		init( b );
	}

	/**
	 * This init method pulls out the record header information,
	 * then sends the as-yet unmodded rkdata record across to the
	 * rktranslate method
	 */
	void init( byte[] rkdata )
	{
		super.init();
		short s;
		byte[] rknum = new byte[4];
		// if this is a 'standalone' RK number, then the byte array
		// contains row, col and ixfe data as well as the number value.
		if( rkdata.length > 6 )
		{
			// get the row information
			super.initRowCol();
			s = ByteTools.readShort( rkdata[4], rkdata[5] );
			ixfe = s;
			System.arraycopy( rkdata, 6, rknum, 0, 4 );
		}
		else
		{
			// get the ixfe information
			s = ByteTools.readShort( rkdata[0], rkdata[1] );
			ixfe = s;
			System.arraycopy( rkdata, 2, rknum, 0, 4 );
		}
		translateRK( rknum );
	}

	/**
	 * This init method pulls out the record header information,
	 * then sends the as-yet unmodded rkdata record across to the
	 * rktranslate method
	 */
	@Override
	public void init()
	{
		super.init();
		getData();
		short s;
		byte[] rknum = new byte[4];
		// if this is a 'standalone' RK number, then the byte array
		// contains row, col and ixfe data as well as the number value.
		if( getLength() > 6 )
		{
			// get the row information

			super.initRowCol();
			s = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
			ixfe = s;
			byte[] numdat = getBytesAt( 6, 4 );
			System.arraycopy( numdat, 0, rknum, 0, 4 );
		}
		else
		{
			// get the ixfe information
			s = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
			ixfe = s;
			byte[] numdat = getBytesAt( 2, 4 );
			System.arraycopy( numdat, 0, rknum, 0, 4 );
		}
		translateRK( rknum );
	}

	/**
	 * figures out the type of RK that we have, does the
	 * little endian thing, then sends it over to getRealVal
	 */
	void translateRK( byte[] rkval )
	{

		int l = 0;
		short num1 = ByteTools.readShort( rkval[0], rkval[1] );
		short num2 = ByteTools.readShort( rkval[2], rkval[3] );
		l = ByteTools.readInt( num2, num1 );

		long num = l;
		// num = num << 1;
		// num = num >>> 1;

		// check what the RK type bits are
		int bitset = l & (1 << 0);
		int bitset2 = l & (1 << 1);
		// add them to get the type
		rkType = bitset + bitset2;
		double d = 1.0;

		Rkdouble = Rk.getRealVal( rkType, num );
		setIsValueForCell( true );
		isDoubleNumber = false;
		switch( rkType )
		{
			case Rk.RK_FP:
				// okay, am i dense or something or is
				// 1 NOT a Float?  see RK Type 0 on pg. 377
				// then tell me I'm not on crack... -jm 11/03
				String newnum = String.valueOf( Rkdouble );
				if( newnum.length() > 12 )
				{
					isDoubleNumber = true;
				}
				int mantindex = newnum.indexOf( "." );
				newnum = newnum.substring( mantindex + 1 );
				try
				{
					if( Integer.parseInt( newnum ) > 0 )
					{ // there's FP digits
						isFPNumber = true;
						isIntNumber = false;
					}
					else
					{
						isFPNumber = false;
						isIntNumber = true;
						RKint = (int) Rkdouble;
					}
				}
				catch( NumberFormatException e )
				{
					isFPNumber = true;
					isIntNumber = false;

					//RKint = (int)Rkdouble;
				}
				if( Rkdouble > Float.MAX_VALUE )
				{
					isDoubleNumber = true;
				}
				break;

			case Rk.RK_FP_100:
				isFPNumber = true;
				isIntNumber = false;
				break;
			case Rk.RK_INT:
				isFPNumber = false;
				isIntNumber = true;
				RKint = (int) Rkdouble;
				break;
			case Rk.RK_INT_100:
				newnum = String.valueOf( Rkdouble );
				if( newnum.toUpperCase().contains( "E" ) )
				{
					// do something intelligent
					NumberFormat nmf = NumberFormat.getInstance();
					try
					{
						Number nm = nmf.parse( newnum );
						float v = nm.floatValue();
							log.debug( "Rk number format: " + v );
					}
					catch( Exception e )
					{
						log.debug( "Exception parsing newnum: {} ", newnum,e );
					}
				}
				else
				{
					mantindex = newnum.indexOf( "." );
					newnum = newnum.substring( mantindex + 1 );
				}
				try
				{
					if( Long.parseLong( newnum ) > 0 )
					{ // there's FP digits
						isFPNumber = true;
						isIntNumber = false;
					}
					else
					{
						isFPNumber = false;
						isIntNumber = true;
						RKint = (int) Rkdouble;
					}
//					happens with big numbers with exponents 
//					they should be ints
				}
				catch( NumberFormatException e )
				{
					isFPNumber = false;
					isIntNumber = true;
					RKint = (int) Rkdouble;
				}
				break;
		}
	}

	/**
	 * returns the position of this record in the array of records
	 * making up this file.
	 */
	@Override
	public int getRecordIndex()
	{
		if( super.getRecordIndex() < 0 )
		{ // this is a MulRk
			if( mymul == null )
			{
				return -1; // throw new InvalidRecordException("Rk without a recidx nor a Mulrk.");
			}
			return mymul.getRecordIndex();
		}
		// standalone RK

		return super.getRecordIndex();
	}

	/**
	 * static method which parses a 4-byte RK number
	 * into a double value using specific MS Rules
	 *
	 * @param byte[] rkbytes - 4 byte Rk Number
	 * @return double - translated from Rk bytes
	 * @see Rk
	 */
	public static double parseRkNumber( byte[] rkbytes )
	{
		int num = 0;
		short num1 = ByteTools.readShort( rkbytes[0], rkbytes[1] );
		short num2 = ByteTools.readShort( rkbytes[2], rkbytes[3] );
		num = ByteTools.readInt( num2, num1 );

		// check what the RK type bits are
		int bitset = num & (1 << 0);
		int bitset2 = num & (1 << 1);
		// add them to get the type
		int rkType = bitset + bitset2;

		return Rk.getRealVal( rkType, num );
	}

	/**
	 * Madness with Microsoft trying to figure out its own format for floating point.
	 * to really understand what is going on RTFM, but basically there are four different
	 * encodings used for RKs, a modified int, which uses its last two bits as an identifer,
	 * this same int divided by 100, a modified float with the Exponent of a double, but with the
	 * last 2 bits used as an identifier, and finally the same as the last divided by 100. Hmmm
	 */
	private static double getRealVal( int RKType, long waknum )
	{
// NOTE isFPNumber, isIntNumber was set in translateRk anyways so take out of here        	
		// we need to mask the number to avoid the final 2 bits messin stuff up
		waknum = (waknum & 0xFFFFFFFC);
		// perform the change for each type
		switch( RKType )
		{
			case Rk.RK_FP:
				// IEEE Number
				waknum = waknum << 32;
				String testq = Long.toBinaryString( waknum );
				return Double.longBitsToDouble( waknum );

			case Rk.RK_FP_100:
				double res = Double.longBitsToDouble( waknum << 32 );
				res /= 100;
				return res;

			case Rk.RK_INT:
				// Integer
				waknum = waknum >> 2;
				return waknum;

			case Rk.RK_INT_100:
				if( waknum >= 4290773292l )
				{
					log.warn( "Erroneous Rk: {}", waknum ); // THIS IS THE CUTOFF NUMBER -- ANYTHING THIS SIZE IS < -10,485.01
				}
				// Integer x 100
				waknum = waknum >> 2;
				double ddd = waknum;
				return (ddd / 100);

			default:
				log.warn( "incorrect RK type for RK record: " + String.valueOf( RKType ) );
		}
		return 0.0;
	}

	@Override
	public int getIntVal() throws RuntimeException
	{
		if( isFPNumber )
		{
			long l = (long) Rkdouble;
			if( l > Integer.MAX_VALUE )
			{
				throw new NumberFormatException( "Cell value is larger than the maximum java signed int size" );
			}
			if( l < Integer.MIN_VALUE )
			{
				throw new NumberFormatException( "Cell value is smaller than the minimum java signed int size" );
			}
			return (int) Rkdouble;
		}
		return RKint;
	}

	@Override
	public double getDblVal()
	{
		if( isIntNumber )
		{
			return RKint;
		}
		return Rkdouble;
	}

	@Override
	public float getFloatVal()
	{
		if( isIntNumber )
		{
			return RKint;
		}
		return (float) Rkdouble;
	}

	/**
	 * Return the string value.  If it is over 99999999999 or under -99999999999
	 * then return a format using scientific notation.  Note this is *not* significant digits,
	 * rather the actual size of the number.  Emulates Excel functionality and display
	 */
	@Override
	public String getStringVal()
	{
		if( isIntNumber )
		{
			return String.valueOf( RKint );
		}
		return ExcelTools.getNumberAsString( Rkdouble );
	}

	@Override
	public void setStringVal( String s )
	{
		try
		{
			if( s.indexOf( "." ) > -1 )
			{
				double f = new Double( s );    // 20080211 KSC: Double.valueOf(s).doubleValue();
				setDoubleVal( f );
			}
			else
			{
				int i = Integer.parseInt( s );
				setIntVal( i );
			}
		}
		catch( java.lang.NumberFormatException f )
		{
			log.warn( "in Rk " + s + " is not a number." );
		}
	}

	/**
	 * @see com.extentech.formats.XLS.XLSRecord#setFloatVal(float)
	 */
	@Override
	public void setFloatVal( float f )
	{
		try
		{
			setRKVal( f );
		}
		catch( Exception x )
		{
			log.warn( "Rk.setFloatVal() problem.  Fallback to floating point Number." );
			Rk.convertRkToNumber( this, f );
		}
	}

	/**
	 * @see com.extentech.formats.XLS.XLSRecord#setIntVal(int)
	 */
	@Override
	public void setIntVal( int f )
	{
		try
		{
			setRKVal( f );
		}
		catch( Exception x )
		{
			log.warn( "Rk.setIntVal() problem.  Fallback to floating point Number." );
			Rk.convertRkToNumber( this, f );
		}
	}

	/**
	 * @see com.extentech.formats.XLS.XLSRecord#setDoubleVal(double)
	 */
	@Override
	public void setDoubleVal( double f )
	{
		try
		{
			setRKVal( f );
		}
		catch( Exception x )
		{
			log.warn( "Rk.setDoubleVal() problem.  Fallback to floating point Number." );
			Rk.convertRkToNumber( this, f );
		}
	}

	public static String getTypeName()
	{
		return "Rkdouble";
	}

	/**
	 * Change the row of this record, including parent Mulrk if it exists.
	 */
	void setMulrkRow( int i )
	{
		super.setRowNumber( i );
		if( mymul != null )
		{
			mymul.setRow( i );
		}
	}

	/**
	 * static method returns the double value converted to Rk-type bytes (a 4-byte structure)
	 * as well as the Rk Type (FP, INT, FP_100, INT_100) in the 5th byte position
	 * @param  double d - double to convert to Rk
	 * @return byte[] array of bytes representing the Rk number (4 bytes) and the Rk type (1 byte)
	 */
	/**
	 * Structure of RkNumber:
	 * A - fX100 (1 bit): A bit that specifies whether num is the value of the RkNumber or 100 times the value of the RkNumber. MUST be a value from the following table:
	 * 0 = The value of RkNumber is the value of num.
	 * 1 = The value of RkNumber is the value of num divided by 100.
	 * B - fInt (1 bit): A bit that specifies the type of num.
	 * 0 =    num is the 30 most significant bits of a 64-bit binary floating-point number as defined in [IEEE754].
	 * The remaining 34-bits of the floating-point number MUST be 0.
	 * 1 =    num is a signed integer
	 * num (30 bits): A variable type field whose type and meaning is specified by the value of fInt, as defined in the following table:
	 */
	public static byte[] getRkBytes( double d )
	{
		long bitlong = Double.doubleToLongBits( d );
		long bigger = Double.doubleToLongBits( (d * 100) );
		long l = (long) d;
		l = java.lang.Math.abs( l );

		byte[] rkbytes = new byte[5];
		rkbytes[4] = -1;    // uninitialized type
		byte[] doublebytes = new byte[8];

		// are low order 34 bits of d = 0?  RK type = 0 (RK_FP)
		// d is a 64 bit num, so move it over a bit...
		if( (bitlong << 30) == 0 )
		{
			doublebytes = ByteTools.doubleToByteArray( d );
			// add the RK type at the end of the last byte
			byte mask = (byte) 0xfc;
			doublebytes[3] = (byte) (doublebytes[3] & mask);
			// bit flipping for the LE switch
			rkbytes[0] = doublebytes[3];
			rkbytes[1] = doublebytes[2];
			rkbytes[2] = doublebytes[1];
			rkbytes[3] = doublebytes[0];
			rkbytes[4] = RK_FP;
		}
		// Can d be represented by a 30 bit integer? RK Type = 2 (RK_INT)
		// ORIGINAL -- else if ((l>>>30) ==0 && ((d%1)==0)){
		else if( (((l << 2) >>> 30) == 0) && ((d % 1) == 0) )
		{
			long lo = (long) d;
			lo = (lo << 2);
			doublebytes = ByteTools.longToByteArray( lo );        // RK_INT
			byte mask = 0x2;
			doublebytes[7] = (byte) (doublebytes[7] | mask);
			rkbytes[0] = doublebytes[7];
			rkbytes[1] = doublebytes[6];
			rkbytes[2] = doublebytes[5];
			rkbytes[3] = doublebytes[4];
			rkbytes[4] = RK_INT;
		}/**/
		// are low order 34 bits of d * 100 = 0?  RK type = 1 (RK_FP_100)
		else if( (bigger << 30) == 0 )
		{
			doublebytes = ByteTools.doubleToByteArray( d * 100 );        // F100
			byte mask = 0x1;
			doublebytes[3] = (byte) (doublebytes[3] | mask);
			rkbytes[0] = doublebytes[3];
			rkbytes[1] = doublebytes[2];
			rkbytes[2] = doublebytes[1];
			rkbytes[3] = doublebytes[0];
			rkbytes[4] = RK_FP_100;
		}
		// Can d * 100 be represented by a 30 bit integer? RK type = 3 (RK_INT_100)
		else if( (((d * 100) % 1) == 0) && (((l * 100) >>> 30) == 0) )
		{
			long lo = (long) (d * 100);        // F100 + INT
			lo = (lo << 2);
			doublebytes = ByteTools.longToByteArray( lo );
			byte mask = 0x3;
			doublebytes[7] = (byte) (doublebytes[7] | mask);
			rkbytes[0] = doublebytes[7];
			rkbytes[1] = doublebytes[6];
			rkbytes[2] = doublebytes[5];
			rkbytes[3] = doublebytes[4];
			rkbytes[4] = RK_INT_100;
		}
		if( (rkbytes[4] != -1) && (d != Rk.parseRkNumber( rkbytes )) )    // if it was processed as an RK, ensure results are accurate
		{
			throw new RuntimeException( d + " could not be translated to Rk value" );
		}

		return rkbytes;
	}

	/**
	 * Allows writing back to the RKRec.  In order to be written correctly
	 * the RK must conform to one of the cases below, if it does not it will need
	 * to be handled as a NUMBER.  Note the last 2 bits of each record need to be
	 * set as an identifier for what type of RK record we are writing.  Types 0 and 1
	 * are stored as a 32 bit modified float, with the mantissa of a double and the final
	 * 2 bits changed to the RK type.  Types 2 and 3 are stored as a 30 bit integer with the
	 * final 2 bits as RK type.  Also note types 1 and 2 are stored in the excel file as x*100.
	 */
	protected void setRKVal( double d )
	{
		byte[] b = Rk.getRkBytes( d );    // returns a 5-byte structure
		rkType = b[4];    // the last byte is the Rk type

		switch( rkType )
		{
			case Rk.RK_FP:
				isFPNumber = true;
				break;
			case Rk.RK_INT:
				isIntNumber = true;
				break;
			case Rk.RK_FP_100:
				isFPNumber = true;
				break;
			case Rk.RK_INT_100:
				isIntNumber = true;
				break;
			default:
				// we need to convert this into a number rec -- not an rk.
				Rk.convertRkToNumber( this, d );
				return;        // it's not an Rk anymore so split
		}

		System.arraycopy( b, 0, getData(), 6, 4 );
		init( getData() );

		// failsafe... if for any reason it did not work
		if( Rkdouble != d )
		{
			log.warn( "Rk.setRKVal() problem.  Fallback to floating point Number." );
			Rk.convertRkToNumber( this, d );
		}
	}

	/**
	 * Converts an RK valrec to a number record.  This allows the number format to be changed within the cell
	 */
	public static void convertRkToNumber( Rk reek, double d )
	{
		int fmt = reek.ixfe;
		String addy = reek.getCellAddress();
		Boundsheet bs = reek.getSheet();
		bs.removeCell( reek );
		BiffRec addedrec = bs.addValue( new Double( d ), addy );
		addedrec.setIxfe( reek.getIxfe() );
	}

	public int getType()
	{
		return rkType;
	}

	/** return a prototype RK record

	 offset  name        size    contents
	 ---
	 4       rw          2       Row Number
	 6       col         2       Column Number of the RK record
	 8       ixfe        2       Index to XF cell format record
	 10      rk          4       RK number

	 public static XLSRecord getPrototype(){

	 Rk newrec = new Rk();
	 // for larger records, we need to read in from a file...
	 byte[] protobytes = {(byte)0x2,
	 (byte)0x0,
	 (byte)0x0,
	 (byte)0x0,
	 (byte)0xf,
	 (byte)0x0,
	 (byte)0x1,
	 (byte)0xc0,
	 (byte)0x5e,
	 (byte) 0x40};
	 newrec.setOpcode(RK);
	 newrec.setLength((short) 0xa);
	 newrec.setData(protobytes);
	 newrec.init(newrec.getData());
	 return newrec;
	 }
	 */

	/**
	 * set the XF (format) record for this rec
	 */
	@Override
	public void setXFRecord( int i )
	{
		if( mymul != null )
		{
			byte[] b = getData();
			ixfe = i;
			byte[] newxfe = ByteTools.cLongToLEBytes( i );
			System.arraycopy( newxfe, 0, b, 4, 2 );
			mymul.updateRks();
		}
		else
		{
			super.setIxfe( i );
		}
		super.setXFRecord();
	}

	@Override
	public void close()
	{
		super.close();
		if( mymul != null )
		{
			mymul.close();
		}
		mymul = null;
	}

	// DEBUGGING - throws exception if

	/**
	 * internal debugging method
	 */
	public static void testVALUES()
	{
		for( int v = Integer.MAX_VALUE; v >= Integer.MIN_VALUE; v-- )
		{
			Rk.getRkBytes( v );    // throws Exception if converted bytes do not match original value
		}
/* problem values:			for (long v=1073741824; v > 536870912; v--)  
				Rk.getRkBytes(v);	// throws Exception if converted bytes do not match original value 
*/
	}

}