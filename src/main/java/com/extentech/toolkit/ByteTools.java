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
package com.extentech.toolkit;

import java.io.ByteArrayOutputStream;
import java.io.DataOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.io.UnsupportedEncodingException;
import java.util.Iterator;
import java.util.List;

/**
 * Helper methods for working with byte arrays and XLS files.
 */

public final class ByteTools implements Serializable
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1220042103372057083L;

	/**
	 * Returns a string representation of the byte array.
	 *
	 * @param bt     the byte array
	 * @param offset - offset into byte array
	 * @return the string representation
	 */
	public static String getByteDump( byte[] bt, int offset )
	{
		return ByteTools.getByteDump( bt, offset, bt.length );
	}

	/**
	 * Returns a string representation of the byte array.
	 *
	 * @param bt     the byte array
	 * @param offset - offset into byte array
	 * @param len    - length of byte array segment to return
	 * @return the string representation
	 */
	public static String getByteDump( byte[] bt, int offset, int len )
	{
		if( bt == null )
		{
			return "";
		}
		StringBuffer buf = new StringBuffer();
		int every4 = 0;
		int every16 = 0;
		buf.append( "\r\n" );// start on a new line
		String offst = (Integer.toHexString( offset ));
		// now calculate where the 4 byte words should be offset so display matches...
		//int remainder = offset%16;
		while( offst.length() < 4 )
		{
			offst = "0" + offst;
		}
		buf.append( offst );
		buf.append( ":    " );
		int origOffset = offset;
		offset += 16;
		for( int i = origOffset; i < /*origOffset+*/len; ++i )
		{
			buf.append( hexits.charAt( (bt[i] >>> 4) & 0xf ) );
			buf.append( hexits.charAt( (bt[i]) & 0xf ) );
			buf.append( " " );
			every4++;
			if( every4 == 4 )
			{
				every4 = 0;
				buf.append( "  " );
				every16++;
				if( every16 == 4 )
				{
					buf.append( "\r\n" );
					offst = (Integer.toHexString( offset ));
					while( offst.length() < 4 )
					{
						offst = "0" + offst;
					}
					buf.append( offst );
					buf.append( ":    " );
					offset += 16;
					every16 = 0;
				}
			}
		}
		return buf + "";
	}

	private static String hexits = "0123456789ABCDEF";

	/**
	 * Returns a string representation of the byte array.
	 *
	 * @param bt  the byte array
	 * @param pad whether to pad the strings so they align
	 * @return the string representation
	 */
	public static String getByteString( byte[] bt, boolean pad )
	{
		if( bt.length == 0 )
		{
			return "null";
		}
		StringBuffer ret = new StringBuffer();
		for( int x = 0; x < bt.length; x++ )
		{
			if( ((x % 8) == 0) && (x > 0) )
			{
				ret.append( "\r\n" );
			}
			//String bstr = Integer.toOctalString(bt[x]); // toBinaryString(bt[x]); // Byte.toString(bt[x]);
			String bstr = Byte.toString( bt[x] ); // toBinaryString(bt[x]); // Byte.toString(bt[x]);
			if( pad )
			{
				while( bstr.length() < 4 )
				{
					bstr = " " + bstr;
				}
			}
			ret.append( bstr );
			ret.append( "," );
		}
		ret.setLength( ret.length() - 1 );
		//ret.append();
		return ret.toString();
	}

	/**
	 * Appends one byte array to another.
	 * If either input (but not both) is null, a clone of the other will be
	 * returned. This method is guaranteed to always return an array different
	 * from either of those passed in.
	 *
	 * @param src  the array which will be appended to <code>dest</code>
	 * @param dest the array to which <code>src</code> will be appended
	 * @throws NullPointerException if both inputs are null
	 */
	public static byte[] append( byte[] src, byte[] dest )
	{
		// Deal with null input correctly
		if( src == null )
		{
			return (byte[]) dest.clone();
		}
		if( dest == null )
		{
			return (byte[]) src.clone();
		}

		int srclen = src.length;
		int destlen = dest.length;

		byte[] ret = new byte[srclen + destlen];
		System.arraycopy( dest, 0, ret, 0, destlen );
		System.arraycopy( src, 0, ret, destlen, srclen );

		return ret;
	}

	/**
	 * append one byte array to an empty array
	 * of the proper size
	 * usage:
	 * newarray = bytetool.append(sourcearray, destinationarray, position to start copy at);
	 */
	public static byte[] append( byte[] src, byte[] dest, int pos )
	{
		int srclen = src.length;
		if( dest == null )
		{
			dest = new byte[srclen];
		}
		int destlen = dest.length;
		if( destlen < srclen )
		{
			Logger.logInfo( "Your destination byte array is too small to copy into: srclen=" + String.valueOf( srclen ) + ": destlen=" + String
					.valueOf( destlen ) );
			srclen = destlen;
		}
		System.arraycopy( src, 0, dest, pos, srclen );
		return dest;
	}

	public static byte[] cLongToLEBytes( int i )
	{
		byte[] ret = new byte[4];
		ret[0] = (byte) (i & 0xff);
		ret[1] = (byte) ((i >> 8) & 0xff);
		ret[2] = (byte) ((i >> 16) & 0xff);
		ret[3] = (byte) ((i >> 24) & 0xff);
		return ret;
	}

	/**
	 * C Longs are only 32 bits, Java Longs are 64.
	 * This method converts a 32-bit 'C' long to a
	 * pair of java shorts.
	 * <p/>
	 * Also performs 'little-endian' conversion.
	 */
	public static short[] cLongToLEShorts( int x )
	{
		short[] buf = new short[2];
		short high = (short) (x >>> 16);
		short low = (short) x;
		buf[0] = low;
		buf[1] = high;
		// if(DEBUG)Logger.logInfo(Info( ("x=" + x + " high=" + high + " low=" + low );
		return buf;
	}

	public static byte[] doubleToLEByteArray( double d )
	{
		byte[] bite = new byte[8]; // A long is 8 bytes
		long l = Double.doubleToLongBits( d );
		int i;
		long t;
		t = l; // variable t will be shifted right each time thru the loop.
		for( i = bite.length - 1; i > -1; i-- )
		{ //High order byte will be in b[0]
			long irr = (t & 0xff);
			bite[i] = Integer.valueOf( (int) irr ).byteValue(); //get the last 8 bits into the byte array.
			t = t >> 8; //Shifts the long 1 byte. Same as divide by 256
		}
		byte[] ret = new byte[bite.length];
		for( int x = 0; x < bite.length; x++ )
		{
			ret[x] = bite[(bite.length - 1) - x];
		}
		return ret;
	}

	public static byte[] doubleToByteArray( double d )
	{
		byte[] bite = new byte[8]; // A long is 8 bytes
		long l = Double.doubleToLongBits( d );
		int i;
		long t;
		t = l; // variable t will be shifted right each time thru the loop.
		for( i = bite.length - 1; i > -1; i-- )
		{ //High order byte will be in b[0]
			long irr = (t & 0xff);
			bite[i] = Integer.valueOf( (int) irr ).byteValue(); //get the last 8 bits into the byte array.
			t = t >> 8; //Shifts the long 1 byte. Same as divide by 256
		}
		return bite;
	}

	/**
	 * converts and bitswaps an eight bite byte array into an IEEE double.
	 */
	public static double eightBytetoLEDouble( byte[] bite )
	{
		byte[] b = new byte[8];
		b[0] = bite[7];
		b[1] = bite[6];
		b[2] = bite[5];
		b[3] = bite[4];
		b[4] = bite[3];
		b[5] = bite[2];
		b[6] = bite[1];
		b[7] = bite[0];
		double d = 0;
		java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream( b );
		java.io.DataInputStream dis = new java.io.DataInputStream( bais );
		try
		{
			Double dbl = dis.readDouble();
			d = dbl.doubleValue();
		}
		catch( java.io.IOException e )
		{
			Logger.logInfo( "io exception in byte to Double conversion" + e );
		}
		return d;
	}

	/**
	 * converts and bitswaps an eight bite byte array into an IEEE double.
	 */
	public static long eightBytetoLELong( byte[] bite )
	{
		byte[] b = new byte[8];
		b[0] = bite[7];
		b[1] = bite[6];
		b[2] = bite[5];
		b[3] = bite[4];
		b[4] = bite[3];
		b[5] = bite[2];
		b[6] = bite[1];
		b[7] = bite[0];
		long l = 0;
		java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream( b );
		java.io.DataInputStream dis = new java.io.DataInputStream( bais );
		try
		{
			Long lg = dis.readLong();
			l = lg.longValue();
		}
		catch( java.io.IOException e )
		{
			Logger.logInfo( "io exception in byte to Double conversion" + e );
		}
		return l;
	}

	/**
	 * Get an array of bytes from a collection of byte arrays
	 * <p/>
	 * Seems slow, why 2 iterations, should be faster way?
	 */
	public static byte[] getBytes( List records )
	{
		Iterator e = records.iterator();
		int buflen = 0;
		while( e.hasNext() )
		{
			byte[] barr = (byte[]) e.next();
			buflen += barr.length;
		}

		byte[] outbytes = new byte[buflen];
		int pos = 0;
		for( int i = 0; i < records.size(); i++ )
		{
			byte[] stream = (byte[]) records.get( i );
			outbytes = append( stream, outbytes, pos );
			pos += stream.length;
		}
		return outbytes;
	}

	/*
		Makes sure unicode strings are in the correct format to match Excel's strings.
		If the string has all low order bytes as 0x0 then return original string, as we do not
		want that extra space.

	 */
	public static byte[] getExcelEncoding( String s )
	{
		byte[] strbytes = null;
		try
		{
			strbytes = s.getBytes( "UnicodeLittleUnmarked" );
		}
		catch( UnsupportedEncodingException e )
		{
			Logger.logInfo( "Error creating encoded string: " + e );
		}
		boolean unicode = false;
		for( int i = 0; i < strbytes.length; i++ )
		{
			i = i + 1;
			if( strbytes[i] != 0x0 )
			{
				unicode = true;
				i = strbytes.length;
			}

		}
		if( unicode )
		{
			//try{
			return strbytes;
			//}catch(UnsupportedEncodingException e){Logger.logInfo("Error creating encoded string: " + e);}
		}
		return s.getBytes();
	}

	/**
	 * This is a working longToLEByteArray.  I'm not sure whats up with the other ones
	 *
	 * @param l
	 * @return
	 */
	public static byte[] longToLEByteArray( long l )
	{
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		DataOutputStream dos = new DataOutputStream( bos );
		try
		{
			dos.writeLong( l );
			dos.flush();
		}
		catch( IOException e )
		{

		}
		byte[] bite = bos.toByteArray();
		byte[] b = new byte[8];
		b[0] = bite[7];
		b[1] = bite[6];
		b[2] = bite[5];
		b[3] = bite[4];
		b[4] = bite[3];
		b[5] = bite[2];
		b[6] = bite[1];
		b[7] = bite[0];
		return b;
	}

	public static boolean isUnicode( String s )
	{
		byte[] strbytes = null;
		try
		{
			strbytes = s.getBytes( "UnicodeLittleUnmarked" );
		}
		catch( UnsupportedEncodingException e )
		{
			;
		}
		for( int i = 0; i < strbytes.length; i++ )
		{
			if( strbytes[i] >= 0x7f )
			{
				return true;  // deal with non-compressible Eastern Strings
			}
			i = i + 1;
			if( strbytes[i] != 0x0 )
			{
				return true; // there is a non-zero high-byte
			}
		}
		return false;
	}

	public static byte[] longToByteArray( long l )
	{
		byte[] bite = new byte[8]; // A long is 8 bytes
		int i;
		long t;
		t = l; // variable t will be shifted right each time thru the loop.
		for( i = bite.length - 1; i > -1; i-- )
		{ //High order byte will be in b[0]
			long irr = (t & 0xff);
			bite[i] = Integer.valueOf( (int) irr ).byteValue(); //get the last 8 bits into the byte array.
			t = t >> 8; //Shifts the long 1 byte. Same as divide by 256
		}
		return bite;
	}

	/**
	 * same as readInt, but takes 4 raw bytes instead of 2 shorts
	 */
	public static int readInt( byte[] bs )
	{
		return readInt( readShort( bs[2], bs[3] ), readShort( bs[0], bs[1] ) );
	}

	/**
	 * same as readInt, but takes 4 raw bytes instead of 2 shorts
	 */
	public static int readInt( byte b1, byte b2, byte b3, byte b4 )
	{
		return readInt( readShort( b3, b4 ), readShort( b1, b2 ) );
	}

	/**
	 * Reads a 4 byte int from a byte array at the specified position
	 * and handles a little endian conversion
	 */
	public static int readInt( byte[] b, int offset )
	{
		return readInt( b[offset++], b[offset++], b[offset++], b[offset++] );
	}

	/**
	 * bit-flipping action converting a 'little-endian'
	 * pair of shorts to a 'big-endian' long.
	 * This is really a java int as it represents a C-language
	 * long value which is only 32 bits, like the java int.
	 */
	public static int readInt( int low, int high )
	{
		if( (low == 0x0) && (high == 0x0) )
		{
			return 0;
		}
		low = low & 0xffff;
		high = high & 0xffff;
		return (int) ((low << 16) | high);
	}

	/**
	 * bit-flipping action converting a 'little-endian'
	 * pair of bytes to a 'big-endian' short.  Returns an int as
	 * excel uses unsigned shorts which can exceed the boundary
	 * of a java signed short
	 */
	public static int readUnsignedShort( byte low, byte high )
	{
		return readInt( low, high, (byte) 0x0, (byte) 0x0 );
	}

	/**
	 * bit-flipping action converting a 'little-endian'
	 * pair of bytes to a 'big-endian' short.
	 * <p/>
	 * This will break if you pass it any values larger than a byte.  Will
	 * probably return a value, but I wouldn't trust it.  Fix in R2  -Rab
	 */
	public static short readShort( int low, int high )
	{
		// 2 bytes
		low = low & 0xff;
		high = high & 0xff;
		return (short) ((high << 8) | low);
	}

	/** bit-flipping action converting a 'little-endian'
	 pair of bytes to a 'big-endian' short.

	 public static short readShort(byte low, byte high)
	 {
	 return (short)(high << 8 | low);
	 }
	 */

	/**
	 * take 16-bit short apart into two 8-bit bytes.
	 */
	public static byte[] shortToLEBytes( short x )
	{
		byte[] buf = new byte[2];
		buf[1] = (byte) (x >>> 8);
		buf[0] = (byte) x;/* cast implies & 0xff */
		return buf;
	}

	public static byte[] toBEByteArray( double d )
	{
		byte[] bite = new byte[8]; // A long is 8 bytes
		long l = Double.doubleToLongBits( d );
		int i;
		long t;
		t = l; // variable t will be shifted right each time thru the loop.
		for( i = bite.length - 1; i > -1; i-- )
		{ //High order byte will be in b[0]
			long irr = (t & 0xff);
			bite[i] = Integer.valueOf( (int) irr ).byteValue(); //get the last 8 bits into the byte array.
			t = t >> 8; //Shifts the long 1 byte. Same as divide by 256
		}
		byte[] b = new byte[8];
		b[0] = bite[7];
		b[1] = bite[6];
		b[2] = bite[5];
		b[3] = bite[4];
		b[4] = bite[3];
		b[5] = bite[2];
		b[6] = bite[1];
		b[7] = bite[0];
		return b;
	}
	//  private boolean DEBUG = false;

	/**
	 * write bytes to a file
	 */
	public static void writeToFile( byte[] b, String fname )
	{
		try
		{
			FileOutputStream fos = new FileOutputStream( fname );
			fos.write( b );
			fos.flush();
			fos.close();
		}
		catch( IOException e )
		{
			Logger.logInfo( "Error writing bytes to file in ByteTools: " + e );
		}

	}

	/**
	 * C Longs are only 32 bits, Java Longs are 64.
	 * This method converts a 32-bit 'C' long to a
	 * byte array.
	 * <p/>
	 * Also performs 'little-endian' conversion.
	 */
	public byte[] cLongToLEBytesOLD( int i )
	{
		//if(true)return Integer.
		short[] sbuf = cLongToLEShorts( i );
		byte[] b1 = shortToLEBytes( sbuf[0] );
		byte[] b2 = shortToLEBytes( sbuf[1] );

		byte[] bbuf = new byte[4];
		bbuf[0] = b1[0];
		bbuf[1] = b1[1];
		bbuf[2] = b2[0];
		bbuf[3] = b2[1];

		// System.arraycopy(b1, 0, bbuf, 0, 2);
		// System.arraycopy(b2, 0, bbuf, 2, 2);
		return bbuf;
	}

	// NICK -- let's put the following in a JUnit Test.  -jm

	/** this is a good test for some of the above methods..
	 byte[] bytes;
	 java.io.ByteArrayInputStream bis = new java.io.ByteArrayInputStream(bite);
	 java.io.DataInputStream dis = new java.io.DataInputStream(bis);
	 try {
	 long lon = dis.readLong();
	 Logger.logInfo("pleeze stop here");
	 } catch (java.io.IOException e){ Logger.logInfo("io exception in byte to long conversion" + e);}
	 */
	/**
	 * generic method to update (set or clear) a short at bitNum
	 *
	 * @param set
	 * @param bitNum
	 */
	public static short updateGrBit( short grbit, boolean set, int bitNum )
	{
		switch( bitNum )
		{
			case 0:
				if( set )
				{
					grbit |= (0x1);
				}
				else
				{
					grbit &= (0xFFFE);
				}
				break;
			case 1:
				if( set )
				{
					grbit |= (0x2);
				}
				else
				{
					grbit &= (0xFFFD);
				}
				break;
			case 2:
				if( set )
				{
					grbit |= (0x4);
				}
				else
				{
					grbit &= (0xFFFB);
				}
				break;
			case 3:
				if( set )
				{
					grbit |= (0x8);
				}
				else
				{
					grbit &= (0xFFF7);
				}
				break;
			case 4:
				if( set )
				{
					grbit |= (0x10);
				}
				else
				{
					grbit &= (0xFFEF);
				}
				break;
			case 5:
				if( set )
				{
					grbit |= (0x20);
				}
				else
				{
					grbit &= (0xFFDF);
				}
				break;
			case 6:
				if( set )
				{
					grbit |= (0x40);
				}
				else
				{
					grbit &= (0xFFBF);
				}
				break;
			case 7:
				if( set )
				{
					grbit |= (0x80);
				}
				else
				{
					grbit &= (0xFF7F);
				}
				break;
			case 8:
				if( set )
				{
					grbit |= (0x100);
				}
				else
				{
					grbit &= (0xFEFF);
				}
				break;
			case 9:
				if( set )
				{
					grbit |= (0x200);
				}
				else
				{
					grbit &= (0xFDFF);
				}
				break;
			case 10:
				if( set )
				{
					grbit |= (0x400);
				}
				else
				{
					grbit &= (0xFBFF);
				}
				break;
			case 11:
				if( set )
				{
					grbit |= (0x800);
				}
				else
				{
					grbit &= (0xF7FF);
				}
				break;
			case 12:
				if( set )
				{
					grbit |= (0x1000);
				}
				else
				{
					grbit &= (0xEFFF);
				}
				break;
			case 13:
				if( set )
				{
					grbit |= (0x2000);
				}
				else
				{
					grbit &= (0xDFFF);
				}
				break;
			case 14:
				if( set )
				{
					grbit |= (0x4000);
				}
				else
				{
					grbit &= (0xBFFF);
				}
				break;
			case 15:
				if( set )
				{
					grbit |= (0x8000);
				}
				else
				{
					grbit &= (0x7FFF);
				}
				break;
		}
		return grbit;
	}
}

     