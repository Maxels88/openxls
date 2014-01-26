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
package org.openxls.toolkit;

/**
 * implementation of MD4 as RFC 1320 by R. Rivest, MIT Laboratory for
 * Computer Science and RSA Data Security, Inc.
 * <p/>
 * <b>NOTE</b>: This algorithm is only included for backwards compatability
 * with legacy applications, it's not secure, don't use it for anything new!
 */

public class MD4Digest
{

	private int H1;
	private int H2;
	private int H3;
	private int H4;         // IV's
	private int[] X = new int[16];
	private int xOff;
	private static final int BYTE_LENGTH = 64;
	private byte[] xBuf;
	private int xBufOff;
	private long byteCount;

	/**
	 * Standard constructor
	 */
	public MD4Digest()
	{
		xBuf = new byte[4];
		xBufOff = 0;
		reset();
	}

	public byte[] getDigest( byte[] data )
	{

		update( data, 0, data.length );
		byte[] digest = new byte[16];
		doFinal( digest, 0 );

		return digest;

	}

	protected void processWord( byte[] in, int inOff )
	{
		X[xOff++] = (in[inOff] & 0xff) | ((in[inOff + 1] & 0xff) << 8) | ((in[inOff + 2] & 0xff) << 16) | ((in[inOff + 3] & 0xff) << 24);

		if( xOff == 16 )
		{
			processBlock();
		}
	}

	protected void processLength( long bitLength )
	{
		if( xOff > 14 )
		{
			processBlock();
		}

		X[14] = (int) (bitLength & 0xffffffff);
		X[15] = (int) (bitLength >>> 32);
	}

	private static void unpackWord( int word, byte[] out, int outOff )
	{
		out[outOff] = (byte) word;
		out[outOff + 1] = (byte) (word >>> 8);
		out[outOff + 2] = (byte) (word >>> 16);
		out[outOff + 3] = (byte) (word >>> 24);
	}

	public int doFinal( byte[] out, int outOff )
	{
		finish();

		unpackWord( H1, out, outOff );
		unpackWord( H2, out, outOff + 4 );
		unpackWord( H3, out, outOff + 8 );
		unpackWord( H4, out, outOff + 12 );

		reset();

		return 16;
	}

	/**
	 * reset the chaining variables to the IV values.
	 */
	public void reset()
	{
		byteCount = 0;

		xBufOff = 0;
		for( int i = 0; i < xBuf.length; i++ )
		{
			xBuf[i] = 0;
		}

		H1 = 0x67452301;
		H2 = 0xefcdab89;
		H3 = 0x98badcfe;
		H4 = 0x10325476;

		xOff = 0;

		for( int i = 0; i != X.length; i++ )
		{
			X[i] = 0;
		}
	}

	//
	// round 1 left rotates
	//
	private static final int S11 = 3;
	private static final int S12 = 7;
	private static final int S13 = 11;
	private static final int S14 = 19;

	//
	// round 2 left rotates
	//
	private static final int S21 = 3;
	private static final int S22 = 5;
	private static final int S23 = 9;
	private static final int S24 = 13;

	//
	// round 3 left rotates
	//
	private static final int S31 = 3;
	private static final int S32 = 9;
	private static final int S33 = 11;
	private static final int S34 = 15;

	/*
	 * rotate int x left n bits.
	 */
	private static int rotateLeft( int x, int n )
	{
		return (x << n) | (x >>> (32 - n));
	}

	/*
	 * F, G, H and I are the basic MD4 functions.
	 */
	private static int F( int u, int v, int w )
	{
		return (u & v) | (~u & w);
	}

	private static int G( int u, int v, int w )
	{
		return (u & v) | (u & w) | (v & w);
	}

	private static int H( int u, int v, int w )
	{
		return u ^ v ^ w;
	}

	protected void processBlock()
	{
		int a = H1;
		int b = H2;
		int c = H3;
		int d = H4;

		//
		// Round 1 - F cycle, 16 times.
		//
		a = rotateLeft( a + F( b, c, d ) + X[0], S11 );
		d = rotateLeft( d + F( a, b, c ) + X[1], S12 );
		c = rotateLeft( c + F( d, a, b ) + X[2], S13 );
		b = rotateLeft( b + F( c, d, a ) + X[3], S14 );
		a = rotateLeft( a + F( b, c, d ) + X[4], S11 );
		d = rotateLeft( d + F( a, b, c ) + X[5], S12 );
		c = rotateLeft( c + F( d, a, b ) + X[6], S13 );
		b = rotateLeft( b + F( c, d, a ) + X[7], S14 );
		a = rotateLeft( a + F( b, c, d ) + X[8], S11 );
		d = rotateLeft( d + F( a, b, c ) + X[9], S12 );
		c = rotateLeft( c + F( d, a, b ) + X[10], S13 );
		b = rotateLeft( b + F( c, d, a ) + X[11], S14 );
		a = rotateLeft( a + F( b, c, d ) + X[12], S11 );
		d = rotateLeft( d + F( a, b, c ) + X[13], S12 );
		c = rotateLeft( c + F( d, a, b ) + X[14], S13 );
		b = rotateLeft( b + F( c, d, a ) + X[15], S14 );

		//
		// Round 2 - G cycle, 16 times.
		//
		a = rotateLeft( a + G( b, c, d ) + X[0] + 0x5a827999, S21 );
		d = rotateLeft( d + G( a, b, c ) + X[4] + 0x5a827999, S22 );
		c = rotateLeft( c + G( d, a, b ) + X[8] + 0x5a827999, S23 );
		b = rotateLeft( b + G( c, d, a ) + X[12] + 0x5a827999, S24 );
		a = rotateLeft( a + G( b, c, d ) + X[1] + 0x5a827999, S21 );
		d = rotateLeft( d + G( a, b, c ) + X[5] + 0x5a827999, S22 );
		c = rotateLeft( c + G( d, a, b ) + X[9] + 0x5a827999, S23 );
		b = rotateLeft( b + G( c, d, a ) + X[13] + 0x5a827999, S24 );
		a = rotateLeft( a + G( b, c, d ) + X[2] + 0x5a827999, S21 );
		d = rotateLeft( d + G( a, b, c ) + X[6] + 0x5a827999, S22 );
		c = rotateLeft( c + G( d, a, b ) + X[10] + 0x5a827999, S23 );
		b = rotateLeft( b + G( c, d, a ) + X[14] + 0x5a827999, S24 );
		a = rotateLeft( a + G( b, c, d ) + X[3] + 0x5a827999, S21 );
		d = rotateLeft( d + G( a, b, c ) + X[7] + 0x5a827999, S22 );
		c = rotateLeft( c + G( d, a, b ) + X[11] + 0x5a827999, S23 );
		b = rotateLeft( b + G( c, d, a ) + X[15] + 0x5a827999, S24 );

		//
		// Round 3 - H cycle, 16 times.
		//
		a = rotateLeft( a + H( b, c, d ) + X[0] + 0x6ed9eba1, S31 );
		d = rotateLeft( d + H( a, b, c ) + X[8] + 0x6ed9eba1, S32 );
		c = rotateLeft( c + H( d, a, b ) + X[4] + 0x6ed9eba1, S33 );
		b = rotateLeft( b + H( c, d, a ) + X[12] + 0x6ed9eba1, S34 );
		a = rotateLeft( a + H( b, c, d ) + X[2] + 0x6ed9eba1, S31 );
		d = rotateLeft( d + H( a, b, c ) + X[10] + 0x6ed9eba1, S32 );
		c = rotateLeft( c + H( d, a, b ) + X[6] + 0x6ed9eba1, S33 );
		b = rotateLeft( b + H( c, d, a ) + X[14] + 0x6ed9eba1, S34 );
		a = rotateLeft( a + H( b, c, d ) + X[1] + 0x6ed9eba1, S31 );
		d = rotateLeft( d + H( a, b, c ) + X[9] + 0x6ed9eba1, S32 );
		c = rotateLeft( c + H( d, a, b ) + X[5] + 0x6ed9eba1, S33 );
		b = rotateLeft( b + H( c, d, a ) + X[13] + 0x6ed9eba1, S34 );
		a = rotateLeft( a + H( b, c, d ) + X[3] + 0x6ed9eba1, S31 );
		d = rotateLeft( d + H( a, b, c ) + X[11] + 0x6ed9eba1, S32 );
		c = rotateLeft( c + H( d, a, b ) + X[7] + 0x6ed9eba1, S33 );
		b = rotateLeft( b + H( c, d, a ) + X[15] + 0x6ed9eba1, S34 );

		H1 += a;
		H2 += b;
		H3 += c;
		H4 += d;

		//
		// reset the offset and clean out the word buffer.
		//
		xOff = 0;
		for( int i = 0; i != X.length; i++ )
		{
			X[i] = 0;
		}
	}

	public void update( byte in )
	{
		xBuf[xBufOff++] = in;

		if( xBufOff == xBuf.length )
		{
			processWord( xBuf, 0 );
			xBufOff = 0;
		}

		byteCount++;
	}

	public void update( byte[] in, int inOff, int len )
	{
		//
		// fill the current word
		//
		while( (xBufOff != 0) && (len > 0) )
		{
			update( in[inOff] );

			inOff++;
			len--;
		}

		//
		// process whole words.
		//
		while( len > xBuf.length )
		{
			processWord( in, inOff );

			inOff += xBuf.length;
			len -= xBuf.length;
			byteCount += xBuf.length;
		}

		//
		// load in the remainder.
		//
		while( len > 0 )
		{
			update( in[inOff] );

			inOff++;
			len--;
		}
	}

	public void finish()
	{
		long bitLength = (byteCount << 3);

		//
		// add the pad bytes.
		//
		update( (byte) 128 );

		while( xBufOff != 0 )
		{
			update( (byte) 0 );
		}

		processLength( bitLength );

		processBlock();
	}

	public static int getByteLength()
	{
		return BYTE_LENGTH;
	}

}






