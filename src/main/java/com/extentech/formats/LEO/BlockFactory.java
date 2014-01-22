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
package com.extentech.formats.LEO;

import com.extentech.toolkit.ByteTools;

import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;

/**
 * Dutifully produces LEO file Blocks.  Complains not.
 */
public class BlockFactory
{
	/**
	 * get a new, empty Block
	 */
	public final static Block getPrototypeBlock( int type )
	{
		Block retblock = null;
		ByteBuffer dta = null;
		switch( type )
		{
			case Block.BIG:
				dta = ByteBuffer.allocate( BIGBLOCK.SIZE );
				retblock = new BIGBLOCK();
				break;
			case Block.SMALL:
				dta = ByteBuffer.allocate( SMALLBLOCK.SIZE );
				retblock = new SMALLBLOCK();
		}
		dta.order( ByteOrder.LITTLE_ENDIAN );
		//for (int x = 0; x < dta.limit(); x++)
		//	dta.put((byte) - 1);
		retblock.setBytes( dta );
		return retblock;
	}

	/**
	 * transform the byte array into Block records.
	 */
	public final static Block[] getBlocksFromOutputStream( OutputStream bbuf, int blen, int type )
	{
		int SIZE = 0;
		switch( type )
		{
			case Block.BIG:
				SIZE = 512;
				break;
			case Block.SMALL:
				SIZE = 64;
		}
		int sz = LEOFile.getSizeInBlocks( blen, SIZE );

		if( bbuf == null )
		{
			return null;
		}
		// int len = (blen-3) / SIZE;
		int len = (blen) / SIZE;
		int pos = 0, sizeDiff = 0, size = 0;

		if( (len * SIZE) < blen )
		{
			len++;
			sizeDiff = (len * SIZE) - blen;
		}
		Block[] blockarr = new Block[len];
		ByteArrayInputStream ins = null;

		// KSC: made a bit simpler upon padding situations ...
		// get ALL blockVect (512 byte chunks of file)
		for( int i = 0; i < len; i++ )
		{
			Block bbd = getPrototypeBlock( type );
			byte[] bb = bbd.getBytes();
			int filepos = i * SIZE;
			if( type == Block.SMALL )    // smallblocks don't need file offset as they are allocated differently than bigblocks
			{
				filepos = 0;
			}

			size = SIZE;
			// make simpler:
			if( blen - pos < size )
			{
				size = (blen - pos);
			}
			ins.read( bb, pos, size );
			bbd.init( ByteBuffer.wrap( bb ), i, filepos );
			pos += SIZE;
			blockarr[i] = bbd;
		}
		return blockarr;
	}

	/**
	 * transform the byte array into Block records.
	 */
	public final static Block[] getBlocksFromByteArray( byte[] bbuf, int type )
	{
		int SIZE = 0;
		switch( type )
		{
			case Block.BIG:
				SIZE = 512;
				break;
			case Block.SMALL:
				SIZE = 64;
		}
		if( bbuf == null )
		{
			return null;
		}
//		int sz = LEOFile.getSizeInBlocks(bbuf.length,SIZE);		

		int len = (bbuf.length) / SIZE;
		int pos = 0, sizeDiff = 0, size = 0;

		if( (len * SIZE) < bbuf.length )
		{
			len++;    // PAD - this most usually hits when called from buildSSAT
			sizeDiff = (len * SIZE) - bbuf.length;
			byte[] bb = new byte[sizeDiff];
			bbuf = ByteTools.append( bb, bbuf );
		}
		Block[] blockarr = new Block[len];

		int start = 1;    // for BigBlocks, skip 1st block, for small
		if( type == Block.SMALL )
		{
			start = 0;
		}
		// get ALL blockVect (512 byte chunks of file)
		for( int i = 0; i < len; i++ )
		{
			Block bbd = getPrototypeBlock( type );
			byte[] bb = bbd.getBytes();
			
/*			int filepos = (i + start) * SIZE;
			if (type==Block.SMALL)	// smallblocks don't need file offset as they are allocated differently than bigblocks
				filepos= 0;
*/
			size = SIZE;
			// make it simpler:
			if( bbuf.length - pos < size )
			{
				size = (bbuf.length - pos);    // account for leftovers (Block padding)
			}
			/*if (i + start == len) {
				size -= sizeDiff; // account for leftovers (Block padding)
			}*/
			// ExcelTools.benchmark("BlockFactory initting new Block@: " + pos + " sz: " +size + " len:" + bbuf.length);
			System.arraycopy( bbuf, pos, bb, 0, size );
			bbd.init( ByteBuffer.wrap( bb ), i, 0 );    // THIS IS A BYTEARRAY ALL OFFSETS ARE 0 filepos);
			pos += size;
			blockarr[i] = bbd;
		}
		return blockarr;
	}
}