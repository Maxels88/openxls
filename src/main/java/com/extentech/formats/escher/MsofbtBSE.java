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
package com.extentech.formats.escher;

import com.extentech.formats.XLS.MSODrawingConstants;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.MD4Digest;

//0xF007
public class MsofbtBSE extends EscherRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -9072016434347371029L;
	private byte[] imageData;
	private int imageType;
	private int numShapes;
	private int cRef = 1;        // default reference count is 1; 0= hidden; > 1= multiple refs to same image   20071120 KSC

	public MsofbtBSE( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	@Override
	protected byte[] getData()
	{
		byte[] imgHeader = new byte[61];    // BSE header = 36 bytes, BLIP record header = 24 bytes, then imageData bytes follow
		imgHeader[0] = (byte) imageType;   //btWin32
		imgHeader[1] = (byte) imageType;   //btMac

		// 20071004 KSC: Unknown type of record, does not follow below rules 
		// and will cause file error ... 
		// TODO: Research more 
		if( imageType == 0 )
		{
			byte[] retData = new byte[36];
			retData[18] = -1;    // tag
			setLength( retData.length );
			return retData;
		}
		//The digest of image data should be here. 2-17
		MD4Digest md4Digest = new MD4Digest();
		byte[] digest = md4Digest.getDigest( imageData );
		System.arraycopy( digest, 0, imgHeader, 2, 16 );

		imgHeader[18] = -1;   //First tag byte is always -1
		imgHeader[19] = 0;    //Second tag byte is always 0

/* 20071119 KSC: original code; this was wrong    		
		int mod = (imageData.length+25)%MAXROWS_BIFF8;
		imgHeader[20] = (byte)((0x000000FF&mod));    
		imgHeader[21] = (byte)((0x0000FF00&mod)>>8);   
		
		//isze is always 
		int size = MAXROWS_BIFF8 + (int)((imageData.length-36)/MAXROWS_BIFF8);
		byte[] tempBytes = ByteTools.cLongToLEBytes(size);
		imgHeader[22]=tempBytes[0];
		imgHeader[23]=tempBytes[1];
		imgHeader[24]=tempBytes[2];
		imgHeader[25]=tempBytes[3];
		
		//		cRef is always zero
		tempBytes = ByteTools.cLongToLEBytes(0);
		imgHeader[26]=tempBytes[0];
		imgHeader[27]=tempBytes[1];
		imgHeader[28]=tempBytes[2];
		imgHeader[29]=tempBytes[3];
		
		//foDelay is always zero
		tempBytes = ByteTools.cLongToLEBytes(0);
		imgHeader[30]=tempBytes[0];
		imgHeader[31]=tempBytes[1];
		imgHeader[32]=tempBytes[2];
		imgHeader[33]=tempBytes[3];
		
		imgHeader[34]=(byte)1;    //usage is numShapes
		imgHeader[35]=(byte)0;    //cbName is always zero
*/
		// new code  
		// Size of data + 25 (for header stuff, I assume) (4 bytes)
		int sz = imageData.length + 25;
		byte[] tempBytes = ByteTools.cLongToLEBytes( sz );
		imgHeader[20] = tempBytes[0];
		imgHeader[21] = tempBytes[1];
		imgHeader[22] = tempBytes[2];
		imgHeader[23] = tempBytes[3];

		// cRef (4 bytes) = reference count; 1 unless it's referenced more than once ... :) 0 if hidden
		tempBytes = ByteTools.cLongToLEBytes( cRef );
		imgHeader[24] = tempBytes[0];
		imgHeader[25] = tempBytes[1];
		imgHeader[26] = tempBytes[2];
		imgHeader[27] = tempBytes[3];

		// foDelay is always zero	= image bytes are not in delay stream
		tempBytes = ByteTools.cLongToLEBytes( 0 );
		imgHeader[28] = tempBytes[0];
		imgHeader[29] = tempBytes[1];
		imgHeader[30] = tempBytes[2];
		imgHeader[31] = tempBytes[3];

		imgHeader[32] = 0;    //usage - should be 0=default usage
		imgHeader[33] = 0;    //cbName is always zero = no name following this header
		// bytes 34 and 35 are unused at this point and should be 0

		// ********************************************************************************
		// BLIP RECORD follows BSE Header UNLESS file delayOffset is > 0 or cbNameLen > 0
		// ********************************************************************************

		// BLIP RECORD for Metafile/PICT Blips (msobiEMF, msobiWMF, or msobiPICT): 
		/* The secondary, or data, UID - should always be set. */
		// BYTE  m_rgbUid[16];
		/* The primary UID - this defaults to 0, in which case the primary ID is
		   that of the internal data. NOTE!: The primary UID is only saved to disk
		   if (blip_instance ^ blip_signature == 1). Blip_instance is MSOFBH.inst and 
		   blip_signature is one of the values defined in MSOBI */
		//BYTE  m_rgbUidPrimary[16]; // optional based on the above check

		/* Metafile Blip overhead = 34 bytes. m_cb gives the number of
		   bytes required to store an uncompressed version of the file, m_cbSave
		   is the compressed size.  m_mfBounds gives the boundary of all the
		   drawing calls within the metafile (this may just be the bounding box
		   or it may allow some whitespace, for a WMF this comes from the
		   SetWindowOrg and SetWindowExt records of the metafile). 
		  */
		/*
		int           m_cb;           // Cache of the metafile size
		RECT          m_rcBounds;     // Boundary of metafile drawing commands
		POINT         m_ptSize;       // Size of metafile in EMUs
		int           m_cbSave;       // Cache of saved size (size of m_pvBits)
		BYTE          m_fCompression; // MSOBLIPCOMPRESSION
		BYTE          m_fFilter;      // always msofilterNone
		void         *m_pvBits;       // Compressed bits of metafile.		
		*/

		// BLIP RECORD for Bitmap Blips (msobiJPEG, msobiPNG, or msobiDIB) :
		/*
		 * They have the same UID header as described in the Metafile Blip case. The data after the header is just a single BYTE "tag"
		 *  value and is followed by the compressed data of the bitmap in the relevant format (JFIF or PNG, bytes as would be stored in a file). 
		 *  For the msobiDIB format, the data is in the standard DIB format as a BITMAPINFO ER, BITMAPCORE ER or BITMAPV4 ER followed by 
		 *  the color map (DIB_RGB_COLORS) and the bits. This data is not compressed (the format is used for very small DIB bitmaps only).

		To determine where the bits are located, refer to the following header:

		/* The secondary, or data, UID - should always be set. */
		/*
		BYTE  m_rgbUid[16];
		/* The primary UID - this defaults to 0, in which case the primary ID is
		   that of the internal data. NOTE!: The primary UID is only saved to disk
		   if (blip_instance ^ blip_signature == 1). Blip_instance is MSOFBH.finst and 
		   blip_signature is one of the values defined in MSOBI
		  */
		/*
		BYTE  m_rgbUidPrimary[16];    // optional based on the above check
		BYTE  m_bTag;            
		void  *m_pvBits;              // raster bits of the blip.
		*/

		// 20080226 KSC: instead of handling only two image types, handle them all :)
		int inst1 = MSODrawingConstants.msobiJPEG;
		switch( imageType )
		{
			case MSODrawingConstants.IMAGE_TYPE_GIF:
				// TODO: ???
				break;
			case MSODrawingConstants.IMAGE_TYPE_EMF:
				inst1 = MSODrawingConstants.msobiEMF;
				break;
			case MSODrawingConstants.IMAGE_TYPE_WMF:
				inst1 = MSODrawingConstants.msobiWMF;
				break;
			case MSODrawingConstants.IMAGE_TYPE_PICT:
				inst1 = MSODrawingConstants.msobiPICT;
				break;
			case MSODrawingConstants.IMAGE_TYPE_JPG:
				inst1 = MSODrawingConstants.msobiJPEG;
				break;
			case MSODrawingConstants.IMAGE_TYPE_PNG:
				inst1 = MSODrawingConstants.msobiPNG;
				break;
			case MSODrawingConstants.IMAGE_TYPE_DIB:
				inst1 = MSODrawingConstants.msobiDIB;
				break;
		}
		//int inst1=(imageType==5)?MsodrawingConstants.msobiJPEG:MsodrawingConstants.msobiPNG;		
		int version1 = 0;
		int fbt1 = MSODrawingConstants.msofbtBlipFirst + imageType;
		int len1 = imageData.length + 17;
		imgHeader[36] = (byte) (((0x0F & inst1) << 4) | (0x0F & version1));
		imgHeader[37] = (byte) ((0xFF0 & inst1) >> 4);
		imgHeader[38] = (byte) ((0xFF & fbt1));
		imgHeader[39] = (byte) ((0xFF00 & fbt1) >> 8);

		//length with size + 25;
		tempBytes = ByteTools.cLongToLEBytes( len1 );
		imgHeader[40] = tempBytes[0];
		imgHeader[41] = tempBytes[1];
		imgHeader[42] = tempBytes[2];
		imgHeader[43] = tempBytes[3];

		///Here will again be the md4 digest of the image data 44-60
		System.arraycopy( digest, 0, imgHeader, 44, 16 );
		imgHeader[60] = (byte) 0xff;   //marker to indicate start of image data
		int len = 61;
		byte[] retData = new byte[len + imageData.length];

		System.arraycopy( imgHeader, 0, retData, 0, len );
		System.arraycopy( imageData, 0, retData, len, imageData.length );

		setLength( retData.length );
		return retData;
	}

	public void setImageData( byte[] d )
	{
		imageData = d;
	}

	/**
	 * set the reference count for this image Data rec
	 *
	 * @param cRef
	 */
	public void setRefCount( int cRef )
	{
		this.cRef = cRef;
	}

	/**
	 * returns the ref count for this image Data Rec
	 *
	 * @return
	 */
	public int getRefCount()
	{
		return cRef;
	}

	public int getImageType()
	{
		return imageType;
	}

	public void setImageType( int imageType )
	{
		this.imageType = imageType;
	}

	public int getNumShapes()
	{
		return numShapes;
	}

	public void setNumShapes( int numShapes )
	{
		this.numShapes = numShapes;
	}

}
