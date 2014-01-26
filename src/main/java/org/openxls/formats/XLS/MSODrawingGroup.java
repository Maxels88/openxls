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

import org.openxls.ExtenXLS.ImageHandle;
import org.openxls.formats.escher.MsofbtBSE;
import org.openxls.formats.escher.MsofbtBstoreContainer;
import org.openxls.formats.escher.MsofbtDgg;
import org.openxls.formats.escher.MsofbtDggContainer;
import org.openxls.formats.escher.MsofbtOPT;
import org.openxls.formats.escher.MsofbtSplitMenuColors;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * <b>MSODrawingGroup: MS Office Drawing Group (EBh)</b><br>
 * <p/>
 * These records contain only data.<p><pre>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       rgMSODrGr    var    MSO Drawing Group Data
 * <p/>
 * </p></pre>
 * <p/>
 * <p/>
 * There is only one drawing group per client document (=MSOFBTDGGCONTAINER, 0xF000 ?).
 * OfficeArtDggContainer:
 * rh (8 bytes): An OfficeArtRecordHeader structure, that specifies the header for this record. The following table specifies the subfields.
 * rh.recVer			A value that MUST be 0xF.
 * rh.recInstance		A value that MUST be 0x000.
 * rh.recType			A value that MUST be 0xF000.
 * rh.recLen			 An unsigned integer specifying the number of bytes following the header that contain document-wide file records.
 * drawingGroup (variable): An OfficeArtFDGGBlock record, that specifies document-wide information about all the drawings that are saved in the file.
 * blipStore (variable): An OfficeArtBStoreContainer record, that specifies the container for all the BLIPs that are used in all the drawings in the parent document.
 * drawingPrimaryOptions (variable): An OfficeArtFOPT record, that specifies the default properties for all drawing objects that are contained in all the drawings in the parent document.
 * drawingTertiaryOptions (variable): An OfficeArtTertiaryFOPT record, that specifies the default properties for all the drawing objects that are contained in all the drawings in the parent document.
 * colorMRU (variable): An OfficeArtColorMRUContainer record, that specifies the most recently used custom colors.
 * splitColors (variable): An OfficeArtSplitMenuColorContainer record, that specifies a container for the colors that were most recently used to format shapes.
 * <p/>
 * <p/>
 * Drawing groups contain drawings.  	(= numDrawings)
 * Drawings in turn contain shapes that are the objects that actually mark a page. (= numShapes)
 * --Each drawing has a collection of rules that govern the shapes in the drawing
 * Shape store their properties in a property table (MSOFBTOPT record of Msodrawing)
 * The actual pictures and images are kept in a separate collection so can load and save separately
 * <p/>
 * Records that are required in the MSODrawingGroup:
 * MSOFBTDGG-		Drawing Group Record- holds total # shapes saved + last or max SPID (shapeID) + number of IDclusters(FIDCLs) + total # drawings saved
 * MSOFBTOPT-		Property Table Record- Default properties of newly created shapes (can be 0'd)
 * MSOFBTBSTORECONTAINER-
 * MSOFBTBSE-		BLIP Store Entry- holds image type, id, size, index, len of blip name ...
 *
 * @see MSODrawing
 */
public final class MSODrawingGroup extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( MSODrawingGroup.class );
	// 20070914 KSC: Save drawing recs here
	private AbstractList msoRecs = new ArrayList();

	// moved from Boundsheet + renamed
	public void addMsodrawingrec( MSODrawing rec )
	{
		msoRecs.add( rec );
	}

	public boolean dirtyflag = false;

	/**
	 * loop through all the Msodrawing recs and return the next valid SPID
	 *
	 * @return
	 */
	public int getNextMsoSPID()
	{
		int spid = 0;
		for( int i = 0; i < msoRecs.size(); i++ )
		{
			spid = Math.max( ((MSODrawing) msoRecs.get( i )).getSPID(), spid );
		}
		return spid + 1;
	}

	/**
	 * remove linked MsoDrawing rec from this drawing group + update image references if necessary
	 * NOTE THIS IS STILL EXPERIMENTAL; MUST BE TESTED WITH A VARIETY OF SCENARIOS
	 */
	public void removeMsodrawingrec( MSODrawing rec, Boundsheet sheet, boolean removeObjRec )
	{
		int imgIdx = rec.getImageIndex() - 1;
		int refCnt = getCRef( imgIdx );
		boolean wasHeader = rec.bIsHeader;
		if( refCnt > 0 )
		{
			decCRef( imgIdx );
		}
		msoRecs.remove( rec );
		updateRecord();    // update msodg rec
		int idx = rec.getRecordIndex();    // chart mso's have been taken out of streamer so idx will be -1
		if( idx > -1 )
		{
			sheet.removeRecFromVec( rec );    // remove Mso rec
			if( removeObjRec ) // also remove associated obj rec 20080804 KSC
			{
				sheet.removeRecFromVec( idx );    // also remove linked Obj record
			}
		}

		if( getMsodrawingrecs().size() == 0 )
		{    // no more drawing recs, delete this msodg
			sheet.removeRecFromVec( this );
			getWorkBook().msodg = null;
			// TODO: Unsure if there are other circumstances where MsodrawingSelection should be removed ... watch out for it
			// KSC: TODO: Necessary????  Appears so for delete chart ...(WHY??????)
			BiffRec b = sheet.getSheetRec( MSODRAWINGSELECTION );
			if( b != null )
			{
				sheet.removeRecFromVec( b );
			}
		}
		else
		{
			if( wasHeader )
			{ // we just removed the header; set 1st one to it
				MSODrawing mso = null;
				for( int z = 0; z < getMsodrawingrecs().size(); z++ )
				{
					mso = (MSODrawing) getMsodrawingrecs().get( z );
					if( mso.getSheet().equals( sheet ) && mso.isShape )
					{
						mso.setIsHeader();    // make this one the header rec
						break;
					}
				}
			}
			wkbook.updateMsodrawingHeaderRec( sheet );
		}
	}

	/**
	 * return the Msodrawing header record for the given sheet
	 *
	 * @param bs
	 * @return
	 */
	public MSODrawing getMsoHeaderRec( Boundsheet bs )
	{
		for( int i = 0; i < msoRecs.size(); i++ )
		{
			MSODrawing msd = (MSODrawing) msoRecs.get( i );    // get index of first Msodrawing rec for this sheet
			// always: the 1st msodrawing rec for the sheet contains the # information ...
			if( msd.getSheet().equals( bs ) )
			{ // got it!
				if( msd.isHeader() )
				{
					return msd;
				} //else 
				//log.error("WorkBook.updateMsodrawingHeaderRec:  Header Record should be first rec in group.");
//				break;
			}
		}
		return null;
	}

	public AbstractList getMsodrawingrecs()
	{
		return msoRecs;
	}

	int spidMax = 1024;
	int numIdClusters = 2;
	int numShapes = 1;
	int numDrawings = 1;

	private ArrayList imageData = new ArrayList();
	private ArrayList imageType = new ArrayList();  // parallel array with imageData
	private ArrayList cRef = new ArrayList();    // 20071120 KSC: keep track of reference count for image data

	@Override
	public void init()
	{
		super.init();
		data = super.getData();
	}

	/* 20070813 KSC: These prototype bytes works for both Images and Charts */
	public byte[] PROTOTYPE_BYTES = {
			15,
			0,
			0,
			-16,
			82,
			0,
			0,
			0,
			0,
			0,
			6,
			-16,
			24,
			0,
			0,
			0,
			2,
			4,
			0,
			0,
			2,
			0,
			0,
			0,
			2,
			0,
			0,
			0,
			1,
			0,
			0,
			0,
			1,
			0,
			0,
			0,
			2,
			0,
			0,
			0,
			51,
			0,
			11,
			-16,
			18,
			0,
			0,
			0,
			-65,
			0,
			8,
			0,
			8,
			0,
			-127,
			1,
			9,
			0,
			0,
			8,
			-64,
			1,
			64,
			0,
			0,
			8,
			64,
			0,
			30,
			-15,
			16,
			0,
			0,
			0,
			13,
			0,
			0,
			8,
			12,
			0,
			0,
			8,
			23,
			0,
			0,
			8,
			-9,
			0,
			0,
			16
	};

	public static XLSRecord getPrototype()
	{
		MSODrawingGroup grp = new MSODrawingGroup();
		grp.setOpcode( MSODRAWINGGROUP );
		grp.setData( grp.PROTOTYPE_BYTES );
		grp.init();
		return grp;
	}

	// Add associated recs necessary for Msodrawing ...
	public void initNewMSODrawingGroup()
	{
		// add new msodg rec to stream (just before sst)
		int index = streamer.getRecordIndex( XLSConstants.SST );
		// add unknown record that appears just before MSODrawingGroup
		XLSRecord rec = new XLSRecord();
		rec.setOpcode( (short) 0x1C1 );
		rec.setData( PROTOTYPE_1C1 );
		streamer.addRecordAt( rec, index++ );
		// add MSODrawingGroup
		streamer.addRecordAt( this, index );
		// also need msymystery record + msoselection ...
		Boundsheet[] b = getWorkBook().getWorkSheets();
		for( Boundsheet aB : b )
		{
			int z = aB.getIndexOf( PHONETIC );
			if( z == -1 )
			{
				Phonetic p = new Phonetic();
				p.setData( p.PROTOTYPE_BYTES );
				p.setOpcode( XLSRecord.PHONETIC );
				p.setStreamer( getStreamer() );
				aB.insertSheetRecordAt( p, aB.getIndexOf( SELECTION ) + 1 );
			}
/* truly necessary???    		if (i==0) { // msodrawingselection only for 1st sheet???????
         		Msodrawingselection msoSelection = new Msodrawingselection();
         		msoSelection.setData(msoSelection.PROTOTYPE_BYTES);
         		msoSelection.setOpcode(XLSRecord.MSODRAWINGSELECTION);
         		msoSelection.setDebugLevel(this.DEBUGLEVEL);
         		msoSelection.setStreamer(this.getStreamer());
         		b[i].insertSheetRecordAt(msoSelection, b[i].getIndexOf(Window2.class));
    		}
*/
		}
	}

	/**
	 *
	 */
	private static final long serialVersionUID = 2378100973014157878L;

	/**
	 * The XF record can either be a style XF or a BiffRec XF.
	 */
	/*These are prototype bytes for record 0x1c1 and 0x863 that seem to accompany when there is MSODrawingGroup data*/
	public byte[] PROTOTYPE_1C1 = { -63, 1, 0, 0, -128, 56, 1, 0 };
	public byte[] PROTOTYPE_863 = { 99, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 21, 0, 0, 0, 0, 0, 0, 0, -46 };

	/**
	 * Parse the MSODrawingGroup bytes and generate state vars:
	 * imageData, imageType, cRef
	 * spidMax, numIdClusters, numDrawings, numShapes
	 */
	public void parse()
	{
		imageData.clear();
		imageType.clear();
		cRef.clear();
		//data = getBytes();
		if( data == null )
		{
			return; // no data!
		}
		
		/*
		 * This represents the MSOFBTDGGCONTAINER record (0xFOOO) which is the Drawing Group Container
		 * 
		 * 
		 * 
		 * The MSOFBTDGGCONTAINER Contains the following records (some are optional):
		 * 			 rh.recVer			A value that MUST be 0xF.
					 rh.recInstance		A value that MUST be 0x000.
					 rh.recType			A value that MUST be 0xF000.
					 rh.recLen			variable
		 * 		MSOFBTDGG (0xF006)		Drawing Group Record, contains number of shapes, drawings and id clusters 
		 * 		MSOFBTCLSID (0xF016) 	Clipboard format (optional)
		 * 		MSOFBTOPT	(0xF00B)	Property Table Record - default props of newly created shapes; only the properties that differ from
		 * 								the per-property defaults are saved.  Format is same as Msodrawing.MSOFBTOPT format
		 *		MSOFBTCOLORMRU (0xF11A)	MRU Color swatch ...	
		 *		MSOFBTSPLITMENUCOLORS (0xF11E)	MRU colors of the top-level ..split menus 
		 *		MSOFBTBSTORECONTAINER (0xF001)	An array of BLIP Store Entry (BSE) Records; Each shape indexes into this array for the BLIP they use 	
		 *		MSOFBTBSE (0xF007)		File BLIP Store Entry Recod FBSE record; Encodes type of BLIP + size + ID + ref. count + file offset ... 
		 *		MSOFBTBLIP (0xF018) 		
		 */
		/* BLIP TYPE ENUM 
		 * msoblipERROR= 0
		 * msoblipUNKNOWN,
		 * msoblipEMF,	// enhanced meta file
		 * msoblipWMF,	// windows meta file
		 * msoblipPICT	// MAC pic
		 * msoblipJPEG,
		 * msoblipPNG,
		 * msoblipDIB,
		 * msoblipFirstClient=32,	// first client-defined BLIP type
		 * msoblipLastClient= 255	// last ""
		 */
		try
		{
			byte[] buf;
			ByteArrayInputStream bis = new ByteArrayInputStream( data );
			while( bis.available() > 0 )
			{
				buf = new byte[8];
				int read = bis.read( buf, 0, 8 );
				int version = (0xF & buf[0]);
				int inst = ((0xFF & buf[1]) << 4) | ((0xF0 & buf[0]) >> 4);
				int fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
				int len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );

				//System.out.println("fbt:"+Integer.toHexString(fbt)+";len:"+len);
				if( fbt < 0xF004 )
				{
					continue;    // under 0xF005 are container recs; we just parse the atoms for needed info ...
				}

				// parse record denoted by fbt
				buf = new byte[len];
				read = bis.read( buf, 0, len );

				switch( fbt )
				{
					case MSODrawingConstants.MSOFBTDGG:    //0xf006:		// MSOFBTDGG - Drawing Group Record
						//  rh.recVer			A value that MUST be 0x0.
						// 	rh.recInstance		A value that MUST be 0x000.
						//	rh.recType			A value that MUST be 0xF006.
						//	rh.recLen			A value that MUST be 0x00000010 + ((head.cidcl - 1) * 0x00000008)
						// 	head (16 bytes): An OfficeArtFDGG record, that specifies document-wide information.
						//		
						//  Rgidcl (variable): An array of OfficeArtIDCL elements, specifying file identifier clusters that are used in the drawing. The number of elements in the array is specified by (head.cidcl – 1). 
						spidMax = ByteTools.readInt( buf[0], buf[1], buf[2], buf[3] );        // maximum shape ID
						numIdClusters = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );    // number of ID clusters
						numShapes = ByteTools.readInt( buf[8], buf[9], buf[10], buf[11] );    // total number of shapes saved
						numDrawings = ByteTools.readInt( buf[12], buf[13], buf[14], buf[15] );    // total number of drawings saved
						// the fixed part is followed by and array of ID clusters used internally for the translation of SPIDs to MSHHSPs (Shape Handles)
						break;

					case MSODrawingConstants.MSOFBTBSE:        //0xf007:		// File BLIP Store Entry Record (FBSE)
						byte[] imgBuf = getImageBytesWithBuffer( buf );
						// strip buffer from data image bytes						
//						numShapes--;	// 20070914 KSC: WHY?? -- (not a real shape? -jm)
						break;
				}
			}
		}
		catch( Exception e )
		{
			log.error( "Msodrawingroup parse error.", e );
		}
	}

	private byte[] getImageBytesWithBuffer( byte[] buf )
	{
//		 Each BLIP in the BStore is serialized to a FBSE record; 
		// btWin32, btMacOS, rgbUid[16] = identifier of blip, tag, size=BLIP size in stream
		int btWin32 = buf[0];
		imageType.add( btWin32 );
		/* parse header for testing purposes
		// inst= encoded type		
		int pos= 1;		
		int btMac= buf[pos++];		
		byte[] UID= new byte[16];
		byte[] UID2= new byte[16];
		System.arraycopy(buf, pos, UID, 0, 16);
		pos+=16;
		short tag= ByteTools.readShort(buf[pos],buf[pos+1]); pos+=2; 
		int size= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4; 
		int cref= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4; 
		int delayOffset= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4;
		byte usage= buf[pos++];
		byte nameLen= buf[pos++];
		byte unused1= buf[pos++];
		byte unused2= buf[pos++];
		if (nameLen==0 && delayOffset==0 && buf.length > 36) {
			// BLIP record follows then
			byte[] BLIPbuf = new byte[24];
			System.arraycopy(buf, pos, BLIPbuf, 0, 24);
			int version = (0xF&BLIPbuf[0]);
		    int inst = ((0xFF&BLIPbuf[1])<<4)|(0xF0&BLIPbuf[0])>>4;
		    int fbt = ((0xFF&BLIPbuf[3])<<8)|(0xFF&BLIPbuf[2]);
		    int len = ByteTools.readInt(BLIPbuf[4], BLIPbuf[5], BLIPbuf[6], BLIPbuf[7]);
		    UID2= new byte[16];
		    System.arraycopy(buf, 8, UID2, 0, 16);    
		} else if (buf.length > 36){
			Logger.logWarn("Delay? " + ((delayOffset==0)?"No":"Yes") + " Namelen=" + nameLen);
		}
		*/
		int ref = ByteTools.readInt( buf[24], buf[25], buf[26], buf[27] );
		cRef.add( ref );

		int HEADERLEN = 61;
		int STARTPOS = HEADERLEN;
		int BYTELEN = buf.length;
		if( HEADERLEN > BYTELEN )
		{
			BYTELEN = 0;
			STARTPOS = 0;
		}
		BYTELEN -= STARTPOS;

		byte[] imgBuf = new byte[BYTELEN];
		System.arraycopy( buf, STARTPOS, imgBuf, 0, BYTELEN );

		imageData.add( imgBuf );
		return imgBuf;
	}

	/**
	 * create a new MSODrawingGroup record based upon image datas defined in imageData/imageType/cRef arrays +
	 * spidMax, numDrawings, numShapes, numIdClusters
	 * <p/>
	 * The squenence of records here are:
	 * F000, F006, F001, F007(xNumImages), F00B ,F11E
	 * MSOFBTDGG MSOFBTBSTORECONTAINER MSOFBTBSE (x numimages)
	 */
	public void updateRecord()
	{
		MsofbtBSE[] BSE = new MsofbtBSE[imageData.size()];    //(0xF007,1,0);
		byte[] imageBytes = null;

		ByteArrayOutputStream bos = null;
		try
		{
			bos = new ByteArrayOutputStream();
			for( int i = 0; i < imageData.size(); i++ )
			{
				BSE[i] = new MsofbtBSE( MSODrawingConstants.MSOFBTBSE, Integer.parseInt( imageType.get( i ).toString() ), 2 );
				BSE[i].setImageData( (byte[]) imageData.get( i ) );
				BSE[i].setImageType( Integer.parseInt( imageType.get( i ).toString() ) );
				BSE[i].setRefCount( (Integer) cRef.get( i ) );        // 20071120 KSC: set the reference count for this image data
				bos.write( BSE[i].toByteArray() );
			}
			imageBytes = bos.toByteArray();
		}
		catch( Exception e )
		{
			log.error( "Msodrawingroup createData error.", e );
		}

		MsofbtDgg dgg = new MsofbtDgg( MSODrawingConstants.MSOFBTDGG, 0, 0 );
		dgg.setSpidMax( spidMax );        // 20071113 KSC
		numDrawings = getNumDrawings(); // 20100324 KSC: changed from: msoRecs.size();	// 20080908 KSC
		dgg.setNumDrawings( numDrawings );
		numShapes = getNumShapes();    // 20080904 KSC: sum up dg's shapes
		dgg.setNumShapes( numShapes );
		// 2008003 KSC: numIdClusters is solely a function of spidMax dgg.setNumIdClusters(numIdClusters);
		byte[] dggBytes = dgg.toByteArray();

		MsofbtOPT OPT = new MsofbtOPT( MSODrawingConstants.MSOFBTOPT, 0, 3 );
		// add the apparent basic shape options
		OPT.setProperty( MSODrawingConstants.msooptfFitTextToShape, false, false, 0x80008, null );
		OPT.setProperty( MSODrawingConstants.msooptfillColor, false, false, 0x8000041, null );
		OPT.setProperty( MSODrawingConstants.msooptlineColor, false, false, 0x8000040, null );
		byte[] OPTBytes = OPT.toByteArray();

    	/* 20070915 KSC not necessary for all msodgs*/
		MsofbtSplitMenuColors SplitMenuColors = new MsofbtSplitMenuColors( MSODrawingConstants.MSOFBTSPLITMENUCOLORS, 4, 0 );
		byte[] SplitMenuColorsBytes = SplitMenuColors.toByteArray();

		int totalLength = imageBytes.length;

		// 20080910 KSC: if no images, don't input n MSOFBTBSTORE
		byte[] BstoreContainerBytes = new byte[0];
		if( totalLength > 0 )
		{
			MsofbtBstoreContainer BstoreContainer = new MsofbtBstoreContainer( MSODrawingConstants.MSOFBTBSTORECONTAINER,
			                                                                   imageData.size(),
			                                                                   15 );
			BstoreContainer.setLength( totalLength );
			BstoreContainerBytes = BstoreContainer.toByteArray();

			// add up the stuff
			totalLength += OPTBytes.length +
					SplitMenuColorsBytes.length +
					BstoreContainerBytes.length +
					dggBytes.length;
		}
		else
		{
			// add up the stuff
			totalLength += OPTBytes.length +
					SplitMenuColorsBytes.length +
					dggBytes.length;
		}
		MsofbtDggContainer dggContainer = new MsofbtDggContainer( MSODrawingConstants.MSOFBTDGGCONTAINER, 0, 15 );
		dggContainer.setLength( totalLength );

		byte[] dggContainerBytes = dggContainer.toByteArray();

		int pos = 0;
		byte[] retData = new byte[totalLength + dggContainerBytes.length];

		System.arraycopy( dggContainerBytes, 0, retData, pos, dggContainerBytes.length );
		pos += dggContainerBytes.length;
		System.arraycopy( dggBytes, 0, retData, pos, dggBytes.length );
		pos += dggBytes.length;
		System.arraycopy( BstoreContainerBytes, 0, retData, pos, BstoreContainerBytes.length );
		pos += BstoreContainerBytes.length;
		// this is the BSE array
		System.arraycopy( imageBytes, 0, retData, pos, imageBytes.length );
		pos += imageBytes.length;
		// default OPT
		System.arraycopy( OPTBytes, 0, retData, pos, OPTBytes.length );
		pos += OPTBytes.length;
		// 20070915 KSC not truly necessary
		System.arraycopy( SplitMenuColorsBytes, 0, retData, pos, SplitMenuColorsBytes.length );
		pos += SplitMenuColorsBytes.length;

		setData( retData );
	}

	/**
	 * sets the underlying image bytes
	 *
	 * @param bts  new image bytes
	 * @param bs   Boundsheet
	 * @param rec  original Msodrawing rec linked to image
	 * @param name original image name (used for lookups)
	 * @return
	 */
	public boolean setImageBytes( byte[] bts, Boundsheet bs, MSODrawing rec, String name )
	{
		// Find original image handle - often is different than getImageIndex due to reuse, etc. of image bytes
		int trueIdx = rec.getImageIndex() - 1;    // true index into imageData and cRef arrays
		if( trueIdx < 0 )
		{
			return false;
		}

		if( (imageData.size()) <= trueIdx )
		{
			return false;
		}

		try
		{
			if( getCRef( trueIdx ) > 1 )
			{    //20080802 KSC: if referenced more than 1x, add new so don't overwrite original
				// create new image handle with new bytes
				ImageHandle im = new ImageHandle( bts, bs );
				// Find original image handle + fill new with original info
				int index = -1;
				ImageHandle[] imgz = bs.getImages();
				for( int i = 0; i < imgz.length; i++ )
				{
					if( imgz[i].getName().equals( name ) )
					{
						index = i;
						break;
					}
				}
				ImageHandle origIm = imgz[index];    // get original image handle
				im.setName( origIm.getName() );        // set new with original's data
				im.setShapeName( origIm.getShapeName() );
				im.setImageType( origIm.getImageType() );
				// insert new image into sheet
				bs.insertImage( im, true );
				im.position( origIm );    // position to original
				// now remove original mso rec
				removeMsodrawingrec( rec, bs, true );
				index = imageData.size() - 1;    // new index
			}
			else    // just set the image bytes
			{
				imageData.set( trueIdx, bts );
			}
		}
		catch( Exception ex )
		{
			log.error( "Msodrawingroup setImageBytes failed.", ex );
			return false;
		}

		updateRecord();
		parse();
		wkbook.initImages();
		return true;
	}

	/**
	 * returns the underlying image bytes
	 *
	 * @param index
	 * @return
	 */
	public byte[] getImageBytes( int index )
	{
		if( index < 0 )
		{
			return null;
		}

		if( index >= imageData.size() )
		{
			return null;
		}

		byte[] ret = null;
		try
		{
			ret = (byte[]) imageData.get( index );
		}
		catch( Exception ex )
		{
			log.error( "Msodrawingroup getImageBytes error.", ex );
		}
		return ret;

	}

	public int getImageType( int index )
	{

		return Integer.parseInt( imageType.get( index ).toString() );
	}

	/**
	 * returns the number of *unique* images in this workbook
	 *
	 * @return
	 */
	public int getNumImages()
	{
		return imageData.size();
	}

	/**
	 * related to number of drawing objects (= images + charts) but unclear how the count goes; may include deleted, etc .
	 *
	 * @return
	 */
	public int getNumDrawings()
	{
		// 20100324 KSC: this is experimental as numDrawings do not follow any obvious logic ...
		numDrawings = 0;
		// 20100511 KSC: this is not correct ...
		for( int i = 0; i < msoRecs.size(); i++ )
		{
			// 20100518 KSC: try this:
			if( ((MSODrawing) msoRecs.get( i )).isHeader() )
			{
				numDrawings++;
			}
			//numDrawings= Math.max(numDrawings, ((Msodrawing)msoRecs.get(i)).getDrawingId());
		}
//		if (numDrawings==0 && msoRecs.size() > 0) numDrawings++;
		return numDrawings;
	}

	/**
	 * count the number of shapes in the document; shape mso's contain a msofbtSpContainer sub-record  (TODO: is this true in every case?)
	 *
	 * @return
	 */
	public int getNumShapes()
	{
		/**
		 * NOTE: I've never found a clear algorithm for numShapes which results in
		 * matching Excel results; however,
		 * it appears that numShapes can be >= to Excel's value and open correctly;
		 * problems occur when numShapes are less than what Excel expects.
		 * So the below is basically the maximum numShapes available and appears to 
		 * result in templates
		 */
		/*
		numShapes= 1;	
		for (int i= 0; i < msoRecs.size(); i++) {
			MSODrawing mso= ((MSODrawing) msoRecs.get(i));
			if (mso.isShape)
				numShapes++;
		}
		*/
		numShapes = msoRecs.size();
		return numShapes;
	}

	/**
	 * return the # of Id Clusters (charts related?)
	 *
	 * @return
	 */
	public int getNumIdClusters()
	{
		return numIdClusters;
	}

	/**
	 * return SpidMax
	 *
	 * @return
	 */
	public int getSpidMax()
	{
		return spidMax;
	}

	/**
	 * set SpidMax
	 *
	 * @param spid
	 */
	public void setSpidMax( int spid )
	{
		spidMax = spid;
	}

	/**
	 * test to see if imageData is in imageArray
	 *
	 * @param imgData byte[] defining image
	 * @return
	 */
	protected int containsImage( byte[] imgData )
	{
		int z = -1;
		for( int i = 0; (i < imageData.size()) && (z < 0); i++ )
		{
			if( java.util.Arrays.equals( imgData, ((byte[]) imageData.get( 0 )) ) )
			{
				z = i;
			}
		}
		return z;
	}

	/**
	 * return the index into the imageData array for the specified image (via byte lookup)
	 *
	 * @param imgData image bytes
	 * @return index into imageData array
	 */
	public int findImage( byte[] imgData )
	{
		return containsImage( imgData );
	}

	/**
	 * if imageData doesn't exist, add to array
	 * otherwise just inc ref count
	 *
	 * @param imgData             byte[] defining image
	 * @param imgType             type of image
	 * @param bAddUnconditionally add new even if already referenced (used in setting image bytes)
	 * @return index to image
	 */
	public int addImage( byte[] imgData, int imgType, boolean bAddUnconditionally )
	{
		int n = -1;
		// 20080908 KSC: done automatically numShapes++;	// 20080208 KSC: if add unconditionally, add even if imageData already exists 
		if( bAddUnconditionally || ((n = containsImage( imgData )) == -1) )
		{ // 20071120 KSC: it's a unique image
			imageData.add( imgData );
			imageType.add( imgType );
			cRef.add( 1 );
			n = imageData.size();
		}
		else
		{    // 20071120 KSC: If not a unqiue image, just update ref count
			incCRef( n );
			n++;
		}
		return n;
	}

	public void clear()
	{
		numShapes = 0;
		imageData.clear();
		imageType.clear();
	}

	/**
	 * If large, MSODrawingGroup will span multiple records; merge data
	 *
	 * @param rec next MSODrawingGroup record in stream
	 */
	public void mergeRecords( MSODrawingGroup rec )
	{
		// Merge and remove continues 
		if( rec.hasContinues() )
		{
			rec.mergeAndRemoveContinues();        // now that data is merged, get rid of continues ...
		}
		byte[] prevData = getBytes();
		byte[] newData = rec.getBytes();
		byte[] totalData = new byte[newData.length];
		if( prevData != null )
		{ // a simple append of the data together
			totalData = new byte[prevData.length + newData.length];
			System.arraycopy( prevData, 0, totalData, 0, prevData.length );
			System.arraycopy( newData, 0, totalData, prevData.length, newData.length );
		}
		else
		{
			totalData = newData;
		}
		setData( totalData );
	}

	/**
	 * show pertinent information for record
	 */
	public String toString()
	{
		StringBuffer sb = new StringBuffer();
		sb.append( "MSODrawingGroup: numShapes=" + numShapes + " numDrawings=" + numDrawings + " numIdCluster=" + numIdClusters + " spidMax=" + spidMax );
		sb.append( "\nNumber of drawing records=" + msoRecs.size() );
		if( data != null )
		{
			sb.append( " Length of data=" + data.length );
		}
		return sb.toString();
	}

	// continue handling
	public void mergeAndRemoveContinues()
	{
		if( !isContinueMerged && hasContinues() )
		{
			super.mergeContinues();
		}
		if( hasContinues() )
		{
			// now that data is merged, get rid of continues ...
			Iterator it = continues.iterator();
			while( it.hasNext() )
			{
				Continue ci = (Continue) it.next();
				getStreamer().removeRecord( ci ); // remove existing continues from stream
			}
			super.removeContinues();
			continues = null;
		}
	}

	/**
	 * increment reference count for specific image data
	 *
	 * @param idx
	 */
	protected void incCRef( int idx )
	{
		if( (idx >= 0) && (idx < cRef.size()) )
		{
			int cr = (Integer) cRef.get( idx ) + 1;
			cRef.remove( idx );
			cRef.add( idx, cr );
		} //else  20071126 KSC: it's OK, can have - indexes ... 
		//log.error("Index error encountered when updating Reference Count");
	}

	/**
	 * return the reference count for the specific image
	 *
	 * @param idx
	 */
	protected int getCRef( int idx )
	{
		if( (idx >= 0) && (idx < cRef.size()) )
		{
			return (Integer) cRef.get( idx );
		}
		//log.error("MSODrawingGroup: error encountered when returning Reference Count");
		return -1;
	}

	/**
	 * decrement the reference count for the specific image
	 *
	 * @param idx
	 */
	protected void decCRef( int idx )
	{
		if( (idx >= 0) && (idx < cRef.size()) )
		{
			int cr = (Integer) cRef.get( idx ) - 1;
			cRef.remove( idx );
			cRef.add( idx, cr );
		}
		else
		{
			log.error( "MSODrawingGroup: error encountered when decrementing Reference Count" );
		}
	}

	/**
	 * add a new Drawing Record based on existing drawing record
	 * i.e. from CopyWorkSheet ...
	 *
	 * @param spidMax
	 * @param rec
	 */
	public void addDrawingRecord( int spidMax, MSODrawing rec )
	{
		numDrawings++;
		incCRef( rec.imageIndex - 1 );  // increment cRef
		this.spidMax = spidMax;
		updateRecord();        // given all information, generate appropriate bytes
	}

	/**
	 * Must ensure that oridinal drawing Id for each drawing record is correct
	 * Plus ensure SPID's are correct
	 * <p/>
	 * Not default prestreaming as we need these values when we assemble sheet recs
	 */
	public void prestream()
	{
		int j = 0;
		if( dirtyflag )
		{
			for( int i = 0; i < msoRecs.size(); i++ )
			{
				MSODrawing mso = (MSODrawing) msoRecs.get( i );
				if( mso.isHeader() )
				{
					mso.setDrawingId( ++j );
				}
			}
		}
	}

	/**
	 * return the number of drawing recs
	 */
	public int getNumDrawingRecs()
	{
		return msoRecs.size();
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		super.close();
		for( int i = 0; i < msoRecs.size(); i++ )
		{
			MSODrawing m = (MSODrawing) msoRecs.get( i );
			m.close();
		}
		msoRecs.clear();
	}
}