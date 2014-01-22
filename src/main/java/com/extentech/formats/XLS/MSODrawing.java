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

import com.extentech.ExtenXLS.ColHandle;
import com.extentech.ExtenXLS.RowHandle;
import com.extentech.formats.escher.MsofbtClientAnchor;
import com.extentech.formats.escher.MsofbtClientData;
import com.extentech.formats.escher.MsofbtDg;
import com.extentech.formats.escher.MsofbtDgContainer;
import com.extentech.formats.escher.MsofbtOPT;
import com.extentech.formats.escher.MsofbtSp;
import com.extentech.formats.escher.MsofbtSpContainer;
import com.extentech.formats.escher.MsofbtSpgr;
import com.extentech.formats.escher.MsofbtSpgrContainer;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

/**
 * <b>Msodrawing: MS Office Drawing (ECh)</b><br>
 * <p/>
 * These records contain only data.<p><pre>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       rgMSODr     var    MSO Drawing Data
 * <p/>
 * </p></pre>
 * <p/>
 * The Msodrawing record represents the MSOFBTDGCONTAINER (0xF002) in the Drawing Layer (Escher) format, and contains all per-sheet
 * types of info, including the shapes themselves.  (A shape=the elemental object that composes a drawing.  All graphical figures on a
 * drawing are shapes).  With few exceptions, shapes are stored hierachically according to how they've been grouped thru the use of
 * the Draw/Group command).
 * Each Msodrawing record contains several sub-records or atoms (atoms are records that are kept inside container records; container
 * records keep atoms and other container records organized).
 * There are several such records that are important to us (these are always present - except for MSOFBTDG in subsequent recs):
 * MSOFBTDG-				Basic Drawing Info-	#shapes in this drawing; last SPID given to an SP in this Drawing Group
 * MSOFBTSPGRCONTAINER-		Patriarch shape	Container - always first MSOFBTSPGRCONTAINER in the drawing container
 * MSOFBTSPCONTAINER-		Shape Container
 * MSOFBTSP-				Shape Atom Record- SPID= shape id + a set of flags
 * MSOFBTOPT-				Property Table Record Associated with Shape Rec- holds image Name, index + many other properties
 * MSOFBTCLIENTANCHOR- 	Client Anchor rec- holds size/bounds info
 * MSOFBTCLIENTDATA-		Host-specific client data record
 * <p/>
 * There are many other records or atoms that are optional and that we will omit for now.
 * <p/>
 * There appears to be 1 msodrawing record per image (there is also one OBJ record per Msodrawing record).
 * The first or header msodrawing record contains MSODBTDG for # shapes
 * There is one MsodrawingGroup record per file.  This MsodrawingGroup record also contains sub-records that hold # shapes
 * <p/>
 * Occasionally, when there is a lot of image data, there can be 2 MSODRAWINGGROUP objects -- continue recs will follow
 * the second MSODG only.
 * <p/>
 * SPIDs are unique per drawing group, and are parceled out by the drawing group to individual drawings in blocks of 1024
 *
 * @see MSODrawingGroup
 */
// TODO: MSOFBTCLIENTANCHOR may be substituted for MSOFBTANCHOR (clipboard), MSOFBTCHILDANCHOR (if shape is a child of a group shape)
public final class MSODrawing extends com.extentech.formats.XLS.XLSRecord
{
	private static final long serialVersionUID = 8275831369787287975L;

	public byte[] PROTOTYPE_BYTES = {
			15,
			0,
			4,
			-16,
			92,
			0,
			0,
			0,
			-78,
			4,
			10,
			-16,
			8,
			0,
			0,
			0,
			2,
			4,
			0,
			0,
			0,
			10,
			0,
			0,
			35,
			0,
			11,
			-16,
			34,
			0,
			0,
			0,
			4,
			65,
			2,
			0,
			0,
			0,
			5,
			-63,
			22,
			0,
			0,
			0,
			66,
			0,
			108,
			0,
			117,
			0,
			101,
			0,
			32,
			0,
			104,
			0,
			105,
			0,
			108,
			0,
			108,
			0,
			115,
			0,
			0,
			0,
			0,
			0,
			16,
			-16,
			18,
			0,
			0,
			0,
			2,
			0,
			2,
			0,
			0,
			0,
			11,
			0,
			0,
			0,
			6,
			0,
			0,
			0,
			22,
			0,
			75,
			0,
			0,
			0,
			17,
			-16,
			0,
			0,
			0,
			0
	};

	int imageIndex = -1;
	boolean bActive = false;        // true if this image is Active (not deleted) - NOTE: setting Active algorithm is not definitively proven
	String imageName = "", shapeName = "";
	short clientAnchorFlag;            // MSOFBTCLIENTANCHOR 1st two bytes - * 0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells. */
	short[] bounds = new short[8];    // MSOFBTCLIENTANCHOR-
	short origHeight, origWidth;    // save original height and width so that if underlying row height(s) or column width(s) change, can still set dimensions correctly ...
	private MsofbtOPT optrec = null;    // 20091209 KSC: save MsofbtOPT records for later updating ...
	private MsofbtOPT secondaryoptrec = null, tertiaryoptrec = null;    // apparently can have secondary and tertiary obt recs depending upon version in which it was saved ...
	short shapeType = 0;                // shape type for this drawing record

	private int SPIDSEED = 1024;
	boolean bIsHeader = false;            // whether "this is the 1st Msodrawing rec in the sheet" - contains several header records
	int SPID = 0;                        // Shape ID 
	int SPCONTAINERLENGTH = 0;            // this shape's container length 
	boolean isShape = true;        // false for Mso's which do not contain an SPCONTAINER sub-record (can be attached textbox or solver container); not included in number of shapes count
	// only applicable to header records ***
	int numShapes = 1;            // TODO: how do we know the value for a new image????
	int lastSPID = SPIDSEED;    // lastSPID is stored at book level so that we can track max SPIDs - useful when images have been deleted and SPIDs are not in order ...
	static final int HEADERRECLENGTH = 24;            //
	int otherSPCONTAINERLENGTH = 0;    // sum of other SPCONTAINERLENGTHs from other Msodrawing recs
	int SOLVERCONTAINERLENGTH = 0;    // Solver Container length 
	int drawingId = 0;                // ordinal # of this drawing record in the workbook

	/**
	 * create a new msodrawing record with the desired SPID, imageName, shapeName and imageIndex
	 * bounds should also be set?
	 * create correct record bytes
	 *
	 * @param spid
	 * @param imageName
	 * @param shapeName
	 * @param imageIndex
	 * @return
	 */
	public byte[] createRecord( int spid, String imageName, String shapeName, int imageIndex )
	{
		this.imageName = imageName;
		this.shapeName = shapeName;
		this.imageIndex = imageIndex;

		byte[] retData;
		// Order of Msodrawing required records:
		/*		MSOFBTDG
         * 		MSOFBTSPGRCONTAINER
    	 * 			MSOFBTSPCONTAINER
         * 			MSOFBTSP
         * 			MSOFBTOPT
         * 			MSOFBTCLIENTANCHOR (MSOFBTCHILDANCHOR, MSOFBTANCHOR)
         * 			MSOFBTCLIENTDATA
         * 
         * 		NOTE that every container has a length field that = the sum of the length of all the atoms (records) it contains as well as the length
         *      of its header 
         */
		// key sub-records to update:
		// MSOFBTSP, MSOFBTOPT (image index, shape name, image name), CLIENTANCHOR if present - bounds
		// plus container which must be calculated from it's sub-records or atoms
		this.SPID = spid;
		//this.SPID = SPIDSEED + imageIndex;// algorithm is incorrect for instances when images are deleted or out of order ...
		//    	Following are present in all msodrawing: MSOFBTSPCONTAINER MSOFBTSP MSOFBTOPT MSOFBTCLIENTANCHOR MSOFBTCLIENTDATA  -- not true! can have CHILDANCHOR, ANCHOR ...
		//Shape Atom; shape type must be msosptPictureFrame = 75
		MsofbtSp msofbtSp1 = new MsofbtSp( MSODrawingConstants.MSOFBTSP, shapeType, 2 );
		msofbtSp1.setId( SPID );
		msofbtSp1.setGrfPersistence( 2560 );               //flag= hasSp type + has anchor -- usual for shape-type msoFbtSp's
		byte[] msofbtSp1Bytes = msofbtSp1.toByteArray();

		// OPT= picture options
		optrec = new MsofbtOPT( MSODrawingConstants.MSOFBTOPT, 0, 3 );    //version is always 3, inst is current count of properties.
		if( imageIndex != -1 )
		{
			optrec.setImageIndex( imageIndex );
		}
		if( (imageName != null) && !imageName.equals( "" ) )
		{
			optrec.setImageName( imageName );
		}
		if( (shapeName != null) && !shapeName.equals( "" ) )
		{
			optrec.setShapeName( shapeName );
		}

		byte[] msofbtOPT1Bytes = optrec.toByteArray();

		// Client Anchor==Bounds
		MsofbtClientAnchor msofbtClientAnchor1 = new MsofbtClientAnchor( MSODrawingConstants.MSOFBTCLIENTANCHOR, 0, 0 );
		msofbtClientAnchor1.setBounds( bounds );
		byte[] msofbtClientAnchor1Bytes = msofbtClientAnchor1.toByteArray();

		MsofbtClientData msofbtClientData1 = new MsofbtClientData( MSODrawingConstants.MSOFBTCLIENTDATA, 0, 0 );  //This is an empty record
		byte[] msofbtClientData1Bytes = msofbtClientData1.toByteArray();

		SPCONTAINERLENGTH = msofbtSp1Bytes.length + msofbtOPT1Bytes.length + msofbtClientAnchor1Bytes.length + msofbtClientData1Bytes.length;

		// 20100412 KSC: must count "oddball" msofbtClientTextBox record that follows this Mso's obj record ... 
		if( shapeType == MSODrawingConstants.msosptTextBox )
		{
			SPCONTAINERLENGTH += 8;
		}
		MsofbtSpContainer msofbtSpContainer1 = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
		msofbtSpContainer1.setLength( SPCONTAINERLENGTH );
		byte[] msofbtSpContainer1Bytes = msofbtSpContainer1.toByteArray();
		SPCONTAINERLENGTH += +msofbtSpContainer1Bytes.length;    // include this rec
		retData = new byte[SPCONTAINERLENGTH];

		int pos = 0;
		System.arraycopy( msofbtSpContainer1Bytes, 0, retData, pos, msofbtSpContainer1Bytes.length );
		pos += msofbtSpContainer1Bytes.length;
		System.arraycopy( msofbtSp1Bytes, 0, retData, pos, msofbtSp1Bytes.length );
		pos += msofbtSp1Bytes.length;
		System.arraycopy( msofbtOPT1Bytes, 0, retData, pos, msofbtOPT1Bytes.length );
		pos += msofbtOPT1Bytes.length;
		System.arraycopy( msofbtClientAnchor1Bytes, 0, retData, pos, msofbtClientAnchor1Bytes.length );
		pos += msofbtClientAnchor1Bytes.length;
		System.arraycopy( msofbtClientData1Bytes, 0, retData, pos, msofbtClientData1Bytes.length );
		pos += msofbtClientData1Bytes.length;    // 20100420 KSC: empty client data record- necessary???

		if( bIsHeader )
		{    //This is only present in the first msodrawing per sheet    		    		
			if( lastSPID < SPID )
			{
				lastSPID = SPID;        // TODO: Shouldn't assume to be SPID
			}
			int totalSPRECORDS = 0;

			// Header also contains Shape Id Seed SP record
			MsofbtSp msofbtSp = new MsofbtSp( MSODrawingConstants.MSOFBTSP,
			                                  MSODrawingConstants.msosptMin,
			                                  2 );  //1st shape rec is of type MSOSPTMIN 
			msofbtSp.setId( SPIDSEED );                      // SPMIN==SPIDSEED          
			msofbtSp.setGrfPersistence( 5 );                    // flag= fPatriarch,                      		
			byte[] msofbtSpBytes = msofbtSp.toByteArray();

			MsofbtSpgr msofbtSpgr = new MsofbtSpgr( MSODrawingConstants.MSOFBTSPGR, 0, 1 );
			msofbtSpgr.setRect( 0, 0, 0, 0 );
			byte[] msofbtSpgrBytes = msofbtSpgr.toByteArray();
			totalSPRECORDS = msofbtSpgrBytes.length + msofbtSpBytes.length;

			MsofbtSpContainer msofbtSpContainer = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
			msofbtSpContainer.setLength( totalSPRECORDS );
			byte[] msofbtSpContainerBytes = msofbtSpContainer.toByteArray();
			totalSPRECORDS += msofbtSpContainerBytes.length;

			SPCONTAINERLENGTH += totalSPRECORDS;
			MsofbtSpgrContainer msofbtSpgrContainer = new MsofbtSpgrContainer( MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15 );
			msofbtSpgrContainer.setLength( SPCONTAINERLENGTH + otherSPCONTAINERLENGTH );
			byte[] msofbtSpgrContainerBytes = msofbtSpgrContainer.toByteArray();

			MsofbtDg msofbtDg = new MsofbtDg( MSODrawingConstants.MSOFBTDG, drawingId, 0 );
			msofbtDg.setNumShapes( numShapes );                //Number of images and drawings.	
			msofbtDg.setLastSPID( lastSPID );                    //lastSPID      			 
			byte[] msofbtDgBytes = msofbtDg.toByteArray();

			MsofbtDgContainer msofbtDgContainer = new MsofbtDgContainer( MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15 );
			msofbtDgContainer.setLength( HEADERRECLENGTH + SPCONTAINERLENGTH + otherSPCONTAINERLENGTH );
			byte[] msofbtDgContainerBytes = msofbtDgContainer.toByteArray();

			byte[] headerRec = new byte[80 + retData.length]; //+retData.length];

			pos = 0;
			System.arraycopy( msofbtDgContainerBytes, 0, headerRec, pos, msofbtDgContainerBytes.length );
			pos += msofbtDgContainerBytes.length;
			System.arraycopy( msofbtDgBytes, 0, headerRec, pos, msofbtDgBytes.length );
			pos += msofbtDgBytes.length;
			System.arraycopy( msofbtSpgrContainerBytes, 0, headerRec, pos, msofbtSpgrContainerBytes.length );
			pos += msofbtSpgrContainerBytes.length;
			System.arraycopy( msofbtSpContainerBytes, 0, headerRec, pos, msofbtSpContainerBytes.length );
			pos += msofbtSpContainerBytes.length;
			System.arraycopy( msofbtSpgrBytes, 0, headerRec, pos, msofbtSpgrBytes.length );
			pos += msofbtSpgrBytes.length;
			System.arraycopy( msofbtSpBytes, 0, headerRec, pos, msofbtSpBytes.length );
			pos += msofbtSpBytes.length;

			System.arraycopy( retData, 0, headerRec, pos, retData.length );
			retData = headerRec;

		}

		this.setData( retData );
		this.setLength( data.length );
		return retData;
	}

	/**
	 * parse the data contained in this drawing record
	 */
	@Override
	public void init()
	{
		// *************************************************************************************************************************************
		// 20070910 KSC: parse MSO record + atom ids (records keep atoms and other containers; atoms contain info and are kept inside containers)
		// common record header for both:  ver, inst, fbt, len; fbt deterimes record type (0xF000 to 0xFFFF)
		// for a specific record, inst differentiates atoms
		// for atoms, ver= version; for records, ver= 0xFFFF
		// for atoms, len= of atom excluding header; for records; sum of len of atoms contained within it
		
		/*	an Msodrawing record may contain the following records and atoms:
		 * 	MSOFBTDG					// drawing record: count + MSOSPID seed
        	MSOFBTREGROUPITEMS			// Mappings to reconstitute groups (for regrouping)
        	MSOFBTCOLORSCHEME
        	MSOFBTSPGRCONTAINER			// Group Shape Container 							
            MSOFBTSPCONTAINER			// Shape Container
        		MSOFBTSPGR				// Group-shape-specific info  (i.e. shapes that are groups) optional
        	**	MSOFBTSP				// A Shape atom record **
		 	**	MSOFBTOPT				// The Property Table for a shape ** - image index + name ... 
            	MSOFBTTEXTBOX			// if the shape has text
            	MSOFBTCLIENTTEXTBOX		// for clipboard stream
            	MSOFBTANCHOR			// Anchor or location fo a shape (if streamed to a clipboard) optional
            	MSOFBTCHILDANCHOR		//   " ", if shape is a child of a group shape optional
            **  MSOFBTCLIENTANCHOR		// Client Anchor/Bounds ** 
            	MSOFBTCLIENTDATA		// content is determined by host optional 	
            	MSOFBTOLEOBJECT			// optional
            	MSOFBTDELETEDPSPL		// optional
		 */
		// *************************************************************************************************************************************
		super.init();

		ByteArrayInputStream bis = new ByteArrayInputStream( super.getData() );
		int version, inst, fbt, len;
		SPCONTAINERLENGTH = 0;        // this shape container length		
		otherSPCONTAINERLENGTH = 0;    // if header, all other SPCONTAINERLENGTHS -- calc from DGCONTAINERLENGTH + SPCONTAINERLENGTHS

		int SPGRCONTAINERLENGTH = 0;        // group shape container length	
		int SPCONTAINERATOMS = 0;        // atoms or sub-records which make up the shape container 
		int DGCONTAINERLENGTH = 0;        // drawing container length
		int DGCONTAINERATOMS = 0;        // atoms or sub-records which, along with SPGRCONTAINER + SOLVERCONTAINER(s), make up the DGCONTAINERLENGHT
		int SOLVERCONTAINERATOMS = 0;    // atoms or sub-records which make up the SOLVERCONTAINERLENGTH
		boolean hasUndoInfo = false;        // true if a non-header Mso contains an SPGRCONTAINER - documentation states:   Shapes that have been deleted but that could be brought back via Undo.
		for(; bis.available() > 0; )
		{
			byte[] dat = new byte[8];
			bis.read( dat, 0, 8 );
			version = (0xF & dat[0]);    // 15 for containers, version for atoms
			inst = ((0xFF & dat[1]) << 4) | (0xF0 & dat[0]) >> 4;
			fbt = ((0xFF & dat[3]) << 8) | (0xFF & dat[2]);    // record type id
			len = ByteTools.readInt( dat[4],
			                         dat[5],
			                         dat[6],
			                         dat[7] );    // for atoms, record length - header length (=8), if container, refers to sum of lengths of atoms inside it, incl. record headers 
			if( version == 15 )
			{// do not parse containers 
				// MSOFBTSPGRCONTAINER:		// Shape Group Container, contains a variable number of shapes (=msofbtSpContainer) + other groups 0xF003 
				// MSOFBTSPCONTAINER:		// Shape Container 0xF004
				// may have several SPCONTAINERs, 1 for background shape, several for deleted shapes ...
				// possible containers
				// DGCONTAINER= DG, REGROUPITEMS, ColorSCHEME, SPGR, SPCONTAINER		    			    	
				// SPGRCONTAINER= SPCONTAINER(s)		    	
				// SPCONTAINER= SPGR, SP, OPT, TEXTBOX, ANCHOR, CHILDANCHOR, CLIENTANCHOR, CLIENTDATA, OLEOBJECT, DeletedPSPL
				// SOLVERCONTAINER= ConnetorRule, AlignRule, ArcRule, ClientRule, CalloutRule, 		    	
				if( fbt == MSODrawingConstants.MSOFBTDGCONTAINER )
				{
					bIsHeader = true;
					otherSPCONTAINERLENGTH = len;    //-HEADERRECLENGTH;
					DGCONTAINERLENGTH = len;
				}
				else if( fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER )
				{//patriarch shape, with all non-background non-deleted shapes in it - may have more than 1 subrecord
					// A group is a collection of other shapes. The contained shapes are placed in the coordinate system of the group. 
					// The group container contains a variable number of shapes (msofbtSpContainer) and other groups (msofbtSpgrContainer, for nested groups). 
					// The group itself is a shape, and always appears as the first msofbtSpContainer in the group container.
					if( SPGRCONTAINERLENGTH == 0 )    // only add 1st container length - others are deleted shapes ... 
					{
						SPGRCONTAINERLENGTH = len;
					}
					if( !bIsHeader )// then this is a grouped shape, must add the header length as is apparent in existing container lengths (see below)
					{
						hasUndoInfo = true;
					}
				}
				else if( fbt == MSODrawingConstants.MSOFBTSPCONTAINER )
				{
					SPCONTAINERLENGTH += (len + 8); //	  add 8 for record header
					if( bIsHeader )    // keep track of total "other sp container length" - necessary to calculate SPGRCONTAINERLENGTH + DGCONTAINERLENGTH 
					{
						otherSPCONTAINERLENGTH -= (len + 8);
					}
					isShape = true;    // any mso that contains a normal "spcontainer" is a shape
				}
				else if( fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER )
				{    // solver container: rules governing shapes
					isShape = false;
					SOLVERCONTAINERLENGTH = len + 8;    // added to dgcontainerlength
				}
				else
				{
					Logger.logInfo( "MSODrawing.init: unknown container encountered: " + fbt );
				}
				continue;
			}
			// parse atoms or sub-records (distinct from container records above)
			dat = new byte[len];
			bis.read( dat, 0, len );
			switch( fbt )
			{
				case MSODrawingConstants.MSOFBTSP:            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
					boolean fGroup, fChild, fPatriarch, fDeleted, fOleShape, fHaveMaster;
					boolean fFlipH, fFlipV, fConnector, fHaveAnchor, fBackground, fHaveSpt;
					int flag;
					SPID = ByteTools.readInt( dat[0], dat[1], dat[2], dat[3] );
					flag = ByteTools.readInt( dat[4], dat[5], dat[6], dat[7] );
					// parse out flag:
					fGroup = (flag & 0x1) == 0x1;        // it's a group shape
					fChild = (flag & 0x2) == 0x2;        // it's not a top-level shape
					fPatriarch = (flag & 0x4) == 0x4;    // topmost group shape ** 1 per drawing *8
					fDeleted = (flag & 0x8) == 0x8;        // had been deleted
					// 4
					fOleShape = (flag & 0x10) == 0x10;    // it's an OLE shape
					fHaveMaster = (flag & 0x20) == 0x20;    // it has a master prop
					fFlipH = (flag & 0x40) == 0x40;        // it's flipped horizontally
					fFlipV = (flag & 0x80) == 0x80;        // it's flipped vertically
					// 8
					fConnector = (flag & 0x100) == 0x100;    // it's a connector type
					fHaveAnchor = (flag & 0x200) == 0x200;    // it's an anchor type
					fBackground = (flag & 0x400) == 0x400;    // it's a background shape
					fHaveSpt = (flag & 0x800) == 0x800;    // it has a shape-type property
					// FYI: there are normally two msofbtsp records for each drawing
					// the first (flag==5) defines fGroup + fPatriarch, 
					// with inst==0, apparently shape type to define SPIDSEED
					// the second msofbtsp record (flag=2560)defines the SPID and contains flags: 
					// fHaveAnchor, fHaveSpt and inst==shape type   
					//NOTE: setting Active algorithm is not definitively proven;
					// if it ISN'T active, why isn't the fDeleted flag set???
					bActive = bActive || fPatriarch;    // if we have fPatriarch, it's active ...
					if( fHaveSpt )
					{
						shapeType = (short) inst;    // save shape type
					}
					if( inst == 0 )    //== shape type	
					{
						SPIDSEED = SPID;        // seed+imageIndex= SPID
					}
					SPCONTAINERATOMS += len + 8;
					break;

				case MSODrawingConstants.MSOFBTCLIENTANCHOR:        // Anchor or location for a shape 
					// NOT SO! sheetIndex = ByteTools.readShort(buf[0],buf[1]);
					/**
					 bounds[0]= column # of top left position (0-based) of the shape
					 bounds[1]= x offset within the top-left column
					 bounds[2]= row # for top left corner
					 bounds[3]= y offset within the top-left corner
					 bounds[4]= column # of the bottom right corner of the shape
					 bounds[5]= x offset within the cell  for the bottom-right corner
					 bounds[6]= row # for bottom-right corner of the shape
					 bounds[7]= y offset within the cell for the bottom-right corner					
					 */
					clientAnchorFlag = ByteTools.readShort( dat[0], dat[1] );
					bounds[COL] = ByteTools.readShort( dat[2], dat[3] );
					bounds[COLOFFSET] = ByteTools.readShort( dat[4], dat[5] );
					bounds[ROW] = ByteTools.readShort( dat[6], dat[7] );
					bounds[ROWOFFSET] = ByteTools.readShort( dat[8], dat[9] );
					bounds[COL1] = ByteTools.readShort( dat[10], dat[11] );
					bounds[COLOFFSET1] = ByteTools.readShort( dat[12], dat[13] );
					bounds[ROW1] = ByteTools.readShort( dat[14], dat[15] );
					bounds[ROWOFFSET1] = ByteTools.readShort( dat[16], dat[17] );
					SPCONTAINERATOMS += len + 8;
					break;

				case MSODrawingConstants.MSOFBTOPT:    // property table atom - for 97 and earlier versions
					// 20091209 KSC: save MsoFbtOpt record for later use in updating, if necessary 
					optrec = new MsofbtOPT( fbt, inst, version );
					optrec.setData( dat );    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
					imageName = optrec.getImageName();
					shapeName = optrec.getShapeName();
					imageIndex = optrec.getImageIndex();
					SPCONTAINERATOMS += len + 8;
					break;

				// 20100519 KSC: later versions can have secondary and tertiary opt blocks
				case MSODrawingConstants.MSOFBTSECONDARYOPT:        // movie (id= 274) 
					secondaryoptrec = new MsofbtOPT( fbt, inst, version );
					secondaryoptrec.setData( dat );    // sets and parses msoFbtOpt data
					SPCONTAINERATOMS += len + 8;
					break;

				case MSODrawingConstants.MSOFBTTERTIARYOPT:    // for Office versions 2000, XP, 2003, and 2007  
					tertiaryoptrec = new MsofbtOPT( fbt, inst, version );
					tertiaryoptrec.setData( dat );    // sets and parses msoFbtOpt data
					SPCONTAINERATOMS += len + 8;
					break;

				case MSODrawingConstants.MSOFBTDG:    // Drawing Record: ID, num shapes + Last SPID for this DG
					drawingId = inst;
					numShapes = ByteTools.readInt( dat[0], dat[1], dat[2], dat[3] );  // number of shapes in this drawing
					lastSPID = ByteTools.readInt( dat[4], dat[5], dat[6], dat[7] );
				case MSODrawingConstants.MSOFBTCOLORSCHEME:
				case MSODrawingConstants.MSOFBTREGROUPITEMS:
					DGCONTAINERATOMS += len + 8;
					otherSPCONTAINERLENGTH -= (len + 8); // these atoms are part of DGCONTAINER, not SPCONTAINER - adjust accordingly					
					break;

				case MSODrawingConstants.MSOFBTCLIENTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...		    	
				case MSODrawingConstants.MSOFBTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...		    	
					isShape = false;    // is treated differently, isn't counted in numShapes calcs 
					break;

				case MSODrawingConstants.MSOFBTCHILDANCHOR:    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape					
					//  If the shape is saved to a clipboard:
				case MSODrawingConstants.MSOFBTSPGR:    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape 
					//  If the shape is a child of a group shape:
				case MSODrawingConstants.MSOFBTANCHOR:        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
				case MSODrawingConstants.MSOFBTCLIENTDATA:
				case MSODrawingConstants.MSOFBTDELETEDPSPL:
					SPCONTAINERATOMS += len + 8;
					break;

				case MSODrawingConstants.MSOFBTCONNECTORRULE:
				case MSODrawingConstants.MSOFBTALIGNRULE:
				case MSODrawingConstants.MSOFBTARCRULE:
				case MSODrawingConstants.MSOFBTCLIENTRULE:
				case MSODrawingConstants.MSOFBTCALLOUTRULE:
					SOLVERCONTAINERATOMS += len + 8;
					break;

				default:
					Logger.logInfo( "MSODrawing.init:  unknown subrecord encountered: " + fbt );
			}
		}
	  	/* //DEBUGGING:  THESE CONTAINER LENGTH CALCULATIONS PASS FOR ALL thus far MSO's ENCOUNTERED
		boolean diff= false;	
		if (isHeader()) {	  	
			int d0= (DGCONTAINERLENGTH-(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8));
			int d1= (SPGRCONTAINERLENGTH+8-(SPCONTAINERLENGTH+otherSPCONTAINERLENGTH));
			if (d0+d1!=0) {
				if (DGCONTAINERLENGTH!=(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8)) {  // this may not be 100% since must account for OTHER record's SOLVERCONTAINER LENGTHS
					System.out.println("DGCONTAINERLENGTH DIFF: " + (DGCONTAINERLENGTH-(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8)));
					diff= true;
				}			  	
				// ******* sum of SPCONTAINERS ***************
				if (SPGRCONTAINERLENGTH+8!=SPCONTAINERLENGTH+otherSPCONTAINERLENGTH) {
					System.out.println("SPGRCONTAINERLENGTH DIFF: " + (SPGRCONTAINERLENGTH+8-(SPCONTAINERLENGTH+otherSPCONTAINERLENGTH)));
					diff= true;
				}
			}
		}
		// one or two header lengths (8 bytes each) have been added to SPCONTAINERLENGTH
		// adjust here:
		int headerlens= 0;
		if (isHeader())
			headerlens= 16;
		else if (isShape) 	// non-shapes don't have SPGRCONTAINERS 
			headerlens= 8;
		if (isHeader() && SPCONTAINERLENGTH==48) // only one SPCONTAINER, decrement by 1 header length
			headerlens= 8;
		if (optrec!=null && optrec.hasTextId()) // shape has an attached text box; must add 8 for following CLIENTTEXTBOX (the !isShape Mso which follows ...)
			SPCONTAINERATOMS+=8;
		if (SPCONTAINERLENGTH-headerlens!=SPCONTAINERATOMS) {	
			System.err.println("SPCONTAINERLEN IS OFF: " + (SPCONTAINERLENGTH-headerlens-SPCONTAINERATOMS));
			System.err.println(this.toString());
			System.err.println(this.debugOutput());
			diff= true;
		}
		/**/
		if( hasUndoInfo )
		{
			SPCONTAINERLENGTH += 8;        // Shapes that have been deleted but that could be brought back via Undo. -- must add to sp container length for total container length calc (see UpdateHeader)
		}
	}

	/**
	 * update the existing mso with the appropriate basic mso data
	 * <br>NOTE: To set other mso data, see setShapeType, setIsHeader ...
	 *
	 * @param spid       Unique Shape Id
	 * @param imageName  String image name or null
	 * @param shapeName  String shape name or null
	 * @param imageIndex int index into the image byte store (for a picture-type) or -1 for null
	 * @param bounds     short[8] the position of this Mso given in rows, cols and offsets
	 */
	public void updateRecord( int spid, String imageName, String shapeName, int imageIndex, short[] bounds )
	{
		this.SPID = spid;
		this.imageName = imageName;
		this.shapeName = shapeName;
		this.imageIndex = imageIndex;
		this.bounds = bounds;
		updateRecord();
	}

	/**
	 * rebuild record bytes for this Msodrawing
	 * aside from updating significant atoms such as MSOFBTOPT, it also recalculates container lengths
	 *
	 * @param spid
	 * @return
	 */
	public void updateRecord()
	{    	
/*// debug: check algorithm:
	System.out.println(this.toString());    	
	System.out.println(this.debugOutput());
	int origSP= SPCONTAINERLENGTH;
	int origDG= 0;
	int origSPGR= 0;
/**/
		byte[] spcontainer1atoms = new byte[0];
		byte[] spcontainer2atoms = new byte[0];    // header specific
		byte[] dgcontaineratoms = new byte[0];    // header specific

		boolean hasUndoInfo = false;        // true if a non-header Mso contains an SPGRCONTAINER - documentation states:   Shapes that have been deleted but that could be brought back via Undo.
		// reset state vars
		SPCONTAINERLENGTH = 0;        // this shape container length		

		// first pass, update significant atoms and sum up container lengths
		int fbt, len;
		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] header = new byte[8];
			bis.read( header, 0, 8 );
			fbt = ((0xFF & header[3]) << 8) | (0xFF & header[2]);
			len = ByteTools.readInt( header[4], header[5], header[6], header[7] );
			if( (0xF & header[0]) == 15 )
			{    // 15 for containers, version for atoms
				// most containers, just skip; some, however, are necessaary for container length calcs (see below 
				if( fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER )
				{
					//if (origSPGR==0) origSPGR= len;
					if( !bIsHeader )// then this is a grouped shape, must add the header length as is apparent in existing container lengths (see below)
					{
						hasUndoInfo = true;
					}
				}
				else if( fbt == MSODrawingConstants.MSOFBTSPCONTAINER )
				{
					isShape = true;    // any mso that contains a normal "spcontainer" is a shape
				}
				else if( fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER )
				{    // solver container: rules governing shapes
					// TODO: is there EVER a reason to update a SOLVERCONTAINER RECORD???
					// STRUCTURE:  
					// MSOFBTSOLVERCONTAINER 61445 15/#/# (15/0/0 for an empty solvercontainer)
					// then 0 or more rules:
					// MSOFBTCONNECTORRULE 61458 1/0/24 ... etc
					SOLVERCONTAINERLENGTH = len + 8;    // added to dgcontainerlength
					isShape = false;
					// testing- remove when done
				}
				else if( fbt == MSODrawingConstants.MSOFBTDGCONTAINER )
				//origDG= len;
				{
					;
				}
			}
			else
			{
				// parse atoms or sub-records (distinct from container records above)
				byte[] data = new byte[len];
				bis.read( data, 0, len );
				switch( fbt )
				{
					case MSODrawingConstants.MSOFBTDG:    // update drawing record atom
						System.arraycopy( ByteTools.cLongToLEBytes( this.numShapes ), 0, data, 0, 4 );
						System.arraycopy( ByteTools.cLongToLEBytes( this.lastSPID ), 0, data, 4, 4 );
						len = data.length;
						data = ByteTools.append( data, header );
						dgcontaineratoms = ByteTools.append( data, dgcontaineratoms );
						break;

					case MSODrawingConstants.MSOFBTSP:            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
						// TODO: necessary to ever update the SPIDSEED?
						int flag = ByteTools.readInt( data[4], data[5], data[6], data[7] );
						boolean fHaveSpt = (flag & 0x800) == 0x800;    // it has a shape-type property
						if( flag != 5 )
						{    // if it's not the SPIDseed, update SPID
							System.arraycopy( ByteTools.cLongToLEBytes( SPID ), 0, data, 0, 4 );
						}
						if( fHaveSpt )
						{    // shape type is contained within inst var
							header[0] = (byte) ((0xF & 2) | (0xF0 & (shapeType << 4)));
							header[1] = (byte) ((0x00000FF0 & shapeType) >> 4);
						}
						data = ByteTools.append( data, header );
						if( flag != 5 )    // if it's not the header
						{
							spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						}
						else
						{
							spcontainer2atoms = ByteTools.append( data, spcontainer2atoms );
						}
						break;

					case MSODrawingConstants.MSOFBTCLIENTANCHOR:        // Anchor or location for a shape 
						// udpate bounds
						System.arraycopy( ByteTools.shortToLEBytes( (short) clientAnchorFlag ), 0, data, 0, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[0] ), 0, data, 2, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[1] ), 0, data, 4, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[2] ), 0, data, 6, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[3] ), 0, data, 8, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[4] ), 0, data, 10, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[5] ), 0, data, 12, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[6] ), 0, data, 14, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[7] ), 0, data, 16, 2 );
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTOPT:    // property table atom - for 97 and earlier versions
						// OPT= picture options
						if( optrec == null )  // if do not have a MsoFbtOPT record, then create new -- shouldn't get here as should always have an existing record
						{
							optrec = new MsofbtOPT( MSODrawingConstants.MSOFBTOPT,
							                        0,
							                        3 );    //version is always 3, inst is current count of properties.
						}
						if( imageIndex != optrec.getImageIndex() )
						{
							optrec.setImageIndex( imageIndex );
						}
						if( (imageName == null) || !imageName.equals( optrec.getImageName() ) )
						{
							optrec.setImageName( imageName );
						}
						if( (shapeName == null) || !shapeName.equals( optrec.getShapeName() ) )
						{
							optrec.setShapeName( shapeName );
						}
						data = optrec.toByteArray();
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTTERTIARYOPT:    // for Office versions 2000, XP, 2003, and 2007  
					case MSODrawingConstants.MSOFBTSECONDARYOPT:        // movie (id= 274) 
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTCOLORSCHEME:
					case MSODrawingConstants.MSOFBTREGROUPITEMS:
						data = ByteTools.append( data, header );
						dgcontaineratoms = ByteTools.append( data, dgcontaineratoms );
						break;

					case MSODrawingConstants.MSOFBTCLIENTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...		    	
					case MSODrawingConstants.MSOFBTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...
						// DON"T ADD LEN **************
						isShape = false;
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTSPGR:    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape
						data = ByteTools.append( data, header );
						spcontainer2atoms = ByteTools.append( data, spcontainer2atoms );
						break;

					case MSODrawingConstants.MSOFBTCHILDANCHOR:    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape					
					case MSODrawingConstants.MSOFBTANCHOR:        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
					case MSODrawingConstants.MSOFBTCLIENTDATA:
					case MSODrawingConstants.MSOFBTDELETEDPSPL:
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTCONNECTORRULE:
					case MSODrawingConstants.MSOFBTALIGNRULE:
					case MSODrawingConstants.MSOFBTARCRULE:
					case MSODrawingConstants.MSOFBTCLIENTRULE:
					case MSODrawingConstants.MSOFBTCALLOUTRULE:
						// SOLVERCONTAINERATOMS+=len+8;
						// TODO: is there EVER a reason to update a SOLVERCONTAINER RECORD???
						//Logger.logInfo("MSODrawing.updateRecord:  encountered solver container atom");
						break;

					default:
						Logger.logInfo( "MSODrawing.updateRecord:  unknown subrecord encountered: " + fbt );
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;
				}
			}
		}
		// container lengths:	
		// are these adjustments necessary??
/*				
	  	if (optrec!=null && optrec.hasTextId()) // shape has an attached text box; must add 8 for following CLIENTTEXTBOX (the !isShape Mso which follows ...)
			SPCONTAINERATOMS+=8;
*/
		SPCONTAINERLENGTH = spcontainer1atoms.length;
		int additionalSP = 0;
		if( hasUndoInfo )
		{
			additionalSP = 8;        // Shapes that have been deleted but that could be brought back via Undo. -- must add to sp container length for total container length calc (see UpdateHeader)
		}
		if( shapeType == MSODrawingConstants.msosptTextBox )
		{
			additionalSP = 8;        // account for attached text mso - which has no SPCONTAINER so must include in "controlling" rec
		}
		// 2nd pass:  now have the important container lengths and their updated 
		// subrecords/atoms, create resulting byte array		

		// build main spcontainer - valid for all records 
		MsofbtSpContainer spcontainer1 = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
		spcontainer1.setLength( SPCONTAINERLENGTH + additionalSP );
		byte[] container = spcontainer1.toByteArray();
		SPCONTAINERLENGTH += container.length;    // include this header length    	
    
    	/*// debugging
    	if (!bIsHeader && SPCONTAINERLENGTH!=origSP)
    		Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));
    	 */
		byte[] retData = new byte[SPCONTAINERLENGTH];
		System.arraycopy( container, 0, retData, 0, container.length );
		System.arraycopy( spcontainer1atoms, 0, retData, container.length, spcontainer1atoms.length );

		SPCONTAINERLENGTH += additionalSP;// necessary when summing up container lengths -- see WorkBook.updateMsoHeaderRecord
		if( bIsHeader )
		{
			// SPCONTAINER
			// sp2 -- SPIDSEED
			// spgr if necessary
			MsofbtSpContainer spcontainer2 = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
			spcontainer2.setLength( spcontainer2atoms.length );
			byte[] container2 = spcontainer2.toByteArray();
			SPCONTAINERLENGTH += spcontainer2atoms.length + container2.length;    // include this header length
        	
        	/*// debugging
        	if (SPCONTAINERLENGTH!=origSP)
        		Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));
			/**/
			// SPGRCONTAINER
			int spgrcontainerlen = (SPCONTAINERLENGTH + otherSPCONTAINERLENGTH) - 8;
        	/*// debugging
    	  	if (spgrcontainerlen!=origSPGR)
    	  		Logger.logErr("SPGRCONTAINERLENTH IS OFF: " + (spgrcontainerlen-origSPGR));
    	  	/**/
			MsofbtSpgrContainer msofbtSpgrContainer = new MsofbtSpgrContainer( MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15 );
			msofbtSpgrContainer.setLength( spgrcontainerlen );
			byte[] spgrcontainer = msofbtSpgrContainer.toByteArray();

			// DGCONTAINER
			int dgcontainerlen = (dgcontaineratoms.length + spgrcontainerlen + SOLVERCONTAINERLENGTH + 8);        // drawing container length
        	/*// debugging
    	  	if (dgcontainerlen!=origDG)
    	  		Logger.logErr("DGCONTAINERLENTH IS OFF: " + (dgcontainerlen-origDG));
    	  	/**/
			MsofbtDgContainer msofbtDgContainer = new MsofbtDgContainer( MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15 );
			msofbtDgContainer.setLength( dgcontainerlen );    //HEADERRECLENGTH + SPCONTAINERLENGTH + otherSPCONTAINERLENGTH);
			byte[] dgcontainer = msofbtDgContainer.toByteArray();

			byte[] header = new byte[((HEADERRECLENGTH + SPCONTAINERLENGTH) - additionalSP) + dgcontainer.length]; //+retData.length];

			int pos = 0;
			System.arraycopy( dgcontainer, 0, header, pos, dgcontainer.length );
			pos += dgcontainer.length;
			System.arraycopy( dgcontaineratoms, 0, header, pos, dgcontaineratoms.length );
			pos += dgcontaineratoms.length;
			System.arraycopy( spgrcontainer, 0, header, pos, spgrcontainer.length );
			pos += spgrcontainer.length;
			System.arraycopy( container2, 0, header, pos, container2.length );
			pos += container2.length;
			System.arraycopy( spcontainer2atoms, 0, header, pos, spcontainer2atoms.length );
			pos += spcontainer2atoms.length;

			System.arraycopy( retData, 0, header, pos, retData.length );
			retData = header;
		}
		this.setData( retData );
		this.setLength( data.length );

		// testing
    	/*
    	System.out.println(this.toString());    	
    	System.out.println(this.debugOutput());
		/**/
	}

	/**
	 * add the set of subrecords necessary to define a Mso header record
	 * <br>used when removing images, charts, etc. and have removed a previous header record
	 */
	public void addHeader()
	{
    	/*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
		bIsHeader = true;
		if( lastSPID < SPID )
		{
			lastSPID = SPID;        // TODO: Shouldn't assume to be SPID
		}

		MsofbtSp msofbtSp = new MsofbtSp( MSODrawingConstants.MSOFBTSP,
		                                  MSODrawingConstants.msosptMin,
		                                  2 );  //1st shape rec is of type MSOSPTMIN 
		msofbtSp.setId( SPIDSEED );                      // SPMIN==SPIDSEED          
		msofbtSp.setGrfPersistence( 5 );                    // flag= fPatriarch,                      		
		byte[] msofbtSpBytes = msofbtSp.toByteArray();
		SPCONTAINERLENGTH += msofbtSpBytes.length;

		MsofbtSpgr msofbtSpgr = new MsofbtSpgr( MSODrawingConstants.MSOFBTSPGR, 0, 1 );
		msofbtSpgr.setRect( 0, 0, 0, 0 );
		byte[] msofbtSpgrBytes = msofbtSpgr.toByteArray();
		SPCONTAINERLENGTH += msofbtSpgrBytes.length;

		// SPCONTAINER
		MsofbtSpContainer msofbtSpContainer = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
		msofbtSpContainer.setLength( msofbtSpgrBytes.length + msofbtSpBytes.length );
		byte[] msofbtSpContainerBytes = msofbtSpContainer.toByteArray();
		SPCONTAINERLENGTH += 8;    // account for the SPCONTAINER header length; 

		// SPGRCONTAINER
		MsofbtSpgrContainer msofbtSpgrContainer = new MsofbtSpgrContainer( MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15 );
		msofbtSpgrContainer.setLength( SPCONTAINERLENGTH + otherSPCONTAINERLENGTH );
		byte[] msofbtSpgrContainerBytes = msofbtSpgrContainer.toByteArray();

		MsofbtDg msofbtDg = new MsofbtDg( MSODrawingConstants.MSOFBTDG, drawingId, 0 );
		msofbtDg.setNumShapes( numShapes );                //Number of images and drawings.	
		msofbtDg.setLastSPID( lastSPID );                    //lastSPID      			 
		byte[] msofbtDgBytes = msofbtDg.toByteArray();

		// DGCONTAINER
		MsofbtDgContainer msofbtDgContainer = new MsofbtDgContainer( MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15 );
		msofbtDgContainer.setLength( HEADERRECLENGTH + SPCONTAINERLENGTH + otherSPCONTAINERLENGTH );
		byte[] msofbtDgContainerBytes = msofbtDgContainer.toByteArray();

		byte[] headerRec = new byte[this.getData().length + 80];  // below records take 80 bytes 

		int pos = 0;
		System.arraycopy( msofbtDgContainerBytes, 0, headerRec, pos, msofbtDgContainerBytes.length );
		pos += msofbtDgContainerBytes.length;// 8
		System.arraycopy( msofbtDgBytes, 0, headerRec, pos, msofbtDgBytes.length );
		pos += msofbtDgBytes.length;
		System.arraycopy( msofbtSpgrContainerBytes, 0, headerRec, pos, msofbtSpgrContainerBytes.length );
		pos += msofbtSpgrContainerBytes.length;// 8
		System.arraycopy( msofbtSpContainerBytes, 0, headerRec, pos, msofbtSpContainerBytes.length );
		pos += msofbtSpContainerBytes.length;// 8
		System.arraycopy( msofbtSpgrBytes, 0, headerRec, pos, msofbtSpgrBytes.length );
		pos += msofbtSpgrBytes.length;
		System.arraycopy( msofbtSpBytes, 0, headerRec, pos, msofbtSpBytes.length );
		pos += msofbtSpBytes.length;

		System.arraycopy( this.getData(), 0, headerRec, pos, this.getData().length );
		setData( headerRec );

    	/*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
	}

	/**
	 * remove the set of subrecords necessary to define a MSO header record
	 * <br>used when removing images, charts, etc. and have removed a previous header record
	 */
	public void removeHeader()
	{
    	/*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	*/
		byte[] spcontainer1atoms = new byte[0];
		// reset state vars
		SPCONTAINERLENGTH = 0;        // this shape container length		

		// first pass, update significant atoms and sum up container lengths
		int fbt, len;
		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] header = new byte[8];
			bis.read( header, 0, 8 );
			fbt = ((0xFF & header[3]) << 8) | (0xFF & header[2]);
			len = ByteTools.readInt( header[4], header[5], header[6], header[7] );
			if( (0xF & header[0]) != 15 )
			{    // skip containers
				// parse atoms or sub-records (distinct from container records above)
				byte[] data = new byte[len];
				bis.read( data, 0, len );
				switch( fbt )
				{
					case MSODrawingConstants.MSOFBTDG:    // update drawing record atom
						break;

					case MSODrawingConstants.MSOFBTSP:            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
						// TODO: necessary to ever update the SPIDSEED?
						int flag = ByteTools.readInt( data[4], data[5], data[6], data[7] );
						if( flag != 5 )
						{    // if it's not the SPIDseed, update SPID
							System.arraycopy( ByteTools.cLongToLEBytes( SPID ), 0, data, 0, 4 );
						}
						data = ByteTools.append( data, header );
						if( flag != 5 )    // if it's not the header
						{
							spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						}
						break;

					case MSODrawingConstants.MSOFBTCLIENTANCHOR:        // Anchor or location for a shape 
						// udpate bounds
						System.arraycopy( ByteTools.shortToLEBytes( (short) clientAnchorFlag ), 0, data, 0, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[0] ), 0, data, 2, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[1] ), 0, data, 4, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[2] ), 0, data, 6, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[3] ), 0, data, 8, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[4] ), 0, data, 10, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[5] ), 0, data, 12, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[6] ), 0, data, 14, 2 );
						System.arraycopy( ByteTools.shortToLEBytes( bounds[7] ), 0, data, 16, 2 );
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTOPT:    // property table atom - for 97 and earlier versions
						// OPT= picture options
						if( optrec == null )  // if do not have a MsoFbtOPT record, then create new -- shouldn't get here as should always have an existing record
						{
							optrec = new MsofbtOPT( MSODrawingConstants.MSOFBTOPT,
							                        0,
							                        3 );    //version is always 3, inst is current count of properties.
						}
						if( (imageIndex != -1) && (imageIndex != optrec.getImageIndex()) )
						{
							optrec.setImageIndex( imageIndex );
						}
						if( (imageName != null) && !imageName.equals( "" ) && !imageName.equals( optrec.getImageName() ) )
						{
							optrec.setImageName( imageName );
						}
						if( (shapeName != null) && !shapeName.equals( "" ) && !shapeName.equals( optrec.getShapeName() ) )
						{
							optrec.setShapeName( shapeName );
						}
						data = optrec.toByteArray();
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTTERTIARYOPT:    // for Office versions 2000, XP, 2003, and 2007  
					case MSODrawingConstants.MSOFBTSECONDARYOPT:        // movie (id= 274) 
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTCOLORSCHEME:
					case MSODrawingConstants.MSOFBTREGROUPITEMS:
						break;

					case MSODrawingConstants.MSOFBTCLIENTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...		    	
					case MSODrawingConstants.MSOFBTTEXTBOX: // msofbtClientTextbox sub-record, contains only this one atom no containers ...
						// DON"T ADD LEN
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTSPGR:    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape
						break;

					case MSODrawingConstants.MSOFBTCHILDANCHOR:    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape					
					case MSODrawingConstants.MSOFBTANCHOR:        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
					case MSODrawingConstants.MSOFBTCLIENTDATA:
					case MSODrawingConstants.MSOFBTDELETEDPSPL:
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;

					case MSODrawingConstants.MSOFBTCONNECTORRULE:
					case MSODrawingConstants.MSOFBTALIGNRULE:
					case MSODrawingConstants.MSOFBTARCRULE:
					case MSODrawingConstants.MSOFBTCLIENTRULE:
					case MSODrawingConstants.MSOFBTCALLOUTRULE:
						// SOLVERCONTAINERATOMS+=len+8;
						//don't add to len
						// TODO: HANDLE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
						break;

					default:
						Logger.logInfo( "MSODrawing.removeHeader:  unknown subrecord encountered: " + fbt );
						data = ByteTools.append( data, header );
						spcontainer1atoms = ByteTools.append( data, spcontainer1atoms );
						break;
				}
			}
		}
		SPCONTAINERLENGTH = spcontainer1atoms.length;
		// 2nd pass:  now have the important container lengths and their updated 
		// subrecords/atoms, create resulting byte array		

		// build main spcontainer - valid for all records 
		MsofbtSpContainer spcontainer1 = new MsofbtSpContainer( MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15 );
		spcontainer1.setLength( SPCONTAINERLENGTH );
		byte[] container = spcontainer1.toByteArray();
		SPCONTAINERLENGTH += container.length;    // include this header length    	

		//if (!bIsHeader && SPCONTAINERLENGTH!=origSP)
		//Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));

		byte[] retData = new byte[SPCONTAINERLENGTH];
		System.arraycopy( container, 0, retData, 0, container.length );
		System.arraycopy( spcontainer1atoms, 0, retData, container.length, spcontainer1atoms.length );

		setData( retData );
    	/*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
	}

	/**
	 * update the header records with new container lengths
	 *
	 * @param otherlength
	 */
	public void updateHeader( int otherSPContainers, int otherContainers, int numShapes, int lastSPID )
	{
		if( !isHeader() )
		{
			Logger.logErr( "Msodrawing.updateHeader is only applicable for the header drawing object" );
			return;
		}	  		  		
	  	/*
		// dgcontainerlength= 	dg(+8) + regroupitems(+8) + spgrcontainer + solvercontainer(s) + colorscheme(+8)
		// spgrcontainerlength= 	sum of spcontainers i.e this spcontainerlength + otherspcontainers		
	  	 */
		this.numShapes = numShapes;
		// this.lastSPID= SPIDSEED + nImages; algorithm is wrong when book contains deleted images, etc.
		this.lastSPID = lastSPID;
		int fbt, len;
		// the two container lengths we are concerned about:  
		//SPGRCONTAINERLENGTH 
		otherSPCONTAINERLENGTH = otherSPContainers;
		int spgrcontainerlength = SPCONTAINERLENGTH + otherSPCONTAINERLENGTH;
		// DGCONTAINERLENGTH= spgrcontainerlength + all the atoms contained within the DG CONTAINER
		int dgcontainerlength = otherContainers + spgrcontainerlength;
		int DGATOMS = 0;
		boolean hasSPGRCONTAINER = false;
	  	/* // debugging container lengths on update
		boolean diff= false;
		int origSPGRL= 0;
		int origDGL= 0;
		/**/
		// count dgcontainer atorms=MSOFBTDG, MSOFBTREGROUPITEMS, MSOFBTCOLORSCHEME (MSOFBTSPGRCONTAINER is added separately)
		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
			// 1st pass count dgcontainer atoms
			if( fbt == MSODrawingConstants.MSOFBTDGCONTAINER )
			{
				; // skip contianers
				//origDGL= len;	// debugging
			}
			else if( fbt == MSODrawingConstants.MSOFBTSPCONTAINER )
			{
				;    // skip containers
			}
			else if( fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER )
			{
				if( hasSPGRCONTAINER )
				{    // already has the 1 required SPGRCONTAINER, if more than 1, multiple groups
					spgrcontainerlength += 8;
				}
				DGATOMS += 8;    // just count header length here
				//if (!hasSPGRCONTAINER)	origSPGRL= len;	// debugging	  			
				hasSPGRCONTAINER = true;
			}
			else if( fbt == MSODrawingConstants.MSOFBTOPT )
			{
				break;    // nothing important after this
			}
			else
			{
				if( (fbt == MSODrawingConstants.MSOFBTREGROUPITEMS) || (fbt == MSODrawingConstants.MSOFBTDG) || (fbt == MSODrawingConstants.MSOFBTCOLORSCHEME) )
				{
					DGATOMS += (len + 8);
				}
				buf = new byte[len];
				bis.read( buf, 0, len );
			}
		}
		dgcontainerlength += DGATOMS;
		/** debugging */
		/*if (origSPGRL!=spgrcontainerlength || origDGL!=dgcontainerlength) {			
			System.out.println(this.toString());
			System.out.println(this.debugOutput());
			System.out.println("ORIGDG=" + origDGL + " ORIGSPL=" + origSPGRL + " DIFF: " + (origDGL-dgcontainerlength) + "-" + (origSPGRL-spgrcontainerlength));
			diff= true;
		}
		/**/

		super.getData();
		bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );

			if( fbt == MSODrawingConstants.MSOFBTDGCONTAINER )
			{
				System.arraycopy( ByteTools.cLongToLEBytes( dgcontainerlength ), 0, data, data.length - bis.available() - 4, 4 );
			}
			else if( fbt == MSODrawingConstants.MSOFBTDG )
			{  // Drawing Record: count + MSOSPID seed 
				buf = new byte[len];
				bis.read( buf, 0, len );
				byte[] newrec = new byte[8];
				System.arraycopy( ByteTools.cLongToLEBytes( this.numShapes ), 0, newrec, 0, 4 );
				System.arraycopy( ByteTools.cLongToLEBytes( this.lastSPID ), 0, newrec, 4, 4 );
				if( len == newrec.length )
				{// should!!!
					// update Msodrawing data ...
					System.arraycopy( newrec, 0, data, data.length - bis.available() - newrec.length, newrec.length );
				}
				else
				{
					Logger.logErr( "UpdateClientAnchorRecord: New Array Size=" + newrec.length );
				}
			}
			else if( fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER )
			{    // sum of all spcontainers on the sheet	  
				System.arraycopy( ByteTools.cLongToLEBytes( spgrcontainerlength ), 0, data, data.length - bis.available() - 4, 4 );
				/*if (diff) {	// debugging container lengths
					System.out.println(this.toString());
					System.out.println(this.debugOutput());
				}/**/
				return;
			}
			else
			{    // skip atoms
				buf = new byte[len];
				bis.read( buf, 0, len );
			}
		}
	}

	/**
	 * update just the image index portion of this record
	 * useful when you don't need to rebuild entire record (and thus possibly lose information)
	 *
	 * @param idx
	 */
	public void updateImageIndex( int idx )
	{
		int version, inst, fbt, len;
		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			inst = ((0xFF & buf[1]) << 4) | (0xF0 & buf[0]) >> 4;
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
			if( fbt < 0xF005 )
			{ // ignore containers  
				continue;
			}
			buf = new byte[len];
			bis.read( buf, 0, len );    // position at next sub-rec 
			if( fbt == MSODrawingConstants.MSOFBTOPT )
			{
				// Find the property that governs image index and update the bytes  
				int propertyId;
				int n = inst;                // number of properties to parse
				int pos = 0;                    // pointer to current property in data/property table  
				for( int i = 0; i < n; i++ )
				{
					propertyId = (0x3F & buf[pos + 1]) << 8 | (0xFF & buf[pos]);
					if( propertyId == MSODrawingConstants.msooptpib )
					{// blip to display = image index						
						// testing int dtx = ByteTools.readInt(dat[pos+2],dat[pos+3],dat[pos+4],dat[pos+5]);  
						int insertPosition = (data.length - bis.available() - len) + pos + 2;
						System.arraycopy( ByteTools.cLongToLEBytes( idx ), 0, data, insertPosition, 4 );
						imageIndex = idx;
						return;
					}
					pos += 6;
				}
			}
		}
	}

	/**
	 * removes the header-specific portion of this Msodrawing record, if any
	 * useful when adding Msodrawng recs from other sources such as adding charts
	 * only one Msodrawing header record is allowed
	 */
	public void makeNonHeader()
	{
		if( !bIsHeader )
		{
			return;    // nothing to do!
		}
		bIsHeader = false;
		this.removeHeader();
	}

	/**
	 * set whether "this is the header drawing object for the sheet"
	 *
	 * @param b
	 */
	public void setIsHeader()
	{
		if( !bIsHeader )
		{
			this.addHeader();
		}
	}

	/**
	 * retrieve the OPT rec for specific option setting
	 *
	 * @return
	 */
	public MsofbtOPT getOPTRec()
	{
		return optrec;
	}

	/**
	 * @return whether "this is the header drawing object for the sheet"
	 */
	public boolean isHeader()
	{
		return bIsHeader;
	}

	public String getName()
	{
		return imageName;
	}

	/**
	 * @return explicitly-set shape name (= name in NamedRange box, I believe)
	 */
	public String getShapeName()
	{
		return shapeName;
	}

	/**
	 * @return true if this image has a border
	 */
	// TODO: detecting a border may be more complicated than this
	// TODO: ability to set a border
	public boolean hasBorder()
	{
		if( optrec != null )
		{
			return optrec.hasBorder();
		}
		return false;
	}

	/**
	 * if has a border, return the border line width
	 *
	 * @return
	 */
	public int borderLineWidth()
	{
		if( optrec != null )
		{
			return optrec.getBorderLineWidth();
		}
		return -1;
	}

	public void setImageIndex( int value )
	{
		imageIndex = value;
		this.updateRecord();
	}

	public void setImageName( String name )
	{
		if( !name.equals( imageName ) )
		{
			imageName = name;
			updateRecord();
		}
	}

	/**
	 * set the ordinal # for this drawing record
	 *
	 * @param id
	 */
	public void setDrawingId( int id )
	{
		drawingId = id;
		updateDGRecord();
	}

	/**
	 * return the ordinal # for this record
	 *
	 * @return
	 */
	public int getDrawingId()
	{
		return drawingId;
	}

	/**
	 * allow setting of "Named Range" name
	 *
	 * @param name
	 */
	public void setShapeName( String name )
	{
		if( !name.equals( shapeName ) )
		{
			shapeName = name;
			updateRecord();
		}
	}

	public int getImageIndex()
	{
		return imageIndex;
	}

	// correct way of setting/getting shape (image and chart) bounds
	/* position methods */
	public static final int COL = 0;
	public static final int COLOFFSET = 1;
	public static final int ROW = 2;
	public static final int ROWOFFSET = 3;
	public static final int COL1 = 4;
	public static final int COLOFFSET1 = 5;
	public static final int ROW1 = 6;
	public static final int ROWOFFSET1 = 7;
	public static final int OFFSETMAX = 1023;

	public void setBounds( short[] b )
	{
		bounds = (short[]) b.clone();
		origHeight = calcHeight();    // 20090831 KSC
		updateClientAnchorRecord( bounds );
	}

	/**
	 * set coordinates as {X, Y, W, H} in 7
	 *
	 * @param coords
	 */
	public void setBoundsInPixels( short[] coords )
	{
		// 20090505 KSC: convert 
		short x = (short) Math.round( coords[0] );    // convert pixels to excel units
		short y = (short) Math.round( coords[1] );    // convert pixels to points
		short w = (short) Math.round( coords[2] );    // convert pixels to excel units
		short h = (short) Math.round( coords[3] );    // convert pixels to points
/*		setX(coords[0]);		
		setY(coords[1]);
		setWidth(coords[2]);
		setHeight(coords[3]);
		*/
		setX( x );
		setY( y );
		setWidth( w );
		setHeight( h );
	}

	// conversions:  approx .136 excel units/pixel, 7.5 pixels/excel unit for X/W/Columns (really 6.4????)
	//               approx .75 pixels/points (Y,H/Rows) (really .60???)
	/**
	 * Untested Info from Excel:
	 * width/height in pixels = (w/h field - 8) * DPI of the display device / 72
	 * DPI for Windows is 96 (Mac= 72)
	 * thus conversion factor= 1.3333 for Windows devices
	 * NOTES: seems correct for y/h, not for x/w
	 * return coordinates as {X, Y, W, H} in pixels
	 *
	 * @return
	 */
	// these are approximate conversion factors to convert excel units to pixels ... 
	public static final double XCONVERSION = 10;        // convert excel units to pixels, garnered from actual comparisions rather than any equation (since it's based upon the normal font, it cannot be 100% for all calculations ...
	public static final double WCONVERSION = 8.8;    // ""
	public static final double PIXELCONVERSION = 1.333;    // see above

	/**
	 * return coordinates as {X, Y, W, H} in pixels
	 *
	 * @param coords
	 */
	public short[] getCoords()
	{
		short x = (short) Math.round( (getX() * 256) / ColHandle.COL_UNITS_TO_PIXELS );
		short y = (short) Math.round( (getY() * 20) / (RowHandle.ROW_HEIGHT_DIVISOR - 4) ); //*** WHY need -4 ?????????
		short w = (short) Math.round( (getWidth() * 256) / ColHandle.COL_UNITS_TO_PIXELS );
		short h = (short) Math.round( (calcHeight() * 20) / (RowHandle.ROW_HEIGHT_DIVISOR - 2) ); //*** WHY need -2 ?????????
		
/* testsing new way above		short x= (short) Math.round(getX()*(XCONVERSION);	// convert excel units to pixels
		short y= (short) Math.round(getY()*PIXELCONVERSION);	// convert points to pixels
		short w= (short) Math.round(getWidth()*WCONVERSION);	// convert excel units to pixels
		short h= (short) Math.round((calcHeight()-8)*PIXELCONVERSION);	// convert points to pixels
*/
		return new short[]{ x, y, w, h };
	}

	/**
	 * set coordinates as {X, Y, W, H} in pixels
	 *
	 * @param coords
	 */
	public void setCoords( short[] coords )
	{
		// TODO should use ABOVE CALC!!!! see getCoords ******************
		short x = (short) Math.round( (coords[0] * ColHandle.COL_UNITS_TO_PIXELS) / 256 );                // convert pixels to excel units
		short y = (short) Math.round( (coords[1] * (RowHandle.ROW_HEIGHT_DIVISOR - 4)) / 20 );                // convert pixels to points
		short w = (short) Math.round( (coords[2] * ColHandle.COL_UNITS_TO_PIXELS) / 256 );        // convert pixels to excel units
		short h = (short) Math.round( (coords[3] * (RowHandle.ROW_HEIGHT_DIVISOR - 2)) / 20 );        // convert pixels to points
		setX( x );
		setY( y );
		setWidth( w );
		setHeight( h );
	}

	/**
	 * returns the bounds of this object
	 * bounds are relative and based upon rows, columns and offsets within
	 */
	public short[] getBounds()
	{
		return bounds;
	}

	public int getCol()
	{
		return bounds[MSODrawing.COL];
	}

	public int getCol1()
	{
		return bounds[MSODrawing.COL1];
	}

	public int getRow0()
	{
		return bounds[MSODrawing.ROW];
	}

	public int getRow1()
	{
		return bounds[MSODrawing.ROW1];
	}

	public void setRow( int row )
	{
		bounds[MSODrawing.ROW] = (short) row;
		updateClientAnchorRecord( bounds );
	}

	public void setRow1( int row )
	{
		bounds[MSODrawing.ROW1] = (short) row;
		updateClientAnchorRecord( bounds );
	}

	/**
	 * get X value of upper left corner
	 * units are in excel units
	 *
	 * @return short x value
	 */
	public short getX()
	{
		int col = bounds[MSODrawing.COL];
		double colOff = bounds[MSODrawing.COLOFFSET] / 1024.0;
		double x = 0.0;
		for( int i = 0; i < col; i++ )
		{
			x += getColWidth( i );
		}
		x += colOff * getColWidth( col );
		return (short) Math.round( x );
	}

	/**
	 * returns the offset within the column in pixels
	 *
	 * @return
	 */
	public short getColOffset()
	{
		double colOff = bounds[MSODrawing.COLOFFSET] / 1024.0;
		int col = bounds[MSODrawing.COL];
		double x = colOff * getColWidth( col );
		return (short) (Math.round( x ) * XCONVERSION);
	}

	/**
	 * set the x position of this object
	 * units are in excel units
	 *
	 * @param x
	 */
	public void setX( int x )
	{
		int z = 0;
		short col = 0;
		short colOffset = 0;
		for( short i = 0; (i < XLSConstants.MAXCOLS) && (z < x); i++ )
		{
			int w = getColWidth( i );
			if( (z + w) < x )
			{
				z += w;
			}
			else
			{
				col = i;
				colOffset = (short) Math.round( 1024 * (((double) (x - z)) / (double) w) );
				z = x;
			}
		}
		bounds[MSODrawing.COL] = col;
		bounds[MSODrawing.COLOFFSET] = colOffset;
		updateClientAnchorRecord( bounds );
	}

	/**
	 * return the y position of this object in points
	 *
	 * @return
	 */
	public short getY()
	{
		int row = bounds[MSODrawing.ROW];
		double y = 0.0;
		for( int i = 0; i < row; i++ )
		{
			y += getRowHeight( i );
		}
		double rowOff = bounds[MSODrawing.ROWOFFSET] / 256.0;
		y += getRowHeight( row ) * rowOff;
		return (short) Math.round( y );
	}

	/**
	 * set the y position of this object
	 * units are in points
	 *
	 * @param y
	 */
	public void setY( int y )
	{
		double z = 0;
		short row = 0;
		short rowOffset = 0;
		for( short i = 0; z < y; i++ )
		{
			double h = getRowHeight( i );
			if( (z + h) < y )
			{
				z += h;
			}
			else
			{
				row = i;
				rowOffset = (short) Math.round( (256 * ((double) (y - z) / getRowHeight( i ))) );
				z = y;
			}
		}
		bounds[MSODrawing.ROW] = row;
		bounds[MSODrawing.ROWOFFSET] = rowOffset;
		updateClientAnchorRecord( bounds );
	}

	/**
	 * public methd that returns saved width value;
	 * useful when calculated method is incorrect due to column width changes ...
	 *
	 * @return
	 */
	public short getOriginalWidth()
	{
		return origWidth;
	}

	/**
	 * calculate width based upon col#'s, coloffsets and col widths
	 * units are in excel column units
	 *
	 * @return short width
	 */
	public short getWidth()
	{
    	/* 
		bounds[0]= column # of top left position (0-based) of the shape
		bounds[1]= x offset within the top-left column
		bounds[2]= row # for top left corner
		bounds[3]= y offset within the top-left corner
		bounds[4]= column # of the bottom right corner of the shape
		bounds[5]= x offset within the cell  for the bottom-right corner
		bounds[6]= row # for bottom-right corner of the shape
		bounds[7]= y offset within the cell for the bottom-right corner		
		*/

		int col = bounds[MSODrawing.COL];
		double colOff = bounds[MSODrawing.COLOFFSET] / 1024.0;
		int col1 = bounds[MSODrawing.COL1];
		double colOff1 = bounds[MSODrawing.COLOFFSET1] / 1024.0;
		double w = getColWidth( col ) - (getColWidth( col ) * colOff);
		for( int i = col + 1; i < col1; i++ )
		{
			w += getColWidth( i );
		}
		if( col1 > col )
		{
			w += getColWidth( col1 ) * colOff1;
		}

		else    //correct????  
		{
			w = getColWidth( col1 ) * (colOff1 - colOff);
		}

		return (short) Math.round( w );
	}

	/**
	 * public method that returns saved height value;
	 * used when calculated method is incorrect due to row height changes ...
	 *
	 * @return
	 */
	public short getHeight()
	{
		return origHeight;
	}

	/**
	 * calculate height based upon row #s, row offsets and row heights
	 * units are in points
	 *
	 * @return short row height
	 */
	private short calcHeight()
	{
		int row = bounds[MSODrawing.ROW];
		int row1 = bounds[MSODrawing.ROW1];
		double rowOff = bounds[MSODrawing.ROWOFFSET] / 256.0;
		double rowOff1 = bounds[MSODrawing.ROWOFFSET1] / 256.0;
		double y = getRowHeight( row ) - (getRowHeight( row ) * rowOff);
		for( int i = row + 1; i < row1; i++ )
		{
			y += getRowHeight( i );
		}
		if( row1 > row )
		{
			y += getRowHeight( row1 ) * rowOff1;
		}
		else
		{
			y = getRowHeight( row1 ) * (rowOff1 - rowOff);
		}
		return (short) Math.round( y );
	}

	/**
	 * set the width of this object
	 * units are in excel units
	 *
	 * @param w
	 */
	public void setWidth( int w )
	{
		int col = bounds[MSODrawing.COL];
		double colOff = bounds[MSODrawing.COLOFFSET] / 1024.0;
		int col1 = col;
		int colOff1 = 0;
		int z = getColWidth( col ) - (int) (getColWidth( col ) * colOff);
		if( z >= w )
		{    // 20100322 KSC: was > w, so the == case never was handled, see TestImages.insertImageOffsetDisappearance bug
			col1 = col;
			colOff1 = (short) Math.round( (1024 * (w / (double) getColWidth( col ))) ) + bounds[MSODrawing.COLOFFSET];
		}
		for( int i = col + 1; (i < XLSConstants.MAXCOLS) && (z < w); i++ )
		{
			int cw = getColWidth( i );
			if( (z + cw) < w )
			{
				z += cw;
			}
			else
			{
				col1 = i;
				colOff1 = (short) (1024 * (((double) (w - z)) / (double) cw));
				z = w;
			}
		}
		bounds[MSODrawing.COL1] = (short) col1;
		bounds[MSODrawing.COLOFFSET1] = (short) colOff1;
		updateClientAnchorRecord( bounds );
		origWidth = getOriginalWidth();    // 20071024 KSC: if change width, record
	}

	/**
	 * set the height of this object
	 * units are in points
	 *
	 * @param h
	 */
	public void setHeight( int h )
	{
		int row = bounds[MSODrawing.ROW];
		double rowOff = bounds[MSODrawing.ROWOFFSET] / 256.0;
		int row1 = row;
		int rowOff1 = 0;
		double rh = getRowHeight( row );
		double y = rh - (rh * rowOff);    // distance from start position to end of row
		if( y > h )
		{
			rowOff1 = (short) (256 * (h / rh)) + bounds[MSODrawing.ROWOFFSET];
		}
		for( int i = row + 1; y < h; i++ )
		{
			rh = getRowHeight( i );
			if( (y + rh) < h )
			{
				y += rh;
			}
			else
			{        // height is met; see what offset into row i is necessary
				row1 = i;
				rowOff1 = (short) Math.round( (256 * ((double) (h - y) / rh)) );
				y = h;    // exit loop
			}
		}
		bounds[MSODrawing.ROW1] = (short) row1;
		bounds[MSODrawing.ROWOFFSET1] = (short) rowOff1;
		updateClientAnchorRecord( bounds );
		origHeight = calcHeight();    // 20071024 KSC: if change height, record
	}

	/* col/row and offset methods */
	public short[] getColAndOffset()
	{
		return new short[]{ bounds[COL], bounds[COLOFFSET] };
	}

	public void setColAndOffset( short[] b )
	{
		bounds[COL] = b[0];
		bounds[COLOFFSET] = b[1];
		updateClientAnchorRecord( bounds );
	}

	public short[] getRowAndOffset()
	{
		return new short[]{ bounds[ROW], bounds[ROWOFFSET] };
	}

	public void setRowAndOffset( short[] b )
	{
		bounds[ROW] = b[0];
		bounds[ROWOFFSET] = b[1];
		updateClientAnchorRecord( bounds );
	}

	/**
	 * return the column width in excel units
	 *
	 * @param col
	 * @return
	 */
	private int getColWidth( int col )
	{
		// MSO column width is in Excel units= 0-255 characters
		// Excel Units are 1/256 of default font width (10 pt Arial)
		double w = Colinfo.DEFAULT_COLWIDTH;
		try
		{
			Colinfo co = this.getSheet().getColInfo( col );
			if( co != null )
			{
				w = co.getColWidth();
			}
		}
		catch( Exception e )
		{    // exception if no col defined
		}

		return (int) Math.round( w / 256 );        // 20090505 KSC: try this	
	}

	/**
	 * return the row height in points
	 *
	 * @param row
	 * @return
	 */
	private double getRowHeight( int row )
	{
		// MSO row height is measured in points, 0-409 
		// in Arial 10 pt, standard row height is 12.75 points
		// .75 points/pixel
		// row height is in twips= 1/20 of a point,1 pt= 1/72 inch
		int h = 255;
		try
		{
			Row r = this.getSheet().getRowByNumber( row );
			if( r != null )
			{
				h = r.getRowHeight();
			}
			else    // no row defined - use default row height 20100504 KSC
			{
				return this.getSheet().getDefaultRowHeight(); // default
			}
		}
		catch( Exception e )
		{    // exception if no row defined // no row defined - use default row height 20100504 KSC		
			return this.getSheet().getDefaultRowHeight(); // default
		}
		return (h / 20.0);        // 20090506 KSC: it's in twips 1/20 of a point	
	}
	/* end position methods */

	private Phonetic phonetic = null;

	/**
	 * @return
	 */
	public Phonetic getMystery()
	{
		return phonetic;
	}

	/**
	 * @param phonetic
	 */
	public void setMystery( Phonetic p )
	{
		phonetic = p;
	}

	/**
	 * @return Returns the numShapes.
	 */
	public int getNumShapes()
	{
		return numShapes;
	}

	/**
	 * set the number of shapes for this drawing rec
	 *
	 * @param n
	 */
	public void setNumShapes( int n )
	{
		numShapes = n;
		updateDGRecord();
	}

	/**
	 * return the lastSPID (only valid for the header msodrawing record
	 * necessary to track so that newly added images have the appropriate SPID
	 *
	 * @return
	 */
	public int getlastSPID()
	{
		return this.lastSPID;
	}

	/**
	 * returns the Shape Container Length for this drawing record
	 * used when setting total size for the header drawing record
	 *
	 * @return
	 */
	public int getSPContainerLength()
	{
		return SPCONTAINERLENGTH;
	}

	/**
	 * returns the solver container length for this drawing record
	 * used when setting the total size for the header drawing record
	 *
	 * @return
	 */
	public int getSOLVERContainerLength()
	{
		return SOLVERCONTAINERLENGTH;
	}

	protected static XLSRecord getPrototype()
	{
		MSODrawing mso = new MSODrawing();
		mso.setOpcode( MSODRAWING );
		mso.setData( mso.PROTOTYPE_BYTES );
		mso.init();
		// prototype bytes contain an image index and an image name - remove here
		// TODO: make prototype bytes correct
//        mso.imageIndex= -1;
//    	mso.imageName= ""; 
//        mso.optrec.setImageName("");
//        mso.optrec.setImageIndex(-1);
//        mso.updateRecord();
		return mso;
	}

	/**
	 * creates a msofbtClientTextbox, a very strange mso record containing only one sub-record or atom;
	 * associated data=the text in the textbox, in a host-defined format i.e. text from a NOTE
	 * <br>
	 * NOTE: this must be flagged as incomplete so that it is not counted in number of shapes and other calcs ...
	 * <br>
	 * NOTE: knowledge of this record is very sketchy ...
	 *
	 * @return
	 */
	protected static XLSRecord getTextBoxPrototype()
	{
		MSODrawing mso = new MSODrawing();
		mso.setOpcode( MSODRAWING );
		mso.setData( new byte[]{
				0, 0, 13, -16, 0, 0, 0, 0
		} );    // only contains 1 sub-record:  0xF00D or 61453= msofbtTextBox==shape has attached text
		mso.init();
		return mso;
	}

	/**
	 * update the bytes of the DG record
	 */
	private void updateDGRecord()
	{
		int fbt, len;

		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
			if( fbt < 0xF005 )
			{ // ignore containers  
				continue;
			}
			buf = new byte[len];
			bis.read( buf, 0, len );    // position at next sub-rec 
			if( fbt == MSODrawingConstants.MSOFBTDG )
			{  // Drawing Record: count + MSOSPID seed 
				data[8] = (byte) (drawingId * 16);    //= the 1st byte of the header portion of the DG record 20080902 KSC				
				byte[] newrec = new byte[8];
				System.arraycopy( ByteTools.cLongToLEBytes( numShapes ), 0, newrec, 0, 4 );
				System.arraycopy( ByteTools.cLongToLEBytes( this.lastSPID ), 0, newrec, 4, 4 );
				if( len == newrec.length )
				{// should!!!
					// update Msodrawing data ...
					System.arraycopy( newrec, 0, data, data.length - bis.available() - newrec.length, newrec.length );
				}
				else
				{
					Logger.logErr( "UpdateClientAnchorRecord: New Array Size=" + newrec.length );
				}
				return;
			}
		}
	}

	/**
	 * update the bytes of the CLIENTANCHOR record
	 *
	 * @param bounds bounds[0]= column # of top left position (0-based) of the shape
	 *               bounds[1]= x offset within the top-left column
	 *               bounds[2]= row # for top left corner
	 *               bounds[3]= y offset within the top-left corner
	 *               bounds[4]= column # of the bottom right corner of the shape
	 *               bounds[5]= x offset within the cell  for the bottom-right corner
	 *               bounds[6]= row # for bottom-right corner of the shape
	 *               bounds[7]= y offset within the cell for the bottom-right corner
	 */
	private void updateClientAnchorRecord( short[] bounds )
	{
		int fbt, len;

		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
			if( fbt < 0xF005 )
			{ // ignore containers  
				continue;
			}
			buf = new byte[len];
			bis.read( buf, 0, len );    // position at next sub-rec 
			if( fbt == MSODrawingConstants.MSOFBTCLIENTANCHOR )
			{        // Anchor or location fo a shape
				// udpate bounds
				ByteArrayOutputStream bos = new ByteArrayOutputStream();
				try
				{
					bos.write( ByteTools.shortToLEBytes( (short) clientAnchorFlag ) );
					bos.write( ByteTools.shortToLEBytes( bounds[0] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[1] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[2] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[3] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[4] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[5] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[6] ) );
					bos.write( ByteTools.shortToLEBytes( bounds[7] ) );
				}
				catch( Exception e )
				{
				}
				byte[] newrec = bos.toByteArray();
				if( buf.length == newrec.length )
				{// should!!!
					// update Msodrawing data ...
					System.arraycopy( newrec, 0, data, data.length - bis.available() - newrec.length, newrec.length );
				}
				else
				{
					Logger.logErr( "UpdateClientAnchorRecord: New Array Size=" + newrec.length );
				}
				return;
			}
		}
	}

	/**
	 * sets a specific OPT subrecord
	 *
	 * @param propertyId   int property id see Msoconstants
	 * @param isBid        true if this is a BLIP id
	 * @param isComplex    true if has complexBytes
	 * @param dtx          if not isComplex, the value; if isComplex, length
	 * @param complexBytes complex bytes if isComplex
	 */
	public void setOPTSubRecord( int propertyId, boolean isBid, boolean isComplex, int dtx, byte[] complexBytes )
	{
		MsofbtOPT optrec = this.getOPTRec();
		// TODO: store optrec length instead of calculating each time
		int origlen = optrec.toByteArray().length;
		optrec.setProperty( MSODrawingConstants.msooptGroupShapeProperties, isBid, isComplex, dtx, complexBytes );
		updateRecord();
		if( origlen != optrec.toByteArray().length )
		{    // must update header
			this.getWorkBook().updateMsodrawingHeaderRec( this.getSheet() );
		}
	}

	/**
	 * set the LastSPID for this drawing record, if it's a header-type record
	 *
	 * @param spid
	 */
	public void setLastSPID( int spid )
	{
		lastSPID = spid;
		updateDGRecord();
	}

	/**
	 * set the SPID for this drawing record
	 * used upon copyworksheet ...
	 *
	 * @param spid
	 */
	public void setSPID( int spid )
	{
		SPID = spid;
		updateSPID();
	}

	/**
	 * return the SPID for this drawing record
	 */
	public int getSPID()
	{
		return SPID;
	}

	public void setShapeType( int shapeType )
	{
		this.shapeType = (short) shapeType;
	}

	public int getShapeType()
	{
		return shapeType;
	}

	/**
	 * change the SPID for this record
	 *
	 * @param spid
	 */
	private void updateSPID()
	{
		int fbt, len, inst;
		super.getData();
		ByteArrayInputStream bis = new ByteArrayInputStream( data );
		for(; bis.available() > 0; )
		{
			byte[] buf = new byte[8];
			bis.read( buf, 0, 8 );
			inst = ((0xFF & buf[1]) << 4) | (0xF0 & buf[0]) >> 4;
			fbt = ((0xFF & buf[3]) << 8) | (0xFF & buf[2]);
			len = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
			if( fbt >= 0xF005 )
			{ // ignore containers  
				buf = new byte[len];
				bis.read( buf, 0, len );

				if( fbt == MSODrawingConstants.MSOFBTSP )
				{
					//byte[] dat = new byte[len];
					//bis.read(dat,0,len);
					int flag = ByteTools.readInt( buf[4], buf[5], buf[6], buf[7] );
//					if (flag!=5) {	// if it's not the SPIDseed
					System.arraycopy( ByteTools.cLongToLEBytes( SPID ), 0, data, data.length - bis.available() - len, 4 );
					//SPID = ByteTools.readInt(dat[0],dat[1],dat[2],dat[3]);		    	
//						return;
//					}
				}
			}
		}
	}

	/**
	 * update the header records with new container lengths
	 *
	 * @param otherlength
	 */
	public String toString()
	{
		StringBuffer sb = new StringBuffer();
		sb.append( "Msodrawing: image=" + this.imageName + ".\t" + this.shapeName + "\timageIndex=" + imageIndex + " sheet=" + ((this.getSheet() != null) ? this
				.getSheet()
				.getSheetName() : "none") );
		sb.append( " ID= " + drawingId + " SPID=" + SPID );
		sb.append( "\tShapeType= " + shapeType );
		if( isHeader() )
		{
			sb.append( "\tNumber of Shapes=" + numShapes + " Last SPID=" + lastSPID + " SPCL=" + SPCONTAINERLENGTH + " otherLen=" + otherSPCONTAINERLENGTH );
		}
		else
		{
			sb.append( " SPCL=" + SPCONTAINERLENGTH );
		}
		if( !isShape )
		{
			sb.append( " NOT A SHAPE" );
		}
		return sb.toString();
	}

	/**
	 * 20081106 KSC: when set sheet, record original height and width
	 * as dependent upon row heights ...
	 */
	@Override
	public void setSheet( Sheet bs )
	{
		super.setSheet( bs );
		// 20081106 Moved from parse as sheet must be set before calcHeight call
		// Record original height and width in case of row height/col width changes
		origHeight = calcHeight();
		origWidth = getOriginalWidth();
	}

	/**
	 * returns true if this drawing is active i.e. not deleted
	 * <br>Note this is experimental
	 *
	 * @return
	 */
	public boolean isActive()
	{
		// TODO: also report falses if height or width is 0??
		return bActive;
	}

	/**
	 * create records sub-records necessary to define an AutoFilter drop-down symbol
	 */
	public void createDropDownListStyle( int col )
	{
		// mso record which - we hope - has the specific options necessary to define the dropdown box
		MsofbtOPT optrec = this.getOPTRec();
		optrec.setInst( 0 ); // clear out
		optrec.setData( new byte[]{ } );
		optrec.setProperty( MSODrawingConstants.msooptfLockAgainstGrouping, false, false, 17039620, null );
		optrec.setProperty( MSODrawingConstants.msooptfFitTextToShape, false, false, 524296, null );
		optrec.setProperty( MSODrawingConstants.msooptfNoLineDrawDash, false, false, 524288, null );
		optrec.setProperty( MSODrawingConstants.msooptGroupShapeProperties, false, false, 131072, null );
		this.setShapeType( MSODrawingConstants.msosptHostControl );    // shape type for these drop-downs
		this.updateRecord( ++this.wkbook.lastSPID, null, null, -1, new short[]{
				(short) col, 0, 0, 0, (short) (col + 1), 0, 1, 0
		} );  // generate msoDrawing using correct values moved from above
		// KSC: keep for testing
		//col++;
		//this.updateRecord(++this.wkbook.lastSPID, null, null, -1, new short[] {(short)col, 0, 0, 0, (short)(col+1), 272, 1, 0});  // generate msoDrawing using correct values moved from above
	}

	/**
	 * create the sub-records necessary to define a Comment (Note) Box
	 */
	public void createCommentBox( int row, int col )
	{
		MsofbtOPT optrec = this.getOPTRec();
		optrec.setInst( 0 ); // clear out	
		optrec.setData( new byte[]{ } );
		//47752196
		// try a random text id ... instead of 48678864
		int id = new java.util.Random().nextInt();
		optrec.setProperty( MSODrawingConstants.msofbtlTxid, false, false, id, null );
		optrec.setProperty( MSODrawingConstants.msofbttxdir, false, false, 2, null );
		optrec.setProperty( MSODrawingConstants.msooptfFitTextToShape, false, false, 524296, null );
		optrec.setProperty( 344, false, false, 0, null );    // no info on this subrecord
		optrec.setProperty( MSODrawingConstants.msooptfillColor, false, false, 134217808, null );
		optrec.setProperty( MSODrawingConstants.msooptfillBackColor, false, false, 134217808, null );
		optrec.setProperty( MSODrawingConstants.msooptfNoFillHitTest, false, false, 1048592, null );
		optrec.setProperty( MSODrawingConstants.msooptfshadowColor, false, false, 0, null );
		optrec.setProperty( MSODrawingConstants.msooptfShadowObscured, false, false, 196611, null );
		// this, strangely, controls note hidden or show- when it's 131074, it's hidden, when it's 131072, it's shown
		optrec.setProperty( MSODrawingConstants.msooptGroupShapeProperties, false, false, 131074, null );
		this.setShapeType( MSODrawingConstants.msosptTextBox );    // shape type for text boxes 
		// position of text box - garnered from Excel examples
		// [1, 240, 0, 30, 3, 496, 4, 196]	A1
		// [4, 240, 2, 105, 6, 496, 7, 15]  D4
		this.updateRecord( ++this.wkbook.lastSPID, null, null, -1, new short[]{
				(short) (col + 1), 240, (short) row, 30, (short) (col + 3), 496, (short) (row + 4), 196
		} );  // generate msoDrawing using correct values moved from above
	}
    /* notes from another attempt:  does this match ours?
_store_mso_opt_comment {
    my $self        = shift;
    my $type        = 0xF00B;
    my $version     = 3;
    my $instance    = 9;
    my $data        = '';
    my $length      = 54;
    my $spid        = $_[0];
    my $visible     = $_[1];
    my $colour      = $_[2] || 0x50;
    $data    = pack "V",  $spid;
    $data   .= pack "H*", '0000BF00080008005801000000008101' ;
    $data   .= pack "C",  $colour;
    $data   .= pack "H*", '000008830150000008BF011000110001' .
                          '02000000003F0203000300BF03';
    $data   .= pack "v",  $visible;
    $data   .= pack "H*", '0A00';
     * 
     */

	/**
	 * test debug output - FOR INTERNAL USE ONLY
	 *
	 * @return
	 */
	public String debugOutput()
	{
		ByteArrayInputStream bis = new ByteArrayInputStream( super.getData() );
		int version, inst, fbt, len;
		StringBuffer log = new StringBuffer();
		try
		{
			for(; bis.available() > 0; )
			{
				byte[] dat = new byte[8];
				bis.read( dat, 0, 8 );
				version = (0xF & dat[0]);
				inst = ((0xFF & dat[1]) << 4) | (0xF0 & dat[0]) >> 4;
				fbt = ((0xFF & dat[3]) << 8) | (0xFF & dat[2]);
				len = ByteTools.readInt( dat[4], dat[5], dat[6], dat[7] );

				log.append( fbt + " " + version + "/" + inst + "/" + len );

//		    if (fbt <= 0xF005) { // ignore containers
				if( version == 15 )
				{ // it's a container - no parsing
					// MSOFBTSPGRCONTAINER:		// Shape Group Container, contains a variable number of shapes (=msofbtSpContainer) + other groups 0xF003 
					// MSOFBTSPCONTAINER:		// Shape Container 0xF004
					if( fbt == MSODrawingConstants.MSOFBTDGCONTAINER )
					{
						log.append( "\tMSOFBTDGCONTAINER" );
					}
					else if( fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER )
					{
						log.append( "\tMSOFBTSPGRCONTAINER" );
					}
					else if( fbt == MSODrawingConstants.MSOFBTSPCONTAINER )
					{
						log.append( "\tMSOFBTSPCONTAINER" );
					}
					else if( fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER )
					{
						log.append( "\tMSOFBTSOLVERCONTAINER" );
					}
					else
					{
						log.append( "\tUNKNOWN CONTAINER" );
					}
					log.append( "\r\n" );
					continue;
				}

				dat = new byte[len];
				bis.read( dat, 0, len );
				switch( fbt )
				{
					case MSODrawingConstants.MSOFBTCALLOUTRULE:
						log.append( "\tMSOFBTCALLOUTRULE" );
						break;

					case MSODrawingConstants.MSOFBTDELETEDPSPL:
						log.append( "\tMSOFBTDELETEDPSPL" );
						break;

					case MSODrawingConstants.MSOFBTREGROUPITEMS:
						log.append( "\tMSOFBTREGROUPITEMS" );
						break;

					case MSODrawingConstants.MSOFBTSP:            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
						log.append( "\tMSOFBTSP" );
						int flag;
						flag = ByteTools.readInt( dat[4], dat[5], dat[6], dat[7] );
						if( (flag & 0x800) == 0x800 )// it has a shape-type property
						{
							log.append( "\tshapeType=" + inst );
						}
						if( inst == 0 )    //== shape type	
						{
							log.append( "\tSPIDSEED=" + ByteTools.readInt( dat[0], dat[1], dat[2], dat[3] ) );
						}
						else
						{
							log.append( "\tSPID=" + +ByteTools.readInt( dat[0], dat[1], dat[2], dat[3] ) );
						}
						log.append( "\tflag=" + flag );
						break;

					case MSODrawingConstants.MSOFBTCLIENTANCHOR:        // Anchor or location fo a shape
						log.append( "\tMSOFBTCLIENTANCHOR" );
						log.append( "\t[" );
						log.append( ByteTools.readShort( dat[2], dat[3] ) + "," );
						log.append( ByteTools.readShort( dat[4], dat[5] ) + "," );
						log.append( ByteTools.readShort( dat[6], dat[7] ) + "," );
						log.append( ByteTools.readShort( dat[8], dat[9] ) + "," );
						log.append( ByteTools.readShort( dat[10], dat[11] ) + "," );
						log.append( ByteTools.readShort( dat[12], dat[13] ) + "," );
						log.append( ByteTools.readShort( dat[14], dat[15] ) + "," );
						log.append( ByteTools.readShort( dat[16], dat[17] ) );
						log.append( "]" );
						break;

					case MSODrawingConstants.MSOFBTOPT:    // property table atom
						log.append( "\tMSOFBTOPT" );
						//MsofbtOPT optrec= new MsofbtOPT(fbt, inst, version);
						//optrec.setData(dat);	// sets and parses msoFbtOpt data, including imagename, shapename and imageindex
						log.append( optrec.debugOutput() );
						break;

					case MSODrawingConstants.MSOFBTSECONDARYOPT:    // property table atom  block 2
						log.append( "\tMSOFBTSECONDARYOPT" );
						MsofbtOPT secondaryoptrec = new MsofbtOPT( fbt, inst, version );
						secondaryoptrec.setData( dat );    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
						log.append( secondaryoptrec.debugOutput() );
						break;

					case MSODrawingConstants.MSOFBTTERTIARYOPT:    // property table atom  block 3
						log.append( "\tMSOFBTTERTIARYOPT" );
						MsofbtOPT tertiaryoptrec = new MsofbtOPT( fbt, inst, version );
						tertiaryoptrec.setData( dat );    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
						log.append( tertiaryoptrec.debugOutput() );
						break;

					case MSODrawingConstants.MSOFBTDG:    // Drawing Record: ID, num shapes + Last SPID for this DG
						log.append( "\tMSOFBTDG" );
						log.append( "\tID=" + inst );
						log.append( "\tns=" + ByteTools.readInt( dat[0], dat[1], dat[2], dat[3] ) );  // number of shapes in this drawing
						log.append( "\tllastSPID=" + ByteTools.readInt( dat[4], dat[5], dat[6], dat[7] ) );
						break;

					case MSODrawingConstants.MSOFBTCLIENTTEXTBOX:
						log.append( "\tMSOFBTCLIENTTEXTBOX" );
						break;

					case MSODrawingConstants.MSOFBTSPGR:
						log.append( "\tMSOFBTSPGR" );
						break;

					case MSODrawingConstants.MSOFBTCLIENTDATA:
						log.append( "\tMSOFBTCLIENTDATA" );
						break;

					case MSODrawingConstants.MSOFBTSOLVERCONTAINER:
						log.append( "\tMSOFBTSOLVERCONTAINER" );
						break;

					case MSODrawingConstants.MSOFBTCHILDANCHOR:
						log.append( "\tMSOFBTCHILDANCHOR" );
						break;

					case MSODrawingConstants.MSOFBTCONNECTORRULE:
						log.append( "\tMSOFBTCONNECTORRULE" );
						break;

					default:    //MSOFBTCONNECTORRULE
						log.append( "\tUNKNOWN ATOM" );
						break;
				}
				log.append( "\r\n" );
			}
		}
		catch( Exception e )
		{
			log.append( "\r\nEXCEPTION: " + e.toString() );
		}
		log.append( "**" );
		return log.toString();
	}

	@Override
	public void close()
	{
		super.close();
		this.bounds = null;
		optrec = null;    // 20091209 KSC: save MsofbtOPT records for later updating ...
		secondaryoptrec = null;
		tertiaryoptrec = null;    // apparently can have secondary and tertiary obt recs depending upon version in which it was saved ...
	}

}