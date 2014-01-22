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

import com.extentech.formats.XLS.charts.Chart;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.Serializable;
import java.util.ArrayList;

/**
 * <b>OBJ Describes a graphic object 0x5D</b><br>
 * <p/>
 * MsoDrawing, MsoDrawngGroup + MsoDrawingSelection contain remaining drawing object data
 * Each subrecord begins with a 2byte id number, ft.  Next, a two-byte length field, b, specifies the length
 * of the subrecord data.  The subrecord data field follws the length field.
 * <p/>
 * In BIFF8, the OBJ record contains a partial description of a drawing object, and the MSODRAWING, MSODRAWINGGROUP, and MSODRAWINGSELECTION records
 * contain the remaining drawing object data.
 * <p/>
 * To store an OBJ record in BIFF8, Microsoft Excel writes a collection of subrecords. The structure of a subrecord is identical to the
 * structure of a BIFF record. Each subrecord begins with a 2-byte id number, ft (see the following table).
 * Next a 2-byte length field, cb, specifies the length of the subrecord data field. The subrecord data field follows the length field.
 * First subrecord is always ftCmo (common object data) and the last is always ftEnd.
 * <p/>
 * For all records other than ftEnd and ftCmo, the subrecord contents are as follows:
 * NOTE: THIS IS NOT NECESSARILY TRUE IF THE SUBRECORD CONTAINS CONTINUES ******** see ftClbsData
 * <p/>
 * Offset	Name	Size	Contents
 * 0		ft		2		= subrecord id (0x4 to 0x14, depending on which record)
 * 2		cb		2		Length of data
 * 4		Reserved var
 * <p/>
 * POSSIBLE SUBRECORDS:
 * ftEnd			00h			End of OBJ record	(4 bytes of 0)
 * (Reserved)		01h
 * (Reserved)		02h
 * (Reserved)		03h
 * ftMacro			04h			Fmla-style macro
 * ftButton		05h			Command button
 * ftGmo			06h			Group marker
 * ftCf			07h			Clipboard format
 * ftPioGrbit		08h			Picture option flags
 * ftPictFmla		09h			Picture fmla-style macro
 * ftCbls			0Ah			Check box link
 * ftRbo			0Bh			Radio button
 * ftSbs			0Ch			Scroll bar
 * ftNts			0Dh			Note structure
 * ftSbsFmla		0Eh			Scroll bar fmla-style macro
 * ftGboData		0Fh			Group box data
 * ftEdoData		10h			Edit control data
 * ftRboData		11h			Radio button data
 * ftCblsData		12h			Check box data
 * ftLbsData		13h			List box data
 * ftCblsFmla		14h			Check box link fmla-style macro
 *
 * @see Obj.init
 * @see Obj.getBasicObjRecord
 */
public final class Obj extends com.extentech.formats.XLS.XLSRecord
{
	private FtCmo ftCmo = null;    // FtCmo object contains common object properties such as object type
	private static final long serialVersionUID = -4442755911399227290L;
	// defines a PICTURE-type Object Record
	public byte[] PROTOTYPE_BYTES = {
			21, 0, 18, 0, 8, 0, 2, 0, 17, 96, 0, 0, 0, 0, 104, 28, 105, 5, 0, 0, 0, 0, 7, 0, 2, 0, -1, -1, 8, 0, 2, 0, 0, 0, 0, 0, 0, 0
	};
	private Chart myChart;
	// 20100430 KSC: the only Continue following an Obj masks an Mso afaik ... see ContinueHandler.addRec

	// OBJECT TYPES
	public static final byte otGroup = 0x0000;
	public static final byte otLine = 0x0001;
	public static final byte otRectangle = 0x0002;
	public static final byte otOval = 0x0003;
	public static final byte otArc = 0x0004;
	public static final byte otChart = 0x0005;
	public static final byte otText = 0x0006;
	public static final byte otButton = 0x0007;
	public static final byte otPicture = 0x0008;
	public static final byte otPolygon = 0x0009;
	public static final byte otCheckbox = 0x000B;
	public static final byte otRadiobutton = 0x000C;
	public static final byte otEditbox = 0x000D;
	public static final byte otLabel = 0x000E;
	public static final byte otDialogbox = 0x000F;
	public static final byte otSpincontrol = 0x0010;
	public static final byte otScrollbar = 0x0011;
	public static final byte otList = 0x0012;
	public static final byte otGroupbox = 0x0013;
	public static final byte otDropdownlist = 0x0014;
	public static final byte otNote = 0x0019;
	public static final byte otOfficeArtobject = 0x001E;

	protected static XLSRecord getPrototype()
	{
		Obj obj = new Obj();
		obj.setOpcode( OBJ );
		obj.setData( obj.PROTOTYPE_BYTES );
		obj.init();
		return obj;
	}

	public Obj()
	{
		this.ftCmo = new FtCmo();
	}

	/**
	 * create a new object of the desired type
	 * <br>NOTE: most object types require one or more sub-records to be complete
	 *
	 * @param obType object type - one of the object type constants
	 * @param obId   object id - must be unique in worksheet or chart substream
	 * @param grbit  object properties
	 */
	public Obj( int obType, int obId, int grbit )
	{
		this.setOpcode( OBJ );
		this.ftCmo = new FtCmo();
		this.ftCmo.setObjId( obId );
		this.ftCmo.setObjType( obType );
		this.ftCmo.setGrbit( grbit );
		this.setData( this.ftCmo.getRec() );
	}

	/**
	 * creates the basic subrecord structure for the desired object type
	 * <br>After the basic Object record is complete, there
	 * <br>NOTE: much of this is not complete
	 *
	 * @param int obType 	- one of the object type (ot) constants
	 * @param int obId		- object id - must be unique per worksheet/chart substream
	 * @return new object record
	 */
	protected static Obj getBasicObjRecord( int obType, int obId )
	{
		// TODO: figure out ftCmo grbit
		int grbit = 8449;        // appears to be standard grbit
		if( obType == 0x19 )// NOTE
		{
			grbit = 16401;
		}
		Obj obj = new Obj( obType, obId, grbit );
/*
obj.ftCmo.getRec()[14]= (byte)0x1C;	// D4
obj.ftCmo.getRec()[15]= (byte)0x17;	// 15
obj.ftCmo.getRec()[16]= (byte)0x92;
obj.ftCmo.getRec()[17]= (byte)0x02;
*/
		ArrayList subrecords = new ArrayList();

		// create the required subrecord(s), if possible, for the desired object type
		switch( obType )
		{
			case otPicture:
				// ftCf (07h)				Picture/Clipboard format  6 bytes; FtCf structure specifies the format of this picture object
				// ftPioGrbit (08h)			Picture option flags  6 bytes; FtPioGrbit structure specifies additional properties of this picture object
				break;
			case otRadiobutton:
				// ftRbo	(0Bh) Radio button; FtRbo structure represents a radio button
				// ftRboData (11h) FtRboData structure specifies additional properties of this radio button object.
			case otCheckbox:

				// ftCbls (0Ah): Check box link; FtCbls structure represents a check box or radio button
				// ftCblsData (12h) FtCblsData structure that specifies the properties of this check box or radio button object
				break;
			case otSpincontrol:
			case otScrollbar:
				// ftSbs	(0Ch) FtSbs structure specifies the properties of this spin control, scrollbar, list, or drop-down list object.
				break;
			case otList:
			case otDropdownlist:
				// ftSbs	(0Ch) FtSbs structure specifies the properties of this spin control, scrollbar, list, or drop-down list object.
				Obj.SubRecord ftSbs = obj.new SubRecord();        // iMin= 0, iMax= 100, dInc= 10, dPage= 0, fHoriz= 0, dxScroll= 16, fDraw
				ftSbs.parse( new byte[]{
						12, 0, 20, 0, 0, 0, 0, 0,	/* unused */
						0, 0, 		/* iVal */
						0, 0, 		/* iMin */
						100, 0,		/* iMax */
						1, 0,		/* dInc */
						10, 0, 		/* dPage */
						0, 0, 		/* fHoriz */
						16, 0, 		/* dxScroll= width in pixels of scrollbar */
						1, 0
				} );		/* fDraw bit 1 all else 0*/
				subrecords.add( ftSbs );
				// ftLbsData (13h) FtLbsData structure that specifies the properties of this listbox or drop-down object.
				Obj.FtLbsData ftlbsdata = obj.new FtLbsData();
				ftlbsdata.parse( new byte[]{
						19, 0, -18, 31, 	/* cbFContinued */
						0, 0,		/* cbFmla */
						0, 0,		/* cLines */
						1, 0, 		/* iSel */
						1, 3, 		/* grbit= 769  ==> lct=3*/
						0, 0, 		/* edit Id */
						2, 0, 8, 0, 87, 0		/* droprecprops: wStyle= 2(autofilter dropdown), cLines, dxMin - width in pixels*/
				} );
				subrecords.add( ftlbsdata );
				break;
			case otEditbox:
				// ftEdoData (10h) FtEdoData structure specifies the properties of this edit box object.
				break;
			case otGroupbox:
				// ftGboData (0Fh) FtGboData structure that specifies the properties of this group box object.
				break;
			case otButton:
				// ftMacro (04h)Fmla-style macro -- optional FtMacro structure that defines the action associated with the object
				break;
			case otNote:
				// ftNts	(0Dh) FtNts structure specifies the properties of this comment object
				Obj.SubRecord ftNts = obj.new SubRecord();
			/* 1.5-specific*/
				java.util.UUID uuid = java.util.UUID.randomUUID();  // global unique identifier for note
				byte[] guid = asByteArray( uuid );

    		/* 1.4 code
    		// fudge UUID creation!
    		long high= (long)Math.random();
    		long low= (long)Math.random();
    		byte[] guid= asByteArray(high, low);
    		*/
				byte[] ntsbytes = new byte[26];
				ntsbytes[0] = 13;    // ftNts
				ntsbytes[2] = 22;    // len
				System.arraycopy( guid, 0, ntsbytes, 4, 16 );
				// bytes 20-21= fSharedNote: 0 if not shared, 1 if shared
				// bytes 22-26= unused
				ntsbytes[22] = 16;    // ?? have no idea what this is but apparently is in comparison records
				ftNts.parse( ntsbytes );
				subrecords.add( ftNts );
				break;
			case otGroup:    // Group
				// ftGmo			06h			Group marker
				break;
		}
		// those object types that contain no subrecords:
		// otLine, otRectangle, otOval, otChart, otText, otOfficeArtObject

		// optional subrecords:
		// optional:  ftPictFmla (09h) Picture fmla-style macro; specifies the location of the data associated with this picture object
		// optional: ftSbsFmla		0Eh			Scroll bar fmla-style macro, ftCblsFmla		14h			Check box link fmla-style macro

		// retrieve all subrecord data and append to Object record
		byte[] newdata = obj.getData();
		for( int i = 0; i < subrecords.size(); i++ )
		{
			Obj.SubRecord sr = (Obj.SubRecord) subrecords.get( i );
			newdata = ByteTools.append( sr.getRec(), newdata );
		}
		// add ending ftEnd subrecord
		newdata = ByteTools.append( new byte[]{ 0, 0, 0, 0 }, newdata );

		obj.setData( newdata );
		return obj;
	}

	public String toString()
	{
		if( ftCmo != null )
		{
			return "OBJ type: " + ftCmo.getType() + " id: " + ftCmo.getObjId();
		}
		return "OBJ record not initialized";
	}

	/**
	 * Associate this record with a worksheet.
	 * \
	 */
	@Override
	public void setSheet( Sheet b )
	{
		this.worksheet = b;
		((Boundsheet) b).lastObjId = Math.max( ((Boundsheet) b).lastObjId, ftCmo.getObjId() );
	}

	/**
	 * Initialize a new OBJ Record
	 */
	@Override
	public void init()
	{
		super.init();

		byte[] data = this.getData();
		ftCmo = new FtCmo();    // Parse 1st FtCmo Record - contains object type, object id and grbit
		ftCmo.parse( data, 0 );

		try
		{    // parse rest of object sub-records
			int pos = 22;
			int cb;
			while( pos < (data.length - 2) )
			{
				int ft = ByteTools.readShort( data[pos], data[pos + 1] );
				cb = 2;
				switch( ft )
				{
					case 0:        //ftEnd
						break;
					case 0x4:    // ftMacro	- ignore for now
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0x5:    // ftButton: 0x5
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0x6:    // ftGmo- Group marker
						cb = 6;
						break;
					case 0x7:    // ftCf= Clipboard Format	- ignore for now
						cb = 6;
						break;
					case 0x8:    // ftPioGrbit=	some picture properties - ignore for now
						cb = 6;
						break;
					case 0x9:    // FtPictFmla = loc of data assoc with picture object - ignore for now
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xA:    // ftCbls= Check box link - ignore for now
						cb = 16;
						break;
					case 0xB:    // btRbo= Radio Button - ignore for now
						cb = 10;
						break;
					case 0xC:    // ftSbs (spin control, scroll bar, list or drop-down list)	// len=24
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xD:    // ftNts= Note structure	- ignore for now
						cb = 26;
						break;
					case 0xE: // ftSbsFmla
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xF:    // ftGboData
						cb = 10;
						break;
					case 0x10:    // ftEboData
						cb = 12;
						break;
					case 0x11:    // ftRboData
						cb = 8;
						break;
					case 0x12:    // ftCblsData - Check Box Data	- ignore for now
						cb = 12;
						break;
					case 0x13:    // ftLbsData= List Box Data
						FtLbsData ftl = new FtLbsData();
						ftl.parse( data, pos, (ftCmo.getObjType() == otDropdownlist) );
						cb = ftl.getCb();
						break;
					case 0x14: // ftCblsFmla
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					default:
						cb = 2;
						break;
				}
				if( ft == 0 )    // == ftEnd
				{
					break;
				}
				pos += cb;
			}
		}
		catch( ArrayIndexOutOfBoundsException ae )
		{
			if( false )
			{
				Logger.logInfo( "Obj encountered in records" );
			}
		}
	}

	/**
	 * update the OBJ record's data
	 * NOTE: For now the only thing we know for sure is 1st sub-record is
	 * ftcmo, with updatable fields objType, objId and grbit
	 * so update those
	 */
	private void updateRecord()
	{
		byte[] retData = new byte[6];
		byte[] b = ByteTools.shortToLEBytes( ftCmo.getObjType() );
		retData[0] = b[0];
		retData[1] = b[1];
		b = ByteTools.shortToLEBytes( ftCmo.getObjId() );
		retData[2] = b[0];
		retData[3] = b[1];
		b = ByteTools.shortToLEBytes( ftCmo.getGrbit() );
		retData[4] = b[0];
		retData[5] = b[1];
		byte[] updated = this.getData();
		System.arraycopy( retData, 0, updated, 4, retData.length );    // update a portion of the initial ftcmo rec starting at byte 4
		this.setData( updated );
	}

	// 20100430 KSC: actually, the only Continue that follows Obj records are masked Mso's -- see ContinueHandler.addRec
	/** ------------------------------------------------------------
	 Continue text, formattingruns;
	 int 	state     					= 0;
	 *
	 * @return Returns the formattingruns.
	public Continue getFormattingruns() {
	return formattingruns;
	}

	/** ------------------------------------------------------------
	 *
	 * @param formattingruns The formattingruns to set.
	public void setFormattingruns(Continue formattingruns) {
	this.formattingruns = formattingruns;
	}

	/** ------------------------------------------------------------
	 *
	 * @return Returns the state.
	public int getState() {
	return state;
	}

	/** ------------------------------------------------------------
	 *
	 * @param state The state to set.
	public void setState(int state) {
	this.state = state;
	}
	 */

	/**
	 * Return the chart for this obj, if exists.  Null if not
	 * ------------------------------------------------------------
	 *
	 * @return
	 */
	public Chart getChart()
	{
		return myChart;
	}

	public void setChart( Chart c )
	{
		myChart = c;
	}

	/**
	 * ------------------------------------------------------------
	 *
	 * @return Returns the objType.
	 */
	public int getObjType()
	{
		return ftCmo.getObjType();
	}

	/**
	 * ------------------------------------------------------------
	 *
	 * @return Returns the objId.
	 */
	public int getObjId()
	{
		return ftCmo.getObjId();
	}

	/**
	 * ------------------------------------------------------------
	 *
	 * @param objId The objId to set.
	 */
	public void setObjId( int objId )
	{
		ftCmo.setObjId( (short) objId );
		updateRecord();
	}

	/**
	 * defines a generic subrecord of the Object record
	 */

	private class SubRecord implements Serializable
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = -9040304321493078260L;
		short ft;
		short cb;
		byte[] rec;

		public byte[] getRec()
		{
			return rec;
		}

		public void parse( byte[] rec )
		{
			ft = ByteTools.readShort( rec[0], rec[1] );
			cb = ByteTools.readShort( rec[2], rec[3] );
			this.rec = rec;
		}

		public int getFt()
		{
			return ft;
		}

		public int getCb()
		{
			return cb;
		}
	}

	/**
	 * ftCmo			15h			Common object data
	 * offset	Name	Size	Contents
	 * 0		ft		2		= ftCmo (0x15)
	 * 2		cb		2		Length of ftCmo Data
	 * 4		ot		2		Object Type
	 * 6		id		2		Object ID Number
	 * 8		grbit	2
	 * <p/>
	 * grbit:
	 * Bits		Mask		Name		Contents
	 * <p/>
	 * 0			0001h		fLocked		= 1 if the object is locked when the sheet is protected
	 * 3  1		000Eh		(Reserved)	Reserved; must be 0 (zero)
	 * 4			0010h		fPrint		= 1 if the object is printable
	 * 12  5		1FE0h		(Reserved)	Reserved; must be 0 (zero)
	 * 13			2000h		fAutoFill	= 1 if the object uses automatic fill style
	 * 14			4000h		fAutoLine	= 1 if the object uses automatic line style
	 * 15			8000h		(Reserved)	Reserved; must be 0 (zero)
	 * <p/>
	 * ot, object type:
	 * 00	= Group
	 * 01	= Line
	 * 02	= Rectangle
	 * 03	= Oval
	 * 04	= Arc
	 * 05	= Chart
	 * 06	= Text
	 * 07	= Button
	 * 08	= Picture
	 * 09	= Polygon
	 * 0B	= CheckBox
	 * 0C	= Option Button
	 * 0D  = Edit Box
	 * OE	= Label
	 * 0F  = Dialog Box
	 * 10	= Spin Control
	 * 11  = Scrollbar
	 * 12 	= List
	 * 13	= Group Box
	 * 14	= Dropdown List Box
	 * 19	= Note
	 * IE	= MSO Drawing
	 * 14		(reserved) 12	must be 0
	 */
	private class FtCmo extends SubRecord
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = 9129235555230900833L;
		short objType;
		short objId;
		short grbit;

		FtCmo()
		{
			ft = 0x15;
			cb = 22;
			rec = new byte[22];
			rec[0] = 21;        // ft
			rec[2] = 18;        // len
			// unknown and undocumented: apparently bytes 16 & 17 are same for all Obj recs in sheet - always true?
			// bytes 14 & 15 change per Obj rec
		}

		public void parse( byte[] rec, int pos )
		{
			if( rec.length < 10 )
			{
				return;    // error! should be at least 22 bytes
			}
//	    	ft= ByteTools.readShort(this.getByteAt(0), this.getByteAt(1));			// 0x15
//	    	cb= ByteTools.readShort(this.getByteAt(2), this.getByteAt(3));    		// len of data in ftcmo == 22 
			objType = ByteTools.readShort( rec[4], rec[5] );        // Object type, see above
			objId = ByteTools.readShort( rec[6], rec[7] );            // Object id
			grbit = ByteTools.readShort( rec[8], rec[9] );            // flags, see above
			// these flags appear important but no documentation on them
			//unknown1= ByteTools.readShort(rec[14], rec[15]);
			//unknown2= ByteTools.readShort(rec[16], rec[17]);
		}

		public short getObjType()
		{
			return objType;
		}

		public short getObjId()
		{
			return objId;
		}

		public short getGrbit()
		{
			return grbit;
		}

		public void setObjId( int objId )
		{
			this.objId = (short) objId;
			byte[] b = ByteTools.shortToLEBytes( this.objId );
			rec[6] = b[0];
			rec[7] = b[1];
		}

		public void setObjType( int objType )
		{
			this.objType = (short) objType;
			byte[] b = ByteTools.shortToLEBytes( this.objType );
			rec[4] = b[0];
			rec[5] = b[1];
		}

		public void setGrbit( int grbit )
		{
			this.grbit = (short) grbit;
			byte[] b = ByteTools.shortToLEBytes( this.grbit );
			rec[8] = b[0];
			rec[9] = b[1];
		}

		//user-friendly object type string
		public String getType()
		{
			switch( objType )
			{
				case 0x00:
					return "Group";
				case 0x01:
					return "Line";
				case 0x02:
					return "Rectangle";
				case 0x03:
					return "Oval";
				case 0x04:
					return "Arc";
				case 0x05:
					return "Chart";
				case 0x06:
					return "Text";
				case 0x07:
					return "Button";
				case 0x08:
					return "Picture";
				case 0x09:
					return "Polygon";
				case 0x0B:
					return "CheckBox";
				case 0x0C:
					return "Option Button";
				case 0x0D:
					return "Edit Box";
				case 0xE:
					return "Label";
				case 0x0F:
					return "Dialog Box";
				case 0x10:
					return "Spin Control";
				case 0x11:
					return "Scrollbar";
				case 0x12:
					return "List";
				case 0x13:
					return "Group Box";
				case 0x14:
					return "Dropdown List Box";
				case 0x19:
					return "Note";
				case 0x1E:
					return "MSO Drawing";
			}
			return "UNKNOWN TYPE " + objType;
		}
	}

	/**
	 * ftLbsData: List Box Data (var):
	 * cbFContinued (2 bytes)
	 * An unsigned integer that indirectly specifies whether some of the data in this structure appear in a subsequent Continue record.
	 * If cbFContinued is 0x0000, all of the fields in this structure except ft and cbFContinued MUST NOT exist.
	 * If this entire structure is contained within the same record, then cbFContinued MUST be greater than or equal to the size, in bytes,
	 * of this structure, not including the four bytes for the ft and cbFContinued fields.
	 * <p/>
	 * If part of this structure is in one or more subsequent Continue records, then the cbFContinued field MUST hold the value
	 * calculated according to the following formula:
	 * cbFContinued = size of the fields of this structure in the current record - 1.
	 * fmla (variable): An ObjFmla that specifies the range of cell values that are the items in this list.
	 * bFmla (2 bytes): An unsigned integer that specifies the number of bytes in this ObjFmla, not counting the two bytes of the cbFmla field itself. This number MUST be even.
	 * fmla (variable): An optional ObjectParsedFormula that specifies the formula. This field MUST exist if and only if cbFmla is greater than 0x0000.
	 * embedInfo (variable): An optional PictFmlaEmbedInfo. This field MUST exist if and only if the structure containing this ObjFmla is an FtPictFmla, the fmla field exists, and the fmla.rgce field starts with a PtgTbl.
	 * padding (variable): An array of bytes whose size is given by:cbFmla minus size of fmla minus size of embedInfo.
	 * cLines (2 bytes): An unsigned integer that specifies the number of items in the list. MUST be less than or equal to 0x7FFF.
	 * iSel (2 bytes): An unsigned integer that specifies the one-based index of the first selected item in this list. A value of 0x0000 specifies there is no currently selected item. MUST be less than or equal to cLines.
	 * A - fUseCB (1 bit): A bit that specifies whether the lct field MUST be ignored. MUST be a value from the following table:
	 * 0	The lct field MUST be ignored.
	 * 1	The lct field MUST NOT be ignored.
	 * B - fValidPlex (1 bit): A bit that specifies whether the rgLines field exists.
	 * C - fValidIds (1 bit): A bit that specifies whether the idEdit field MUST be ignored. MUST be a value from the following table:
	 * 0	 The idEdit field MUST be ignored.
	 * 1	The idEdit field MUST NOT be ignored.
	 * D - fNo3d (1 bit): A bit that specifies whether this control is displayed without 3-dimensional effects. MUST be a value from the following table:
	 * 0	The control is displayed with 3-dimentional effects.
	 * 1	The control is not displayed with 3-dimentional effects.
	 * E - wListSelType (2 bits): An unsigned integer that specifies the type of selection behavior this list control is expected to support. MUST be a value from the following table:
	 * 0	The list control is only allowed to have one selected item.
	 * 1	The list control is allowed to have multiple items selected by clicking on each item.
	 * 2	The list control is allowed to have multiple items selected by holding the CTRL key and clicking on each item.
	 * F - unused (1 bit): Undefined and MUST be ignored.
	 * G - reserved (1 bit): MUST be zero, and MUST be ignored.
	 * lct (8 bits): An unsigned integer that specifies the behavior class of this list. MUST be ignored if the fUseCB field is 0.
	 * Otherwise, MUST be a value from the following table:
	 * Expected behavior of the control
	 * 0x00	Regular sheet dropdown control (like a list box object).
	 * 0x01	PivotTable page field dropdown.
	 * 0x03	AutoFilter dropdown. The lct field MUST NOT have this value unless this object is in a worksheet or macro sheet.
	 * 0x05	AutoComplete dropdown.
	 * 0x06	Data validation list dropdown. The lct field MUST NOT have this value unless this object is in a worksheet or macro sheet.
	 * 0x07	PivotTable row or column field dropdown.
	 * 0x09	Dropdown for the Total Row of a table.
	 * <p/>
	 * idEdit (2 bytes): An ObjId that specifies the edit box associated with this list. A value of idEdit.id equal to 0x0000 or a value of fValidIds equal to 0 specifies that there is no edit box associated with this list.
	 * dropData (variable): An optional LbsDropData that specifies properties for this dropdown control. This field MUST exist if and only if the containing Obj’s cmo.ot is equal to 0x0014.
	 * rgLines (variable): An optional array of XLUnicodeString. Each string in this array specifies an item in the list.
	 * This array MUST exist if and only if the fValidPlex field is equal to 1.
	 * The number of elements in this array, if it exists, MUST be cLines.
	 * The cch field of each string in this array MUST be less than or equal to 0x00FF.
	 * If this array does not fit in the owning Obj record, Continue records are used. Each string in this array MUST be entirely contained within the same record.
	 * bsels (variable): An optional array of one-byte Booleans that specifies which items in the list are part of a multiple selection. This array MUST exist if and only if the wListType field is not equal to 0. The number of elements in this array, if it exists, MUST be cLines. The nth byte in this array specifies whether the nth list item is part of the multiple selection. The value of each element MUST be taken from the following table:
	 * 0x00	List item is not part of the multiple selection.
	 * 0x01	List item is part of the multiple selection.
	 * If this array does not fit in the current record, or would come within eight bytes of the end of the maximum allowable size of that record, Continue records are used.
	 */
	private class FtLbsData extends SubRecord
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = 4965884944498060366L;

		FtLbsData()
		{
			ft = 0x13;
		}

		public void parse( byte[] rec, int pos, boolean isListBox )
		{
			int origPos = pos;
			pos += 2;
			int cbFContinued = ByteTools.readShort( rec[pos], rec[pos + 1] );
			if( cbFContinued != 0 )
			{
				pos += 2;
				int cbFmla = ByteTools.readShort( rec[pos], rec[pos + 1] );
				// read cbFmla bytes    			    				
				pos += cbFmla + 2;
				int cLines = ByteTools.readShort( rec[pos], rec[pos + 1] );
				pos += 2;
				int isel = ByteTools.readShort( rec[pos], rec[pos + 1] );
				pos += 2;
				short grbit = ByteTools.readShort( rec[pos], rec[pos + 1] );
				// fUseCB= 		bit 1
				// fValidPlex 	bit 2
				// fValidIds	bit 3
				// fNo3d		bit 4
				// wListSelType bits 5, 6
				// bits 7, 8 - ignore
				// lct			8 bits
				if( (grbit & 0x1) == 0x1 )
				{ // fUseCB means lct is valid
					short lct = (short) ((grbit) >> 8);
					// 0= regular drop-down
					// 1= pivot table page field drop-down
					// 3= AutoFilter drop-down
					// 5= AutoComplete drop-down
					// 6= rec Validation drop-down
					// 7= pivot table drop-down
					// 9= Table total row drop-down    			
				}
				pos += 2;
				int iEditId = ByteTools.readShort( rec[pos], rec[pos + 1] );
				if( isListBox )
				{
					// droprec - var
					pos += 2;
					int droprecProps = ByteTools.readShort( rec[pos], rec[pos + 1] );
					// 0= combo, 1= combo edit dropdown control, 2= simple dropdown (just the button)
					pos += 2;
					int cLine = ByteTools.readShort( rec[pos], rec[pos + 1] );
					pos += 2;
					int dxMin = ByteTools.readShort( rec[pos], rec[pos + 1] );
					// str= unicode string 
				}
				if( (grbit & 0x2) == 0x2 )
				{    // fValidPlex
					// rgLines - var
				}
				// bsels - var
				cb = (short) ((pos + 2) - origPos);
			}
			else        //	the rest are in continues record
			{
				cb = 4;
			}
		}

	}

	/**
	 * returns debug output for this Obj record
	 * <br>FOR INTERNAL USE ONLY
	 *
	 * @return
	 */
	public String debugOutput()
	{
		byte[] data = this.getData();
		StringBuffer log = new StringBuffer();
		FtCmo ftCmo = new FtCmo();    // Parse 1st FtCmo Record - contains object type, object id and grbit
		ftCmo.parse( data, 0 );
		log.append( this.toString() );
		log.append( "\r\n\t" );
		log.append( "FtCmo" );
		log.append( "\t" + ByteTools.getByteDump( data, 0, 22 ) );
		try
		{    // parse rest of object sub-records
			int pos = 22;
			int cb;
			while( pos < (data.length - 2) )
			{
				log.append( "\r\n\t" );
				int ft = ByteTools.readShort( data[pos], data[pos + 1] );
				cb = 2;
				switch( ft )
				{
					case 0:        //ftEnd
						log.append( "ftEnd" );
						break;
					case 0x4:    // ftMacro	- ignore for now
						log.append( "ftMacro" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0x5:    // ftButton: 0x5
						log.append( "ftButton" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0x6:    // ftGmo- Group marker
						log.append( "ftGmo" );
						cb = 6;
						break;
					case 0x7:    // ftCf= Clipboard Format	- ignore for now
						log.append( "ftCf" );
						cb = 6;
						break;
					case 0x8:    // ftPioGrbit=	some picture properties - ignore for now
						log.append( "ftPioGrbit" );
						cb = 6;
						break;
					case 0x9:    // FtPictFmla = loc of data assoc with picture object - ignore for now
						log.append( "ftPictFmla" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xA:    // ftCbls= Check box link - ignore for now
						log.append( "ftCbls" );
						cb = 16;
						break;
					case 0xB:    // btRbo= Radio Button - ignore for now
						log.append( "ftRbo" );
						cb = 10;
						break;
					case 0xC:    // ftSbs (spin control, scroll bar, list or drop-down list)	// len=24
						log.append( "ftSbs" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xD:    // ftNts= Note structure	- ignore for now
						log.append( "ftNts" );
						cb = 26;
						break;
					case 0xE: // ftSbsFmla
						log.append( "ftSbsFmla" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					case 0xF:    // ftGboData
						log.append( "ftGboData" );
						cb = 10;
						break;
					case 0x10:    // ftEboData
						log.append( "ftEboData" );
						cb = 12;
						break;
					case 0x11:    // ftRboData
						log.append( "ftRboData" );
						cb = 8;
						break;
					case 0x12:    // ftCblsData - Check Box Data	- ignore for now
						log.append( "ftCblsData" );
						cb = 12;
						break;
					case 0x13:    // ftLbsData= List Box Data
						log.append( "ftLbsData" );
						FtLbsData ftl = new FtLbsData();
						ftl.parse( data, pos, (ftCmo.getObjType() == otDropdownlist) );
						cb = ftl.getCb();
						break;
					case 0x14: // ftCblsFmla
						log.append( "ftCblsFmla" );
						cb = 4 + ByteTools.readShort( data[pos + 2], data[pos + 3] );
						break;
					default:
						log.append( "Unknown" );
						cb = 2;
						break;
				}
				if( ft == 0 )    // == ftEnd
				{
					break;
				}
				log.append( ByteTools.getByteDump( data, pos, cb ) );
				pos += cb;
			}
		}
		catch( ArrayIndexOutOfBoundsException ae )
		{
			if( false )
			{
				Logger.logInfo( "Obj encountered in records" );
			}
		}
		return log.toString();
	}

	/* 1.5 specific code*/
	private static java.util.UUID toUUID( byte[] byteArray )
	{
		long msb = 0;
		long lsb = 0;
		for( int i = 0; i < 8; i++ )
		{
			msb = (msb << 8) | (byteArray[i] & 0xff);
		}
		for( int i = 8; i < 16; i++ )
		{
			lsb = (lsb << 8) | (byteArray[i] & 0xff);
		}
		java.util.UUID uuid = new java.util.UUID( msb, lsb );
		return uuid;
	}

	private static byte[] asByteArray( java.util.UUID uuid )
	{
		long msb = uuid.getMostSignificantBits();
		long lsb = uuid.getLeastSignificantBits();
		byte[] buffer = new byte[16];

		for( int i = 0; i < 8; i++ )
		{
			buffer[i] = (byte) (msb >>> (8 * (7 - i)));
		}
		for( int i = 8; i < 16; i++ )
		{
			buffer[i] = (byte) (lsb >>> (8 * (7 - i)));
		}

		return buffer;
	}
/* 1.4 specific
	private static byte[] asByteArray(long m, long l) { 	 
	    byte[] buffer = new byte[16];

	    for (int i = 0; i < 8; i++) {
	            buffer[i] = (byte) (m >>> 8 * (7 - i));
	    }
	    for (int i = 8; i < 16; i++) {
	            buffer[i] = (byte) (l >>> 8 * (7 - i));
	    }

	    return buffer;
	}
*/
}


