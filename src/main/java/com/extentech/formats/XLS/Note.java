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

import com.extentech.formats.OOXML.Text;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;

/**
 * <b>Note: A cell annotation (1CDh)</b><br>
 * <p/>
 * Note records contain the row/col position of the annotation, plus an ordinal id that links it to the
 * set of records that define it:  Mso/obj/mso/Txo/Continue/Continue.  These record sets appear one after
 * another after the DIMENSIONS record and before WINDOW2.  After the associated records, all Note records
 * appear, in order.
 * <p/>
 * <p/>
 * Kaia's notes on notes:
 * <p/>
 * The Note record is the last of a set of records necessary to define each note:
 * Mso - shapeType= msosptTextBox.  See Msodrawing.createTextBoxStyle for a list of the necessary sub-records
 * Obj containing an ftNts sub-record
 * Mso - odd Mso cotnaining only 8 bytes - see Msodrawing.getAttachedTextPrototype
 * Txo - contains the size of the text + options such as text rotation + size of the formatting runs
 * Continue - 0 + text
 * Continue - formatting runs for text
 * <p/>
 * For each note, the above records repeat, then each Note record appears in order after all of the above recs
 * Notes appear to have an ordinal id that is necessary for proper display
 * It's hidden/shown state appears to be defined by byte 5 but there is also other necessary info that I don't know yet
 * <p/>
 * The tricky part of Notes is the Mso msosptTextBox plus the Mso msosptClientTextBox; they don't follow the
 * normal rules we are used to - the SPCONTAINERLENGTH of the 1st Mso MUST BE inc by 8, the length of the
 * second Mso ... but when multiple Notes are on the same sheet, that logic doesn't seem to be followed 100%
 * plus the number of shapes of the mso header is NOT incremented by the secondary Mso ...
 * <p/>
 * Also, the Mso's don't appear to follow the normal rules for the drawing Id - seem to be repeated ...
 * <p/>
 * Lastly, the 1st Mso contains an OPT sub-record msofbtlTxid; this contains the host-defined id for the note;
 * I'm not sure what to set for this so I put a random int.
 * <p/>
 * More info:
 * <p/>
 * Some templates define notes differently; instead of the above set of records, it goes
 * Obj (ftNts)
 * Continue [00, 00, (byte)0x0D, (byte)0xF0, 00, 00, 00, 00]  --> masks the 2nd mso
 * Txo
 * Continue
 * Continue
 * Continue --> masks the 1st mso
 * <p/>
 * Also:
 * formatting runs: This is almost finished but - of course - doc on formatting runs format doesn't appear to match reality ((:)
 * Also:
 * There needs to be a good, easy interface for setting formatting runs for a text string
 * <p/>
 * <p/>
 * offset  name        size    contents
 * ---
 * 0 		2 Index to row
 * 2 		2 Index to column
 * 4		2 grbit controls hidden state - 0=hidden, 2=shown
 * 6		2 ordinal id MUST MATCH Obj record id
 * 8		1 lenght of author string or /2 if encoding=1
 * 9       1 encoding 0= non-unicode? 1= unicode?
 * 10		(var) 0 + author bytes
 * (author encoding)
 * </p></pre>
 *
 * @see Note
 */
public final class Note extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 *
	 *
	 */
	private short id;
	boolean hidden = true;    // default
	private String author;
	private byte auth_encoding;
	private MSODrawing mso = null;
	private static final long serialVersionUID = -1571461658267879478L;
	// pointer to associated Txo, stores actual text + formatting runs
	Txo txo = null;

	@Override
	public void init()
	{
		super.init();
		rw = ByteTools.readShort( this.getData()[0], this.getData()[1] );
		col = ByteTools.readShort( this.getData()[2], this.getData()[3] );
		hidden = (this.getData()[4] != (byte) 2);    // not entirely sure of this
		// bytes 5 unknnown
		id = ByteTools.readShort( this.getData()[6], this.getData()[7] );
		// rest are fairy known :)
		short authorlen = this.getData()[8];
		auth_encoding = this.getData()[9];
		byte[] authorbytes = new byte[authorlen];
		System.arraycopy( this.getData(), 11, authorbytes, 0, authorlen );
		if( auth_encoding == 0 )
		{
			author = new String( authorbytes );
		}
		else
		{
			try
			{
				author = new String( authorbytes, WorkBookFactory.UNICODEENCODING );
			}
			catch( Exception e )
			{
				;
			}
		}
	}

	@Override
	public void setSheet( Sheet bs )
	{
		super.setSheet( bs );
		txo = getAssociatedTxo();
	}

	/**
	 * Is this a useful to string?  possibly just the note itself, at a minimum include it
	 *
	 * @see com.extentech.formats.XLS.XLSRecord#toString()
	 */
	public String toString()
	{
		String s = "";
		if( txo != null )
		{
			s = txo.getStringVal();
		}
		if( author != null )
		{
			s += " author:" + author;
		}
		return "NOTE at [" + this.getCellAddressWithSheet() + "]: " + s;
	}

	public static XLSRecord getPrototype( String author )
	{
		Note n = new Note();
		n.setOpcode( NOTE );
		/*
         *  $row, $col, $visible, $obj_id,
                                 $num_chars, $author_enc
    		$length     = length($data) + length($author);
    		my $header  = pack("vv", $record, $length);
         */
		byte[] authbytes = author.getBytes();
		byte[] data = new byte[n.PROTOTYPE_BYTES.length + authbytes.length + 1];
		data[8] = (byte) author.length();
		// encoding= 0
		System.arraycopy( authbytes, 0, data, 11, authbytes.length );
		n.setData( data );
		n.init();
		return n;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 	/* row */
			0, 0, 	/* col */
			0,		/* hidden state */
			0,		/* unknown */
			1, 0,	/* ordinal id */
			4,	    /* length of author */
			0,     /* string encoding*/
			//		/* 0-padded author string*/
	};

	/**
	 * set the row and column that this note is attached to
	 *
	 * @param row
	 * @param col
	 */
	public void setRowCol( int row, int col )
	{
		this.rw = (short) row;
		this.col = (short) col;
		byte[] b = ByteTools.shortToLEBytes( (short) this.rw );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		b = ByteTools.shortToLEBytes( this.col );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
	}

	/**
	 * return the ordinal id of this note
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		this.id = (short) id;
		byte[] b = ByteTools.shortToLEBytes( this.id );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
	}

	/**
	 * retrieve the ordinal id of this note
	 *
	 * @return
	 */
	public int getId()
	{
		return this.id;
	}

	/**
	 * returns true if this note is hidden (default state)
	 *
	 * @return
	 */
	public boolean getHidden()
	{
		return this.hidden;
	}

	/**
	 * show or hide this note
	 * <br>NOTE: this is still experimental
	 *
	 * @param hidden
	 */
	public void setHidden( boolean hidden )
	{
		this.hidden = hidden;
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		if( this.hidden )
		{    // hide
			this.getData()[4] = 0;
			// ALSO MUST SET the associated Mso opt subrec to actually hide the note textbox ...
			mso.setOPTSubRecord( MSODrawingConstants.msooptGroupShapeProperties, false, false, 131074, null );
		}
		else
		{            //show
			this.getData()[4] = 2;
			// ALSO MUST SET the associated Mso opt subrec to actually show the note textbox ...
			mso.setOPTSubRecord( MSODrawingConstants.msooptGroupShapeProperties, false, false, 131072, null );
		}
	}

	/**
	 * returns the Text of this Note or Comment
	 *
	 * @return String
	 */
	public String getText()
	{
		if( txo != null )
		{
			return txo.getStringVal();
		}
		return null;
	}

	/**
	 * sets the Text of this Note or Comment
	 *
	 * @param txt
	 */
	public void setText( String txt )
	{
		if( txo != null )
		{
			try
			{
				txo.setStringVal( txt );
			}
			catch( IllegalArgumentException e )
			{
				Logger.logErr( e.toString() );
			}
		}
	}

	/**
	 * /** set the text of this Note (Comment) as a unicode string,
	 * with formatting information
	 *
	 * @param txt
	 */
	public void setText( Unicodestring txt )
	{
		if( txo != null )
		{
			try
			{
				txo.setStringVal( txt );
			}
			catch( IllegalArgumentException e )
			{
				Logger.logErr( e.toString() );
			}
		}

	}

	/**
	 * return the author of this note, if set
	 *
	 * @return
	 */
	public String getAuthor()
	{
		return author;
	}

	/**
	 * sets the author of this note
	 *
	 * @param author
	 */
	public void setAuthor( String author )
	{
		this.author = author;
		byte[] authbytes = this.author.getBytes();
		byte[] oldData = this.getData();
		byte[] newData = new byte[authbytes.length + 11];
		System.arraycopy( oldData, 0, newData, 0, 8 );
		newData[8] = (byte) author.length();
		// encoding= 0
		System.arraycopy( authbytes, 0, newData, 11, authbytes.length );
		this.setData( newData );
		this.init();
	}

	/**
	 * returns the Txo associated with this Note
	 *
	 * @return
	 */
	private Txo getAssociatedTxo()
	{
		Boundsheet bs = this.getSheet();
		int idx = -1;
		if( bs != null )
		{// shouldn't!
			idx = bs.getIndexOf( OBJ );
			if( idx == -1 )
			{
				return null;    // should't!
			}
			while( idx < bs.getSheetRecs().size() )
			{
				if( ((BiffRec) bs.getSheetRecs().get( idx )).getOpcode() == OBJ )
				{
					Obj o = ((Obj) bs.getSheetRecs().get( idx ));
					// if it's of type Note + has the same id, this is it
					if( (o.getObjType() == 0x19) && (o.getObjId() == this.id) )
					{ // got it!
						// now find the next TXO
						idx++;
						while( (idx < bs.getSheetRecs().size()) && ((((BiffRec) bs.getSheetRecs().get( idx ))).getOpcode() != TXO) )
						{
							idx++;
						}
						break;
					}
				}
				idx++;
			}
		}
		if( idx < bs.getSheetRecs().size() )
		{
			return (Txo) bs.getSheetRecs().get( idx );
		}
		return null;
	}

	/**
	 * returns the bounds (size and position) of the Text Box for this Note
	 * <br>bounds are relative and based upon rows, columns and offsets within
	 * <br>bounds are as follows:
	 * <br>bounds[0]= column # of top left position (0-based) of the shape
	 * <br>bounds[1]= x offset within the top-left column (0-1023)
	 * <br>bounds[2]= row # for top left corner
	 * <br>bounds[3]= y offset within the top-left corner	(0-1023)
	 * <br>bounds[4]= column # of the bottom right corner of the shape
	 * <br>bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
	 * <br>bounds[6]= row # for bottom-right corner of the shape
	 * <br>bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
	 *
	 * @return
	 */
	public short[] getTextBoxBounds()
	{
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		return mso.getBounds();
	}

	/**
	 * sets the bounds (size and position) of the Text Box for this Note
	 * <br>bounds are relative and based upon rows, columns and offsets within
	 * <br>bounds are as follows:
	 * <br>bounds[0]= column # of top left position (0-based) of the shape
	 * <br>bounds[1]= x offset within the top-left column (0-1023)
	 * <br>bounds[2]= row # for top left corner
	 * <br>bounds[3]= y offset within the top-left corner	(0-1023)
	 * <br>bounds[4]= column # of the bottom right corner of the shape
	 * <br>bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
	 * <br>bounds[6]= row # for bottom-right corner of the shape
	 * <br>bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
	 *
	 * @param bounds
	 */
	public void setTextBoxBounds( short[] bounds )
	{
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		mso.setBounds( bounds );
	}

	/**
	 * sets the text box width for the note in pixels
	 *
	 * @param width
	 * @return short[] end column and end column offset
	 */
	public void setTextBoxWidth( short width )
	{
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		mso.setWidth( (short) Math.round( width / 6.4 ) );    // convert pixels to points
	}

	/**
	 * sets the text box width for the note in pixels
	 *
	 * @param width
	 * @return short[] end column and end column offset
	 */
	public void setTextBoxHeight( short height )
	{
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		mso.setHeight( (short) Math.round( height ) );
	}

	/**
	 * Store formatting runs for this Note (Comment)
	 * <br>Formatting Runs are an Internal Structure and are not relavent to the end user
	 *
	 * @param formattingRuns
	 */
	public void setFormattingRuns( ArrayList formattingRuns )
	{
		if( txo != null )
		{
			txo.setFormattingRuns( formattingRuns );
		}
	}

	/**
	 * Method to retrieve formatting runs for this Note (Comment)
	 * <br>Formatting Runs are an Internal Structure and are not relavent to the end user
	 *
	 * @param formattingRuns
	 */
	public ArrayList getFormattingRuns()
	{
		if( txo != null )
		{
			txo.getFormattingRuns();
		}
		return null;
	}

	/**
	 * get the id of drawing object which defines the text box for this Note
	 * <br>For Internal Use Only
	 *
	 * @return
	 */
	public int getSPID()
	{
		if( mso == null )
		{
			mso = this.getAssociatedMso();
		}
		return mso.getSPID();
	}

	private MSODrawing getAssociatedMso()
	{
		Boundsheet bs = this.getSheet();
		int idx = -1;
		if( bs != null )
		{// shouldn't!
			idx = bs.getIndexOf( OBJ );
			if( idx == -1 )
			{
				return null;    // should't!
			}
			while( idx < bs.getSheetRecs().size() )
			{
				if( ((BiffRec) bs.getSheetRecs().get( idx )).getOpcode() == OBJ )
				{
					Obj o = ((Obj) bs.getSheetRecs().get( idx ));
					// if it's of type Note + has the same id, this is it
					if( (o.getObjType() == 0x19) && (o.getObjId() == this.id) )
					{ // got it!
						// first check if this is one of the odd configurations
						int opcodeprev = ((BiffRec) bs.getSheetRecs().get( idx - 1 )).getOpcode();
						int opcodenext = ((BiffRec) bs.getSheetRecs().get( idx + 1 )).getOpcode();
						if( opcodeprev == MSODRAWING )    // normal case: mso, obj (note), mso, txo, continue, continue
						{
							return (MSODrawing) bs.getSheetRecs().get( idx - 1 );
						}
						else if( opcodenext == CONTINUE )
						{// continue masking mso - AFTER object record
							if( (idx + 5) < bs.getSheetRecs().size() )
							{
								return (((Continue) bs.getSheetRecs()
								                      .get( idx + 5 ))).maskedMso;    // the 1st mso is actually at position 5: Obj (note)/Continue/Txo/Continue/Continue/Continue
							}
						}
					}
				}
				idx++;
			}
		}
		return null;
	}

	/**
	 * returns the OOXML representation of this Note
	 *
	 * @return
	 */
	public String getOOXML()
	{
		if( txo == null )
		{
			return "";
		}
		Text t = new Text( Sst.createUnicodeString( this.getText(), txo.getFormattingRuns(), Sst.STRING_ENCODING_UNICODE ) );
		return t.getOOXML( this.getWorkBook() );
	}
}
