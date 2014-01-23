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

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.MSODrawingConstants;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;
//0xf00b

/**
 * shape properties table
 */
public class MsofbtOPT extends EscherRecord
{
	private static final Logger log = LoggerFactory.getLogger( MsofbtOPT.class );
	private static final long serialVersionUID = 465530579513265882L;
	byte[] recordData = new byte[0];
	boolean bBackground, bActive, bPrint;
	int imageIndex = -1;
	java.awt.Color fillColor = null;
	int fillType = 0;
	String imageName = "", shapeName = "", alternateText = "";
	int[] lineprops = null;            // Line properties -- weight, color, style ...
	static final int LINEPROPTS_STYLE = 0;
	static final int LINEPROPTS_WEIGHT = 1;
	static final int LINEPROPTS_COLOR = 2;
	boolean hasTextId = false;    // true if OPT contains msofbtlTxid - necessary to calculate container lengths correctly - see MSODrawing.updateRecord
	LinkedHashMap props = new LinkedHashMap();        // parsed property table note: properties are ordered via property set;

	public MsofbtOPT( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	public void setImageIndex( int value )
	{
		imageIndex = value;
		if( imageIndex > -1 )
		{
			setProperty( MSODrawingConstants.msooptpib, true, false, value, null );
		}
		else // remove property
		{
			props.remove( Integer.valueOf( MSODrawingConstants.msooptpib ) );
		}

	}

	/**
	 * returns true if this OPT subrecord contains an msofbtlTxid entry-
	 * necessary to calculate container lengths correctly - see MSODrawing.updateRecord
	 *
	 * @return
	 */
	public boolean hasTextId()
	{
		return hasTextId;
	}

	/**
	 * set the image name for this shape
	 * msooptpibName
	 *
	 * @param name
	 */
	public void setImageName( String name )
	{
		try
		{
			imageName = name;
			if( (imageName == null) || imageName.equals( "" ) )
			{
				// remove property
				props.remove( Integer.valueOf( MSODrawingConstants.msooptpibName ) );
				return;
			}
			byte[] imageNameBytes = name.getBytes( XLSConstants.UNICODEENCODING );
			byte[] newbytes = new byte[imageNameBytes.length];
			System.arraycopy( imageNameBytes, 0, newbytes, 0, imageNameBytes.length );
			imageNameBytes = newbytes;
			setProperty( MSODrawingConstants.msooptpibName, true, true, imageNameBytes.length, imageNameBytes );
		}
		catch( Exception e )
		{
			log.error( "Msofbt.setImageName failed.", e );
		}
	}

	/**
	 * set the shape name atom in this OPT record
	 *
	 * @param name
	 */
	public void setShapeName( String name )
	{
		try
		{
			shapeName = name;
			if( (shapeName == null) || shapeName.equals( "" ) )
			{
				// remove property
				props.remove( Integer.valueOf( MSODrawingConstants.msooptwzName ) );
				return;
			}
			byte[] shapeNameBytes = name.getBytes( XLSConstants.UNICODEENCODING );
			byte[] newbytes = new byte[shapeNameBytes.length];
			System.arraycopy( shapeNameBytes, 0, newbytes, 0, shapeNameBytes.length );
			shapeNameBytes = newbytes;
			setProperty( MSODrawingConstants.msooptwzName, true, true, shapeNameBytes.length, shapeNameBytes );
		}
		catch( Exception e )
		{
			log.error( "Msofbt.setShapeName failed.", e );
		}
	}

	/**
	 * generate the recordData from the stored props hashmap if anything has changed
	 */
	@Override
	protected byte[] getData()
	{
		if( isDirty )
		{    // regenerate recordData as contents have changed
			byte[] tmp = new byte[inst * 6];        // basic property table
			byte[] complexData = new byte[0];    // extra complex data, if any, after basic property table
			int pos = 0;
			// try to extract properties in order
			java.util.ArrayList keys = new java.util.ArrayList( props.keySet() );
			// try in numerical order-- appears to MOSTLY be the case ...
			Object[] k = keys.toArray();
			Arrays.sort( k );

			// write out properties in (numerical) order
			for( Object aK : k )
			{
				Integer propId = ((Integer) aK);
				Object[] o = (Object[]) props.get( propId );
				boolean isComplex = (Boolean) o[0];
				boolean isBid = (Boolean) o[2];
				int flag = 0;
				if( isComplex )
				{
					flag = flag | 0x80;
				}
				if( isBid )
				{
					flag = flag | 0x40;
				}
				int dtx;
				if( !isComplex )
				{
					dtx = (Integer) o[1];    // non-complex data is just an integer
				}
				else
				{
					dtx = ((byte[]) o[1]).length + 2;        // stored data is a byte array; get length + 2
					complexData = ByteTools.append( (byte[]) o[1], complexData );
					complexData = ByteTools.append( new byte[]{ 0, 0 }, complexData );
				}
				// the basic part of the property table
				tmp[pos++] = ((byte) (0xFF & propId));
				tmp[pos++] = (byte) (flag | ((0x3F00 & propId) >> 8));
				byte[] dtxBytes = ByteTools.cLongToLEBytes( dtx );
				System.arraycopy( dtxBytes, 0, tmp, pos, 4 );
				pos += 4;
			}
			recordData = new byte[tmp.length + complexData.length];
			System.arraycopy( tmp, 0, recordData, 0, tmp.length );
			// after the basic property table (PropID, IsBID, IsCOMPEX, dtx), store the complex data 
			System.arraycopy( complexData, 0, recordData, tmp.length, complexData.length );
			isDirty = false;
		}
		this.setLength( recordData.length );
		return recordData;
	}

	public void setData( byte[] b )
	{
		recordData = b;
		props.clear();
		imageIndex = -1;
		imageName = "";
		shapeName = "";
		alternateText = "";
		parseData();
	}

	/**
	 * given property table bytes, parse into props hashmap
	 */
	private void parseData()
	{
		/* 
		 * First part of an OPT record is an array of FOPTEs (propertyId, fBid, fComplex, data) 
		 * If fComplex is set, the actual data (Unicode strings, arrays, etc.) is stored AFTER the last FOPTE (sorted by property id???);
		 * the length of the complex data is stored in the data field.
		 * if fComplex is not set, the meaining of the data field is dependent upon the propertyId
		 * if fBid is set and fComplex is not set, the data = a BLIP id (= an index into the BLIP store)
		 * The number of FOPTES is the inst field read above
		 */
		int propertyId, fBid, fComplex;
		//int n= inst;				// number of properties to parse
		int pos = 0;                    // pointer to current property in data/property table
		if( (inst == 0) && (recordData.length > 0) )
		{    // called from GelFrame ...
			byte[] dat = new byte[8];    // read header
			System.arraycopy( recordData, 0, dat, 0, 8 );
			version = (0xF & dat[0]);    // 15 for containers, version for atoms
			inst = ((0xFF & dat[1]) << 4) | (0xF0 & dat[0]) >> 4;
			fbt = ((0xFF & dat[3]) << 8) | (0xFF & dat[2]);    // record type id==0xF00B
			pos = 8;    // skip header
		}
		for( int i = 0; i < inst; i++ )
		{
			propertyId = (0x3F & recordData[pos + 1]) << 8 | (0xFF & recordData[pos]);    // 14 bits
			fBid = ((0x40 & recordData[pos + 1]) >> 6);        // specifies whether the value in the dtx field is a BLIP identifier- only valid if fComplex= FALSE
			fComplex = ((0x80 & recordData[pos + 1]) >> 7);    // if complex property, value is length.  Data is parsed after.
			int dtx = ByteTools.readInt( recordData[pos + 2], recordData[pos + 3], recordData[pos + 4], recordData[pos + 5] );
			// TODO: if property number is of type bool/long/msoarray/... parse accordingly ..
			if( propertyId == MSODrawingConstants.msooptpib )    // blip to display
			{
				imageIndex = dtx;
			}
			else if( propertyId == MSODrawingConstants.msooptFillType )
			{
				fillType = dtx;
			}
			else if( propertyId == MSODrawingConstants.msooptfillColor )
			{
				fillColor = setFillColor( dtx );
			}
			else if( propertyId == MSODrawingConstants.msooptfBackground )
			{
				bBackground = (dtx != 0);
			}
			else if( propertyId == MSODrawingConstants.msooptGroupShapeProperties )
			{
				//bPrint= (dtx!=0); // NOT TRUE!! TODO: parse real GroupShapeProperties (many)
			}
			else if( propertyId == MSODrawingConstants.msooptpictureActive )
			{
				bActive = (dtx != 0);
			}
			else if( propertyId == MSODrawingConstants.msooptlineWidth )
			{    // appears that this controls display of line
				if( lineprops == null )
				{
					lineprops = new int[3];
				}
				lineprops[LINEPROPTS_WEIGHT] = dtx;
			}
			else if( propertyId == MSODrawingConstants.msooptlineColor )
			{    // appears that this is always present, even if no line
				if( lineprops == null )
				{
					lineprops = new int[3];
				}
				lineprops[LINEPROPTS_COLOR] = dtx;
			}
			else if( propertyId == MSODrawingConstants.msooptLineStyle )
			{
				if( lineprops == null )
				{
					lineprops = new int[3];
				}
				lineprops[LINEPROPTS_STYLE] = dtx;
			}
			else if( propertyId == MSODrawingConstants.msofbtlTxid )
			{
				hasTextId = true;
			} // msooptFillWidth
			props.put( propertyId,
			           new Object[]{ fComplex != 0, dtx, fBid != 0 } );
			pos += 6;
		}

		// now parse complex data after all "tightly packed" properties have been parsed.  Order of data is original order
		Iterator ii = props.keySet().iterator();
		while( ii.hasNext() )
		{
			Integer propId = (Integer) ii.next();
			Object[] o = (Object[]) props.get( propId );    // Object[]:  0= isComplex, 1= dtx (value or len of complex data -- filled in below), 2= isBid
			if( (Boolean) o[0] )
			{
				int len = (Integer) o[1];
				if( len >= 2 )
				{
					// apparently each record is delimited by a double byte 0 -- so decrement by 2 here and increment pos by 2 below 
					byte[] complexdata = new byte[len - 2];    // retrieve complex data at end of record
					System.arraycopy( recordData,
					                  pos,
					                  complexdata,
					                  0,
					                  complexdata.length );    // get property data after main property table
					props.put( propId, new Object[]{ o[0], complexdata, o[2] } );    //store complex data for later retrieval
					if( propId == MSODrawingConstants.msooptpibName )
					{ // = image name
						try
						{
							imageName = new String( complexdata, XLSConstants.UNICODEENCODING );
						}
						catch( Exception e )
						{
							imageName = "Unnamed";
						}
					}
					else if( propId == MSODrawingConstants.msooptwzName )
					{ // = shape name
						try
						{
							shapeName = new String( complexdata, XLSConstants.UNICODEENCODING );
						}
						catch( Exception e )
						{
							;
						}
					}
					else if( propId == MSODrawingConstants.msooptwzDescription )
					{ // = Alternate Text
						try
						{
							alternateText = new String( complexdata, XLSConstants.UNICODEENCODING );
						}
						catch( Exception e )
						{
							;
						}
					}
					else if( propId == MSODrawingConstants.msooptFillBlipName )
					{    // = the comment, file name, or the full URL that is used as a fill
						try
						{
							String fillName = new String( complexdata, XLSConstants.UNICODEENCODING );
						}
						catch( Exception e )
						{
							;
						}
					}

					pos += complexdata.length + 2;
				}
			}
		}
	}

	/**
	 * @param propId       msofbtopt property ide see Msoconstants
	 * @param isBid        value is a BLIP id - only valid if isComplex is false
	 * @param isComplex    if false, dtx is used; if true, complexBytes are used and dtx=length
	 * @param dtx          if not iscomplex, the value of property id; if iscomplex, length of complex data following the property table
	 * @param complexBytes if iscomplex, holds value of complex property e.g. shape name
	 */
	public void setProperty( int propId, boolean isBid, boolean isComplex, int dtx, byte[] complexBytes )
	{
		// a general order of common properties is (via property id):
		/*
		 * 127
		 * 267
		 * 261
		 * 262
		 * 128
		 * 133
		 * 139
		 * 191
		 * 385
		 * 447
		 * 448
		 * 459
		 * 511 
		 */
		if( isComplex ) // complexBytes shouldn't be null
		{
			props.put( propId, new Object[]{ isComplex, complexBytes, isBid } );
		}
		else
		{
			props.put( propId, new Object[]{
					isComplex, dtx, isBid
			} );
		}

		this.inst = props.size();
		isDirty = true;    // flag to regenerate recordData
	}

	public boolean hasBorder()
	{
		if( (lineprops != null) && (lineprops[LINEPROPTS_WEIGHT] > 1) )
		{
			return true;
		}
		return false;
	}

	public int getBorderLineWidth()
	{
		if( lineprops != null )
		{
			return lineprops[LINEPROPTS_WEIGHT];
		}
		return -1;
	}

	public int getImageIndex()
	{
		return imageIndex;
	}

	public String getImageName()
	{
		return imageName;
	}

	public String getShapeName()
	{
		return shapeName;
	}

	/**
	 * Debug Output -- For Internal Use Only
	 */
	public String debugOutput()
	{
		int propertyId;
		StringBuffer log = new StringBuffer();
/*		java.util.ArrayList keys= new java.util.ArrayList(props.keySet());
		int n= keys.size();
		for (int i= 0; i < n; i++) {
			log.append("\r\n");			
			propertyId = ((Integer)keys.get(i)).intValue(); //   (0x3F&recordData[pos+1])<<8|(0xFF&recordData[pos]);
			Object[] o= (Object[]) props.get(keys.get(i));
			boolean isComplex= ((Boolean)o[0]).booleanValue();
			boolean isBid= ((Boolean)o[2]).booleanValue();
			if (isComplex) fComplex=1;
			if (isBid)	fBid= 1;
			int dtx; 
			if (!isComplex)
				dtx= ((Integer)o[1]).intValue();	// non-complex data is just an integer
			else 
				dtx= ((byte[])o[1]).length + 2;		// stored data is a byte array; get length + 2					
			//fBid =  ((0x40&recordData[pos+1])>>6);		// value is a BLIP ID - only valid if fComplex= FALSE
			//fComplex = ((0x80&recordData[pos+1])>>7);	// if complex property, value is length.  Data is parsed after.
			//int dtx = ByteTools.readInt(recordData[pos+2],recordData[pos+3],recordData[pos+4],recordData[pos+5]);
			log.append("\t\t" + propertyId + "/" + fBid + "/" + fComplex + "/" + dtx);
			//pos+=6;
		}
*/
		int n = inst;                // number of properties to parse
		// pointer to current property in data/property table
		int fBid = 0, fComplex = 0;
		int end = recordData.length;
		for( int pos = 0; pos < end; pos += 6 )
		{
			propertyId = (0x3F & recordData[pos + 1]) << 8 | (0xFF & recordData[pos]);
			fBid = ((0x40 & recordData[pos + 1]) >> 6);
			fComplex = ((0x80 & recordData[pos + 1]) >> 7);    // if complex property, value is length.  Data is parsed after.
			int dtx = ByteTools.readInt( recordData[pos + 2], recordData[pos + 3], recordData[pos + 4], recordData[pos + 5] );
			if( fComplex != 0 )
			{
				end -= dtx;
			}
			log.append( "\t\t" + propertyId + "/" + fBid + "/" + fComplex + "/" + dtx );
		}
		return log.toString();
	}

	/**
	 * Interpret an OfficeArtCOLORREF (used in fillColor, lineColor and many other opts)
	 *
	 * @param clrStructure <br> More information
	 *                     The OfficeArtCOLORREF structure specifies a color. The high 8 bits MAY be set to 0xFF, in which case the color MUST be ignored.
	 *                     The color properties that are specified in the following table have a set of extended-color properties. The color property specifies the main color.
	 *                     The colorExt and colorExtMod properties specify the extended colors that can be used to define the main color more precisely.
	 *                     If neither extended-color property is set, the main color property contains the full color definition.
	 *                     Otherwise, the colorExt property specifies the base color, and the colorExtMod property specifies a tint or shade modification that is applied to the colorExt property.
	 *                     In this case, the main color property contains the flattened RGB color that is computed by applying the specified tint or shade modification to the specified base color.
	 *                     <p/>
	 *                     <p/>
	 *                     A - unused1 (1 bit): A bit that is undefined and MUST be ignored.
	 *                     <p/>
	 *                     B - unused2 (1 bit): A bit that is undefined and MUST be ignored.
	 *                     <p/>
	 *                     C - unused3 (1 bit): A bit that is undefined and MUST be ignored.
	 *                     <p/>
	 *                     D - fSysIndex (1 bit): A bit that specifies whether the system color scheme will be used to determine the color.
	 *                     A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index into the system color table. Values less than 0x00F0 map directly to system colors.
	 *                     For more information, see [MSDN-GetSysColor] (below)
	 *                     The following table specifies values that have special meaning.
	 *                     Value		Meaning
	 *                     0x00F0		Use the fill color of the shape.
	 *                     0x00F1		If the shape contains a line, use the line color of the shape. Otherwise, use the fill color.
	 *                     0x00F2		Use the line color of the shape.
	 *                     0x00F3		Use the shadow color of the shape.
	 *                     0x00F4	    Use the current, or last-used, color.
	 *                     0x00F5	    Use the fill background color of the shape.
	 *                     0x00F6	    Use the line background color of the shape.
	 *                     0x00F7	    If the shape contains a fill, use the fill color of the shape. Otherwise, use the line color.
	 *                     The following table specifies values that indicate special procedural properties that are used to modify the color components of another color.
	 *                     These values are combined with those in the preceding table or with a user-specified color. The first six values are mutually exclusive.
	 *                     Value		Meaning
	 *                     0x0100	    Darken the color by the value that is specified in the blue field. A blue value of 0xFF specifies that the color is to be left unchanged, whereas a blue value of 0x00 specifies that the color is to be completely darkened.
	 *                     0x0200	    Lighten the color by the value that is specified in the blue field. A blue value of 0xFF specifies that the color is to be left unchanged, whereas a blue value of 0x00 specifies that the color is to be completely lightened.
	 *                     0x0300  	Add a gray level RGB value. The blue field contains the gray level to add:    NewColor = SourceColor + gray
	 *                     0x0400	    Subtract a gray level RGB value. The blue field contains the gray level to subtract:	NewColor = SourceColor - gray
	 *                     0x0500		Reverse-subtract a gray level RGB value. The blue field contains the gray level from which to subtract:	    NewColor = gray - SourceColor
	 *                     0x0600	    If the color component being modified is less than the parameter contained in the blue field, set it to the minimum intensity.
	 *                     If the color component being modified is greater than or equal to the parameter, set it to the maximum intensity.
	 *                     0x2000	    After making other modifications, invert the color.
	 *                     0x4000	    After making other modifications, invert the color by toggling just the high bit of each color channel.
	 *                     0x8000      Before making other modifications, convert the color to grayscale.
	 *                     E - fSchemeIndex (1 bit): A bit that specifies whether the current application-defined color scheme will be used to determine the color.
	 *                     A value of 0x1 specifies that red will be treated as an index into the current color scheme table. If this value is 0x1, green and blue MUST be 0x00.
	 *                     F - fSystemRGB (1 bit): A bit that specifies whether the color is a standard RGB color. The following table specifies the meaning of each value for this field.
	 *                     Value		Meaning
	 *                     0x0			The RGB color MAY use halftone dithering to display.
	 *                     0x1		    The color MUST be a solid color.
	 *                     G - fPaletteRGB (1 bit): A bit that specifies whether the current palette will be used to determine the color.
	 *                     A value of 0x1 specifies that red, green, and blue contain an RGB value that will be matched in the current color palette. This color MUST be solid.
	 *                     H - fPaletteIndex (1 bit): A bit that specifies whether the current palette will be used to determine the color.
	 *                     A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index into the current color palette. This color MAY<1> be dithered.
	 *                     If this value is 0x1, blue MUST be 0x00.
	 *                     blue (1 byte): An unsigned integer that specifies the intensity of the blue color channel. A value of 0x00 has the minimum blue intensity. A value of 0xFF has the maximum blue intensity.
	 *                     green (1 byte): An unsigned integer that specifies the intensity of the green color channel. A value of 0x00 has the minimum green intensity. A value of 0xFF has the maximum green intensity.
	 *                     red (1 byte): An unsigned integer that specifies the intensity of the red color channel. A value of 0x00 has the minimum red intensity. A value of 0xFF has the maximum red intensity.
	 *                     <p/>
	 *                     ...
	 *                     <p/>
	 *                     MSDN-GetSysColor
	 *                     Value						Meaning
	 *                     COLOR_3DDKSHADOW	21		Dark shadow for three-dimensional display elements.
	 *                     COLOR_3DFACE		15		Face color for three-dimensional display elements and for dialog box backgrounds.
	 *                     COLOR_3DHIGHLIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
	 *                     COLOR_3DHILIGHT		20		Highlight color for three-dimensional display elements (for edges facing the light source.)
	 *                     COLOR_3DLIGHT		22		Light color for three-dimensional display elements (for edges facing the light source.)
	 *                     COLOR_3DSHADOW		16		Shadow color for three-dimensional display elements (for edges facing away from the light source).
	 *                     COLOR_ACTIVEBORDER	10		Active window border.
	 *                     COLOR_ACTIVECAPTION	2		Active window title bar.	Specifies the left side color in the color gradient of an active window's title bar if the gradient effect is enabled.
	 *                     COLOR_APPWORKSPACE	12		Background color of multiple document interface (MDI) applications.
	 *                     COLOR_BACKGROUND	1		Desktop.
	 *                     COLOR_BTNFACE		15		Face color for three-dimensional display elements and for dialog box backgrounds.
	 *                     COLOR_BTNHIGHLIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
	 *                     COLOR_BTNHILIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
	 *                     COLOR_BTNSHADOW		16		Shadow color for three-dimensional display elements (for edges facing away from the light source).
	 *                     COLOR_BTNTEXT		18		Text on push buttons.
	 *                     COLOR_CAPTIONTEXT	9		Text in caption, size box, and scroll bar arrow box.
	 *                     COLOR_DESKTOP		1		Desktop.
	 *                     COLOR_GRADIENTACTIVECAPTION	27	Right side color in the color gradient of an active window's title bar.
	 *                     COLOR_ACTIVECAPTION specifies the left side color. Use SPI_GETGRADIENTCAPTIONS with the SystemParametersInfo function to determine whether the gradient effect is enabled.
	 *                     COLOR_GRADIENTINACTIVECAPTION	28	Right side color in the color gradient of an inactive window's title bar. COLOR_INACTIVECAPTION specifies the left side color.
	 *                     COLOR_GRAYTEXT		17		Grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color.
	 *                     COLOR_HIGHLIGHT		13		Item(s) selected in a control.
	 *                     COLOR_HIGHLIGHTTEXT	14		Text of item(s) selected in a control.
	 *                     COLOR_HOTLIGHT		26		Color for a hyperlink or hot-tracked item.
	 *                     COLOR_INACTIVEBORDER11		Inactive window border.
	 *                     COLOR_INACTIVECAPTION3		Inactive window caption.
	 *                     Specifies the left side color in the color gradient of an inactive window's title bar if the gradient effect is enabled.
	 *                     COLOR_INACTIVECAPTIONTEXT19	Color of text in an inactive caption.
	 *                     COLOR_INFOBK		24		Background color for tooltip controls.
	 *                     COLOR_INFOTEXT		23		Text color for tooltip controls.
	 *                     COLOR_MENU			4		Menu background.
	 *                     COLOR_MENUHILIGHT	29		The color used to highlight menu items when the menu appears as a flat menu (see SystemParametersInfo). The highlighted menu item is outlined with COLOR_HIGHLIGHT.
	 *                     Windows 2000:  This value is not supported.
	 *                     COLOR_MENUBAR		30		The background color for the menu bar when menus appear as flat menus (see SystemParametersInfo). However, COLOR_MENU continues to specify the background color of the menu popup.
	 *                     Windows 2000:  This value is not supported.
	 *                     COLOR_MENUTEXT		7		Text in menus.
	 *                     COLOR_SCROLLBAR		0		Scroll bar gray area.
	 *                     COLOR_WINDOW		5		Window background.
	 *                     COLOR_WINDOWFRAME	6		Window frame.
	 *                     COLOR_WINDOWTEXT	8		Text in windows.
	 */
	private java.awt.Color setFillColor( int clrStructure )
	{
		byte[] b = ByteTools.longToByteArray( clrStructure );
		boolean bPaletteIndex, bSchemeIndex, bSysIndex;
		short fillclr;

		bPaletteIndex = (b[4] & 0x1) == 0x1;    // 	specifies whether the current palette will be used to determine the color
		bSchemeIndex = (b[4] & 0x8) == 0x8;    //  specifies whether the current application defined color scheme will be used to determine the color.
		bSysIndex = (b[4] & 0x10) == 0x10;    //  specifies whether the system color scheme will be used to determine the color.

		if( bPaletteIndex )
		{    // // GREEN and RED are treated as an unsigned 16-bit index into the current color palette. This color MAY be dithered. BLUE MUST be 0x00.
		}
		if( bSchemeIndex )
		{        //  RED is an index into the current scheme color table. GREEN and BLUE MUST be 0x00.
			fillclr = b[7];            // what does 80 mean??????
			if( fillclr > FormatHandle.COLORTABLE.length )
			{
				fillclr = FormatHandle.interpretSpecialColorIndex( fillclr );
			}
			fillColor = FormatHandle.COLORTABLE[fillclr];
			return fillColor;
		}
		if( bSysIndex )
		{    // GREEN and RED will be treated as an unsigned 16-bit index into the system color table. Values less than 0x00F0 map directly to system colors.
			fillclr = ByteTools.readShort( b[6], b[7] );
			if( (fillclr == 0x00F0) //		Use the fill color of the shape.
					|| (fillclr == 0x00F1)        //If the shape contains a line, use the line color of the shape. Otherwise, use the fill color.
					|| (fillclr == 0x00F2)        //Use the line color of the shape.
					|| (fillclr == 0x00F3)        //Use the shadow color of the shape.
					|| (fillclr == 0x00F4)        //Use the current, or last-used, color.
					|| (fillclr == 0x00F5)        //Use the fill background color of the shape.
					|| (fillclr == 0x00F6)        //Use the line background color of the shape.
					|| (fillclr == 0x00F7) )        //If the shape contains a fill, use the fill color of the shape. Otherwise, use the line color.
			{
				fillclr = FormatConstants.COLOR_WHITE;
			}
			if( fillclr == 0x40 )    // default fg color
			{
				fillclr = FormatConstants.COLOR_WHITE;
			}
			else if( fillclr == 0x41 )    // default bg color
			{
				fillclr = FormatConstants.COLOR_WHITE;
			}
			else if( fillclr == 0x4D )
			{        // default CHART fg color -- INDEX SPECIFIC!
				fillColor = null;    // flag to map via series (bar) color defaults
				return fillColor;
			}
			else if( fillclr == 0x4E )    // default CHART fg color
			{
				fillclr = FormatConstants.COLOR_WHITE;
			}
			else if( fillclr == 0x4F )    // chart neutral color == black
			{
				fillclr = FormatConstants.COLOR_BLACK;
			}

			if( (fillclr < 0) || (fillclr > FormatHandle.COLORTABLE.length) )
			{
				fillclr = FormatConstants.COLOR_WHITE;
			}
			fillColor = FormatHandle.COLORTABLE[fillclr];
			return fillColor;
		}

		// otherwise, r, g and blue are color values 0-255
		int bl = ((b[5] < 0) ? (255 + b[5]) : b[5]);
		int g = ((b[6] < 0) ? (255 + b[6]) : b[6]);
		int r = ((b[7] < 0) ? (255 + b[7]) : b[7]);
		fillColor = new java.awt.Color( r, g, bl );
		return fillColor;
	}

	public java.awt.Color getFillColor()
	{
		return fillColor;
	}

	public int getFillType()
	{
		return fillType;
	}
}
