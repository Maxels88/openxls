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
package com.extentech.ExtenXLS;

import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.OOXML.TwoCellAnchor;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.MSODrawing;
import com.extentech.formats.XLS.MSODrawingConstants;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.Iterator;

//OOXML-specific structures

/**
 * The ImageHandle provides access to an Image embedded in a spreadsheet.<br>
 * <br>
 * Use the ImageHandle to work with images in spreadsheet.<br>
 * <br>  <br>
 * With an ImageHandle you can:
 * <br><br>
 * <blockquote>
 * insert images into your spreadsheet
 * set the position of the image
 * set the width and height of the image
 * write spreadsheet image files to any outputstream
 * <p/>
 * </blockquote>
 * <br>
 * <br>
 *
 * @see com.extentech.ExtenXLS.WorkBookHandle
 * @see com.extentech.ExtenXLS.WorkSheetHandle
 */
public class ImageHandle implements Serializable
{
	private static final Logger log = LoggerFactory.getLogger( ImageHandle.class );
	/**
	 * returns width divided by height for the aspect ratio
	 *
	 * @return the aspect ratio for the image
	 */
	public double getAspectRatio()
	{
		double height, width;
		double aspectRatio = 0;
		height = getHeight() / 122.27;  //(convert to in)
		width = getWidth() / 57.06;     //(convert to in)
		aspectRatio = width / height;
		return aspectRatio;
	}

	/**
	 *
	 *
	 */
	private static final long serialVersionUID = 3177017738178634238L;
	private byte[] imageBytes;
	private Boundsheet mysheet;

	// coordinates
	public static final int X = 0;
	public static final int Y = 1;
	public static final int WIDTH = 2;
	public static final int HEIGHT = 3;

	private String imageName = " ", shapeName = "";
	private short height, width;
	private short x, y;
	private int image_type = -1;

	// OOXML-specific
	private SpPr imagesp;
	private String editMovement = null;

	private MSODrawing thisMsodrawing;    //20070924 KSC: link to actual msodrawing rec that describes this image

	public void setMsgdrawing( MSODrawing rec )
	{
		thisMsodrawing = rec;
	}

	public MSODrawing getMsodrawing()
	{
		return thisMsodrawing;
	}

	/**
	 * Returns the image type
	 *
	 * @return
	 */
	public String getType()
	{
		switch( image_type )
		{
			case com.extentech.formats.XLS.MSODrawingConstants.IMAGE_TYPE_GIF:
				return "gif";
			case com.extentech.formats.XLS.MSODrawingConstants.IMAGE_TYPE_JPG:
				return "jpeg";
			case com.extentech.formats.XLS.MSODrawingConstants.IMAGE_TYPE_PNG:
				return "png";
			case com.extentech.formats.XLS.MSODrawingConstants.IMAGE_TYPE_EMF:
				return "emf";
			default:
				return "undefined";
		}
	}

	/**
	 * Returns the image mime type
	 *
	 * @return
	 */
	public String getMimeType()
	{
		return "image/" + getImageType();
	}

	/**
	 * Constructor which takes image file bytes and inserts into
	 * specific sheet
	 *
	 * @param imageBytes
	 * @param sheet
	 */
	public ImageHandle( InputStream imagebytestream, WorkSheetHandle _sheet )
	{
		this( imagebytestream, _sheet.getMysheet() );
	}

	/**
	 * Constructor  which takes image file bytes and associates it with the
	 * specified boundsheet
	 *
	 * @param imagebytestream
	 * @param bs
	 */
	public ImageHandle( InputStream imagebytestream, Boundsheet bs )
	{
		mysheet = bs;
		try
		{
			imageBytes = new byte[imagebytestream.available()];
			imagebytestream.read( imageBytes );
		}
		catch( Exception ex )
		{
			log.error( "Failed to create new ImageHandle in sheet" + bs.getSheetName() + " from InputStream:" + ex.toString() );
		}
		initialize();
	}

	/**
	 * Constructor which takes image file bytes and inserts into
	 * specific sheet
	 *
	 * @param imageBytes
	 * @param sheet
	 */
	public ImageHandle( byte[] _imageBytes, Boundsheet sheet )
	{
		mysheet = sheet;
		imageBytes = _imageBytes;

		initialize();

	}

	/**
	 * Override equals so that Sheet cannot contain dupes.
	 *
	 * @see java.lang.Object#equals(java.lang.Object)
	 */
	public boolean equals( Object another )
	{
		if( another.toString().equals( this.toString() ) )
		{
			return true;
		}
		return false;
	}

	private void initialize()
	{
		// after instantiating the image you should have
		// access to the width and height bounds
		String imageFormat = "";
		try
		{
			ByteArrayInputStream bis = new ByteArrayInputStream( imageBytes );

			imageFormat = getImageFormat( bis );
			if( imageFormat == null )
			{
				// 20080128 KSC: Occurs when image format= EMF ????? We cannot interpret, will most likely crash file
					log.error( "ImageHandle.initialize: Unrecognized Image Format" );
				return;
			}
			if( !(imageFormat.equalsIgnoreCase( "jpeg" ) || imageFormat.equalsIgnoreCase( "png" )) )
			{
				bis.reset();
				imageBytes = convertData( bis );
				image_type = MSODrawingConstants.IMAGE_TYPE_PNG;
			}
			else
			{
				if( imageFormat.equalsIgnoreCase( "jpeg" ) )
				{
					image_type = MSODrawingConstants.IMAGE_TYPE_JPG;
				}
				else
				{
					this.image_type = MSODrawingConstants.IMAGE_TYPE_PNG;
				}

			}

			try
			{
				BufferedImage bi = ImageIO.read( new ByteArrayInputStream( imageBytes ) );
				width = (short) bi.getWidth();
				height = (short) bi.getHeight();
				bi = null;
			}
			catch( IOException ex )
			{
				log.warn( "Java ImageIO could not decode image bytes for " + imageName + ":" + ex.toString() );
				if( false )
				{    // 20081028 KSC: don't overwrite original image bytes
					String imgname = "failed_image_read_" + this.imageName + "." + imageFormat;
					FileOutputStream outimg = new FileOutputStream( imgname );
					outimg.write( imageBytes );
					outimg.flush();
					outimg.close();

				}
			}
			// 20070924 KSC: bounds[2] = width;
			// "" bounds[3] = height;
			this.imageName = "UnnamedImage";
		}
		catch( Exception e )
		{
			log.warn( "Problem creating ImageHandle:" + e.toString() + " Please see BugTrack article: http://extentech.com/uimodules/docs/docs_detail.jsp?meme_id=1431&showall=true", e );
			if( false )
			{ // debug image parse probs
				String imgname = "failed_image_read_" + this.imageName + "." + imageFormat;
				try
				{
					FileOutputStream outimg = new FileOutputStream( imgname );
					outimg.write( imageBytes );
					outimg.flush();
					outimg.close();
				}
				catch( Exception ex )
				{
					;
				}
			}
		}
	}

	/**
	 * update the underlying image record
	 */
	public void update() throws Exception
	{
		if( this.thisMsodrawing == null )
		{
			throw new Exception( "ImageHandle.Update: Image Not initialzed" );
		}
		this.thisMsodrawing.updateRecord();    //this.thisMsodrawing.getSPID());
		this.mysheet.getWorkBook().updateMsodrawingHeaderRec( this.mysheet );
	}

	/**
	 * returns true if this drawing is active i.e. not deleted
	 * <br>Note this is experimental
	 *
	 * @return
	 */
	public boolean isActive()
	{
		if( this.thisMsodrawing != null )
		{
			return this.thisMsodrawing.isActive();
		}
		return true;    // default
	}

	/**
	 * converts image data into byte array used by
	 * <p/>
	 * Jan 22, 2010
	 *
	 * @param imagebytestream
	 * @return
	 */
	public byte[] convertData( InputStream imagebytestream )
	{
		try
		{
			ByteArrayOutputStream bos = new ByteArrayOutputStream();

			BufferedImage bi = ImageIO.read( imagebytestream );

			ImageIO.write( bi, "png", bos );

			return bos.toByteArray();
		}
		catch( Exception e )
		{
			log.error( "ImageHandle.convertData: " + e.toString(), e );
			return null;
		}

	}

	/**
	 * returns the format name of the image data
	 * <p/>
	 * Jan 22, 2010
	 *
	 * @param imagebytestream
	 * @return
	 */
	public String getImageFormat( InputStream imagebytestream )
	{

		try
		{
			// Create an image input stream on the image
			ImageInputStream iis = ImageIO.createImageInputStream( imagebytestream );

			// Find all image readers that recognize the image format
			Iterator iter = ImageIO.getImageReaders( iis );
			if( !iter.hasNext() )
			{

				return null;
			}

			// Use the first reader
			ImageReader reader = (ImageReader) iter.next();

			// Close stream
			iis.close();

			// Return the format name
			return reader.getFormatName();
		}
		catch( IOException e )
		{
		}
		// The image could not be read
		return null;
	}

    
    
    /*
	 * 20070924 KSC: image bounds are set via column/row # + offsets within the respective cell
	 * added many position methods that access msodrawing for actual positions
	 */
	/* position methods */

	/**
	 * return the image bounds
	 * images bounds are as follows:
	 * bounds[0]= column # of top left position (0-based) of the shape
	 * bounds[1]= x offset within the top-left column	(0-1023)
	 * bounds[2]= row # for top left corner
	 * bounds[3]= y offset within the top-left corner (0-1023)
	 * bounds[4]= column # of the bottom right corner of the shape
	 * bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
	 * bounds[6]= row # for bottom-right corner of the shape
	 * bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
	 */
	public short[] getBounds()
	{
		return thisMsodrawing.getBounds();
	}

	// short is not big enough to handle MAXROWS... correct? -jm

	/**
	 * sets the image bounds
	 * images bounds are as follows:
	 * bounds[0]= column # of top left position (0-based) of the shape
	 * bounds[1]= x offset within the top-left column (0-1023)
	 * bounds[2]= row # for top left corner
	 * bounds[3]= y offset within the top-left corner	(0-1023)
	 * bounds[4]= column # of the bottom right corner of the shape
	 * bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
	 * bounds[6]= row # for bottom-right corner of the shape
	 * bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
	 */
	public void setBounds( short[] bounds )
	{
		thisMsodrawing.setBounds( bounds );
	}

	/**
	 * return the image bounds in x, y, width, height format in pixels
	 *
	 * @return
	 */
	public short[] getCoords()
	{
		if( thisMsodrawing != null )
		{
			return thisMsodrawing.getCoords();
		}
		return new short[]{ x, y, width, height };
	}

	/**
	 * set the image x, w, width and height in pixels
	 *
	 * @param x
	 * @param y
	 * @param w
	 * @param h
	 */
	public void setCoords( int x, int y, int w, int h )
	{
		if( thisMsodrawing != null )
		{
			thisMsodrawing.setCoords( new short[]{ (short) x, (short) y, (short) w, (short) h } );
		}
		else
		{// save for later
			this.x = (short) x;
			this.y = (short) y;
			this.width = (short) w;
			this.height = (short) h;
		}
	}

	/**
	 * set the image upper x coordinate in pixels
	 *
	 * @param x
	 */
	public void setX( int x )
	{
		if( thisMsodrawing != null )
		{
			thisMsodrawing.setX( (short) Math.round( x / 6.4 ) ); // 20090506 KSC: convert pixels to excel units
		}
		else // save for later
		{
			this.x = (short) x;
		}
	}

	/**
	 * set the image upper y coordinate in pixels
	 *
	 * @param y
	 */
	public void setY( int y )
	{
		if( thisMsodrawing != null )
		{
			thisMsodrawing.setY( (short) Math.round( y * 0.60 ) );    //convert pixels to points
		}
		else // save for later
		{
			this.y = (short) y;
		}
	}

	/**
	 * return the topmost row of the image
	 *
	 * @return
	 */
	public int getRow()
	{
		return thisMsodrawing.getRow0();
	}

	/**
	 * return the lower row of the image
	 *
	 * @return
	 */
	public int getRow1()
	{
		return thisMsodrawing.getRow1();
	}

	/**
	 * set the topmost row of the image
	 *
	 * @param row
	 */
	public void setRow( int row )
	{
		thisMsodrawing.setRow( row );
	}

	/**
	 * set the lower row of the image
	 *
	 * @param row
	 */
	public void setRow1( int row )
	{
		thisMsodrawing.setRow1( row );
	}

	/**
	 * return the leftmost column of the image
	 */
	public int getCol()
	{
		return thisMsodrawing.getCol();
	}

	/**
	 * return the rightmost column of the image
	 */
	public int getCol1()
	{
		return thisMsodrawing.getCol1();
	}

	public int getOriginalWidth()
	{
		return thisMsodrawing.getOriginalWidth();
	}

	/**
	 * return the width of the image in pixels
	 *
	 * @return
	 */
	public short getWidth()
	{
		return (short) Math.round( thisMsodrawing.getWidth() * 6.4 ); // 20090506 KSC: Convert excel units to pixels
	}

	/**
	 * set the width of the image in pixels
	 *
	 * @param w
	 */
	public void setWidth( int w )
	{
		thisMsodrawing.setWidth( (short) Math.round( w / 6.4 ) ); // 20090506 KSC: convert pixels to excel units
	}

	/**
	 * return the height of the image in pixels
	 *
	 * @return
	 */
	public short getHeight()
	{
		return (short) Math.round( thisMsodrawing.getHeight() / 0.60 ); // 20090506 KSC: Convert points to pixels
	}

	/**
	 * set the height of the image in pixels
	 *
	 * @param h
	 */
	public void setHeight( int h )
	{
		thisMsodrawing.setHeight( (short) Math.round( h * 0.60 ) ); // 20090506 KSC: convert pixels to points
	}

	/**
	 * return the upper x coordinate of the image in pixels
	 *
	 * @return
	 */
	public short getX()
	{
		return (short) Math.round( thisMsodrawing.getX() * 6.4 );  // 20090506 KSC: convert excel units to pixels
	}

	/**
	 * return the upper y coordinate of the image in pixels
	 *
	 * @return
	 */
	public short getY()
	{
		return (short) Math.round( thisMsodrawing.getY() * 0.60 );  // 20090506 KSC: convert points to pixels
	}

	/**
	 * get the position of the top left corner of the image in the sheet
	 */

	public short[] getRowAndOffset()
	{
		return thisMsodrawing.getRowAndOffset();
	}

	/**
	 * get the position of the top left corner of the image in the sheet
	 */
	public short[] getColAndOffset()
	{
		return thisMsodrawing.getColAndOffset();
	}

	/**
	 * Internal method that converts the image bounds appropriate for saving
	 * by Excel in the MsofbtOPT record.
	 */
	public int getImageIndex()
	{
		return thisMsodrawing.getImageIndex();
	}

	/**
	 * write the image bytes to an outputstream such as a file
	 *
	 * @param out
	 */
	//Modified by Bikash
	public void write( OutputStream out ) throws IOException
	{
		// TODO: write the image bytes out
		out.write( this.imageBytes );
	}

	/**
	 * Need to figure out a unique identifier for these -- if there
	 * is one in the BIFF8 information, that would be best
	 *
	 * @see java.lang.Object#toString()
	 */
	public String toString()
	{
		return getName();    // 20071026 KSC: imageName;
	}

	/**
	 * removes this Image from the WorkBook.
	 *
	 * @return whether the removal was a success
	 */
	public boolean remove()
	{
		this.mysheet.removeImage( this );
		// blow out the image rec
		this.thisMsodrawing.remove( true );
		return true;
	}

	/**
	 * @return Returns the image type.
	 */
	public int getImageType()
	{
		return image_type;
	}

	/**
	 * @param image_type The image type to set.
	 */
	public void setImageType( int type )
	{
		image_type = type;
	}

	/**
	 * returns the WorkSheet this image is contained in
	 *
	 * @return Returns the sheet.
	 */
	public Boundsheet getSheet()
	{
		return mysheet;
	}

	/**
	 * @return Returns the name (either explicitly set name or image name).
	 */
	public String getName()
	{    // 20071025 KSC: use Explicitly set name = shape name, if present.  Otherwise, use imageName
		if( !shapeName.equals( "" ) )
		{
			return shapeName;
		}
		return this.imageName;
	}

	/**
	 * return the imageName (= original file name, I believe)
	 *
	 * @return
	 */
	public String getImageName()
	{
		return this.imageName;
	}

	/**
	 * returns the explicitly set shape name (set by entering text in the named range field in Excel)
	 *
	 * @return the shape name
	 */
	public String getShapeName()
	{
		return shapeName;
	}

	/**
	 * @param name The name to set.
	 */
	public void setName( String name )
	{
		this.imageName = name;
		// NOTE: updating mso code causes corruption in certain Infoteria files- must look at
		if( (this.thisMsodrawing != null) && !this.thisMsodrawing.getName().equals( name ) )
		{
			this.thisMsodrawing.setImageName( name );    // update is done in setImageName
			this.thisMsodrawing.getWorkBook()
			                   .updateMsodrawingHeaderRec( this.getSheet() );    // must update header if change mso's - Claritas image insert regression bug (testImages.testInsertImageCorruption)
		}
	}

	/**
	 * allow setting of image name as seen in the Named Range Box
	 *
	 * @param name
	 */
	public void setShapeName( String name )
	{
		if( name != null )
		{
			shapeName = name;
		}
		else
		{
			shapeName = "";
		}
		// only set name/update record if names have changed
		if( (this.thisMsodrawing != null) && !shapeName.equals( this.thisMsodrawing.getShapeName() ) )
		{
			this.thisMsodrawing.setShapeName( shapeName );    // update is done in setShapeName
			this.thisMsodrawing.getWorkBook()
			                   .updateMsodrawingHeaderRec( this.getSheet() );    // 20100202 KSC: must update header if change mso's - Claritas image insert regression bug (testImages.testInsertImageCorruption)
		}
	}

	/**
	 * @return Returns the imageBytes.
	 */
	public byte[] getImageBytes()
	{
		return imageBytes;
	}

	/**
	 * sets the underlying image bytes to the bytes from a new image
	 * <p/>
	 * essentially swapping the image bytes for another, while retaining
	 * all other image properties (borders, etc.)
	 *
	 * @param imageBytes The imageBytes to set.
	 */
	public void setImageBytes( byte[] imageBytes )
	{
		this.imageBytes = imageBytes;
		mysheet.getWorkBook().getMSODrawingGroup().setImageBytes( this.imageBytes, mysheet, this.thisMsodrawing, this.getName() );
	}

	/**
	 * Get a JSON representation of the format
	 *
	 * @param cr
	 * @return
	 */
	public JSONObject getJSON()
	{
		JSONObject ch = new JSONObject();
		try
		{
			ch.put( "name", this.getImageName() );
			short[] coords = this.getCoords();

			// short[] coords =  { x, y, width, height };

			ch.put( "x", coords[ImageHandle.X] );
			ch.put( "y", coords[ImageHandle.Y] );

			ch.put( "width", coords[ImageHandle.WIDTH] );
			// ch.put("width", width); // for some reason COORDS wrong for width

			ch.put( "height", coords[ImageHandle.HEIGHT] );
			ch.put( "type", this.getType() );

		}
		catch( JSONException e )
		{
			log.error( "Error getting imageHandle JSON", e );
		}
		return ch;
	}

	/**
	 * set ImageHandle position based on another
	 *
	 * @param im source ImageHandle
	 * @return
	 */
	public void position( ImageHandle im )
	{
	    /* one way, just set x and y and keep original w and h
        //      set x and y, keep original width and height
        short[] origcoords= im.getCoords();
        short[] coords= this.getCoords();
        coords[0]= origcoords[0];	
        coords[1]= origcoords[1];
        this.setCoords(coords[0], coords[1], coords[2], coords[3]);
        */ 
    	/* other way, set with all original coordinates */
		this.setBounds( im.getBounds() );
	}

	/**
	 * return the XML representation of this image
	 *
	 * @return String
	 */
	public String getXML( int rId )
	{
		StringBuffer sb = new StringBuffer();
		short[] bounds = this.getBounds();
		final int EMU = 1270;    // 1 pt= 1270 EMUs		-- for consistency with OOXML

		sb.append( "<twoCellAnchor editAs=\"oneCell\">" );
		sb.append( "\r\n" );
		// top left coords
		sb.append( "<from>" );
		sb.append( "<col>" + bounds[0] + "</col>" );
		sb.append( "<colOff>" + (bounds[1] * EMU) + "</colOff>" );
		sb.append( "<row>" + bounds[2] + "</row>" );
		sb.append( "<rowOff>" + (bounds[3] * EMU) + "</rowOff>" );
		sb.append( "</from>" );
		sb.append( "\r\n" );
		// bottom right coords
		sb.append( "<to>" );
		sb.append( "<col>" + bounds[4] + "</col>" );
		sb.append( "<colOff>" + (bounds[5] * EMU) + "</colOff>" );
		sb.append( "<row>" + bounds[6] + "</row>" );
		sb.append( "<rowOff>" + (bounds[7] * EMU) + "</rowOff>" );
		sb.append( "</to>" );
		sb.append( "\r\n" );
		// Picture details - req. child elements= nvPicPr (non-visual picture properties), blipFill (links to image), spPr (shape properties)
		sb.append( "<pic>" );
		sb.append( "\r\n" );
		sb.append( "<nvPicPr>" );
		sb.append( "<cNvPr id=\"" + this.getMsodrawing().getSPID() + "\"" );
		sb.append( " name=\"" + this.getImageName() + "\"" );
		sb.append( " descr=\"" + this.getShapeName() + "\"/>" );
		sb.append( "<cNvPicPr>" );
		sb.append( "<picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>" );
		sb.append( "</cNvPicPr>" );
		sb.append( "</nvPicPr>" );
		sb.append( "\r\n" );
		// Picture relationship Id and relationship to the package
		sb.append( "<blipFill>" );
		sb.append( "<blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId" + rId + "\"/>" );
		//<a:srcRect/>	//If the picture is cropped, these details are stored in the <srcRect/> element
		sb.append( "<stretch><fillRect/></stretch>" );
		sb.append( "</blipFill>" );
		sb.append( "\r\n" );
		// shape properties
		sb.append( "<spPr>"/* bwMode=\"auto\">"*/ );
		sb.append( "<xfrm>" );
		int x = this.getX() * EMU, y = this.getY() * EMU, cx = this.getWidth() * EMU, cy = this.getHeight() * EMU;
		sb.append( "<off x=\"" + x + "\" y=\"" + y + "\"/>" );        // offsets= location
		sb.append( "<ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>" );    // extents= size of bounding box enclosing pic in EMUs
		sb.append( "</xfrm>" );
		sb.append( "<prstGeom prst=\"rect\">" );        // preset geometry, appears necessary for images ...
		sb.append( "<avLst/>" );
		sb.append( "</prstGeom>" );
		sb.append( "<noFill/>" );
		sb.append( "</spPr>" );
		sb.append( "</pic>" );
		sb.append( "\r\n" );
		sb.append( "<clientData/>" );
		sb.append( "\r\n" );
		sb.append( "</twoCellAnchor>" );
		return sb.toString();
	}

	/**
	 * return the (00)XML (or DrawingML) representation of this image
	 *
	 * @param rId relationship id for image file
	 * @return String
	 */
	public String getOOXML( int rId )
	{
		TwoCellAnchor t = new TwoCellAnchor( this.editMovement );
		t.setAsImage( rId, this.getImageName(), this.getShapeName(), this.getMsodrawing().getSPID(), this.getSpPr() );
		t.setBounds( TwoCellAnchor.convertBoundsFromBIFF8( this.getSheet(), this.getBounds() ) );    // adjust BIFF8 bounds to OOXML units
		return t.getOOXML();
		// missing in <xdr:pic>
		// <xdr:nvPicPr><xdr:cNvPicPr><a:picLocks noChangeArrowheads="1">
		// <xdr:blipFill><a:blip cstate="print">
		// <xdr:blipFill><a:srcRect/><a:stretch><a:fillRect/></a:stretch>
        
    	    	
/*    	
    	StringBuffer sb= new StringBuffer();
    	int[] bounds= twoCellAnchor.convertBoundsFromBIFF8(this.getSheet(), this.getBounds());	// adjust BIFF8 bounds to OOXML units
        final int EMU= 1270;	// 1 pt= 1270 EMUs

        // 20081008 KSC:  Added namespaces for Excel7 Use **********************
    	// TODO: create a twoCellAnchor from this ImageHandle and use twoCellAnchor.getOOXML
    	sb.append("<xdr:twoCellAnchor");
    		if (editMovement!=null) sb.append(" editAs=\"" + editMovement + "\"");	// how to resize or move upon editing	
    		sb.append(">\r\n");
    		// top left coords  
    		sb.append("<xdr:from>");
    			sb.append("<xdr:col>" + bounds[0] + "</xdr:col>");	// 1-based column
    			sb.append("<xdr:colOff>" + bounds[1]+ "</xdr:colOff>");
    			sb.append("<xdr:row>" + bounds[2] + "</xdr:row>");
    			sb.append("<xdr:rowOff>" + bounds[3] + "</xdr:rowOff>");
    		sb.append("</xdr:from>");	sb.append("\r\n");
    		// bottom right coords
    		sb.append("<xdr:to>");
    			sb.append("<xdr:col>" + bounds[4] + "</xdr:col>");		// 1-based column
    			sb.append("<xdr:colOff>" + bounds[5] + "</xdr:colOff>");
    			sb.append("<xdr:row>" + bounds[6] + "</xdr:row>");
    			sb.append("<xdr:rowOff>" + bounds[7] + "</xdr:rowOff>");
    		sb.append("</xdr:to>");		sb.append("\r\n");
    		// Picture details - req. child elements= nvPicPr (non-visual picture properties), blipFill (links to image), spPr (shape properties)
    		sb.append("<xdr:pic>");		sb.append("\r\n");
    			sb.append("<xdr:nvPicPr>");
    	        	sb.append("<xdr:cNvPr id=\"" + this.getMsodrawing().getSPID() + "\"");
    	        	sb.append(" name=\"" + this.getImageName() + "\"");
    	        	sb.append(" descr=\"" + this.getShapeName() + "\"/>");
    	        	sb.append("<xdr:cNvPicPr>");
    	        	sb.append("<a:picLocks noChangeAspect=\"1\"  noChangeArrowheads=\"1\"/>");
    	        	sb.append("</xdr:cNvPicPr>");
    	        sb.append("</xdr:nvPicPr>");	sb.append("\r\n");
    	        // Picture relationship Id and relationship to the package
    	        sb.append("<xdr:blipFill>");
    	        	sb.append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId" + rId + "\"/>"); 
    	        	sb.append("<a:srcRect/>");			//If the picture is cropped, these details are stored in the <srcRect/> element
    	        	sb.append("<a:stretch><a:fillRect/></a:stretch>");
    	        sb.append("</xdr:blipFill>");	sb.append("\r\n");
    	        // shape properties 
    	        if (imagesp!=null) 
    	        	sb.append(imagesp.getOOXML());
    	        else { // default basic 
	    	        sb.append("<xdr:spPr bwMode=\"auto\">");    	        
    	        	sb.append("<a:xfrm>");	
	    	        int x=this.getX()*EMU, y=this.getY()*EMU, cx= this.getWidth()*EMU, cy= this.getHeight()*EMU;
	    	        sb.append("<a:off x=\"" + x + "\" y=\"" + y + "\"/>");		// offsets= location
	    	        sb.append("<a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>");	// extents= size of bounding box enclosing pic in EMUs
	    	        sb.append("</a:xfrm>");    	        
	    	        sb.append("<a:prstGeom prst=\"rect\">");		// preset geometry, appears necessary for images ...
	    	        sb.append("<a:avLst/>");
	    	        sb.append("</a:prstGeom>");
	    	        sb.append("</xdr:spPr>");	sb.append("\r\n");
    	        }
    	    sb.append("</xdr:pic>");	sb.append("\r\n");
    	    sb.append("<xdr:clientData/>");	sb.append("\r\n");
    	sb.append("</xdr:twoCellAnchor>");	
    	return sb.toString();
*/
	}

	/**
	 * return the OOXML shape property for this image
	 *
	 * @return
	 */
	public SpPr getSpPr()
	{
		return imagesp;
	}

	/**
	 * define the OOXML shape property for this image from an existing spPr element
	 */
	public void setSpPr( SpPr sp )
	{
		imagesp = sp;
//    	imagesp.setNS("xdr");
	}

	/**
	 * specify how to resize or move upon edit OOXML specific
	 *
	 * @param editMovement
	 */
	public void setEditMovement( String editMovement )
	{
		this.editMovement = editMovement;
	}

}