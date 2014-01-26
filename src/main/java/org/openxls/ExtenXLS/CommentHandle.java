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
package org.openxls.ExtenXLS;

import org.openxls.formats.XLS.Note;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <pre>
 * CommentHandle allows for manipulation of the Note or Comment feature of Excel
 *
 * In order to create CommentHandles programatically use the methods in WorkSheetHandle or CellHandle
 *
 * </pre>
 */
public class CommentHandle implements Handle
{
	private static final Logger log = LoggerFactory.getLogger( CommentHandle.class );
	private Note note;

	/**
	 * Creates a new CommentHandle object
	 * <br>For internal use only
	 *
	 * @param n
	 */
	protected CommentHandle( Note n )
	{
		note = n;
	}

	/**
	 * Returns the text of the Note (Comment)
	 *
	 * @return String value or text of Note
	 */
	public String getCommentText()
	{
		if( note != null )
		{
			return note.getText();
		}
		return null;
	}

	/**
	 * Sets the text of the Note (Comment).
	 * <p>The text may contain embedded formatting information as follows:
	 * <br>< font specifics>text segment< font specifics for next segment>text segment...
	 * <br>where font specifics can be one or more of (all are optional):
	 * <ul>b - bold
	 * <br>i - italic
	 * <br>s - strikethru
	 * <br>u - underlined
	 * <br>f="" - font name surrounded by quotes e.g. "Tahoma"
	 * <br>sz="" - font size in points surrounded by quotes e.g. "10"
	 * <br>Each option must be delimited by ;'s
	 * <br>For Example:
	 * <br><ul>"&#60;f=\"Tahoma\";b;sz=\"16\">Note: &#60;f=\"Cambria\";sz=\"12\">This is an important point"</ul>
	 * To reset to the default font, input an empty format: <> e.g.:
	 * <br><ul>"&#60;b;i;sz=\"8\">Note:<>This is an important comment"</ul>
	 *
	 * @param text - String text of Note
	 */
	public void setCommentText( String text )
	{
		if( note != null )
		{
			try
			{
				note.setText( text );
			}
			catch( IllegalArgumentException e )
			{
				log.error( e.toString() );
			}
		}
	}

	/**
	 * returns the author of this Note (Comment) if set
	 *
	 * @return String author
	 */
	public String getAuthor()
	{
		if( note != null )
		{
			return note.getAuthor();
		}
		return null;
	}

	/**
	 * sets the author of this Note (Comment)
	 *
	 * @param author
	 */
	public void setAuthor( String author )
	{
		if( note != null )
		{
			note.setAuthor( author );
		}
	}

	/**
	 * Removes or deletes this Note (Comment) from the worksheet
	 */
	public void remove()
	{
		note.getSheet().removeNote( note );
		note = null;
	}

	/**
	 * Sets this Note (Comment) to always show, even when the attached cell loses focus
	 */
	public void show()
	{
		if( note != null )
		{
			note.setHidden( false );
		}
	}

	/**
	 * Sets this Note (Comment) to be hidden until the attached cell has focus
	 * <p/>
	 * This is the default state of note records
	 */
	public void hide()
	{
		if( note != null )
		{
			note.setHidden( true );
		}
	}

	/**
	 * Returns true if this Note (Comment) is hidden until focus
	 *
	 * @return
	 */
	public boolean getIsHidden()
	{
		if( note != null )
		{
			return note.getHidden();
		}
		return false;
	}

	/**
	 * Sets this Note (Comment) to be attached to a cell at [row, col]
	 *
	 * @param row int row number (0-based)
	 * @param col int column number (0-based)
	 */
	public void setRowCol( int row, int col )
	{
		if( note != null )
		{
			note.setRowCol( row, col );
		}
	}

	/**
	 * Returns the address this Note (Comment) is attached to
	 *
	 * @return String Cell Address
	 */
	public String getAddress()
	{
		if( note != null )
		{
			return note.getCellAddressWithSheet();
		}
		return null;
	}

	/**
	 * return the String representation of this CommentHandle
	 */
	public String toString()
	{
		if( note != null )
		{
			return note.toString();
		}
		return "Not initialized";
	}

	/**
	 * return the Row number (0-based) this Note is attached to
	 *
	 * @return 0-based row number
	 */
	public int getRowNum()
	{
		if( note != null )
		{
			return note.getRowNumber();
		}
		return -1;
	}

	/**
	 * return the Column this note is attached to
	 *
	 * @return Column number as an integer e.g. A=0, B=1 ...
	 */
	public int getColNum()
	{
		if( note != null )
		{
			return note.getColNumber();
		}
		return -1;
	}

	/**
	 * Sets the width and height of the bounding text box of the note
	 * <br>Units are in pixels
	 * <br>NOTE: the height algorithm w.r.t. varying row heights is not 100%
	 *
	 * @param width  short desired text box width in pixels
	 * @param height short desired text box height in pixels
	 */
	public void setTextBoxSize( int width, int height )
	{
		if( note != null )
		{
			note.setTextBoxWidth( (short) width );
			note.setTextBoxHeight( (short) height );
		}
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
		if( note != null )
		{
			return note.getTextBoxBounds();
		}
		return null;
	}

	/**
	 * Sets the bounds (size and position) of the Text Box for this Note
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
	 */
	public void setTextBoxBounds( short[] bounds )
	{
		if( note != null )
		{
			note.setTextBoxBounds( bounds );
		}
	}

	/**
	 * Returns the internal note record for this CommentHandle.
	 * <p/>
	 * Be aware that this note record should not be modified directly, and that this
	 * method is only for internal application use.
	 *
	 * @return Name record
	 */
	public Note getInternalNoteRec()
	{
		return note;
	}

	/**
	 * Returns the OOXML representation of this Note object
	 *
	 * @param authId 0-based author index for the author linked to this Note
	 * @return String OOMXL representation
	 */
	public String getOOXML( int authId )
	{
		StringBuffer ooxml = new StringBuffer();
		// TODO: Handle FORMATS
		ooxml.append( "<comment ref=\"" + ExcelTools.formatLocation( new int[]{
				note.getRowNumber(), note.getColNumber()
		} ) + "\" authorId=\"" + authId + "\">" );
		ooxml.append( note.getOOXML() );
		ooxml.append( "</comment>" );
		return ooxml.toString();
	}

}

