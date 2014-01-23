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

import com.extentech.formats.OOXML.Fill;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Cf;
import com.extentech.formats.XLS.Condfmt;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.FormatConstantsImpl;
import com.extentech.formats.XLS.Xf;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.StringTool;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Provides methods for querying and changing cell formatting information. Cell
 * formating includes fonts, borders, text alignment, background colors, cell
 * protection (locking), etc.
 * <p/>
 * The mutator methods of a FormatHandle object directly and immediately change
 * the formatting of the cell(s) to which it is applied. Under no circumstances
 * will a FormatHandle alter the formatting of cells to which it is not applied.
 * The list of cells to which a FormatHandle applies may be queried with the
 * {@link #getCells} method. Additional cells may be added with {@link #addCell}
 * and related methods.
 * <p/>
 * A FormatHandle may also be obtained for a row or a column. In such a case the
 * FormatHandle sets the default formats for the row or column. The default
 * format for a row or column is the format which appears when no other formats
 * for the cell are specified, such as for newly created cells.
 */
public class FormatHandle implements Handle, FormatConstants
{
	private static final Logger log = LoggerFactory.getLogger( FormatHandle.class );
	private CompatibleVector mycells = new CompatibleVector(); // all the Cells sharing this format e.g. CellRange
	private Xf myxf;
	private int xfe;
	// see Xf.usedCount instead	private boolean canModify = false;
	private ColHandle mycol = null;
	private RowHandle myrow = null;
	private boolean writeImmediate = false;
	boolean underlined = false;

	private com.extentech.formats.XLS.WorkBook wkbook;
	private com.extentech.ExtenXLS.WorkBook wbh;

	public static final Map<String, String> numericFormatMap;
	public static final Map<String, String> currencyFormatMap;
	public static final Map<String, String> dateFormatMap;

	static
	{
		Map<String, String> formats = new HashMap<>();
		for( String[] formatArr : FormatConstants.NUMERIC_FORMATS )
		{
			if( formatArr.length == 3 )
			{
				formats.put( formatArr[0].toLowerCase(), formatArr[2] );
			}
		}
		numericFormatMap = Collections.unmodifiableMap( formats );

		formats = new HashMap<>();
		for( String[] formatArr : FormatConstants.CURRENCY_FORMATS )
		{
			if( formatArr.length == 3 )
			{
				formats.put( formatArr[0].toLowerCase(), formatArr[2] );
			}
		}
		currencyFormatMap = Collections.unmodifiableMap( formats );

		formats = new HashMap<>();
		for( String[] formatArr : FormatConstants.DATE_FORMATS )
		{
			if( formatArr.length == 3 )
			{
				formats.put( formatArr[0].toLowerCase(), formatArr[2] );
			}
			else
			{
				// only 2 elements in date format string array: pattern and hex id
				formats.put( formatArr[0].toLowerCase(), formatArr[0] );
			}
		}
		dateFormatMap = Collections.unmodifiableMap( formats );
	}

	/**
	 * Nullary constructor for use in bean context. <b>This is not part of the
	 * public API and should not be called.</b>
	 */
	public FormatHandle()
	{
	}

	/**
	 * Constructs a FormatHandle for the given WorkBook and format record.
	 * <b>This is not part of the public API and should not be called.</b>
	 */
	public FormatHandle( com.extentech.ExtenXLS.WorkBook book, Xf xfr )
	{
		myxf = xfr;
		xfe = myxf.getIdx();
		wbh = book;
		if( book != null )
		{
			wkbook = book.getWorkBook();
		}
		else
		{
			if( xfr.getWorkBook() != null )
			{
				wkbook = xfr.getWorkBook();
			}
			else
			{
				log.error( "FormatHandle constructed with null WorkBook." );
			}
		}
	}

	/**
	 * Constructs a FormatHandle for the given WorkBook's default format.
	 */
	public FormatHandle( com.extentech.ExtenXLS.WorkBook book )
	{
		this( book, book.getWorkBook().getDefaultIxfe() );
	}

	/**
	 * Constructs a FormatHandle for the given format ID and WorkBook.
	 * This is useful for creating a FormatHandle with the same parameters as
	 * a cell that does not refer to the cell. For example:
	 * <pre>
	 * CellHandle cell = &lt;get cell here&gt;;
	 * FormatHandle format = new FormatHandle(
	 *         cell.getWorkBook(), cell.getFormatId() );
	 * </pre>
	 *
	 * @param book  the WorkBook from which the format should be retrieved
	 * @param xfnum the ID of the format
	 */
	public FormatHandle( com.extentech.formats.XLS.WorkBook book, int xfnum )
	{
		wkbook = book;
		if( (xfnum > -1) && (xfnum < wkbook.getNumXfs()) )
		{
			myxf = wkbook.getXf( xfnum );
			xfe = myxf.getIdx();
		}
		if( myxf == null )
		{ // add new xf if necessary
			myxf = duplicateXf( null );
		}
		myxf.getFont(); // will set to default (0th) font if not already set
	}

	/** Constructs a FormatHandle for the given format ID and WorkBook.
	 * This is useful for creating a FormatHandle with the same parameters as
	 * a cell that does not refer to the cell. For example:
	 * <pre>
	 * CellHandle cell = &lt;get cell here&gt;;
	 * FormatHandle format = new FormatHandle(
	 *         cell.getWorkBook(), cell.getFormatId() );
	 * </pre>
	 *
	 * @param book
	 *            the WorkBook from which the format should be retrieved
	 * @param xfnum
	 *            the ID of the format public FormatHandle(WorkBook book, int
	 *            xfnum ,int x) { this(book.getWorkBook(),xfnum); wbh = book; }
	 */

	/**
	 * Constructs a FormatHandle for the given format index and WorkBook.
	 * <b>This is not part of the public API and should not be called.</b>
	 */
	/*
	 * This constructor is just used from XML parsing, due to some dedupe
	 * errors. possibly errors in .xlsx too?
	 */
	protected FormatHandle( com.extentech.ExtenXLS.WorkBook book, int xfnum, boolean dedupe )
	{
		this( book, xfnum );
		writeImmediate = dedupe;
	}

	/**
	 * Constructs a dummy FormatHandle for the given conditional format. <b>This
	 * is not part of the public API and should not be called.</b>
	 * <p/>
	 * This unique flavor of FormatHandle is only used to display the formatting
	 * values of a Conditional format record.
	 * <p/>
	 * Creates a dummy Xf to store values, otherwise has no effect on the
	 * WorkBook record stream which are manipulated through CellHandle.
	 *
	 * @param book containing the conditional formats
	 * @param the  index to the conditional format in the book collection
	 */
	protected FormatHandle( Condfmt cx, com.extentech.ExtenXLS.WorkBook book, int xfnum, CellHandle c )
	{
		cx.setFormatHandle( this );
		xfe = xfnum;
		wbh = book;
		wkbook = book.getWorkBook();
		if( c == null )
		{
			// ok, this is a horrible hack, as its only correct if the top left
			// cell of the range has the same background format
			// as the cell a user is hitting. Lame, but i've been handed this at
			// the last moment and am patching things. weak effort guys.
			try
			{
				int[] rc = cx.getEncompassingRange();
				c = new CellHandle( cx.getSheet().getCell( rc[0], rc[1] ), book );
			}
			catch( Exception e )
			{
			}
			;
		}
		myxf = duplicateXf( book.getWorkBook().getXf( c.getFormatId() ) );

		// set the format from the cf
		List lx = cx.getRules();
		Iterator itx = lx.iterator();
		while( itx.hasNext() )
		{
			Cf format = (Cf) itx.next();
			this.updateFromCF( format, book );
		}
	}

	/**
	 * updates this format handle via a Cf rule
	 *
	 * @param cf   Cf rule
	 * @param book workbook
	 */
	protected void updateFromCF( Cf cf, com.extentech.ExtenXLS.WorkBook book )
	{
		// border colors
		Color[] clr = cf.getBorderColors();
		if( clr != null )
		{
			setBorderColors( clr );
		}

		// line style
		int[] xs = cf.getBorderStyles();
		if( xs != null )
		{
			this.setBorderLineStyle( xs );
		}

		/*
		 * // cf.getBorderSizes() int[] b = cf.getBorderSizes(); if(b!=null)
		 * setBorderLineStyle(b);
		 */
		// cf.getFont()
		Font f = cf.getFont();
		if( f != null )
		{
			setFont( f );
		}
		else
		{
			this.setFontHeight( 180 ); // why????????
		}

		if( cf.getFontItalic() )
		{
			this.setItalic( true );
		}

		if( cf.getFontStriken() )
		{
			this.setStricken( true );
		}

		int fsup = cf.getFontEscapement();
		// super/sub (0 = none, 1 = super, 2 = sub)
		if( fsup > -1 )
		{
			this.setScript( fsup );
		}

		// handle underlines
		int us = cf.getFontUnderlineStyle();

		if( us > -1 )
		{
			this.setUnderlineStyle( us );
			this.setUnderlined( true );
		}

		// number cf
		if( cf.getFormatPattern() != null )
		{
			setFormatPattern( cf.getFormatPattern() );
		}

		if( cf.getFill() != null )
		{
			this.setFill( cf.getFill() );
		}
		else
		{
			int fill = cf.getPatternFillStyle(); // Now -1 is a valid entry:
			if( fill > -1 )
			{
				/* If the fill style is solid: When solid is specified, the
				foreground color (fgColor) is the only color rendered,
				even when a background color (bgColor) is also
				specified. */
				int bg = cf.getPatternFillColorBack();
				int fg = cf.getPatternFillColor();
				this.setFill( fill, fg, bg );
			}
			else
			{
				int fg = cf.getForegroundColor();
				if( fg > -1 )
				{
					setForegroundColor( fg );
				}
			}
		}
	}

	/**
	 * Creates a FormatHandle for the given cell. <b>This is not part of the
	 * public API and should not be called.</b> Customers should use
	 * {@link CellHandle#getFormatHandle()} instead.
	 */
	public FormatHandle( CellHandle c )
	{
		wkbook = c.getCell().getWorkBook();
		if( c.getCell().getXfRec() != null )
		{
			myxf = c.getCell().getXfRec();
			xfe = myxf.getIdx(); // update the pointer - 20071010 KSC
		}
		else
		{ // ?? create new
			// 20090512 KSC: Shigeo NPE error formaterror646694, create new
			// rather than outputting warning
			// Logger.logWarn("No XF for cell " + c.toString());
			myxf = wkbook.getXf( c.getCell().getIxfe() );
			xfe = myxf.getIdx();
		}
		// 20101201 KSC: only add to cache when adding xf's addToCache();
	}

	/**
	 * overrides the equals method to perform equality based on format
	 * properties ------------------------------------------------------------
	 *
	 * @param Object another - the FormatHandle to compare with this FormatHandle
	 * @return true if this FormatHandle equals another
	 */
	public boolean equals( Object another )
	{
		return another.toString().equals( toString() );
	}

	/**
	 * Locks the cell attached to this FormatHandle for editing (makes
	 * read-only) lock cell and make read-only
	 *
	 * @param boolean locked - true if cells should be locked for this
	 *                FormatHandle
	 */
	public void setLocked( boolean locked )
	{
		Xf xf = cloneXf( myxf );
		xf.setLocked( locked );
		updateXf( xf );
	}

	/**
	 * sets the cell attached to this FormatHandle to hide or show formula
	 * strings;
	 *
	 * @param boolean b- true if formulas should be hidden for this FormatHandle
	 */
	public void setFormulaHidden( boolean b )
	{
		Xf xf = cloneXf( myxf );
		xf.setFormulaHidden( b );
		updateXf( xf );
	}

	/**
	 * returns whether this FormatHandle is set to hide formula strings
	 *
	 * @return true if the formula strings are hidden, false otherwise
	 */
	public boolean isFormulaHidden()
	{
		return myxf.isFormulaHidden();
	}

	/**
	 * returns whether this Format Handle specifies that cells are locked for
	 * changing
	 *
	 * @return true if cells are locked
	 */
	public boolean isLocked()
	{
		return myxf.isLocked();
	}

	/**
	 * provides a mapping between Excel formats and Java formats <br>
	 * see: <a
	 * href="tutorial">http://java.sun.com/docs/books/tutorial/i18n/format
	 * /decimalFormat.html</a>
	 * <p/>
	 * Note there are slight Excel-specific differences in the format strings
	 * returned. Several numeric and currency formats in excel have different
	 * formatting for postive and negative numbers. In these cases, the java
	 * format string is split by semicolons and may contain text [Red] which is
	 * to specify the negative number should be displayed in red. Remove this
	 * from the string before passing into the Format class;
	 * <p/>
	 * <pre>
	 *         G  Era designator  Text  AD
	 *         y  Year  Year  1996; 96
	 *         M  Month in year  Month  July; Jul; 07
	 *         w  Week in year  Number  27
	 *         W  Week in month  Number  2
	 *         D  Day in year  Number  189
	 *         d  Day in month  Number  10
	 *         F  Day of week in month  Number  2
	 *         E  Day in week  Text  Tuesday; Tue
	 *         a  Am/pm marker  Text  PM
	 *         H  Hour in day (0-23)  Number  0
	 *         k  Hour in day (1-24)  Number  24
	 *         K  Hour in am/pm (0-11)  Number  0
	 *         h  Hour in am/pm (1-12)  Number  12
	 *         m  Minute in hour  Number  30
	 *         s  Second in minute  Number  55
	 *         S  Millisecond  Number  978
	 *         z  Time zone  General time zone  Pacific Standard Time; PST; GMT-08:00
	 *         Z  Time zone  RFC 822 time zone  -0800
	 * </pre>
	 *
	 * @return String the formatting pattern for the cell
	 */
	public String getJavaFormatString()
	{
		String pat = getFormatPattern();
		if( pat == null )
		{
			return null;
		}

		// toLowerCase is a simplistic way to implement the case insensitivity
		// of the pattern tokens. It could cause issues with string literals.
		Object patty = convertFormatString( pat.toLowerCase() );
		if( patty != null )
		{
			patty = StringTool.qualifyPatternString( patty.toString() );
		}
		if( myxf.isDatePattern() )
		{
			if( patty != null )
			{
				return (String) patty;
			}
			return "M/d/yy h:mm";
		}
		if( patty != null )
		{
			return (String) patty;
		}
		/*
		 * If we reached here, we don't have a mapping for this particular
		 * format. Send a warning to the system then make sure the pattern we
		 * are sending back is valid. Many excel patterns have 4 patterns
		 * separated by semicolons. We only can pass 2 into the formatter
		 * (positive and negative). This usually works out to be the first two
		 * patterns in the string.
		 */
		pat = StringTool.qualifyPatternString( pat.toString() );
		int firstParens = pat.indexOf( ";" );
		if( firstParens != -1 )
		{
			int secondParens = pat.indexOf( ";", firstParens + 1 );
			if( secondParens != -1 )
			{
				pat = pat.substring( 0, secondParens );
			}
			else
			{ // yet another hackaround -jm
				pat = pat.substring( firstParens + 2, pat.length() - 1 );
			}
		}
		return pat;
	}

	/**
	 * converts an Excel-style format string to a Java Format string.
	 *
	 * @param String pattern - Excel Format String
	 * @return String that can be used with the Java Format classes.
	 * @see getJavaFormatString
	 */
	public static String convertFormatString( String pattern )
	{
		String ret = numericFormatMap.get( pattern );
		if( ret != null )
		{
			return ret;
		}
		ret = currencyFormatMap.get( pattern );
		if( ret != null )
		{
			return ret;
		}
		ret = dateFormatMap.get( pattern );
		if( ret != null )
		{
			return ret;
		}
		return null;
	}

	/*
	 * FIXME: Border Issues (marker)
	 * 
	 * The methods to set individual border line styles do not follow a
	 * consistent naming convention with the rest of the border methods.
	 * 
	 * There are no methods for getting or setting the inside borders. I don't
	 * know whether this is a problem or a design choice. It depends on whether
	 * the inside borders are stored in the format record or as separate cell
	 * formats.
	 * 
	 * There are no methods to get the diagonal borders. There is a method to
	 * set the diagonal border line style, but it affects both diagonals.
	 */

	/**
	 * sets the border color for all borders (top, left, bottom and right) from
	 * a Color array
	 * <p/>
	 * NOTE: this setting will affect every cell which refers to this
	 * FormatHandle
	 *
	 * @param java .awt.Color array - 4-element array of desired border colors
	 *             [T, L, B, R]
	 */
	public void setBorderColors( Color[] bordercolors )
	{
		Xf xf = cloneXf( myxf );
		if( bordercolors[0] != null )
		{
			xf.setTopBorderColor( getColorInt( bordercolors[0] ) );
		}
		if( bordercolors[1] != null )
		{
			xf.setLeftBorderColor( getColorInt( bordercolors[1] ) );
		}
		if( bordercolors[2] != null )
		{
			xf.setBottomBorderColor( getColorInt( bordercolors[2] ) );
		}
		if( bordercolors[3] != null )
		{
			xf.setRightBorderColor( getColorInt( bordercolors[3] ) );
		}
		updateXf( xf );
	}

	/**
	 * remove borders for this format
	 */
	public void removeBorders()
	{
		Xf xf = cloneXf( myxf );
		xf.removeBorders();
		updateXf( xf );
	}

	/**
	 * sets the border color for all borders (top, left, bottom and right) from
	 * an int array containing color constants
	 * <p/>
	 * <p/>
	 * NOTE: this setting will affect every cell which refers to this
	 * FormatHandle
	 *
	 * @param int[] bordercolors - 4-element array of desired border color
	 *              constants [T, L, B, R]
	 * @see FormatHandle.COLOR_* constants
	 */
	public void setBorderColors( int[] bordercolors )
	{
		Xf xf = cloneXf( myxf );
		xf.setTopBorderColor( bordercolors[0] );
		xf.setLeftBorderColor( bordercolors[1] );
		xf.setBottomBorderColor( bordercolors[2] );
		xf.setRightBorderColor( bordercolors[3] );
		updateXf( xf );

	}

	/**
	 * set the border color for all borders (top, left, bottom, and right) to
	 * one color via color constant
	 * <p/>
	 * <p/>
	 * NOTE: this setting will affect every cell which refers to this
	 * FormatHandle
	 *
	 * @param int x - color constant which represents the color to set all
	 *            border sides
	 * @see FormatHandle.COLOR_* constants
	 */
	public void setBorderColor( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setRightBorderColor( (short) x );
		xf.setLeftBorderColor( (short) x );
		xf.setTopBorderColor( (short) x );
		xf.setBottomBorderColor( (short) x );
		updateXf( xf );

	}

	/**
	 * set the border color for all borders (top, left, bottom, and right) to
	 * one java.awt.Color
	 * <p/>
	 * NOTE: this setting will affect every cell which refers to this
	 * FormatHandle
	 *
	 * @param Color col - color to set all border sides
	 */
	public void setBorderColor( Color col )
	{
		Xf xf = cloneXf( myxf );
		short x = (short) getColorInt( col );
		xf.setRightBorderColor( x );
		xf.setLeftBorderColor( x );
		xf.setTopBorderColor( x );
		xf.setBottomBorderColor( x );
		updateXf( xf );

	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderRightColor( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setRightBorderColor( (short) x );
		updateXf( xf );
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderRightColor( Color x )
	{
		Xf xf = cloneXf( myxf );
		xf.setRightBorderColor( (short) getColorInt( x ) );
		updateXf( xf );
	}

	/**
	 * Get the Right border color
	 *
	 * @return color constant
	 */
	public Color getBorderRightColor()
	{
		if( myxf.getRightBorderLineStyle() == 0 )
		{
			return null;
		}
		int x = myxf.getRightBorderColor();
		if( x < this.getWorkBook().colorTable.length )
		{
			return this.getWorkBook().getColorTable()[x];
		}
		return this.getWorkBook().getColorTable()[0]; // black i'm afraid
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderLeftColor( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setLeftBorderColor( (short) x );
		updateXf( xf );
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderLeftColor( Color x )
	{
		Xf xf = cloneXf( myxf );
		xf.setLeftBorderColor( (short) getColorInt( x ) );
		updateXf( xf );
	}

	/**
	 * returns Border Colors of Cell ie: top, left, bottom, right
	 * <p/>
	 * returns null or 1 color for each of 4 sides
	 * <p/>
	 * 1,1,1,1 represents a border all around the cell 1,1,0,0 represents on the
	 * top left edge of the cell
	 *
	 * @return int array representing Cell borders
	 */
	public Color[] getBorderColors()
	{
		Color[] colors = new Color[4];
		colors[0] = getBorderTopColor();
		colors[1] = getBorderLeftColor();
		colors[2] = getBorderBottomColor();
		colors[3] = getBorderRightColor();
		return colors;
	}

	/**
	 * Get the Left border color
	 *
	 * @return color constant
	 */
	public Color getBorderLeftColor()
	{
		if( myxf.getLeftBorderLineStyle() == 0 )
		{
			return null;
		}
		return this.getWorkBook().getColorTable()[myxf.getLeftBorderColor()];
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderTopColor( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setTopBorderColor( x );
		updateXf( xf );
	}

	/**
	 * Get the Top border color
	 *
	 * @return color constant
	 */
	public Color getBorderTopColor()
	{
		if( myxf.getTopBorderLineStyle() == 0 )
		{
			return null;
		}
		int xt = myxf.getTopBorderColor();
		if( xt > this.getWorkBook().getColorTable().length ) // guards
		{
			xt = 0;
		}
		return this.getWorkBook().getColorTable()[xt];
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderTopColor( Color x )
	{
		Xf xf = cloneXf( myxf );
		xf.setTopBorderColor( getColorInt( x ) );
		updateXf( xf );
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderBottomColor( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBottomBorderColor( x );
		updateXf( xf );
	}

	/**
	 * Set the top border color
	 *
	 * @param color constant
	 */
	public void setBorderBottomColor( Color x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBottomBorderColor( getColorInt( x ) );
		updateXf( xf );
	}

	/**
	 * Returns true if the value should be red due to a combination of a format
	 * pattern and a negative number
	 *
	 * @return
	 */
	public boolean isRedWhenNegative()
	{
		String pattern = myxf.getFormatPattern();
		if( pattern.indexOf( "Red" ) > -1 )
		{
			return true;
		}
		return false;
	}

	/**
	 * Get the Right border color
	 *
	 * @return color constant
	 */
	public Color getBorderBottomColor()
	{
		if( myxf.getBottomBorderLineStyle() == 0 )
		{
			return null;
		}
		int x = myxf.getBottomBorderColor();
		if( x < this.getWorkBook().getColorTable().length )
		{
			return this.getWorkBook().getColorTable()[x];
		}
		return this.getWorkBook().getColorTable()[0]; // black i'm afraid

	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setTopBorderLineStyle( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setTopBorderLineStyle( (short) x );
		updateXf( xf );
	}

	/**
	 * Get the border line style
	 *
	 * @return line style constant
	 */
	public int getTopBorderLineStyle()
	{
		return myxf.getTopBorderLineStyle();
	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setBottomBorderLineStyle( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBottomBorderLineStyle( (short) x );
		updateXf( xf );
	}

	/**
	 * Get the border line style
	 *
	 * @return line style constant
	 */
	public int getBottomBorderLineStyle()
	{
		return myxf.getBottomBorderLineStyle();
	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setLeftBorderLineStyle( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setLeftBorderLineStyle( (short) x );
		updateXf( xf );
	}

	/**
	 * Get the border line style
	 *
	 * @return line style constant
	 */
	public int getLeftBorderLineStyle()
	{
		return myxf.getLeftBorderLineStyle();
	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setRightBorderLineStyle( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setRightBorderLineStyle( (short) x );
		updateXf( xf );
	}

	/**
	 * Get the border line style
	 *
	 * @return line style constant
	 */
	public int getRightBorderLineStyle()
	{
		return myxf.getRightBorderLineStyle();
	}

	/**
	 * Sets the border line style using static BORDER_ shorts within
	 * FormatHandle
	 *
	 * @param line style constant
	 */
	public void setBorderLineStyle( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBorderLineStyle( (short) x );
		updateXf( xf );
	}

	/**
	 * set border line styles via array of ints representing border styles
	 * order= left, right, top, bottom, [diagonal]
	 *
	 * @param b int[]
	 */
	public void setBorderLineStyle( int[] b )
	{
		myxf.setAllBorderLineStyles( b );
	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setBorderLineStyle( short x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBorderLineStyle( x );
		updateXf( xf );
	}

	/**
	 * return the 5 border lines styles (l, r, t, b, diag)
	 *
	 * @return
	 */
	public int[] getAllBorderLineStyles()
	{
		int[] ret = new int[5];
		ret[0] = myxf.getLeftBorderLineStyle();
		ret[1] = myxf.getRightBorderLineStyle();
		ret[2] = myxf.getTopBorderLineStyle();
		ret[3] = myxf.getBottomBorderLineStyle();
		ret[4] = myxf.getDiagBorderLineStyle();
		return ret;
	}

	/**
	 * return the 5 border line colors (l, r, t, b, diag)
	 *
	 * @return
	 */
	public int[] getAllBorderColors()
	{
		int[] ret = new int[5];
		ret[0] = myxf.getLeftBorderColor();
		ret[1] = myxf.getRightBorderColor();
		ret[2] = myxf.getTopBorderColor();
		ret[3] = myxf.getBottomBorderColor();
		ret[4] = myxf.getDiagBorderColor();
		return ret;
	}

	/**
	 * Set the border line style
	 *
	 * @param line style constant
	 */
	public void setBorderDiagonal( int x )
	{
		Xf xf = cloneXf( myxf );
		xf.setBorderDiag( x );
		updateXf( xf );
	}

	/**
	 * Set a column handle on this format handle, so all changes applied to this
	 * format will be applied to the entire column
	 *
	 * @param c
	 */
	public void setColHandle( ColHandle c )
	{
		mycol = c;
	}

	/**
	 * Set a row handle on this format handle, so all changes applied to this
	 * format will be applied to the entire row
	 *
	 * @param c
	 */
	public void setRowHandle( RowHandle c )
	{
		myrow = c;
	}

	/**
	 * Create a copy of this FormatHandle with its own Xf
	 *
	 * @return the copied FormatHandle
	 */
	@Override
	public Object clone()
	{
		FormatHandle ret = null;
		if( wbh == null )
		{ // who knew???
			wkbook = myxf.getWorkBook(); // Changed to myxf since myfont is no
			// longer
			ret = new FormatHandle();
			ret.myxf = myxf;
			ret.xfe = myxf.getIdx();
			ret.wkbook = wkbook;
		}
		else
		{
			ret = new FormatHandle( wbh, myxf ); // no need to duplicate it - just
			// use all formatting of
			// original xf
		}

		return ret;
	}

	public String toString()
	{
		return myxf.toString();
	}

	public FormatHandle( com.extentech.ExtenXLS.WorkBook book, String fontname, int fontstyle, int fontsize )
	{
		this( book );
		setFont( fontname, fontstyle, fontsize );
	}

	/**
	 * Jan 27, 2011
	 *
	 * @param workBook
	 * @param i
	 */
	public FormatHandle( com.extentech.ExtenXLS.WorkBook workBook, int i )
	{
		this( workBook.getWorkBook(), i );
		wbh = workBook;
	}

	/**
	 * /** Set the weight of the font in 1/20 point units 100-1000 range. 400 is
	 * normal, 700 is bold.
	 *
	 * @param wt
	 */
	public void setFontWeight( int wt )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setFontWeight( wt );
		updateFont( f );
	}

	/**
	 * Get the font weight the weight of the font is in 1/20 point units
	 *
	 * @return
	 */
	public int getFontWeight()
	{
		return myxf.getFont().getFontWeight();
	}

	/**
	 * Set the Font for this Format.
	 * <p/>
	 * As adding a new Font and format increases the file size, try using this
	 * once for each distinct font used in the file, then use 'setFormatId' to
	 * share this font with other Cells.
	 * <p/>
	 * Roughly matches the functionality of the java.awt.Font class. Currently
	 * the style parameter is only useful for bold/normal weights. Italics,
	 * underlines, etc must be modified elsewhere.
	 * <p/>
	 * Note that in order to maintain java.awt.Font compatibility for
	 * bold/normal styles, defaults for weight/style have been mapped to 0 =
	 * normal (excel 200 weight) and 1 = bold (excel 700 weight)
	 *
	 * @param String font name
	 * @param int    font style
	 * @param int    font size
	 */
	public void setFont( Font f )
	{
		setXFToFont( f );
	}

	/**
	 * Set new font to XF, handling duplication and caching ...
	 *
	 * @param Font f
	 */
	private void setXFToFont( Font f )
	{
		int fti = addFontIfNecessary( f );
		if( myxf == null )
		{ // shouldn't!!
			myxf = duplicateXf( null );
			myxf.setFont( fti );
		}
		else if( myxf.getIfnt() != fti )
		{ // if not using font already,
			// duplicate xf and set to new font
			Xf xf = cloneXf( myxf );
			xf.setFont( fti );
			updateXf( xf );
		}

	}

	/**
	 * Sets this format handle to a font
	 *
	 * @param fn  font name e.g. 'Arial'
	 * @param stl font style either Font.PLAIN or Font.BOLD
	 * @param sz  font size or height in 1/20 point units
	 */
	public void setFont( String fn, int stl, double sz )
	{
		sz *= 20;
		if( stl == 0 )
		{
			stl = 200;
		}
		if( stl == 1 )
		{
			stl = 700;
		}
		Font f = new Font( fn, stl, (int) sz );
		setXFToFont( f );
	}

	/**
	 * add font to record streamer if cant find exact font already in there
	 *
	 * @param f
	 * @return
	 */
	private int addFontIfNecessary( Font f )
	{
		if( wkbook == null )
		{
			log.error( "AddFontIfNecessary: workbook is null" );
			return -1;
		}
		int fti = wkbook.getFontIdx( f );
		// don't use the built-ins.
		if( fti == 3 )
		{
			fti = 0; // use initial default font instead of last ...
		}
		if( fti == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 ); // flag to insert
			fti = wkbook.insertFont( f ) + 1;
		}
		else
		{
			f.setIdx( fti );
		}
		return fti;
	}

	/**
	 * adds a font to the global font store only if exact font is not already
	 * present
	 *
	 * @param f  Font
	 * @param bk WorkBookHandle
	 */
	public static int addFont( Font f, WorkBookHandle bk )
	{
		if( bk == null )
		{
			log.error( "addFont: workbook is null" );
		}
		int fti = bk.getWorkBook().getFontIdx( f );
		// if (fti > 3) {// don't use the built-ins.
		if( fti == 3 )
		{
			fti = 0; // use initial default font instead of last ... 20070827
		}
		// KSC
		if( fti == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 ); // flag to insert
			fti = bk.getWorkBook().insertFont( f ) + 1;
		}
		else
		{
			f.setIdx( fti );
		}
		return fti;
	}

	/**
	 * Adds a font internally to the workbook
	 *
	 * @param f
	 *
	 *            private void addFont(Font f) {
	 *
	 *            }
	 */

	/**
	 * Apply this Format to a Range of Cells
	 *
	 * @param CellRange to apply the format to
	 */
	public void addCellRange( CellRange cr )
	{
		CellHandle[] crcells = cr.getCells();
		for( CellHandle crcell : crcells )
		{
			addCell( crcell );
		}
	}

	/**
	 * Apply this Format to a Range of Cells
	 *
	 * @param CellHandle array to apply the format to
	 */
	public void addCellArray( CellHandle[] crcells )
	{
		for( CellHandle crcell : crcells )
		{
			addCell( crcell );
		}
	}

	/**
	 * add a Cell to this FormatHandle thus applying the Format to the Cell
	 *
	 * @param CellHandle to apply the format to
	 */
	public void addCell( CellHandle c )
	{
		c.setFormatHandle( this );
	}

	/**
	 * Add a List of Cells to this FormatHandle
	 *
	 * @param cx
	 */
	public void addCells( List cx )
	{
		Iterator itx = cx.iterator();
		while( itx.hasNext() )
		{
			addCell( (BiffRec) itx.next() );
			mycells.add( itx.next() );
		}
	}

	void addCell( BiffRec c )
	{
		if( myxf != null )
		{
			c.setXFRecord( myxf.getIdx() );
		}
		else
		{
			log.error( "FormatHandle.addCell() - You MUST call setFont() to initialize the FormatHandle's font before adding Cells." );
		}
		mycells.add( c );
	}

	/**
	 * Applies the format to a cell without establishing a relationship. The
	 * format represented by this <code>FormatHandle</code> will be applied to
	 * the cell but it will not be updated with any future changes. If you want
	 * that behavior use {@link #addCell(CellHandle) addCell} instead.
	 */
	public void stamp( CellHandle cell )
	{
		cell.setFormatId( xfe );
	}

	/**
	 * Applies the format to a cell range without establishing a relationship.
	 * The format represented by this <code>FormatHandle</code> will be applied
	 * to the cells but they will not be updated with any future changes. If you
	 * want that behavior use {@link #addCell(CellHandle) addCell} instead.
	 */
	public void stamp( CellRange range )
	{
		try
		{
			range.setFormatID( xfe );
		}
		catch( Exception e )
		{
			// This can't actually happen
		}
	}

	/**
	 * set the Background Pattern for this Format
	 *
	 * @param int Excel color constant
	 */
	public void setBackgroundPattern( int t )
	{
		Xf xf = cloneXf( myxf );
		// 20080103 KSC: handle solid (=filled) backgrounds, in which Excel
		// switches fg and bg colors (!!!)
		if( t != FormatConstants.PATTERN_FILLED )
		{
			xf.setPattern( t );
		}
		else
		{
			int bg = xf.getBackgroundColor();
			xf.setBackgroundSolid();
			xf.setForeColor( bg, null );
		}
		updateXf( xf );
	}

	/**
	 * returns whether this Format is formatted as a Date
	 *
	 * @return boolean true if this Format is formatted as a Date
	 */
	public boolean isDate()
	{
		if( myxf == null )
		{
			return false;
		}
		return myxf.isDatePattern();
	}

	/**
	 * returns whether this Format is formatted as a Currency
	 *
	 * @return boolean true if this Format is formatted as a currency
	 */
	public boolean isCurrency()
	{
		if( myxf == null )
		{
			return false;
		}
		return myxf.isCurrencyPattern();
	}

	/**
	 * set the underline style for this font
	 *
	 * @param int u underline style one of the Font.STYLE_UNDERLINE constants
	 */
	public void setUnderlineStyle( int u )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setUnderlineStyle( (byte) u );
		updateFont( f );
	}

	/**
	 * super/sub (0 = none, 1 = super, 2 = sub)
	 *
	 * @param int script type for Format Font
	 */
	public void setScript( int ss )
	{
		if( ss > 2 )
		{
			ss = 2; // deal with invalid numbers
		}
		if( ss < 0 )
		{
			ss = 0; // deal with invalid numbers
		}
		myxf.getFont().setScript( ss );
	}

	/**
	 * set the Font Color for this Format via indexed color constant
	 *
	 * @param int Excel color constant
	 */
	public void setFontColor( int t )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setColor( t );
		updateFont( f );
	}

	/**
	 * set the Font Color for this Format
	 *
	 * @param AWT Color color constant
	 */
	public void setFontColor( Color colr )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setColor( colr );
		updateFont( f );
	}

	/**
	 * sets the Font color for this Format via web Hex String
	 *
	 * @param clr
	 */
	public void setFontColor( String clr )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setColor( clr );
		updateFont( f );
	}

	/**
	 * Get the Font foreground (text) color as a java.awt.Color
	 *
	 * @return
	 */
	public Color getFontColor()
	{
		return myxf.getFont().getColorAsColor();
	}

	public String getFontColorAsHex()
	{
		return myxf.getFont().getColorAsHex();
	}

	/**
	 * set the Foreground Color for this Format NOTE: Foreground color = the
	 * CELL BACKGROUND color color for all patterns and Background color= the
	 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
	 * color=CELL BACKGROUND color and Background Color=64 (white).
	 *
	 * @param int Excel color constant
	 */
	public void setForegroundColor( int t )
	{
		Xf xf = cloneXf( myxf );
		xf.setForeColor( t, null );
		updateXf( xf );
	}

	/**
	 * set the foreground Color for this Format NOTE: Foreground color = the
	 * CELL BACKGROUND color color for all patterns and Background color= the
	 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
	 * color=CELL BACKGROUND color and Background Color=64 (white).
	 *
	 * @param AWT Color constant
	 */
	public void setForegroundColor( Color colr )
	{
		Xf xf = cloneXf( myxf );
		int clrz = getColorInt( colr );
		xf.setForeColor( clrz, colr );
		updateXf( xf );
	}

	/**
	 * set the background color for this Format NOTE: Foreground color = the
	 * CELL BACKGROUND color color for all patterns and Background color= the
	 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
	 * color=CELL BACKGROUND color and Background Color=64 (white).
	 *
	 * @param t Excel color constant
	 */
	public void setBackgroundColor( int t )
	{
		Xf xf = cloneXf( myxf );
		if( xf.getFillPattern() == Xf.PATTERN_SOLID )
		{
			xf.setForeColor( t, null );
		}
		else
		{
			xf.setBackColor( t, null );
		}

		updateXf( xf );
	}

	/**
	 * set the Cell Background Color for this Format
	 * <p/>
	 * NOTE: Foreground color = the CELL BACKGROUND color color for all patterns
	 * and Background color= the PATTERN color for all patterns != Solid For
	 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
	 * Color=64 (white).
	 *
	 * @param awt Color color constant
	 */
	public void setCellBackgroundColor( Color colr )
	{
		int clrz = getColorInt( colr );
		setCellBackgroundColor( clrz );

	}

	/**
	 * makes the cell a solid pattern background if no pattern was already
	 * present NOTE: Foreground color = the CELL BACKGROUND color color for all
	 * patterns and Background color= the PATTERN color for all patterns !=
	 * Solid For PATTERN_SOLID, Foreground color=CELL BACKGROUND color and
	 * Background Color=64 (white).
	 *
	 * @param int Excel color constant
	 */
	public void setCellBackgroundColor( int t )
	{
		Xf xf = cloneXf( myxf );

		if( xf.getFillPattern() == 0 )
		{
			xf.setBackgroundSolid();
		}

		if( xf.getFillPattern() == Xf.PATTERN_SOLID )
		{
			xf.setForeColor( t, null );
		}
		else
		{
			xf.setBackColor( t, null );
		}
		updateXf( xf );
	}

	/**
	 * sets this fill pattern from an existing OOXML (2007v) fill element
	 *
	 * @param f
	 */
	protected void setFill( Fill f )
	{
		Xf xf = cloneXf( myxf );
		xf.setFill( f );
		updateXf( xf );
	}

	/**
	 * sets the fill for this format handle if fill==Xf.PATTERN_SOLID then fg is
	 * the PATTERN color i.e the CELL BG COLOR
	 *
	 * @param fillpattern
	 * @param fg
	 * @param bg
	 */
	public void setFill( int fillpattern, int fg, int bg )
	{
		Xf xf = cloneXf( myxf );

		xf.setPattern( fillpattern );

		/**
		 * If the fill style is solid: When solid is specified, the foreground
		 * color (fgColor) is the only color rendered, even when a background
		 * color (bgColor) is also specified.
		 */
		if( xf.getFillPattern() == Xf.PATTERN_SOLID )
		{ // is reversed
			xf.setForeColor( bg, null );
			xf.setBackColor( 64, null );
		}
		else
		{
			/**
			 * or cell fills with patterns specified, then the cell fill color
			 * is specified by the bgColor element
			 */
			xf.setForeColor( fg, null );
			xf.setBackColor( bg, null );
		}
		updateXf( xf );
	}

	/**
	 * Get the Pattern Background Color for this Format Pattern
	 *
	 * @return the Excel color constant
	 */
	public int getBackgroundColor()
	{
		return myxf.getBackgroundColor();
	}

	/**
	 * get the Pattern Background Color for this Format Pattern as a hex string
	 *
	 * @return Hex Color String
	 */
	public String getBackgroundColorAsHex()
	{
		return myxf.getBackgroundColorHEX();
	}

	/**
	 * get the Pattern Background Color for this Format Pattern as an awt.Color
	 *
	 * @return background Color
	 */
	public java.awt.Color getBackgroundColorAsColor()
	{
		return HexStringToColor( this.getBackgroundColorAsHex() );
	}

	/**
	 * returns the foreground color setting regardless of format pattern (which
	 * can switch fg and bg)
	 *
	 * @return
	 */
	public int getTrueForegroundColor()
	{
		return myxf.getForegroundColor();    // 20080814 KSC: getForegroundColor() does the swapping so use base method
	}

	/**
	 * get the Pattern Background Color for this Formatted Cell
	 * <p/>
	 * This method handles display of conditional formats for the cell
	 * <p/>
	 * checks for conditional format, then applies it if conditions are true.
	 *
	 * @return the Excel color constant
	 */
	public int getCellBackgroundColor()
	{
		int fp = getFillPattern();

		if( fp == Xf.PATTERN_SOLID )
		{
			return myxf.getForegroundColor(); // this.getForegroundColor() does
		}
		// the swapping so use base
		// method

		return myxf.getBackgroundColor();
	}

	/**
	 * get the Pattern Background Color for this Format as a Hex Color String
	 *
	 * @return Hex Color String
	 */
	public String getCellBackgroundColorAsHex()
	{
		int fp = getFillPattern();

		if( fp == Xf.PATTERN_SOLID )
		{
			return myxf.getForegroundColorHEX(); // this.getForegroundColor() does
		}
		return myxf.getBackgroundColorHEX();
	}

	/**
	 * get the Pattern Background Color for this Format as an awt.Color
	 *
	 * @return cell background color
	 */
	public java.awt.Color getCellBackgroundColorAsColor()
	{
		return HexStringToColor( this.getCellBackgroundColorAsHex() );
	}

	/**
	 * get the Background Color for this Format NOTE: Foreground color = the
	 * CELL BACKGROUND color color for all patterns and Background color= the
	 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
	 * color=CELL BACKGROUND color and Background Color=64 (white).
	 *
	 * @return the Excel color constant
	 */
	public int getForegroundColor()
	{
		// if it's SOLID pattern, fg/bg are swapped
		if( getFillPattern() == Xf.PATTERN_SOLID )
		{
			return myxf.getBackgroundColor();
		}

		return myxf.getForegroundColor();
	}

	/**
	 * get the Background Color for this Format as a Hex Color String NOTE:
	 * Foreground color = the CELL BACKGROUND color color for all patterns and
	 * Background color= the PATTERN color for all patterns != Solid For
	 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
	 * Color=64 (white).
	 *
	 * @return Hex Color String
	 */
	public String getForegroundColorAsHex()
	{
		if( getFillPattern() == Xf.PATTERN_SOLID ) // if it's SOLID pattern,
		{
			return myxf.getBackgroundColorHEX();
		}

		return myxf.getForegroundColorHEX();
	}

	/**
	 * get the Background Color for this Format as a Color NOTE:
	 * Foreground color = the CELL BACKGROUND color color for all patterns and
	 * Background color= the PATTERN color for all patterns != Solid For
	 * PATTERN_SOLID, Foreground color=CELL BACKGROUND color and Background
	 * Color=64 (white).
	 *
	 * @return Hex Color String
	 */
	public Color getForegroundColorAsColor()
	{
		return HexStringToColor( this.getForegroundColorAsHex() );
	}

	/**
	 * set the Background Color for this Format NOTE: Foreground color = the
	 * CELL BACKGROUND color color for all patterns and Background color= the
	 * PATTERN color for all patterns != Solid For PATTERN_SOLID, Foreground
	 * color=CELL BACKGROUND color and Background Color=64 (white).
	 *
	 * @param awt Color color constant
	 */
	public void setBackgroundColor( Color colr )
	{
		Xf xf = cloneXf( myxf );
		int clrz = getColorInt( colr );
		xf.setBackColor( clrz, colr );
		updateXf( xf );
	}

	/**
	 * Set the format handle to use standard bold text
	 *
	 * @param boolean isBold
	 */
	public void setBold( boolean isBold )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setBold( isBold );
		updateFont( f );
	}

	/**
	 * Get if this format is bold or not
	 *
	 * @return boolean whether the cell font format is bold
	 */
	public boolean getIsBold()
	{
		return myxf.getFont().getIsBold();
	}

	/**
	 * Return an int representing the underline style
	 * <p/>
	 * These map to the STYLE_UNDERLINE static integers *
	 *
	 * @return int underline style
	 */
	public int getUnderlineStyle()
	{
		return myxf.getFont().getUnderlineStyle();
	}

	/**
	 * Get the font height in points
	 *
	 * @return font height
	 */
	public double getFontHeightInPoints()
	{
		return getFont().getFontHeightInPoints();
	}

	/**
	 * Returns the Font's height in 1/20th point increment
	 *
	 * @return font height
	 */
	public int getFontHeight()
	{
		return getFont().getFontHeight();
	}

	/**
	 * Set the Font's height in 1/20th point increment
	 *
	 * @param new font height
	 */
	public void setFontHeight( int fontHeight )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setFontHeight( fontHeight );
		updateFont( f );
	}

	/**
	 * Returns the Font's name
	 *
	 * @return font name
	 */
	public String getFontName()
	{
		return getFont().getFontName();
	}

	/**
	 * Set the Font's name
	 * <p/>
	 * To be valid, this font name must be available on the client system.
	 *
	 * @param font name
	 */
	public void setFontName( String fontName )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setFontName( fontName );
		updateFont( f );
	}

	/**
	 * Determine if the format handle refers to a font stricken out
	 *
	 * @return boolean representing if the FormatHandle is striking out a cell.
	 */
	public boolean getStricken()
	{
		if( myxf.getFont() == null )
		{
			return false;
		}
		return myxf.getFont().getStricken();
	}

	/**
	 * Set if the format handle is stricken out
	 *
	 * @param isStricken boolean representing if the formatted cell should be stricken
	 *                   out.
	 */
	public void setStricken( boolean isStricken )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setStricken( isStricken );
		updateFont( f );
	}

	/**
	 * Get if the font is italic
	 *
	 * @return boolean representing if the formatted cell is italic.
	 */
	public boolean getItalic()
	{
		if( myxf.getFont() == null )
		{
			return false;
		}
		return myxf.getFont().getItalic();
	}

	/**
	 * Set if the font is italic
	 *
	 * @param isItalic boolean representing if the formatted cell should be italic.
	 */
	public void setItalic( boolean isItalic )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setItalic( isItalic );
		updateFont( f );
	}

	/**
	 * Get if the font is underlined
	 *
	 * @return boolean representing if the formatted cell is underlined.
	 */
	public boolean getUnderlined()
	{
		if( myxf.getFont() == null )
		{
			return underlined;
		}
		return myxf.getFont().getUnderlined();
	}

	/**
	 * Set underline attribute on the font
	 *
	 * @param isUnderlined boolean representing if the formatted cell should be
	 *                     underlined
	 */
	public void setUnderlined( boolean isUnderlined )
	{
		Font f = cloneFont( myxf.getFont() );
		f.setUnderlined( isUnderlined );
		updateFont( f );

	}

	/**
	 * Get if the font is bold
	 *
	 * @return boolean representing if the formatted cell is bold
	 */
	public boolean getBold()
	{
		return myxf.getFont().getBold();
	}

	/**
	 * returns the existing font record for this Format
	 * <p/>
	 * Font is an internal record and should not be accessed by end users
	 *
	 * @return the XLS Font record associated with this Format
	 */
	public com.extentech.formats.XLS.Font getFont()
	{
		// should this be a protected method?
		if( myxf != null ) // shouldn't!
		{
			return myxf.getFont();
		}
		return null;
	}

	/**
	 * Sets the number format pattern for this format. All Excel built-in number
	 * formats are supported. Custom formats will not be applied by ExtenXLS
	 * (e.g. {@link CellHandle#getFormattedStringVal}) but they will be written
	 * correctly to the output file. For more information on number format
	 * patterns see <a href="http://support.microsoft.com/kb/264372">Microsoft
	 * KB264372</a>.
	 *
	 * @param pat the Excel number format pattern to apply
	 * @see FormatConstantsImpl#getBuiltinFormats
	 */
	public void setFormatPattern( String pat )
	{
		Xf xf = cloneXf( myxf );
		xf.setFormatPattern( pat );
		updateXf( xf );
	}

	public void setPattern( int pat )
	{
		Xf xf = cloneXf( myxf );
		xf.setPattern( pat );
		updateXf( xf );
	}

	/**
	 * The format ID allows setting of the format for a Cell without adding it
	 * to the Format's Cell Collection.
	 * <p/>
	 * Use to decrease memory requirements when dealing with large collections
	 * of cells.
	 * <p/>
	 * Usage:
	 * <p/>
	 * mycell.setFormatId(myformat.getFormatId());
	 *
	 * @return the format ID for this Format
	 */
	public int getFormatId()
	{
		return xfe;
	}

	/**
	 * Gets the number format pattern for this format, if set. For more
	 * information on number format patterns see <a
	 * href="http://support.microsoft.com/kb/264372">Microsoft KB264372</a>.
	 *
	 * @return the Excel number format pattern for this cell or
	 * <code>null</code> if none is applied
	 */
	public String getFormatPattern()
	{
		return myxf.getFormatPattern();
	}

	/**
	 * get the fill pattern for this format
	 */
	public int getFillPattern()
	{
		return myxf.getFillPattern();
	}

	/**
	 * returns the index of the Color within the Colortable
	 *
	 * @param the index of the color in the colortable
	 * @return the color
	 */
	public static Color getColor( int col )
	{
		if( (col > -1) && (col < FormatHandle.COLORTABLE.length) )
		{
			return FormatHandle.COLORTABLE[col];
		}
		return FormatHandle.COLORTABLE[0];
	}

	/**
	 * returns the index of the Color within the Colortable
	 *
	 * @param col the color
	 * @return the index of the color in the colortable
	 */
	public static int getColorInt( Color col )
	{
		for( int i = 0; i < FormatHandle.COLORTABLE.length; i++ )
		{
			if( col.equals( FormatHandle.COLORTABLE[i] ) )
			{
				return i;
			}
		}
		int R = col.getRed();
		int G = col.getGreen();
		int B = col.getBlue();
		int colorMatch = -1;
		double colorDiff = Integer.MAX_VALUE;
		for( int i = 0; i < FormatHandle.COLORTABLE.length; i++ )
		{
			double curDif = Math.pow( R - FormatHandle.COLORTABLE[i].getRed(), 2 ) + Math.pow( G - FormatHandle.COLORTABLE[i].getGreen(),
			                                                                                   2 ) + Math.pow( B - FormatHandle.COLORTABLE[i]
					.getBlue(), 2 );
			if( curDif < colorDiff )
			{
				colorDiff = curDif;
				colorMatch = i;
			}
		}
		return colorMatch;
	}

	/**
	 * Sets the internal format to the FormatHandle. For internal use only, not
	 * supported.
	 */
	public void setXf( Xf xrec )
	{
		myxf = xrec;
	}

	/**
	 * Set the horizontal alignment for this FormatHandle
	 *
	 * @param align - an int representing the alignment. Please review the
	 *              FormatHandle.ALIGN*** static int's
	 */
	public void setHorizontalAlignment( int align )
	{
		Xf xf = cloneXf( myxf );
		xf.setHorizontalAlignment( align );
		updateXf( xf );
	}

	/**
	 * Returns an int representing the current horizontal alignment in this
	 * FormatHandle. These values are mapped to the FormatHandle.ALIGN*** static
	 * int's
	 */
	public int getHorizontalAlignment()
	{
		return myxf.getHorizontalAlignment();
	}

	/**
	 * set indent (1= 3 spaces)
	 *
	 * @param indent
	 */
	public void setIndent( int indent )
	{
		Xf xf = cloneXf( myxf );
		xf.setIndent( indent );
		updateXf( xf );
	}

	/**
	 * increase or decrease the precision of this numeric or curerncy format pattern
	 * <br>If the format pattern is "General", converts to a basic number pattern
	 * <br>If the precision is already 0 and !increase, this method does nothing
	 *
	 * @param increase true if increase the precsion (number of decimals to display)
	 */
	public void adjustPrecision( boolean increase )
	{
		// TODO:  if decimal is contained within quotes ...
		String pat = getFormatPattern();
		if( pat.equals( "General" ) && increase )
		{
			pat = "0.0";    // the most basic numeric pattern
			this.setFormatPattern( pat );
			return;
		}

		try
		{
			// split pattern and deal with positive, negative, zero and text separately
			// for each, find decimal place and increment/decrement; if not found, find last digit placeholder
			String[] pats = pat.split( ";" );
			String newPat = "";
			for( int i = 0; i < pats.length; i++ )
			{
				if( i > 0 )
				{
					newPat += ';';
				}
				int z = pats[i].indexOf( '.' );    // position of decimal
				boolean foundit = false;
				if( z != -1 )
				{    // found decimal place
					z++;
					for(; z < pats[i].length(); z++ )
					{
						char c = pats[i].charAt( z );
						if( ((c == '0') || (c == '#') || (c == '?')) )    // numeric placeholders
						{
							foundit = true;
						}
						else if( foundit && !((c == '0') || (c == '#') || (c == '?')) )    // numeric placeholders. if hit last one, either inc or dec
						{
							break;
						}
					}
					if( increase )
					{
						newPat += new StringBuffer( pats[i] ).insert( z, "0" )
						                                     .toString();    //pats[i].substring(0, z) + "0" + pats[i].substring(z+1);
					}
					else
					{
						if( pats[i].charAt( z - 2 ) != '.' )
						{
							newPat += new StringBuffer( pats[i] ).deleteCharAt( z - 1 )
							                                     .toString();    //  .pats[i].substring(0, z-1) + pats[i].substring(z+1);
						}
						else
						{
							newPat += new StringBuffer( pats[i] ).delete( z - 2, z ).toString();
						}
					}

				}
				else if( increase )
				{ // no decimal yet.  If decrease, ignore.  if increase, add
					z = pats[i].length() - 1;
					for(; z >= 0; z-- )
					{
						char c = pats[i].charAt( z );
						if( ((c == '0') || (c == '#') || (c == '?')) )
						{    // found last numeric placeholder
							foundit = true;
							break;
						}
					}
					if( foundit )    // if had ANY numeric placeholders
					{
						newPat += new StringBuffer( pats[i] ).insert( z + 1, ".0" )
						                                     .toString();    //pats[i].substring(0, z) + ".0" + pats[i].substring(z+1);
					}
					else
					{
						newPat += pats[i];    // keep original
					}
				}
				else            // if decrease and no decimal found, leave alone
				{
					newPat += pats[i];    // leave alone
				}
			}

			//System.out.println("Old Style" + pat + ".  New Style: " + newPat + ". Increase?" +  (increase?"yes":"no"));	// KSC: TESETING: TAKE OUT WHEN DONE
			this.setFormatPattern( newPat );
		}
		catch( Exception e )
		{
			log.error( "Error setting style", e );    // KSC: TESETING: TAKE OUT WHEN DONE
		}
	}

	/**
	 * return indent (1 = 3 spaces)
	 *
	 * @return
	 */
	public int getIndent()
	{
		return myxf.getIndent();
	}

	/**
	 * sets the Right to Left Text Direction or reading order of this style
	 *
	 * @param rtl possible values:
	 *            <br>0=Context Dependent
	 *            <br>1=Left-to-Right
	 *            <br>2=Right-to-Let
	 * @param rtl possible values: <br>
	 *            0=Context Dependent <br>
	 *            1=Left-to-Right <br>
	 *            2=Right-to-Let
	 */
	public void setRightToLeftReadingOrder( int rtl )
	{
		Xf xf = cloneXf( myxf );
		xf.setRightToLeftReadingOrder( rtl );
		updateXf( xf );
	}

	/**
	 * returns true if this style is set to Right-to-Left text direction
	 * (reading order)
	 *
	 * @return
	 */
	public int getRightToLetReadingOrder()
	{
		return myxf.getRightToLeftReadingOrder();
	}

	/**
	 * Set the Vertical alignment for this FormatHandle
	 *
	 * @param align - an int representing the alignment. Please review the
	 *              FormatHandle.ALIGN*** static int's
	 */
	public void setVerticalAlignment( int align )
	{
		Xf xf = cloneXf( myxf );
		xf.setVerticalAlignment( align );
		updateXf( xf );
	}

	/**
	 * Returns an int representing the current Vertical alignment in this
	 * FormatHandle. These values are mapped to the FormatHandle.ALIGN*** static
	 * int's
	 */
	public int getVerticalAlignment()
	{
		return myxf.getVerticalAlignment();
	}

	/**
	 * DEPRECATED and non functional. Not neccesary as this occurs automatically
	 * <p/>
	 * Consolidates this Format with other identical formats in workbook
	 * <p/>
	 * <p/>
	 * There is a limit to the number of distinct formats in an Excel workbook.
	 * <p/>
	 * This method allows you to share Formats between identically formatted
	 * cells.
	 *
	 * @return
	 * @deprecated
	 */
	public FormatHandle pack()
	{
		return this; // wkbook.cache.get(this);
	}

	/**
	 * Set the workbook for this FormatHandle
	 *
	 * @param bk
	 */
	public void setWorkBook( com.extentech.formats.XLS.WorkBook bk )
	{
		wkbook = bk;
	}

	/**
	 * Set the cell wrapping behavior for this FormatHandle. Default is false
	 */
	public void setWrapText( boolean wrapit )
	{
		Xf xf = cloneXf( myxf );
		xf.setWrapText( wrapit );
		updateXf( xf );
	}

	/**
	 * Get the cell wrapping behavior for this FormatHandle. Default is false
	 */
	public boolean getWrapText()
	{
		return myxf.getWrapText();
	}

	/**
	 * Set the rotation of the cell in degrees. Values 0-90 represent rotation
	 * up, 0-90degrees. Values 91-180 represent rotation down, 0-90 degrees.
	 * Value 255 is vertical
	 *
	 * @param align - an int representing the rotation.
	 */
	public void setCellRotation( int align )
	{
		Xf xf = cloneXf( myxf );
		xf.setRotation( align );
		updateXf( xf );
	}

	/**
	 * Get the rotation of the cell. Value 0 means no rotation (horizontal
	 * text). Values 1-90 mean rotation up (couter-clockwise) by 1-90 degrees.
	 * Values 91-180 mean rotation down (clockwise) by 1-90 degrees. Value 255
	 * means vertical text.
	 */
	public int getCellRotation()
	{
		return myxf.getRotation();
	}

	/**
	 * Get a JSON representation of the format
	 *
	 * @param cr
	 * @return
	 */
	public String getJSON( int XFNum )
	{
		return getJSONObject( XFNum ).toString();
	}

	/**
	 * Get a JSON representation of the format
	 * <p/>
	 * font height is represented as HTML pt size
	 *
	 * @param cr
	 * @return
	 */
	public JSONObject getJSONObject( int XFNum )
	{
		Font myf = getFont();
		JSONObject theStyle = new JSONObject();
		try
		{
			theStyle.put( "style", XFNum );
			// handle the font
			JSONObject theFont = new JSONObject();
			theFont.put( "name", myf.getFontName() );

			// round out the font size...
			long sz = Math.round( myf.getFontHeight() / 22.0 );

			theFont.put( "size", sz ); // adjust smaller
			theFont.put( "color", colorToHexString( myf.getColorAsColor() ) );
			theFont.put( "weight", myf.getFontWeight() );
			if( getIsBold() )
			{
				theFont.put( "bold", "1" );
			}
			if( this.getUnderlined() )
			{
				theFont.put( "underline", "1" );
			}
			else
			{
				theFont.put( "underline", "0" );
			}

			if( getItalic() )
			{
				theFont.put( "italic", "1" );
			}

			theStyle.put( "font", theFont );

			// <Borders>
			JSONObject border = new JSONObject();
			if( getRightBorderLineStyle() != 0 )
			{
				JSONObject rBorder = new JSONObject();
				rBorder.put( "style", BORDER_STYLES_JSON[getRightBorderLineStyle()] );
				rBorder.put( "color", colorToHexString( getBorderRightColor() ) );
				border.put( "right", rBorder );
			}
			if( getBottomBorderLineStyle() != 0 )
			{
				JSONObject bBorder = new JSONObject();
				bBorder.put( "style", BORDER_STYLES_JSON[getBottomBorderLineStyle()] );
				bBorder.put( "color", colorToHexString( getBorderBottomColor() ) );
				border.put( "bottom", bBorder );
			}
			if( getLeftBorderLineStyle() != 0 )
			{
				JSONObject lBorder = new JSONObject();
				lBorder.put( "style", BORDER_STYLES_JSON[getLeftBorderLineStyle()] );
				lBorder.put( "color", colorToHexString( getBorderLeftColor() ) );
				border.put( "left", lBorder );
			}
			if( getTopBorderLineStyle() != 0 )
			{
				JSONObject tBorder = new JSONObject();
				tBorder.put( "style", BORDER_STYLES_JSON[getTopBorderLineStyle()] );
				tBorder.put( "color", colorToHexString( getBorderTopColor() ) );
				border.put( "top", tBorder );
			}
			theStyle.put( "borders", border );

			// <Alignment>
			JSONObject alignment = new JSONObject();
			alignment.put( "horizontal", HORIZONTAL_ALIGNMENTS[getHorizontalAlignment()] );
			alignment.put( "vertical", VERTICAL_ALIGNMENTS[getVerticalAlignment()] );
			if( getWrapText() )
			{
				alignment.put( "wrap", "1" );
			}
			theStyle.put( "alignment", alignment );

			if( getIndent() != 0 )
			{
				theStyle.put( "indent", getIndent() );
			}
			// <Interior> colors + background patterns
			if( getFillPattern() >= 0 )
			{ // KSC: added >= 0 as some conditional formats have patternFillStyle==0 even though they have a pattern block and
				JSONObject interior = new JSONObject();
				// weird black/white case
				if( myxf.getForegroundColor() == 65 )
				{
					interior.put( "color", "#FFFFFF" );
				}
				else
				{
					// KSC: use color string if it exists; create if doesn't
					interior.put( "color", myxf.getForegroundColorHEX() );
				}
				interior.put( "pattern", myxf.getFillPattern() );
				interior.put( "fg", myxf.getForegroundColor() ); // Excel-2003 Color Table index
				interior.put( "patterncolor", myxf.getBackgroundColorHEX() );
				interior.put( "bg", myxf.getBackgroundColor() ); // Excel-2003 Color Table index
				theStyle.put( "interior", interior );
			}

			if( myxf.getIfmt() != 0 )
			{ // only input user defined formats ...
				JSONObject nFormat = new JSONObject();
				String fmtpat = getFormatPattern();
				try
				{
					if( !fmtpat.equals( "General" ) )
					{
						nFormat.put( "format", fmtpat ); // convertXMLChars?
						nFormat.put( "formatid", myxf.getIfmt() );
						if( isDate() )
						{
							nFormat.put( "isdate", "1" );
						}
						if( isCurrency() )
						{
							nFormat.put( "iscurrency", "1" );
						}
						if( isRedWhenNegative() )
						{
							nFormat.put( "isrednegative", "1" );
						}
					}

				}
				catch( Exception e )
				{ // it's possible that getFormatPattern
					// returns null
				}
				theStyle.put( "numberformat", nFormat );
			}
			// <Protection>
			JSONObject protection = new JSONObject();
			try
			{
				if( this.myxf.isLocked() )
				{
					protection.put( "Protected", true );
				}
				if( this.myxf.isFormulaHidden() )
				{
					protection.put( "HideFormula", true );
				}
			}
			catch( Exception e )
			{
			}
			theStyle.put( "protection", protection );

		}
		catch( JSONException e )
		{
			log.error( "Error getting cellRange JSON: ", e );
		}
		return theStyle;
	}

	/**
	 * Returns an XML fragment representing the FormatHandle
	 */
	public String getXML( int XFNum )
	{
		return getXML( XFNum, false );
	}

	/**
	 * Returns an XML fragment representing the FormatHandle
	 *
	 * @param convertToUnicodeFont if true, font family will be changed to ArialUnicodeMS
	 *                             (standard unicode) for non-ascii fonts
	 */
	public String getXML( int XFNum, boolean convertToUnicodeFont )
	{
		Font myf = getFont();
		// <Style =main element
		StringBuffer sb = new StringBuffer( "<Style" );
		sb.append( " ID=\"s" + XFNum + "\"" );
		sb.append( ">" );
		// <Font>
		sb.append( "<Font " + myf.getXML( convertToUnicodeFont ) + "/>" );

		// <Borders>
		sb.append( "<Borders>" );
		if( getRightBorderLineStyle() != 0 )
		{
			sb.append( "		<Border" );
			sb.append( " Position=\"right\"" );
			sb.append( " LineStyle=\"" + BORDER_NAMES[getRightBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + colorToHexString( getBorderRightColor() ) + "\"" );
			sb.append( " Weight=\"" + BORDER_SIZES_HTML[getRightBorderLineStyle()] + "\"" );
			sb.append( "/>" );
		}
		if( getBottomBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"bottom\"" );
			sb.append( " LineStyle=\"" + BORDER_NAMES[getBottomBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + colorToHexString( getBorderBottomColor() ) + "\"" );
			sb.append( " Weight=\"" + BORDER_SIZES_HTML[getBottomBorderLineStyle()] + "\"" );
			sb.append( "/>" );
		}
		if( getLeftBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"left\"" );
			sb.append( " LineStyle=\"" + BORDER_NAMES[getLeftBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + colorToHexString( getBorderLeftColor() ) + "\"" );
			sb.append( " Weight=\"" + BORDER_SIZES_HTML[getLeftBorderLineStyle()] + "\"" );
			sb.append( "/>" );
		}
		if( getTopBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"top\"" );
			sb.append( " LineStyle=\"" + BORDER_NAMES[getTopBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + colorToHexString( getBorderTopColor() ) + "\"" );
			sb.append( " Weight=\"" + BORDER_SIZES_HTML[getTopBorderLineStyle()] + "\"" );
			sb.append( "/>" );
		}
		sb.append( "</Borders>" );

		// <Alignment>
		sb.append( "<Alignment" );
		sb.append( " Horizontal=\"" + HORIZONTAL_ALIGNMENTS[getHorizontalAlignment()] + "\"" );
		sb.append( " Vertical=\"" + VERTICAL_ALIGNMENTS[getVerticalAlignment()] + "\"" );
		if( getWrapText() )
		{
			sb.append( " Wrap=\"1\"" );
		}
		if( myxf.getIndent() > 0 )
		{
			sb.append( " Indent=\"" + (myxf.getIndent() * 10) + "px\"" ); // indent // # * 3 // spaces // - do // an approximate conversion
		}
		sb.append( " />" );

		// <Interior> colors + background patterns
		/*
		 * NOTE: Foreground color = the CELL BACKGROUND color color for all
		 * patterns and Background color= the PATTERN color for all patterns !=
		 * Solid For PATTERN_SOLID, Foreground color=CELL BACKGROUND color and
		 * Background Color=64 (white).
		 * 
		 * NOTE: Interior Color (html color string)===fg (color #) and
		 * PatternColor (html color string)===bg (color #) because it's too hard
		 * to interpret Excel's "automatic" concept, sometimes it need white as
		 * 1 or 9 or 64 ... very difficult to figure out how to convert back to
		 * correct color int given just a color string ...
		 */
		sb.append( "<Interior" );
		if( getFillPattern() > 0 )
		{ // 20070201 KSC: Background colors only valid if there is a fill pattern
			// 20080815 KSC: use myxf.getForeground/getBackgroundColor as FH
			// vers will switch depending on fill pattern
			// Also put Pattern element last in so re-creatng will NOT switch
			// fg/bg and create a confusing morass :)
			int fg = myxf.getForegroundColor();
			// ****************************************
			// PDf processing shouldn't output white background due to z-order and overwriting image/chart objects
			// ... possibly other uses need the white bg set ...???
			// ****************************************
			if( !((myxf.getFillPattern() == PATTERN_FILLED) && this.getWorkBook().getColorTable()[fg].equals( Color.WHITE )) )
			{
				sb.append( " Color=\"" + colorToHexString( this.getWorkBook().getColorTable()[fg] ) + "\"" + " Fg=\"" + fg + "\"" );
			}
			sb.append( " PatternColor=\"" + colorToHexString( this.getWorkBook()
			                                                      .getColorTable()[myxf.getBackgroundColor()] ) + "\"" + " Bg=\"" + myxf.getBackgroundColor() + "\"" );
			sb.append( " Pattern=\"" + myxf.getFillPattern() + "\"" );
		}
		sb.append( " />" );
		// <NumberFormat>
		if( myxf.getIfmt() != 0 )
		{ // only input user defined formats ...
			String fmtpat = getFormatPattern();
			try
			{
				sb.append( "<NumberFormat" );
				if( !fmtpat.equals( "General" ) )
				{
					sb.append( " Format=\"" + StringTool.convertXMLChars( fmtpat ) + "\"" );
					sb.append( " FormatId=\"" + myxf.getIfmt() + "\"" );
					if( isDate() )
					{
						sb.append( " IsDate=\"1\"" );
					}
					if( isCurrency() )
					{
						sb.append( " IsCurrency=\"1\"" );
					}
				}
			}
			catch( Exception e )
			{ // it's possible that getFormatPattern
				// returns null
			}
			finally
			{
				sb.append( " />" );
			}
		}
		// <Protection>
		// only input user defined formats ...
		boolean locked = this.myxf.isLocked();
		int lck = 0;
		if( locked )
		{
			lck = 1;
		}

		boolean formulahidden = this.myxf.isFormulaHidden();
		int fmlz = 0;
		if( formulahidden )
		{
			fmlz = 1;
		}

		try
		{
			sb.append( "<Protection" );
			sb.append( " Protected=\"" + lck + "\"" );
			sb.append( " HideFormula=\"" + fmlz + "\"" );
		}
		catch( Exception e )
		{
			log.warn( "FormatHandle.getXML problem with protection setting: " + e.toString() );
		}
		finally
		{
			sb.append( " />" );
		}

		sb.append( "</Style>" );
		return sb.toString();
	}

	/**
	 * Gets the Excel format ID for this format's number format pattern.
	 *
	 * @return the Excel format identifier number for the number format pattern
	 */
	public int getFormatPatternId()
	{
		return myxf.getIfmt();
	}

	/**
	 * Sets the number format pattern based on the format ID number. This method
	 * is recommended for advanced users only. In most cases you should use
	 * {@link #setFormatPattern(String)} instead.
	 *
	 * @param fmt the format ID number for the desired number format pattern
	 */
	public void setFormatPatternId( int fmt )
	{
		myxf.setFormat( (short) fmt );
	}

	/**
	 * Convert a java.awt.Color to a hex string.
	 *
	 * @return String representation of a Color
	 */
	public static String colorToHexString( Color c )
	{
		int r = c.getRed();
		int g = c.getGreen();
		int b = c.getBlue();
		String rh = Integer.toHexString( r );
		if( rh.length() < 2 )
		{
			rh = "0" + rh;
		}
		String gh = Integer.toHexString( g );
		if( gh.length() < 2 )
		{
			gh = "0" + gh;
		}
		String bh = Integer.toHexString( b );
		if( bh.length() < 2 )
		{
			bh = "0" + bh;
		}
		return ("#" + rh + gh + bh).toUpperCase();
	}

	public static short colorFONT = 0;
	public static short colorBACKGROUND = 1;
	public static short colorFOREGROUND = 2;
	public static short colorBORDER = 3;

	/**
	 * convert hex string RGB to Excel colortable int format if an exact match
	 * is not find, does color-matching to try and obtain closest match
	 *
	 * @param s
	 * @param colorType
	 * @return
	 */
	public static int HexStringToColorInt( String s, short colorType )
	{
		if( s.length() > 7 )
		{
			s = "#" + s.substring( s.length() - 6 );
		}
		if( s.indexOf( "#" ) == -1 )
		{
			s = "#" + s;
		}
		if( s.length() == 7 )
		{
			String rs = s.substring( 1, 3 );
			int r = Integer.parseInt( rs, 16 );
			String gs = s.substring( 3, 5 );
			int g = Integer.parseInt( gs, 16 );
			String bs = s.substring( 5, 7 );
			int b = Integer.parseInt( bs, 16 );
			// Handle exceptions for black, white and color indexes 9 (see
			// FormatConstants for more info)

			if( (r == 255) && (r == g) && (r == b) )
			{
				if( colorType == colorFONT )
				{
					return 9;
				}
			}

			Color c = new Color( r, g, b );
			return getColorInt( c );
		}
		return 0;
	}

	/**
	 * convert hex string RGB to Excel colortable int format if no exact match
	 * is found returns -1
	 *
	 * @param s         HTML color string (#XXXXXX) format or (#FFXXXXXX - OOXML-style
	 *                  format)
	 * @param colorType
	 * @return index into color table or -1 if not found
	 */
	public static int HexStringToColorIntExact( String s, short colorType )
	{
		if( s.length() > 7 )
		{
			s = "#" + s.substring( s.length() - 6 );
		}
		if( s.indexOf( "#" ) == -1 )
		{
			s = "#" + s;
		}
		if( s.length() == 7 )
		{
			String rs = s.substring( 1, 3 );
			int r = Integer.parseInt( rs, 16 );
			String gs = s.substring( 3, 5 );
			int g = Integer.parseInt( gs, 16 );
			String bs = s.substring( 5, 7 );
			int b = Integer.parseInt( bs, 16 );

			// Handle exceptions for black, white and color indexes 9 (see
			// FormatConstants for more info)
			if( (r == 255) && (r == g) && (r == b) )
			{
				if( colorType == colorFONT )
				{
					return 9;
				}
			}

			Color c = new Color( r, g, b );
			for( int i = 0; i < FormatHandle.COLORTABLE.length; i++ )
			{
				if( c.equals( FormatHandle.COLORTABLE[i] ) )
				{
					return i;
				}
			}
		}
		return -1; // match NOT FOUND
	}

	/**
	 * convert hex string RGB to a Color
	 *
	 * @param s web-style hex color string, including #, or OOXML-style color string, FFXXXXXX
	 * @return Color
	 * @see Color
	 */
	public static Color HexStringToColor( String s )
	{
		if( s.length() > 7 )
		{    // transform OOXML-style color strings to Web-style
			s = "#" + s.substring( s.length() - 6 );
		}
		Color c = null;
		if( s.length() == 7 )
		{
			String rs = s.substring( 1, 3 );
			int r = Integer.parseInt( rs, 16 );
			String gs = s.substring( 3, 5 );
			int g = Integer.parseInt( gs, 16 );
			String bs = s.substring( 5, 7 );
			int b = Integer.parseInt( bs, 16 );
			c = new Color( r, g, b );
		}
		else
		{
			c = new Color( 0, 0, 0 ); // default to black?
		}
		return c;
	}

	/**
	 * interpret color table special entries
	 *
	 * @param clr color index in range of 0x41-0x4F (charts) 0x40 ...
	 * @return index into color table
	 */
	public static short interpretSpecialColorIndex( int clr )
	{
		switch( clr )
		{
			case 0x0041: // Default background color. This is the window background
				// color in the sheet display and is the default
				// background color for a cell.
			case 0x004E: // Default chart background color. This is the window
				// background color in the chart display.
			case 0x0050: // WHAT IS THIS ONE????????
				return FormatConstants.COLOR_WHITE;
			case 0x0040: // Default foreground color
			case 0x004F: // Chart neutral color which is black, an RGB value of
				// (0,0,0).
			case 0x004D: // Default chart foreground color. This is the window text
				// color in the chart display.
			case 0x0051: // ToolTip text color. This is the automatic font color for
				// comments.
			case 0x7FFF: // Font automatic color. This is the window text color
				return FormatConstants.COLOR_BLACK;
			default: // 67(=0x43) ???
				return FormatConstants.COLOR_WHITE;
		}

		/*
		 * switch (icvFore) { case 0x40: // default fg color return
		 * FormatConstants.COLOR_WHITE; case 0x41: // default bg color return
		 * FormatConstants.COLOR_WHITE; case 0x4D: // default CHART fg color --
		 * INDEX SPECIFIC! return -1; // flag to map via series (bar) color
		 * defaults case 0x4E: // default CHART fg color return icvFore; case
		 * 0x4F: // chart neutral color == black return
		 * FormatConstants.COLOR_BLACK; }
		 * 
		 * switch (icvBack) { case 0x40: // default fg color return
		 * FormatConstants.COLOR_WHITE; case 0x41: // default bg color return
		 * FormatConstants.COLOR_WHITE; case 0x4D: // default CHART fg color --
		 * INDEX SPECIFIC! return -1; // flag to map via series (bar) color
		 * defaults case 0x4E: // default CHART bg color //return
		 * FormatConstants.COLOR_WHITE; // is this correct? return icvBack; case
		 * 0x4F: // chart neutral color == black return
		 * FormatConstants.COLOR_BLACK; }
		 */
	}

	public static int BorderStringToInt( String s )
	{
		for( int i = 0; i < FormatConstants.BORDER_NAMES.length; i++ )
		{
			if( FormatConstants.BORDER_NAMES[i].equals( s ) )
			{
				return i;
			}
		}
		return 0;
	}

	// 20060412 KSC: added for access
	public com.extentech.formats.XLS.WorkBook getWorkBook()
	{
		return wkbook;
	}

	/**
	 * return truth of "this Xf rec is a style xf"
	 *
	 * @return
	 */
	public boolean isStyleXf()
	{
		return myxf.isStyleXf();
	}

	/**
	 * creates a new font based on an existing one, adds to workbook recs
	 */
	private Font createNewFont( Font f )
	{
		f.setIdx( -1 ); // flag to insert anew (see Font.setWorkBook)
		wkbook.insertFont( f );
		f.setWorkBook( wkbook );
		return f;
	}

	/**
	 * create a new font based on existing font - does not add to workbook recs
	 *
	 * @param src font
	 * @return cloned font
	 */
	private Font cloneFont( Font src )
	{
		Font f = new Font();
		f.setOpcode( com.extentech.formats.XLS.XLSConstants.FONT );
		f.setData( src.getBytes() ); // use default font as basis of new font
		f.setIdx( -2 );    // avoid adding to fonts array when call setW.b. below
		f.setWorkBook( this.getWorkBook() );
		f.init();
		return f;
	}

	/**
	 * update the font for this FormatHandle, including updating the xf if
	 * necessary
	 *
	 * @param f
	 */
	private void updateFont( Font f )
	{
		int idx = wkbook.getFontIdx( f );
		if( idx == -1 )
		{ // can't find it so add new
			f = createNewFont( f );
		}
		else
		{
			f = wkbook.getFont( idx );
		}
		if( f.getIdx() != myxf.getIfnt() )
		{ // then updated the font, must create new xf to link to
			Xf xf = cloneXf( myxf );
			xf.setFont( f.getIdx() );
			updateXf( xf );
		}
	}

	/**
	 * Create or Duplicate Xf rec so can alter (pattern, font, colrs ...)
	 *
	 * @param xf xf to base off of, or null (will create new)
	 * @return new Xf
	 */
	private Xf duplicateXf( Xf xf )
	{
		int fidx = 0;
		if( xf != null )
		{
			fidx = xf.getFont().getIdx();
		}
		xf = Xf.updateXf( xf, fidx, wkbook ); // clones/creates new based upon
		// original
		xfe = xf.getIdx(); // update the pointer
// not used anymore		canModify = true; // if duplicated, it's new and unlinked (thus far)
		return xf;
	}

	/**
	 * adds all formatting represented by the sourceXf to this workbook, if not
	 * already present <br>
	 * This is used internally for transferring formats from one workbook to
	 * another
	 *
	 * @param xf - sourceXf
	 * @return ixfe of added Xf
	 */
	public int addXf( Xf sourceXf )
	{
		// must handle font first in order to create xf below
		// check to see if the font needs to be added in current workbook
		int fidx = addFontIfNecessary( sourceXf.getFont() );

		/** XF **/
		Xf localXf = FormatHandle.cloneXf( sourceXf, wkbook.getFont( fidx ), wkbook ); // clone xf so modifcations don't affect original

		/** NUMBER FORMAT **/
		String fmt = sourceXf.getFormatPattern(); // number format pattern
		if( fmt != null )
		{
			localXf.setFormatPattern( fmt ); // adds new format pattern if not // found
		}

		// now check out to see if this particular xf pattern exists; if not, add
		updateXf( localXf );
		return xfe;
	}

	/**
	 * if existing format matches, reuse. otherwise, create new Xf record and
	 * add to cache
	 *
	 * @param xf
	 */
	private void updateXf( Xf xf )
	{
		if( !myxf.toString().equals( xf.toString() ) )
		{
			if( (myxf.getUseCount() <= 1) && (xfe > 15) )
			{ // used only by one cell, OK to modify
				if( writeImmediate || (wkbook.getFormatCache().get( xf.toString() ) == null) )
				{
					// myxf hasn't been used yet; modify bytes and re-init ***
					byte[] xfbytes = xf.getBytes();
					myxf.setData( xfbytes );
					if( xf.fill != null )
					{
						myxf.fill = (com.extentech.formats.OOXML.Fill) xf.fill.cloneElement();
					}
					myxf.init();
					myxf.setFont( myxf.getIfnt() ); // set font as well ..
					wkbook.updateFormatCache( myxf ); // ensure new xf signature
					// is stored
				}
				else
				{
					if( myxf.getUseCount() > 0 )
					{
						myxf.decUseCoount();    // flag original xf that 1 less record is referencing it
					}
					myxf = (Xf) wkbook.getFormatCache().get( xf.toString() );
					xfe = myxf.getIdx(); // update the pointer
					if( xfe == -1 ) // hasn't been added to wb yet - should this ever happen???
					{
						myxf = duplicateXf( xf ); // create a duplicate and leave original
					}
					else
					{
						myxf.incUseCount();
					}
				}
			}
			else
			{ // cannot modify original - either find matching or create new
				if( myxf.getUseCount() > 0 )
				{
					myxf.decUseCoount();    // flag original xf that 1 less record is referencing it
				}
				if( wkbook.getFormatCache().get( xf.toString() ) == null )
				{ // doesn't exist yet
					myxf = duplicateXf( xf ); // create a duplicate and leave original
				}
				else
				{
					myxf = (Xf) (wkbook.getFormatCache().get( xf.toString() ));
					xfe = myxf.getIdx(); // update the pointer
					if( xfe == -1 ) // hasn't been added to the record store yet 	// - should ever happen???
					{
						myxf = duplicateXf( xf ); // create a duplicate and leave original
					}
					else
					{
						myxf.incUseCount();
					}
				}
			}

			for( Object mycell : mycells )
			{
				((BiffRec) mycell).setXFRecord( xfe ); // make sure all linked cells are updated as well
			}
			if( mycol != null )
			{
				mycol.setFormatId( xfe );
			}
			if( myrow != null )
			{
				myrow.setFormatId( xfe );
			}
		}
	}

	/**
	 * create a duplicate xf rec based on existing ...
	 *
	 * @param xf
	 * @return cloned xf
	 */
	private Xf cloneXf( Xf xf )
	{
		Xf clone = new Xf( xf.getFont().getIdx(), wkbook );
		byte[] data = xf.getBytesAt( 0, xf.getLength() - 4 );
		clone.setData( data );
		if( xf.fill != null )
		{
			clone.fill = (com.extentech.formats.OOXML.Fill) xf.fill.cloneElement();
		}
		clone.init();
		return clone;
	}

	/**
	 * static version of cloneXf
	 *
	 * @param xf
	 * @param wkbook
	 * @return
	 */
	public static Xf cloneXf( Xf xf, com.extentech.formats.XLS.WorkBook wkbook )
	{
		Xf clone = new Xf( xf.getFont().getIdx(), wkbook );
		byte[] data = xf.getBytesAt( 0, xf.getLength() - 4 );
		clone.setData( data );
		clone.init();
		return clone;
	}

	/**
	 * static version of cloneXf
	 *
	 * @param xf
	 * @param wkbook
	 * @return
	 */
	public static Xf cloneXf( Xf xf, Font f, com.extentech.formats.XLS.WorkBook wkbook )
	{
		Xf clone = new Xf( f, wkbook );
		byte[] data = xf.getBytesAt( 0, xf.getLength() - 4 );
		clone.setData( data );
		clone.setFont( f.getIdx() ); // font idx is overwritten by xf data; must reset
		clone.init();
		return clone;
	}

	/**
	 * set the pointer to the XFE or Conditional format
	 *
	 * @param xfe
	 */
	public void setFormatId( int x )
	{
		this.xfe = x;
	}

	/**
	 * clear out object references
	 */
	public void close()
	{
		mycells.clear();
		mycells = new CompatibleVector(); // all the Cells sharing this format
		myxf = null;
		mycol = null;
		myrow = null;
		wkbook = null;
		wbh = null;
	}

	@Override
	protected void finalize() throws Throwable
	{
		try
		{
			close(); // close open files
		}
		finally
		{
			super.finalize();
		}
	}

	/**
	 * For internal usage only, return the internal XF record that
	 * represents this FormatHandle
	 *
	 * @return the myxf
	 */
	protected Xf getXf()
	{
		return myxf;
	}
}