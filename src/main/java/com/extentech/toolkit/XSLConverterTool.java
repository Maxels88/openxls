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
/*
 * XSLConverterTool is a collection of methods that are of use in the xml/xsl conversion for sheetster.
 * 
 * These methods are called from XSL to populate various fields correctly
 * 
 * Created on Mar 17, 2006
 *
*/
package com.extentech.toolkit;

import com.extentech.ExtenXLS.ExcelTools;

import java.text.NumberFormat;
import java.util.Hashtable;

public class XSLConverterTool
{

	private Hashtable styles = new Hashtable();
	private String lastCell = "A1";

	/**
	 * Gets a date format pattern based off of an ifmt.  This applies only to built in dates.
	 *
	 * @param ifmt from XF record
	 * @return date format pattern
	 */
	public String getDateFormatPattern( String ifmt )
	{
		if( ifmt.equals( "14" ) )
		{
			return "m/d/yy";
		}
		if( ifmt.equals( "15" ) )
		{
			return "d-mmm-yy";
		}
		if( ifmt.equals( "16" ) )
		{
			return "d-mmm";
		}
		if( ifmt.equals( "17" ) )
		{
			return "mmm-yy";
		}
		if( ifmt.equals( "22" ) )
		{
			return "m/d/yy h:mm";
		}
		return "m/d/yy";

	}

	/**
	 * Gets a calendar format pattern based off of an ifmt.  This applies only to built in dates.
	 *
	 * @param ifmt from XF record
	 * @return calendar format pattern
	 */
	public String getJsCalendarFormatPattern( String ifmt )
	{
		if( ifmt.equals( "14" ) )
		{
			return "%m/%d/%Y";
		}
		if( ifmt.equals( "15" ) )
		{
			return "%d-%b-%y";
		}
		if( ifmt.equals( "16" ) )
		{
			return "%d-%b";
		}
		if( ifmt.equals( "17" ) )
		{
			return "%m-%y";
		}
		if( ifmt.equals( "22" ) )
		{
			return "%m/%d/%Y %h:%M";
		}
		return "%m/%d/%Y";
	}

	/**
	 * Returns a formatted currency string based off the local format and the string format
	 * passed in.
	 * TODO: implement formatting patterns
	 *
	 * @param fmt
	 * @return
	 */
	public String getCurrencyFormat( String fmt, String value )
	{
		try
		{
			NumberFormat nf = NumberFormat.getCurrencyInstance();
			Double d = new Double( value );
			String retStr = nf.format( d );
			return retStr;
		}
		catch( NumberFormatException e )
		{
			return value;
		}
	}

	/**
	 * Get a style based of a style ID.   If the style does not yet exist, create a new one, and
	 * add it to the hashtable of styles
	 *
	 * @param styleId
	 * @return Style
	 */
	private Style getStyle( String styleId )
	{
		Object o = styles.get( styleId );
		if( o != null )
		{
			return (Style) o;
		}
		Style thisStyle = new Style( styleId );
		styles.put( styleId, thisStyle );
		return thisStyle;
	}

	/**
	 * returns a String populated with cell data for missing cells since the last cell read.
	 * Requires the first and last cells to exist for a row.
	 *
	 * @return html fragment
	 */
	public String getPreviousCellData( String sheet, String currCell, String colspan )
	{
		StringBuffer returnString = new StringBuffer();
		int colSpan = 1;
		if( (colspan != null) && (colspan != "") )
		{
			colSpan = Integer.parseInt( colspan );
		}
		int[] newCell = ExcelTools.getRowColFromString( currCell );
		int[] oldCell = ExcelTools.getRowColFromString( lastCell );
		int newCol = newCell[1];
		int oldCol = oldCell[1] + 1;
		for(; oldCol < newCol; oldCol++ )
		{
			String newAddress = ExcelTools.getAlphaVal( oldCol ) + (oldCell[0] + 1);
			returnString.append( getEmptyCellHTML( sheet, newAddress ) );
		}
		newCell[1] = newCol + colSpan - 1;
		currCell = ExcelTools.formatLocation( newCell );
		lastCell = currCell;
		return returnString.toString();
	}

	private String getEmptyCellHTML( String sheet, String address )
	{
		return "<td id=\"" + address + "\" > </td>";
	}

	/**
	 * ****************************************************************************************************
	 * **********************  DELEGATING METHODS ***********************************************
	 * *****************************************************************************************************
	 */
	public void setStyleColor( String styleId, String color )
	{
		Style thisStyle = getStyle( styleId );
		thisStyle.setColor( color );
	}

	public String getStyleColor( String styleId )
	{
		Style thisStyle = getStyle( styleId );
		return thisStyle.getColor();
	}

	public void setIsDate( String styleId, String isDate )
	{
		Style thisStyle = getStyle( styleId );
		thisStyle.setIsDate( isDate );
	}

	public String getIsDate( String styleId )
	{
		Style thisStyle = getStyle( styleId );
		String s = thisStyle.getIsDate();
		if( s == null )
		{
			return "0";
		}
		return s;
	}

	public void setIsCurrency( String styleId, String isCurrency )
	{
		if( isCurrency.equals( "" ) )
		{
			isCurrency = "0";
		}
		Style thisStyle = getStyle( styleId );
		thisStyle.setIsCurrency( isCurrency );
	}

	public String getIsCurrency( String styleId )
	{
		Style thisStyle = getStyle( styleId );
		String s = thisStyle.getIsCurrency();
		if( s == null )
		{
			return "0";
		}
		return s;
	}

	public void setFormatId( String styleId, String formatId )
	{
		Style thisStyle = getStyle( styleId );
		thisStyle.setFormatId( formatId );
	}

	public String getFormatId( String styleId )
	{
		Style thisStyle = getStyle( styleId );
		String s = thisStyle.getFormatId();
		if( s == null )
		{
			return "0";
		}
		return s;
	}

	public String getFormatPattern( String styleId )
	{
		Style thisStyle = getStyle( styleId );
		return thisStyle.getFormatPattern();
	}

	public void setFormatPattern( String styleId, String pattern )
	{
		Style thisStyle = getStyle( styleId );
		thisStyle.setFormatPattern( pattern );
	}

	/**
	 * Style holds style information about a certain style in the xsl spreadsheet
	 */
	private class Style
	{
		private String ID;
		private String formatId;
		private String color = "";
		private String test;
		private String fontFamily;
		private String fontSize;
		private String fontWeight;
		private String fontColor;
		private String textAlign;
		private String isDate;
		private String isCurrency;
		private String formatPattern;

		public String getFormatPattern()
		{
			return formatPattern;
		}

		public void setFormatPattern( String pattern )
		{
			formatPattern = pattern;
		}

		protected Style( String id )
		{
			ID = id;
		}

		public String getColor()
		{
			return color;
		}

		public void setColor( String color )
		{
			this.color = color;
		}

		public String getFontColor()
		{
			return fontColor;
		}

		public void setFontColor( String fontColor )
		{
			this.fontColor = fontColor;
		}

		public String getFontFamily()
		{
			return fontFamily;
		}

		public void setFontFamily( String fontFamily )
		{
			this.fontFamily = fontFamily;
		}

		public String getFontSize()
		{
			return fontSize;
		}

		public void setFontSize( String fontSize )
		{
			this.fontSize = fontSize;
		}

		public String getFontWeight()
		{
			return fontWeight;
		}

		public void setFontWeight( String fontWeight )
		{
			this.fontWeight = fontWeight;
		}

		public String getID()
		{
			return ID;
		}

		public void setID( String id )
		{
			ID = id;
		}

		public String getFormatId()
		{
			return formatId;
		}

		public void setFormatId( String formatId )
		{
			this.formatId = formatId;
		}

		public String getTextAlign()
		{
			return textAlign;
		}

		public void setTextAlign( String textAlign )
		{
			this.textAlign = textAlign;
		}

		public String getIsCurrency()
		{
			return isCurrency;
		}

		public void setIsCurrency( String isCurrency )
		{
			this.isCurrency = isCurrency;
		}

		public String getIsDate()
		{
			return isDate;
		}

		public void setIsDate( String isDate )
		{
			this.isDate = isDate;
		}

	}
}
