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
package com.extentech.formats.OOXML;

import com.extentech.ExtenXLS.PivotTableHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.SxStreamID;
import com.extentech.formats.XLS.Sxvd;
import com.extentech.formats.XLS.Sxview;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import org.xmlpull.v1.XmlPullParserFactory;

import java.io.InputStream;

public class PivotTableDefinition implements OOXMLElement
{

	private static final long serialVersionUID = -5070227633357072878L;
	Sxview ptview = null;

	/**
	 * NOT COMPLETED DO NOT USE
	 * parse a pivotTable OOXML element <br>
	 * Top-level attributes <li>Location information <li>Collection of fields
	 * <li>Fields on the row axis <li>Items on the row axis (specific values)
	 * <li>Fields on the column axis <li>Items on the column axis (specific
	 * values) <li>Fields on the report filter region <li>Fields in the values
	 * region <li>Style information <br>
	 * Outline of the XML for a pivotTableDefinition (sequence) <li>
	 * pivotTableDefinition <li>location <li>pivotFields <li>rowFields <li>
	 * rowItems <li>colFields <li>colItems <li>pageFields <li>dataFields <li>
	 * conditionalFormats <li>pivotTableStyleInfo
	 * <p/>
	 * <br>
	 * <p/>
	 * <pre>
	 * A PivotTable report that has more than one row field has one inner row field,
	 * 	the one closest to the data area.
	 *
	 * 	Any other row fields are outer row fields
	 *
	 * 	Items in the outermost row field are displayed only once,
	 * 	but items in the rest of the row fields are repeated as needed
	 *
	 * 	Page fields allow you to filter the entire PivotTable report to
	 * 	display data for a single item or all the items.
	 *
	 * 	Data fields provide the data values to be summarized. Usually data fields contain numbers,
	 * 	which are combined with the Sum summary function, but data fields can also contain text,
	 * 	in which case the PivotTable report uses the Count summary function.
	 *
	 * 	If a report has more than one data field, a single field button named Data
	 * 	appears in the report for access to all of the data fields.
	 * </pre>
	 *
	 * @param bk
	 * @param sheet
	 * @param ii
	 */
	public static PivotTableHandle parseOOXML( WorkBookHandle bk, /*Object cacheid, */Boundsheet sheet, InputStream ii )
	{
		Sxview ptview = null;

		try
		{
			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();
			xpp.setInput( ii, null ); // using XML 1.0 specification
			int eventType = xpp.getEventType();

			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pivotTableDefinition" ) )
					{ // get attributes
						String cacheId = "", tablename = "";
						for( int z = 0; z < xpp.getAttributeCount(); z++ )
						{
							String nm = xpp.getAttributeName( z );
							String v = xpp.getAttributeValue( z );
							if( nm.equals( "name" ) )
							{
								tablename = v;
							}
							else if( nm.equals( "cacheId" ) )
							{
								cacheId = v;
							}
							//dataOnRows
							//applyNumberFormats
							//applyBorderFormats
							//applyFontFormats
							//applyPatternFormats
							//applyAlignmentFormats
							//applyWidthHeightFormats
							//dataCaption
							//updatedVersion
							//showMemberPropertyTips
							//useAutoFormatting
							//itemPrintTitles
							//createdVersion
							//indent
							//compact
							//compactData
							//gridDropZones
						}
// KSC: TESTING!!!
						cacheId = "0";
						short cid = (short) (Integer.valueOf( cacheId ).shortValue());
						SxStreamID ptstream = bk.getWorkBook().getPivotStream( cid + 1 );
						ptview = sheet.addPivotTable( ptstream.getCellRange().toString(), bk, cid + 1, tablename );
						ptview.setDataName( "Values" );    // default
						ptview.setICache( cid );
					}
					else if( tnm.equals( "location" ) )
					{
						parseLocationOOXML( xpp, ptview, bk );
					}
					else if( tnm.equals( "pivotFields" ) )
					{ // Represents the collection of fields that appear on the PivotTable.
						parsePivotFields( xpp, ptview );
					}
					else if( tnm.equals( "pageFields" ) )
					{    // Represents the collection of items in the page or report filter region of the PivotTable.
//						short count = Integer.valueOf(xpp.getAttributeValue(0)).shortValue();
//						ptview.setCDimPg(count);
					}
					else if( tnm.equals( "pageField" ) )
					{ // count: # of pageField elements
						parsePageFieldOOXML( xpp, ptview );
					}
					else if( tnm.equals( "dataFields" ) )
					{    // Represents the collection of items in the data region of the PivotTable.
//						short count = Integer.valueOf(xpp.getAttributeValue(0)).shortValue();
//						ptview.setCDimData(count);
					}
					else if( tnm.equals( "dataField" ) )
					{ // count: # of dataField elements
						parseDataFieldOOXML( xpp, ptview );
					}
					else if( tnm.equals( "colFields" ) || tnm.equals( "rowFields" ) )
					{    // the collection of fields on the column or row axis
						parseFieldOOXML( xpp, ptview );
					}
					else if( tnm.equals( "rowItems" ) || tnm.equals( "colItems" ) )
					{        // the collection of column items or row items
						parseLineItemOOXML( xpp, ptview );
					}
					else if( tnm.equals( "formats" ) )
					{
						parseFormatsOOXML( xpp, ptview );
					}
					else if( tnm.equals( "chartFormats" ) )
					{
						parseFormatsOOXML( xpp, ptview );
					}
					else if( tnm.equals( "pivotTableStyleInfo" ) )
					{
						;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{ // go to end of file
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "PivotTableDefinition.parseOOXML: " + e.toString() );
		}
		PivotTableHandle pth = new PivotTableHandle( ptview, bk );
		return pth;
	}

	/**
	 * parses a pivotTableDefinition location firstDataCol, firstDataRow,
	 * firstHeaderRow, ref [all req], colPageCount, rowPageCount
	 *
	 * @param xpp
	 * @param pth
	 * @return
	 * @throws XmlPullParserException
	 */
	private static void parseLocationOOXML( XmlPullParser xpp, Sxview ptview, WorkBookHandle bk ) throws XmlPullParserException
	{
		try
		{
			for( int i = 0; i < xpp.getAttributeCount(); i++ )
			{
				String nm = xpp.getAttributeName( i );
				String v = xpp.getAttributeValue( i );
				if( nm.equalsIgnoreCase( "ref" ) ) // req; Specifies the first row of the actual PivotTable (NOT the data)
				{
					ptview.setLocation( v );
				}
				else if( nm.equalsIgnoreCase( "firstDataCol" ) ) // req
				// Specifies the first column of the PivotTable data, relative to the top left cell in the ref value.
				{
					ptview.setColFirstData( (short) (Integer.valueOf( v ).shortValue()) );
				}
				else if( nm.equalsIgnoreCase( "firstDataRow" ) ) // req
				// Specifies the first row of the PivotTable data, relative to the top left cell in the ref value.
				{
					ptview.setRwFirstData( (short) (Integer.valueOf( v ).shortValue()) );
				}
				else if( nm.equalsIgnoreCase( "firstHeaderRow" ) ) // req
				// Specifies the first row of the PivotTable header relative to the top left cell in the ref value.
				{
					ptview.setRwFirstHead( (short) (Integer.valueOf( v ).shortValue()) );
				}
				else if( nm.equalsIgnoreCase( "rowPageCount" ) ) // def= 0
				// Specifies the number of rows per page for this PivotTable that the filter area will occupy. By default there is a
				// single column of filter fields per page and the fields occupy as many rows as there are fields.
				{
					;
				}
				else if( nm.equalsIgnoreCase( "colPageCount" ) ) // def= 0
				// Specifies the number of columns per page for this PivotTable that the filter area will occupy. By default
				// there is a single column of filter fields per page and the fields occupy as many rows as there are fields.
				{
					;
				}
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "PivotTableHandle.parseLocation:" );
		}
	}

	/**
	 * parses the pivotFields element of the pivotTableDefinition parent
	 * <br>Represents the collection of fields that appear on the PivotTable.
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parsePivotFields( XmlPullParser xpp, Sxview ptview ) throws XmlPullParserException
	{
		try
		{
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			int fcount = 0;
			Sxvd curAxis = null;                        // up to 4 axes:  ROW, COL, PAGE or DATA
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pivotField" ) )
					{ // Represents a single field in the PivotTable. This complex type contains information about the field, including the collection of items in the field.
						curAxis = null;
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{        // TODO: HANDLE ALL ATTRIBUTES *****
							String nm = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( nm.equalsIgnoreCase( "axis" ) ) // Specifies the region of the PivotTable that this field is displayed
							{
								curAxis = ptview.addPivotFieldToAxis( axisLookup( v ), fcount++ );    // axisPage, axisRow, axisCol
							}
							else if( nm.equalsIgnoreCase( "showAll" ) )    // Specifies a boolean value that indicates whether to show all items for this field.
							{
								;                                        // A value of off, 0, or false indicates items be shown according to user specified criteria
							}
							else if( nm.equalsIgnoreCase( "defaultSubtotal" ) )    // Specifies a boolean value that indicates whether the default subtotal aggregation // function is displayed for this field.
							{
								;
							}
							else if( nm.equalsIgnoreCase( "numFmtId" ) )        // Specifies the identifier of the number format to apply to this field.
							{
								;
							}
							else if( nm.equalsIgnoreCase( "dataField" ) )
							{    // Specifies a boolean value that indicates whether this field appears in the data region of the PivotTable.
								if( v.equals( "1" ) && curAxis == null )
								{
									curAxis = ptview.addPivotFieldToAxis( axisLookup( "axisValues" ), fcount++ );        // DATA axis
								}
							}
							else if( nm.equalsIgnoreCase( "multipleItemSelectionAllowed" ) ) // Specifies a boolean value that indicates whether the field can have multiple items selected in the page field.
							{
								;
							}
							else if( nm.equalsIgnoreCase( "sortType" ) )    // ascending, descending or manual
							{
								;
							}
						}

					}
					else if( tnm.equals( "items" ) )
					{    // Represents the collection of items in a PivotTable field. The items in the collection are ordered by index. Items
						// represent the unique entries from the field in the source data.
						parsePivotItemOOXML( xpp, ptview, curAxis );
					}
					else if( tnm.equals( "pivotArea" ) )
					{ // parent= autoSortScope, which has no attributes or other children
						parsePivotAreaOOXML( xpp );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( elname ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parsePivotFields:" + e.toString() );
		}
	}

	/**
	 * parses pivot items -- Represents the collection of items in a PivotTable
	 * field. The items in the collection are ordered by index. Items represent
	 * the unique entries from the field in the source data The order in which
	 * the items are listed is the order they would appear on a particular axis
	 * (row or column, for example). parent= pivotField
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parsePivotItemOOXML( XmlPullParser xpp, Sxview ptview, Sxvd axis ) throws XmlPullParserException
	{
		try
		{
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			//int count = Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// count # of item elements
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "item" ) )
					{ // Represents a single item in PivotTable field.
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( nm.equals( "c" ) )    // Specifies a boolean value that indicates whether the approximate number of child items for this item is greater than zero.
							{
								;
							}
							else if( nm.equals( "d" ) ) // Specifies a boolean value that indicates whether this item has been expanded in the PivotTable view.
							{
								;
							}
							else if( nm.equals( "e" ) ) // Specifies a boolean value that indicates whether attribute hierarchies nested next to each other on a PivotTable row or column will offer drilling "across" each other or not.
							{
								;
							}
							else if( nm.equals( "f" ) ) // Specifies a boolean value that indicates whether this item is a calculated member.
							{
								;
							}
							else if( nm.equals( "h" ) ) // Specifies a boolean value that indicates whether the item is hidden.
							{
								;
							}
							else if( nm.equals( "m" ) ) // Specifies a boolean value that indicate whether the item has a missing value.
							{
								;
							}
							else if( nm.equals( "n" ) ) // Specifies the user caption of the item.
							{
								;
							}
							else if( nm.equals( "s" ) ) // Specifies a boolean value that indicates whether the item has a character value.
							{
								;
							}
							else if( nm.equals( "sd" ) ) // Specifies a boolean value that indicates whether the details are hidden for this item.
							{
								;
							}
							else if( nm.equals( "t" ) ) // Specifies the type of this item. A value of 'default' indicates the subtotal or total item.
							{
								if( v.equals( "default" ) )
								{
									ptview.addPivotItem( axis, 1, -1 );
								}
								else
								{
									Logger.logWarn( "PivitItem: Unknown type" );    // REMOVE WHEN TESTED
								}
							}
							else if( nm.equals( "x" ) ) // Specifies the item index in pivotFields collection in the PivotCache. Applies only non- OLAP PivotTables.
							{
								ptview.addPivotItem( axis, 0, Integer.valueOf( v ).intValue() );
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( elname ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parsePivotItemOOXML:" + e.toString() );
		}
	}

	/**
	 * parses the format element of the pivotTableDefinition.  Represents the collection of formats applied to PivotTable.
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parseFormatsOOXML( XmlPullParser xpp, Sxview ptview ) throws XmlPullParserException
	{
		try
		{
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			//int count = Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// count # of item elements
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "format" ) )
					{    // parent= formats
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							//String v = xpp.getAttributeValue(i);
							if( nm.equals( "dxfId" ) )
							{ // Specifies the identifier of the format the application is currently using for the PivotTable. Formatting information is written to the styles part. See the Styles section (§3.8) for more information on formats.

							}
							else if( nm.equals( "action" ) )
							{
								/* 	Specifies the formatting behavior for the area indicated in the pivotArea element. The
									default value for this attribute is "formatting," which indicates that the specified cells
									have some formatting applied. The format is specified in the dxfId attribute. If the
									formatting is cleared from the cells, then the value of this attribute becomes "blank."
								 */
							}
						}
					}
					else if( tnm.equals( "chartFormat" ) )
					{    // parent= chartFormats
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							//String v = xpp.getAttributeValue(i);
							if( nm.equals( "chart" ) )
							{ // Specifies the index of the chart part to which the formatting applies.
							}
							else if( nm.equals( "format" ) )
							{    // Specifies the index of the pivot format that is currently in use. This index corresponds to a dxf element in the Styles part.
							}
							else if( nm.equals( "series" ) )
							{    // Specifies a boolean value that indicates whether format applies to a series. (default=false)
							}
						}
					}
					else if( tnm.equals( "pivotArea" ) )
					{
						parsePivotAreaOOXML( xpp );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( elname ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parsePivotAreaOOXML:" + e.toString() );
		}
	}

	/**
	 * parse pivotArea element:  Rule describing a PivotTable selection (format, pivotField ...)
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parsePivotAreaOOXML( XmlPullParser xpp ) throws XmlPullParserException
	{
		try
		{
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pivotArea" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							//String v = xpp.getAttributeValue(i);
							if( nm.equals( "field" ) )    // Index of the field that this selection rule refers to.
							{
								;
							}
							else if( nm.equals( "type" ) )    // all, button, data, none, normal, origin, topRight
							{
								;
							}
							else if( nm.equals( "dataOnly" ) )    // Flag indicating whether only the data values (in the data area of the view) for an item selection are selected and does not include the item labels.
							{
								;
							}
							else if( nm.equals( "labelOnly" ) )    // Flag indicating whether only the item labels for an item selection are selected and does not include the data values (in the data area of the view).
							{
								;
							}
							else if( nm.equals( "outline" ) )    // Flag indicating whether the rule refers to an area that is in outline mode.
							{
								;
							}
							else if( nm.equals( "axis" ) )    //The region of the PivotTable to which this rule applies.
							{
								;
							}
							else if( nm.equals( "fieldPosition" ) )    // Position of the field within the axis to which this rule applies.
							{
								;
							}
							// grandRow, grandCol, cacheIndex, offset, collapsedLevelsAreSubtotals
						}
					}
					else if( tnm.equals( "references" ) )
					{ // Represents the set of selected fields and the selected items within those fields
						// count
					}
					else if( tnm.equals( "reference" ) )
					{
						/*
						 * <attribute name="field" use="optional" type="xsd:unsignedInt"/>
8 <attribute name="count" type="xsd:unsignedInt"/>
9 <attribute name="selected" type="xsd:boolean" default="true"/>
10 <attribute name="byPosition" type="xsd:boolean" default="false"/>
11 <attribute name="relative" type="xsd:boolean" default="false"/>
12 <attribute name="defaultSubtotal" type="xsd:boolean" default="false"/>
13 <attribute name="sumSubtotal" type="xsd:boolean" default="false"/>
14 <attribute name="countASubtotal" type="xsd:boolean" default="false"/>
15 <attribute name="avgSubtotal" type="xsd:boolean" default="false"/>
16 <attribute name="maxSubtotal" type="xsd:boolean" default="false"/>
17 <attribute name="minSubtotal" type="xsd:boolean" default="false"/>
18 <attribute name="productSubtotal" type="xsd:boolean" default="false"/>
19 <attribute name="countSubtotal" type="xsd:boolean" default="false"/>
20 <attribute name="stdDevSubtotal" type="xsd:boolean" default="false"/>
21 <attribute name="stdDevPSubtotal" type="xsd:boolean" default="false"/>
22 <attribute name="varSubtotal" type="xsd:boolean" default="false"/>
23 <attribute name="varPSubtotal" type="xsd:boolean" default="false"/>
						 */
					}
					else if( tnm.equals( "x" ) )
					{
						//int index= parseItemIndexOOXML(xpp);
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( elname ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parsePivotAreaOOXML:" );
		}

	}

	/**
	 * element i parent= rowItem or colItem
	 * <br>the collection of items in the row axis -- index corresponds to that in the location range
	 * <br>OR
	 * <br>the collection of column items-- index corresponds to that in the location range
	 * <br>The first <i> collection represents all item values for the first column in the column axis area
	 * <br>The first <x> in the first <i> corresponds to the first field in the columns area
	 * or
	 * <br>Represents the collection of items in the row region of the PivotTable.
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parseLineItemOOXML( XmlPullParser xpp, Sxview ptview ) throws XmlPullParserException
	{
		try
		{
			// number of row or col lines == cRw or cCol of SxVIEW
			//int linecount= Integer.valueOf(xpp.getAttributeValue(0)).intValue();	// parent element= rowItems or colItems
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			boolean isRowItems = (elname.equals( "rowItems" ));

			final Class<ITEMTYPES> enumType = ITEMTYPES.class;
			int type = 0, repeat = 0;
			short[] indexes = null;
			int nIndexes = (isRowItems) ? ptview.getCDimRw() : ptview.getCDimCol();
			int index = 0, nLines = 0;
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "i" ) )
					{ // a colItem or rowItem
						indexes = new short[nIndexes];
						nLines = nIndexes;
						index = 0;
						type = repeat = 0;
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{    // i, r, t
							String nm = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( nm.equals( "i" ) )
							{    // Specifies a zero-based index indicating the referenced data item it in a data field with multiple data items.
							}
							else if( nm.equals( "r" ) )
							{// Specifies the number of items to repeat from the previous row item. Note: The first item has no @r explicitly written. Since a default of "0" is specified in the schema, for any item
								// whose @r is missing, a default value of "0" is implied.
								repeat = Integer.valueOf( v ).intValue();
								index += repeat;
								nLines -= repeat;
							}
							else if( nm.equals( "t" ) )
							{// Specifies the type of the item. Value of 'default' indicates a grand total as the last row item value
								// default= data, avg, blank, count, countA, data, grand, max, min, product, stdDev, stdDevP, sum, var, varP
								type = Enum.valueOf( enumType, "_" + v ).ordinal();
								if( type != 0 || type != 0xE )
								{
//									index++;	stil confused on this ...
									nLines--;
								}
							}

						}
					}
					else if( tnm.equals( "x" ) )
					{
						indexes[index++] = (short) parseItemIndexOOXML( xpp );            // v
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endname = xpp.getName();
					if( endname.equals( elname ) )
					{
						break;
					}
					else if( endname.equals( "i" ) )
					{
						if( isRowItems )
						{
							ptview.addPivotLineToROWAxis( repeat, nLines, type, indexes );
						}
						else
						{
							ptview.addPivotLineToCOLAxis( repeat, nLines, type, indexes );
						}
					}

				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parseItemOOXML:" + e.toString() );
		}
	}

	/*
	 * item types enum -- add an underscore because "default" is not a valid entry
	 */
	enum ITEMTYPES
	{
		_data,
		_default,
		_sum,
		_countA,
		_count,
		_avg,
		_max,
		_min,
		_product,
		_stdDev,
		_stdDevP,
		_var,
		_varP,
		_grand,
		_blank,
	}

	;

	/**
	 * Represents a generic field that can appear either on the column or the
	 * row region of the PivotTable. There will be as many <x> elements as there
	 * are item values in any particular column or row. attribute: x: Specifies
	 * the index to a pivotField item value
	 *
	 * @param xpp
	 * @throws XmlPullParserException
	 */
	private static void parseFieldOOXML( XmlPullParser xpp, Sxview ptview ) throws XmlPullParserException
	{
		try
		{
			//int fieldcount = Integer.valueOf(xpp.getAttributeValue(0)).intValue();
			int eventType = xpp.getEventType();
			String elname = xpp.getName();			
/*			if (elname.equals("rowFields"))
				ptview.setCDimRw((short)fieldcount);				
			else
				ptview.setCDimCol((short)fieldcount);*/
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "field" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "x" ) )
							{
								String v = xpp.getAttributeValue( i );
								if( elname.equals( "rowFields" ) )
								{
									ptview.addRowField( Integer.valueOf( v ) );
								}
								else
								{
									ptview.addColField( Integer.valueOf( v ) );
								}
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( elname ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			throw new XmlPullParserException( "parseFieldOOXML:" + e.toString() );
		}
	}

	/**
	 * This element represents an array of indexes to cached shared item values
	 * <br>element x
	 * <br>child of reference (...pivotField), i (rowField, colField)
	 *
	 * @param xpp
	 * @return int index
	 * @throws XmlPullParserException
	 */
	private static int parseItemIndexOOXML( XmlPullParser xpp )
	{
		int v = 0;
		try
		{
			if( xpp.getAttributeCount() > 0 )
			{
				v = Integer.valueOf( xpp.getAttributeValue( 0 ) ).intValue();    // index
			}
		}
		catch( Exception e )
		{
		}
		return v;
	}

	private static void parseDataFieldOOXML( XmlPullParser xpp, Sxview ptview )
	{
		/*
 <attribute name="subtotal"
		 * type="ST_DataConsolidateFunction" default="sum"/> 9
		 * <attribute name="showDataAs" type="ST_ShowDataAs"
		 * default="normal"/> 10 <attribute name="baseField"
		 * type="xsd:int" default="-1"/> 11 <attribute
		 * name="baseItem" type="xsd:unsignedInt"
		 * default="1048832"/> 12 <attribute name="numFmtId"
		 * type="ST_NumFmtId" use="optional"/>
		 */
		int fieldIndex = 0;
		String aggregateFunction = null;
		String name = null;
		for( int z = 0; z < xpp.getAttributeCount(); z++ )
		{
			String nm = xpp.getAttributeName( z );
			String v = xpp.getAttributeValue( z );
			if( nm.equals( "name" ) )
			{
				name = v;
			}
			else if( nm.equals( "fld" ) )
			{
				fieldIndex = Integer.valueOf( xpp.getAttributeValue( z ) );
			}
			else if( nm.equals( "subtotal" ) )    // default= "sum"
			{
				aggregateFunction = v;
			}
		}
// TODO: 
// showDataAs, baseItem, baseField		--> display format
// numFmtId		
		ptview.addDataField( fieldIndex, aggregateFunction, name );
	}

	/**
	 * parse the pageField element, which defines a pivot field on the PAGE axis
	 *
	 * @param xpp
	 * @param ptview
	 */
	private static void parsePageFieldOOXML( XmlPullParser xpp, Sxview ptview )
	{
		int fieldIndex = 0, itemIndex = 0x7FFD;
		for( int i = 0; i < xpp.getAttributeCount(); i++ )
		{
			String nm = xpp.getAttributeName( i );
			if( nm.equals( "fld" ) )
			{
				fieldIndex = Integer.valueOf( xpp.getAttributeValue( i ) );
			}
			else if( nm.equals( "item" ) )
			{
				itemIndex = Integer.valueOf( xpp.getAttributeValue( i ) );
			}
		}
		ptview.addPageField( fieldIndex, itemIndex );
	}
	/**
	 * Changes a range in a PivotTable to expand until it includes the cell
	 * address from CellHandle.
	 *
	 * Example:
	 *
	 * CellHandle cell = new Cellhandle("D4") PivotTable = SUM(A1:B2)
	 * addCellToRange("A1:B2",cell); would change the PivotTable to look like
	 * "SUM(A1:D4)"
	 *
	 * Returns false if PivotTable does not contain the PivotTableLoc range.
	 *
	 * @param String
	 *            - the Cell Range as a String to add the Cell to
	 * @param CellHandle
	 *            - the CellHandle to add to the range
	 *
	 *            public boolean addCellToRange(String PivotTableLoc, CellHandle
	 *            handle) throws PivotTableNotFoundException{ int[]
	 *            PivotTableaddr = ExcelTools.getRangeRowCol(PivotTableLoc);
	 *            String handleaddr = handle.getCellAddress(); int[] celladdr =
	 *            ExcelTools.getRowColFromString(handleaddr);
	 *
	 *            // check existing range and set new range vals if the new Cell
	 *            is outside if(celladdr[0] >
	 *            PivotTableaddr[2])PivotTableaddr[2] = celladdr[0];
	 *            if(celladdr[0] < PivotTableaddr[0])PivotTableaddr[0] =
	 *            celladdr[0]; if(celladdr[1] >
	 *            PivotTableaddr[3])PivotTableaddr[3] = celladdr[1];
	 *            if(celladdr[1] < PivotTableaddr[1])PivotTableaddr[1] =
	 *            celladdr[1]; String newaddr =
	 *            ExcelTools.formatRange(PivotTableaddr); boolean b =
	 *            this.changePivotTableLocation(PivotTableLoc, newaddr); return
	 *            b; }
	 */

	/**
	 * Returns the "SXVIEW" record for this PivotTable.
	 *
	 * @return
	 *
	 *         public Sxview getPt() { return pt; }
	 */

	/**
	 * @param sxview public void setPt(Sxview sxview) { pt= sxview; }
	 */

	public OOXMLElement cloneElement()
	{
		return null;
	}

	public String getOOXML()
	{
		// TODO: Finish
		return null;
	}

	private static int axisLookup( String axis )
	{
		if( axis.equals( "axisRow" ) )
		{
			return Sxvd.AXIS_ROW;
		}
		else if( axis.equals( "axisCol" ) )
		{
			return Sxvd.AXIS_COL;
		}
		else if( axis.equals( "axisPage" ) )
		{
			return Sxvd.AXIS_PAGE;
		}
		else if( axis.equals( "axisValues" ) )
		{
			return Sxvd.AXIS_DATA;
		}
		return Sxvd.AXIS_NONE;

	}

}
