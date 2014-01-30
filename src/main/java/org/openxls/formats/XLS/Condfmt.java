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

import org.openxls.ExtenXLS.ExcelTools;
import org.openxls.ExtenXLS.FormatHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.formats.OOXML.CfRule;
import org.openxls.formats.OOXML.Dxf;
import org.openxls.toolkit.ByteTools;
import org.openxls.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * <b>Condfmt:  Conditional Formatting Range Information 0x1B0</b><br>
 * <p/>
 * This record stores a conditional format, including conditions and formatting info.
 * <p/>
 * And, no it does not just point to an Xf record because that would be easy.
 * <p/>
 * <p/>
 * OFFSET       NAME            SIZE        CONTENTS
 * -----
 * 4                ccf                 2           Number of Conditional formats
 * 6                grbit               2           Option flags (not a byte?) [1 = Conditionally formatted cells need recalculation or redraw]
 * 8                rwFirst             2           First row to conditionally format (0 based)
 * 10               rwLast              2           Last row to conditionally format (0 based)
 * 12               colFirst            2           First column to conditionally format (0 based)
 * 14               colLast             2           Last column to conditionally format (0 based)
 * 16               sqrefCount          2           Count of sqrefs *
 * 18               rgbSqref            var         Array of sqref structures
 * <p/>
 * <p/>
 * Sqref Structures
 * <p/>
 * OFFSET       NAME            SIZE        CONTENTS
 * -----
 * 0            rwFirst             2           First row in reference
 * 2            rwLast              2           Last row in reference
 * 4            colFirst            2           First column in reference
 * 6            colLast             2           Last column in reference
 *
 * @see Cf
 */
public final class Condfmt extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Condfmt.class );
	private FormatHandle formatHandle = null;

	private static final long serialVersionUID = -7923448634000437926L;
	short grbit = 0;        //      Option flags (not a byte?)

	private int ccf;
	DiscontiguousRefStruct refs = null;
	private ArrayList cfRules = new ArrayList(); // 2003-version Cf recs OR OOXML cfRules TODO: eventually will generate Cf records instead
	private int cfxe = -1; // a fake ixfe for use by ExtenXLS to track formats
	boolean isdirty = false;    // if any changes to underlying record is made, set to true

	/**
	 * set dirty flag to rebuild condfmt record
	 * used when updated the underlying ranges wit
	 */
	public void setDirty()
	{
		isdirty = true;
	}

	/**
	 * initialize the condfmt record
	 * <p/>
	 * Please note that the sqref structure is not initialized in this location, but is required for cfmt functionality.
	 * <p/>
	 * It happens on parse after worksheet is set
	 */
	@Override
	public void init()
	{
		super.init();
		rw = 0;
		ccf = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );     // SHOULD BE # cf's but appears to be 1+ ??
		grbit = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );   // SHOULD BE 1 to recalc but has been 3, 5, ...??
	}

	/**
	 * As the init() call occurs before worksheet is set upon this conditionalformat record,
	 * we have to initialze the references after init in order for referenceTracker to work correctly
	 */
	public void initializeReferences()
	{
		data = getData();
		int sqrefCount = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );
		byte[] sqrefdata = new byte[sqrefCount * 8];
		System.arraycopy( data, 14, sqrefdata, 0, sqrefdata.length );
		refs = new DiscontiguousRefStruct( sqrefdata, this );
	}

	/**
	 * default constructor
	 */
	public Condfmt()
	{
	}

	/**
	 * @return Returns the formatHandle.
	 */
	public FormatHandle getFormatHandle()
	{
		return formatHandle;
	}

	/**
	 * @param formatHandle The formatHandle to set.
	 */
	public void setFormatHandle( FormatHandle formatHandle )
	{
		this.formatHandle = formatHandle;
	}

	/**
	 * This is an overall not ideal situation where we want a formatid that we can use in sheetster
	 * to identify this conditional format
	 * <p/>
	 * // TODO: Perfect this algorithm!! :)  cfxe should be constant for this Condfmt
	 * // ... if address changes?  if sheet # changes
	 *
	 * @return Returns the cfxe.
	 */
	public int getCfxe()
	{
		int[] rc = refs.getRowColBounds();
		cfxe = 50000 + (getSheet().getSheetNum() * 10000) + ByteTools.readShort( rc[0], rc[1] );    // base cxfe on cell address
		return cfxe;
	}

	/**
	 * @param cfxe The cfxe to set.
	 */
	public void setCfxe( int c )
	{
		cfxe = c;
	}

	/**
	 * returns the rules associated with this record
	 *
	 * @return
	 */
	public ArrayList getRules()
	{
		return cfRules;
	}

	/**
	 * add a new CF rule to this conditional format
	 */
	public void addRule( Cf c )
	{
		if( cfRules.indexOf( c ) == -1 )
		{
			cfRules.add( c );
		}
		c.setCondfmt( this );
	}

	/**
	 * Return all ranges as strings
	 *
	 * @return
	 */
	public String[] getAllRanges()
	{
		return refs.getRefs();
	}

	/**
	 * Returns the entire range this conditional format refers to  in Row[0]Col[0]Row[n]Col[n] format.
	 *
	 * @return
	 */
	public int[] getEncompassingRange()
	{
		return refs.getRowColBounds();
	}

	/**
	 * update data for streaming
	 *
	 * @param loc
	 */
	private void updateRecord()
	{
		if( !isdirty )
		{
			return;
		}
		// get the size of our output
		byte[] outdata = new byte[(refs.getNumRefs() * 8) + 14];
		byte[] tmp = ByteTools.shortToLEBytes( (short) (getRules().size()) );
		int offset = 0;
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];

		tmp = ByteTools.shortToLEBytes( grbit );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];

		int[] rowcols = refs.getRowColBounds();
		tmp = ByteTools.shortToLEBytes( (short) rowcols[0] );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];
		tmp = ByteTools.shortToLEBytes( (short) rowcols[2] );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];
		tmp = ByteTools.shortToLEBytes( (short) rowcols[1] );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];
		tmp = ByteTools.shortToLEBytes( (short) rowcols[3] );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];

		tmp = ByteTools.shortToLEBytes( (short) refs.getNumRefs() );
		outdata[offset++] = tmp[0];
		outdata[offset++] = tmp[1];

		byte[] sqrefbytes = refs.getRecordData();
		System.arraycopy( sqrefbytes, 0, outdata, offset, sqrefbytes.length );
		setData( outdata );
	}

	/**
	 * Add a location to the conditional format record, this
	 * is a string representation that can be either a cell, ie "A1", or a range
	 * ie "A1:A12";
	 *
	 * @param location string representing the added range
	 */
	public void addLocation( String location )
	{
		refs.addRef( location );
		isdirty = true;
	}

	/**
	 * get the bounding range of this conditional format
	 *
	 * @return
	 */
	public String getBoundingRange()
	{
		int[] rowcols = refs.getRowColBounds();
		return ExcelTools.formatRangeRowCol( rowcols );
	}

	/**
	 * Set this cf to a new enclosing cell range
	 * This should only be used for inital creation of a conditional format
	 * record or when all other internal ranges should be cleared as it removes
	 * all others
	 *
	 * @param range
	 */
	public void resetRange( String range )
	{
		refs = new DiscontiguousRefStruct( range, this );
		isdirty = true;
	}

	private byte[] PROTOTYPE_BYTES = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	/**
	 * Create a Condfmt record & populate with prototype bytes
	 *
	 * @return
	 */
	protected static XLSRecord getPrototype()
	{
		Condfmt cf = new Condfmt();
		cf.setOpcode( CONDFMT );
		cf.setData( cf.PROTOTYPE_BYTES );
		cf.init();
		return cf;
	}

	/**
	 * update the bytes
	 */
	@Override
	public void preStream()
	{
		updateRecord();
	}

	/**
	 * OOXML conditionalFormatting (Conditional Formatting)
	 * A Conditional Format is a format, such as cell shading or font color,
	 * that a spreadsheet application can
	 * automatically apply to cells if a specified condition is true.
	 * This collection expresses conditional formatting rules
	 * applied to a particular cell or range.
	 *
	 * parent:   worksheet
	 * children: cfRule  (1 or more)
	 * attributes: pivot (flag indicating this cf is assoc with a pivot table), sqref
	 */

	/**
	 * create one or more Data Validation records based on OOXML input
	 */
	// TODO: finish pivot option, create Cf recs on each cfRule
	public static Condfmt parseOOXML( XmlPullParser xpp, WorkBookHandle wb, Boundsheet bs )
	{
		Condfmt condfmt = null;
		List dxfs = wb.getWorkBook().getDxfs();
		if( dxfs == null )
		{
			dxfs = new ArrayList();    // shouldn't!
		}
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "conditionalFormatting" ) )
					{      // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( n.equals( "sqref" ) )
							{    // series of references
								condfmt = bs.createCondfmt( "", wb );    //(Condfmt) Condfmt.getPrototype();
								condfmt.initializeReferences();
								String[] ranges = StringTool.splitString( v, " " );
								for( String range : ranges )
								{
									condfmt.addCellRange( bs.getSheetName() + "!" + range );
								}
							}
							else if( n.equals( "pivot" ) )
							{
								// ???
							}
						}
						// create a Cf record based upon cfRule info
					}
					else if( tnm.equals( "cfRule" ) )
					{  // one or more
						CfRule cfRule = (CfRule) CfRule.parseOOXML( xpp ).cloneElement();
						Cf cf = bs.createCf( condfmt );    // creates a new cf rule and links to the current condfmt
						cf.setOperator( Cf.translateOperator( cfRule.getOperator() ) );    // set the cf rule operator	(greater than, equals ...)
						cf.setType( Cf.translateOOXMLType( cfRule.getType() ) );                // set the cf rule type (cell is, exrpression ...)
						if( cf.getType() == 3 )// containsText
						{
							cf.setContainsText( cfRule.getContainsText() );
						}
						if( cfRule.getFormula1() != null )
						{
							cf.setCondition1( cfRule.getFormula1() );
						}
						if( cfRule.getFormula2() != null )
						{
							cf.setCondition2( cfRule.getFormula2() );
						}
						int dxfId = cfRule.getDxfId();
						if( dxfId > -1 )
						{    // it's not required to have a dxf
							Dxf dxf = (Dxf) dxfs.get( dxfId );    // dxf= differential format, contains the specific styles to define this cf rule
							Cf.setStylePropsFromDxf( dxf, cf );
//	                        String dxfStyleString= dfx.getStyleProps();	// returns a string representation of the dxf or differential styles
//	                        Cf.setStylePropsFromString(dxfStyleString,cf);	// set the dxf styles to the cf rule
						}
						// original code that didn't input CfRules into Cf's just stored CfRule objects ...                        condfmt.cfRules.add((CfRule.parseOOXML(xpp).cloneElement()));
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "conditionalFormatting" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "Condfmt.parseOOXML: " + e.toString() );
		}
		if( condfmt != null )
		{
			bs.addConditionalFormat( condfmt );   // add this conditional format to the sheet
		}
		return condfmt;
	}

	/**
	 * Add a cell range to this conditional format.
	 *
	 * @param string
	 */
	private void addCellRange( String range )
	{
		refs.addRef( range );
		isdirty = true;
		updateRecord();
	}

	/**
	 * returns EXML for the Conditional Format
	 * <p/>
	 * <p/>
	 * <p/>
	 * <ConditionalFormatting>
	 * <Range>R12C2:R16C2</Range>
	 * <Condition>
	 * <Qualifier>Between</Qualifier>
	 * <Value1>2</Value1>
	 * <Value2>4</Value2>
	 * <Format Style='color:#002060;font-weight:700;text-line-through:none;
	 * border:.5pt solid windowtext;background:#00B0F0'/>
	 * </Condition>
	 * </ConditionalFormatting>
	 *
	 * @return XML string for this record
	 */
	public String getXML()
	{
		return getXML( false );
	}

	/**
	 * returns XMLSS for the Conditional Format
	 * <p/>
	 * <p/>
	 * <p/>
	 * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
	 * <Range>R12C2:R16C2</Range>
	 * <Condition>
	 * <Qualifier>Between</Qualifier>
	 * <Value1>2</Value1>
	 * <Value2>4</Value2>
	 * <Format Style='color:#002060;font-weight:700;text-line-through:none;
	 * border:.5pt solid windowtext;background:#00B0F0'/>
	 * </Condition>
	 * </ConditionalFormatting>
	 *
	 * @return XML string for this record
	 */
	public String getXMLSS()
	{
		return getXML( true );
	}

	/**
	 * returns EXML (XMLSS) for the Conditional Format
	 * <p/>
	 * <p/>
	 * <p/>
	 * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
	 * <Range>R12C2:R16C2</Range>
	 * <Condition>
	 * <Qualifier>Between</Qualifier>
	 * <Value1>2</Value1>
	 * <Value2>4</Value2>
	 * <Format Style='color:#002060;font-weight:700;text-line-through:none;
	 * border:.5pt solid windowtext;background:#00B0F0'/>
	 * </Condition>
	 * </ConditionalFormatting>
	 *
	 * @return
	 */
	public String getXML( boolean useXMLSSNameSpace )
	{

		String ns = "";
		if( useXMLSSNameSpace )
		{
			ns = "xmlns=\"urn:schemas-microsoft-com:office:excel\"";
		}

		StringBuffer xml = new StringBuffer( "<ConditionalFormatting" + ns + ">" );
		// cf's
		Iterator its = getRules().iterator();
		while( its.hasNext() )
		{
			Cf c = (Cf) its.next();
			xml.append( c.getXML() );
		}
		xml.append( "</ConditionalFormatting>" );
		return xml.toString();
	}

	/**
	 * generate the proper OOXML to define this set of Conditional Formatting
	 *
	 * @return
	 */
	public String getOOXML( WorkBookHandle bk, int[] priority )
	{
		updateRecord();
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<conditionalFormatting" );
		if( refs != null )
		{
			ooxml.append( " sqref=\"" );
			String[] refStrs = refs.getRefs();
			for( int i = 0; i < refStrs.length; i++ )
			{
				if( i > 0 )
				{
					ooxml.append( " " );
				}
				ooxml.append( refStrs[i] );
			}
			ooxml.append( "\"" );
		}
		ooxml.append( ">" );
		// cf's
		// NOTE:  cf.getDxfId/setDxfId links this conditional formatting rule with the proper incremental style
		// NOTE:  cfRules must have a valid dxfId or the output file will open with errors
		// NOTE:  for now, dxfs can only be saved from the original styles.xml;
		List dxfs = getWorkBook().getDxfs();
		if( dxfs == null )
		{
			dxfs = new ArrayList();
			getWorkBook().setDxfs( dxfs );
		}
		if( cfRules != null )
		{
			for( Object cfRule : cfRules )
			{
				ooxml.append( ((Cf) cfRule).getOOXML( bk, priority[0]++, dxfs ) );
			}
		}
		ooxml.append( "</conditionalFormatting>" );
		return ooxml.toString();
	}

	/**
	 * Checks if the conditional format contains the row/col passed in
	 *
	 * @param rowColFromString
	 * @return
	 */
	public boolean contains( int[] rowColFromString )
	{
		return refs.containsReference( rowColFromString );
	}

	/**
	 * clear out object referencse
	 */
	@Override
	public void close()
	{
		super.close();
		refs = null;
		if( cfRules != null )
		{
			for( Object cfRule : cfRules )
			{
				Cf cf = (Cf) cfRule;
				cf.close();
				cf = null;

			}
		}
	}

}
