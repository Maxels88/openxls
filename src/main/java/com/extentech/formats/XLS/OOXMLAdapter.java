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

/**
 *  OOXMLAdapter generates OOXML (SpreadsheetML packaged in ZIP) for a given workbook adhering to the following specification:
 *row
 *  Concept of Open Package Convention (OPC) is identification of Relationships:
 *      .rels files specify Relationship type, Relatioship id rId and Target filename
 *      in specific XML file, these rIds are specified, and the .rels files are used to look up the particular file 
 *
 * OOXML for a workbook consists of multiple files contained in a ZIP
 *
 *
 *  Open Package Convention Structure of SpreadsheetML ZIP:
 *
 *      [Content_types].xml             specifies the content types and the Parts (i.e. files) used in the ZIP package
 *      \_rels directory
 *              workbook.xml.rels       relationship file for workbook, lists target files and their relationship type
 *      \xl directory
 *          styles.xml                      contains font, fill and xf specs
 *          sharedStrings.xml               contains info re: ssts
 *          workbook.xml                    specifies sheet rIds
 *          \rels\workbook.xml.rels         specifies sheet targets 
 *          \worksheets directory           
 *                  sheetXX.xml             contains row and cell info ... 
 *                                          if contains drawing ml-specific data will contain "rid" linking
 *              \rels
 *                  sheetXX.xml.rels        if necessary for linking drawingML rIds, printerSetting rIDs 
 * (if necessary also contains:)
 *          \theme directory
 *          \charts directory               
 *                  chart1.xml
 *          \chartsheets directory
 *          \drawings directory
 *                  drawingXX.xml           chart and image info
 *                  \rels               
 *                      drawing1.xml.rels   links rid's to Target chart XMLs
 *          \media directory
 *                  contains image files
 *          \printerSettings
 *                  contains PrinterSettingsX.bin
 *          \dialogsheets directory
 *          \macrosheets directory
 *  \docProps
 */

import com.extentech.ExtenXLS.ChartHandle;
import com.extentech.ExtenXLS.Document;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.XML.UnicodeInputStream;
import com.extentech.toolkit.StringTool;
import com.extentech.toolkit.TempFileManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import org.xmlpull.v1.XmlPullParserFactory;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.CharArrayReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Stack;
import java.util.TreeMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

// TODO: finish Style options, Chart options, Japanese Chars ..., 
// TODO: dataTables
// TODO: external names
// TODO: drawing shapes
// TODO: Pivot Tables, Macro Sheets ...
// TODO: handle generation of: printer settings, doc properties, external refs ... on xls --> xlsx

/* ******************************************************************************************************
 * Dissertation on one of the main complications in the below code:
 * 
 * "External Objects" are XMLs and associated files that are external to the main workbook and worksheets, 
 * and essentially are treated as pass-throughs (except for charts, sharedString table + styles, which are generated)
 * 
 * Because these are not created upon output, pass-through external objects (xml files, bins, etc) are written to a directory upon input,
 * and pulled from that directory upon output.  Data associated with the external object is stored in either WorkBook.getOOMXLObjects or BoundSheet.getOOXMLObjects.
 * Linked data associated with the pass-through object is stored and .rels are re-created   
 * 
 * Here is a list of the External Objects:
 * 
 * (Workbook Level)
 * vbaProject
 * externalLink**
 * calcChain
 * sharedStrings
 * styles
 * theme**
 * 
 * (Sheet Level)
 * printerSettings
 * drawing**
 * control (=activeX)
 * oleObject
 * 
 * **external items may also have linked items to them e.g. themes may have object associated with them
 * 
 * All associated objects are listed in the associated .rels file
 */
public class OOXMLAdapter implements OOXMLConstants
{
	private static final Logger log = LoggerFactory.getLogger( OOXMLAdapter.class );

	ZipOutputStream zip;
	Writer writer;
	Map deferredFiles = new HashMap();
	// map of original external filename and new filename on disk (may be different as images, etc. must have consecutive indexes ...
//    HashMap externalFiles= new HashMap();  not used anylonger

	// content lists used to create all .rels + [Content_Types].xml
	ArrayList mainContentList = new ArrayList();        // [Content_Types].xml
	ArrayList wbContentList = new ArrayList();            // workbook.xml.rels	- written to [Content_Types].xml
	ArrayList drContentList = new ArrayList();            // drawingX.xml.rels	- written to [Content_Types].xml
	ArrayList shContentList = new ArrayList();            // sheetX.xml.rels
	ArrayList sheetsContentList = new ArrayList();        // total contents of each sheet - written to [Content_Types].xml
	// also have vmContentList and chContentList for fairly rare occurrences of vmldrawings and charts containing embeds

	// External OOXML Object such as Vba projects, Ole Objects, Printer Settings, etc.
	// links external ref "extra info" with the external reference id
//     Map   shExternalLinkInfo= new HashMap();

	String externalDir = ""; // store "pass-through" files i.e. files we cannot process into our BIFF8 rec structure (vbaProject.bin, for example)
	// ordinal numbers for sheet-level objects (rid is stored in sheetX.xml and file stored in appropriate directory, liked via sheetX.xml.rels
	// Each item (images, charts, etc) has a very specific and ordered name format e.g. image1.jpg, printerSettings2.bin ...
	int drawingId = 0;
	int vmlId = 0;                    // vmlDrawing.vml
	int commentsId = 0;
	int activeXId = 0;
	int activeXBinaryId = 0;        // activeX.bin
	int printerSettingsId = 0;
	int oleObjectsId = 0;
	int chartId = 0;
	int imgId = 0;

	static final double rowHtFactor = 20.0;
	static final double colWFactor = 256.0;

	// TODO: finish Styles OOXML --> cellStyleXfs, MANY options not handled ...
	// TODO: finish charts and images (many options not handled)
	// TODO: handle tableData
	// TODO: handle themes, doc properties (create? alter?)
	// TODO: handle shapes

	// ***************************************************************************************
	// contents of External OOXMLObject arraylist
	int EX_TYPE = 0;            // type of External Object - must be listed in OOXMLConstants
	int EX_PATH = 1;            // path in ZIP
	int EX_FNAME = 2;            // file name
	// 3= rid
	int EX_EXTRAINFO = 4;        // any extra information associated - object specific
	int EX_EMBEDINFO = 5;        // string of embedd
	// ***************************************************************************************
	int format = WorkBookHandle.FORMAT_XLSX;    // default format is non-macro-enabled workbook
	static String inputEncoding = "UTF-8";    // default

	/**
	 * set the XLSX format for this WorkBook
	 * <br> either FORMAT_XLSX, FORMAT_XLSM (Macro-enabled), FORMAT_XLTS (template)  or FORMAT_XLTM (Macro-enabled template)
	 * <br>NOTE: If file extension is .XLSM format FORMAT_XLSM must be set
	 * <br>either because there are macros present or because the filename
	 * <br>is unconditionally set to .XLSM
	 *
	 * @param format
	 */
	public void setFormat( int format )
	{
		this.format = format;
	}

	public int getFormat()
	{
		return format;
	}

	/**
	 * Parses an xsd:boolean value.
	 *
	 * @param value the string to parse
	 * @return the boolean value of the given string
	 * @throws IllegalArgumentException if the given string is not a valid
	 *                                  boolean value
	 */
	public static final boolean parseBoolean( String value )
	{
		String trimmed = value.trim();
		if( trimmed.equals( "true" ) || trimmed.equals( "1" ) )
		{
			return true;
		}
		if( trimmed.equals( "false" ) || trimmed.equals( "0" ) )
		{
			return false;
		}
		throw new IllegalArgumentException( "'" + value + "' is not a valid boolean value" );
	}

	/**
	 * get a standalone ChartML document
	 *
	 * @param ch
	 * @return
	 */
	public static String getStandaloneChartDrawingOOXML( ChartHandle ch )
	{
		String ret = "";
		try
		{
			// trap package contents for drawing.xml
			StringBuffer chartml = new StringBuffer();
			chartml.append( xmlHeader );
			chartml.append( "\r\n" );
			chartml.append( "<c:chartSpace xmlns:c=\"" + chartns + "\" xmlns:a=\"" + drawingmlns + "\" xmlns:r=\"" + relns + "\">" );
			chartml.append( "\r\n" );
			chartml.append( ch.getOOXML( 1 ) );
			chartml.append( "</c:chartSpace>" );
			chartml.append( "\r\n" );
			return chartml.toString();
		}
		catch( Exception e )
		{
			log.error( "OOXMLAdapter.getStandaloneChartDrawingOOXML: " + e.toString() );
			ret = "";
		}
		return ret;
	}

	/**
	 * generic utility to take the original external object filename and parse
	 * it using a new ordinal id e.g. image10.emf may become image3.emf NOTE:
	 * Increments global ordinal ids
	 *
	 * @param f Original filename with path
	 * @return New filename
	 */
	protected String getExOOXMLFileName( String f )
	{
		String fname = f.substring( f.lastIndexOf( "/" ) + 1 );
		String ext = fname.substring( fname.lastIndexOf( "." ) );
		String root = fname.substring( 0, fname.indexOf( "." ) );
		int z = root.length() - 1; // now skip # appended to end of root (number
		// will be regenerated from id)
		while( Character.isDigit( root.charAt( z ) ) )
		{
			z--;
		}
		root = root.substring( 0, z + 1 );
		if( root.equalsIgnoreCase( "image" ) )
		{
			fname = root + (++imgId) + ext;
		}
		else if( root.equalsIgnoreCase( "oleObject" ) )
		{
			fname = root + (++oleObjectsId) + ext;
		}
		else if( root.equalsIgnoreCase( "activeX" ) )
		{
			if( ext.toLowerCase().equals( ".xml" ) )
			{
				fname = root + (++activeXId) + ext;
			}
			else if( ext.toLowerCase().equals( ".bin" ) )
			{
				fname = root + (++activeXBinaryId) + ext;
			}
		}
		else if( root.equalsIgnoreCase( "printerSettings" ) )
		{
			fname = root + (++printerSettingsId) + ext;
		}
		else if( root.equalsIgnoreCase( "drawing" ) )
		{
			fname = root + (++drawingId) + ext;
			// if workbook-level file, no incrementing id, just use original
			// filename
		}
		else if( root.equalsIgnoreCase( "app" ) || root.equalsIgnoreCase( "core" ) || root.equalsIgnoreCase( "theme" ) || root.equalsIgnoreCase(
				"themeOverride" ) || root.equalsIgnoreCase( "app" ) || root.equalsIgnoreCase( "custom" ) || root.equalsIgnoreCase(
				"connections" ) || root.equalsIgnoreCase( "externalLink" ) || root.equalsIgnoreCase( "calcChain" ) || root.equalsIgnoreCase(
				"styles" ) || root.equalsIgnoreCase( "sharedStrings" ) || root.equalsIgnoreCase( "vbaProject" ) )
		{
			; // do nothing, use original fname *** these do not have ordinal id's
		}
		else
		// TESTING - remove when done!
		// *********************************************************************
		{
			log.error( "Unknown External Type: " + root );
		}
		return fname;
	}

	/**
	 * write a temp file for later inclusion in the zip file
	 *
	 * @param sb
	 * @param fn
	 */
	protected void addDeferredFile( StringBuffer sb, String fn )
	{
		try
		{
			File fx = addDeferredFile( fn );
			FileOutputStream fos = new FileOutputStream( fx );
/* new way            
            OOXMLAdapter.writeSBToStreamEfficiently(sb,fos);
            fos.flush();
            fos.close();
*/           // have to use writer to set encoding - vital to non-utf8 input files 
			Writer out = new OutputStreamWriter( fos, inputEncoding );    //"UTF-8" );
			OOXMLAdapter.writeSBToStreamEfficiently( sb, out );
			out.close();
/* old way             
            Writer out = new OutputStreamWriter( fos, inputEncoding);	//"UTF-8" ); 
          out.write( sb.toString() );
            out.close();
*/
		}
		catch( Exception e )
		{
			log.error( "OOXMLAdapter addDeferredFile failed.", e );
		}
	}

	/**
	 * A memory efficient way to write a StringBuffer to an OutputStream
	 * without creating Strings and other Objects.
	 * <p/>
	 * Mar 9, 2012
	 *
	 * @param aSB
	 * @param ous
	 * @throws IOException
	 */
//    public static void writeSBToStreamEfficiently(StringBuffer aSB, OutputStream ous) throws IOException{
	public static void writeSBToStreamEfficiently( StringBuffer aSB, Writer ous ) throws IOException
	{

		final int aLength = aSB.length();
		final int aChunk = 1024;
		final char[] aChars = new char[aChunk];    //aPosEnd-aPosStart];	//aChunk];

		for( int aPosStart = 0; aPosStart < aLength; aPosStart += aChunk )
		{
			final int aPosEnd = Math.min( aPosStart + aChunk, aLength );
			aSB.getChars( aPosStart, aPosEnd, aChars, 0 );                 // Create no new buffer
			final CharArrayReader aCARead = new CharArrayReader( aChars ); // Create no new buffer

			// This may be slow but it will not create any more buffer (for bytes)
			int aByte;
			int i = 0;
			while( ((aByte = aCARead.read()) != -1) && (i++ < (aPosEnd - aPosStart)) )
			{
				ous.write( aByte );
			}
		}
	}

	protected File addDeferredFile( String name ) throws IOException
	{
		File file = TempFileManager.createTempFile( "OOXMLOutput_", ".tmp" );
		deferredFiles.put( name.toString(), file.getAbsolutePath() );
		return file;
	}

	/**
	 * write a temp file for later inclusion in the zip file
	 *
	 * @param sb
	 * @param fn
	 * @throws IOException
	 */
	protected void addDeferredFile( byte[] b, String fn ) throws IOException
	{
		File fx = addDeferredFile( fn );
		// write to temp file
		FileOutputStream fos = new FileOutputStream( fx );
		BufferedOutputStream bos = new BufferedOutputStream( fos );
		bos.write( b );
		bos.close();
	}

	/**
	 * returns truth of "Book Contains External OOXML Object named type"     *
	 *
	 * @param bk
	 * @param type e.g. "vba", "custprops"
	 * @return
	 */
	private boolean hasObject( WorkBookHandle bk, String type )
	{
		List externalOOXML = bk.getWorkBook().getOOXMLObjects();
		for( Object anExternalOOXML : externalOOXML )
		{
			String[] s = (String[]) anExternalOOXML;
			if( (s != null) && (s.length == 3) )
			{   // id, dir, filename
				if( s[0].equalsIgnoreCase( type ) )
				{
					return true;
				}
			}
		}
		return false;
	}

	/**
	 * return true of workbook bk contains macros
	 *
	 * @param bk
	 * @return
	 */
	public static boolean hasMacros( Document bk )
	{
		if( bk instanceof WorkBookHandle )
		{
			List externalOOXML = ((WorkBookHandle) bk).getWorkBook().getOOXMLObjects();
			for( Object anExternalOOXML : externalOOXML )
			{
				String[] s = (String[]) anExternalOOXML;
				if( (s != null) && (s.length == 3) )
				{   // id, dir, filename
					if( s[0].equalsIgnoreCase( "vba" ) || s[0].equalsIgnoreCase( "macro" ) )
					{
						return true;
					}
				}
			}
		}
		return false;
	}

	/**
	 * utility which, given an an exisitng file f2write,
	 * creates zipEntry and writes to  zip var
	 * <p/>
	 * NOTE: global zip ZipOutputStream must be open
	 *
	 * @param f2write name and path of exisitng file
	 * @param fname   desired name in zip
	 * @throws IOException
	 */
	protected void writeFileToZIP( String f2write, String fname ) throws IOException
	{
		nextZipEntry( fname );

		FileInputStream fis = new FileInputStream( f2write );
		BufferedInputStream bis = new BufferedInputStream( fis );

		int i = bis.read();
		while( i != -1 )
		{
			zip.write( i );
			i = bis.read();
		}

		bis.close();
	}

	protected void nextZipEntry( String name ) throws IOException
	{
		// Flush the writer to ensure data ends up in the right entry
		try
		{
			writer.flush();
		}
		catch( Exception e )
		{
			log.error( "Flush failing on zip entry, likely due to streaming first sheet " + e );
		}

		// Start the new entry in the ZIP file
		zip.putNextEntry( new ZipEntry( name ) );
	}

	/**
	 * utility to look up Content Type string for type abbreviation
	 *
	 * @param type
	 * @return
	 */
	protected String getContentType( String type )
	{
		for( String[] contentType : contentTypes )
		{
			if( contentType[0].equalsIgnoreCase( type ) )
			{
				return contentType[1];
			}
		}
		return "UNKNOWN TYPE " + type;
	}

	/**
	 * utility to look up Relationship Type string for type abbreviation
	 *
	 * @param type
	 * @return
	 * @see OOXMLConstnats
	 */
	protected String getRelationshipType( String type )
	{
		for( String[] relsContentType : relsContentTypes )
		{
			if( relsContentType[0].equalsIgnoreCase( type ) )
			{
				return relsContentType[1];
			}
		}
		return "UNKNOWN TYPE " + type;
	}

	/**
	 * utility to retrieve correct relationship type abbreviation string from verbose type string
	 *
	 * @param type
	 * @return
	 * @see OOXMLConstants
	 */
	protected static String getRelationshipTypeAbbrev( String type )
	{
		for( String[] relsContentType : relsContentTypes )
		{
			if( relsContentType[1].equalsIgnoreCase( type ) )
			{
				return relsContentType[0];
			}
		}
		return "UNKNOWN TYPE " + type;
	}

	// this is for testing purposes only, not used so no need to comment out logger msg :)
	private void getZipEntries( ZipFile zf )
	{
		// testing!!
		try
		{
			java.util.Enumeration e = zf.entries();
			while( e.hasMoreElements() )
			{
				ZipEntry ze = (ZipEntry) e.nextElement();
				log.info( ze.getName() );
			}
		}
		catch( Exception e )
		{
			log.error( "getZipEntries: " + e.toString() );
		}
	}

	/**
	 * Strip non-ascii (i.e. xml non-valid) chars from Strings
	 * This is utilized for xml attributes, node values can contain quote symbols
	 *
	 * @param s
	 * @return
	 */
	public static StringBuffer stripNonAscii( String s )
	{
		StringBuffer out = new StringBuffer();
		if( s == null )
		{
			return out;
		}
		/** FROM MS:
		 * "Special character" refers to any character outside the standard ASCII character set
		 * range of 0x00 - 0x7F, such as Latin characters with accents, umlauts, or other diacritics.
		 * The default encoding scheme for XML documents is UTF-8, which encodes ASCII characters with a
		 * value of 0x80 or higher differently than other standard encoding schemes.
		 Most often, you see this problem if you are working with data that uses the
		 simple "iso-8859-1" encoding scheme. In this case, the quickest solution is usually the first
		 or example, use the following XML declaration:
		 */
		// Legal characters are tab, carriage return, line feed, and the legal characters of Unicode and ISO/IEC 10646
		// XML processors MUST accept any character in the range specified for Char
		// Char	   ::=   	#x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]	/* any Unicode character, excluding the surrogate blocks, FFFE, and FFFF. */

/* fro wiki
 * Unicode code points in the following ranges are valid in XML 1.0 documents:[1]

    U+0009, U+000A, U+000D: these are the only C0 controls accepted in XML 1.0;
    U+0020–U+D7FF, U+E000–U+FFFD: this excludes some (not all) non-characters in the BMP (all surrogates, U+FFFE and U+FFFF are forbidden);
    U+10000–U+10FFFF: this includes all code points in supplementary planes, including non-characters.

The preceding code points ranges contain the following controls which are only valid in certain contexts in XML 1.0 documents, and whose usage is restricted and highly discouraged:
    U+007F–U+0084, U+0086–U+009F: this includes a C0 control characters, and most (not all) C1 controls.
 */
		for( int i = 0; i < s.length(); i++ )
		{
			char c = s.charAt( i );
			int charCode = c;
			if( (charCode == 0x9) ||
					(charCode == 0xA) ||
					(charCode == 0xD) ||
					((charCode >= 0x20) && (charCode <= 0xD7FF)) ||
					((charCode >= 0xE000) && (charCode <= 0xFFFD)) ||
					((charCode >= 0x10000) && (charCode <= 0x10FFFF)) )
			{
				if( charCode == '&' )
				{
					out.append( "&amp;" );
				}
				else if( charCode == '"' )
				{
					out.append( "&quot;" );
				}
				else if( charCode == '<' )
				{
					out.append( "&lt;" );
				}
				else if( charCode == '>' )
				{
					out.append( "&gt;" );
				}
				else if( charCode == '\'' )
				{
					out.append( "&apos;" );
				}
				else
				{
					out.append( c );
				}
			} /*else { // Encoding Q??????:
	        	System.out.println("Skipping Special Char: " + charCode); // skip it
            }*/
		}
/* these translations do not seem to make any difference for Baxter's issue
            } else if (charCode==8220) {	// smart quotes BAXTER ISSUE - TRY 
                out.append("&#8220;");
            } else if (charCode==8221) {	// ""
                out.append("&#8221;");
            } else if (charCode==8216) {	// ""
                out.append("&#8216;");
            } else if (charCode==8217) {	// ""
                out.append("&#8217;");
            } else if (charCode==8211){ // en dash
                out.append("&#8211;");            	
            } else if (charCode==8212){ // em dash
                out.append("&#8212;");            
            } else if (charCode==8242) { // single prime
            	out.append("&#8242;");
            } else if (charCode==8364) { // Euro Symbol
            	out.append("&#8364;");
            } else        // copyright= 169  // 10 \n  176=small circle 
            	out.append(c);		// 169=&copy;  		
        } 
*/
		return out;
	}

	/**
	 * Strip non-ascii (i.e. xml non-valid) chars from Strings
	 * Node values can contain quote symbols
	 *
	 * @param s
	 * @return
	 */
	public static StringBuffer stripNonAsciiRetainQuote( String s )
	{
		StringBuffer out = new StringBuffer();
		if( s == null )
		{
			return out;
		}
		for( int i = 0; i < s.length(); i++ )
		{
			char c = s.charAt( i );
			int charCode = c;
			if( (charCode >= 32) && (charCode <= 126) )
			{
				if( charCode == '&' )
				{
					out.append( "&amp;" );
				}
				else if( charCode == '<' )
				{
					out.append( "&lt;" );
				}
				else if( charCode == '>' )
				{
					out.append( "&gt;" );
				}
				else
				{
					out.append( c );
				}
			}
			else
			{
				out.append( c );
			}
		}
		return out;
	}

	/**
	 * deal with the "BOM" input streams...
	 * <p/>
	 * http://bugs.sun.com/bugdatabase/view_bug.do?bug_id=4508058
	 * <p/>
	 * <p/>
	 * Oct 14, 2010
	 *
	 * @param in
	 * @return
	 */
	protected static InputStream wrapInputStream( InputStream in )
	{
		UnicodeInputStream uin = new UnicodeInputStream( in, "UTF-8" );
		// String enc = uin.getEncoding(); // check for BOM mark and skip
		return uin;
	}

	/**
	 * used as a way to monitor Zip Entry fetching
	 * <p/>
	 * Oct 14, 2010
	 *
	 * @param f
	 * @param name
	 * @return
	 */
	protected static ZipEntry getEntry( ZipFile f, String name )
	{
		return f.getEntry( name );
	}

	/**
	 * parses any .rels file into content List array list
	 *
	 * @param ii
	 * @return
	 */
	protected static ArrayList parseRels( InputStream ii )
	{
		ArrayList contentList = new ArrayList();
		try
		{
			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();

			xpp.setInput( ii, null ); // using XML 1.0 specification

//            if(DEBUG) Logger.logInfo("parseRels InputStream has available bytes: " + ii.available());

			int eventType = xpp.getEventType();

//            if(DEBUG) Logger.logInfo("parseRels XPP Name: " + xpp.getName() );

//            if(DEBUG) Logger.logInfo("parseRels XPP Event Type: " + eventType );

			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				try
				{
					if( eventType == XmlPullParser.START_TAG )
					{
						String tnm = xpp.getName();
						if( (tnm != null) && tnm.equals( "Relationship" ) )
						{
							String type = "";
							String target = "";
							String rId = "";
							for( int i = 0; i < xpp.getAttributeCount(); i++ )
							{
								String nm = xpp.getAttributeName( i ); // id, Type,
								// Target
								String v = xpp.getAttributeValue( i );
								if( nm.equalsIgnoreCase( "Type" ) )
								{
									type = getRelationshipTypeAbbrev( v );
								}
								else if( nm.equalsIgnoreCase( "Target" ) )
								{
									target = v;
								}
								else if( nm.equalsIgnoreCase( "id" ) )
								{
									rId = v;
								}
							}
							// 20100426 KSC: unfortunately, need to ensure that commentsX.xml is processed before vmlDrawingX.xml
							if( target.indexOf( "comments" ) == -1 )
							{
								contentList.add( new String[]{ type, target, rId } );
							}
							else // ensure comments are before vmlDrawing so that Notes can be created
							{
								contentList.add( 0, new String[]{ type, target, rId } );
							}
						}
						else
						{
//                            if(DEBUG) Logger.logInfo("parseRels null entry name");
						}
					}
				}
				catch( Exception ea )
				{
					log.error( "XML Exception in OOXMLAdapter.parseRels. Input file is out of spec.", ea );
				}
				eventType = xpp.next();
			}

		}
		catch( org.xmlpull.v1.XmlPullParserException ex )
		{
			log.error( "XML Exception in OOXMLAdapter.parseRels. Input file is out of spec.", ex );

		}
		catch( Exception e )
		{
			log.error( "OOXMLAdapter.parseRels. " + e.toString() );
		}
		return contentList;
	}

	/**
	 * utility to retrieve Text element for tag
	 *
	 * @param xpp
	 * @return
	 * @throws IOException
	 */
	public static String getNextText( XmlPullParser xpp ) throws IOException, XmlPullParserException
	{
		int eventType = xpp.next();
		String ret = "";
		while( (eventType != XmlPullParser.END_DOCUMENT) &&
				(eventType != XmlPullParser.END_TAG) &&
				(eventType != XmlPullParser.START_TAG) && (eventType != XmlPullParser.TEXT) )
		{
			eventType = xpp.next();
		}
		if( eventType == XmlPullParser.TEXT )
		{
			ret = xpp.getText();
		}

		try
		{
			return new String( ret.getBytes(), inputEncoding );
/* KSC: replaced with above *SHOULD* be correct
        	 if (!isUTFEncoding)
        		 if (xpp.getInputEncoding().equals("UTF-8"))
        			 return new String(ret.getBytes(), "UTF-8");	// ensure encoding
        		 else
        			 return new String(ret.getBytes(), xpp.getInputEncoding());	// ensure encoding*/
		}
		catch( Exception e )
		{
		}    // inputEncoding can be null
		return ret;
	}

	/**
	 * simple utility that ensures that sheets are last in the workbookcontent list
	 * also ensure that theme(s) are parsed first, as are used in styles etc.
	 * in order to create all dependent objects first
	 *
	 * @param wbContentList
	 */
	protected void reorderWbContentList( ArrayList wbContentList )
	{
		for( int j = 0; j < wbContentList.size(); j++ )
		{
			String[] wb = (String[]) wbContentList.get( j );
			if( !wb[0].equals( "sheet" ) )
			{    // sheets come last
				wbContentList.remove( j );
				wbContentList.add( 0, wb );
			}
		}
		for( int j = 0; j < wbContentList.size(); j++ )
		{
			String[] wb = (String[]) wbContentList.get( j );
			if( wb[0].equals( "theme" ) )
			{    //  goes before styles
				wbContentList.remove( j );
				wbContentList.add( 0, wb );
				break;
			}
		}

	}

	/**
	 * return the file matching rId in the ContentList (String[] type, filename, rId)
	 *
	 * @param contentList
	 * @param rId
	 * @return
	 */
	protected static String getFilename( ArrayList contentList, String rId )
	{
		for( Object aContentList : contentList )
		{
			String[] s = (String[]) aContentList;
			if( s[2].equals( rId ) )
			{
				return s[1];
			}
		}
		return null;
	}

	/**
	 * creates a copy of a stack so changes won't affect original
	 *
	 * @param origStack
	 * @return
	 */
	protected Stack cloneStack( Stack origStack )
	{
		Stack s = new Stack();
		for( int i = 0; i < origStack.size(); i++ )
		{
			s.push( origStack.elementAt( i ) );
		}
		return s;
	}

	public static boolean deleteDir( File f )
	{
		if( f.isDirectory() )
		{
			String[] children = f.list();
			for( String aChildren : children )
			{
				boolean success = deleteDir( new File( f, aChildren ) );
				if( !success )
				{
					//return false;
				}
			}
		}
		// The directory is now empty so delete it
		f.deleteOnExit();
		return f.delete();
	}

	public static String getTempDir( String f ) throws IOException
	{
    	/*
    	File fx = TempFileManager.createTempFile("OOXMLA",".tmp");
    	File fdir = fx.getParentFile();
    	String s = "";
    	if(fdir.isDirectory())
    		s = fdir.getAbsolutePath();
    	else{
    		*/
		String s = System.getProperty( "java.io.tmpdir" );

		if( !(s.endsWith( "/" ) || s.endsWith( "\\" )) )
		{
			s += "/";
		}
		f = StringTool.stripPath( f );
		s += "extentech/";
		if( f.indexOf( '.' ) > 0 )
		{
			s += f.substring( 0, f.indexOf( '.' ) ) + "/";
		}
		else
		{
			s += f + "/";
		}

		return s;
	}

	/**
	 * sorts the sheets for incoming workbook xlsx -- used for eventMode only
	 * <p/>
	 * Jan 19, 2011
	 *
	 * @param cl
	 */
	protected static void sortSheets( ArrayList cl ) throws Exception
	{
		// take the array of storages and find sheets

		// use natural sort of the Tree
		TreeMap sorted = new TreeMap();

		Iterator its = cl.iterator();

		while( its.hasNext() )
		{
			String[] c = (String[]) its.next();
			String shtnm = c[1];
			// parse out sheet number
			// xxxxsheet1.xmlxxxx
			int st = shtnm.indexOf( "worksheets/sheet" );
			if( st > -1 )
			{
				st += 16;
				shtnm = shtnm.substring( st );
				shtnm = shtnm.substring( 0, shtnm.toLowerCase().indexOf( ".xml" ) );
				try
				{
					int ti = Integer.parseInt( shtnm );
					// we know the sheet number, add to the tree
					sorted.put( ti, c );
				}
				catch( Exception e )
				{
					log.error( "Could not sort sheets", e );
					return;
				}
			}
		}
		// now we have the sorted map of "CLs" we can re-order sheets in arraylist
		its = cl.iterator();
		Iterator sort = sorted.values().iterator();
		int clpos = -1;
		while( sort.hasNext() )
		{
			String[] c = (String[]) sort.next();
			boolean found = false;
			// now find in cl
			while( its.hasNext() && !found )
			{
				clpos++;
				String[] cx = (String[]) its.next();
				if( cx[0].equals( "sheet" ) )
				{ // replace the sheet entries one by one
					cl.set( clpos, c );
					found = true;
				}
			}
		}
	}

	/**
	 * re-save to temp directroy any "pass-through"
	 * OOXML files i.e. files/entities not present in 2003-version and thus not processed
	 * <p/>
	 * <br>(app.xml, theme.xml ...)
	 *
	 * @param wbh
	 */
	public static void refreshPassThroughFiles( WorkBookHandle wbh )
	{
		try
		{
			// retrieve source zip
			java.util.zip.ZipFile sourceZip = new java.util.zip.ZipFile( wbh.getFile() );
			OOXMLReader.refreshExternalFiles( sourceZip, OOXMLAdapter.getTempDir( wbh.getWorkBook().getFactory().getFileName() ) );
			wbh.setFile( null );
		}
		catch( Exception e )
		{    // wbh.getFile() can be an XLS file (as source) so Exception is almost always OK (do not report)
			//log.error("OOXMLAdapter.refreshPassThroughFiles: could not retrieve source ooxml: " + e.toString());
		}
	}

}

/**
 * helper class for array compares ...
 */
class intArray
{

	private int[] a = null;

	public intArray( int[] a )
	{
		this.a = new int[a.length];
		for( int i = 0; i < a.length; i++ )
		{
			this.a[i] = a[i];
		}
	}

	public intArray( short[] a )
	{
		this.a = new int[a.length];
		for( int i = 0; i < a.length; i++ )
		{
			this.a[i] = a[i];
		}
	}

	public boolean isZero()
	{
		if( a == null )
		{
			return true;
		}
		for( int anA : a )
		{
			if( anA != 0 )
			{
				return false;
			}
		}
		return true;
	}

	public int[] get()
	{
		return a;
	}

	public boolean equals( Object o )
	{
		int[] testa = ((intArray) o).get();
		if( (testa == null) || (a == null) || (testa.length != a.length) )
		{
			return false;
		}
		for( int i = 0; i < a.length; i++ )
		{
			if( a[i] != testa[i] )
			{
				return false;
			}
		}
		return true;
	}
}

class objArray
{
	private Object[] a = null;

	public objArray( Object[] a )
	{
		this.a = a;
        /*
        this.a= new Object[a.length];
        for (int i= 0; i < a.length; i++) {
            this.a[i]= a[i];
        }
        */
	}

	public Object[] get()
	{
		return a;
	}

	public boolean equals( Object o )
	{
		Object[] testa = ((objArray) o).get();
		return java.util.Arrays.equals( a, testa );
	}

}
