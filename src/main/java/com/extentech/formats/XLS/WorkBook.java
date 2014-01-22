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

import com.extentech.ExtenXLS.Cell;
import com.extentech.ExtenXLS.DateConverter;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.ImageHandle;
import com.extentech.ExtenXLS.WorkBookException;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.Theme;
import com.extentech.formats.XLS.charts.Ai;
import com.extentech.formats.XLS.charts.Chart;
import com.extentech.formats.XLS.charts.Fontx;
import com.extentech.formats.XLS.charts.GenericChartObject;
import com.extentech.formats.XLS.formulas.IlblListener;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgArea3d;
import com.extentech.formats.XLS.formulas.PtgExp;
import com.extentech.formats.XLS.formulas.PtgNameX;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.toolkit.FastAddVector;
import com.extentech.toolkit.Logger;

import java.awt.Color;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.TreeMap;
import java.util.Vector;

/**
 * <pre>
 * The WorkBook record represents an XLS workbook substream containing worksheets and associated records.
 *
 * </pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Index
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */

public class WorkBook implements Serializable, XLSConstants, Book
{

	private static final long serialVersionUID = 2282017774412632087L;

	DateConverter.DateFormat dateFormat = DateConverter.DateFormat.LEGACY_1900;

	// public constants
	public static int STRING_ENCODING_AUTO = 0;
	public static int STRING_ENCODING_UNICODE = 1;
	public static int STRING_ENCODING_COMPRESSED = 2;
	public static int ALLOWDUPES = 0;
	public static int SHAREDUPES = 1;

	private int DEBUGLEVEL = 0;
	private int bofct = 0, eofct = 0, indexnum = 0;
	private int defaultIxfe = 15;
	private int CalcMode = CALCULATE_AUTO;
	private int defaultLanguage = 0;    // default language code for the current workbook
	private boolean copying = false;
	private boolean sharedupes = false;
	private AbstractList xfrecs = new Vector();
	private AbstractList indexes = new Vector();
	private AbstractList names = new Vector(); // ALL of the names (including worksheet scoped, etc)
	private AbstractList orphanedPtgNames = new Vector();
	private AbstractList externalnames = new Vector();
	AbstractList formulas = new FastAddVector();
	private AbstractList fonts = new Vector();
	public Color[] colorTable;

	/**
	 * The list of Format records indexed by format ID.
	 */
	private TreeMap formats = new TreeMap();
	private AbstractList charts = new Vector();
	/**
	 * OOXML-specific
	 */
	private AbstractList ooxmlObjects = new Vector();    // stores OOXML objects external to workbook e.g. oleObjects,
	private String ooxmlcodename = null;            // stores OOXML codename
	private ArrayList dxfs = null;            // 20090622 KSC: stores dxf's (incremental style info) per workbook				
	private int firstSheet = 0;                // specifies first sheet (ooxml)
	private Theme theme = null;

	// Reference Tracking
	private ReferenceTracker refTracker = new ReferenceTracker();

	// various
	private AbstractList boundsheets = new Vector();  //TODO:  remove this variable?  its duplicated in workSheets
	private AbstractList hlinklookup = new Vector( 20 );
	private AbstractList mergecelllookup = new Vector( 20 );

	/*Indirect formulas that need to be calculated after load for reftracker*/
	private AbstractList indirectFormulas = new ArrayList();
	private AbstractList supBooks = new Vector();

	/* This is here only to allow client code to compile, should be removed */
	public static String CONVERTMULBLANKS = "deprecated";

	// We should consider using an ordered collections class for some of these?
	// an enumeration of worksheets for instance will not always be in the same
	// order, causing tests to fail...
	private Hashtable workSheets = new Hashtable( 100, .950f );
	private HashMap bookNameRecs = new HashMap();
	private Hashtable formulashash = new Hashtable();

	/**
	 * Maps number format patterns to their IDs.
	 */
	private Hashtable formatlookup = new Hashtable( 30 );

	private Hashtable ptViews = new Hashtable( 20 );// TODO: move to sheet - sheet-level pivot table view recordss    
	private ArrayList ptstream = new ArrayList();        // wb-level pivot cache definitions usually only 1
	private PivotCache ptcache = null;                    // if has pivot tables this is the one and only pivot cache

	public HashMap formatCache = new HashMap();

	private Index lastidx;
	private Sst stringTable;
	private Bof lastBOF;
	private Externsheet myexternsheet;
	private Bof firstBOF;
	private ByteStreamer streamer = createByteStreamer();
	private TabID tabs;
	private Window1 win1;
	private CalcMode calcmoderec;        // determines recalculation mode for workbook - Manual, Auto ...
	private DefaultRowHeight drh;
	private Chart currchart;
	private Ai currai;
	private Supbook myADDINSUPBOOK = null; // for external names SUPBOOK
	WorkBookFactory factory;
	private ContinueHandler contHandler = new ContinueHandler( this );
	private Usersviewbegin usersview;
	private Eof lasteof;
	private BiffRec xl2k = null;
	public AbstractList msodgMerge = new Vector();
	public MSODrawingGroup msodg = null;
	public int lastSPID = 1024;    // 20071030 last or next SPID (= shape ID or image ID); incremented upon new images ... appropriate to store at book level (?) 
	private MSODrawing currdrw = null;

	private BookProtectionManager protector;

	protected void reflectiveClone( WorkBook source )
	{
		for( Field field : WorkBook.class.getDeclaredFields() )
		{
			if( Modifier.isStatic( field.getModifiers() ) )
			{
				continue;
			}

			try
			{
				field.set( this, field.get( source ) );
			}
			catch( IllegalAccessException e )
			{
				throw new RuntimeException( e );
			}
		}
	}

	protected ByteStreamer createByteStreamer()
	{
		return new ByteStreamer( this );
	}

	/**
	 * Gets this sheet's SheetProtectionManager.
	 */
	public BookProtectionManager getProtectionManager()
	{
		if( protector == null )
		{
			protector = new BookProtectionManager( this );
		}
		return protector;
	}

	/**
	 * init the ImageHandles
	 */
	void initImages()
	{
		lastSPID = Math.max( lastSPID,
		                     msodg.getSpidMax() ); // 20090508 KSC: lastSPID should also account for charts [BUGTRACKER 2372 copyChartToSheet error]
		// 20071217 KSC: clear out imageMap before inputting!
		for( int x = 0; x < this.getWorkSheets().length; x++ )
		{
			(this.getWorkSheets()[x]).imageMap.clear();
		}
		for( int i = 0; i < msodg.getMsodrawingrecs().size(); i++ )
		{    // 20070914 KSC: store msodrawingrecs with MSODrawingGroup instead of here
			MSODrawing rec = (MSODrawing) msodg.getMsodrawingrecs().get( i );
			lastSPID = Math.max( lastSPID, rec.getlastSPID() );    // valid for header msodrawing record(s)
			int imgdx = rec.getImageIndex() - 1;    // it's 1-based
			byte[] imageData = msodg.getImageBytes( imgdx );
			if( imageData != null )
			{
				ImageHandle im = new ImageHandle( imageData, rec.getSheet() );
				im.setMsgdrawing( rec );                    //Link 2 actual Msodrawing rec
				im.setName( rec.getName() );                // set image name from rec ...
				im.setShapeName( rec.getShapeName() );    // set shape name as well ...
				im.setImageType( msodg.getImageType( imgdx ) );// 20100519 KSC: added!
				rec.getSheet().imageMap.put( im, imgdx );
			}
		}
	}

	/*
	 * Once all the records are parsed, the msodrawinggroup needs to be parsed.
	 * We first merge all the msodrawinggroup records. From the till now experiment,
	 * there are maximum of two such records, adjacent to each other. If there are two,
	 * the size of the first is only 8228 and the remaining data is in second record.
	 * However, when writing, it doesn't seem important to break it into two and can be written
	 * as one. 
	 *  
	 *  Once msodrawingrecords are parsed, it is not used. However, we have a msodrawingglobal class that 
	 *  records the current count of msodrawing related records, like shape count, max shape id till now etc.
	 *  We need to maintain that for writing (creating the msodrawinggroup records) later on.
	 */
	public void mergeMSODrawingRecords()
	{
		// 20070915 KSC: Now don't re-initialize Msodrawing recs; just merge and parse	
		if( msodg != null )
		{
			msodg.mergeAndRemoveContinues();
			for( int i = 1; i < msodgMerge.size(); i++ )
			{
				msodg.mergeRecords( (MSODrawingGroup) msodgMerge.get( i ) );  // merges and removes secondary msodrawinggroups 
				// 20071003 KSC: get rid of secondary msodg's, will be created upon createMSODGContinues
				this.getStreamer().removeRecord( (MSODrawingGroup) msodgMerge.get( i ) ); // remove existing continues from stream				
			}
			while( msodgMerge.size() > 1 )
			{
				msodgMerge.remove( msodgMerge.size() - 1 );
			}
			msodg.parse();
			this.initImages();
		}
	}

	/**
	 * after changing the MSODrawings on a sheet, this is called to
	 * update the header Msodrawing rec.
	 * <br>must sum up all other mso's on the particular sheet SPCONTAINERLEN and update the sheet mso header ...
	 * <br>this calculation is quite experimental, so far it's working in all known cases ...
	 *
	 * @param bs
	 */
	public void updateMsodrawingHeaderRec( Boundsheet bs )
	{
		MSODrawing msdHeader = msodg.getMsoHeaderRec( bs );
		if( msdHeader != null )
		{
			int spContainerLength = 0;    // count of all other spcontainer lengths (sum=header spgroupcontainer)
			int otherContainerLength = 0;    // count of other containers (solvercontainer,etc) added to dgcontainerlength
			int numshapes = 2;    // 20100324 KSC: total guess, really
			int totalDrawingRecs = msodg.getMsodrawingrecs().size();
			for( int z = 0; z < totalDrawingRecs; z++ )
			{
				MSODrawing rec = (MSODrawing) msodg.getMsodrawingrecs().get( z );
				if( rec.getSheet().equals( bs ) )
				{
					if( (rec != msdHeader) && !rec.isHeader() )
					{    // added header check- seems like *can* have multiple header recs in charts!
						spContainerLength += rec.getSPContainerLength();
						otherContainerLength += rec.getSOLVERContainerLength();
						if( rec.isShape )  // if it's a shape-type mso, count; there are other mso-types that are not SPCONTAINERS; apparently don't count these ...
						{
							numshapes++;
						}
					}
				}
			}
			msdHeader.updateHeader( spContainerLength, otherContainerLength, numshapes, msdHeader.getlastSPID() );
		}
	}

	/**
	 * Get the MsoDrawingGroup for this workbook
	 *
	 * @return msodrawinggroup
	 */
	public MSODrawingGroup getMSODrawingGroup()
	{
		return msodg;
	}

	/**
	 * Set the msodrawinggroup for this workbook
	 *
	 * @param msodg
	 */
	public void setMSODrawingGroup( MSODrawingGroup msodg )
	{
		this.msodg = msodg;
		this.msodg.setWorkBook( this );
		this.msodg.setStreamer( this.getStreamer() );
	}

	/**
	 * For workbooks that do not contain an MSODrawing group create a new one,
	 * if the drawing group already exists return the existing
	 *
	 * @return
	 */
	public MSODrawingGroup createMSODrawingGroup()
	{
		if( msodg != null )
		{
			return msodg;
		}
		this.setMSODrawingGroup( (MSODrawingGroup) MSODrawingGroup.getPrototype() );
		return msodg;
	}

	/**
	 * Return some useful statistics about the WorkBook
	 *
	 * @return
	 */
	public String getStats()
	{
		return getStats( false );
	}

	/**
	 * Return some useful statistics about the WorkBook
	 *
	 * @return
	 */
	public String getStats( boolean usehtml )
	{
		String rex = "\r\n";
		if( usehtml )
		{
			rex = "<br/>";
		}

		String ret = "-------------------------------------------" + rex;
		ret += "ExtenXLS Version:     " + WorkBookHandle.getVersion() + rex;
		ret += "Excel Version:        " + getXLSVersionString() + rex;
		ret += "-------------------------------------------------\r\n";
		ret += "Statistics for:       " + this.toString() + rex;
		ret += "Number of Worksheets: " + this.getNumWorkSheets() + rex;
		ret += "Number of Cells:      " + this.getNumCells() + rex;
		ret += "Number of Formulas:   " + this.getNumFormulas() + rex;
		ret += "Number of Charts:     " + this.getChartVect().size() + rex;
		ret += "Number of Fonts:      " + this.getNumFonts() + rex;
		ret += "Number of Formats:    " + this.getNumFormats() + rex;
		ret += "Number of Xfs:        " + this.getNumXfs() + rex;
		// ret += "StringTable:          " + this.stringTable.toString() + rex;        
		ret += "-------------------------------------------------\r\n";
		return ret;
	}

	public String getXLSVersionString()
	{
		return lastBOF.getXLSVersionString();
	}

	private void addMergedcells( Mergedcells c )
	{
		this.mergecelllookup.add( c );
	}

	private void addHlink( Hlink r )
	{
		this.hlinklookup.add( r );
	}

	/**
	 * take care of any lazy updating before output
	 */
	public void prestream()
	{
		if( this.getMSODrawingGroup() != null )
		{
			this.getMSODrawingGroup().prestream();
		}
	}

	public Chart[] getCharts()
	{
		Chart[] chts = new Chart[this.charts.size()];
		return (Chart[]) charts.toArray( chts );
	}

	public Supbook[] getSupBooks()
	{
		Supbook[] sbs = new Supbook[this.supBooks.size()];
		return (Supbook[]) this.supBooks.toArray( sbs );
	}

	/**
	 * returns the list of OOXML objects which are external or auxillary to the main workbook
	 * e.g. theme, doc properties
	 *
	 * @return
	 */
	public List getOOXMLObjects()
	{
		return ooxmlObjects;
	}

	/**
	 * adds the object-specific signature of the external or auxillary OOXML object
	 * Object should be of String[] form,
	 * key, path, local path  + filename [, rid,  [extra info], [embedded information]]
	 * e.g. theme, doc properties
	 *
	 * @param o
	 */
	public void addOOXMLObject( Object o )
	{
		if( !((String[]) o)[0].equals( "externalLink" ) )
		{
			ooxmlObjects.add( o );
		}
		else
		{
			ooxmlObjects.add( 0, o );    // ensure ExternalLinks are 1st because they are linked via rId in workbook.xml
		}
	}

	/**
	 * return the OOXML theme for this workbook, if any
	 *
	 * @return
	 */
	public Theme getTheme()
	{
		return theme;
	}

	/**
	 * sets the OOXML theme for this 2007 verison workbook
	 *
	 * @param t
	 */
	public void setTheme( Theme t )
	{
		theme = t;
	}

	/**
	 * return the External Supbook record associated with the desired externalWorkbook
	 * will create if bCreate
	 *
	 * @param externalWorkbook String URL (name) of External Workbook
	 * @param bCreate          if true, will create an external SUPBOOK record for the externalWorkbook
	 * @return Supbook
	 */
	public Supbook getExternalSupbook( String externalWorkbook, boolean bCreate )
	{
		Supbook sb = null;
		if( externalWorkbook == null )
		{
			return null;
		}
		for( int i = 0; i < this.supBooks.size(); i++ )
		{
			if( ((Supbook) this.supBooks.get( i )).isExternalRecord() )
			{
				if( externalWorkbook.equalsIgnoreCase( ((Supbook) this.supBooks.get( i )).getExternalWorkBook() ) )
				{
					sb = (Supbook) this.supBooks.get( i );
				}
			}
		}
		if( (sb == null) && bCreate )
		{  // create
			sb = (Supbook) Supbook.getExternalPrototype( externalWorkbook );
			int loc = ((Supbook) supBooks.get( supBooks.size() - 1 )).getRecordIndex();    // must have at least one global supbook present
			streamer.addRecordAt( sb, loc + 1 );    // external supbooks appear to be before "normal" supbooks [BugTracker 1434]
			this.supBooks.add( sb );    // 20080714 KSC: add at beginning -- correct in all cases? 
//          this.addRecord(sb,false); // "" no need
		}
		return sb;
	}

	/**
	 * return the index into the Supbook records for this supbook
	 *
	 * @param sb
	 * @return int
	 */
	public int getSupbookIndex( Supbook sb )
	{
		for( int i = 0; i < this.supBooks.size(); i++ )
		{
			if( this.supBooks.get( i ) == sb )
			{
				return i;
			}
		}
		return -1;
	}

	public AbstractList getChartVect()
	{
		return charts;
	}

	public Sxview getPivotTableView( String nm )
	{
		return (Sxview) ptViews.get( nm );
	}

	protected void addPivotTable( Sxview sx )
	{
		this.ptViews.put( sx.getTableName(), sx );    // Pivot Table View ==Top-level record for a Pivot Table
	}

	/**
	 * return all pivot table views (==Sxview records)
	 * <br>SxView is the top-level record of a Pivot Table
	 * as distinct from the PivotCache (stored data source in a LEOFile Storage)
	 * and PivotTable Stream (SxStream top-level record)
	 *
	 * @return
	 */
	public Sxview[] getAllPivotTableViews()
	{
		Sxview[] sv = new Sxview[ptViews.size()];
		Enumeration x = ptViews.elements();
		int t = 0;
		while( x.hasMoreElements() )
		{
			sv[t++] = (Sxview) x.nextElement();
		}
		return sv;
	}

	public int getNPivotTableViews()
	{
		return ptViews.size();
	}

	/**
	 * return the Externsheet for this book
	 *
	 * @param create a new Externsheet if it does not exist
	 * @return the Externsheet
	 */
	public Externsheet getExternSheet( boolean create )
	{
		if( (myexternsheet == null) && create )
		{
			addExternsheet();
		}
		return myexternsheet;
	}

	/**
	 * get the Externsheet
	 */
	public Externsheet getExternSheet()
	{
		if( myexternsheet == null )
		{
			addExternsheet();
		}
		return myexternsheet;
	}

	/**
	 * default constructor -- do init
	 */
	public WorkBook()
	{

		Object cm = System.getProperties().get( WorkBook.CALC_MODE_PROP );
		if( cm != null )
		{
			try
			{
				this.CalcMode = Integer.parseInt( cm.toString() );
			}
			catch( Exception e )
			{
				Logger.logWarn( "Invalid Calc Mode Setting in System properties:" + cm );
			}
		}
		if( System.getProperties().get( "com.extentech.ExtenXLS.sharedupes" ) != null )
		{
			this.sharedupes = System.getProperties().get( "com.extentech.ExtenXLS.sharedupes" ).equals( "true" );
			if( this.sharedupes )
			{
				this.setDupeStringMode( WorkBook.SHAREDUPES );
			}
		}
		this.initBuiltinFormats();
		// re-init color table: initial state of color table if Pallete record exists, changes may occur
		colorTable = new java.awt.Color[FormatHandle.COLORTABLE.length];
		for( int i = 0; i < FormatHandle.COLORTABLE.length; i++ )
		{
			colorTable[i] = FormatHandle.COLORTABLE[i];
		}
	}

	/**
	 * Gets the format ID for a given number format pattern.
	 * This lookup is completely case-insensitive. For most patterns this
	 * correctly reflects the case-insensitivity of the tokens. Custom patterns
	 * containing string literals could be matched incorrectly.
	 *
	 * @param pattern the number format pattern to look up
	 * @return the format ID of the given pattern or -1 if it's not recognized
	 */
	public short getFormatId( String pattern )
	{
		Short res = (Short) formatlookup.get( pattern.toUpperCase() );
		if( res != null )
		{
			return res;
		}
		return -1;
	}

	/**
	 * Initializes the format lookup to contain the built-in formats.
	 */
	private void initBuiltinFormats()
	{
		String[][] formats = FormatConstantsImpl.getBuiltinFormats();

		for( String[] format : formats )
		{
			formatlookup.put( format[0].toUpperCase(), Short.valueOf( format[1], 16 ) );
		}
	}

	/**
	 * Init names at Post-load
	 */
	public void initializeNames()
	{
		for( int i = 0; i < names.size(); i++ )
		{
			Name n = (Name) names.get( i );
			n.parseExpression();    // evaluate expression at postload, after sheet recs are loaded
		}
	}

	/**
	 * add a Name object to the collection of names
	 */
	public int addName( Name n )
	{
		if( n.getItab() != 0 )
		{
			// its a sheet level name
			try
			{
				Boundsheet b = this.getWorkSheetByNumber( n.getItab() - 1 );// one based pointer
				b.addLocalName( n );
				n.setSheet( b );
			}
			catch( WorkSheetNotFoundException e )
			{
			}
		}
		else
		{
			String sName = n.getNameA();    // returns upper case name
			Object existo = bookNameRecs.get( sName );
			if( existo != null )
			{ // handle duplicate named ranges
				String bnam = n.toString();
				if( bnam.indexOf( "Built-in:" ) != 0 )
				{
					try
					{
						if( ((Name) existo).getLocation() != null )  // use original - as good a guess as any
						{
							return -1; // an invalid sheet
						}
					}
					catch( Exception e )
					{
					}
					// if original does not have a location set, use this one instead
					this.names.remove( this.names.indexOf( existo ) );
					this.bookNameRecs.remove( sName );
				}
			}
			this.bookNameRecs.put( sName, n );
		}
		this.names.add( n );
		if( (myexternsheet != null) && (n.getExternsheet() == null) )
		{
			try
			{
				n.setExternsheet( myexternsheet ); // update sheet reference
			}
			catch( WorkSheetNotFoundException e )
			{
				Logger.logWarn( "WorkBookHandle.addName() setting Externsheet failed for new Name: " + e.toString() );
			}
		}
		return names.size() - 1;
	}

	public void addNameUpdateSheetRefs( Name n, String origWorkBookName )
	{
		if( bookNameRecs.get( n.getNameA() ) == null )
		{
			Name newName = new Name( this, n.getName() );
			try
			{
				newName.setLocation( n.getLocation() );
			}
			catch( Exception e )
			{
			}
			if( (myexternsheet != null) && (newName.getExternsheet() == null) )
			{
				try
				{
					newName.setExternsheet( myexternsheet ); // update sheet reference
				}
				catch( WorkSheetNotFoundException e )
				{
					Logger.logWarn( "WorkBookHandle.addName() setting Externsheet failed for new Name: " + e.toString() );
				}
			}
			newName.updateSheetRefs( origWorkBookName );
		}
	}

	/**
	 * Store an external name
	 *
	 * @param n = String describing the name
	 * @return int location of the name
	 */
	public int addExternalName( String n )
	{
		this.externalnames.add( n );
		return externalnames.size();    // one-based index
	}

	/**
	 * Get a string array of external names
	 *
	 * @return externalNames
	 */
	public String[] getExternalNames()
	{
		String[] n = new String[externalnames.size()];
		externalnames.toArray( n );
		return n;
	}

	/**
	 * Get the external name at the specified index.
	 *
	 * @param t index of the name
	 * @return name at the index, empty string if it doesn't exist.
	 * <p/>
	 * Why are we calling getExternalName one based, then removing for ordinal,  internal processes should always
	 * be 0,1,2,3...  -NR 1/06
	 */
	public String getExternalName( int t )
	{
		if( t > 0 )
		{
			return (String) this.externalnames.get( t - 1 );    // one-based index
		}
		return "";
	}

	/**
	 * For workbooks that do not contain an externsheet
	 * this creates the externsheet record with one 0000 record
	 * and the related Supbook rec
	 */
	public void addDefaultExternsheet()
	{
		Supbook sbb = (Supbook) Supbook.getPrototype( this.getNumWorkSheets() );
		int l = this.stringTable.getRecordIndex();
		streamer.addRecordAt( sbb, l++ );
		supBooks.add( sbb );
		Externsheet ex = (Externsheet) Externsheet.getPrototype( 0x0000, 0x0000, this );
		streamer.addRecordAt( ex, l );
		this.addRecord( ex, false );
	}

	/**
	 * apparently this method adds an External name rec and returns the ilbl
	 * <p/>
	 * Correct structure is
	 * Supbook
	 * Externname
	 * Supbook
	 * Externsheet
	 *
	 * @param s
	 * @return
	 * @see PtgNameX
	 */
	public int getExtenalNameNumber( String s )
	{
		int i = externalnames.indexOf( s );
		if( i > -1 ) // got it
		{
			return i + 1;
		}
		if( this.getExternSheet() == null )
		{
			this.addDefaultExternsheet();
		}

		// not found; add a new EXTERNNAME record to list of add-ins
		int n = addExternalName( s );
		try
		{
			int loc;
			if( myADDINSUPBOOK == null )
			{
				Supbook sb = (Supbook) Supbook.getAddInPrototype();
				loc = this.getExternSheet().getRecordIndex();
				streamer.addRecordAt( sb, loc++ );
				this.addRecord( sb, false );
				myADDINSUPBOOK = sb;
				supBooks.add( sb );
				int externref = this.getExternSheet().getVirtualReference();
				// Add EXTERNNAME record after ADD-IN SUPBOOK record and after existing EXTERNNAME records
				Externname exn = (Externname) Externname.getPrototype( s );
				streamer.addRecordAt( exn, loc++ );
				this.addRecord( exn, false );
			}
			else
			{
				loc = streamer.getRecordIndex( myADDINSUPBOOK );
				// Add EXTERNNAME record after ADD-IN SUPBOOK record and after existing EXTERNNAME records
				Externname exn = (Externname) Externname.getPrototype( s );
				streamer.addRecordAt( exn, loc + externalnames.size() );
				this.addRecord( exn, false );
			}

		}
		catch( Exception e )
		{
			Logger.logWarn( "Error adding externname: " + e );
		}
		return n;
	}

	/**
	 * Get a collection of all names in the workbook
	 */
	public Name[] getNames()
	{
		Name[] n = new Name[names.size()];
		names.toArray( n );
		return n;
	}

	/**
	 * Get a collection of all names in the workbook
	 */
	public Name[] getWorkbookScopedNames()
	{
		ArrayList a = new ArrayList( this.bookNameRecs.values() );
		Name[] n = new Name[a.size()];
		a.toArray( n );
		return n;
	}

	/**
	 * returns the List of Formulas in the book
	 *
	 * @return
	 */
	public List getFormulaList()
	{
		return formulas;
	}

	/**
	 * returns the array of Formulas in the book
	 *
	 * @return
	 */
	public Formula[] getFormulas()
	{
		Formula[] n = new Formula[formulas.size()];
		formulas.toArray( n );
		return n;
	}

	/**
	 * remove a formula from the book
	 *
	 * @param fmla
	 */
	public void removeFormula( Formula fmla )
	{
		this.formulashash.remove( fmla.getCellAddressWithSheet() );
		formulas.remove( fmla );
		fmla.destroy();
	}

	/**
	 * Returns the recalculation mode for the Workbook:
	 * <br>0= Manual
	 * <br>1= Automatic
	 * <br>2= Automatic except for multiple table operations
	 *
	 * @return int
	 */
	public int getRecalculationMode()
	{
		return this.calcmoderec.getRecalcuationMode();
	}

	/**
	 * Sets the recalculation mode for the Workbook:
	 * <br>0= Manual
	 * <br>1= Automatic
	 * <br>2= Automatic except for multiple table operations
	 */
	public void setRecalcuationMode( int mode )
	{
		this.calcmoderec.setRecalculationMode( mode );
	}

	/**
	 * returns a Named range by number
	 *
	 * @param t
	 * @return
	 */
	public Name getName( int t )
	{
		return (Name) this.names.get( t - 1 );
	}

	/**
	 * rename the NamedRange in the lookup map
	 *
	 * @param t
	 * @return
	 */
	public void setNewName( String oldname, String newname )
	{
		if( oldname.equals( newname ) )
		{
			return;
		}
		oldname = oldname.toUpperCase();        // case-insensitive
		newname = newname.toUpperCase();        // 	""
		Object old = bookNameRecs.get( oldname );
		if( old == null )
		{
			return; // new name?
		}
		bookNameRecs.remove( oldname );
		bookNameRecs.put( newname, old );

	}

	/**
	 * Re-assocates ptgrefs that are pointing to a name that has been deleted then
	 * is recreated
	 *
	 * @param name
	 */
	public void associateDereferencedNames( Name name )
	{
		Iterator i = orphanedPtgNames.iterator();
		String theName = name.getName();
		while( i.hasNext() )
		{
			IlblListener x = (IlblListener) i.next();
			if( x.getStoredName().equalsIgnoreCase( theName ) )
			{
				x.setIlbl( (short) this.getNameNumber( theName ) );
				x.addListener();
			}
		}
	}

	/**
	 * returns a named range by name string
	 * <p/>
	 * This method will first attempt to look in the book names, then the sheet names,
	 * obviously different scoped names can have the same identifying name, so this could return
	 * one of multiple names if this is the case
	 *
	 * @param t
	 * @return
	 */
	public Name getName( String nameRef )
	{
		nameRef = nameRef.toUpperCase();    // case-insensitive
		Object o = this.bookNameRecs.get( nameRef );
		if( o == null )
		{
			Boundsheet[] shts = this.getWorkSheets();
			for( Boundsheet sht : shts )
			{
				o = sht.getName( nameRef );
				if( o != null )
				{
					return (Name) o;
				}
			}
		}
		return (Name) o;
	}

	/**
	 * returns a scoped named range by name string
	 *
	 * @param t
	 * @return
	 */
	public Name getScopedName( String nameRef )
	{
		Object o = this.bookNameRecs.get( nameRef.toUpperCase() );    // case-insensitive
		if( o == null )
		{
			return null;
		}
		return (Name) o;
	}

	/**
	 * Returns the ilbl of the name record associated with the string passed in.
	 * If the name does not exist, it get's created without a location reference.
	 * This is needed to support formula creation with non-existent names referenced.
	 *
	 * @param t, the name record to search for
	 * @return the index of the name
	 */
	public int getNameNumber( String nameStr )
	{
		for( int i = 0; i < names.size(); i++ )
		{
			Name n = (Name) this.names.get( i );
			if( n.getName().equalsIgnoreCase( nameStr ) )
			{
				return i + 1;
			}
		}
		// no name exists, we need to create one.
		Name myName;
		myName = new Name( this, nameStr );
//        myName = new Name(this, true);
//        myName.setName(nameStr);
		Name[] nmx = this.getNames();
		int namepos = -1;
		if( nmx.length >= 1 )
		{
			namepos = nmx[nmx.length - 1].getRecordIndex();
		}
		else
		{
			namepos = this.getExternSheet().getRecordIndex();
		}
		namepos++;
		this.getStreamer().addRecordAt( myName, namepos );
		return this.getNameNumber( nameStr );
	}

	/**
	 * Get's the index for this particular front.
	 * <p/>
	 * NOTE:  this doesn't actually get a "matching" font, it has to be the exact font.
	 * 20070826 KSC: changed to match font characterstics, not just return exact matching font
	 */
	public int getFontIdx( Font f )
	{
		// 20070819 KSC: Try this to see if better! Matches 6 key attributes (size, name, color, etc.)
		for( int i = fonts.size() - 1; i >= 0; i-- )
		{    // start from the back so don't initially match defaults...
			if( f.matches( (Font) fonts.get( i ) ) )
			{
				return ((i > 3) ? (i + 1) : i);
			}
		}
		// return  fonts.indexOf(f);
		return -1;
	}

	/**
	 * Get's the index for this font, based on matching through
	 * xml strings.  If the font doesn't exist in the book it returns -1;
	 *
	 * @return KSC: is this method necessary now with above getFontIdx changes?
	 */
	public int getMatchingFontIndex( Font f )
	{
		Map fontmap = this.getFontRecsAsXML();
		Object o = fontmap.get( "<FONT><" + f.getXML() + "/></FONT>" );
		if( o != null )
		{
			Integer I = (Integer) o;
			return I;
		}
		return -1;
	}

	/**
	 * InsertFont inserts a Font record into the workbook level stream,
	 * For some reason, the addFont only puts it into an array that is never accessed
	 * on output.  This may have a reason, so I am not overwriting it currently, but
	 * let's check it out?
	 */
	public int insertFont( Font x )
	{
		int insertIdx = this.getFont( this.getNumFonts() ).getRecordIndex();
		// perform default add rec actions
		this.getStreamer().addRecordAt( x, insertIdx + 1 );
		x.setIdx( -1 ); // flag to add into font array
		this.addRecord( x, false );    // also adds to font array so no need for additional addFont below 
		return fonts.indexOf( x );
	}

	/**
	 * add a Font object to the collection of Fonts
	 */
	public int addFont( Font f )
	{
		fonts.add( f );
		if( fonts.size() > 4 )    // fake the evil 4!
		{
			return fonts.size();
		}
		return fonts.size() - 1;
	}

	public int getNumFonts()
	{
		return fonts.size();
	}

	/**
	 * Get the font at the specified index.  Note that the number 4 does not exist, so index correctly based of that.
	 * <p/>
	 * So,  if you call getFont(5), you are really doing getFont(4) from the internal array
	 *
	 * @param t
	 * @return
	 */
	public Font getFont( int t )
	{
		if( t >= 4 )
		{
			t--;
		}
		if( this.fonts.size() >= t )
		{
			if( t >= fonts.size() )
			{
				Logger.logWarn( "font " + t + " not found. Workbook contains only: " + fonts.size() + " defined fonts." );
				return (Font) fonts.get( 0 );
			}
			return (Font) fonts.get( t );
		}
		return (Font) this.fonts.get( 0 );
	}

	/**
	 * Inserts a newly created Format record into the workbook.
	 * This method handles assigning the format ID and adding the record to the
	 * workbook. If the record is already part of the workbook use
	 * {@link #addFormat} instead.
	 */
	public int insertFormat( Format format )
	{
		Format last;
		try
		{
			last = (Format) formats.get( formats.lastKey() );
		}
		catch( NoSuchElementException e )
		{
			/* There are no other Format records in the workbook.
			 * This shouldn't happen because most (all?) Excel files contain
    		 * Format records for the locale-specific (and thus not implied)
    		 * built-in formats. If it does happen, either we need to re-assess
    		 * the above assumption or this method was called before the Format
    		 * records were parsed. Either way we need to know about it. 
    		 */
			throw new AssertionError( "WorkBook.insertFormat called but no " +
					                          "Format records exist. This should not happen. Please " +
					                          "report this error to support@extentech.com." );
		}

		// Add it to the streamer and workbook
		getStreamer().addRecordAt( format, last.getRecordIndex() + 1 );
		addRecord( format, false );

		// Give it a format ID
		if( format.getIfmt() == -1 )
		{
			format.setIfmt( (short) Math.max( last.getIfmt() + 1, 164 ) );
		}

		// Add it to the format lookups
		addFormat( format );

		return format.getIfmt();
	}

	public Formula getFormula( String cellAddress ) throws FormulaNotFoundException
	{
		Formula formula = (Formula) formulashash.get( cellAddress );
		if( formula == null )
		{
			throw new FormulaNotFoundException( "no formula found at " + cellAddress );
		}

		return (formula);
	}

	/**
	 * Adds an existing format record to the list of known formats.
	 * This method does not add the record to the workbook! If the format is
	 * not already in the workbook use {@link #insertFormat} instead.
	 *
	 * @param format the Format record to add
	 */
	public int addFormat( Format format )
	{
		Short ifmt = format.getIfmt();

		// Add it to the format record lookup
		formats.put( ifmt, format );

		// Add it to the format string lookup
		formatlookup.put( format.getFormat().toUpperCase(), ifmt );

		return format.getIfmt();
	}

	/**
	 * Gets the number of custom number formats registered on this book.
	 */
	public int getNumFormats()
	{
		return formats.size();
	}

	/**
	 * Gets a custom number format by its format ID.
	 */
	public Format getFormat( int id )
	{
		return (Format) formats.get( Short.valueOf( (short) id ) );
	}

	public TabID getTabID()
	{
		return this.tabs;
	}

	/**
	 * associate default row/col size recs
	 */
	void setDefaultRowHeightRec( DefaultRowHeight dr )
	{
		this.drh = dr;
	}

	/**
	 * set Default row height in twips (1/20 of a point)
	 */
	// should be a double as Excel units are 1/20 of what is stored in defaultrowheight 
	// e.g. 12.75 is Excel Units, twips =  12.75*20 = 256 (approx)
	// should expect users to use Excel units and target method do the 20* conversion
	public void setDefaultRowHeight( int t )
	{
		drh.setDefaultRowHeight( t );
	}

	/**
	 * set Default col width for all worksheets in the workbook,
	 * <p/>
	 * Default column width can also be set on individual worksheets
	 */
	public void setDefaultColWidth( int t )
	{
		Boundsheet[] b = this.getWorkSheets();
		for( Boundsheet aB : b )
		{
			aB.setDefaultColumnWidth( t );
		}
	}

	/**
	 * sets the selected worksheet
	 */
	public int getSelectedSheetNum()
	{
		return win1.getCurrentTab();
	}

	/**
	 * sets the selected worksheet
	 */
	public void setSelectedSheet( Boundsheet bs )
	{
		Boundsheet[] bsx = this.getWorkSheets();
		for( Boundsheet aBsx : bsx )
		{
			if( aBsx != bs )
			{
				aBsx.setSelected( false );
			}
		}
		this.win1.setCurrentTab( bs );
	}

	Formula lastFormula = null;
	XLSRecord countryRec = null;
	boolean inChartSubstream = false;
	ArrayList chartTemp = new ArrayList();

	/**
	 * for those cases where a formula calculation adds a new string rec
	 * need to explicitly set lastFormula before calling addRecord
	 *
	 * @param f
	 */
	protected void setLastFormula( Formula f )
	{
		lastFormula = f;
	}

	/**
	 * Associate a record with its containers and related records.
	 */
	@Override
	public BiffRec addRecord( BiffRec rec, boolean addtorec )
	{
		short opc = rec.getOpcode();
		rec.setStreamer( streamer );
		rec.setWorkBook( this );

		Boundsheet bs = null;
		Long lbplypos = 0l;

		// get the relevant Boundsheet for this rec
		if( rec instanceof Bof )
		{
			if( this.getFirstBof() == null )
			{
				this.setFirstBof( (Bof) rec );
			}
			if( bofct == eofct )
			{ // not a chart or other non Sheet Bof
				this.setLastBOF( (Bof) rec );
			}
			if( ((Bof) rec).isChartBof() )
			{
				inChartSubstream = true;
			}
		}

		if( this.lastBOF == null )
		{
			Logger.logWarn( "WorkBook: NULL Last BOF" );
		}
		long lb = this.lastBOF.getLbPlyPos();
		if( !this.lastBOF.isValidBIFF8() )
		{
			lb += (long) 8;
		}
		lbplypos = lb; // use last

		bs = getSheetFromRec( rec, lbplypos );
		if( bs != null )
		{
			lbplypos = bs.getLbPlyPos();
		}

		if( (bs != null) )
		{ // &&){ 
			lastbound = bs;
			if( addtorec )
			{
				rec.setSheet( bs );// we don't include Bof or other Book-recs because it lives in the Streamer recvec
			}
			if /*((rec.isValueForCell()) 
            && */( !copying )
			{
				if( (lastFormula != null) && (opc == STRINGREC) )
				{
					lastFormula.addInternalRecord( rec );
				}
				else if( (lastFormula != null) && (opc == ARRAY) )
				{
					lastFormula.addInternalRecord( rec );
				}
				else if( rec.isValueForCell() )
				{
					if( currchart == null )
					{
						bs.addCell( (CellRec) rec );
					}

				}
			}
		}

		if( inChartSubstream )
		{
			if( currchart == null )
			{
				if( rec.getOpcode() == CHART )
				{
					charts.add( rec );
					if( bs != null )
					{
						bs.addChart( (Chart) rec );
					}
					currchart = (Chart) rec;
					currchart.setPreRecords( chartTemp );
					chartTemp = new ArrayList();    // clear out
				}
				else
				{
					chartTemp.add( rec );
				}
			}
			else
			{
				currchart.addInitialChartRecord( rec );
				if( rec.getOpcode() == EOF )
				{
					currchart.initChartRecords();    // finished
					currchart = null;
					inChartSubstream = false;
				}
			}
			addtorec = false;
		}

		// Rows, valrecs, dbcells, and muls are stored in the row, not the byte streamer
		if( (opc == XLSRecord.DBCELL) || (opc == XLSRecord.ROW) || rec.isValueForCell() || (opc == XLSRecord.MULRK) || (opc == CHART) || (opc == XLSRecord.FILEPASS) || (opc == XLSRecord.SHRFMLA) || (opc == XLSRecord.ARRAY) || (opc == XLSRecord.STRINGREC) )
		{
			addtorec = false;
		}

		// add it to the record stream
		if( addtorec )
		{

			if( lbplypos > 0 )
			{
				streamer.addRecord( rec );
			}
			else
			{
				streamer.records.add( rec );
			}
		}

		switch( opc )
		{
			case AUTOFILTER:
				bs.getAutoFilters().add( rec );
				break;

			case CONDFMT:
				bs.getConditionalFormats().add( rec );
				((Condfmt) rec).initializeReferences();
				break;

			case CF:
				Condfmt cfmt = (Condfmt) bs.getConditionalFormats().get( bs.getConditionalFormats().size() - 1 );
				cfmt.addRule( (Cf) rec );
				break;

			case MERGEDCELLS:
				bs.addMergedCellsRec( (Mergedcells) rec );
				this.addMergedcells( (Mergedcells) rec );
				break;

			// give protection records to the relevant ProtectionManager
			case PASSWORD:
			case PROTECT:
			case PROT4REV:
			case OBJPROTECT:
			case SCENPROTECT:
			case FEATHEADR:
				ProtectionManager manager;
				if( bs != null )
				{
					manager = bs.getProtectionManager();
				}
				else
				{
					manager = this.getProtectionManager();
				}
				manager.addRecord( rec );
				break;

			case DVAL:
				if( bs != null )
				{
					bs.setDvalRec( (Dval) rec );
				}
				break;

			case DV:
				if( bs != null )
				{
					if( bs.getDvalRec() != null )
					{
						bs.getDvalRec().addDvRec( (Dv) rec );
					}
				}
				break;

			case INDEX:
				Index id = (Index) rec;
				id.setIndexNum( indexnum++ );
				indexes.add( indexes.size(), id );
				this.setLastINDEX( id );
				if( bs == null )
				{
					Logger.logWarn( "ERROR: WorkBook.addRecord( Index ) error: BAD LBPLYPOS.  The wrong LB:" + lbplypos );
					try
					{
						bs = this.getWorkSheetByNumber( indexnum - 1 );
						Logger.logInfo( " The RIGHT LB:" + bs.getLbPlyPos() );
					}
					catch( WorkSheetNotFoundException e )
					{
						Logger.logInfo( "problem getting WorkSheetByNumber: " + e );
					}

				}
				bs.setSheetIDX( id );
				break;

			case ROW:
				Row rw = (Row) rec;
				if( bs != null )
				{
					bs.addRowRec( rw );
				}

				break;

			case FORMULA:
				this.addFormula( (Formula) rec );
				lastFormula = (Formula) rec;
				break;

			case ARRAY:
				Array arr = (Array) rec;
				if( bs != null )
				{
					bs.addArrayFormula( arr );
				}
				arr.setParentRec( lastFormula );    // [BugTracker 1869] link array formula to it's parent formula rec
				break;
                	
/*                case SHRFMLA : done in shrfmla.init
                    Shrfmla form = (Shrfmla) rec;
                    try{ // throws exceptipon during pullparse
                    	form.setHostCell( lastFormula );
                    }catch(Exception e){;}
                    break;*/

			case DATE1904:
				if( ((NineteenOhFour) rec).is1904 )
				{
					this.dateFormat = DateConverter.DateFormat.LEGACY_1904;
				}
				break; 
                    
                /*case PALETTE : // palette now correctly read into COLORTABLE
                    this.pal = (Palette) rec;
                    break;*/

			case HLINK:
				Hlink hl = (Hlink) rec;
				this.addHlink( hl );
				break;

			case DSF:
				Dsf dsf = (Dsf) rec;
				if( dsf.fDSF == 1 )
				{
					Logger.logErr( "DOUBLE STREAM FILE DETECTED!" );
					Logger.logErr( "  ExtenXLS is compatible with Excel 97 and above only." );
					throw new WorkBookException( "ERROR: DOUBLE STREAM FILE DETECTED!  ExtenXLS is compatible with Excel 97 + only.",
					                             WorkBookException.DOUBLE_STREAM_FILE );
				}
				break;

			case GUTS:
				if( bs != null )
				{
					bs.setGuts( (Guts) rec );
				}
				break;

			case DBCELL:
				break;

			case BOF:
				if( DEBUGLEVEL > 5 )
				{
					Logger.logInfo( "BOF:" + bofct + " - " + rec );
				}
				if( eofct == bofct )
				{
					if( bs != null )
					{
						bs.setBOF( (Bof) rec );
					}
				}
				bofct++;
				break;

			case EXTERNSHEET:
				myexternsheet = (Externsheet) rec;
				break;

			case DEFCOLWIDTH:
				if( bs != null )
				{
					bs.setDefColWidth( (DefColWidth) rec );
				}
				break;

			case EOF:
				this.lasteof = (Eof) rec;
				eofct++;
				if( eofct == bofct )
				{
					if( bs != null )
					{
						bs.setEOF( (Eof) rec );
					}
					eofct--;
					bofct--;
				}
				break;

			case SELECTION: // only used for Recvec index
				bs.setLastselection( (Selection) rec );
				break;

			case COUNTRY:
				// Added to save position of 1st bound sheet, which is 1 record
				// before COUNTRY RECORD (= 2 before 1st SUPBOOK record - true in all cases?)
				countryRec = (XLSRecord) rec;
				// USA=1, Canada=1, Japan=81, China=86, Thailand= 66, Korea= 82, India=91 ... 
				this.defaultLanguage = ((Country) rec).getDefaultLanguage();
				break;

			case SUPBOOK:   // KSC: must store ordinal positions of SupBooks, for adding Externsheets
				supBooks.add( rec );
				if( myADDINSUPBOOK == null )
				{ // see if this is the ADD-IN SUPBOOK rec
					Supbook sb = (Supbook) rec;
					if( sb.isAddInRecord() )
					{
						myADDINSUPBOOK = sb;
					}
				}
				break;

			case BOUNDSHEET:
				Boundsheet sh = (Boundsheet) rec;
                    
                    /*  Here we need to set the selected variable,
                        but not mess with selected tabs
                        when all of the sheets aren't in the book yet.
                        -jm 
                    */
				int ctab = 1;    // default- select 1st sheet if no Windows1 record
				if( win1 != null )
				{// Windows1 record is optional 20101004 TestCorruption.TestNPEOnOpen 
					ctab = win1.getCurrentTab();
				}
				int shts = boundsheets.size();
				if( ctab == shts )
				{
					sh.selected = true;
				}

				this.addWorkSheet( sh.getLbPlyPos(), sh );
				break;

			case MULRK:
				Mulrk mul = (Mulrk) rec;
				Iterator xit = mul.getRecs().iterator();
				while( xit.hasNext() )
				{
					this.addRecord( ((Rk) xit.next()), false );
				}
				break;

			case SST:
				this.setSharedStringTable( (Sst) rec );
				break;

			case EXTSST:
				((Extsst) rec).setSst( this.getSharedStringTable() );
				break;

			case SXSTREAMID:
				this.ptstream.add( rec );    // Pivot Stream                   
				break;

			case SXVS:
			case DCONREF:
			case DCONNAME:
			case DCONBIN:
				try
				{
					SxStreamID sid = (SxStreamID) ptstream.get( ptstream.size() - 1 );
					sid.addSubrecord( rec );
				}
				catch( Exception e )
				{
				}
				break;

			case SXVIEW:
				addPivotTable( ((Sxview) rec) ); // Pivot Table View ==Top-level record for a Pivot Table                    
				break;

			// all* possible records associated with SxView (=PivotTable View)	(*hopefully)
			case SXVD:
				//case SXVI:	// subrecords of SxVD
				//case SXVDEX:
			case SXIVD:
			case SXPI:
			case SXDI:
			case SXLI:
			case SXEX:
			case SXVIEWEX9:
			case QSISXTAG:
				try
				{
					Sxview sx = (Sxview) this.ptViews.values().toArray()[this.ptViews.size() - 1];
					sx.addSubrecord( rec );
				}
				catch( Exception e )
				{
					;
				}
				break;

			case TABID:
				this.tabs = (TabID) rec;
				break;

			case NAME:
				addName( (Name) rec );
				break;

			case CALCMODE:
				this.calcmoderec = (CalcMode) rec;
				break;

			case WINDOW1:
				this.win1 = (Window1) rec;
				break;

			case WINDOW2:
				if( bs != null )
				{
					bs.setWindow2( (Window2) rec );
				}
				break;

			case SCL: // scl is for zoom
				if( bs != null )
				{
					bs.setScl( (Scl) rec );
				}
				break;

			case PANE:
				if( bs != null )
				{
					bs.setPane( (Pane) rec );
				}
				break;
			case EXCEL2K:
				xl2k = rec;
				break;

			case PHONETIC:
				// TODO: this isn't necessary anymore!  look at and remove 
				if( this.currdrw != null )
				{
					this.currdrw.setMystery( (Phonetic) rec );
				}
				break;

			case MSODRAWINGGROUP:
				if( msodg == null )
				{
					msodg = (MSODrawingGroup) rec;
				}
				msodgMerge.add( rec );

				break;

			case MSODRAWING:
				rec.setSheet( bs );
				if( msodg != null )
				{
					msodg.addMsodrawingrec( (MSODrawing) rec );
				}
				else
				{
					;
				}
				// do what???System.out.println("PROBLEM with MSODG!");
				break;

			case COLINFO:
				bs.addColinfo( (Colinfo) rec );
				break;

			case USERSVIEWBEGIN:
				this.usersview = (Usersviewbegin) rec;
				break;

			case WSBOOL:
				if( bs != null )
				{
					bs.setWsBool( (WsBool) rec );
				}
				break;

			// Handle continue records which are actually masked Mso's
			case CONTINUE:
				if( ((Continue) rec).maskedMso != null )
				{
					((Continue) rec).maskedMso.setSheet( bs );
					if( msodg != null )
					{
						msodg.addMsodrawingrec( ((Continue) rec).maskedMso );
					}
				}
				break;

			case XF:
				try
				{
					this.addXf( (Xf) rec );
				}
				catch( Exception e )
				{
					;// throws exceptions during PullParse
				}
				break;

			default:
				// DO NOTHING

		}

		// finish up
		rec.setIndex( getLastINDEX() );
		rec.setXFRecord();
		return rec;
	}

	/**
	 * Dec 15, 2010
	 *
	 * @param rec
	 * @return
	 */
	@Override
	public Boundsheet getSheetFromRec( BiffRec rec, Long lbplypos )
	{
		Boundsheet bs = null;

		if( rec.getSheet() != null )
		{
			bs = rec.getSheet();
		}
		else if( lbplypos != null )
		{
			bs = this.getWorkSheet( lbplypos );
		}
		else
		{
			bs = lastbound;
		}
		return bs;
	}

	/**
	 * get a handle to the ContinueHandler
	 */
	@Override
	public ContinueHandler getContinueHandler()
	{
		return this.contHandler;
	}

	/**
	 * get a handle to the Usersviewbegin for the workbook
	 */
	public Usersviewbegin getUsersviewbegin()
	{
		if( usersview == null )
		{
			usersview = new Usersviewbegin();
			streamer.addRecord( usersview );
			this.addRecord( usersview, false );
		}
		return usersview;
	}

	/**
	 * set the Debug level
	 */
	@Override
	public void setDebugLevel( int i )
	{
		this.DEBUGLEVEL = i;
	}

	/**
	 * get a handle to the Reader for this
	 * WorkBook.
	 */
	@Override
	public void setFactory( WorkBookFactory r )
	{
		this.factory = r;
	}

	/**
	 * Get the number of worksheets in this WorkBook
	 */
	public int getNumWorkSheets()
	{
		return this.boundsheets.size();
	}

	/**
	 * get the number of formulas in this WorkBook
	 *
	 * @return
	 */
	public int getNumFormulas()
	{
		return formulas.size();
	}

	/**
	 * get the number of Cells in this WorkBook
	 */
	public int getNumCells()
	{
		int cellnum = 0;
		Enumeration e = workSheets.elements();
		while( e.hasMoreElements() )
		{
			Boundsheet b = (Boundsheet) e.nextElement();
			cellnum += b.getNumCells();
		}
		return cellnum;
	}

	/**
	 * get all of the Cells in this WorkBook
	 */
	public BiffRec[] getCells()
	{
		List cellz = new FastAddVector();
		for( int i = 0; i < workSheets.size(); i++ )
		{
			try
			{
				Boundsheet b = this.getWorkSheetByNumber( i );
				BiffRec[] cz = b.getCells();
				for( BiffRec aCz : cz )
				{
					cellz.add( aCz );
				}
			}
			catch( Exception e )
			{
				Logger.logErr( "Error retrieving worksheet for getCells: " + e );
			}
		}
		BiffRec[] cellzr = new BiffRec[cellz.size()];
		cellz.toArray( cellzr );
		return cellzr;
	}

	/**
	 * get the cell by the following String Pattern
	 * <p/>
	 * BiffRec c = getCell("SheetName!C17");
	 */
	public BiffRec getCell( String cellname ) throws CellNotFoundException, WorkSheetNotFoundException
	{
		int semi = cellname.indexOf( "!" );
		String sname = cellname.substring( 0, semi );
		String cname = cellname.substring( semi + 1 );
		return getCell( sname, cname );
	}

	/**
	 * get the cell by the following String Pattern
	 * <p/>
	 * BiffRec c = getCell("Sheet1", "C17");
	 */
	public BiffRec getCell( String sheetname, String cellname ) throws CellNotFoundException, WorkSheetNotFoundException
	{
		cellname = cellname.toUpperCase();
		try
		{
			Boundsheet bs = this.getWorkSheetByName( sheetname );
			BiffRec ret = bs.getCell( cellname );
			if( ret == null )
			{
				throw new CellNotFoundException( sheetname + ":" + cellname );
			}
			return ret;
		}
		catch( WorkSheetNotFoundException a )
		{
			throw new WorkSheetNotFoundException( sheetname + " not found" );
		}
		catch( NullPointerException e )
		{
			throw new CellNotFoundException( sheetname + ":" + cellname );
		}
	}

	private Boundsheet lastbound = null;

	/**
	 * add a Boundsheet to the WorkBook
	 */
	private void addWorkSheet( Long lbplypos, Boundsheet sheet )
	{
		if( sheet == null )
		{
			Logger.logWarn( "WorkBook.addWorkSheet() attempting to add null sheet." );
			return;
		}
		this.lastbound = sheet;
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( "Workbook Adding Sheet: " + sheet.getSheetName() + ":" + String.valueOf( lbplypos ) );
		}
		workSheets.put( lbplypos, sheet );
		boundsheets.add( sheet );
	}

	/**
	 * @param n
	 * @return
	 */
	public boolean removeName( Name n )
	{
		if( this.names.contains( n ) )
		{
			this.names.remove( n );
			if( n.getItab() != 0 )
			{
				try
				{
					this.getWorkSheetByNumber( n.getItab() - 1 ).removeLocalName( n );
				}
				catch( WorkSheetNotFoundException e )
				{
				}
			}
			else
			{
				this.bookNameRecs.remove( n.toString().toUpperCase() );    // case-insensitive
			}
		}
		ArrayList al = n.getIlblListeners();
		this.orphanedPtgNames.addAll( al );

		this.updateNameIlbls();
		return this.getStreamer().removeRecord( n );
	}

	/**
	 * Add a sheet-scoped name record to the boundsheet
	 * <p/>
	 * Note this is not that primary repository for names, it just contains the name records
	 * that are bound to this book, adding them here will not add them to the workbook;
	 *
	 * @param bookNameRecs
	 */
	public void addLocalName( Name name )
	{
		bookNameRecs.put( name.getNameA(), name );
	}

	/**
	 * Remove a sheet-scoped name record from the boundsheet.
	 * <p/>
	 * Note this is not that primary repository for names, it just contains the name records
	 * that are bound to this book, removing them here will not remove them completely from the workbook.
	 * <p/>
	 * In order to do that you will need to call book.removeName
	 *
	 * @param bookNameRecs
	 */
	public void removeLocalName( Name name )
	{
		bookNameRecs.remove( name.getNameA() );
	}

	/**
	 * After any changes in the name records
	 * this method needs to be called in order to
	 * update ilbl records
	 */
	public void updateNameIlbls()
	{
		for( int i = 0; i < names.size(); i++ )
		{
			Name n = (Name) names.get( i );
			n.updateIlblListeners();
		}
	}

	/**
	 * remove a Boundsheet from the WorkBook
	 */
	public void removeWorkSheet( Boundsheet sheet )
	{

		int sheetNum = sheet.getSheetNum();
		// remove the sheet
		// automatically deletes Named ranges scoped to the sheet
		Name[] namesOnSheet = sheet.getAllNames();
		for( Name aNamesOnSheet : namesOnSheet )
		{
			this.removeName( aNamesOnSheet );
		}

		//Remove Externsheet ref before removing sheet
		// update any Externsheet references...
		try
		{
			Externsheet ext = this.getExternSheet();
			if( ext != null )
			{
				ext.removeSheet( sheet.getSheetNum() );
			}
		}
		catch( WorkSheetNotFoundException e )
		{
			Logger.logInfo( "could not update Externsheet reference from " + sheet.toString() + " : " + e.toString() );
		}

		sheet.removeAllRecords();
		streamer.removeRecord( sheet );
		workSheets.remove( new Long( sheet.getLbPlyPos() ) );
		boundsheets.remove( sheet );
		// we need to reset the lastbound for adding new worksheets.  Currently assume it is
		// the last one in the vector.
		if( boundsheets.size() > 0 )
		{
			lastbound = (Boundsheet) (boundsheets.get( boundsheets.size() - 1 ));
			lastBOF = lastbound.getMyBof();
		}

		// decrement the tab ids...
		this.tabs.removeRecord();
		this.updateScopedNamedRanges();

		// update wb chart cache - remove charts referenced by deleted sheet
		for( int i = getChartVect().size() - 1; i >= 0; i-- )
		{
			if( ((Chart) this.getChartVect().get( i )).getSheet() == sheet )
			{
				this.getChartVect().remove( i );
			}
		}

		if( this.getNumWorkSheets() == 0 )
		{
			return; // empty book
		}
		try
		{ // set the next sheet selected...
			while( sheetNum <= this.getNumWorkSheets() )
			{
				Boundsheet s2 = this.getWorkSheetByNumber( sheetNum++ );
				s2.setSelected( true );
				if( !s2.getHidden() )
				{
					break;
				}

			}
		}
		catch( WorkSheetNotFoundException e )
		{
			try
			{
				Boundsheet s2 = this.getWorkSheetByNumber( 0 );
				s2.setSelected( true );
			}
			catch( Exception ee )
			{
				throw new WorkBookException( "Invalid WorkBook.  WorkBook must contain at least one Sheet.",
				                             WorkBookException.RUNTIME_ERROR );
			}
			;
		}
	}

	/**
	 * Updates all the name records in the workbook that are bound to a
	 * worksheet scope (as opposed to a workbook scope).  Name records use
	 * their own non-externsheet based sheet references, so need to be modified
	 * whenever a sheet delete (or non-last sheet insert) operation occurs
	 */
	private void updateScopedNamedRanges()
	{
		for( int i = 0; i < boundsheets.size(); i++ )
		{
			((Boundsheet) boundsheets.get( i )).updateLocalNameReferences();
		}
	}

	/**
	 * returns the Boundsheet identified by its
	 * offset to the BOF record indicating the
	 * start of the Boundsheet data stream.
	 * <p/>
	 * used internally to access the Sheets to
	 * ensure that the lbplypos is correct -- essential
	 * to proper operation of XLS file.
	 *
	 * @param Long lbplypos of Boundsheet
	 */
	private Boundsheet getWorkSheet( Long lbplypos )
	{
		return (Boundsheet) workSheets.get( lbplypos );
	}

	/**
	 * returns the Boundsheet with the specific name
	 *
	 * @param String name of Boundsheet
	 */
	public Boundsheet getWorkSheetByName( String bstr ) throws WorkSheetNotFoundException
	{
		try
		{
			if( bstr.startsWith( "'" ) || bstr.startsWith( "\"" ) )
			{
				bstr = bstr.substring( 1, bstr.trim().length() - 1 );
			}
			Iterator bs = boundsheets.iterator();
			while( bs.hasNext() )
			{
				Boundsheet bsi = (Boundsheet) bs.next();
				String bsin = bsi.getSheetName();
				// TODO: check if we can have dupe names different case
				if( bsin.equalsIgnoreCase( bstr ) )
				{
					return bsi;
				}
			}
		}
		catch( Exception ex )
		{
			Logger.logWarn( "WorkBook.getWorkSheetByName failed: " + ex.toString() );
		}
		throw new WorkSheetNotFoundException( "Worksheet " + bstr + " not found in " + this.toString() );
	}

	/**
	 * returns the Boundsheet with the specific Hashname
	 *
	 * @param String hashname of Boundsheet
	 */
	public Boundsheet getWorkSheetByHash( String s ) throws WorkSheetNotFoundException
	{
		Boundsheet[] bs = this.getWorkSheets();
		for( Boundsheet b : bs )
		{
			if( b.getSheetHash().equalsIgnoreCase( s ) )
			{
				return b;
			}
		}
		return null;
	}

	/**
	 * returns the Boundsheet at the specific index
	 *
	 * @param int index of Boundsheet
	 */
	public Boundsheet getWorkSheetByNumber( int i ) throws WorkSheetNotFoundException
	{
		Boundsheet bs = null;
		try
		{
			bs = (Boundsheet) boundsheets.get( i );
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
		}
		;
		if( bs == null )
		{    // External Sheet Ref NOT FOUND
			throw new WorkSheetNotFoundException( String.valueOf( i ) + " not found" );
		}
		return bs;
	}

	/**
	 * set the last processed Index record
	 */
	private void setLastINDEX( Index id )
	{
		lastidx = id;
	}

	Index getLastINDEX()
	{
		return lastidx;
	}

	void setSharedStringTable( Sst s )
	{
		stringTable = s;
	}

	public Sst getSharedStringTable()
	{
		return stringTable;
	}

	/**
	 * returns the Vector of Boundsheets
	 */
	public AbstractList getSheetVect()
	{
		return this.boundsheets;
	}

	/**
	 * returns the boundsheets for this book as an array
	 */
	Boundsheet[] getWorkSheets()
	{
		Boundsheet[] ret = new Boundsheet[boundsheets.size()];
		return (Boundsheet[]) boundsheets.toArray( ret );
	}

	/**
	 * set the last BOF read in the stream
	 */
	void setLastBOF( Bof b )
	{
		lastBOF = b;
	}
	/** return the last BOF read in the stream */
	// Bof lastBOF{return lastBOF;}

	/**
	 * get a handle to the first BOF to perform offset functions which don't know where the
	 * start of the file is due to the compound file format.
	 * <p/>
	 * Referred to in Boundsheet as the 'lbPlyPos', this
	 * is the position of the BOF for the Boundsheet relative
	 * to the *first* BOF in the file (the firstBOF of the WorkBook)
	 *
	 * @see Boundsheet
	 */
	@Override
	public void setFirstBof( Bof b )
	{
		firstBOF = b;
	}

	Bof getFirstBof()
	{
		return firstBOF;
	}

	public String toString()
	{
		return this.getFileName();
	}

	@Override
	public String getFileName()
	{
		if( this.factory != null ) // 2003-vers
		{
			return this.factory.getFileName();
		}
		return "New Spreadsheet";
	}

	/**
	 * Returns whether the sheet selection tabs should be shown.
	 */
	public boolean showSheetTabs()
	{
		return win1.showSheetTabs();
	}

	/**
	 * Sets whether the sheet selection tabs should be shown.
	 */
	public void setShowSheetTabs( boolean show )
	{
		win1.setShowSheetTabs( show );
	}

	/**
	 * set the first visible tab
	 */
	public void setFirstVisibleSheet( Boundsheet bs2 )
	{
		win1.setFirstTab( bs2.getSheetNum() );
	}

	// Associate related records

	/**
	 * return the XF record at the specified index
	 */
	public Xf getXf( int i )
	{
		if( xfrecs.size() < (i - 1) )
		{
			return null;
		}
		return (Xf) xfrecs.get( i );
	}

	public int getNumXfs()
	{
		return xfrecs.size();
	}

	/**
	 * InsertXF inserts an XF record into the workbook level stream,
	 * For some reason, the addXf only puts it into an array that is never accessed
	 * on output.  This may have a reason, so I am not overwriting it currently, but
	 * let's check it out?
	 */
	int insertXf( Xf x )
	{
		int insertIdx = this.getXf( this.getNumXfs() - 1 ).getRecordIndex();
		// perform default add rsec actions
		this.getStreamer().addRecordAt( x, insertIdx + 1 );
		this.addRecord( x, false );    // updates xfrecs + formatcache
		x.ixfe = x.tableidx;
		return x.tableidx;
	}

	/**
	 * internally used in preparation for reading an 2007 and above workbook
	 */
	public void removeXfRecs()
	{
		// must keep the 1st xf rec as default 
		for( int i = xfrecs.size() - 1; i > 0; i-- )
		{
			Xf xf = (Xf) xfrecs.get( i );
			this.streamer.removeRecord( xf );
			xfrecs.remove( i );

		}
	}

	/**
	 * TODO: Does this function as desired?   See comment for insertXf() above...
	 * tracks existing xf recs, used when testing whether xfrec exists or not ...
	 * -NR 1/06
	 * ya should - called now from addRecord every time an xf record is added
	 * NOTE: this is the only place addXf is called
	 *
	 * @param xf
	 * @return
	 */
	int addXf( Xf xf )
	{
		xfrecs.add( xf );
		xf.tableidx = xfrecs.size() - 1;    // flag that it's been added to records
		this.updateFormatCache( xf );    // links tostring of xf to xf rec for updating/reuse purposes         		
		return xf.tableidx;
	}

	/**
	 * formatCache:
	 * links tostring of xf to xf rec for updating/reuse purposes
	 *
	 * @param xf
	 * @see FormatHandle.updateXf
	 * @see WorkBook.addXf
	 */
	public void updateFormatCache( Xf xf )
	{
		if( xf.tableidx != -1 )
		{    // if this xf has been already added to the workbook 
			if( formatCache.containsValue( xf ) )
			{    // xf signature has changed/it's been updated
				Iterator ii = formatCache.keySet().iterator();    // remove and update below
				while( ii.hasNext() )
				{
					String key = (String) ii.next();
					Xf x = (Xf) formatCache.get( key );
					if( x.equals( xf ) )
					{
						formatCache.remove( key );
						break;
					}
				}
			}
			String formatStr = xf.toString();

			if( !formatCache.containsKey( formatStr ) )
			{
				formatCache.put( formatStr, xf );
			}
		}
	}

	/**
	 * retrieve the format cache - links string vers. of xf to xf rec
	 * used for resusing xf's
	 *
	 * @return
	 * @see FormatHandle.updateXf
	 */
	public HashMap getFormatCache()
	{
		return formatCache;
	}

	/**
	 * Get a substream by name.
	 */
	@Override
	public ByteStreamer getStreamer()
	{
		return this.streamer;
	}

	public void setSubStream( ByteStreamer s )
	{
		this.streamer = s;
	}

	/**
	 * Get the typename for this object.
	 */
	String getTypeName()
	{
		return "WorkBook";
	}

	/**
	 * Write the contents of the WorkBook bytes to an OutputStream
	 *
	 * @param _out
	 */
	@Override
	public int stream( OutputStream _out )
	{
		return streamer.streamOut( _out );
	}

	/**
	 * get a handle to the factory
	 */
	@Override
	public WorkBookFactory getFactory()
	{
		return this.factory;
	}

	/**
	 * Copies a complete boundsheet within the workbook
	 * <p/>
	 * If a name exists that refers directly to this sheet then duplicate it, otherwise workbook scoped names are
	 * not copied
	 */
	public void copyWorkSheet( String SourceSheetName, String NewSheetName ) throws Exception
	{
		Boundsheet origSheet = null;
		origSheet = this.getWorkSheetByName( SourceSheetName );
		List chts = origSheet.getCharts();    // 20080630 KSC: Added
		for( Object cht : chts )
		{
			Chart cxi = (Chart) cht;
			cxi.populateForTransfer();
		}
		byte[] inbytes = origSheet.getSheetBytes();
		this.addBoundsheet( inbytes, SourceSheetName, NewSheetName, null, false );
		Boundsheet bnd = this.getWorkSheetByName( NewSheetName );
		// handle moving the built-in name records.  These handle such items as print area, header/footer, etc
		Name[] ns = this.getNames();
		for( Name n1 : ns )
		{ // 20100404 KSC: take out +1?
			if( n1.getItab() == (origSheet.getSheetNum() + 1) )
			{
				// it's a built in record, move it to the new sheet
				int sheetnum = bnd.getSheetNum();
				int xref = this.getExternSheet( true ).insertLocation( sheetnum, sheetnum );
				Name n = (Name) n1.clone();
				n.setExternsheetRef( xref );
				n.updateSheetReferences( bnd );
				n.setSheet( bnd );
				n.setItab( (short) (bnd.getSheetNum() + 1) );
				this.insertName( n );
			}
		}
	}

	/**
	 * Inserts a newly created Name record into the correct location in the streamer.
	 *
	 * @param n
	 */
	public void insertName( Name n )
	{
		int namepos = -1;
		Name[] nmx = getNames();
		if( nmx.length > 0 )
		{
			if( nmx[nmx.length - 1].getSheet() != null )
			{
				namepos = nmx[nmx.length - 1].getRecordIndex();
			}
			else
			{
				namepos = getExternSheet( true ).getRecordIndex() + nmx.length;
			}
		}
		else
		{
			namepos = getExternSheet( true ).getRecordIndex();
		}
		namepos++;
		getStreamer().addRecordToBookStreamerAt( n, namepos );
		addRecord( n, false );
	}

	/**
	 * Copies an existing Chart to another WorkSheet
	 *
	 * @param chartname
	 * @param sheetname
	 */
	public void copyChartToSheet( String chartname, String sheetname ) throws ChartNotFoundException, WorkSheetNotFoundException
	{
		Chart ct = this.getChart( chartname );
		Boundsheet sht = this.getWorkSheetByName( sheetname );
		byte[] bt = ct.getChartBytes();
		sht.addChart( bt, chartname, ct.getCoords() );
	}

	/**
	 * Inserts a serialized boundsheet chart into the workboook
	 */
	public Chart addChart( Chart destChart, String NewChartName, Boundsheet boundsht )
	{
		destChart.setWorkBook( this );
		destChart.setSheet( boundsht );
		List recs = destChart.getXLSrecs();
		for( Object rec1 : recs )
		{
			XLSRecord rec = (XLSRecord) rec1;
			rec.setWorkBook( this );
			rec.setSheet( boundsht );
			if( rec.getOpcode() == MSODRAWING )
			{
				addChartUpdateMsodg( (MSODrawing) rec, boundsht );
				continue;
			}
			if( !(rec instanceof Bof) )    // TODO: error/problem with the BOF record!!!
			{
				rec.init();
			}
			if( rec instanceof Dimensions )
			{
				destChart.setDimensions( (Dimensions) rec );
			}
			try
			{
				((GenericChartObject) rec).setParentChart( destChart );
			}
			catch( ClassCastException e )
			{    // Scl, Obj and others are not chart objects
			}

		}
		destChart.setTitle( NewChartName );
		destChart.setId( boundsht.lastObjId + 1 );    // track last obj id per sheet ...
		this.charts.add( destChart );
		boundsht.getCharts().add( destChart );    // should really have two lists???
		return destChart;
	}

	/**
	 * updates Mso (MSODrawingGroup + Msodrawing) records upon add/copy worksheet and add/copy charts
	 * NOTE: this code is mainly garnered via trial and error, works
	 *
	 * @param mso Msodrawing record that is being added or copied
	 * @param sht Boundsheet
	 */
	public void addChartUpdateMsodg( MSODrawing mso, Boundsheet sht )
	{
		if( msodg == null )
		{
			this.setMSODrawingGroup( (MSODrawingGroup) MSODrawingGroup.getPrototype() );
			msodg.initNewMSODrawingGroup();    // generate and add required records for drawing records        	
		}
		msodg.addMsodrawingrec( mso );
		MSODrawing hdr = msodg.getMsoHeaderRec( sht );
		if( (hdr != null) && (hdr != mso) )
		{ // already have a header rec
			if( sht.getCharts().size() > 0 )
			{
				mso.makeNonHeader();
				hdr.setNumShapes( hdr.getNumShapes() + 1 );
			}
		}
		else if( hdr == null )
		{
			mso.setIsHeader();
			hdr = mso;
		}
		this.updateMsodrawingHeaderRec( sht );
		msodg.dirtyflag = true;    // flag to reset SPIDs on write
		msodg.setSpidMax( ++lastSPID );
		msodg.updateRecord();
	}

	/**
	 * JM -
	 * Add the requisite records in the book streamer for the chart. \
	 * Supbook, externsheet & msodrawinggroup
	 * <p/>
	 * I think this is due to the fact that the referenced series are usually stored
	 * in the fashon 'Sheet1!A4:B6' The sheet1 reference requires a supbook, though the
	 * reference is internal.
	 */
	public void addPreChart()
	{
		this.addExternsheet();
		if( msodg == null )
		{
			this.setMSODrawingGroup( (MSODrawingGroup) MSODrawingGroup.getPrototype() );
			msodg.initNewMSODrawingGroup();    // generate and add required records for drawing records        	
		}

	}

	/**
	 * remove an existing chart from the workbook
	 * NOTE: STILL EXPERIMENTAL TESTS OK IN BASIC CIRCUMSTANCES BUT MUST BE TESTED FURTHER
	 */
	public void deleteChart( String chartname, Boundsheet sheet ) throws ChartNotFoundException
	{
		Chart chart = this.getChart( chartname );
		// TODO: Update Dimensions record??
		List recs = chart.getXLSrecs();
		// first rec SHOULD BE MsoDrawing!!!        
		try
		{
			MSODrawing rec = (MSODrawing) recs.get( 0 );
			msodg.removeMsodrawingrec( rec, sheet, true );    // also remove associated Obj record        
		}
		catch( Exception e )
		{
			Logger.logErr( "deleteChart: expected Msodrawing record" );
		}
        /* shouldn't be necessary to remove chart recs as they are separated upon init of workbook and reassebmbled upon write*/
		this.removeChart( chartname );
	}

	/**
	 * Inserts a serialized boundsheet into the workboook.
	 *
	 * @param inbytes          original sheet bytes
	 * @param NewSheetName     new Sheet Name
	 * @param origWorkBookName original WorkBook Name (nec. for resolving possible external references)     *
	 * @throws ClassNotFoundException
	 * @throws IOException
	 * @boolean SSTPopulatedBoundsheet - a boundsheet that has all of it's sst data saved off in LabelSST records.
	 * one would use this when moving a sheet from a workbook that was not an original or created from getNoSheetWorkBook.
	 * Do not use this if the data already exists in the SST, you are just causing bloat!
	 */
	public void addBoundsheet( byte[] inbytes,
	                           String origSheetName,
	                           String NewSheetName,
	                           String origWorkBookName,
	                           boolean SSTPopulatedBoundsheet ) throws IOException, ClassNotFoundException
	{
		Boundsheet destSheet = null;
		ByteArrayInputStream bais = new ByteArrayInputStream( inbytes );
		BufferedInputStream bufstr = new BufferedInputStream( bais );
		ObjectInputStream o = new ObjectInputStream( bufstr );
		destSheet = (Boundsheet) o.readObject();

		if( destSheet != null )
		{
			this.addBoundsheet( destSheet, origSheetName, NewSheetName, origWorkBookName, SSTPopulatedBoundsheet );
		}
	}

	/**
	 * change the tab order of a boundsheet
	 */
	public void changeWorkSheetOrder( Boundsheet bs, int idx )
	{
		// reorder the sheet vector
		if( (idx >= 0) && (idx < boundsheets.size()) )
		{
			boundsheets.remove( bs );
			boundsheets.add( idx, bs );
			for( int x = 0; x < boundsheets.size(); x++ )
			{
				Boundsheet bs1 = (Boundsheet) boundsheets.get( x );
				boolean udpatewin1 = bs1.selected();
				if( udpatewin1 )
				{
					bs1.setSelected( true );
				}
			}
		}

		int insertLoc = Integer.MAX_VALUE;
		//remove the existing boundsheet records in the streamer
		for( int i = 0; i < boundsheets.size(); i++ )
		{
			Boundsheet bound = (Boundsheet) boundsheets.get( i );
			int position = bound.getRecordIndex();
			insertLoc = Math.min( insertLoc, position );
			streamer.removeRecord( (XLSRecord) boundsheets.get( i ) );
		}
		// enter the boundsheet records back in the streamer in correct order
		for( int i = 0; i < boundsheets.size(); i++ )
		{
			Boundsheet bound = (Boundsheet) boundsheets.get( i );
			streamer.addRecordAt( bound, insertLoc + i );
		}
	}

	/**
	 * add a deserialized boundsheet  to this workbook.
	 *
	 * @param bound            new (copied) sheet
	 * @param newSheetName     new sheetname
	 * @param origWorkBookName original WorkBook Name (nec. for resolving possible external references)     *
	 * @boolean SSTPopulatedBoundsheet - the boundsheet has all of it's sst data saved off in LabelSST records.
	 * one would use this when moving a sheet from a workbook that was not an original or created from getNoSheetWorkBook.
	 * Do not use this if the data already exists in the SST, you are just causing bloat!
	 */
	public void addBoundsheet( Boundsheet bound,
	                           String origSheetName,
	                           String newSheetName,
	                           String origWorkBookName,
	                           boolean SSTPopulatedBoundsheet )
	{
		bound.streamer = streamer;
		boolean old_allowdupes = this.isSharedupes();
		this.setDupeStringMode( ALLOWDUPES );

		bound.mc.clear();
		bound.setWorkBook( this );

		//Check if sheetname already exists!        
		try
		{
			while( this.getWorkSheetByName( newSheetName ) != null )
			{
				newSheetName = newSheetName + "Copy";
			}
		}
		catch( WorkSheetNotFoundException we )
		{
			; /* good !!!*/
		}
		bound.setSheetName( newSheetName );

		// get a hold of the lbplypos number that we will need for the new boundsheet
		int recvecOffset = streamer.getRecVecSize() - 1;
		XLSRecord x = null;
		if( lastbound != null )
		{
			try
			{    // lastbound must be reset because other operations could alter
				lastbound = this.getWorkSheetByNumber( this.getNumWorkSheets() - 1 );
				x = (XLSRecord) lastbound.getSheetRecs().get( lastbound.getSheetRecs().size() - 1 );
			}
			catch( Exception e )
			{
				;
			}
		}
		else if( countryRec != null )
		{
			x = (XLSRecord) streamer.getRecordAt( countryRec.getRecordIndex() );
		}
		else
		{
			x = this.lasteof;
		}
		// last record is a junkrec.  We are going to move that down and put in the new BOF here
		// TODO: recvecOffset position when no sheets??????
		if( x.getOpcode() != EOF )
		{
			recvecOffset -= 1;
		}
		int newloc = x.offset;
		// modify the boundsheet rec for its new location/info/name
		int listenerpos = -1, newoffset = -1;
		if( lastbound != null )
		{
			listenerpos = lastbound.getRecordIndex();
			newoffset = lastbound.offset + lastbound.getLength() + 4;
			// offset + reclen + headerlen
		}
		else if( x != null )
		{
			listenerpos = x.getRecordIndex() - 1;    // account for +1 below     
			newoffset = x.offset + x.getLength() + 4;
		}
		else
		{
			listenerpos = recvecOffset - 1;
			newoffset = streamer.getRecordAt( recvecOffset ).getLength() + 4;
		}

		//  put the serialized recs from localrecs into the normal SheetRecs
		bound.setLocalRecs( new FastAddVector() ); // reset localrecs
		List newrecs = bound.getSheetRecs();
		Bof newbof = (Bof) newrecs.get( 0 );

		newloc += newbof.getLength() + 4;
		newbof.setOffset( newloc );
		bound.setBOF( newbof );
		this.addRecord( newbof, false );

		recvecOffset += 1; // move it past that last Eof
		newoffset = newloc + newbof.getLength() + 4;
		lastbound = bound;
		// insert the actual boundsheet record into the recvec
		streamer.addRecordAt( bound, listenerpos + 1 );

		this.addRecord( bound, false );

		// modify the TabID record to reflect new sheet
		tabs.addNewRecord();
		recvecOffset = newbof.getRecordIndex();

		int tout = 0;
		copying = true;

		// Add an externsheet ref for the new sheet         
		if( this.myexternsheet == null )
		{
			this.addExternsheet();
		}
		try
		{
			int sheetref = this.getNumWorkSheets() - 1;
			myexternsheet.insertLocation( sheetref, sheetref );
		}
		catch( Exception e )
		{
			Logger.logWarn( "Adding new sheetRef failed in addBoundsheet()" + e.toString() );
		}

		// update the chart references + add to wb
		List chts = bound.getCharts();
		for( Object cht : chts )
		{
			Chart chart = (Chart) cht;    // obviously algorithm has changed and chart is NOT removed :) [discovered by Shigeo/Infoteria/formatbroken273193.sce] // 20080702 KSC: since it's removed, don't inc index
			chart.updateSheetRefs( bound.getSheetName(), origWorkBookName );
			this.charts.add( chart );
		}

		bound.lastObjId = 1;    //  see if resetting obj id helps in file open errors; if so, must reset Note obj id's as well ... 

		/********** This loop handles Boundsheet records contained in the sheet level streamer, that is, not the valrecs *****/
		int numrecs = newrecs.size();
		for( int z = 1; z < numrecs; z++ )
		{
			XLSRecord xl = (XLSRecord) newrecs.get( z );
			this.addRecord( xl, false );
			if( DEBUGLEVEL > 5 )
			{
				try
				{
					Logger.logInfo( "Copying: " + xl.toString() + ":" + String.valueOf( newoffset ) + ":" + xl.getLength() );
				}
				catch( Exception e )
				{
					;
				}
			}
			if( xl instanceof Codename )
			{
				Codename secretagent = (Codename) xl;  // lol -nr
				secretagent.setName( newSheetName );
			}
			else if( xl instanceof Name )
			{
				// Name records specify data ranges -- update to point to new sheet
				Name n = (Name) xl;
				int refnum = myexternsheet.getcXTI();
				n.setExternsheetRef( refnum );
			}
			else if( xl instanceof Cf )
			{    // must check Conditional Format formula refs and handle any external references 
				try
				{
					updateFormulaPtgRefs( ((Cf) xl).getFormula1(), origSheetName, newSheetName, origWorkBookName );
					// NOTE: FORMULA2 can be null -- TODO: should check here 
					updateFormulaPtgRefs( ((Cf) xl).getFormula2(), origSheetName, newSheetName, origWorkBookName );
				}
				catch( Exception e )
				{
				}
			}
			else if( xl.getOpcode() == OBJ )
			{
				((Obj) xl).setObjId( bound.lastObjId++ );
			}
			else if( (xl.getOpcode() == MSODRAWING) || ((xl.getOpcode() == CONTINUE) && (((Continue) xl).maskedMso != null)) )
			{    // 20100510 KSC: handle masked mso's            	
				MSODrawing mso;
				if( xl.getOpcode() == MSODRAWING )
				{
					mso = (MSODrawing) xl;
				}
				else
				{
					mso = ((Continue) xl).maskedMso;
				}
				if( msodg == null )
				{
					this.setMSODrawingGroup( (MSODrawingGroup) MSODrawingGroup.getPrototype() );
					msodg.initNewMSODrawingGroup();    // generate and add required records for drawing records
					msodg.addMsodrawingrec( mso );    // only add when msodg is null b/c otherwise it's added via the addRecord statement above
				}
				if( mso.getImageIndex() > 0 )
				{ //  add image bytes as well, if any        			
					ImageHandle im = bound.getImageByMsoIndex( mso.getImageIndex() );
					int idx = msodg.addImage( im.getImageBytes(), im.getImageType(), false );
					bound.imageMap.put( im, idx );    // 20100518 KSC: makes more sense? im.getImageIndex()));	// add new image to map and link to actual imageIndex - moved from above
					if( idx != mso.getImageIndex() )
					{
						mso.updateImageIndex( idx );
					}
				}
				mso.setSPID( this.lastSPID );
				msodg.setSpidMax( ++this.lastSPID );
				// resets drawing id's - necessarily correct?				msodg.dirtyflag= true;	// flag to reset SPIDs on write
			}
			xl.setOffset( newoffset );
			tout += xl.getLength();
			newoffset += xl.getLength();
		}
		if( msodg != null )
		{// Moved from above so don't udpate at every mso addition
			// necessary?  all mso sub-records on the sheet should have stayed the same ...this.updateMsodrawingHeaderRec(bound);
			msodg.updateRecord();
		}
		/*************** END handling of boundsheet streamer records *************************/

		/*************** HANDLE Formats + PtgRefs in Cell Records ****************************/
		updateTransferedCellReferences( bound, origSheetName, origWorkBookName );

		// associate the records in the sheet
		this.setSharedupes( old_allowdupes );

		if( SSTPopulatedBoundsheet )
		{
			// bring over the sst
			Sst sst = this.getSharedStringTable();
			BiffRec[] b = bound.getCells();
			for( BiffRec aB : b )
			{
				aB.setWorkBook( this );
				if( aB.getOpcode() == XLSConstants.LABELSST )
				{
					Labelsst s = (Labelsst) aB;
					s.insertUnsharedString( sst );
				}
			}
		}

		if( this.getNumWorkSheets() > 1 )
		{
			bound.setSelected( false );
		}
		else
		{
			bound.setSelected( true );
		}

		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "changesize for  new boundsheet: " + bound.getSheetName() + ": " + tout );
		}
		copying = false;
	}

	/**
	 * traverses all rows and their associated cells in the newly transfered sheet,
	 * ensuring formula/cell references and format references are correctly transfered
	 * into the current workbook
	 *
	 * @param bound source sheet
	 */
	private void updateTransferedCellReferences( Boundsheet bound, String origSheetName, String origWorkBookName )
	{
		HashMap localFonts = (HashMap) this.getFontRecsAsXML();
		List boundFonts = bound.getTransferFonts();        // ALL fonts in the source workbook
		HashMap localXfs = (HashMap) this.getXfrecsAsString();
		List boundXfs = bound.getTransferXfs();
		// Set the workbook on all the cells
		Row[] rows = bound.getRows();
		for( Row row : rows )
		{
			row.setWorkBook( this );
			if( row.getIxfe() != this.getDefaultIxfe() )
			{
				transferFormatRecs( row, localFonts, boundFonts, localXfs, boundXfs );    // 20080709 KSC: handle default ixfe for row
			}
			Iterator rowcells = row.getCells().iterator();
			Mulblank aMul = null;
			short c = 0;
			while( rowcells.hasNext() )
			{
				BiffRec b = (BiffRec) rowcells.next();
				if( b.getOpcode() == MULBLANK )
				{
					if( aMul == b )
					{
						c++;
					}
					else
					{
						aMul = (Mulblank) b;
						c = (short) aMul.getColFirst();
					}
					aMul.setCurrentCell( c );
				}
				b.setWorkBook( this );   // Moved to before updateFormulaPtgRefs [BugTracker 1434] 
				if( b instanceof Formula )
				{ // Examine Ptg Refs to handle external sheet references not contained in this workbook
					updateFormulaPtgRefs( (Formula) b, origSheetName, bound.getSheetName(), origWorkBookName );
					if( ((Formula) b).shared != null )
					{
						((Formula) b).shared.setWorkBook( this );
					}

				}
				// 20080226 KSC: transfer format, fonts and xf here instead of populateWorkbookWithRemoteData()
				transferFormatRecs( b, localFonts, boundFonts, localXfs, boundXfs );
			}
		}
		// 20080226 KSC: handle xf's for columns
		for( Colinfo co : bound.getColinfos() )
		{
			transferFormatRecs( co, localFonts, boundFonts, localXfs, boundXfs );
		}
		List c = bound.getCharts();
		for( Object aC : c )
		{
			Chart cht = (Chart) aC;
			ArrayList fontrefs = cht.getFontxRecs();
			for( Object fontref : fontrefs )
			{
				Fontx fontx = (Fontx) fontref;
				int fid = fontx.getIfnt();
				if( fid > 3 )
				{
					fid = bound.translateFontIndex( fid, localFonts );
					fontx.setIfnt( fid );
				}
			}
		}
	}

	/**
	 * examine all Ptg's referenced by this formula, looking for hanging or missing sheet references
	 * if found, sets sheet reference to the current sheet (TODO: a better way?)
	 *
	 * @param f Formula Rec
	 */
	private void updateFormulaPtgRefs( Formula f, String origSheetName, String newSheetName, String origWorkBookName )
	{
		try
		{
			if( f == null )
			{
				return;    // 20100222 KSC
			}
			f.populateExpression();
			Ptg[] p = f.getCellRangePtgs();
			for( Ptg aP : p )
			{
				if( aP instanceof PtgRef )
				{
					PtgRef ptg = (PtgRef) aP;
					try
					{
						if( !(ptg instanceof PtgArea3d) || ((PtgArea3d) ptg).getFirstSheet().equals( ((PtgArea3d) ptg).getLastSheet() ) )
						{
							String sheetName = ptg.getSheetName();
							if( sheetName.equals( origSheetName ) )
							{
								ptg.setSheetName( newSheetName );
							}
							ptg.addToRefTracker();
/* changed to use above.  don't understand this:
    						if (!sheetName.equals(origSheetName)) {
								this.getWorkSheetByName(ptg.getSheetName());
							ptg.setSheetName(newSheetName);
    						} else
    							ptg.setSheetName(newSheetName);
*/
						}
						else
						{ // uncommon case of two sheet range
							PtgArea3d pref = (PtgArea3d) ptg;
//						this.getWorkSheetByName(pref.getFirstPtg().getSheetName());
// don't understand this			this.getWorkSheetByName(pref.getLastPtg().getSheetName());
							ptg.setLocation( ptg.toString() );    // reset ixti if nec.
						}
					}
					catch( WorkSheetNotFoundException we )
					{
						Logger.logWarn( "External Reference encountered upon updating formula references:  Worksheet Reference Found: " + ptg
								.getSheetName() );
						ptg.setExternalReference( origWorkBookName );
					}
				}
				else if( aP instanceof PtgExp )
				{
					PtgExp ptgexp = (PtgExp) aP;
					try
					{
						Ptg[] pe = ptgexp.getConvertedExpression();    // will fail if ShrFmla hasn't been input yet 
						for( Ptg aPe : pe )
						{
							if( aPe instanceof PtgRef )
							{
								PtgRef ptg = (PtgRef) aPe;
								try
								{
									if( ptg instanceof PtgArea3d )
									{ // PtgRef3d, etc.
										this.getWorkSheetByName( ptg.getSheetName() );
										ptg.setLocation( ptg.toString() );    // reset ixti if nec.
									}
									// otherwise, we're good
								}
								catch( WorkSheetNotFoundException we )
								{
									Logger.logWarn(
											"External References Not Supported:  UpdateFormulaReferences: External Worksheet Reference Found: " + ptg
													.getSheetName() );
									ptg.setExternalReference( origWorkBookName );
								}
							}
						}
					}
					catch( Exception e )
					{
						//if links to "main" ShrFmla, won't be set yet and will give exception - see Shrfmla WorkBook.addRecord 
					}
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "WorkBook.updateFormulaRefs: error parsing expression: " + e );
		}
	}

	/**
	 * given a record in an previously external workbook, ensure that xf and font records
	 * are correctly input into the current workbook and that the pointers are correctly updated
	 *
	 * @param b          BiffRec
	 * @param localFonts HashMap of string version of all fonts, font nums in current workbook
	 * @param boundFonts List of string version of all fonts, font nums in external workbook
	 * @param localXfs   HashMap of string version of all xfs, xf nums in current workbook
	 * @param boundXfs   List of string version of all xfs, xf nums in external workbook
	 */
	private void transferFormatRecs( BiffRec b,
	                                 HashMap<String, Integer> localFonts,
	                                 List boundFonts,
	                                 HashMap<String, Integer> localXfs,
	                                 List boundXfs )
	{
		int oldXfNum = b.getIxfe();
		int localNum = this.transferFormatRecs( oldXfNum, localFonts, boundFonts, localXfs, boundXfs );
		if( localNum != -1 )
		{
			b.setIxfe( localNum );
		}
	}

	/**
	 * given a record in an previously external workbook, ensure that xf and font records
	 * are correctly input into the current workbook and that the pointers are correctly updated
	 *
	 * @param b          BiffRec
	 * @param localFonts HashMap of string version of all fonts, font nums in current workbook
	 * @param boundFonts List of string version of all fonts, font nums in external workbook
	 * @param localXfs   HashMap of string version of all xfs, xf nums in current workbook
	 * @param boundXfs   List of string version of all xfs, xf nums in external workbook
	 */
	private int transferFormatRecs( int oldXfNum,
	                                HashMap<String, Integer> localFonts,
	                                List boundFonts,
	                                HashMap<String, Integer> localXfs,
	                                List boundXfs )
	{
		int localNum = -1;
		if( boundXfs.size() > oldXfNum )
		{// if haven't populatedForTransfer i.e. haven't opted to transfer formats ...
			Xf origxf = (Xf) boundXfs.get( oldXfNum );        // clone xf so modifcations don't affect original
			if( origxf != null )
			{
				/** FONT **/
				// must handle font first in order to create xf below
				// see if referenced xf + fonts are already in workbook; if not, add
				int localfNum;
				// check to see if the font needs to be added 
				int fnum = origxf.getIfnt();
				if( fnum > 3 )
				{
					fnum--;
				}
				Font thisFont = (Font) boundFonts.get( fnum );
				String xmlFont = "<FONT><" + thisFont.getXML() + "/></FONT>";
				Object fontNum = localFonts.get( xmlFont );
				if( fontNum != null )
				{ // then get the fontnum in this book
					localfNum = (Integer) fontNum;
				}
				else
				{ // it's a new font for this workbook, add it in
					localfNum = this.insertFont( thisFont ) + 1;
					localFonts.put( xmlFont, localfNum );
				}

				/** XF **/
				Xf localxf = FormatHandle.cloneXf( origxf, origxf.getFont(), this );    // clone xf so modifcations don't affect original
				// input "local" versions of format and font

				/** FORMAT **/
				Format fmt = origxf.getFormat();    // number format - is null if format is general ...
				if( fmt != null ) // add if necessary
				{
					localxf.setFormatPattern( fmt.getFormat() );    // adds new format pattern if not found
				}
				localxf.setFont( localfNum );

				// now check out to see if xf needs to be added 
				String xmlxf = localxf.toString();
				Object xfNum = localXfs.get( xmlxf );
				if( xfNum == null )
				{ // insert it into the book
					localNum = this.insertXf( localxf );
					localXfs.put( xmlxf, localNum );
				}
				else  // already exists in the destination
				{
					localNum = (Integer) xfNum;
				}

			}
		}
		return localNum;
	}

	public void setStringEncodingMode( int mode )
	{
		this.getSharedStringTable().setStringEncodingMode( mode );
	}

	public void setDupeStringMode( int mode )
	{
		if( mode == ALLOWDUPES )
		{
			this.setSharedupes( false );
		}
		else if( mode == SHAREDUPES )
		{
			this.setSharedupes( true );
		}
	}

	/**
	 * Returns a Chart Handle
	 *
	 * @return ChartHandle a Chart in the WorkBook
	 */
	public Chart getChart( String chartname ) throws ChartNotFoundException
	{
		AbstractList cv = this.getChartVect();
		Chart cht = null;
		// Get by MSODG Drawing Name
		for( int x = 0; x < cv.size(); x++ )
		{
			cht = (Chart) cv.get( x );
			MSODrawing titlemso = cht.getMsodrawobj();
			if( titlemso != null )
			{
				String mson = titlemso.getName();    //shapeName;
				if( mson.equalsIgnoreCase( chartname ) )
				{
					return cht;
				}
			}
		}
		boolean untitled = chartname.equals( "[Untitled]" );
		// Try to get by title
		for( int x = 0; x < cv.size(); x++ )
		{
			cht = (Chart) cv.get( x );
			String cname = cht.getTitle();
			if( cname.equalsIgnoreCase( chartname ) )
			{
				return cht;
			}
			if( untitled && cname.equals( "" ) )
			{
				return cht;
			}
		}
		throw new ChartNotFoundException( chartname );
	}

	/**
	 * removes the desired chart from the list of charts
	 *
	 * @param chartname
	 * @throws ChartNotFoundException
	 */
	public void removeChart( String chartname ) throws ChartNotFoundException
	{
		AbstractList cv = this.getChartVect();
		Chart cht = null;
		for( int x = 0; x < cv.size(); x++ )
		{
			cht = (Chart) cv.get( x );
			if( cht.getTitle().equalsIgnoreCase( chartname ) )
			{
				cv.remove( x );
				return;
			}
		}
		throw new ChartNotFoundException( chartname );
	}

	/**
	 * NOT 100% IMPLEMENTE YET
	 * creates the initial records for a Pivot Cache
	 * <br>A Pivot Cache identifies the data used in a Pivot Table
	 * <br>NOTE: only SHEET cache sources are supported at this time
	 *
	 * @param ref       String reference: either reference or named range
	 * @param sheetName String sheetname
	 * @param cacheid   if > 0, the desired cacheid (useful only in OOXML parsing)
	 * @return int cacheid
	 */
	public int addPivotStream( String ref, String sheetName, int sid )
	{
		//in wb substream, DIRECTLY AFTER STYLE records:
		// STYLE/STYLEEX [TableStyle TableStyleElement] [Palette] [ClrtClient]
		if( sid < 0 )
		{
			sid = 0;    // initial cache id if none already present
		}
		List records = this.getStreamer().records;
		int z = -1;
		for( int i = records.size() - 1; (i > 0) && (z == -1); i-- )
		{
			int opcode = ((BiffRec) records.get( i )).getOpcode();
			if( opcode == SXADDL )
			{        // find last cache id and increment
/*    			while (i > 0 && opcode!=SXSTREAMID)
    				opcode= ((BiffRec) records.get(i--)).getOpcode();
    			if (opcode==SXSTREAMID) {
    				sid= ((SxStreamID) records.get(i+1)).getStreamID() + 1;
    			}*/
				z = i + 1;
			}
			else if( opcode == 4188 )    // ClrtClient
			{
				z = i + 1;
			}
			else if( opcode == PALETTE )
			{
				z = i + 1;
			}
			else if( opcode == 2192 )    // TableStyleElement
			{
				z = i + 1;
			}
			else if( opcode == 2194 )    // StyleEx
			{
				z = i + 1;
			}
			else if( opcode == STYLE )    // Style
			{
				z = i + 1;
			}
		}

		TableStyles tx = (TableStyles) TableStyles.getPrototype();    // see if this is really necessary ...
		this.getStreamer().addRecordAt( tx, z++ );
		SxStreamID sxid = (SxStreamID) SxStreamID.getPrototype();
		this.getStreamer().addRecordAt( sxid, z++ );
		this.ptstream.add( sxid );    // Pivot Cache -                     
		sxid.setStreamID( sid );    // cache id
		this.getStreamer().records.addAll( z, sxid.addInitialRecords( this, ref, sheetName ) );

		return sid;
	}

	/**
	 * adds the Pivot Cache Directory Storage +Stream records necessary to
	 * define the pivot cache (==pivot table data) for pivot table(s)
	 * <br>NOTE: at this time only 1 pivot cache is supported
	 *
	 * @param ref Cell Range which identifies pivot table data range
	 * @param wbh
	 * @param sId Stream or cachid Id -- links back to SxStream set of records
	 */
	public void addPivotCache( String ref, WorkBookHandle wbh, int sId )
	{
		if( ptcache == null )
		{
			ptcache = new PivotCache();
			ptcache.createPivotCache( factory.myLEO.getDirectoryArray(), wbh, ref, sId );
		}
	}

	/**
	 * returns the start of the stream defining the desired pivot cache
	 *
	 * @param cacheid
	 * @return
	 */
	public SxStreamID getPivotStream( int cacheid )
	{
//int z= 0;    	
		for( Object aPtstream : ptstream )
		{
			int sid = ((SxStreamID) aPtstream).getStreamID();
			if( sid == cacheid )
//    		if (z++==cacheid)
			{
				return (SxStreamID) aPtstream;
			}
		}
/*    	List records= this.getStreamer().records;
    	for (int i= 0; i < records.size(); i++) {
    		int opcode= ((BiffRec) records.get(i)).getOpcode();
    		if (opcode==SXSTREAMID) { 
    			int sid= ((SxStreamID) records.get(i)).getStreamID() + 1;
    			if (sid==cacheid)
    				return (SxStreamID) records.get(i);
    		}
    	}*/
		return null;
	}

	/**
	 * @return
	 */
	public AbstractList getHlinklookup()
	{
		return hlinklookup;
	}

	/**
	 * @return
	 */
	public AbstractList getMergecelllookup()
	{
		return mergecelllookup;
	}

	public void addIndirectFormula( Formula f )
	{
		indirectFormulas.add( f );
	}

	/**
	 * Initialize the indirect functions in this workbook by calculating the formulas
	 */
	public void initializeIndirectFormulas()
	{
		Iterator i = indirectFormulas.iterator();    // contains all INDIRECT funcvars + params 
		while( i.hasNext() )
		{
			Formula f = (Formula) i.next();
			f.calculateIndirectFunction();
		}
		indirectFormulas = new ArrayList(); // clear out
	}

	/**
	 * Inserts an externsheet into the recvec, provided one does not yet exist.
	 * also calls add supBook
	 */
	public void addExternsheet()
	{
		if( myexternsheet == null )
		{
			int numsheets = this.getNumWorkSheets();
			Supbook sb = (Supbook) Supbook.getPrototype( numsheets );
			// put it in after the last boundsheet record
			try
			{
				Boundsheet b = this.getWorkSheetByNumber( numsheets - 1 );
				int loc = b.getRecordIndex() + 1;
				if( streamer.getRecordAt( loc ).getOpcode() == COUNTRY )
				{
					loc++;
				}
				streamer.addRecordAt( sb, loc );    // 20080306 KSC: do first 'cause externsheet now references global sb store 
				this.addRecord( sb, false );
				Externsheet ex = (Externsheet) Externsheet.getPrototype( 0, 0, this );
				streamer.addRecordAt( ex, loc + 1 );// 20080306 KSC: must inc loc since now inserting after sb           
				this.addRecord( ex, false );
				myexternsheet = ex;
			}
			catch( WorkSheetNotFoundException e )
			{
				Logger.logWarn( "WorkBook.addExternSheet() locating Sheet for adding Externsheet failed: " + e );
			}

		}
	}

	/**
	 * Sets the ExtenXLS calculation mode for the workbook.
	 *
	 * @param CalcMode
	 * @see WorkBookHandle.setFormulaCalculationMode()
	 */
	public void setCalcMode( int mode )
	{
		CalcMode = mode;
	}

	/**
	 * Sets the calculation mode for the workbook.
	 *
	 * @param CalcMode
	 * @see WorkBookHandle.setFormulaCalculationMode()
	 */
	public int getCalcMode()
	{
		return CalcMode;
	}

	/**
	 * @return Returns the xfrecs.
	 */
	public AbstractList getXfrecs()
	{
		return xfrecs;
	}

	/**
	 * Return the font records
	 */
	public AbstractList getFontRecs()
	{
		return fonts;
	}

	/**
	 * Returns a map of xml strings representing the XF's in this workbook/Integer of lookup.
	 * These are used as a comparitor to determine if additional xf's need to be brought in or
	 * not and to give the new XF number if the xf exists.
	 * <p/>
	 * Changed 20080226 KSC: to use XF toString as XML is limited in format, toString is more complete
	 *
	 * @return map (String XfXml, Integer xfLookup)
	 */
	public Map getXfrecsAsString()
	{
		Map retMap = new HashMap();
		for( int xfNum = 1; xfNum < this.getNumXfs(); xfNum++ )
		{
			Xf x = this.getXf( xfNum );
			String xml = x.toString();
			retMap.put( xml, xfNum );
		}
		return retMap;
	}

	/**
	 * Returns a map of xml strings representing the XF's in this workbook/Integer of lookup.
	 * These are used as a comparitor to determine if additional xf's need to be brought in or
	 * not and to give the new XF number if the xf exists.
	 *
	 * @return map (String XfXml, Integer xfLookup)
	 */
	public Map getFontRecsAsXML()
	{
		Map retMap = new HashMap();
		for( int i = this.fonts.size() - 1; i >= 0; i-- )
		{
			Font fnt = (Font) this.fonts.get( i );
			String xml = "<FONT><" + fnt.getXML() + "/></FONT>";
			retMap.put( xml, fnt.getIdx() );
		}
		return retMap;
	}

	/**
	 * @return Returns the lastbound.
	 */
	public Boundsheet getLastbound()
	{
		return lastbound;
	}

	/**
	 * @param lastbound The lastbound to set.
	 */
	public void setLastbound( Boundsheet lastbound )
	{
		this.lastbound = lastbound;
	}

	/**
	 * add formula to book and init the ptgs
	 *
	 * @param rec
	 */
	public void addFormula( Formula rec )
	{
		this.formulas.add( rec );
		String shn = rec.getSheet().getSheetName() + "!" + rec.getCellAddress();
		this.formulashash.put( shn, rec );
	}

	public boolean isSharedupes()
	{
		return sharedupes;
	}

	public void setSharedupes( boolean sharedupes )
	{
		this.sharedupes = sharedupes;
	}

	/**
	 * Returns whether this book uses the 1904 date format.
	 *
	 * @deprecated Use {@link #getDateFormat()} instead.
	 */
	public boolean is1904()
	{
		return this.dateFormat == DateConverter.DateFormat.LEGACY_1904;
	}

	/**
	 * Gets the date format used by this book.
	 */
	public DateConverter.DateFormat getDateFormat()
	{
		return this.dateFormat;
	}

	public ReferenceTracker getRefTracker()
	{
		return refTracker;
	}

	// OOXML Additions
	private boolean isExcel2007 = false;

	/**
	 * set truth of "Is Excel 2007"
	 * true increases maximums of base storage, etc.
	 *
	 * @param b
	 */
	public void setIsExcel2007( boolean b )
	{
		isExcel2007 = b;
	}

	/**
	 * returns truth of "Excel 2007" format
	 */
	public boolean getIsExcel2007()
	{
		return isExcel2007;
	}

	/**
	 * returns the workbook codename used by vba macros OOXML-specific
	 */
	public String getCodename()
	{
		return ooxmlcodename;
	}

	/**
	 * returns true if the default language selected in Excel is one of
	 * the DBCS (Double-Byte Code Set) languages
	 * <br>
	 * The languages that support DBCS include
	 * <br>
	 * Japanese, Chinese (Simplified), Chinese (Traditional), and Korean
	 * <pre>
	 * Language        Country code    Countries/regions
	 * -------------------------------------------------------------
	 *
	 * Arabic                966       (Saudi Arabia)
	 * Czech                 42        (Czech Republic)
	 * Danish                45        (Denmark)
	 * Dutch                 31        (The Netherlands)
	 * English               1         (The United States of America)
	 * Farsi                 98        (Iran)
	 * Finnish               358       (Finland)
	 * French                33        (France)
	 * German                49        (Germany)
	 * Greek                 30        (Greece)
	 * Hebrew                972       (Israel)
	 * Hungarian             36        (Hungary)
	 * Indian                91        (India)
	 * Italian               39        (Italy)
	 * Japanese              81        (Japan)
	 * Korean                82        (Korea)
	 * Norwegian             47        (Norway)
	 * Polish                48        (Poland)
	 * Portuguese (Brazil)   55        (Brazil)
	 * Portuguese            351       (Portugal)
	 * Russian               7         (Russian Federation)
	 * Simplified Chinese    86        (People's Republic of China)
	 * Spanish               34        (Spain)
	 * Swedish               46        (Sweden)
	 * Thai                  66        (Thailand)
	 * Traditional Chinese   886       (Taiwan)
	 * Turkish               90        (Turkey)
	 * Urdu                  92        (Pakistan)
	 * Vietnamese            84        (Vietnam)
	 * </pre>
	 *
	 * @return boolean
	 */
	public boolean defaultLanguageIsDBCS()
	{
		return ((this.defaultLanguage == 81) ||
				(this.defaultLanguage == 886) ||
				(this.defaultLanguage == 86) ||
				(this.defaultLanguage == 82));
		// PROBLEM WITH THIS:  POSSIBLE TO BE SET AS DBCS DEFAULT LANGUAGE
		// BUT HAVE NON-DBCS TEXT or VISA VERSA
    	
    	/*
    	 * In a double-byte character set, some characters require two bytes, 
    	 * while some require only one byte. 
    	 * The language driver can distinguish between these two types of characters by designating 
    	 * some characters as "lead bytes." 
    	 * A lead byte will be followed by another byte (a "tail byte") to create a 
    	 * Double-Byte Character (DBC). 
    	 * The set of lead bytes is different for each language. 
    	 * 
    	 * Lead bytes are always guaranteed to be extended characters; no 7-bit ASCII characters 
    	 * can be lead bytes. 
    	 * The tail byte may be any byte except a NULL byte. 
    	 * The end of a string is always defined as the first NULL byte in the string. 
    	 * Lead bytes are legal tail bytes; the only way to tell if a byte is acting as a 
    	 * lead byte is from the context.
    	 */
	}

	/**
	 * sets the workbook codename used by vba macros OOXML-specific
	 *
	 * @param s
	 */
	public void setCodename( String s )
	{
		ooxmlcodename = s;
	}    // TODO: input into Codename record

	/**
	 * sets the first sheet (int) in the workbook OOXML-specific
	 * <p/>
	 * naive implementaiton does not account for hidden sheets
	 * <p/>
	 * Mar 15, 2010
	 *
	 * @param f
	 */
	public void setFirstSheet( int f )
	{
		firstSheet = f;
	}

	/**
	 * returns the first non-hidden
	 * <p/>
	 * sheet (int) in the workbook OOXML-specific
	 * <p/>
	 * Mar 15, 2010
	 *
	 * @return
	 */
	public int getFirstSheet()
	{
		try
		{
			if( !getWorkSheetByNumber( firstSheet ).getHidden() )
			{
				return this.firstSheet;
			}
		}
		catch( Exception x )
		{
			;
		}

		// first sheet is hidden -- fix
		for( int t = 0; t < this.getNumWorkSheets(); t++ )
		{
			try
			{
				if( !getWorkSheetByNumber( t ).getHidden() )
				{
					firstSheet = t;
					return firstSheet;
				}
			}
			catch( Exception x )
			{
				;
			}
		}
		// all else failed
		return firstSheet;
	}

	/**
	 * sets the list of dxf's (incremental style info) (int) OOXML-specific
	 */
	public void setDxfs( ArrayList dxfs )
	{
		this.dxfs = dxfs;
	}

	/**
	 * returns the list of dxf's (incremental style info) (int) OOXML-specific
	 */
	public ArrayList getDxfs()
	{
		return dxfs;
	}

	/**
	 * Returns all strings that are in the SharedStringTable for this workbook.  The SST contains
	 * all standard string records in cells, but may not include such things as strings that are contained
	 * within formulas.   This is useful for such things as full text indexing of workbooks
	 *
	 * @return Strings in the workbook.
	 */
	public String[] getAllStrings()
	{
		ArrayList al = this.getSharedStringTable().getAllStrings();
		String[] s = new String[al.size()];
		s = (String[]) al.toArray( s );
		return s;
	}

	/**
	 * set the pivot cache pointer
	 *
	 * @param pc initialized pivot cache
	 */
	public void setPivotCache( PivotCache pc )
	{
		ptcache = pc;
	}

	/**
	 * return the pivot cache, if any
	 *
	 * @return
	 */
	public PivotCache getPivotCache()
	{
		return ptcache;
	}

	/**
	 * clear out ALL sheet object refs
	 */
	public void closeSheets()
	{
		for( int i = boundsheets.size() - 1; i > 0; i-- )
		{
			Boundsheet b = (Boundsheet) boundsheets.get( i );
			if( b.streamer != null )
			{    // do separately because may call boundsheet close
				b.streamer.close();
				b.streamer = null;
			}
			b.close();
			if( tabs != null )
			{
				tabs.removeRecord();
			}
		}
		Object[] recs = this.getStreamer().getBiffRecords();
		boolean resetVars = (recs[0] != null);
		boundsheets.clear();
		for( int i = 0; i < formulas.size(); i++ )
		{
			Formula f = (Formula) formulas.get( i );
			f.close();
		}
		this.formulas.clear();
		if( this.refTracker != null )
		{
			this.refTracker.clearCaches();
		}
		this.formulashash.clear();
		// TODO: handle
		this.indirectFormulas.clear();
		this.charts.clear();
		this.chartTemp.clear();
		if( firstBOF != null )
		{
			firstBOF.close();
			this.firstBOF = null;
		}
		if( lasteof != null )
		{
			lasteof.close();
			lasteof = null;
		}
		if( resetVars )
		{    // just clearing out sheets instead of closing workbook
			// reset lasteof for possible new insertion of sheets (if not removing workbook) ...
			int i = recs.length - 1;
			while( (i > 0) && (lasteof == null) )
			{
				if( recs[i] != null )
				{
					if( ((BiffRec) recs[i]).getOpcode() == EOF )
					{
						lasteof = (Eof) recs[i];
					}
				}
				i--;
			}
		}
		if( lastBOF != null )
		{
			lastBOF.close();
			lastBOF = null;
		}
		if( resetVars )
		{    // just clearing out sheets instead of closing workbook
			if( ((BiffRec) recs[0]).getOpcode() == BOF )
			{// should!!
				lastBOF = (Bof) recs[0];
				firstBOF = (Bof) recs[0];
			}
		}
		if( lastbound != null )
		{
			lastbound.close();
			this.lastbound = null;
		}
		if( lastFormula != null )
		{
			lastFormula.close();
			this.lastFormula = null;
		}
		this.workSheets.clear();
	}

	protected void closeRecords()
	{
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		closeSheets();

		if( this.isExcel2007 )
		{
			try
			{
				String externalDir = OOXMLAdapter.getTempDir( this.getFactory().getFileName() );
				OOXMLAdapter.deleteDir( new File( externalDir ) );
			}
			catch( Exception e )
			{
				;
			}
		}
		contHandler.close();
		contHandler = new ContinueHandler( this );

		// clear out array list references
		Iterator ii = bookNameRecs.keySet().iterator();
		while( ii.hasNext() )
		{
			Name n = (Name) bookNameRecs.get( ii.next() );
			n.close();
		}
		this.bookNameRecs.clear();
		for( int i = 0; i < xfrecs.size(); i++ )
		{
			Xf x = (Xf) xfrecs.get( i );
			x.close();
		}
		this.xfrecs.clear();
		this.formatCache.clear();
		this.formatlookup.clear();
		this.formats.clear();
		this.fonts.clear();
		if( this.dxfs != null )
		{
			this.dxfs.clear();
		}
		for( int i = 0; i < msodgMerge.size(); i++ )
		{
			MSODrawingGroup m = (MSODrawingGroup) msodgMerge.get( i );
			m.close();
		}
		this.msodgMerge.clear();
		if( this.msodg != null )
		{
			msodg.close();
			msodg = null;
		}
		// TODO: handle
		this.mergecelllookup.clear();
		this.hlinklookup.clear();
		this.externalnames.clear();
		for( int i = 0; i < names.size(); i++ )
		{
			Name n = (Name) names.get( i );
			n.close();
		}
		this.names.clear();
		for( int i = 0; i < indexes.size(); i++ )
		{
			Index ind = (Index) indexes.get( i );
			ind.close();
		}
		this.indexes.clear();
		if( lastidx != null )
		{
			lastidx.close();
			lastidx = null;
		}
		if( this.stringTable != null )
		{
			this.stringTable.close();
			this.stringTable = null;
		}

		if( countryRec != null )
		{
			this.countryRec.close();
			this.countryRec = null;
		}
		if( win1 != null )
		{
			win1.close();
			win1 = null;
		}

		if( this.drh != null )
		{
			this.drh.close();
			drh = null;
		}

		// integration point for subclasses
		this.closeRecords();

		this.contHandler = new ContinueHandler( this );
		if( this.currai != null )
		{
			this.currai.close();
			this.currai = null;
		}
		if( calcmoderec != null )
		{
			calcmoderec.close();
			calcmoderec = null;
		}
		this.currchart = null;
		this.currdrw = null;
		if( this.protector != null )
		{
			this.protector.close();
			this.protector = null;
		}
		if( this.myexternsheet != null )
		{
			this.myexternsheet.close();
			this.myexternsheet = null;
		}
		if( this.refTracker != null )
		{
			this.refTracker.close();
			this.refTracker = null;
		}
		if( tabs != null )
		{
			tabs.close();
			tabs = null;
		}
		// TODO: deal
		this.factory = null;
		streamer = createByteStreamer();
		if( xl2k != null )
		{
			((XLSRecord) xl2k).close();
		}
		xl2k = null;

	}

	public Color[] getColorTable()
	{
		return colorTable;
	}

	public void setColorTable( Color[] clrtable )
	{
		colorTable = clrtable;
	}

	public int getDefaultIxfe()
	{
		return defaultIxfe;
	}

	public void setDefaultIxfe( int defaultIxfe )
	{
		this.defaultIxfe = defaultIxfe;
	}

}
