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

import org.openxls.formats.XLS.charts.Ai;
import org.openxls.formats.XLS.charts.AlRuns;
import org.openxls.formats.XLS.charts.Area;
import org.openxls.formats.XLS.charts.AreaFormat;
import org.openxls.formats.XLS.charts.AttachedLabel;
import org.openxls.formats.XLS.charts.Axcent;
import org.openxls.formats.XLS.charts.Axesused;
import org.openxls.formats.XLS.charts.Axis;
import org.openxls.formats.XLS.charts.AxisLineFormat;
import org.openxls.formats.XLS.charts.AxisParent;
import org.openxls.formats.XLS.charts.Bar;
import org.openxls.formats.XLS.charts.Begin;
import org.openxls.formats.XLS.charts.Boppop;
import org.openxls.formats.XLS.charts.BoppopCustom;
import org.openxls.formats.XLS.charts.CatLab;
import org.openxls.formats.XLS.charts.CatserRange;
import org.openxls.formats.XLS.charts.Chart;
import org.openxls.formats.XLS.charts.Chart3DBarShape;
import org.openxls.formats.XLS.charts.ChartFormat;
import org.openxls.formats.XLS.charts.ChartFormatLink;
import org.openxls.formats.XLS.charts.ChartFrtInfo;
import org.openxls.formats.XLS.charts.ChartLine;
import org.openxls.formats.XLS.charts.CrtLayout12;
import org.openxls.formats.XLS.charts.CrtLayout12A;
import org.openxls.formats.XLS.charts.Dat;
import org.openxls.formats.XLS.charts.DataFormat;
import org.openxls.formats.XLS.charts.DataLabExt;
import org.openxls.formats.XLS.charts.DataLabExtContents;
import org.openxls.formats.XLS.charts.DefaultText;
import org.openxls.formats.XLS.charts.Dropbar;
import org.openxls.formats.XLS.charts.End;
import org.openxls.formats.XLS.charts.EndBlock;
import org.openxls.formats.XLS.charts.EndObject;
import org.openxls.formats.XLS.charts.Fbi;
import org.openxls.formats.XLS.charts.Fontx;
import org.openxls.formats.XLS.charts.Frame;
import org.openxls.formats.XLS.charts.FrtFontList;
import org.openxls.formats.XLS.charts.FrtWrapper;
import org.openxls.formats.XLS.charts.GelFrame;
import org.openxls.formats.XLS.charts.Ifmt;
import org.openxls.formats.XLS.charts.Legend;
import org.openxls.formats.XLS.charts.Legendxn;
import org.openxls.formats.XLS.charts.Line;
import org.openxls.formats.XLS.charts.LineFormat;
import org.openxls.formats.XLS.charts.MarkerFormat;
import org.openxls.formats.XLS.charts.ObjectLink;
import org.openxls.formats.XLS.charts.Picf;
import org.openxls.formats.XLS.charts.Pie;
import org.openxls.formats.XLS.charts.PieFormat;
import org.openxls.formats.XLS.charts.PivotChartBits;
import org.openxls.formats.XLS.charts.PivotChartLink;
import org.openxls.formats.XLS.charts.PlotArea;
import org.openxls.formats.XLS.charts.PlotGrowth;
import org.openxls.formats.XLS.charts.Pos;
import org.openxls.formats.XLS.charts.Radar;
import org.openxls.formats.XLS.charts.RadarArea;
import org.openxls.formats.XLS.charts.SbaseRef;
import org.openxls.formats.XLS.charts.Scatter;
import org.openxls.formats.XLS.charts.SerParent;
import org.openxls.formats.XLS.charts.SerToCrt;
import org.openxls.formats.XLS.charts.SerauxErrBar;
import org.openxls.formats.XLS.charts.SerauxTrend;
import org.openxls.formats.XLS.charts.Serfmt;
import org.openxls.formats.XLS.charts.Series;
import org.openxls.formats.XLS.charts.SeriesList;
import org.openxls.formats.XLS.charts.SeriesText;
import org.openxls.formats.XLS.charts.ShtProps;
import org.openxls.formats.XLS.charts.SiIndex;
import org.openxls.formats.XLS.charts.StartBlock;
import org.openxls.formats.XLS.charts.StartObject;
import org.openxls.formats.XLS.charts.Surface;
import org.openxls.formats.XLS.charts.SxViewLink;
import org.openxls.formats.XLS.charts.TextDisp;
import org.openxls.formats.XLS.charts.ThreeD;
import org.openxls.formats.XLS.charts.Tick;
import org.openxls.formats.XLS.charts.Units;
import org.openxls.formats.XLS.charts.ValueRange;
import org.openxls.formats.XLS.charts.YMult;
import org.openxls.formats.XLS.formulas.Ptg;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Factory to create XLSRecords and Ptgs.
 */
public class XLSRecordFactory implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( XLSRecordFactory.class );
	/**
	 * This class is static only, prohibit construction.
	 */
	private XLSRecordFactory()
	{
		throw new UnsupportedOperationException( "XLSRecordFactory is purely static" );
	}

	// 20060504 KSC: add separate array for unary prefix operators: 
	//				 maps regular operator to unary pseudo version located in ptgLookup below 
	public static Object[][] ptgPrefixOperators = {
			{ "+", "u+" }, { "-", "u-" },
			// should $ be in here???
	};

	/**
	 * Maps BIFF8 record opcodes to classes.
	 */
	// Removed static object instance creations as very difficult (impossible) to dereference so as can release memeory ...
// TRY :	private static final Map records;

	// subset of ptgLookup below used for pattern matching in formula strings
	public static String[][] ptgOps = {
			{ "#VALUE!", "PtgErr" },
			{ "#NULL!", "PtgErr" },
			{ "#DIV/0!", "PtgErr" },
			{ "#VALUE!", "PtgErr" },
			{ "#REF!", "PtgErr" },
			{ "#NUM!", "PtgErr" },
			{ "#N/A", "PtgErr" },
			// operators
			{ "(", "PtgParen" },
			{ "*", "PtgMlt" },
			{ "/", "PtgDiv" },
			{ "^", "PtgPower" },
			{ "&", "PtgConcat" },
			{ "<>", "PtgNE" },
			{ "<=", "PtgLE" },
			{ "<", "PtgLT" },
			{ ">=", "PtgGE" },
			{ ">", "PtgGT" },
			{ "=", "PtgEQ" },
			{ "!=", "PtgNE" },
			{ "+", "PtgAdd" },
			// moved to AFTER other operators
			{ "-", "PtgSub" },
	};
	// this is how you init a 2D array of Objects
	public static String[][] ptgLookup = {
			// 20070215 KSC: added constant strings to avoid PtgRef3d matches
			{ "#VALUE!", "PtgErr" },
			{ "#NULL!", "PtgErr" },
			{ "#DIV/0!", "PtgErr" },
			{ "#VALUE!", "PtgErr" },
			{ "#REF!", "PtgErr" },
			{ "#NUM!", "PtgErr" },
			{ "#N/A", "PtgErr" },
			// operators
			{ "(", "PtgParen" },
			{ "*", "PtgMlt" },
			{ "/", "PtgDiv" },
			{ "^", "PtgPower" },
			{ "&", "PtgConcat" },
			{ "<>", "PtgNE" },
			{ "<=", "PtgLE" },
			{ "<", "PtgLT" },
			{ ">=", "PtgGE" },
			{ ">", "PtgGT" },
			{ "=", "PtgEQ" },
			{ "!=", "PtgNE" },
			//{" ","PtgIsect"}, intersection operator, need to work out how to use a space as operator-- SEE PTGMEMFUNC/MEMAREA -- only valid when parsing complex ranges
			//{",","PtgUnion"},  problems matching as a separator in string parsing
			//{":","PtgRange"}, // may have issues with ptgArea?
			{ "+", "PtgAdd" },
			// moved to AFTER other operators
			{ "-", "PtgSub" },
			//operands
			{ "PtgStr", "PtgStr" },
			{ "PtgNumber", "PtgNumber" },
			{ "PtgInt", "PtgInt" },
			{ "PtgRef", "PtgRef" },
			{ "PtgArea", "PtgArea" },
			{ "false", "PtgBool" },
			{ "true", "PtgBool" },
			{ "PtgArea3d", "PtgArea3d" },
			{ "PtgRef3d", "PtgRef3d" },
			{ "PtgArray", "PtgArray" },
			{ "PtgMissArg", "PtgMissArg" },
			{ "PtgMemFunc", "PtgMemFunc" },
			{ "PtgAtr", "PtgAtr" },
			//functions
			{ "PtgFunc", "PtgFunc" },
			{ "PtgFuncVar", "PtgFuncVar" },
			// unary prefix operators:  see FormulaParser.splitString
			{ "u+", "PtgUPlus" },
			{ "u-", "PtgUMinus" },
			{ ")", "PtgParen" },

	};

	/*  DO DIFFERENTLY SO TO AVOID HANGING OBJECT REFERENCES
		static {
			HashMap recmap = new HashMap();

			// Most Frequent
			case BLANK), new Blank() );
			case ROW), new Row() );
			case XF), new Xf() );
			case INDEX), new Index() );
			case COUNTRY), new Country() );
			case CALCMODE), new CalcMode() );
			case DIMENSIONS), new Dimensions() );
			case SELECTION), new Selection() );
			case DEFAULTROWHEIGHT),  new DefaultRowHeight() );
			case DEFCOLWIDTH), new DefColWidth() );
			case DBCELL), new Dbcell() );
			case BOF), new Bof() );
			case BOUNDSHEET), new Boundsheet() );
			case EOF), new Eof() );
			case FORMAT), new Format() );
			case STYLE), new Style() );
			case PASSWORD), new Password() );
			case PALETTE), new Palette() );
			case ARRAY), new Array() );
			case BOOLERR), new Boolerr() );
			case EXTERNSHEET), new Externsheet() );
			case EXTERNNAME), new Externname() );
			case FORMULA), new Formula() );
			case LABEL), new Label() );
			case TXO), new Txo() );
			case CONTINUE), new Continue() );
			case SST), new Sst() );
			case GUTS), new Guts() );
			case EXTSST), new Extsst() );
			case HLINK), new Hlink() );
			case LABELSST), new Labelsst() );
			case NUMBER), new NumberRec() );
			case MERGEDCELLS), new Mergedcells() );
			case MULBLANK), new Mulblank() );
			case MULRK), new Mulrk() );
			case RK), new Rk() );
			case RSTRING), new Rstring() );
			case SHRFMLA), new Shrfmla() );
			case STRINGREC), new StringRec() );
			case SUPBOOK), new Supbook() );
			case DV), new Dv() );
			case DVAL), new Dval() );
			case SETUP), new Setup() );
			case HCENTER), new HCenter() );
			case VCENTER), new VCenter() );
			case LEFTMARGIN), new LeftMargin() );
			case RIGHTMARGIN), new RightMargin() );
			case TOPMARGIN), new TopMargin() );
			case BOTTOMMARGIN), new BottomMargin() );
			case PRINTGRID), new PrintGrid() );
			case PRINTROWCOL), new PrintRowCol() );
			case XCT), new Xct() );
			case CRN), new Crn() );
			case NOTE), new Note() );

			// Named Ranges and References
			case NAME), new Name() );

			// Workbook Settings
			case FONT), new Font() );
			case DSF), new Dsf() );
			case WINDOW1), new Window1() );
			case WINDOW2), new Window2() );
			case CODENAME), new Codename() );
			case PROTECT), new Protect() );
			case OBJPROTECT), new ObjProtect() );
			case SCENPROTECT), new ScenProtect() );
			case FEATHEADR), new FeatHeadr() );
			case PROT4REV), new Prot4rev() );
			case COLINFO), new Colinfo() );
			case USERSVIEWBEGIN), new Usersviewbegin() );
			case USERSVIEWEND), new Usersviewend() );
			case WSBOOL), new WsBool() );
			case BOOKBOOL), new BookBool() );
			case USR_EXCL), new UsrExcl() );
			case INTERFACE_HDR), new InterfaceHdr() );
			case RRD_INFO), new RrdInfo() );
			case RRD_HEAD), new RrdHead() );
			case FILE_LOCK), new FileLock() );
			case PLS), new Pls() );
			case HEADERREC), new Headerrec() );
			case DATE1904), new NineteenOhFour() );

			// Sheet Settings
			case OBJ), new Obj() );
			case OBPROJ), new Obproj() );
			case FOOTERREC), new Footerrec() );
			case TABID), new TabID() );
			case PANE),  new Pane() );
			case SCL),  new Scl() );

			// Protection settings
			case FILEPASS), new Filepass() );

			// Conditional Formatting
			case CF), new Cf() );
			case CONDFMT), new Condfmt() );

			// Auto filter
			case AUTOFILTER),  new AutoFilter() );

			// Chart Records
			case CHART), new Chart() );
			case SERIES), new Series() );
			case SERIESTEXT), new SeriesText() );
			case SERIESLIST), new SeriesList() );
			case AI), new Ai() );
			case BEGIN), new Begin() );
			case END), new End() );
			case UNITS), new Units() );
			case CHART), new Chart() );
			case DATAFORMAT), new DataFormat() );
			case LINEFORMAT), new LineFormat() );
			case MARKERFORMAT), new MarkerFormat() );
			case AREAFORMAT), new AreaFormat() );
			case PIEFORMAT), new PieFormat() );
			case ATTACHEDLABEL), new AttachedLabel() );
			case CHARTFORMAT), new ChartFormat() );
			case LEGEND), new Legend() );
			case BAR), new Bar() );
			case LINE), new Line() );
			case PIE), new Pie() );
			case AREA), new Area() );
			case SCATTER), new Scatter() );
			case CHARTLINE), new ChartLine() );
			case AXIS), new Axis() );
			case TICK), new Tick() );
			case VALUERANGE), new ValueRange() );
			case CATSERRANGE), new CatserRange() );
			case AXISLINEFORMAT), new AxisLineFormat() );
			case CHARTFORMATLINK), new ChartFormatLink() );
			case DEFAULTTEXT), new DefaultText() );
			case TEXTDISP), new TextDisp() );
			case FONTX), new Fontx() );
			case OBJECTLINK), new ObjectLink() );
			case FRAME), new Frame() );
			case BEGIN), new Begin() );
			case END), new End() );
			case PLOTAREA), new PlotArea() );
			case THREED), new ThreeD() );
			case PICF), new Picf() );
			case DROPBAR), new Dropbar() );
			case RADAR), new Radar() );
			case SURFACE), new Surface() );
			case RADARAREA), new RadarArea() );
			case AXISPARENT), new AxisParent() );
			case LEGENDXN), new Legendxn() );
			case SHTPROPS), new ShtProps() );
			case SERTOCRT), new SerToCrt() );
			case AXESUSED), new Axesused() );
			case SBASEREF), new SbaseRef() );
			case SERPARENT), new SerParent() );
			case SERAUXTREND), new SerauxTrend() );
			case IFMT), new Ifmt() );
			case POS), new Pos() );
			case ALRUNS), new AlRuns() );
			case AI), new Ai() );
			case SERAUXERRBAR), new SerauxErrBar() );
			case SERFMT), new Serfmt() );
			case CHART3DBARSHAPE), new Chart3DBarShape() );
			case FBI), new Fbi() );
			case BOPPOP), new Boppop() );
			case AXCENT), new Axcent() );
			case DAT), new Dat() );
			case PLOTGROWTH), new PlotGrowth() );
			case SIIINDEX), new SiIndex() );
			case GELFRAME), new GelFrame() );
			case BOPPOPCUSTOM), new BoppopCustom() );
			case FONTBASIS),  new FontBasis() );

			// PivotTable Records
			case SXVIEW), new Sxview() );
			case SXFORMAT), new Sxformat() );
			case SXLI), new Sxli() );
			case SXVI), new Sxvi() );
			case SXVD), new Sxvd() );
			case SXIVD), new Sxivd() );

			// Object and Picture Records
			case PHONETIC), new Phonetic() );
			case MSODRAWING), new MSODrawing() );
			case MSODRAWINGGROUP), new MSODrawingGroup() );
			case MSODRAWINGSELECTION), new MSODrawingSelection() );

			// Excel 9 Chart Records
			case CHARTFRTINFO), new ChartFrtInfo() );
			case FRTWRAPPER), new FrtWrapper() );
			case STARTBLOCK), new StartBlock() );
			case ENDBLOCK), new EndBlock() );
			case STARTOBJECT), new StartObject() );
			case ENDOBJECT), new EndObject() );
			case CATLAB), new CatLab() );
			case YMULT), new YMult() );
			case SXVIEWLINK), new SxViewLink() );
			case PIVOTCHARTBITS), new PivotChartBits() );
			case FRTFONTLIST), new FrtFontList() );
			case PIVOTCHARTLINK), new PivotChartLink() );
			case DATALABEXTCONTENTS), new DataLabExtContents() );
			case DATALABEXT), new DataLabExt() );

			records = Collections.unmodifiableMap( recmap );
		}
	*/
	/*
	 * Create a ptg record from a name, will be init'ed elsewhere if needed.  
	 * I am keeping this seperate from getBiffRecord for performance reasons.  Why search
	 * through all the ptg/formula stuff every time you deal with a XLS record an vice-versa.  Also,
	 * init'ing may be different.  Small duplication of code, but I think it is worth it.
	 */
	public static Ptg getPtgRecord( String name ) throws InvalidRecordException
	{
		for( String[] aPtgLookup : ptgLookup )
		{
			if( aPtgLookup[0].equalsIgnoreCase( name ) )
			{
				try
				{
					String classname = aPtgLookup[1];
					return (Ptg) Class.forName( "org.openxls.formats.XLS.formulas." + classname ).newInstance();
				}
				catch( Exception e )
				{
					throw new InvalidRecordException( "ERROR: Creating Record: " + name + "failed: " + e.toString() );
				}
			}
		}
		return null;
	}

	/**
	 * Get an instance of the record type corresponding to the given opcode.
	 *
	 * @param opcode the BIFF8 record opcode to be resolved
	 * @return an instance of the class corresponding to the given opcode
	 * or an XLSRecord if the opcode is unknown
	 * @throws RuntimeException if instantiation of the record fails
	 */
	public static BiffRec getBiffRecord( short opcode )
	{
		// TRY THIS:
		BiffRec record = null;
		try
		{
			switch( opcode )
			{
				case BLANK:
					record = new Blank();
					break;
				case ROW:
					record = new Row();
					break;
				case XF:
					record = new Xf();
					break;
				case INDEX:
					record = new Index();
					break;
				case COUNTRY:
					record = new Country();
					break;
				case CALCMODE:
					record = new CalcMode();
					break;
				case DIMENSIONS:
					record = new Dimensions();
					break;
				case SELECTION:
					record = new Selection();
					break;
				case DEFAULTROWHEIGHT:
					record = new DefaultRowHeight();
					break;
				case DEFCOLWIDTH:
					record = new DefColWidth();
					break;
				case DBCELL:
					record = new Dbcell();
					break;
				case BOF:
					record = new Bof();
					break;
				case BOUNDSHEET:
					record = new Boundsheet();
					break;
				case EOF:
					record = new Eof();
					break;
				case FORMAT:
					record = new Format();
					break;
				case STYLE:
					record = new Style();
					break;
				case PASSWORD:
					record = new Password();
					break;
				case PALETTE:
					record = new Palette();
					break;
				case ARRAY:
					record = new Array();
					break;
				case BOOLERR:
					record = new Boolerr();
					break;
				case EXTERNSHEET:
					record = new Externsheet();
					break;
				case EXTERNNAME:
					record = new Externname();
					break;
				case FORMULA:
					record = new Formula();
					break;
				case LABEL:
					record = new Label();
					break;
				case TXO:
					record = new Txo();
					break;
				case CONTINUE:
					record = new Continue();
					break;
				case SST:
					record = new Sst();
					break;
				case GUTS:
					record = new Guts();
					break;
				case EXTSST:
					record = new Extsst();
					break;
				case HLINK:
					record = new Hlink();
					break;
				case LABELSST:
					record = new Labelsst();
					break;
				case NUMBER:
					record = new NumberRec();
					break;
				case MERGEDCELLS:
					record = new Mergedcells();
					break;
				case MULBLANK:
					record = new Mulblank();
					break;
				case MULRK:
					record = new Mulrk();
					break;
				case RK:
					record = new Rk();
					break;
				case RSTRING:
					record = new Rstring();
					break;
				case SHRFMLA:
					record = new Shrfmla();
					break;
				case STRINGREC:
					record = new StringRec();
					break;
				case SUPBOOK:
					record = new Supbook();
					break;
				case DV:
					record = new Dv();
					break;
				case DVAL:
					record = new Dval();
					break;
				case SETUP:
					record = new Setup();
					break;
				case HCENTER:
					record = new HCenter();
					break;
				case VCENTER:
					record = new VCenter();
					break;
				case LEFTMARGIN:
					record = new LeftMargin();
					break;
				case RIGHTMARGIN:
					record = new RightMargin();
					break;
				case TOPMARGIN:
					record = new TopMargin();
					break;
				case BOTTOMMARGIN:
					record = new BottomMargin();
					break;
				case PRINTGRID:
					record = new PrintGrid();
					break;
				case PRINTROWCOL:
					record = new PrintRowCol();
					break;
				case XCT:
					record = new Xct();
					break;
				case CRN:
					record = new Crn();
					break;
				case NOTE:
					record = new Note();
					break;

				// Named Ranges and References
				case NAME:
					record = new Name();
					break;

				// Workbook Settings
				case FONT:
					record = new Font();
					break;
				case DSF:
					record = new Dsf();
					break;
				case WINDOW1:
					record = new Window1();
					break;
				case WINDOW2:
					record = new Window2();
					break;
				case PLV:
					record = new PLV();
					break;
				case CODENAME:
					record = new Codename();
					break;
				case PROTECT:
					record = new Protect();
					break;
				case OBJPROTECT:
					record = new ObjProtect();
					break;
				case SCENPROTECT:
					record = new ScenProtect();
					break;
				case FEATHEADR:
					record = new FeatHeadr();
					break;
				case PROT4REV:
					record = new Prot4rev();
					break;
				case COLINFO:
					record = new Colinfo();
					break;
				case USERSVIEWBEGIN:
					record = new Usersviewbegin();
					break;
				case USERSVIEWEND:
					record = new Usersviewend();
					break;
				case WSBOOL:
					record = new WsBool();
					break;
				case BOOKBOOL:
					record = new BookBool();
					break;
				case USR_EXCL:
					record = new UsrExcl();
					break;
				case INTERFACE_HDR:
					record = new InterfaceHdr();
					break;
				case RRD_INFO:
					record = new RrdInfo();
					break;
				case RRD_HEAD:
					record = new RrdHead();
					break;
				case FILE_LOCK:
					record = new FileLock();
					break;
				case PLS:
					record = new Pls();
					break;
				case HEADERREC:
					record = new Headerrec();
					break;
				case DATE1904:
					record = new NineteenOhFour();
					break;

				// Sheet Settings
				case OBJ:
					record = new Obj();
					break;
				case OBPROJ:
					record = new Obproj();
					break;
				case FOOTERREC:
					record = new Footerrec();
					break;
				case TABID:
					record = new TabID();
					break;
				case PANE:
					record = new Pane();
					break;
				case SCL:
					record = new Scl();
					break;

				// Conditional Formatting
				case CF:
					record = new Cf();
					break;
				case CONDFMT:
					record = new Condfmt();
					break;

				// Auto filter
				case AUTOFILTER:
					record = new AutoFilter();
					break;

				// Chart Records
				case CHART:
					record = new Chart();
					break;
				case SERIES:
					record = new Series();
					break;
				case SERIESTEXT:
					record = new SeriesText();
					break;
				case SERIESLIST:
					record = new SeriesList();
					break;
				case AI:
					record = new Ai();
					break;
				case UNITS:
					record = new Units();
					break;
				case DATAFORMAT:
					record = new DataFormat();
					break;
				case LINEFORMAT:
					record = new LineFormat();
					break;
				case MARKERFORMAT:
					record = new MarkerFormat();
					break;
				case AREAFORMAT:
					record = new AreaFormat();
					break;
				case PIEFORMAT:
					record = new PieFormat();
					break;
				case ATTACHEDLABEL:
					record = new AttachedLabel();
					break;
				case CHARTFORMAT:
					record = new ChartFormat();
					break;
				case LEGEND:
					record = new Legend();
					break;
				case BAR:
					record = new Bar();
					break;
				case LINE:
					record = new Line();
					break;
				case PIE:
					record = new Pie();
					break;
				case AREA:
					record = new Area();
					break;
				case SCATTER:
					record = new Scatter();
					break;
				case CHARTLINE:
					record = new ChartLine();
					break;
				case AXIS:
					record = new Axis();
					break;
				case TICK:
					record = new Tick();
					break;
				case VALUERANGE:
					record = new ValueRange();
					break;
				case CATSERRANGE:
					record = new CatserRange();
					break;
				case AXISLINEFORMAT:
					record = new AxisLineFormat();
					break;
				case CHARTFORMATLINK:
					record = new ChartFormatLink();
					break;
				case DEFAULTTEXT:
					record = new DefaultText();
					break;
				case TEXTDISP:
					record = new TextDisp();
					break;
				case FONTX:
					record = new Fontx();
					break;
				case OBJECTLINK:
					record = new ObjectLink();
					break;
				case FRAME:
					record = new Frame();
					break;
				case BEGIN:
					record = new Begin();
					break;
				case END:
					record = new End();
					break;
				case PLOTAREA:
					record = new PlotArea();
					break;
				case THREED:
					record = new ThreeD();
					break;
				case PICF:
					record = new Picf();
					break;
				case DROPBAR:
					record = new Dropbar();
					break;
				case RADAR:
					record = new Radar();
					break;
				case SURFACE:
					record = new Surface();
					break;
				case RADARAREA:
					record = new RadarArea();
					break;
				case AXISPARENT:
					record = new AxisParent();
					break;
				case LEGENDXN:
					record = new Legendxn();
					break;
				case SHTPROPS:
					record = new ShtProps();
					break;
				case SERTOCRT:
					record = new SerToCrt();
					break;
				case AXESUSED:
					record = new Axesused();
					break;
				case SBASEREF:
					record = new SbaseRef();
					break;
				case SERPARENT:
					record = new SerParent();
					break;
				case SERAUXTREND:
					record = new SerauxTrend();
					break;
				case IFMT:
					record = new Ifmt();
					break;
				case POS:
					record = new Pos();
					break;
				case ALRUNS:
					record = new AlRuns();
					break;
				case SERAUXERRBAR:
					record = new SerauxErrBar();
					break;
				case SERFMT:
					record = new Serfmt();
					break;
				case CHART3DBARSHAPE:
					record = new Chart3DBarShape();
					break;
				case FBI:
					record = new Fbi();
					break;
				case BOPPOP:
					record = new Boppop();
					break;
				case AXCENT:
					record = new Axcent();
					break;
				case DAT:
					record = new Dat();
					break;
				case PLOTGROWTH:
					record = new PlotGrowth();
					break;
				case SIIINDEX:
					record = new SiIndex();
					break;
				case GELFRAME:
					record = new GelFrame();
					break;
				case BOPPOPCUSTOM:
					record = new BoppopCustom();
					break;
				case FONTBASIS:
					record = new FontBasis();
					break;

				// PivotTable Records
				case SXVIEW:
					record = new Sxview();
					break;
				case TABLESTYLES:
					record = new TableStyles();
					break;
				case SXFORMAT:
					record = new Sxformat();
					break;
				case SXLI:
					record = new Sxli();
					break;
				case SXVI:
					record = new Sxvi();
					break;
				case SXVD:
					record = new Sxvd();
					break;
				case SXIVD:
					record = new Sxivd();
					break;
				case SXSTREAMID:
					record = new SxStreamID();
					break;
				case SXVS:
					record = new SxVS();
					break;
				case SXADDL:
					record = new SxAddl();
					break;
				case SXVDEX:
					record = new SxVdEX();
					break;
				case SXPI:
					record = new SxPI();
					break;
				case SXDI:
					record = new SxDI();
					break;
				case SXDB:
					record = new SxDB();
					break;
				case SXFDB:
					record = new SxFDB();
					break;
				case SXDBEX:
					record = new SXDBEx();
					break;
				case SXFDBTYPE:
					record = new SXFDBType();
					break;
				case SXSTRING:
					record = new SXString();
					break;
				case SXNUM:
					record = new SXNum();
					break;
				case SXDBB:
					record = new SxDBB();
					break;
				case SXEX:
					record = new SxEX();
					break;
				case QSISXTAG:
					record = new QsiSXTag();
					break;
				case SXVIEWEX9:
					record = new SxVIEWEX9();
					break;
				case DCONREF:
					record = new DConRef();
					break;
				case DCONNAME:
					record = new DConName();
					break;
				case DCONBIN:
					record = new DConBin();
					break;

				// Object and Picture Records
				case PHONETIC:
					record = new Phonetic();
					break;
				case MSODRAWING:
					record = new MSODrawing();
					break;
				case MSODRAWINGGROUP:
					record = new MSODrawingGroup();
					break;
				case MSODRAWINGSELECTION:
					record = new MSODrawingSelection();
					break;

				// Excel 9 Chart Records
				case CHARTFRTINFO:
					record = new ChartFrtInfo();
					break;
				case FRTWRAPPER:
					record = new FrtWrapper();
					break;
				case STARTBLOCK:
					record = new StartBlock();
					break;
				case ENDBLOCK:
					record = new EndBlock();
					break;
				case STARTOBJECT:
					record = new StartObject();
					break;
				case ENDOBJECT:
					record = new EndObject();
					break;
				case CATLAB:
					record = new CatLab();
					break;
				case YMULT:
					record = new YMult();
					break;
				case SXVIEWLINK:
					record = new SxViewLink();
					break;
				case PIVOTCHARTBITS:
					record = new PivotChartBits();
					break;
				case FRTFONTLIST:
					record = new FrtFontList();
					break;
				case PIVOTCHARTLINK:
					record = new PivotChartLink();
					break;
				case DATALABEXTCONTENTS:
					record = new DataLabExtContents();
					break;
				case DATALABEXT:
					record = new DataLabExt();
					break;
				case CRTLAYOUT12A:
					record = new CrtLayout12A();
					break;
				case CRTLAYOUT12:
					record = new CrtLayout12();
					break;
				default:
					record = new XLSRecord();
			}
			record.setOpcode( opcode );
			log.debug( "Record: " + record.getClass().getName() );
		}
		catch( Exception e )
		{
			throw new RuntimeException( "failed to instantiate record", e );
		}
		return record;
	}

}