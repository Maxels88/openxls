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

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.charts.Ai;
import com.extentech.formats.XLS.charts.Chart;
import com.extentech.formats.XLS.formulas.GenericPtg;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgAreaErr3d;
import com.extentech.formats.XLS.formulas.PtgErr;
import com.extentech.formats.XLS.formulas.PtgName;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.formats.XLS.formulas.PtgRefErr;
import com.extentech.formats.XLS.formulas.PtgRefErr3d;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.Vector;

/**
 * This class is responsible for registering cell references (Ptgs) and managing
 * reference updates etc.
 * <p/>
 * <p/>
 * Here's the scenario:
 * <p/>
 * as we parse workbooks, WorkBook.addFormula() extracts all PtgRefs and
 * puts them in various ReferenceTracker maps.
 * <p/>
 * as cells are inserted/and or changed, referenced cells can be updated
 * after calling 'getAffectedCells'.
 * <p/>
 * TODO: retrofit all methods to use sheet_map cache
 * <p/>
 * <p/>
 * <p/>
 * All PtgRefs and Areas
 */
public class ReferenceTracker
{
	private static final Logger log = LoggerFactory.getLogger( ReferenceTracker.class );
	// the sheets allow for faster refs
	// each sheet contains a collection of rows.
	private Map sheetMap = new HashMap();
	// store ptgNames
	private Map nameRefs = new HashMap();

	// Database calc caches
	private Map criteriaDBs = new HashMap();
	private Map CollectionDBs = new HashMap();
	private Map vlookups = new HashMap();
	private Collection crs = new Vector();

	// VLOOKUPs and other lookups need to calc col ptgs
	private Map lookupColsCache = new HashMap();

	public Map getLookupColCache()
	{
		return lookupColsCache;
	}

	/**
	 * @return Returns the vlookups.
	 */
	public Map getVlookups()
	{
		return vlookups;
	}

	/**
	 * @return Returns the criteriaDBs.
	 */
	public Map getCriteriaDBs()
	{
		return criteriaDBs;
	}

	/**
	 * @return Returns the CollectionDBs.
	 */
	public Map getListDBs()
	{
		return CollectionDBs;
	}

	/**
	 * Blow out all the cacheing
	 */
	public void clearCaches()
	{
		// Should we?
		// crPtgMap     =   new HashMap();
		// refPtgMap    =   new HashMap();

		// Databases
		criteriaDBs = new HashMap();
		CollectionDBs = new HashMap();
		vlookups = new HashMap();
	}

	/**
	 * clear out VLOOKUP and related function caches
	 */
	public void clearLookupCaches()
	{
		lookupColsCache.clear();
		lookupColsCache = new HashMap();
		criteriaDBs.clear();
		criteriaDBs = new HashMap();
		CollectionDBs.clear();
		CollectionDBs = new HashMap();
		vlookups.clear();
		vlookups = new HashMap();
	}

	/**
	 * Returns ALL formulas on the cellhandles sheet that reference this CellHandle.
	 * <p/>
	 * Clears the cached value on said cells so a recalc will be forced upon any getVal method
	 * <p/>
	 * Please note that these cells have already been calculated, so in order
	 * to get their values without re-calculating them
	 * Extentech suggests setting the book level non-calculation flag, ie
	 * <p/>
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
	 * <p/>
	 * or
	 * <p/>
	 * FormulaHandle.getCachedVal()
	 *
	 * @return Collection of of calculated cells
	 */
	public synchronized Map clearAffectedFormulaCellsOnSheet( CellHandle cx, String sheetname )
	{
		HashMap hm = (HashMap) clearAffectedFormulaCells( cx );
		HashMap retmap = new HashMap();
		Iterator i = hm.keySet().iterator();
		while( i.hasNext() )
		{
			String s = (String) i.next();
			if( s.indexOf( sheetname ) > -1 )
			{
				retmap.put( s, hm.get( s ) );
			}
		}
		return retmap;
	}

	/**
	 * Returns a Collection Map of cells that are affected by formula
	 * references to this CellHandle.
	 * <p/>
	 * Clears the cached value on said cells and recalcs.
	 * <p/>
	 * Please note that these cells are calculated before return.
	 * <p/>
	 * To get their values without re-calculating them set the book level non-calculation flag, ie
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
	 * or
	 * FormulaHandle.getCachedVal()
	 *
	 * @return Map of calculated cells
	 */
	public synchronized Map clearAffectedFormulaCells( CellHandle cx )
	{
		return clearAffectedFormulaCells( cx.getCell(), new HashMap() );
	}

	/**
	 * Returns a Collection Map of cells that are affected by formula
	 * references to this CellHandle.
	 * <p/>
	 * Clears the cached value on said cells and recalcs.
	 * <p/>
	 * Please note that these cells are calculated before return.
	 * <p/>
	 * To get their values without re-calculating them set the book level non-calculation flag, ie
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT);
	 * or
	 * FormulaHandle.getCachedVal()
	 *
	 * @return Map of calculated cells
	 */
	public synchronized Map clearAffectedFormulaCells( BiffRec cx )
	{
		return clearAffectedFormulaCells( cx, new HashMap() );
	}

	/**
	 * Returns ALL formulas on the cellhandle's sheet that reference
	 * record changedRec.
	 * <p/>
	 * Clears the cached value on said cells so a recalc will be
	 * forced upon any getVal method
	 * <p/>
	 * Please note that these cells have not yet been calculated, so
	 * in order pass a null value of sheetname in order to get all sheets
	 *
	 * @return Collection of of calculated cells
	 */
	private synchronized Map clearAffectedFormulaCells( BiffRec changedRec, Map affectedCellHandles )
	{

		if( affectedCellHandles == null )
		{
			affectedCellHandles = new HashMap();
		}

		String newRecSheetName = changedRec.getSheet().getSheetName();
		// get ref collection for the sheet
		TrackedPtgs ptgRefs = (TrackedPtgs) sheetMap.get( GenericPtg.qualifySheetname( newRecSheetName ) );    // now tracked ptgs are stored per sheet
		if( ptgRefs == null )
		{
			return affectedCellHandles;
		}
		Iterator parents = ptgRefs.getParents( changedRec );    // finds ALL parents affected by cell newRec
		while( parents.hasNext() )
		{
			BiffRec br = (BiffRec) parents.next();
			short op = br.getOpcode();
			if( op == XLSRecord.NAME )
			{
				String theName = ((Name) br).getNameA();
				if( nameRefs.containsKey( theName ) )
				{
					// add all formulas that refer
					ArrayList list = (ArrayList) (nameRefs.get( theName )); // gets the ptgname
					for( Object aList : list )
					{
						BiffRec ptgParent = ((Ptg) aList).getParentRec();
						if( ptgParent.getOpcode() == XLSConstants.NAME )
						{
							continue; // a Named Range referencing another named range ... will be caught later
						}
						String adr = ptgParent.getSheet().getSheetName() + "!" + ptgParent.getCellAddress();
						if( affectedCellHandles.get( adr ) == null )
						{
							ReferenceTracker.addRec( ptgParent, affectedCellHandles );
							affectedCellHandles = clearAffectedFormulaCells( ptgParent,
							                                                 affectedCellHandles ); // recurse parent formula and get cells it affects
						}
					}
				}
			}
			else if( (op == XLSConstants.CONDFMT) || (op == XLSConstants.AI) )
			{    // ignore since these records are not themselves referenced
			}
			else if( op == XLSConstants.SHRFMLA )
			{     // Shared Formula references are now reference-tracked; to find specific formula affected, use Shrfmla.getAffected
				Shrfmla sh = (Shrfmla) br;
				Formula f = sh.getAffected( changedRec );
				if( f != null )
				{
					String adr = f.getSheet().getSheetName() + "!" + f.getCellAddress();
					if( !affectedCellHandles.containsKey( adr ) )
					{
						ReferenceTracker.addRec( f, affectedCellHandles );
						affectedCellHandles = clearAffectedFormulaCells( f,
						                                                           affectedCellHandles );    // recurse parent formula and get cells it affects
					}
				}
			}
			else
			{  // regular Formula
				if( br.getSheet() != null )
				{
					String adr = br.getSheet().getSheetName() + "!" + br.getCellAddress();
					if( !affectedCellHandles.containsKey( adr ) )
					{
						ReferenceTracker.addRec( br, affectedCellHandles );
						affectedCellHandles = clearAffectedFormulaCells( br,
						                                                           affectedCellHandles );    // recurse parent formula and get cells it affects
					}
				} // ignore no sheet
			}
		}
		return affectedCellHandles;
	}

	/**
	 * retrieve all chart-related (==Ai) references to the particular cell
	 *
	 * @param newRec cell to lookup references
	 * @return list of Ai records that reference cell
	 */
	public List<Ai> getChartReferences( BiffRec newRec )
	{
		String newRecSheetName = newRec.getSheet().getSheetName();
		ArrayList<Ai> ret = new ArrayList();
		// get ref collection for the sheet
		TrackedPtgs ptgRefs = (TrackedPtgs) sheetMap.get( GenericPtg.qualifySheetname( newRecSheetName ) );    // now tracked ptgs are stored per sheet
		if( ptgRefs == null )
		{
			return ret;
		}
		Iterator parents = ptgRefs.getParents( newRec );    // finds ALL parents affected by cell newRec
		while( parents.hasNext() )
		{
			BiffRec br = (BiffRec) parents.next();
			short op = br.getOpcode();
			if( op == XLSConstants.AI )
			{
				ret.add( (Ai) br );
			}
		}
		return ret;
	}

	/**
	 * Add to the collection of PtgNames for referenceTracker
	 */
	public void addPtgNameReference( PtgName p )
	{
		String name = p.getTextString().toUpperCase(); // case-insensitive
		Object refs = nameRefs.get( name );
		if( refs == null )
		{
			refs = new ArrayList();
			((ArrayList) refs).add( p );
			nameRefs.put( name, refs );
		}
		else
		{
			ArrayList ptgNames = (ArrayList) refs;
			if( !ptgNames.contains( p ) )
			{
				ptgNames.add( p );
			}
		}
	}

	/**
	 * Add a cell to the Collection of afected cell handles
	 *
	 * @param celly
	 */
	private static void addRec( BiffRec celly, Map affectedCellHandles )
	{
		String address = celly.getSheet().getSheetName() + "!" + celly.getCellAddress();
		affectedCellHandles.put( address, celly );
		// add the new val to the reftracker and clear any cached
		// formula vals pointing to it, but do not recalc them
		// yes get rid of them all...
		try
		{
			((Formula) celly).clearCachedValue();
		}
		catch( ClassCastException e )
		{
		}
		;
	}

	/**
	 * adds a cellrange Ptg (Area, Area3d etc.) to be tracked
	 * <p/>
	 * These records are stored in a row based lookup,
	 * this row map is looked up off actual row number, ie row 1 is get(1).
	 *
	 * @param cr
	 * @return
	 */
	public Ptg addCellRange( Ptg ptgRef )
	{
		// system setting to disable ref tracking...
		String trackprop = System.getProperty( WorkBookHandle.REFTRACK_PROP );
		if( trackprop != null )
		{
			if( trackprop.equals( "false" ) )
			{
				return ptgRef;
			}
		}

		if( !(ptgRef instanceof PtgRef) )
		{
			return ptgRef;
		}
		if( (ptgRef instanceof PtgAreaErr3d) || (ptgRef instanceof PtgRefErr3d) || (ptgRef instanceof PtgRefErr) )
		{
			return ptgRef;
		}

		String sheetname = "";
		try
		{
			try
			{
				sheetname = ((PtgRef) ptgRef).getSheetName();
				sheetname = GenericPtg.qualifySheetname( sheetname );
			}
			catch( Exception ex )
			{
				sheetname = "WorkBookRanges";
			}
			// fast fail erroneous Sheet refs
			if( sheetname.equals( "#REF!" ) )
			{
				return ptgRef;
			}

			TrackedPtgs ptgs = (TrackedPtgs) sheetMap.get( sheetname ); // now tracked ptgs are stored per sheet not per row
			if( ptgs == null )
			{
				ptgs = new TrackedPtgs( new LocationComparer() );
				sheetMap.put( sheetname, ptgs );
			}
			if( !ptgs.contains( ptgRef ) )                    // **no duplicates allowed** (matches on location+parent rec)
			{
				ptgs.add( ptgRef );
			}
		}
		catch( Exception e )
		{
		}
		return ptgRef;

	}

	/**
	 * Clears out the cached location of ptgrefs in the target
	 * sheet.  This is required for copied worksheets with ptgrefs
	 * contained internally
	 *
	 * @param targetSheet
	 */
	public void clearPtgLocationCaches( String targetSheet )
	{
		try
		{
			targetSheet = GenericPtg.qualifySheetname( targetSheet );
			Iterator ptgs = ((TrackedPtgs) sheetMap.get( targetSheet )).values().iterator();
//            Iterator ptgs= ((TrackedPtgs) sheetMap.get(targetSheet)).iterator();
			while( ptgs.hasNext() )
			{
				try
				{
					PtgRef p = (PtgRef) ptgs.next();
					p.clearLocationCache();
				}
				catch( Exception ex )
				{
				}
			}
		}
		catch( Exception e )
		{

		}
	}

	/**
	 * removes a cellrange Ptg (Area, Area3d etc.) to be tracked
	 *
	 * @param cr
	 */
	public void removeCellRange( Ptg cr )
	{
		if( !(cr instanceof PtgRef) )
		{
			return;
		}
		try
		{
			String sheetname = "";
			try
			{
				sheetname = ((PtgRef) cr).getSheetName();
				sheetname = GenericPtg.qualifySheetname( sheetname );
			}
			catch( Exception ex )
			{
				sheetname = "WorkBookRanges";
			}
			TrackedPtgs ptgs = (TrackedPtgs) sheetMap.get( sheetname );
			if( ptgs != null )
			{
				ptgs.remove( cr );
			}
		}
		catch( Exception e )
		{
			// this is common and not a problem normally - then we won't report a warning
			//Logger.logWarn("ReferenceTracker.removeCellRange failed for: " + cr.toString() +":"+ e);
		}
	}

	/**
	 * updates the tracked ptg by using a new parent record
	 *
	 * @param pr     original ptg contained in tracker
	 * @param parent new parent record of ptg
	 */
	public void updateInRefTracker( PtgRef pr, XLSRecord parent )
	{
		if( (pr instanceof PtgRefErr) || (pr instanceof PtgRefErr3d) )
		{
			return;
		}
		try
		{
			String sheetname = "";
			try
			{
				sheetname = pr.getSheetName();
				sheetname = GenericPtg.qualifySheetname( sheetname );
			}
			catch( Exception ex )
			{
				sheetname = "WorkBookRanges";
			}

			TrackedPtgs ptgs = (TrackedPtgs) sheetMap.get( sheetname );
			if( ptgs != null )
			{
				ptgs.update( pr, parent );
			}
		}
		catch( Exception e )
		{
			// this is common and not a problem normally - then we won't report a warning
			//Logger.logWarn("ReferenceTracker.removeCellRange failed for: " + cr.toString() +":"+ e);
		}
	}

	    
    
    /* CELL RANGE SECTION */

	/**
	 * Returns an Array of the CellRanges existing in this WorkBook
	 * specifically the Ranges referenced in Formulas, Charts, and
	 * Named Ranges.
	 * <p/>
	 * This is necessary to allow for automatic updating of references
	 * when adding/removing/moving Cells within these ranges, as well
	 * as shifting references to Cells in Formulas when Formula records
	 * are moved.
	 *
	 * @return all existing Cell Range references used in Formulas, Charts, and Names
	 */
	public CellRange[] getCellRanges()
	{
		CellRange[] ret = new CellRange[crs.size()];
		return (CellRange[]) crs.toArray( ret );
	}

	/**
	 * updateReferences
	 * Shifts, Expands or Contracts ALL affected ranges upon a row or col insert or delete.
	 * This will eventually eliminate the need for subequent
	 * shifting done in moveFormulaCellReferences]
	 * <p/>
	 * this will move the entire range down if range first row > startrow
	 * it will expand the range if range first row <= startrow and range second row <= startrow
	 *
	 * @param start     0-based start row
	 * @param shift     shift amount can be + or -
	 * @param thissheet
	 * @param shiftRow  true if shifting rows (false for columns)
	 */

	public static void updateReferences( int start, int shiftamount, Boundsheet thissheet, boolean shiftRow )
	{
		// shift is 0-based, so that references to row shiftrow-1 and up are shifted
		// claritas is different, 1-based + shifts shiftrow+1 hence shiftInclusive setting
		// NOTE: shared formula references are the only ones that are NOT shifted via
		// updateReferences since PtgRefN and PtgAreaN's are NOT included in the
		// referenceTracker collection
		boolean shiftInclusive = thissheet.isShiftInclusive();    // claritas-specific setting which directs us to expand ranges rather than shift when start of range==start
		boolean isExcel2008 = thissheet.getWorkBook().getIsExcel2007();    // limits are different between BIFF8 and Excel 2007
		if( shiftInclusive )
		{
			start++;    // make 1-based
		}

		HashSet updated = new HashSet(); // tracks which Ptgs have been already updated

		String sheetname = GenericPtg.qualifySheetname( thissheet.getSheetName() );
		TrackedPtgs trackedptgs = (TrackedPtgs) thissheet.getWorkBook().getRefTracker().sheetMap.get( sheetname );
		if( (trackedptgs == null) || (trackedptgs.size() == 0) )
		{
			return;
		}
		Object[] ptgs = null;

		ptgs = trackedptgs.toArray();

		int i = ((shiftamount > 0) ? (ptgs.length - 1) : 0);
		int end = ((shiftamount > 0) ? 0 : ptgs.length);
		int inc = ((shiftamount > 0) ? -1 : +1);
		boolean done = false;
		while( !done )
		{
			Ptg p = (Ptg) ptgs[i];
			// skip these
			if( (p instanceof PtgRefErr) || (p instanceof PtgRefErr3d) )  // these shouldn't be in the reference tracker ...
			{
				continue;
			}

			PtgRef pr = (PtgRef) p;
			if( !updated.contains( pr ) )
			{
				String sht;
				try
				{
					sht = pr.getSheetName();
				}
				catch( Exception e )
				{    // shouldn't happen
					log.error( "ReferenceTracker.updateReferences:  Error in Formula Reference Location: " + e.toString() );
					continue;
				}
				sht = GenericPtg.qualifySheetname( sht );
				if( sheetname.equals( sht ) )
				{
					if( shiftPtg( pr, sht, start, shiftamount, isExcel2008, shiftRow ) )
					{
						updated.add( pr );        // record which has already been updated to avoid incorrect expansion or movement
					}
				}
			}
			i += inc;
			if( shiftamount > 0 )
			{
				done = (i < 0);
			}
			else
			{
				done = (i == ptgs.length);
			}
		}
		// also update merged ranges which fall within range
		if( thissheet.hasMergedCells() )
		{
			// update mergedcells first, as they may grow in record size and are not handled by continues.
			Iterator itx = thissheet.getMergedCellsRecs().iterator();
			while( itx.hasNext() )
			{
				Mergedcells mrg = (Mergedcells) itx.next();
				CellRange[] rngs = mrg.getMergedRanges();
				for( CellRange rng : rngs )
				{
					try
					{
						int[] rc = rng.getRangeCoords();
						rc[0]--;
						rc[2]--;    // 1-based ...?
						boolean isRange = (rc.length > 2);
						boolean bUpdated = false;
						if( shiftRow )
						{
							if( rc[0] >= start )
							{    // shift
								rc[0] += shiftamount;
								if( isRange )
								{
									rc[2] += shiftamount;
								}
								bUpdated = true;
							}
							else if( isRange && (rc[2] >= start) )
							{ // expand
								rc[2] += shiftamount;
								bUpdated = true;
							}
						}
						if( bUpdated )
						{
							String newrange = thissheet + "!" + ExcelTools.formatLocation( rc );
							rng.setRange( newrange );
						}
					}
					catch( CellNotFoundException e )
					{
					}
				}
			}
		}
	}

	/**
	 * given a PtgRef, shifts correctly given start (row or col), shiftamount (+ or - 1) and truth of "shiftRow"
	 *
	 * @param ptgref      = the ptgref to move
	 * @param sht         = the sheet in which the ptgref resides
	 * @param start       = Start row for shifting - This is a 0 based value
	 * @param shiftamount = amount to shift the ptgref
	 * @param isExcel2007 true if use Excel-2007 maximums
	 * @param shiftRow    = ?
	 * @return true if updated PtgRef location
	 */
	public static boolean shiftPtg( PtgRef ptgref, String sht, int start, int shiftamount, boolean isExcel2007, boolean shiftRow )
	{
		int[] rc;
		int iParent = ptgref.getParentRec().getOpcode();
		boolean isNamedRange = (iParent == XLSConstants.NAME);
		boolean isAi = (iParent == XLSConstants.AI);
		boolean isShared = (iParent == XLSConstants.SHRFMLA);
		try
		{
			rc = ptgref.getIntLocation();
		}
		catch( Exception e )
		{    // shouldn't happen!
			if( !(ptgref instanceof PtgAreaErr3d) )    // if it's not already an error Ptg report error
			{
				log.error( "ReferenceTracker.shiftPtg:  Error in Formula Reference Location: " + e.toString() );
			}
			return false;
		}
		boolean isRange = (rc.length > 2);
		boolean bUpdated = false;
		if( shiftRow )
		{
			if( !isNamedRange && !isAi && !ptgref.isRowRel() )
			{
				return false;    // if absolute don't shift (except for names and ai/charting refs, which should expand or shift in all cases)
			}
			if( ((rc[0] + 1) == start) && isRange )
			{ // expand don't shift
				rc[2] += shiftamount;
				if( isAi && (rc[1] != rc[3]) )    // Series in ROWS get shifted, not expanded ...
				{
					rc[0] += shiftamount;
				}
				bUpdated = true;
			}
			else if( (rc[0] + 1) >= start )
			{    // shift
				rc[0] += shiftamount;
				if( isRange )
				{
					rc[2] += shiftamount;
				}
				bUpdated = true;
			}
			else if( isRange && ((rc[2] + 1) >= start) )
			{ // expand
				rc[2] += shiftamount;
				if( isAi && (rc[1] != rc[3]) )    // Series in ROWS get shifted, not expanded ...
				{
					rc[0] += shiftamount;
				}
				bUpdated = true;
			}
			// SHIFTING EXCEPTION: if the parent formula cell is located ON the shifting row, do not shift 
			if( bUpdated && (iParent == XLSConstants.FORMULA) && (ptgref.getParentRec().getRowNumber() == (start - 1)) )
			{
				bUpdated = false;
			}
		}
		else
		{ // deal with columns in same way as above
			if( !isNamedRange && !isAi && !ptgref.isColRel() )
			{
				return false;    // if absolute don't shift (except for names and ai/charting refs, which should expand or shift in all cases)
			}
			if( (rc[1] + 1) >= start )
			{
				rc[1] += shiftamount;
				if( isRange )
				{
					rc[3] += shiftamount;
				}
				bUpdated = true;
			}
			else if( isRange && ((rc[3] + 1) >= start) )
			{
				rc[3] += shiftamount;
				bUpdated = true;
			}
		}
		if( bUpdated )
		{
			// deal with limits
			if( isExcel2007 )
			{
				if( rc[0] >= XLSConstants.MAXROWS )
				{
					rc[0] = XLSConstants.MAXROWS - 1;
				}
				if( isRange && (rc[2] >= XLSConstants.MAXROWS) )
				{
					rc[2] = XLSConstants.MAXROWS - 1;
				}
			}
			else
			{
				if( rc[0] >= XLSConstants.MAXROWS_BIFF8 )
				{
					rc[0] = XLSConstants.MAXROWS_BIFF8 - 1;
				}
				if( isRange && (rc[2] >= XLSConstants.MAXROWS_BIFF8) )
				{
					rc[2] = XLSConstants.MAXROWS_BIFF8 - 1;
				}
			}
			String newaddr = ExcelTools.formatLocation( rc, ptgref.isRowRel(), ptgref.isColRel() );
			if( isRange && (newaddr.indexOf( ":" ) == -1) ) // handle special case of SHOULD be a range but 1st and last match
			{
				newaddr = newaddr + ":" + newaddr;
			}
			newaddr = sht + "!" + newaddr;
			// NOW UPDATE THE PTG LOCATION
			try
			{
				if( !isAi && !isShared )
				{
					ptgref.setLocation( newaddr );    // should update ref tracker (remove and add new) appropriately
					if( isNamedRange )
					{
						// find formula references to the named range -- contained in reftracker.nameRefs
						ReferenceTracker rt = ptgref.getParentRec().getWorkBook().getRefTracker();
						String theName = ((Name) ptgref.getParentRec()).getNameA();
						if( rt.nameRefs.containsKey( theName ) )
						{
							// add all formulas that refer
							ArrayList list = (ArrayList) (rt.nameRefs.get( theName )); // gets the ptgname
							for( Object aList : list )
							{
								BiffRec ptgParent = ((Ptg) aList).getParentRec();
								if( ptgParent.getOpcode() == XLSConstants.NAME )
								{
									continue; // a Named Range referencing another named range ... will be caught later
								}
								rt.clearAffectedFormulaCells( ptgParent ); // recurse parent formula and get cells it affects
							}
						}
					}
					else if( iParent == XLSConstants.CONDFMT )
					{
						((Condfmt) ptgref.getParentRec()).setDirty();    // flag to rebuild record
					}
				}
				else if( isShared )
				{
					((Shrfmla) ptgref.getParentRec()).updateLocation( shiftamount, ptgref );
				}
				else
				{    // Ai (chart) reference
					((Ai) ptgref.getParentRec()).changeAiLocation( ptgref, newaddr );
				}
			}
			catch( Exception e )
			{
				log.error( "ReferenceTracker.shiftPtg:  Shifting Formula Reference failed: " + e.toString() );
			}
		}

		return bUpdated;
	}

	/**
	 * insert chart series upon an insert row
	 * called by WSH.shiftRow
	 * if series are row-based
	 */
	public static void insertChartSeries( Chart c, String sht, int rownum )
	{
		int[] rc;
		PtgRef pr = null;
		String cursheet;
		boolean inserted = false;
		HashMap seriesmap = c.getSeriesPtgs();
		Iterator ii = seriesmap.keySet().iterator();
		while( ii.hasNext() && !inserted )
		{
			com.extentech.formats.XLS.charts.Series s = (com.extentech.formats.XLS.charts.Series) ii.next();
			Ptg[] ptgs = (Ptg[]) seriesmap.get( s );
			for( Ptg ptg : ptgs )
			{
				try
				{
					pr = (PtgRef) ptg;
					cursheet = pr.getSheetName();
					rc = pr.getIntLocation();
				}
				catch( Exception e )
				{    // shouldn't happen unless it's a PtgErr-type
					continue;
				}
				if( sht.equalsIgnoreCase( cursheet ) )
				{ // if series are in rows, if existing series fall within inserted row, add new series
					if( rc[0] == rownum )
					{ // already shifted ai range matches inserted row; take shifted range, shift backwards to get desired range to insert
						boolean isRange = (rc.length > 2);
						rc[0]--;
						if( isRange )
						{
							rc[2]--;
						}
						// Adjust Series/Values Rane
						String newseries = ExcelTools.formatLocation( rc, pr.isRowRel(), pr.isColRel() );
						if( isRange && (newseries.indexOf( ":" ) == -1) ) // handle special case of SHOULD be a range but 1st and last match
						{
							newseries = newseries + ":" + newseries;
						}
						newseries = sht + "!" + newseries;
						// Adjust Legend Range
						Ai legend = s.getLegendAi();
						rc = ExcelTools.getRowColFromString( legend.getDefinition() );
						if( rc[0] == rownum )
						{
							rc[0]--;
						}
						String legendRange = ExcelTools.formatLocation( rc );
						// Adjust Bubble Range if present
						Ai bubble = s.getBubbleValueAi();
						String bubbleRange = "";
						if( (bubble != null) && !bubble.getDefinition().equals( "" ) )
						{
							rc = ExcelTools.getRowColFromString( bubble.getDefinition() );
							if( rc[0] == rownum )
							{
								rc[0]--;
							}
							if( rc.length > 2 )
							{
								rc[2]--;
							}
							bubbleRange = ExcelTools.formatLocation( rc );
						}
						// Get Category range but don't alter
						String categoryRange = s.getCategoryValueAi().getDefinition();        // category shouldn't shift
						c.addSeries( newseries, categoryRange, bubbleRange, legendRange, "", 0 );    // 0= default chart
						c.setDimensionsRecord();
						inserted = true;
					}
				}
			}
		}
	}
	
/* FORMULA REFS */

	/**
	 * Update the address in a Ptg using the policy defined for the Ptg in an RDF if any.
	 * called by changeFormulaLocation, addCellToRange
	 *
	 * @param thisptg
	 * @param newaddr
	 */
	public static void updateAddressPerPolicy( Ptg thisptg, String newaddr )
	{
		int pl = thisptg.getLocationPolicy();
		// Logger.logInfo("ReferenceTracker.updateAddressPerPolicy called.");
		switch( pl )
		{
			case Ptg.PTG_LOCATION_POLICY_UNLOCKED:
				// if this ptg has not been locked, update it
				if( newaddr.indexOf( "#REF!" ) > -1 )
				{
					Formula formula = (Formula) thisptg.getParentRec();
					formula.replacePtg( thisptg, new PtgErr( PtgErr.ERROR_REF ) );
				}
				else
				{
					thisptg.setLocation( newaddr );
				}
				break;

			case Ptg.PTG_LOCATION_POLICY_TRACK:
				// this ptg tracks the cell that belongs to it...
				thisptg.updateAddressFromTrackerCell();
				break;

			case Ptg.PTG_LOCATION_POLICY_LOCKED:
				// do nothing
				break;
		}
//		}
	}

	/**
	 * IF the newly inserted cell is a Formula, we need to update its cell range(s)
	 * called by shiftRow
	 *
	 * @param copycell
	 * @param oldaddr
	 * @param newaddr
	 * @param rownum
	 * @param shiftRow true if shifting rows and not columns
	 * @throws Exception
	 * @see WorkSheetHandle.shiftRow
	 */
	public static void adjustFormulaRefs( CellHandle newcell, int newrownum, int offset, boolean shiftRow ) throws Exception
	{
		try
		{
			String sheet = newcell.getWorkSheetName();
			boolean isExcel2007 = newcell.getWorkBook().getWorkBook().getIsExcel2007();
			Ptg[] locptgs = newcell.getFormulaHandle().getFormulaRec().getCellRangePtgs();
			for( Ptg locptg : locptgs )
			{
				if( locptg instanceof PtgRef )
				{
					PtgRef pr = (PtgRef) locptg;
					shiftPtg( pr, sheet, newrownum, offset, isExcel2007, shiftRow );
				}
			}
		}
		catch( FormulaNotFoundException e )
		{
			//
		}
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		sheetMap.clear();
		nameRefs.clear();
		criteriaDBs.clear();
		CollectionDBs.clear();
		vlookups.clear();
		crs.clear();
		lookupColsCache.clear();
		sheetMap = new HashMap();
		nameRefs = new HashMap();
		// Database calc caches
		criteriaDBs = new HashMap();
		CollectionDBs = new HashMap();
		vlookups = new HashMap();
		crs = new Vector();
		lookupColsCache = new HashMap();
	}
}

/**
 * TrackedPtgs is a TreeMap override specific for PtgRefs to compare them more completely
 * (via location + parent-rec ...)
 * <p/>
 * Each PtgRef is identified via a hashcode created from it's row/column (or row/column pairs for a PtgArea)
 * In this way a given PtgRef can be determined whether it is contained already, can be removed, updated and
 * it's parent records gathered in an efficient way.
 * <p/>
 * TrackedPtgs must be instantiated with the custom Comparitor LocationComparer as:
 * new TrackedPtgs(new LocationComparer())
 * LocationComparer will return the correct compare for a PtgRef object, based upon it's location and parent record, in other words
 * PtgRef-A == PtgRef-B when both the location and the parent records are equal.
 */
class TrackedPtgs extends TreeMap
{
	private static final long serialVersionUID = 1L;
	static final long SECONDPTGFACTOR = ((XLSRecord.MAXCOLS + ((long) XLSRecord.MAXROWS * XLSRecord.MAXCOLS)));

	/**
	 * set the custom Comparitor for tracked Ptgs
	 * Tracked Ptgs are referened by a unique key that is based upon it's location and it's parent record
	 *
	 * @param c
	 */
	public TrackedPtgs( Comparator c )
	{
		super( c );
	}

	/**
	 * single access point for unique key creation from a ptg location and the ptg's parent
	 * would be so much cleaner without the double precision issues ... sigh
	 *
	 * @param o
	 * @return
	 */
	private Object getKey( Object o ) throws IllegalArgumentException
	{
		long loc = ((PtgRef) o).hashcode;
		if( loc == -1 )    // may happen on a referr (should have been caught earlier) or if not initialized yet
		{
			throw new IllegalArgumentException();
		}
		long ploc = ((PtgRef) o).getParentRec().hashCode();
		return new long[]{ loc, ploc };
	}

	/**
	 * access point for unique key creation from separate identities
	 *
	 * @param loc  -- location key for a ptgref
	 * @param ploc -- location key for a parent rec
	 * @return Object key
	 */
	private Object getKey( long loc, long ploc )
	{
		return new long[]{ loc, ploc };
	}

	/**
	 * override of add to record ptg location hash + parent has for later lookups
	 */
	public boolean add( Object o )
	{
		try
		{
			super.put( getKey( o ), o );
		}
		catch( IllegalArgumentException e )
		{    // SHOULD NOT HAPPEN -- happens upon RefErrs but they shouldnt be added ...
			// 	TESTING: report error
//	System.err.println("Illegal PtgRef Location: " + o.toString());			
		}
		return true;
	}

	/**
	 * override of the contains method to look up ptg location + parent record via hashcode
	 * to see if it is already contained within the store
	 */
	public boolean contains( Object o )
	{
		if( super.containsKey( getKey( o ) ) )
		{
			return true;
		}
		return false;
	}

	/**
	 * returns EVERY parent that references the cell "cell"
	 *
	 * @param cell
	 * @return iterator of biffrec parents of cell
	 */
	// attempt to avoid concurrentmod exception by generating arrays, but doesn't work 
//    public Object[] getParents(BiffRec cell) {
	public Iterator getParents( BiffRec cell )
	{
		ArrayList parents = new ArrayList();
		int[] rc = { cell.getRowNumber(), cell.getColNumber() };
		long loc = PtgRef.getHashCode( rc[0], rc[1] );    // get location in hashcode notation
		// first see if have tracked ptgs at the test location -- match all regardless of parent rec ...
		Object key;
		Map m = Collections.synchronizedMap( subMap( getKey( loc, 0 ), getKey( loc + 1, 0 ) ) );        // +1 for max parent
		if( (m != null) && (m.size() > 0) )
		{
			Iterator ii = m.keySet().iterator();
			while( ii.hasNext() )
			{
				key = ii.next();
				long testkey = ((long[]) key)[0];
/*			Object[] keys= m.keySet().toArray();
			for (int i= 0; i < keys.length; i++) {
				long testkey= ((long[])keys[i])[0];*/
				if( testkey == loc )
				{    // longs to remove parent hashcode portion of double
					parents.add( ((PtgRef) get( key )).getParentRec() );
//					parents.add(((PtgRef)this.get(keys[i])).getParentRec());
//System.out.print(": Found ptg" + this.get((Integer)locs.get(key)));					
				}
				else
				{
					break;    // shouldn't hit here
				}
			}
		}
		// now see if test cell falls into any areas

		m = Collections.synchronizedMap( tailMap( getKey( SECONDPTGFACTOR, 0 ) ) ); // NOW GET ALL PTGAREAS ...
		if( m != null )
		{
			synchronized(m.keySet())
			{
				Iterator ii = m.keySet().iterator();
				while( ii.hasNext() )
				{
					key = ii.next();
					long testkey = ((long[]) key)[0];
/*			Object[] keys= m.keySet().toArray();
			for (int i= 0; i < keys.length; i++) {
					long testkey= ((long[])keys[i])[0];*/
					double firstkey = testkey / SECONDPTGFACTOR;
					double secondkey = (testkey % SECONDPTGFACTOR);
					if( ((long) firstkey <= loc) && ((long) secondkey >= loc) )
					{
						int col0 = (int) firstkey % XLSRecord.MAXCOLS;
						int col1 = (int) secondkey % XLSRecord.MAXCOLS;
						int rw0 = ((int) (firstkey / XLSRecord.MAXCOLS)) - 1;
						int rw1 = ((int) (secondkey / XLSRecord.MAXCOLS)) - 1;
						if( isaffected( rc, new int[]{ rw0, col0, rw1, col1 } ) )
						{
							parents.add( ((PtgRef) get( key )).getParentRec() );
//							parents.add(((PtgRef)this.get(keys[i])).getParentRec());
						}
					}
					else if( firstkey > loc ) // we're done
					{
						break;
					}
				}
			}
		}
//System.out.println("");		
//		return parents.toArray();	//iterator();
		return parents.iterator();
	}

	/**
	 * returns true if cell coordinates are contained within the area coordinates
	 *
	 * @param cellrc
	 * @param arearc
	 * @return
	 */
	private boolean isaffected( int[] cellrc, int[] arearc )
	{
		if( cellrc[0] < arearc[0] )
		{
			return false; // row above the first ref row?
		}
		if( cellrc[0] > arearc[2] )
		{
			return false; // row after the last ref row?
		}

		if( cellrc[1] < arearc[1] )
		{
			return false; // col before the first ref col?
		}
		if( cellrc[1] > arearc[3] )
		{
			return false; // col after the last ref col?
		}
		return true;
	}

	/**
	 * remove this PtgRef object via it's key
	 */
	@Override
	synchronized public Object remove( Object o )
	{
		return super.remove( getKey( o ) );
	}

	/**
	 * update the location key for this PtgRef based upon a new parent record
	 *
	 * @param o      ptgref object
	 * @param parent parent of ptgref
	 */
	public void update( Object o, XLSRecord parent )
	{
		try
		{
			remove( o );
			long newloc = parent.hashCode();
			put( getKey( ((PtgRef) o).hashcode, newloc ), o );
		}
		catch( IllegalArgumentException e )
		{
// TESTING: report error
			//System.err.println("Illegal PtgRef Location: " + o.toString());			
		}
	}

	public Object[] toArray()
	{
		return values().toArray();
	}

	/**
	 * avoid double comparisons by converting double value to long
	 * @param value
	 * @return
	 * NOTE: doesn't work in all cases -- shelve for now
	 *
	private long hashCode(long location, double parentlocation) {
	long bits = Double.doubleToLongBits(location+parentlocation);
	return (long)(bits ^ (bits >>> 32));
	}*/
}

/**
 * custom comparitor which compares keys for TrackerPtgs
 * consisting of a long ptg location hash, long parent record hashcode
 */
class LocationComparer implements Comparator
{
	@Override
	public int compare( Object o1, Object o2 )
	{
		long[] key1 = (long[]) o1;
		long[] key2 = (long[]) o2;
		if( key1[0] < key2[0] )
		{
			return -1;
		}
		if( key1[0] > key2[0] )
		{
			return 1;
		}
		if( key1[0] == key2[0] )
		{
			if( key1[1] == key2[1] )
			{
				return 0;    // equals
			}
			if( key1[1] < key2[1] )
			{
				return -1;
			}
			if( key1[1] > key2[1] )
			{
				return 1;
			}
		}
		return -1;
	}
}




