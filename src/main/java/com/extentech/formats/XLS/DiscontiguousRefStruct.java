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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.formulas.PtgArea;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.toolkit.ByteTools;

import java.io.Serializable;
import java.util.Comparator;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

/**
 * DiscontiguousRefStruct manages discontiguous ranges within one sheet.  Each of the ranges is managed via a Ref object,
 * be it range or single cell references.
 */
public class DiscontiguousRefStruct implements Serializable
{

	private static final long serialVersionUID = -7923448634000437926L;
	// KSC: try to decrease processing time by using refPtgs treemap  private ArrayList sqrefs = new ArrayList();
	private refPtgs allrefs = new refPtgs( new refPtgComparer() );
	XLSRecord parentRec = null;

	/**
	 * Handles creating a SqRefStruct via a string passed in, this should be in the format
	 * reference1,reference2,
	 * or
	 * 'A1:A5,G21,H33'
	 */
	public DiscontiguousRefStruct( String ranges, XLSRecord parentRec )
	{
		this.parentRec = parentRec;
		String[] refs = ranges.split( "," );
		for( int i = 0; i < refs.length; i++ )
		{
			try
			{
				if( !refs[i].equals( "" ) )
				{
/* KSC: try to decrease processing time by using refPtgs treemap         			SqRef pref = new SqRef(refs[i],parentRec);
        			sqrefs.add(pref);*/
					allrefs.add( refs[i], parentRec );
				}
			}
			catch( NumberFormatException e )
			{
				;//keep going
			}
		}
	}

	/**
	 * Returns the refs in R1C1 format
	 *
	 * @return
	 */
	public String[] getRefs()
	{
/* KSC: try to decrease processing time by using refPtgs treemap         			    	String[] s = new String[sqrefs.size()];
        for (int i=0;i<s.length;i++) {
                s[i] = ((SqRef)sqrefs.get(i)).getLocation();
        }*/

		String[] s = new String[allrefs.size()];
		Iterator ptgs = allrefs.values().iterator();
		int i = 0;
		while( ptgs.hasNext() )
		{
			try
			{
				PtgRef pr = (PtgRef) ptgs.next();
				s[i++] = pr.getLocation();
			}
			catch( Exception ex )
			{
			}
		}
		return s;
	}

	/**
	 * Takes a binary array representing a sqref struct.
	 * <p/>
	 * 2               rgbSqref            var         Array of 8 byte sqref structures, format shown below
	 * in sqref class
	 *
	 * @param sqrefs
	 */
	public DiscontiguousRefStruct( byte[] sqrefrec, XLSRecord parentRec )
	{
		this.parentRec = parentRec;
		for( int i = 0; i < sqrefrec.length; )
		{
			byte[] sref = new byte[8];
			System.arraycopy( sqrefrec, i, sref, 0, 8 );
/* KSC: try to decrease processing time by using refPtgs treemap            SqRef s = new SqRef(sref, parentRec);
            sqrefs.add(s);*/
			allrefs.add( sref, parentRec );
			i += 8;
		}
	}

	/**
	 * Returns the binary record of this SQRef for output to Biff8
	 */
	public byte[] getRecordData()
	{
/* KSC: try to decrease processing time by using refPtgs treemap 
 		short s = (short) sqrefs.size();
        byte[] retData = ByteTools.shortToLEBytes(s);
        for (int i=0;i<sqrefs.size();i++) {
            SqRef sqr = (SqRef)sqrefs.get(i);
            retData = ByteTools.append(sqr.getRecordData(), retData);
         }*/
//        byte[] retData = ByteTools.shortToLEBytes((short)allrefs.size()); now done in updateData
//        Object[] refs= allrefs.values().toArray();
//        for (int i= refs.length-1; i >= 0; i--) {

		byte[] retData = new byte[0];
		Iterator ptgs = allrefs.values().iterator();
		while( ptgs.hasNext() )
		{
			try
			{
				PtgRef pr = (PtgRef) ptgs.next();
//              retData = ByteTools.append(getRecordData((PtgRef)refs[i]), retData);
				retData = ByteTools.append( retData, getRecordData( pr ) );
			}
			catch( Exception ex )
			{
			}
		}
		return retData;
	}

	/**
	 * Return the number of references this structure contains
	 *
	 * @return
	 */
	public int getNumRefs()
	{
//        return sqrefs.size();
		return allrefs.size();
	}

	/**
	 * Return toString as an array of references
	 */
	public String toString()
	{
		String result = "[";
/* KSC: TODO FINISH        Iterator i = sqrefs.iterator();
        while(i.hasNext()) {
            result += i.next().toString();
            if(i.hasNext())result += ", ";
        }*/
		result += "]";
		return result;
	}

	/**
	 * Adds a ref to the existing group of refs
	 *
	 * @param range
	 */
	public void addRef( String range )
	{
/* KSC: try to decrease processing time by using refPtgs treemap            SqRef sr = new SqRef(range, this.parentRec);
        sqrefs.add(sr);*/
		allrefs.add( range, this.parentRec );
	}

	/**
	 * Determines if the reference structure encompasses the
	 * reference value passed in.
	 *
	 * @param rowcol
	 * @return
	 */
	public boolean containsReference( int[] rowcol )
	{
/* KSC: try to decrease processing time by using refPtgs treemap                	
    	for (int i=0;i<sqrefs.size();i++) {
            SqRef sqr = (SqRef)sqrefs.get(i);
            if (sqr.contains(rowcol))return true;
        }*/
		if( allrefs.containsReference( rowcol ) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Get the bounding rectangle of all references in this structure in the format
	 * {topRow,leftCol,bottomRow,rightCol}
	 * <p/>
	 * Note that if you are including 3dReferences this could return inconsistent results
	 *
	 * @return
	 */
	public int[] getRowColBounds()
	{
/* KSC: try to decrease processing time by using refPtgs treemap                	
        int[] retValues = {0,0,0,0};
        for (int i=0;i<sqrefs.size();i++) {
            SqRef sref = (SqRef)sqrefs.get(i);
            int[] locs = sref.getIntLocation();
            for (int x=0;x<2;x++) {
                if((locs[x]<retValues[x])||i==0)retValues[x]=locs[x];
            }
            for (int x=2;x<4;x++) {
                if((locs[x]>retValues[x])||i==0)retValues[x]=locs[x];
            }
        }*/
/*        Object[] refs= allrefs.values().toArray();
        for (int i= refs.length-1; i >= 0; i--) {*/
		int[] retValues = { Integer.MAX_VALUE, Integer.MAX_VALUE, 0, 0 };
		Iterator ptgs = allrefs.values().iterator();
		int i = 0;
		while( ptgs.hasNext() )
		{
			PtgRef pr = (PtgRef) ptgs.next();
//        	PtgRef pr= (PtgRef)refs[i];
			int[] locs = pr.getIntLocation();
			for( int x = 0; x < 2; x++ )
			{
				if( (locs[x] < retValues[x]) )
				{
					retValues[x] = locs[x];
				}
			}
			for( int x = 2; x < 4; x++ )
			{
				if( (locs[x] > retValues[x]) )
				{
					retValues[x] = locs[x];
				}
			}
			i++;
		}
		return retValues;
	}

	private byte[] getRecordData( PtgRef myPtg )
	{
		byte[] retData = new byte[0];
		int[] rc = myPtg.getRowCol();
		if( rc[0] >= 65536 )// TODO: if XLSX, would this get here????
		{
			rc[0] = -1;
		}
		if( rc[2] >= 65536 )
		{
			rc[2] = -1;
		}
		retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[0] ), retData );
		retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[2] ), retData );
		retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[1] ), retData );
		retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[3] ), retData );
		return retData;
	}

	/**
	 * SQRef is a Ref (PtgArea) with specific methods to get and set via a different byte array than
	 * ptgArea normally uses.
	 * <p/>
	 * Sqref Structures
	 * <p/>
	 * OFFSET       NAME            SIZE        CONTENTS
	 * -----
	 * 2            rwFirst             2           First row in reference
	 * 4            rwLast              2           Last row in reference
	 * 6            colFirst            2           First column in reference
	 * 8            colLast             2           Last column in reference
	 */
	class SqRef implements Serializable
	{
		private static final long serialVersionUID = -7923448634000437926L;
		PtgRef myPtg;

		/**
		 * Construct an Sqref structure from a range string
		 *
		 * @param range
		 * @param parentRec
		 */
		public SqRef( String range, XLSRecord parentRec )
		{
			// add handling for different ref types, for now utilize a PtgArea
			if( (range != null) && !range.equals( "" ) )
			{
				myPtg = new PtgArea( range, parentRec );
				myPtg.addToRefTracker();
			}
		}

		public String toString()
		{
			return myPtg.toString();
		}

		/**
		 * Constructor for usage with any sort of ptgRef
		 *
		 * @param ptg
		 */
		public SqRef( PtgRef ptg )
		{
			myPtg = ptg;
		}

		public String getLocation()
		{
			return myPtg.getLocation();
		}

		/**
		 * Determine if the ref contains the rowcol passed in
		 *
		 * @return
		 */
		public boolean contains( int[] rowcol )
		{
			if( myPtg instanceof PtgArea )
			{
				return ((PtgArea) myPtg).contains( rowcol );
			}
			return false;
		}

		public int[] getIntLocation()
		{
			return myPtg.getIntLocation();
		}

		/**
		 * Construct an Sqref structure from a byte array
		 * in the format specified above for an sqref
		 * <p/>
		 * We do a string conversion for the absolute reference translation
		 * which all (?) sqrefs share.
		 *
		 * @param range
		 * @param parentRec
		 */
		public SqRef( byte[] b, XLSRecord parentRec )
		{
			int row = ByteTools.readShort( b[0], b[1] );
			int col = ByteTools.readShort( b[4], b[5] );
			int row2 = ByteTools.readShort( b[2], b[3] );
			if( row2 < 0 ) // if row truly references MAXROWS_BIFF8 comes out -
			{
				row2 += XLSConstants.MAXROWS_BIFF8;
			}

			int col2 = ByteTools.readShort( b[6], b[7] );
			int[] rc = { row, col, row2, col2 };
			boolean[] bool = { true, true, true, true };
			String location = ExcelTools.formatRangeRowCol( rc, bool );
			myPtg = new PtgArea( location, parentRec );
			myPtg.addToRefTracker();
		}

		/**
		 * Get the sqref as a byte array in the standardized SqRef structure
		 *
		 * @return
		 */
		public byte[] getRecordData()
		{
			byte[] retData = new byte[0];
			int[] rc = myPtg.getRowCol();
			if( rc[0] >= 65536 )// TODO: if XLSX, would this get here????
			{
				rc[0] = -1;
			}
			if( rc[2] >= 65536 )
			{
				rc[2] = -1;
			}
			retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[0] ), retData );
			retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[2] ), retData );
			retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[1] ), retData );
			retData = ByteTools.append( ByteTools.shortToLEBytes( (short) rc[3] ), retData );
			return retData;
		}

	}

}

class refPtgs extends TreeMap implements Serializable
{
	private static final long serialVersionUID = -7923448634000437926L;
	static final long SECONDPTGFACTOR = (((long) XLSRecord.MAXCOLS + ((long) XLSRecord.MAXROWS * XLSRecord.MAXCOLS)));

	/**
	 * set the custom Comparitor for tracked Ptgs
	 * Tracked Ptgs are referened by a unique key that is based upon it's location and it's parent record
	 *
	 * @param c
	 */
	public refPtgs( Comparator c )
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
			super.put( this.getKey( o ), o );
		}
		catch( IllegalArgumentException e )
		{    // SHOULD NOT HAPPEN -- happens upon RefErrs but they shouldnt be added ...
		}
		return true;
	}

	public boolean add( byte[] b, XLSRecord parentRec )
	{
		int row = ByteTools.readShort( b[0], b[1] );
		int col = ByteTools.readShort( b[4], b[5] );
		int row2 = ByteTools.readShort( b[2], b[3] );
		if( row2 < 0 ) // if row truly references MAXROWS_BIFF8 comes out -
		{
			row2 += XLSConstants.MAXROWS_BIFF8;
		}

		int col2 = ByteTools.readShort( b[6], b[7] );
		int[] rc = { row, col, row2, col2 };
		boolean[] bool = { true, true, true, true };
		String location = ExcelTools.formatRangeRowCol( rc, bool );
		PtgArea pa = new PtgArea( location, parentRec );
		pa.addToRefTracker();
		return this.add( pa );
	}

	public boolean add( String range, XLSRecord parentRec )
	{
		// add handling for different ref types, for now utilize a PtgArea
		if( (range != null) && !range.equals( "" ) )
		{
			PtgArea pa = new PtgArea( range, parentRec );
			pa.addToRefTracker();
			return this.add( pa );
		}
		return false;
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

	public boolean containsReference( int[] rc )
	{
		long loc = PtgRef.getHashCode( rc[0], rc[1] );    // get location in hashcode notation
		Map m = this.subMap( getKey( loc, 0 ), getKey( loc + 1, 0 ) );    // +1 for max parent
		Object key;
		if( (m != null) && (m.size() > 0) )
		{
			Iterator ii = m.keySet().iterator();
			while( ii.hasNext() )
			{
				key = ii.next();
				long testkey = ((long[]) key)[0];
				if( testkey == loc )
				{    // longs to remove parent hashcode portion of double
//System.out.print(": Found ptg" + this.get((Integer)locs.get(key)));
					return true;
				}
				break;    // shouldn't hit here
			}
		}
		// now see if test cell falls into any areas
		m = this.tailMap( getKey( SECONDPTGFACTOR, 0 ) ); // ALL AREAS ...
		if( m != null )
		{
			Iterator ii = m.keySet().iterator();
			while( ii.hasNext() )
			{
				key = ii.next();
				long testkey = ((long[]) key)[0];
				double firstkey = testkey / SECONDPTGFACTOR;
				double secondkey = (testkey % SECONDPTGFACTOR);
				if( ((long) firstkey <= (long) loc) && ((long) secondkey >= (long) loc) )
				{
					int col0 = (int) firstkey % XLSRecord.MAXCOLS;
					int col1 = (int) secondkey % XLSRecord.MAXCOLS;
					int rw0 = ((int) (firstkey / XLSRecord.MAXCOLS)) - 1;
					int rw1 = ((int) (secondkey / XLSRecord.MAXCOLS)) - 1;
					if( this.isaffected( rc, new int[]{ rw0, col0, rw1, col1 } ) )
					{
//System.out.print(": Found area " + ((PtgRef)this.get(index)));
						return true;
					}
				}
				else if( firstkey > loc ) // we're done
				{
					break;
				}
			}
		}
//System.out.println("");		

		return false;
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
	public Object remove( Object o )
	{
		return super.remove( getKey( o ) );
	}

	public Object[] toArray()
	{
		return this.values().toArray();
	}
}

/**
 * custom comparitor which compares keys for TrackerPtgs
 * consisting of a long ptg location hash, long parent record hashcode
 */
class refPtgComparer implements Comparator, Serializable
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
