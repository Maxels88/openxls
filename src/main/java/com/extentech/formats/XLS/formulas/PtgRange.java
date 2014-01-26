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
package com.extentech.formats.XLS.formulas;

import com.extentech.ExtenXLS.ExcelTools;

import java.util.ArrayList;

/*
   Computes the minimal bounding rectangle of the top two operands.
   This is excel's ":" colon operator.
   
   
 * @see Ptg
 * @see Formula

    
*/
public class PtgRange extends GenericPtg implements Ptg
{

	private static final long serialVersionUID = 7181427387507157013L;

	@Override
	public boolean getIsOperator()
	{
		return true;
	}

	@Override
	public boolean getIsBinaryOperator()
	{
		return true;
	}

	@Override
	public boolean getIsPrimitiveOperator()
	{
		return true;
	}    // 20091019 KSC
   /*? public boolean getIsOperand(){return true;}
    public boolean getIsControl(){return true;}
    */

	public PtgRange()
	{
		ptgId = 0x11;
		record = new byte[1];
		record[0] = 0x11;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return ":";
	}

	@Override
	public int getLength()
	{
		return PTG_RANGE_LENGTH;
	}

	/**
	 * The RANGE operator (:)
	 * Returns the minimal rectangular range that contains both parameters.
	 * A1:B2:B2:C3 ==>A1:C3
	 * <p/>
	 * NOTE: assumption is NO 3d refs i.e. all on same sheet *******
	 */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		if( form.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}

		try
		{
			String sheet = null;
			String sourceSheet = null;
			try
			{
				sourceSheet = getParentRec().getSheet().getSheetName();
			}
			catch( NullPointerException ne )
			{
				;
			}
			ArrayList first = null;
			ArrayList last = null;
			for( int i = 0; i < 2; i++ )
			{
				Ptg p = form[i];
				ArrayList a = new ArrayList();
				if( p instanceof PtgArea )
				{
					a.add( p );
				}
				else if( p instanceof PtgRef )
				{
					a.add( p );
				}
				else if( p instanceof PtgName )
				{
					Ptg[] pc = p.getComponents();
					for( Ptg aPc : pc )
					{
						a.add( aPc );
					}
				}
				else if( p instanceof PtgStr )
				{
					String[] comps = (p.toString()).split( "," );
					for( String comp : comps )
					{
						if( comp.indexOf( ":" ) == -1 )
						{
							if( !comp.equals( "#REF!" ) && !comp.equals( "#NULL!" ) )
							{
								PtgRef3d pr = new PtgRef3d( false );
								pr.setParentRec( getParentRec() );
								pr.setLocation( comp );
								a.add( pr );
							}
							else
							{
								PtgRefErr3d pr = new PtgRefErr3d();
								pr.setParentRec( getParentRec() );
								a.add( pr );
							}
						}
						else
						{
							PtgArea3d pa = new PtgArea3d( false );
							pa.setParentRec( getParentRec() );
							pa.setLocation( comp );
							Ptg[] pcs = pa.getComponents();
							if( pcs != null )
							{
								for( Ptg pc : pcs )
								{
									((PtgRef) pc).setSheetName( pa.getSheetName() );
									a.add( pc );
								}
							}
						}
					}
				}
				else if( p instanceof PtgArray )
				{
					// parse array components and create refs
					Ptg[] pc = p.getComponents();
					for( Ptg aPc : pc )
					{
						String loc = aPc.toString();
						if( loc.indexOf( ":" ) == -1 )
						{
							if( loc.indexOf( "!" ) == -1 )
							{
								PtgRef pr = new PtgRef();
								pr.setUseReferenceTracker( false );
								pr.setParentRec( getParentRec() );
								pr.setLocation( loc );
								a.add( pr );
							}
							else
							{
								PtgRef3d pr = new PtgRef3d( false );
								pr.setParentRec( getParentRec() );
								pr.setLocation( loc );
								a.add( pr );
							}
						}
						else
						{
							PtgArea3d pa = new PtgArea3d( false );
							pa.setParentRec( getParentRec() );
							pa.setLocation( loc );
							a.add( pa );
						}
					}
				}
				else if( (p instanceof PtgErr) || (p instanceof PtgRefErr) || (p instanceof PtgAreaErr3d) )
				{
					// DO WHAT???
					; // ignore
				}
				else
				{        // if an intermediary value returned from PtgRange, PtgUnion or PtgIsect, will be a GenericPtg which holds intermediary values in its vars array
					Ptg[] pc = ((GenericPtg) p).vars;
					for( Ptg aPc : pc )
					{
						if( (aPc instanceof PtgArea) & !(aPc instanceof PtgAreaErr3d) )
						{
							Ptg[] pa = aPc.getComponents();
							for( Ptg aPa : pa )
							{
								a.add( aPa );
							}
						}
						else
						{
							a.add( aPc );
						}
					}
				}

				if( first == null )
				{
					first = a;
				}
				else
				{
					last = a;
				}
			}
			// now have components for both operands
			// range op returns the range that encompasses all referenced ptgs
			int[] rng = new int[]{ Short.MAX_VALUE, Short.MAX_VALUE, 0, 0 };
			for( Object aFirst : first )
			{
				PtgRef pr = (PtgRef) aFirst;
				if( sheet == null )
				{
					sheet = pr.getSheetName();    // TODO: 3d ranges??????
				}
				int[] rc = pr.getIntLocation();
				if( rc.length > 2 )
				{ // it's a range
					int numrows = (rc[2] - rc[0]) + 1;
					int numcols = (rc[3] - rc[1]) + 1;
					int numcells = numrows * numcols;
					if( numcells < 0 )
					{
						numcells *= -1; // handle swapped cells ie: "B1:A1"
					}
					int rowctr = rc[0];
					int cellctr = rc[1] - 1;
					for( int i = 0; i < numcells; i++ )
					{
						if( cellctr == rc[3] )
						{// if its the end of the row,increment row.
							cellctr = rc[1] - 1;
							rowctr++;
						}
						++cellctr;
						int[] addr = new int[]{ rowctr, cellctr };
						adjustRange( addr, rng );
					}
				}
				else
				{
					adjustRange( rc, rng );
				}
			}
			for( Object aLast : last )
			{
				PtgRef pr = (PtgRef) aLast;
				if( sheet == null )
				{
					sheet = pr.getSheetName();    // TODO: 3d ranges??????
				}
				int[] rc = pr.getIntLocation();
				if( rc.length > 2 )
				{ // it's a range
					if( rc.length > 2 )
					{ // it's a range
						int numrows = (rc[2] - rc[0]) + 1;
						int numcols = (rc[3] - rc[1]) + 1;
						int numcells = numrows * numcols;
						if( numcells < 0 )
						{
							numcells *= -1; // handle swapped cells ie: "B1:A1"
						}
						int rowctr = rc[0];
						int cellctr = rc[1] - 1;
						for( int i = 0; i < numcells; i++ )
						{
							if( cellctr == rc[3] )
							{// if its the end of the row,increment row.
								cellctr = rc[1] - 1;
								rowctr++;
							}
							++cellctr;
							int[] addr = new int[]{ rowctr, cellctr };
							adjustRange( addr, rng );
						}
					}
				}
				else
				{
					adjustRange( rc, rng );
				}
			}
			// For performance reasons, instantiate a PtgMystery as a lightweight GenericPtg which holds intermediary values in it's vars
			GenericPtg retp = new PtgMystery();
			PtgArea3d pa = new PtgArea3d( false );
			pa.setParentRec( getParentRec() );
			// TODO: 3d ranges????
			pa.setSheetName( sheet );
			pa.setLocation( rng );
			retp.setVars( new Ptg[]{ pa } );
			return retp;
		}
		catch( NumberFormatException e )
		{
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
		catch( Exception e )
		{    // handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}

	private static void adjustRange( int[] rc, int[] rng )
	{
		if( ExcelTools.isBeforeRange( rc, rng ) )
		{
			rng[0] = rc[0];
			rng[1] = rc[1];
		}
		if( ExcelTools.isAfterRange( rc, rng ) )
		{
			rng[2] = rc[0];
			rng[3] = rc[1];
		}
	}

}