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
package org.openxls.formats.XLS.formulas;

import java.util.ArrayList;
import java.util.Arrays;

/*
  Computes the intersection of the two top operands.  Essentially
  this is a space operator.  Makes me think of space and drums, just 
  about the only thing more boring than these binary operand PTG's.
   
   
 * @see Ptg
 * @see Formula

    
*/
public class PtgIsect extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2131759675781833457L;

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
	}

	public PtgIsect()
	{
		ptgId = 0xF;
		record = new byte[1];
		record[0] = 0xF;
	}

	public String toString()
	{
		return getString();
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return " ";
	}

	@Override
	public int getLength()
	{
		return PTG_ISECT_LENGTH;
	}

	/**
	 * Intersection = Where A and B are shared.
	 * The ISECT operator (space)
	 * Returns the intersected range of two ranges. If the resulting cell
	 * range is empty, the formula will return the error code “#NULL!” (for instance A1:A2 B3).
	 * A1:B2 B2:C3 ==> B2
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
					if( !((PtgArea) p).wholeCol || ((PtgArea) p).wholeRow )
					{
						Ptg[] pc = p.getComponents();
						if( pc != null )
						{
							for( Ptg aPc : pc )
							{
								((PtgRef) aPc).setSheetName( ((PtgArea) p).getSheetName() );
								a.add( aPc );
							}
						}
					}
					else
					{
						a.add( p );    // TODO: what?????????
					}
				}
				else if( p instanceof PtgRef3d )
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
					if( pc != null )
					{
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
								Ptg[] pcs = pa.getComponents();
								for( Ptg pc1 : pcs )
								{
									((PtgRef) pc1).setSheetName( pa.getSheetName() );
									a.add( pc1 );
								}
							}
						}
					}
				}
				else if( (p instanceof PtgErr) || (p instanceof PtgRefErr) || (p instanceof PtgAreaErr3d) )
				{
					// DO WHAT???
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
			// For performance reasons, instantiate a PtgMystery as a lightweight GenericPtg which holds intermediary values in it's vars
			GenericPtg retp = new PtgMystery();
			ArrayList retptgs = new ArrayList();
			for( Object aFirst : first )
			{
				PtgRef pr = (PtgRef) aFirst;
				int[] rc = pr.getIntLocation();
				for( int m = 0; m < last.size(); m++ )
				{
					PtgRef pc = (PtgRef) last.get( m );
					int[] rc2 = pc.getIntLocation();
					if( Arrays.equals( rc, rc2 ) )
					{
						retptgs.add( pc );
						last.remove( m );
						m--;
					}
				}
			}
			Ptg[] ptgs = new Ptg[retptgs.size()];
			retptgs.toArray( ptgs );
			retp.setVars( ptgs );
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
}