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

import java.util.ArrayList;

/*
 * Computes the Union of the two top operands. Supposedly this is
 * "Microsoft Excel's comma operator" Brought to you by the same people that
 * created those space saving RK's
 * 
 * 
 * "http://www.extentech.com">Extentech Inc.</a>
 * 
 * 
 * 
 * @see Ptg
 * 
 * @see Formula
 */
public class PtgUnion extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8333035819099274707L;

	public boolean getIsOperator()
	{
		return true;
	}

	public boolean getIsBinaryOperator()
	{
		return true;
	}

	public boolean getIsPrimitiveOperator()
	{
		return true;
	}

	public PtgUnion()
	{
		ptgId = 0x10;
		record = new byte[1];
		record[0] = 0x10;
	}

	/**
	 * return the human-readable String representation of
	 */
	public String getString()
	{
		return ",";
	}

	public int getLength()
	{
		return PTG_UNION_LENGTH;
	}

	/**
	 * Union = All items from A and B, shared and not shared. Range list
	 * operator, represented by the system's list separator sign (for example
	 * comma sign). Treats two ranges as one operator (A1:B2,B2:C3) B1:B3
	 * ==>B1:B2, B2:B3 (A1:B2,B2:C3) ==>A1:B2, B2:C3
	 */
	// just add together? seems that is the case
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
				if( this.getParentRec().getSheet() != null )    // could be a Name rec ...
				{
					sourceSheet = this.getParentRec().getSheet().getSheetName();
				}
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
					Ptg[] pc = ((PtgName) p).getComponents();
					for( int j = 0; j < pc.length; j++ )
					{
						a.add( pc[j] );
					}
				}
				else if( p instanceof PtgStr )
				{
					String[] comps = (p.toString()).split( "," );
					for( int j = 0; j < comps.length; j++ )
					{
						if( comps[j].indexOf( ":" ) == -1 )
						{
							if( !comps[j].equals( "#REF!" ) && !comps[j].equals( "#NULL!" ) )
							{
								PtgRef3d pr = new PtgRef3d( false );
								pr.setParentRec( this.getParentRec() );
								pr.setLocation( comps[j] );
								a.add( pr );
							}
							else
							{
								PtgRefErr3d pr = new PtgRefErr3d();
								pr.setParentRec( this.getParentRec() );
								a.add( pr );
							}
						}
						else
						{
							PtgArea3d pa = new PtgArea3d( false );
							pa.setParentRec( this.getParentRec() );
							pa.setLocation( comps[j] );
							Ptg[] pcs = pa.getComponents();
							if( pcs != null )
							{
								for( int k = 0; k < pcs.length; k++ )
								{
									((PtgRef) pcs[k]).setSheetName( pa.getSheetName() );
									a.add( pcs[k] );
								}
							}
						}
					}
				}
				else if( p instanceof PtgArray )
				{
					// parse array components and create refs
					Ptg[] pc = ((PtgArray) p).getComponents();
					for( int j = 0; j < pc.length; j++ )
					{
						String loc = ((PtgStr) pc[j]).toString();
						if( !loc.startsWith( "#" ) )
						{ // skip errors -- TODO: IS THIS CORRECT?
							if( loc.indexOf( ":" ) == -1 )
							{
								if( loc.indexOf( "!" ) == -1 )
								{
									PtgRef pr = new PtgRef();
									pr.setUseReferenceTracker( false );
									pr.setParentRec( this.getParentRec() );
									pr.setLocation( loc );
									a.add( pr );
								}
								else
								{
									PtgRef3d pr = new PtgRef3d( false );
									pr.setParentRec( this.getParentRec() );
									pr.setLocation( loc );
									a.add( pr );
								}
							}
							else
							{
								PtgArea3d pa = new PtgArea3d( false );
								pa.setParentRec( this.getParentRec() );
								pa.setLocation( loc );
								a.add( pa );
							}
						}
					}
				}
				else if( p instanceof PtgErr || p instanceof PtgRefErr || p instanceof PtgAreaErr3d )
				{
					// DO WHAT???
					;
				}
				else
				{        // if an intermediary value returned from PtgRange, PtgUnion or PtgIsect, will be a GenericPtg which holds intermediary values in its vars array
					Ptg[] pc = ((GenericPtg) p).vars;
					for( int j = 0; j < pc.length; j++ )
					{
						if( pc[j] instanceof PtgArea & !(pc[j] instanceof PtgAreaErr3d) )
						{
							Ptg[] pa = pc[j].getComponents();
							for( int k = 0; k < pa.length; k++ )
							{
								a.add( pa[k] );
							}
						}
						else
						{
							a.add( pc[j] );
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
			// now have components for both operands - SUM and package into 1 ptg for return
			// For performance reasons, instantiate a PtgMystery as a lightweight GenericPtg which holds intermediary values in it's vars
			GenericPtg retp = new PtgMystery();
			ArrayList retptgs = new ArrayList();
			for( int k = 0; k < first.size(); k++ )
			{
				retptgs.add( (PtgRef) first.get( k ) );
			}
			for( int k = 0; k < last.size(); k++ )
			{
				retptgs.add( (PtgRef) last.get( k ) );
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