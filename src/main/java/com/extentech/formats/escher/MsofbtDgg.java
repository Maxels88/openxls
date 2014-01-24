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
package com.extentech.formats.escher;

import com.extentech.toolkit.ByteTools;

//0xF006
public class MsofbtDgg extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7933328640935994167L;
	int spidMax = 1024;
	int numIdClusters;
	int numShapes;
	int numDrawings;
	/* 20071115 KSC: Unused at this point IdClusters[] clusters = new IdClusters[1];
	
	public class IdClusters{
		public int dgid,cspidCur;
	}
	*/

	public MsofbtDgg( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
		
		/* 20071115 KSC: THIS IS UNUSED RIGHT NOW
		//Data from experimental results
		this.clusters = new IdClusters[1];
		this.clusters[0] = new IdClusters();
		this.clusters[0].dgid = 1;
		this.clusters[0].cspidCur = 4;
		*/
	}

	@Override
	protected byte[] getData()
	{
		byte[] spidMaxBytes;
		byte[] cidclBytes;
		byte[] cspSavedBytes;
		byte[] cdgSavedBytes;
//		spidMaxBytes = ByteTools.cLongToLEBytes(spidMax+numShapes);	20071113 KSC: can't assume this
		spidMaxBytes = ByteTools.cLongToLEBytes( spidMax );
		numIdClusters = spidMax / 1024 + (((spidMax % 1024) != 0) ? 1 : 0);    // 20080903 KSC: # id clusters is based upon # shapes used
		cidclBytes = ByteTools.cLongToLEBytes( numIdClusters );
		cspSavedBytes = ByteTools.cLongToLEBytes( numShapes );
		cdgSavedBytes = ByteTools.cLongToLEBytes( numDrawings );

		// new code
		int lenOfFIDCL = numIdClusters - 1;
		byte[] retBytes = new byte[16 + (lenOfFIDCL * 8)];            // HOLDS initial info plus FIDCL array of 8 bytes*numIdClusters
		System.arraycopy( spidMaxBytes, 0, retBytes, 0, 4 );
		System.arraycopy( cidclBytes, 0, retBytes, 4, 4 );
		System.arraycopy( cspSavedBytes, 0, retBytes, 8, 4 );
		System.arraycopy( cdgSavedBytes, 0, retBytes, 12, 4 );
		int pos = 16;
		
	    /* 20071120 KSC: new code */
		//20071115 KSC: try to interpret -- these changes seems correct in byte comparisons ... but unsure if it's correct for all cases
		for( int i = 0; i < lenOfFIDCL; i++ )
		{
			System.arraycopy( ByteTools.cLongToLEBytes( 1 ), 0, retBytes, pos, 4 );        // dgid owning the SPID's in this cluster
			System.arraycopy( ByteTools.cLongToLEBytes( ((i == 0) ? (numShapes) : 1) ),
			                  0,
			                  retBytes,
			                  pos + 4,
			                  4 );    // # SPID's used so far
			pos += 8;
		}
		
	    /* old code 
	    byte[] b1,b2;
	   
	    b1 = ByteTools.cLongToLEBytes(1);	    
	    System.arraycopy(b1, 0, retBytes, pos, 4);
	    b2 = ByteTools.cLongToLEBytes(numShapes);  //TODO: Can you find any logic for this??	-- No!! remove!!
		System.arraycopy(b2, 0, retBytes, pos+4, 4);
		pos+=8;
		*/
		setLength( retBytes.length );
		return retBytes;
	}
	
	
	/* 20071115 KSC: THIS IS UNUSED RIGHT NOW
	public void setIdClusters(IdClusters[] c){
		clusters = c;
	}
	*/

	public int getNumDrawings()
	{
		return numDrawings;
	}

	public void setNumDrawings( int numDrawings )
	{
		this.numDrawings = numDrawings;
	}

	public int getNumIdClusters()
	{
		return numIdClusters;
	}
/* 20080903 KSC: numIdClusters is soley a function of spidMax/1024
	public void setNumIdClusters(int numIdClusters) {
		this.numIdClusters = numIdClusters;
	}
*/

	public int getNumShapes()
	{
		return numShapes;
	}

	public void setNumShapes( int numShapes )
	{
		this.numShapes = numShapes;
	}

	// 20071113 KSC

	/**
	 * set SpidMax
	 */
	public void setSpidMax( int spid )
	{
		spidMax = spid;
	}
}
