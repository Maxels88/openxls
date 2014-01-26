package org.openxls.formats.XLS;

import java.util.List;

/**
 * User: npratt
 * Date: 1/24/14
 * Time: 12:56
 */
public interface CellsByCol
{
	public List<CellRec> get( int col );
	public void add( CellRec cell );
	public CellRec remove( CellRec cell );
}
