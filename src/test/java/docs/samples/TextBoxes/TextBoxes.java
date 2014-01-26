package docs.samples.TextBoxes;

import org.openxls.ExtenXLS.CommentHandle;
import org.openxls.ExtenXLS.ImageHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.formats.XLS.Boundsheet;
import org.openxls.formats.XLS.MSODrawing;
import org.openxls.formats.XLS.MSODrawingGroup;
import org.openxls.formats.XLS.WorkBook;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.openxls.formats.escher.MsofbtOPT;

import java.util.AbstractList;
import java.util.HashMap;
import java.util.List;

/**
 * User: npratt
 * Date: 11/27/13
 * Time: 16:26
 */
public class TextBoxes
{
	private String wdir = System.getProperty( "user.dir" ) + "/docs/samples/TextBoxes/";

	public static void main( String[] args ) throws WorkSheetNotFoundException
	{
		TextBoxes textBoxes = new TextBoxes();
		textBoxes.run();
	}

	private void run() throws WorkSheetNotFoundException
	{
		System.out.println( "wdir = " + wdir );

		WorkBookHandle wb = new WorkBookHandle( wdir + "TextBoxes.xls" );
		WorkSheetHandle sheet = wb.getWorkSheet( 0 );

		WorkBook workBook = wb.getWorkBook();
		MSODrawingGroup msoDrawingGroup = workBook.getMSODrawingGroup();
		System.out.println( "msoDrawingGroup = " + msoDrawingGroup );
		msoDrawingGroup.getNumShapes();

		CommentHandle[] commentHandles = sheet.getCommentHandles();
		System.out.println( "commentHandles = " + commentHandles.length );

		Boundsheet boundSheet = workBook.getWorkSheetByNumber( 0 );
		AbstractList<MSODrawing> msodrawingrecs = msoDrawingGroup.getMsodrawingrecs();
		for( MSODrawing msodrawingrec : msodrawingrecs )
		{
			msodrawingrec.createCommentBox( 1,1 );
			int i = 0;
			MsofbtOPT optRec = msodrawingrec.getOPTRec();
			if( optRec != null )
			{
				boolean b = optRec.hasTextId();
				System.out.println( "b = " + b );

			}
		}
		MSODrawing msoHeaderRec = msoDrawingGroup.getMsoHeaderRec( boundSheet );
		HashMap ooxmlShapes = boundSheet.getOOXMLShapes();
		List charts = boundSheet.getCharts();
		List dvRecs = boundSheet.getDvRecs();

//		System.out.println( "msoHeaderRec = " + msoHeaderRec.debugOutput() );
		String shapeName = msoHeaderRec.getShapeName();
		int shapeType = msoHeaderRec.getShapeType();

		int numShapes = msoHeaderRec.getNumShapes();
		System.out.println( "numShapes = " + numShapes );
		ImageHandle[] images = sheet.getImages();
		System.out.println( "images = " + images.length );
	}
}
