import java.io.IOException;

import com.qoppa.pdf.PDFException;
import com.qoppa.word.WordDocument;
import com.qoppa.word.WordException;
import com.sun.xml.internal.fastinfoset.algorithm.BuiltInEncodingAlgorithm.WordListener;

import jPDFImagesSamples.PDFImagesSample;


public class Test {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		 PDFImagesSample sample = new PDFImagesSample();
		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.jpg",
				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.pdf");
		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.jpeg",
				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.jpeg.pdf");
		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.gif",
				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.gif.pdf");
//No Reader
//		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.tiff",
//				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.tiff.pdf");
//not supported		 
//		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.bmp",
//				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.bmp.pdf");
		 sample.createPDFFromImage("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.png",
				 "c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.png.pdf");
		 
		 WordDocument wd = null;
		try {
			wd = new WordDocument ("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.doc");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WordException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
			
		// Save the document as a PDF file
		try {
			wd.saveAsPDF("c:/PerforceWorkspaces/jlucuik_JLUCUIK-XP_5324/jlucuik_JLUCUIK-XP_5324/pdf/test.word.pdf");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (PDFException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
