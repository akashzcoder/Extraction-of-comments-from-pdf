/* * * * * * * * *
 *     Zcoder    *
 * * * * * * * * */

//Please clone/fork the code before enhancing or using it. Thanks :)


import com.itextpdf.text.pdf.PdfArray;
import com.itextpdf.text.pdf.PdfDictionary;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfObject;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfString;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow; 
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.ListIterator;

public class CommentsTest {

public static void main(String args[]) throws IOException {
	String outputFileName= "c:\\temp\\output.xls";
	 ArrayList<PdfString> comment = new ArrayList<PdfString>();
	 HSSFWorkbook myWorkBook_m = new HSSFWorkbook();
     HSSFSheet mySheet_m = myWorkBook_m.createSheet("Comments");
     HSSFRow myRow_m = null;
     HSSFCell myCell_m = null;
try {

//System.out.println("Test for this block");
PdfReader reader = new PdfReader("c:\\temp\\2.pdf");// input file

for(int i = 1; i <= reader.getNumberOfPages(); i++)
{

PdfDictionary page = reader.getPageN(i);
PdfArray annotsArray = null;

if(page.getAsArray(PdfName.ANNOTS)==null)
continue;

annotsArray = page.getAsArray(PdfName.ANNOTS);
for (ListIterator iter = annotsArray.listIterator(); iter.hasNext();)
{
PdfDictionary annot = (PdfDictionary) PdfReader.getPdfObject((PdfObject) iter.next());
PdfString content = (PdfString) PdfReader.getPdfObject(annot.get(PdfName.CONTENTS));
if (content != null) {
//System.out.println(content);
comment.add(content);
}
}
}


} catch (Exception e) {
e.printStackTrace();
}
int k=0;
for (int j=0; j <comment.size();j++){

    myRow_m = mySheet_m.createRow(k);
   myCell_m = myRow_m.createCell(0);
   String cmnt= comment.get(j).toString();
   myCell_m.setCellValue(cmnt);

   k++;
    }
   
    FileOutputStream out_m = new FileOutputStream(outputFileName);
    myWorkBook_m.write(out_m);
    out_m.close();
}
}