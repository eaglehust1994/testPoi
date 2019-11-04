package testTemplate.PoiJava;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
/**
 * Hello world!
 *
 */
@SuppressWarnings("unused")
public class App 
{
    public static void main( String[] args ) throws IOException
    {
    	   Date date = new Date();
   	    Long da = date.getTime();
    	   XWPFDocument document = new XWPFDocument();
    	    // Create new Paragraph
    	    XWPFParagraph appendix = document.createParagraph();
    	    appendix.setAlignment(ParagraphAlignment.CENTER);
    	    genFontTextTitle(appendix,"Phụ lục",14,"Times New Roman",true);
    	    XWPFParagraph titleProcPackage = document.createParagraph();
    	    titleProcPackage.setAlignment(ParagraphAlignment.CENTER);
    	    genFontTextTitle(titleProcPackage, "Danh mục hàng hóa thuộc gói thầu số ", 14, "Times New Roman", true);
    	    
    	    // Write the Document in file system
    	    FileOutputStream out = new FileOutputStream(new File("demo_POI" + da +".docx"));
    	    XWPFTable table = document.createTable();
    	    	String font  = "Times New Roman";
    	      //create  row title 
    	      XWPFTableRow tableRowTitle = table.getRow(0);
    	      XWPFParagraph indexTitle = tableRowTitle.getCell(0).addParagraph();
    	      indexTitle.setAlignment(ParagraphAlignment.CENTER);
    	      genFontTextTitle(indexTitle,"STT",12,font,true);
    	      
    	      XWPFParagraph itemTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(itemTitle,"Tên hàng hóa",12,font,true);
    	      
    	      XWPFParagraph countTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(countTitle,"Số lượng",12,font,true);
    	      
    	      XWPFParagraph unitPriceTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(unitPriceTitle,"Đơn giá (trước VAT)",12,font,true);
    	      
    	      XWPFParagraph amountBeforeVatTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(amountBeforeVatTitle,"Thành tiền (trước VAT)",12,font,true);
    	      
    	      XWPFParagraph vatTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(vatTitle,"VAT",12,font,true);
    	      
    	      XWPFParagraph totalPriceTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(totalPriceTitle,"Thành tiền (Sau VAT)",12,font,true);
    	      
    	      XWPFParagraph currencyNameTitle = tableRowTitle.addNewTableCell().addParagraph();
    	      genFontTextTitle(currencyNameTitle,"Loại tiền",12,font,true); 
    	      
    	      
    	    document.write(out);
    	    out.close();
    	    document.close();
    	    System.out.println("successully");
    }
    
    private static void genFontTextTitle (XWPFParagraph genText, String text, Integer fontSize, String fontFamily,Boolean bold) {
    	XWPFRun run = genText.createRun();
    	run.setText(text);
    	run.setFontSize(fontSize);
    	run.setFontFamily(fontFamily);
    	run.setBold(bold);
    	
    }
    
    private static void genFontDataTable () {
    	
    }
}
