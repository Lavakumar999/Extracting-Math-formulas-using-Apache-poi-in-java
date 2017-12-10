package docoperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class DocOperations {

    public static void main(String[] args) {
        
        try {
            
            FileInputStream fis=new FileInputStream("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\Docdemo.docx");
            //this class is used to extract the content
            XWPFDocument docx=new XWPFDocument(fis);
            List<XWPFParagraph> paragraphList=docx.getParagraphs();
            for(XWPFParagraph paragraph:paragraphList){
                System.out.println(paragraph.getParagraphText());
            }
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocOperations.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocOperations.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
}
