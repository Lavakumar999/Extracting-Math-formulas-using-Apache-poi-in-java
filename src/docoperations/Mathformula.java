/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package docoperations;

import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;

/**
 *
 * @author Lava Kumar
 */
public class Mathformula {
    static long startTime = System.currentTimeMillis();
    static String val="";
    static File stylesheet = new File("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\OMML2MML.XSL");
    static File stylesheet1=new File("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\MML2OMML.XSL");
    static TransformerFactory tFactory = TransformerFactory.newInstance();
    static StreamSource stylesource = new StreamSource(stylesheet);
    static StreamSource stylesource1 = new StreamSource(stylesheet1);
    static String getMathML(CTOMath ctomath) throws Exception {
            Transformer transformer = tFactory.newTransformer(stylesource);
            
                DOMSource source = new DOMSource(ctomath.getDomNode());
                StringWriter stringwriter = new StringWriter();
                StreamResult result = new StreamResult(stringwriter);
                transformer.setOutputProperty("omit-xml-declaration", "yes");
                transformer.transform(source, result);
                String mathML = stringwriter.toString();
                stringwriter.close();
                   mathML = mathML.replaceAll("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
                   mathML = mathML.replaceAll("xmlns:mml", "xmlns");
                   mathML = mathML.replaceAll("mml:", "");
         return mathML;
    }
//    static CTOMath getOMML(String mathML) throws Exception {
//            Transformer transformer = tFactory.newTransformer(stylesource1);
//
//            StringReader stringreader = new StringReader(mathML);
//            StreamSource source = new StreamSource(stringreader);
//
//            StringWriter stringwriter = new StringWriter();
//            StreamResult result = new StreamResult(stringwriter);
//            transformer.transform(source, result);
//
//            String ooML = stringwriter.toString();
//            stringwriter.close();
//            System.out.println(ooML);
//            CTOMathPara ctOMathPara = CTOMathPara.Factory.parse(ooML);
//            CTOMath ctOMath = ctOMathPara.getOMathArray(1000);
//            //for making this to work with Office 2007 Word also, special font settings are necessary
//            XmlCursor xmlcursor = ctOMath.newCursor();
//            while (xmlcursor.hasNextToken()) {
//             XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
//             if (tokentype.isStart()) {
//              if (xmlcursor.getObject() instanceof CTR) {
//               CTR cTR = (CTR)xmlcursor.getObject();
//               cTR.addNewRPr2().addNewRFonts().setAscii("Cambria Math");
//               cTR.getRPr2().getRFonts().setHAnsi("Cambria Math");
//              }
//             }
//            }
//        return ctOMath;
//    }

    public static void main(String[] args) {
         try {
            FileInputStream fis=new FileInputStream("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\Docdemo.docx");
            XWPFDocument document=new XWPFDocument(fis);
            List<String> mathMLList = new ArrayList<String>();
       for (IBodyElement ibodyelement : document.getBodyElements()) {
            if (ibodyelement.getElementType().equals(BodyElementType.PARAGRAPH)) {
               XWPFParagraph paragraph = (XWPFParagraph)ibodyelement;
               for (CTOMath ctomath : paragraph.getCTP().getOMathList()) {
                     mathMLList.add(getMathML(ctomath));
               }
               for (CTOMathPara ctomathpara : paragraph.getCTP().getOMathParaList()) {
                       for (CTOMath ctomath : ctomathpara.getOMathList()) {
                       mathMLList.add(getMathML(ctomath));
                       }
              }
     }
     else if (ibodyelement.getElementType().equals(BodyElementType.TABLE)) {
              XWPFTable table = (XWPFTable)ibodyelement; 
              for (XWPFTableRow row : table.getRows()) {
                  for (XWPFTableCell cell : row.getTableCells()) {
                      for (XWPFParagraph paragraph : cell.getParagraphs()) {
                           for (CTOMath ctomath : paragraph.getCTP().getOMathList()) {
                               mathMLList.add(getMathML(ctomath));
                           }
                           for (CTOMathPara ctomathpara : paragraph.getCTP().getOMathParaList()) {
                                 for (CTOMath ctomath : ctomathpara.getOMathList()) {
                                        mathMLList.add(getMathML(ctomath));
                                 }
                           }
                      }
                  }
              }
           }
         }
       int itr=1;
          for (String mathML : mathMLList) {
//              CTOMath ctOMath = getOMML(mathML);
//              System.out.println(ctOMath);
              String target = mathML.replaceAll("<[^>]*>", "");
              //System.out.println("formula "+mathML);
              System.out.println("formula"+" "+itr+" "+target);
                  itr++;    
          }  
          String encoding = "UTF-8";
            FileOutputStream fos = new FileOutputStream("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\result.html");
            OutputStreamWriter writer = new OutputStreamWriter(fos, encoding);
            writer.write("<!DOCTYPE html>\n");
            writer.write("<html lang=\"en\">");
            writer.write("<head>");
            writer.write("<meta charset=\"utf-8\"/>");

            //using MathJax for helping all browsers to interpret MathML
            writer.write("<script type=\"text/javascript\"");
            writer.write(" async src=\"https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.1/MathJax.js?config=MML_HTMLorMML\"");
            writer.write(">");
            writer.write("</script>");

            writer.write("</head>");
            writer.write("<body>");
            writer.write("<p>Following formulas was found in Word document: </p>");

            int i = 1;
            for (String mathML : mathMLList) {
             writer.write("<p>Formula" + i++ + ":</p>");
             writer.write(mathML);
             writer.write("<p/>");
            }
            writer.write("<input type='button' value='print' onclick='window.print();'/>");
            writer.write("</body>");
            writer.write("</html>");
            writer.close();

            Desktop.getDesktop().browse(new File("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\result.html").toURI());
//              FileReader fr=new FileReader("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\result.html");
//              BufferedReader br= new BufferedReader(fr);
//             StringBuilder content=new StringBuilder(1024);
//             String s="";
//             while((s=br.readLine())!=null)
//                 {
//                 content.append(s);
//                 }
//              String k=content.toString();
//              OutputStream file = new FileOutputStream(new File("C:\\Users\\Lava Kumar\\Desktop\\web programs\\Learn\\Test.pdf"));
//              com.itextpdf.text.Document doc = new com.itextpdf.text.Document();
//              PdfWriter pdfwriter = PdfWriter.getInstance(doc, file);
//              doc.open();
//              InputStream is = new ByteArrayInputStream(k.getBytes());
//              XMLWorkerHelper worker = XMLWorkerHelper.getInstance();
//              worker.parseXHtml(pdfwriter, doc, new StringReader(k));
//              doc.close();
             
//                ITextRenderer renderer = new ITextRenderer();
//                renderer.setDocument(k);
//                renderer.layout();
//                renderer.createPDF(file);
//                file.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocOperations.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocOperations.class.getName()).log(Level.SEVERE, null, ex);
        } catch (Exception ex) {
            Logger.getLogger(Mathformula.class.getName()).log(Level.SEVERE, null, ex);
        }
      long endTime   = System.currentTimeMillis();
      long totalTime = endTime - startTime;
      System.out.println("total time ="+totalTime+"  Milliseconds");  
        
    }
}
