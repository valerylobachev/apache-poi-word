/**
 * Created by valery on 05.03.17.
 */
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class HWPFTest {
    public static void main(String[] args){
        String filePath = "template.docx";
        String outFilePath = "result.docx";
        POIFSFileSystem fs = null;
        try {

            OPCPackage pack = POIXMLDocument.openPackage(filePath);
            XWPFDocument doc = new XWPFDocument(pack);
            List<XWPFParagraph> paragraphList = doc.getParagraphs();
            processParagraphs(paragraphList, "<<companyName>>", "OXY Consult");
            FileOutputStream fos = new FileOutputStream(outFilePath);
            doc.write(fos);
            fos.flush();
            fos.close();
//            fs = new POIFSFileSystem(new FileInputStream(filePath));
//            HWPFDocument doc = new HWPFDocument(fs);
//            //doc = replaceText(doc, "<<companyName>>", "OXY Consult");
//            saveWord(outFilePath, doc);
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }


    private static void processParagraphs(List<XWPFParagraph> paragraphList,
                                   String key, String replaceText) {
        for (XWPFParagraph paragraph : paragraphList) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                boolean isSetText = false;
                if (text != null && text.indexOf(key) != -1) {
                    isSetText = true;
                    text = text.replace(key, replaceText);
                }
                if (isSetText) {
                    run.setText(text, 0);
                }
            }
        }
    }
    private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
        Range r1 = doc.getRange();
        r1.replaceText(findText, replaceText);

        /*for (int i = 0; i < r1.numSections(); ++i ) {
            Section s = r1.getSection(i);
            for (int x = 0; x < s.numParagraphs(); x++) {
                Paragraph p = s.getParagraph(x);
                for (int z = 0; z < p.numCharacterRuns(); z++) {
                    CharacterRun run = p.getCharacterRun(z);
                    String text = run.text();
                    if(text.contains(findText)) {
                        run.replaceText(findText, replaceText);
                    }
                }
            }
        }*/
        return doc;
    }

    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            out.close();
        }
    }
}