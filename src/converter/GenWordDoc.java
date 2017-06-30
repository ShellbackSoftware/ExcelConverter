/* 
Copyright (c) [2017] 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */
package converter;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * This class just handles creating & editing the word document. The lists are what is written
 * out to the Word file, and the list being used changes depending on some user input.
 * 
 * @author shellbacksoftware@gmail.com
 */
public class GenWordDoc {
    private String outputPath;          // Path to the document
    private List<Question> questions;   // List of questions
    private List<Presenter> presenters; // List of presenters
    private Boolean hasPresenter;       // If true, follows a certain template
    
    /* Constructor. Initializations throw exceptions, so it looks messier than it is */
    public GenWordDoc(List que, String path, String name, Boolean presenter){
        hasPresenter = presenter;
        try {
             if(!(path.equals(""))){
                outputPath = path.replace("\\", "/");
                outputPath = outputPath+"/"+name;
            }else{
                outputPath = name;
            }
            if(!hasPresenter){
                questions = que;
            }else{
                presenters = que;
            }
            createDoc();
        } catch (IOException ex) {
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /* Creates the document */
    private void createDoc() throws IOException{
        FileOutputStream out = null;
        try {
            XWPFDocument doc = new XWPFDocument();
            out = new FileOutputStream(outputPath);
            if(!hasPresenter){
                for(Question q : questions){
                    XWPFParagraph par = doc.createParagraph();
                    XWPFRun r1 = par.createRun();
                    r1.setText(q.toString());                   // Currently, question's toString
                    XWPFParagraph c = doc.createParagraph();
                    XWPFRun r2 = c.createRun();
                }
            }else{
                for(Presenter p : presenters){
                    XWPFParagraph par = doc.createParagraph();
                    XWPFRun r1 = par.createRun();
                    r1.setText(p.toString());                   // Currently, question's toString
                    XWPFParagraph c = doc.createParagraph();
                    XWPFRun r2 = c.createRun();
                    XWPFParagraph br = doc.createParagraph();
                    br.setPageBreak(true);
                }
            }
        doc.write(out);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
        out.close();
    }
}
