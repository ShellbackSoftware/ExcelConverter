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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
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
    private String templatePath;        // Path to template file
    private List<Question> questions;   // List of questions
    private List<Presenter> presenters; // List of presenters
    
    private boolean hasPresenter;
    private boolean wantsCounts;
    private boolean wantsComments;
    
    
    /* Constructor. Initializations throw exceptions, so it looks messier than it is */
    public GenWordDoc(List que, String path, String name, Boolean presenter, boolean counts, 
            boolean comments, String tPath){
        hasPresenter = presenter;
        wantsCounts = counts;
        wantsComments = comments;
        templatePath = tPath;
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
            boolean testing = false;
            if(testing) createDoc();
            else{
                XWPFDocument templateDoc = new XWPFDocument(OPCPackage.openOrCreate(new File(templatePath)));
                XWPFWordExtractor extractor = new XWPFWordExtractor(templateDoc);
                String template = extractor.getText();
                generatePages(template);
            }
        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /* Creates the necessary amount of pages, copying the template to each */
    private void generatePages(String temp) throws IOException{
        FileOutputStream out = null;
        try {
            XWPFDocument wdoc = new XWPFDocument();
            out = new FileOutputStream(outputPath);
            if(!hasPresenter){
                for(Question q : questions){
                    XWPFParagraph par = wdoc.createParagraph();
                    XWPFRun r1 = par.createRun();
                    r1.setText(temp);
                    String newInfo = temp;
                    XWPFParagraph c = wdoc.createParagraph();
                    XWPFRun r2 = c.createRun();
                }
            }else{
                for(Presenter p : presenters){
                    XWPFParagraph par = wdoc.createParagraph();
                    XWPFRun r1 = par.createRun();
                    String newInfo = temp;
                    // Names
                    StringBuilder names = new StringBuilder();
                    for (String s : p.getQuestionNames()){
                        names.append(s);
                        names.append("\n");
                    }
                    XWPFParagraph c = wdoc.createParagraph();
                    // Ugly, but it works - gets required information to replace stuff
                    newInfo = replacePresenter(newInfo,p.getName());
                    newInfo = replaceQName(newInfo,names.toString());
                    ArrayList<String> avgScores = new ArrayList<>();
                    ArrayList<String> favScores = new ArrayList<>();
                        for(Question q : p.getQuestions()){
                            String qAvgScore = String.format("%.1f",q.getScore());
                            String qFavScore = String.format("%.0f%%",q.getPercent());
                            avgScores.add(qAvgScore);
                            favScores.add(qFavScore);
                        }
                    // Average scores
                    StringBuilder aScores = new StringBuilder();
                    for (String s : avgScores){
                        aScores.append(s);
                        aScores.append("\n");
                    }
                    // Favorable scores
                    StringBuilder fScores = new StringBuilder();
                    for (String s : favScores){
                        fScores.append(s);
                        fScores.append("\n");
                    }
                    newInfo = replaceScores(newInfo, aScores.toString(), fScores.toString());   
                    String qCount = "Counts go here";
                    ArrayList<String> qComments = p.getComments();
                    r1.setText(newInfo);
                    XWPFParagraph br = wdoc.createParagraph();
                    br.setPageBreak(true);
                }
            }
        wdoc.write(out);
        }catch (IOException ex) {
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
        out.close();
    }
    
    /* Replaces the presenter name in the template file */
    private String replacePresenter(String targetString, String newText){
       return targetString.replace("%PRESENTER%", newText);
    }
    
    /* Replaces question names */
    private String replaceQName(String targetString, String newText){
        return targetString.replace("%QUESTION%", newText);
    }
    
    /* Replaces scores */
    private String replaceScores(String targetString, String avgScore, String favScore){
        String temp = targetString.replace("%FAVSCORE%", favScore);
        return temp.replace("%AVGSCORE%", avgScore);
    }
    
    /* Replaces the counts */
    private String replaceCounts(String targetString, String newText){
        return targetString.replace("%COUNTS%", newText);
    }
    
    /* Replaces the comments */
    private String replaceComments(String targetString, String newText){
        return targetString.replace("%COMMENTS%", newText);
    }
    
    /* Creates the document, used for testing. Will be deleted for distribution */
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
