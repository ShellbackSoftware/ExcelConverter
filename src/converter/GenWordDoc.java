package converter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

/**
 * This class just handles creating & editing the word document. It takes in the template file,
 * then makes a copy of it and replaces the text in the copy.
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
            XWPFDocument templateDoc = new XWPFDocument(OPCPackage.openOrCreate(new File(templatePath)));
            runProgram(templateDoc);
        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /* Does the work of replacing the text in the duplicated template, and printing out the result */
    private void runProgram(XWPFDocument doc){
        FileOutputStream out = null;
        try{
            XWPFDocument newdoc = new XWPFDocument();
            out = new FileOutputStream(outputPath);
            if(hasPresenter){
                for(Presenter p : presenters){
                    List<XWPFParagraph> paras = doc.getParagraphs();
                    for (XWPFParagraph para : paras) {

                        if (!para.getParagraphText().isEmpty()) {
                            XWPFParagraph newpara = newdoc.createParagraph();
                            generatePages(para, newpara);
                        }
                    }
                // Presenter
                replaceText(p.getName(), null, "%DELIMITER",newdoc);
                
                // Question Scores
                ArrayList<String> avgScores = new ArrayList<>();
                ArrayList<String> favScores = new ArrayList<>();
                
                // Used for counts
                ArrayList<String> countKeys = new ArrayList<>();
                ArrayList<String> countVals = new ArrayList<>();
                
                    for(Question q : p.getQuestions()){
                        String qAvgScore = String.format("%.1f",q.getScore());
                        String qFavScore = String.format("%.0f%%",q.getPercent());
                        avgScores.add(qAvgScore);
                        favScores.add(qFavScore);
                        if(wantsCounts){
                            Map<Double,Integer> counts = q.getCounts();
                            for (Map.Entry<Double, Integer> entry : counts.entrySet()) {
                                Double key = entry.getKey();
                                Integer value = entry.getValue();
                                countKeys.add(key.toString());
                                countVals.add(value.toString());
                            }
                        }
                    }
                replaceTable(newdoc, p.getQuestionNames(), avgScores, favScores,"%INFO"); 
                
                if(wantsCounts){
                    replaceCountsTable(newdoc, countKeys, countVals,"%COUNTS");
                }

                if(wantsComments){
                    replaceText(null, p.getComments(),"%COMMENTS",newdoc);
                }
                
                XWPFParagraph br = newdoc.createParagraph();
                br.setPageBreak(true);
                }
                newdoc.write(out);
            }else{
                List<XWPFParagraph> paras = doc.getParagraphs();
                    for (XWPFParagraph para : paras) {

                        if (!para.getParagraphText().isEmpty()) {
                            XWPFParagraph newpara = newdoc.createParagraph();
                            generatePages(para, newpara);
                        }
                    }
                ArrayList<String> qNames = new ArrayList<>();
                for(Question q: questions){
                    // Questions
                    qNames.add(q.getQuestion());
                }
                    // Question Scores
                    ArrayList<String> avgScores = new ArrayList<>();
                    ArrayList<String> favScores = new ArrayList<>();
                    
                    // Used for counts
                    ArrayList<String> countKeys = new ArrayList<>();
                    ArrayList<String> countVals = new ArrayList<>();
                for (Question q : questions){
                    String qAvgScore = String.format("%.1f",q.getScore());
                    String qFavScore = String.format("%.0f%%",q.getPercent());
                    avgScores.add(qAvgScore);
                    favScores.add(qFavScore);
                    if(wantsCounts){
                        Map<Double,Integer> counts = q.getCounts();
                        for (Map.Entry<Double, Integer> entry : counts.entrySet()) {
                            Double key = entry.getKey();
                            Integer value = entry.getValue();
                            countKeys.add(key.toString());
                            countVals.add(value.toString());
                        }
                    }
                }
                replaceTable(newdoc, qNames, avgScores, favScores,"%INFO"); 
                
                if(wantsCounts){
                    replaceCountsTable(newdoc, countKeys, countVals,"%COUNTS");
                }
                newdoc.write(out);
            }
            out.flush();
            out.close();
        }catch(IOException ex){
            Logger.getLogger(GenWordDoc.class.getName()).log(Level.SEVERE, null, ex);
        }
    }   
    
    /* Generates as many pages as required from template */
    private void generatePages(XWPFParagraph oldPar, XWPFParagraph newPar) {
        final int DEFAULT_FONT_SIZE = 10;

        for (XWPFRun run : oldPar.getRuns()) {  
            String textInRun = run.getText(0);
            if (textInRun == null || textInRun.isEmpty()) {
                continue;
            }
            int fontSize = run.getFontSize();

            XWPFRun newRun = newPar.createRun();

            // Copy text
            newRun.setText(textInRun);

            // Apply the same style
            newRun.setFontSize( ( fontSize == -1) ? DEFAULT_FONT_SIZE : run.getFontSize() );    
            newRun.setFontFamily( run.getFontFamily() );
            newRun.setBold( run.isBold() );
            newRun.setItalic( run.isItalic() );
            newRun.setStrike( run.isStrike() );
            newRun.setColor( run.getColor() );
        }
    }

    /* Replaces Counts table */
    private void replaceCountsTable(XWPFDocument doc, ArrayList<String> keys, ArrayList<String> values, String find) {
        XWPFTable table = null;
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            TextSegement found = paragraph.searchText(find, new PositionInParagraph());
            if ( found != null ) {
                if ( found.getBeginRun() == found.getEndRun() ) {
                 // whole search string is in one Run
                 XmlCursor cursor = paragraph.getCTP().newCursor();
                 table = doc.insertNewTbl(cursor);
                 XWPFRun run = runs.get(found.getBeginRun());
                 // Clear the keyword from doc
                String runText = run.getText(run.getTextPosition());
                String replaced = runText.replace(find, "");
                run.setText(replaced, 0);
                } else {
                 // The search string spans over more than one Run
                 StringBuilder b = new StringBuilder();
                 for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
                   XWPFRun run = runs.get(runPos);
                   b.append(run.getText(run.getTextPosition()));
                }                       
                 String connectedRuns = b.toString();
                 XmlCursor cursor = paragraph.getCTP().newCursor();
                 table = doc.insertNewTbl(cursor);
                 String replaced = connectedRuns.replace(find, ""); // Clear search text

                 // The first Run receives the replaced String of all connected Runs
                 XWPFRun partOne = runs.get(found.getBeginRun());
                 partOne.setText(replaced, 0);
                 // Removing the text in the other Runs.
                 for (int runPos = found.getBeginRun()+1; runPos <= found.getEndRun(); runPos++) {
                   XWPFRun partNext = runs.get(runPos);
                   partNext.setText("", 0);
                 }
               }
            }     
        }
        fillCountsTable(table, keys, values);
    }
    
    /* Replaces questions table */
    private void replaceTable(XWPFDocument doc, ArrayList<String> qs, ArrayList<String> avgScores, ArrayList<String> favScores, String find) {
        XWPFTable table = null;
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            TextSegement found = paragraph.searchText(find, new PositionInParagraph());
            if ( found != null ) {
                if ( found.getBeginRun() == found.getEndRun() ) {
                 // whole search string is in one Run
                 XmlCursor cursor = paragraph.getCTP().newCursor();
                 table = doc.insertNewTbl(cursor);
                 XWPFRun run = runs.get(found.getBeginRun());
                 // Clear the keyword from doc
                String runText = run.getText(run.getTextPosition());
                String replaced = runText.replace(find, "");
                run.setText(replaced, 0);
                } else {
                 // The search string spans over more than one Run
                 StringBuilder b = new StringBuilder();
                 for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
                   XWPFRun run = runs.get(runPos);
                   b.append(run.getText(run.getTextPosition()));
                }                       
                 String connectedRuns = b.toString();
                 XmlCursor cursor = paragraph.getCTP().newCursor();
                 table = doc.insertNewTbl(cursor);
                 String replaced = connectedRuns.replace(find, ""); // Clear search text

                 // The first Run receives the replaced String of all connected Runs
                 XWPFRun partOne = runs.get(found.getBeginRun());
                 partOne.setText(replaced, 0);
                 // Removing the text in the other Runs.
                 for (int runPos = found.getBeginRun()+1; runPos <= found.getEndRun(); runPos++) {
                   XWPFRun partNext = runs.get(runPos);
                   partNext.setText("", 0);
                 }
               }
            }     
        }
        fillTable(table, qs, avgScores, favScores);
    }
   
    
    /* Replace text in the document */
    private long replaceText(String rep, ArrayList<String> comments, String targ, XWPFDocument doc) {
    long count = 0;
    String repl = "";
    for (XWPFParagraph paragraph : doc.getParagraphs()) {
      List<XWPFRun> runs = paragraph.getRuns();

        StringBuilder sb = new StringBuilder();
        if(comments != null){
            for (String c : comments){
                sb.append(c);
                sb.append("|");
            }
            repl = sb.toString();
        }else{
            repl = rep;
        }
        String find = targ;
        TextSegement found = paragraph.searchText(find, new PositionInParagraph());
        if ( found != null ) {
          count++;
          if ( found.getBeginRun() == found.getEndRun() ) {
            // whole search string is in one Run
            XWPFRun run = runs.get(found.getBeginRun());
            String runText = run.getText(run.getTextPosition());
            String replaced = runText.replace(find, repl);
            run.setText(replaced, 0);
          } else {
            // The search string spans over more than one Run
            // Put the Strings together
            StringBuilder b = new StringBuilder();
            for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
              XWPFRun run = runs.get(runPos);
              b.append(run.getText(run.getTextPosition()));
            }                       
            String connectedRuns = b.toString();
            String replaced = connectedRuns.replace(find, repl);

            // The first Run receives the replaced String of all connected Runs
            XWPFRun partOne = runs.get(found.getBeginRun());
            partOne.setText(replaced, 0);
            // Removing the text in the other Runs.
            for (int runPos = found.getBeginRun()+1; runPos <= found.getEndRun(); runPos++) {
              XWPFRun partNext = runs.get(runPos);
              partNext.setText("", 0);
            }                          
          }
        }     
    }
    return count;
  }  

    /* Fills counts table */
    private void fillCountsTable(XWPFTable table, ArrayList<String> scores, ArrayList<String> values){
        int currRow = 0;
        scores.add(0," Score ");
        for(String q : scores){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.getCell(0).setText(q);
            if(currRow < scores.size()-1){
                table.createRow();
                currRow++;
            }else{
                currRow++;
            }
        }
        currRow = 0;
        values.add(0," Amount ");
        for(String v : values){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            curRow.getCell(1).setText(v);
            currRow++;    
        }
    }
    
    /* Fills the questions table */
    private void fillTable(XWPFTable table, ArrayList<String> qs, ArrayList<String> avgScores, ArrayList<String> favScores) {
        int currRow = 0;
        for(String q : qs){
            XWPFTableRow curRow = table.getRow(currRow);
            if(currRow == 0){
                curRow.getCell(0).setText(" Question ");
            }else{
            curRow.getCell(0).setText(q);
            }
            if(currRow < qs.size()-1){
                table.createRow();
                currRow++;
            }else{
                currRow++;
            }
        }
        currRow = 0;
        for(String avg : avgScores){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            if(currRow == 0){
                curRow.getCell(1).setText(" Average Score ");
            }else{
                curRow.getCell(1).setText(avg);
            }
            currRow++;    
        }
        currRow = 0;
        for(String fav : favScores){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            if(currRow == 0){
                curRow.getCell(2).setText(" Favorable Score ");
            }else{
                curRow.getCell(2).setText(fav);
            }
            currRow++;
        }
    }
}