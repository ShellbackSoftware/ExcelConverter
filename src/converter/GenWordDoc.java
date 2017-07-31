package converter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

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
    
    private ConverterGUI gui = new ConverterGUI();
    
    
    /* Constructor. Initializations throw exceptions, so it looks messier than it is */
    public GenWordDoc(List que, String path, String name, Boolean presenter, boolean counts, 
            boolean comments, String tPath){
        hasPresenter = presenter;
        wantsCounts = counts;
        wantsComments = comments;
        templatePath = tPath;
        try {
            if(name.equals(".docx")){
            throw new EmptyFileNameException("The file name can't be blank!");
            }   
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
        } catch (IOException ex ) {
            gui.printError("Output File Path Error", "The output file path is incorrect. Please verify that the location exists.");
        } catch( InvalidFormatException ex){
            gui.printError("InvalidFormatException", "Invalid format");
        } catch(POIXMLException ex){
            gui.printError("Template Path error",
                    "The template file path appears to be incorrect. Please verify that the path is correct, and that the file exists.");
        } catch(EmptyFileNameException ex){
            gui.printError("Output File Name Error", "Please enter a file name for the output file.");
        }        
    }
    
    /* Does the work of replacing the text in the duplicated template, and printing out the result */
    private void runProgram(XWPFDocument doc){
        FileOutputStream out = null;
        try{
            XWPFDocument newdoc = new XWPFDocument();
            // Sets page margins; 720L = 1/2"
            CTSectPr sectPr = newdoc.getDocument().getBody().addNewSectPr();
            CTPageMar pageMar = sectPr.addNewPgMar();
            pageMar.setLeft(BigInteger.valueOf(1080L));
            //pageMar.setTop(BigInteger.valueOf(360L));
            pageMar.setRight(BigInteger.valueOf(1080L));
            //pageMar.setBottom(BigInteger.valueOf(360L));
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
                replaceText(p.getName(),"%DELIMITER",newdoc);
                
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
                    replaceComments(p.getComments(),"%COMMENTS",newdoc);
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
        }catch (IOException ex ) {
            gui.printError("Output File Path Error", "The output file path is incorrect. Please verify that the location exists.");
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

    
    /* Currently, comments are a table with no border */
    private void replaceComments(ArrayList<String> comments, String find, XWPFDocument doc) {
        XWPFTable table = null;
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            TextSegement found = paragraph.searchText(find, new PositionInParagraph());
            if ( found != null ) {
                if ( found.getBeginRun() == found.getEndRun() ) {
                 // whole search string is in one Run
                 XmlCursor cursor = paragraph.getCTP().newCursor();
                 table = doc.insertNewTbl(cursor);

                 CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
                 cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));

                 // Bullet list
                CTLvl cTLvl = cTAbstractNum.addNewLvl();
                cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
                cTLvl.addNewLvlText().setVal("•");
                
                XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);

                XWPFNumbering numbering = doc.createNumbering();

                BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);

                BigInteger numID = numbering.addNum(abstractNumID);

                for (String string : comments) {
                    XWPFTableRow lnewRow = table.createRow();
                    XWPFTableCell lnewCell = lnewRow.getCell(0);
                    XWPFParagraph lnewPara =lnewCell.getParagraphs().get(0);
                    lnewPara.setNumID(numID);
                    XWPFRun lnewRun = lnewPara.createRun();
                    lnewRun.setText(string); 
                }
                 
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
                 int pad = (int) (.1 * 1440);
                 table.setCellMargins(0, pad, 0, pad);  // Top, left, bottom, right
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
        formatComments(table);
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
                 int pad = (int) (.1 * 1440);
                 table.setCellMargins(0, pad, 0, pad); // Top, left, bottom, right
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
                 int pad = (int) (.1 * 1440);
                 table.setCellMargins(0, pad, 0, pad); // Top, left, bottom, right
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
                 int pad = (int) (.1 * 1440);
                 table.setCellMargins(0, pad, 0, pad);  // Top, left, bottom, right
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
    private void replaceText(String repl, String targ, XWPFDocument doc) {
    for (XWPFParagraph paragraph : doc.getParagraphs()) {
      List<XWPFRun> runs = paragraph.getRuns();
        String find = targ;
        TextSegement found = paragraph.searchText(find, new PositionInParagraph());
        if ( found != null ) {
          if ( found.getBeginRun() == found.getEndRun() ) {
            // whole search string is in one Run
            XWPFRun run = runs.get(found.getBeginRun());
            String runText = run.getText(run.getTextPosition());
            String replaced = runText.replace(find, repl);
            run.setText(replaced, 0);
          } else {
            // The search string spans over more than one Run
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
  }  

    /* Fills counts table */
    private void fillCountsTable(XWPFTable table, ArrayList<String> scores, ArrayList<String> values){
        int currRow = 0;
        scores.add(0,"Response");
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
        values.add(0,"Votes");
        for(String v : values){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            curRow.getCell(1).setText(v);
            currRow++;    
        }
        // Format the table
        formatTable(table,true);
    }
    
    /* Fills the questions table */
    private void fillTable(XWPFTable table, ArrayList<String> qs, ArrayList<String> avgScores, ArrayList<String> favScores) {
        int currRow = 0;
        qs.add(0,"Question");
        for(String q : qs){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.getCell(0).setText(q);
            if(currRow < qs.size()-1){
                table.createRow();
                currRow++;
            }else{
                currRow++;
            }
        }
        currRow = 0;
        avgScores.add(0,"Average Score");
        for(String avg : avgScores){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            curRow.getCell(1).setText(avg);
            currRow++;    
        }
        currRow = 0;
        favScores.add(0,"Favorable Percent");
        for(String fav : favScores){
            XWPFTableRow curRow = table.getRow(currRow);
            curRow.addNewTableCell();
            curRow.getCell(2).setText(fav);
            currRow++;
        }
        // Format the table
        formatTable(table,false);
    }
    
    /* Formats the comments */
    private void formatComments(XWPFTable table){
        table.setWidth(5000);
        table.removeRow(0);     // Deletes empty first row
        table.getCTTbl().getTblPr().unsetTblBorders();
    }
    
    // Formats the table to make it purdy.
    private void formatTable(XWPFTable table, boolean counts){
    List<XWPFTableRow> rows = table.getRows();          // List of rows in table
    int rowCt = 0;
    int colCt = 0;
    for (XWPFTableRow row : rows) {
        row.setHeight(360);
        row.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT); //set w:hRule="exact"
        List<XWPFTableCell> cells = row.getTableCells();        // Cells in this row
            for (XWPFTableCell cell : cells) {
                CTTcPr tcpr = cell.getCTTc().addNewTcPr();      // get a table cell properties element (tcPr)
                CTVerticalJc va = tcpr.addNewVAlign();  
                va.setVal(STVerticalJc.CENTER);                 // Center vert align

                XWPFParagraph para = cell.getParagraphs().get(0);   // First paragraph in cell
                if (rowCt == 0) {          
                   para.setAlignment(ParagraphAlignment.CENTER);
                }else if (colCt == 0 && !counts){
                   para.setAlignment(ParagraphAlignment.LEFT);
                }else{
                   para.setAlignment(ParagraphAlignment.RIGHT);
                }
                 cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2000));
                 colCt++;
            } // End cell	
            rowCt++;
            colCt=0;
        } // End row
    }
}