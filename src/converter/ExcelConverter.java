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
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Main class of the program. This class reads in the Excel file and deals with user input from
 * the GUI. For the Word file manipulation, see the GenWordDoc class. For more information, see the readme.
 *
 * @author shellbacksoftware@gmail.com
 */
public final class ExcelConverter {

        private DataFormatter dataf;                        // Formats cells to use their data

    
        private boolean hasPresenter = false;               // Set by user if presenters are part of the criteria
        private boolean last;                               // Used to denote the last of the data in the file
        //Following two are unused, will be implemented later
        private boolean multSheets = false;                 // Whether or not there are multiple sheets being used
        private boolean firstSheet = true;                  // Data is on the first sheet in the book or not
        
        private boolean wantsCounts;                        // Whether user wants to know counts per question
        private boolean hasComments;                        // Whether the sheet has comments included
        
        private int totalSheets;                            // Total sheets in the workbook (Unused, to be implemented later)
        private double favScore;                            // User provided cutoff for the favorable score
        
        private ArrayList<Integer> listOfSheets;            // List of sheets the data is on if not on the first sheet (Unused, to be implemented later)
        private ArrayList<Integer> particips;               // Used for the number of participants per chunk of data
        private ArrayList<Integer> endRows;                 // List that holds the values of the rows where each data ends    
        private ArrayList<Double> scores;                   // ArrayList to hold the score values
        
        private ArrayList<String> genComments;              // List of comments for a question if there is no presenter
        private ArrayList<String> countsColumns;            // Columns that the user wants extra info for
        private ArrayList<String> presComments;             // List of comments for a presenter
        
        private Presenter[] presenters;                     // Array of presenters
        
        private Map<String, ArrayList<Question>> qMap;      // Connects lists of questions to each presenter
        private Map<Double, Integer> extraColumns;          // Stores the counts of each value

        
        /* Constructor, initializes everything */
        public ExcelConverter(String path, boolean presenter, boolean multiSheets, boolean fSheet, String sheets, double favoScore, 
                boolean coms, boolean counts, String colCounts, String outName, String outPath, String templatePath) throws IOException{
            firstSheet = fSheet;
            hasPresenter = presenter;
            multSheets = multiSheets;
            favScore = favoScore;
            hasComments = coms;
            wantsCounts = counts;
            countsColumns = new ArrayList<>(Arrays.asList(colCounts.split(",")));
            listOfSheets = new ArrayList<>();
            extraColumns = new HashMap();
            scores = new ArrayList<>();
            genComments = new ArrayList<>();
            particips = new ArrayList<>();
            endRows = new ArrayList<>();
            dataf = new DataFormatter();   
            qMap = new HashMap();
            ArrayList<Question> questions = new ArrayList<>();
            if(hasPresenter){
                readExcel(path);    
                List<Presenter> p = new ArrayList<>(Arrays.asList(presenters));
                p.removeAll(Collections.singleton(null));
                GenWordDoc gwd = new GenWordDoc(p, outPath, outName+".docx", hasPresenter, wantsCounts, hasComments, templatePath);
            }else{
                questions = readExcel(path);
                GenWordDoc gwd = new GenWordDoc(questions, outPath, outName+".docx", hasPresenter, wantsCounts, hasComments, templatePath);
            }
        }
        
        /* Guts of the program. Takes in the file, and does everything needed */
        public ArrayList<Question> readExcel(String filePath) throws IOException{
            ArrayList<Question> qList = new ArrayList<>();                      // List of questions (no presenter)
            FileInputStream input = new FileInputStream(new File(filePath));
            Workbook workbook = getWorkbook(input,filePath);            
            if(multSheets){
                totalSheets = workbook.getNumberOfSheets();
            }
            else{
                totalSheets = 1;
            }
            // Start sheet loop
            for(int s = 0; s < totalSheets; s++){
                Sheet sheet = workbook.getSheetAt(s);               // Loop through each sheet in the workbook
                int numCols = sheet.getRow(0).getLastCellNum();     // Find the number of columns
                int numRows = sheet.getPhysicalNumberOfRows();      // Number of rows
                ArrayList<String> presents = new ArrayList<>();     // Temp list to hold presenters. Unused if no presenters. 
                // Collect presenter names
                if(hasPresenter){
                    presents = getPresenters(sheet);
                    presComments = new ArrayList<>();
                    presenters = new Presenter[presents.size()];
                }
                int numEnds = getBreakAmount(sheet);       
                String[] qNames = new String[numCols];
               // The next 3 arrays are for the values per question
                double[] sums = new double[numEnds];
                double[] favCounts = new double[numEnds];
                int[] counts = new int[numEnds];
                last = false;
                    for(int i = 0; i < numCols; i++){                   // Start columns loop
                        double maxNum = 0;                              // Used later for counts
                        int presenterNum = 0;                           // Which presenter the info belongs to
                        if(last == true){
                            break;
                        }
                        int count = 0;                      
                        double sum = 0;                                 // Sum of the scores
                        double favCount = 0;                            // Favorable scores (4 or 5)
                        Cell cell = null;
                        if(i == 0 && hasPresenter){ i++; }              // Increment to get past the presenter column 
                        for(Row row : sheet){                           // Start Row iterator
                            Question aQuestion;
                            cell = row.getCell(i);
                            String qu = dataf.formatCellValue(cell);
                            if(row.getRowNum() == 0){                   // Skip title row & set question
                               qNames[i] = qu;
                               row = sheet.getRow(1);
                               cell = sheet.getRow(1).getCell(i);
                            }
                            if(rowEmpty(row) || isEmpty(cell)){         // Skip blank rows and cells
                                continue;
                            }
                            String str = dataf.formatCellValue(cell);                                           // Reformat as a string
                            if("end".equalsIgnoreCase(str) && row.getRowNum()+1 == endRows.get(presenterNum)){  // Last value for this question & presenter
                                Presenter pre = new Presenter();
                                counts[presenterNum] = count;
                                sums[presenterNum] = sum;
                                favCounts[presenterNum] = favCount;
                                if(hasPresenter){
                                   pre.setName(presents.get(presenterNum+1));           // If presenter exists, set it
                                }
                                // Following chunk just sets the various fields, and adds the question to an Array List
                                if(!(qNames[i] == null)){
                                    aQuestion = createQuestion(hasPresenter, wantsCounts, qNames[i],(sums[presenterNum]/counts[presenterNum]),
                                            ((favCounts[presenterNum]/counts[presenterNum])*100), (particips.get(presenterNum)) );
                                    if(!hasPresenter){
                                        aQuestion.setComments(genComments);
                                    }
                                    if(wantsCounts && convertToIndex(countsColumns).contains(i)){
                                        for(Double cur : scores){
                                            addExtraInfo(cur,getScores(scores,cur));
                                        }
                                        aQuestion.setCounts(extraColumns);
                                    }
                                    qList.add(aQuestion);
                                    if(hasPresenter){
                                        if(!(hasComments && aQuestion.getQuestion().contains("comment"))){
                                            addQuestion(presents.get(presenterNum+1), aQuestion);
                                        }
                                        if(qNames[i].contains("comment")){
                                            cleanPresComments();
                                            pre.setComments(presComments);
                                        }
                                        if(pre != null){
                                            pre.setQuestions(qMap.get(presents.get(presenterNum+1)));
                                            if(!(qList.isEmpty())){
                                                presenters[presenterNum] = pre;
                                            }
                                        }
                                    }
                                }
                                if(presenterNum < numEnds){
                                    presenterNum++;
                                }
                                if(row.getRowNum() == numRows){ // End of column
                                    last = true; 
                                    if(i == numCols-1){         // End of file
                                        return qList;
                                    }
                                    break;
                                }
                                extraColumns = new HashMap();
                                presComments = new ArrayList<>();
                                genComments = new ArrayList<>();
                                scores = new ArrayList<>();
                                count = 0;
                                sum = 0;
                                favCount = 0;
                                continue;
                            }
                            if(cell != null && cell.getCellType() != 1){
                                if(maxNum < (double) getCellValue(cell)){
                                    maxNum = (double) getCellValue(cell);
                                }
                                sum = sum + (double) getCellValue(cell);
                                scores.add((double) getCellValue(cell));
                                count++;
                                if((double) getCellValue(cell) > favScore){
                                    favCount ++;                              // Favorable score
                                }
                            }if (!isEmpty(cell) && hasComments && cell.getCellType() == 1){
                                if(hasPresenter){
                                    presComments.add(str);                        
                                }else{
                                    genComments.add(str);
                                }
                            }
                        } // End row iterator
                    } // End columns loop    
                qList = cleanList(qList);
                if(s == 0 && firstSheet){
                    workbook.close();
                    input.close();
                    return qList;
                }
            } // End sheets loop
                        
            workbook.close();
            input.close();

            return qList;
        }
        
        /* Converts the list of columns into indices*/
        public static ArrayList<Integer> convertToIndex(ArrayList<String> input) {
            ArrayList<Integer> output = new ArrayList<>();
            for(String s : input){
                s = s.toUpperCase();
                int value = 0;
                for (int i = 0; i < s.length(); i++) {
                    int delta = (s.charAt(i)) - 64;
                    value = value*26+ delta;
                }
                output.add(value);
            }
            return output;
        }
        /*public static int getColumIndex(String columnName) {
            columnName = columnName.toUpperCase();
            int value = 0;
            for (int i = 0; i < columnName.length(); i++) {
                int delta = (columnName.charAt(i)) - 64;
                value = value*26+ delta;
            }
            return value-1;
        }*/
        
        /* Adds the extra info corresponding to the columns specified */
        private void addExtraInfo(Double num, int amt){
            if(!extraColumns.containsKey(num)){
                extraColumns.put(num, amt);
            }
            extraColumns.put(num,amt);
        }
        
        /* Adds question to the list that corresponds to the presenter */
        private void addQuestion(String pre, Question que){
            if(!qMap.containsKey(pre)){
                qMap.put(pre, new ArrayList<Question>());
            }
            qMap.get(pre).add(que);
        }
        
        /* Removes null values from the list and cleans out non-scoring questions*/
        private ArrayList<Question> cleanList(ArrayList<Question> q){
            ArrayList<Question> toDelete = new ArrayList<>();
            for(int i = 0; i < q.size(); i++){
                if(Double.isNaN(q.get(i).getScore())){  // Non-scoring
                    toDelete.add(q.get(i));
                }
                if (q.get(i) == null){                  // Null
                    toDelete.add(q.get(i));
                }
            }
            for(Question qu : toDelete){
                q.remove(qu);
            }
            return q;
        }
        
        /* Cleans the list up by removing nulls */
        private void cleanPresComments(){
            for(int c = 0; c < presComments.size(); c++){         // Checks for null values
                if (presComments.get(c) == null){
                    presComments.remove(c);
                }
            }
        }
        
        /* Sets all the fields for a Question object */
        private Question createQuestion(boolean hp,boolean wc,String qName, double score, double percent, int parts){
            Question q = new Question(hp, wc);
            q.setQuestion(qName);                                   
            q.setScore(score);
            q.setPercent(percent);
            q.setNumberParts(parts);
            return q;
        }
        
        /* Loops through the first column to get presenter names, then adds the names to a list */
        private ArrayList<String> getPresenters(Sheet sheet){
            ArrayList<String> names = new ArrayList<>();
            for(Row r : sheet) {    
                Cell c = r.getCell(0);
                if(!isEmpty(c)) {
                    names.add(c.getStringCellValue()); 
                }
            }
            return names;
        }
        
        /* Returns the amount of chunks of data on the provided sheet */
        private int getBreakAmount(Sheet sheet){
            int k = 0;
            int ends = 0;
                // Collect number of chunks of data
                for(Row r : sheet) {
                    Cell c = r.getCell(1);
                    String str = dataf.formatCellValue(c);
                    k++;
                    if(r.getRowNum()==0)k--;
                    if("end".equalsIgnoreCase(str)) {
                        particips.add(k-1);
                        k=0;
                        endRows.add(r.getRowNum()+1);
                        ends++;
                    }
                }
            return ends;
        }
        
        /* Returns the occurences of the given score */
        private int getScores(ArrayList<Double> allScores, double target){
            return Collections.frequency(allScores, target);
        }
        
        /* Simply tests for a null row */
        private boolean rowEmpty(Row r){
            return r == null;
        }
        
        /* Checks if a cell is empty or not */
        private static boolean isEmpty(final Cell cell) {
            if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) { // Legitimately empty cell
                return true;
            }
            if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().isEmpty()) { // Empty string in the cell
                return true;
            }
            return false;
        }
            
        /* Processes the cell, returning the type and value of the cell */
        private Object getCellValue(Cell cell) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                case Cell.CELL_TYPE_NUMERIC:
                    return cell.getNumericCellValue();
            }   
 
        return null;
        }
        
        /* Factory method for checking excel file type */
        private Workbook getWorkbook(FileInputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
            if (excelFilePath.endsWith("xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            } else if (excelFilePath.endsWith("xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                throw new IllegalArgumentException("The specified file is not Excel file");
            }
        return workbook;
    }
}