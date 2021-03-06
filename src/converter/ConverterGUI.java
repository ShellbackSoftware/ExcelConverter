package converter;

import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * GUI for the Excel converter
 * 
 * @author shellbacksoftware@gmail.com
*/
public class ConverterGUI extends javax.swing.JFrame {

    private String inputFilePath;           // Input file path
    private boolean hasPresenter = false;   // The data has presenters or not
    private boolean hasComments = false;    // Whether user wants comments included
    private boolean wantsCounts = false;    // Whether user wants counts of any columns
    private boolean firstSheet = true;      // Data is on first sheet or not
    private boolean multiSheets = false;    // Using multiple sheets or not
    private String columnCounts;            // Column numbers where the user wants the counts of each score
    private String sheets;                  // Sheet(s) the data is on
    private String outputName;              // Name of the output file
    private String outputFilePath;          // Location of the output file
    private String templatePath;            // Template word doc to be edited
    private double favScore;                // Cutoff for favorable score
    
    /**
     * Creates new form DataGUI
     */
    public ConverterGUI() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        instructionsPanel = new javax.swing.JPanel();
        instructions = new javax.swing.JLabel();
        options = new javax.swing.JPanel();
        FILEPATH = new javax.swing.JTextField();
        browseInputButton = new javax.swing.JButton();
        fileLabel = new javax.swing.JLabel();
        presenterCheck = new javax.swing.JRadioButton();
        countsCheck = new javax.swing.JRadioButton();
        colCounts = new javax.swing.JTextField();
        countCol = new javax.swing.JLabel();
        checkComments = new javax.swing.JRadioButton();
        outputFileName = new javax.swing.JTextField();
        outputFileLocLabel = new javax.swing.JLabel();
        outputFileNameLabel = new javax.swing.JLabel();
        outputLocExp = new javax.swing.JLabel();
        outputBrowseButton = new javax.swing.JButton();
        outputFileLoc = new javax.swing.JTextField();
        jSeparator1 = new javax.swing.JSeparator();
        jSeparator2 = new javax.swing.JSeparator();
        runButton = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        favScoreSet = new javax.swing.JTextField();
        jSeparator4 = new javax.swing.JSeparator();
        jLabel4 = new javax.swing.JLabel();
        browseTemplate = new javax.swing.JButton();
        templateDoc = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Excel Converter");
        setResizable(false);

        instructionsPanel.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
        instructionsPanel.setToolTipText("");
        instructionsPanel.setName("Instructions"); // NOI18N

        instructions.setText("<html>Before selecting a file, make sure that the first row in the sheet contains the question names / titles you want. Additionally, if a<br/> section of data pertains to the same host (i.e. presenter), make sure that the column to the left of the first set of data contains<br/> the names. At the end of each section of data, add in the word \"end\", and that way the program will know to move on to the next<br/> section. If you have comments, just make sure the word \"comment\" is somewhere in the column title.  If you need any more information,<br/> feel free to read the included readme, and look at the examples that are packaged with this program.</html>");
        instructions.setToolTipText("");
        instructions.setVerticalAlignment(javax.swing.SwingConstants.TOP);

        javax.swing.GroupLayout instructionsPanelLayout = new javax.swing.GroupLayout(instructionsPanel);
        instructionsPanel.setLayout(instructionsPanelLayout);
        instructionsPanelLayout.setHorizontalGroup(
            instructionsPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(instructionsPanelLayout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(instructions, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(10, 10, 10))
        );
        instructionsPanelLayout.setVerticalGroup(
            instructionsPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(instructionsPanelLayout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(instructions, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(10, 10, 10))
        );

        options.setBorder(javax.swing.BorderFactory.createTitledBorder("Options"));

        FILEPATH.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                FILEPATHActionPerformed(evt);
            }
        });

        browseInputButton.setText("Browse...");
        browseInputButton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        browseInputButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                browseInputButtonActionPerformed(evt);
            }
        });

        fileLabel.setFont(new java.awt.Font("sansserif", 0, 14)); // NOI18N
        fileLabel.setText("Input File:");

        presenterCheck.setText("Are there data delimiters?");
        presenterCheck.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                presenterCheckActionPerformed(evt);
            }
        });

        countsCheck.setText("Would you like the counts for any questions?");
        countsCheck.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                countsCheckActionPerformed(evt);
            }
        });

        colCounts.setText("A");
        colCounts.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                colCountsActionPerformed(evt);
            }
        });

        countCol.setText("If so, which column letter?");

        checkComments.setText("Would you like comments from the surveys included?");
        checkComments.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                checkCommentsActionPerformed(evt);
            }
        });

        outputFileName.setText("SurveyResults");
        outputFileName.setPreferredSize(new java.awt.Dimension(70, 28));
        outputFileName.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outputFileNameActionPerformed(evt);
            }
        });

        outputFileLocLabel.setFont(new java.awt.Font("sansserif", 0, 14)); // NOI18N
        outputFileLocLabel.setText("Output File Location:");

        outputFileNameLabel.setFont(new java.awt.Font("sansserif", 0, 14)); // NOI18N
        outputFileNameLabel.setText("Output File Name:");

        outputLocExp.setText("Leave blank or as \"Desktop\" to save on your Desktop");

        outputBrowseButton.setText("Browse...");
        outputBrowseButton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        outputBrowseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outputBrowseButtonActionPerformed(evt);
            }
        });

        outputFileLoc.setText("Desktop");
        outputFileLoc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outputFileLocActionPerformed(evt);
            }
        });

        runButton.setText("Run");
        runButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                runButtonActionPerformed(evt);
            }
        });

        jLabel2.setText("What would you like the favorable cutoff score to be?**");

        favScoreSet.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        favScoreSet.setText("3");
        favScoreSet.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                favScoreSetActionPerformed(evt);
            }
        });

        jLabel4.setFont(new java.awt.Font("sansserif", 0, 14)); // NOI18N
        jLabel4.setText("<html>Template Word file:* </html>");

        browseTemplate.setText("Browse...");
        browseTemplate.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        browseTemplate.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                browseTemplateActionPerformed(evt);
            }
        });

        templateDoc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                templateDocActionPerformed(evt);
            }
        });

        jLabel3.setText("<html>The cutoff score is where any values above the one specified</br> are factored into the favorable percentage.</html>");

        javax.swing.GroupLayout optionsLayout = new javax.swing.GroupLayout(options);
        options.setLayout(optionsLayout);
        optionsLayout.setHorizontalGroup(
            optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(optionsLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(optionsLayout.createSequentialGroup()
                        .addComponent(countsCheck, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(colCounts, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(12, 12, 12)
                        .addComponent(countCol, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(optionsLayout.createSequentialGroup()
                        .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                .addComponent(jSeparator4, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 722, Short.MAX_VALUE)
                                .addGroup(javax.swing.GroupLayout.Alignment.LEADING, optionsLayout.createSequentialGroup()
                                    .addComponent(fileLabel)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                    .addComponent(FILEPATH, javax.swing.GroupLayout.PREFERRED_SIZE, 412, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGap(18, 18, 18)
                                    .addComponent(browseInputButton))
                                .addGroup(javax.swing.GroupLayout.Alignment.LEADING, optionsLayout.createSequentialGroup()
                                    .addComponent(presenterCheck, javax.swing.GroupLayout.PREFERRED_SIZE, 301, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                    .addComponent(checkComments))
                                .addGroup(javax.swing.GroupLayout.Alignment.LEADING, optionsLayout.createSequentialGroup()
                                    .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(outputFileNameLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 122, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(outputFileLocLabel))
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                    .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                        .addComponent(outputFileName, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(outputFileLoc, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                    .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addGroup(optionsLayout.createSequentialGroup()
                                            .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                            .addComponent(templateDoc, javax.swing.GroupLayout.PREFERRED_SIZE, 155, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                            .addComponent(browseTemplate))
                                        .addGroup(optionsLayout.createSequentialGroup()
                                            .addComponent(outputBrowseButton)
                                            .addGap(10, 10, 10)
                                            .addComponent(outputLocExp))))
                                .addComponent(jSeparator1, javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(jSeparator2, javax.swing.GroupLayout.Alignment.LEADING))
                            .addGroup(optionsLayout.createSequentialGroup()
                                .addComponent(jLabel2)
                                .addGap(18, 18, 18)
                                .addComponent(favScoreSet, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 357, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, optionsLayout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(runButton, javax.swing.GroupLayout.PREFERRED_SIZE, 126, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(295, 295, 295))
        );
        optionsLayout.setVerticalGroup(
            optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(optionsLayout.createSequentialGroup()
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(FILEPATH, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(browseInputButton)
                    .addComponent(fileLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(presenterCheck)
                    .addComponent(checkComments))
                .addGap(5, 5, 5)
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(3, 3, 3)
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(countsCheck)
                    .addComponent(colCounts, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(countCol, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(favScoreSet, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator4, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(optionsLayout.createSequentialGroup()
                        .addGap(2, 2, 2)
                        .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(templateDoc, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(browseTemplate))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(outputBrowseButton)
                            .addComponent(outputLocExp)))
                    .addGroup(optionsLayout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(optionsLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(optionsLayout.createSequentialGroup()
                                .addComponent(outputFileNameLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(outputFileLocLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(optionsLayout.createSequentialGroup()
                                .addComponent(outputFileName, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(outputFileLoc, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(runButton))
        );

        jLabel5.setText("<html>* The Word file has certain keywords the program searches for; see the readme for examples.</html>");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(6, 6, 6)
                        .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(instructionsPanel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(options, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(instructionsPanel, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(options, javax.swing.GroupLayout.PREFERRED_SIZE, 325, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 21, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    /* Text box that contains the target word doc */
    private void templateDocActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_templateDocActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_templateDocActionPerformed

    /* */
    public void printError(String title, String error){
        JOptionPane.showMessageDialog(null,error, title, JOptionPane.WARNING_MESSAGE);
    }
    
    /* Opens file chooser for the target Word doc */
    private void browseTemplateActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_browseTemplateActionPerformed
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Word document", new String[]{"doc", "docx"});
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Select a file");
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            templateDoc.setText(chooser.getSelectedFile().getPath());
        }
    }//GEN-LAST:event_browseTemplateActionPerformed

    /* Box that user sets favorable score cutoff */
    private void favScoreSetActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_favScoreSetActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_favScoreSetActionPerformed

    /* Button to run the program */
    private void runButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_runButtonActionPerformed
        try {
            inputFilePath = FILEPATH.getText();
            favScore = Double.parseDouble(favScoreSet.getText());
            outputName = outputFileName.getText();
            if(outputFileLoc.getText().equalsIgnoreCase("Desktop") || (outputFileLoc.getText().equals(""))){
                outputFilePath = System.getProperty("user.home");
                outputFilePath = outputFilePath + "/Desktop";
            }else{
                outputFilePath = outputFileLoc.getText();                
            }
            columnCounts = colCounts.getText();
            templatePath = templateDoc.getText();
            sheets = "0";
            ExcelConverter ec = new ExcelConverter(inputFilePath, hasPresenter, multiSheets, firstSheet, sheets, favScore, hasComments, wantsCounts, columnCounts, outputName, outputFilePath, templatePath);
        } catch (IOException ex) { // Show a popup, telling user that the file was invalid.
            JOptionPane.showMessageDialog(null, 
                              "File was not found. Please make sure that your path is correct, and that the file exists.", 
                              "Invalid Path", 
                              JOptionPane.WARNING_MESSAGE);
        }
    }//GEN-LAST:event_runButtonActionPerformed

    /* Text box for the output file */
    private void outputFileLocActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outputFileLocActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_outputFileLocActionPerformed

    /* Opens file chooser for the output directory */
    private void outputBrowseButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outputBrowseButtonActionPerformed
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Select a directory");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            outputFileLoc.setText(chooser.getCurrentDirectory().getAbsolutePath());
        }
    }//GEN-LAST:event_outputBrowseButtonActionPerformed

    /* Desired name of output file */
    private void outputFileNameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outputFileNameActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_outputFileNameActionPerformed

    /* Users sets if they want comments included */
    private void checkCommentsActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_checkCommentsActionPerformed
        if(hasComments){
            hasComments = false;
        }
        else{
            hasComments = true;
        }
    }//GEN-LAST:event_checkCommentsActionPerformed

    /* User sets which column(s) they want the counts for */
    private void colCountsActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_colCountsActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_colCountsActionPerformed

    /* User sets if they want counts for any questions*/
    private void countsCheckActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_countsCheckActionPerformed
        if(wantsCounts){
            wantsCounts = false;
        }
        else{
            wantsCounts = true;
        }
    }//GEN-LAST:event_countsCheckActionPerformed

    /* User sets hasPresenter field in main class */
    private void presenterCheckActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_presenterCheckActionPerformed
        if(hasPresenter){
            hasPresenter = false;
        }
        else{
            hasPresenter = true;
        }
    }//GEN-LAST:event_presenterCheckActionPerformed

    /* Opens file chooser for input file */
    private void browseInputButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_browseInputButtonActionPerformed
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("XLS file", new String[]{"xls", "xlsx"});
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Select a file");
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            FILEPATH.setText(chooser.getSelectedFile().getPath());
        }
    }//GEN-LAST:event_browseInputButtonActionPerformed

    /* Text box that contains the filepath */
    private void FILEPATHActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_FILEPATHActionPerformed
        // Autogenerated code, no need to do anything here
    }//GEN-LAST:event_FILEPATHActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        ConverterGUI gui = new ConverterGUI();
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            gui.printError("ClassNotFound Exception", "Class not found, something went wrong!");
        } catch (InstantiationException ex) {
            gui.printError("Instantiation Exception", "A class couldn't be instantiated. Woops!");
        } catch (IllegalAccessException ex) {
            gui.printError("IllegalAccess Exception", "A class has been changed somehow. Sorry!");
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ConverterGUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ConverterGUI().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextField FILEPATH;
    private javax.swing.JButton browseInputButton;
    private javax.swing.JButton browseTemplate;
    private javax.swing.JRadioButton checkComments;
    private javax.swing.JTextField colCounts;
    private javax.swing.JLabel countCol;
    private javax.swing.JRadioButton countsCheck;
    private javax.swing.JTextField favScoreSet;
    private javax.swing.JLabel fileLabel;
    private javax.swing.JLabel instructions;
    private javax.swing.JPanel instructionsPanel;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JSeparator jSeparator4;
    private javax.swing.JPanel options;
    private javax.swing.JButton outputBrowseButton;
    private javax.swing.JTextField outputFileLoc;
    private javax.swing.JLabel outputFileLocLabel;
    private javax.swing.JTextField outputFileName;
    private javax.swing.JLabel outputFileNameLabel;
    private javax.swing.JLabel outputLocExp;
    private javax.swing.JRadioButton presenterCheck;
    private javax.swing.JButton runButton;
    private javax.swing.JTextField templateDoc;
    // End of variables declaration//GEN-END:variables
}
