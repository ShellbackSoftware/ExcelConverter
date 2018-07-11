# ExcelConverter
To get started, there are a few example files packaged with this program. 

If you want to see an example Excel input file, simply go to Excel2Word Converter/TestExcelSheets. The first 
file, delimiter.xlsx, is an example of a file with delimiters. For example, presenters at a conference, where
each chunk of data correlates to a specific presenter. The second file, nodelimiter.xls, is an example of a 
file with only questions. 

Example templates can be found in Excel2Word Converter/Templates. There's an example of a template while using
a delimiter, and an example without a delimiter. They're just generic templates to show how to use the search
strings, along with some lorem ipsum text for filler.

Now, on to preparing the data:
For the Excel file, there are two possible formats. One with delimiters, and one without. 

1. With delimiters. If you have delimiters (i.e. presenters in a conference), make sure the delimiters are in
the left-most column (Column A). At the end of the delimiters column, just put "end" in the cell below the last 
delimiter. That tells the program to switch gears away from delimiters. Then, at the end of each chunk of data
for the delimiter, put "end" in the cell below the last value. You can look at "delimiter.xlsx" to see what 
the format should look like. 

2. Without delimiters. If you aren't using delimiters, so it's just a bunch of questions, all you need to do
is add "end" to the bottom of each column. A very easy way to do this, is just type "end" under the first 
column, then click the bottom right corner of that cell, and drag it to the right until all the columns that
have a question are selected. That'll copy/paste "end" into all the cells, saving you a lot of typing and time.

Currently, in Version 1.0, there is no way to handle comments from a file using format #2. Unfortunately, if
you're using format #2 and you have a question with comments, you'll have to manually copy and paste them
into the word document. However, if you are using format #1 and have comments, just make sure that the word
"comment" is somewhere in the question title. 

No matter which format you are using, the basic layout must be the same. The question titles are across Row 1, 
in each column. The data for each question goes down the column. Then, at the end of each chunk of data, type
"end" in the cell.

For the template, there's an example of each type in the /Templates folder. You can make one however you like,
just use the following key words:

%DELIMITER - The delimiter. You only need to make one page, it'll make the same amount of pages as number of
delimiters, and fill in the appropriate information for each one.

%INFO - The table of questions with their corresponding scores. This is the only required keyword.

%COUNTS - The table of counts for a specific column. 

%COMMENTS - List of comments for the delimiter

There are a few restrictions on this, however.  In the current version, as I mentioned earlier, %COMMENTS 
can't be used when there is no delimiter.  Additionally, the table for counts only works for one column at
the moment.  You can put in multiple questions for counts, but they all populate the same table. If the need
arises in the future, that will be fixed in a future version.  For now, if you put in just one column letter,
then the table shows the counts from just the one question. If you put in a list of letters, with a comma in
between, you'll get one table with a lot of data in it.

After the files are set up, the interface of the program will explain the rest. If you want to see the program
in action, just match up an example input file with the proper example template, and look at the resulting 
file on the desktop.

Lastly, due to some intricacies with Java, if you would like to save the output file to your desktop, simply 
leave the output file location as "Desktop" or blank.

Hopefully this makes your day easier, and thank you for using it!
