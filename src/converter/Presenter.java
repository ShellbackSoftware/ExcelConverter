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

import java.util.ArrayList;

/**
 * Skeletal class to hold a presenter and their information
 * 
 * @author shellbacksoftware@gmail.com
 */
public class Presenter {
    private String name;
    private ArrayList<Question> questions;  // List of questions that belong to the presenter
    private ArrayList<String> comments;     // List of comments for the presenter
    
    public Presenter(){
    }
    
    /*@Override
    public String toString() {
        return String.format("Presentation: %s, Question: %s, Comments: %s", name, questions, comments);
    }*/
    
    public void setComments(ArrayList<String> c){
        comments = c;
    }
    
    public ArrayList<String> getComments(){
        return comments;
    }
    
    public void setQuestions(ArrayList<Question> q){
        questions = q;
    }
    
    public ArrayList<Question> getQuestions(){
        return questions;
    }
    
    public ArrayList<String> getQuestionScore(){
        ArrayList<String> results = new ArrayList<>();
        for(Question q : questions){
            results.add(Double.toString(q.getScore()));
        }
        return results;
    }
    
    public ArrayList<String> getQuestionNames(){
        ArrayList<String> qs = new ArrayList<>();
        for(Question q : questions){
            qs.add(q.getQuestion());
        }
        return qs;
    }
    
    public void setName(String n){
        name = n;
    }
    
    public String getName(){
        return name;
    }
}
