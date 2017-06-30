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
import java.util.HashMap;
import java.util.Map;

/*
 * Class to represent the question information. Each method does exactly what it's called.
 *
 * @author shellbacksoftware@gmail.com
 */
public class Question {
    private String question;                    // Question being asked (New object for each question)
    private double avgScore;                    // Average score of all surveys
    private double avgPercent;                  // Not actually average, favorable percentage of surveys 
    private double numPartic;                   // Number of participants
    private ArrayList<String> coms;             // List of comments for specific questions
    private boolean hasPresenters;
    private Map<Double, Integer> counts;
    private boolean wantsCounts;
    
    /* Default constructor */
    public Question(boolean hp, boolean wc){
        coms = new ArrayList<>();
        counts = new HashMap();
        hasPresenters = hp;
        wantsCounts = wc;
    }
    
    @Override
    public String toString() {
        if(hasPresenters && wantsCounts){
            return String.format("Question: %s \n Score: %.1f \n Favorable Percent: %.0f%% \n Counts: %s",question, avgScore, avgPercent, getCounts());
        }else if(hasPresenters && !wantsCounts){
                    return String.format("Question: %s, Score: %.1f, Favorable Percent: %.0f%%",question, avgScore, avgPercent);
        }else if(!hasPresenters && wantsCounts){
            return String.format("Participants: %.0f, Question: %s, Score: %.1f, Favorable Percent: %.0f%%, Counts: %s",numPartic, question, avgScore, avgPercent, getCounts());
        }else{
            return String.format("Participants: %.0f, Question: %s, Score: %.1f, Favorable Percent: %.0f%%",numPartic, question, avgScore, avgPercent);
        }
    }
    
    public void setCounts(Map<Double,Integer> c){
        counts = c;
    }
    
    public Map getCounts(){
        return counts;
    }
    
    public void setComments(ArrayList<String> in){
        coms = in;
    }
    
    public void setNumberParts(double p){
        numPartic = p;
    }

    public double getNumberParts(){
        return numPartic;
    }
    
    public double getScore(){
        return avgScore;
    }
    
    public void setScore(double a){
        this.avgScore = a;
    }
    
    public double getPercent(){
        return avgPercent;
    }
    
    public void setPercent(double p){
        this.avgPercent = p;
    }
    
    public void setQuestion(String qu){
        this.question = qu;
    }

    public String getQuestion(){
        return question;
    }
}