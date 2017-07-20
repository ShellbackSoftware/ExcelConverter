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
    
    public void setCounts(Map<Double,Integer> c){
        counts = c;
    }
    
    public Map getCounts(){
        return counts;
    }
    
    public ArrayList<String> getComments(){
        return coms;
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