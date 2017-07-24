package converter;

import java.util.ArrayList;

/**
 * Skeletal class to hold a delimiter (initially, a presenter, hence the class name) 
 * and the information relating to that delimiter.
 * 
 * @author shellbacksoftware@gmail.com
 */
public class Presenter {
    private String name;
    private ArrayList<Question> questions;  // List of questions that belong to the presenter
    private ArrayList<String> comments;     // List of comments for presenter
    
    public Presenter(){
    }
    
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
