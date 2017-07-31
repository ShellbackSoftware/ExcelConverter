package converter;

/*
 * Exception that is thrown if the user leaves the Output Name field blank
 */
public class EmptyFileNameException extends Exception {
    public EmptyFileNameException(String msg){
        super(msg);
    }
}
