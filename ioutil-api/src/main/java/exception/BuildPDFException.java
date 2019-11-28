package exception;

/**
 * @author Sherlock.Wu
 * @date 2019/11/27
 */
public class BuildPDFException extends RuntimeException{
    public BuildPDFException(){
        super();
    }
    public BuildPDFException(String msg){
        super(msg);
    }
    public BuildPDFException(String msg, Throwable cause){
        super(msg,cause);
    }
}
