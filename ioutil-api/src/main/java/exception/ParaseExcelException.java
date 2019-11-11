package exception;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ParaseExcelException extends RuntimeException{
    public ParaseExcelException(){
        super();
    }
    public ParaseExcelException(String msg){
        super(msg);
    }
    public  ParaseExcelException(String msg,Throwable cause){
        super(msg,cause);
    }
}
