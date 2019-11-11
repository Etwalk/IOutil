package exception;

/**
 * 生成excel异常类
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ExcelBuilderException extends RuntimeException {
    public ExcelBuilderException(){
        super();
    }
    public ExcelBuilderException(String msg){
        super(msg);
    }
    public  ExcelBuilderException(String msg,Throwable cause){
        super(msg,cause);
    }
}
