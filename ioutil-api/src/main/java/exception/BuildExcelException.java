package exception;

/**
 * 生成excel异常类
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class BuildExcelException extends RuntimeException {
    public BuildExcelException(){
        super();
    }
    public BuildExcelException(String msg){
        super(msg);
    }
    public BuildExcelException(String msg, Throwable cause){
        super(msg,cause);
    }
}
