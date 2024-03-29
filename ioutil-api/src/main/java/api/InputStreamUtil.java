package api;

import java.io.*;

/**
 * @author Sherlock.Wu
 * @date 2019/11/26
 */
public class InputStreamUtil {
    /**
     *
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static byte[] getBytes(InputStream inputStream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        byte[] dataBytes;
        while ((len = inputStream.read(buffer)) != -1) {
            baos.write(buffer, 0, len);
        }
        baos.flush();
        dataBytes = baos.toByteArray();
        return dataBytes;
    }

    /**
     * 
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static InputStream getNewStream(InputStream inputStream) throws IOException {
        byte[] dataBytes = getBytes(inputStream);
        return new ByteArrayInputStream(dataBytes);
    }

    /**
     * 把inputStream 中的数据放到outputStream中
     * @param inputStream
     * @param outputStream
     * @throws IOException
     */
    public static void inputToOutputStream(InputStream inputStream,OutputStream outputStream)throws IOException{
        byte[] buffer = new byte[1024];
        int len;
        byte[] dataBytes;
        while ((len = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, len);
        }
        outputStream.flush();
    }
    public static void close(InputStream inputStream){
        try {
            if(null != inputStream){
                inputStream.close();
            }
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
