package api;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

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
}
