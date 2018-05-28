
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Administrator on 2018/5/23 0023.
 */
public class Main {
    public static void main(String[] args) {
        try {
            System.out.println("ppt2pdf只允许单线程运行");
            SimpleDateFormat dateFormat= new SimpleDateFormat("yyyyMMdd");
            String nowDirName = "";
            String path="";
            while(true) {
                Date date = new Date();
                File directory = new File("");
                String dirName = dateFormat.format(date);
                if (!dirName.equals(nowDirName)) {
                    nowDirName = dirName;
                    path = directory.getAbsolutePath();
                    System.out.println(nowDirName);
                    System.out.println("当前监视目录"+directory.getAbsolutePath()+"\\?\\"+nowDirName);//获取路径
                }
                File directory2 = new File(path);
                String[] files = directory2.list();
                for (String dirName2:files) {
                    String checkPath = directory.getAbsolutePath()+"\\"+dirName2+"\\";
                    if (new File(checkPath).isDirectory()) {
                        checkPath +=  nowDirName+"\\";
                        if (new File(checkPath).isDirectory()) {
                            String[] files2 = new File(checkPath).list();
                            for (String file:files2) {
                                String suffix = file.substring(file.lastIndexOf("."));//后缀名
                                if (".ppt".equals(suffix)||".pptx".equals(suffix)) {
                                    String fileName = file.substring(0,file.lastIndexOf("."));
                                    String tartName = checkPath+fileName;
                                    File exFile = new File(tartName+".pdf");
                                    if (!exFile.exists()) {
                                        System.out.println("开始转换"+tartName+suffix);
                                        Word2Pdf(tartName+suffix, tartName+".pdf");
                                    }
                                }
                            }
                        }

                    }
                }
                Thread.sleep(800);
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
            File file = new File("errLog.txt");
            PrintStream stream = null;
            e.printStackTrace(stream);
            stream.flush();
            stream.close();
        }

    }

    // 将word格式的文件转换为pdf格式
    public static void Word2Pdf(String srcPath, String desPath) throws IOException {
        // 源文件目录
        File inputFile = new File(srcPath);
        if (!inputFile.exists()) {
            System.out.println("源文件不存在！");
            return;
        }
        // 输出文件目录
        File outputFile = new File(desPath);
        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().exists();
        }
        // 调用openoffice服务线程
        String command = "C:/Program Files (x86)/OpenOffice 4/program/soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
/*        String command = sofPath+"soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";*/
        Process p = Runtime.getRuntime().exec(command);

        // 连接openoffice服务
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(
                "127.0.0.1", 8100);
        connection.connect();

        // 转换word到pdf
        DocumentConverter converter = new OpenOfficeDocumentConverter(
                connection);
        converter.convert(inputFile, outputFile);

        // 关闭连接
        connection.disconnect();

        // 关闭进程
        p.destroy();
        System.out.println("转换完成！");
    }

}
