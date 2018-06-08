


import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.*;
import org.artofsolving.jodconverter.process.ProcessManager;
import org.artofsolving.jodconverter.process.ProcessQuery;
import org.artofsolving.jodconverter.process.PureJavaProcessManager;

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
/*            SimpleDateFormat dateFormat= new SimpleDateFormat("yyyyMMdd");
            String nowDirName = "";
            String path="";*/
//            while(true) {
/*                Date date = new Date();
                File directory = new File("");*/
/*                String dirName = dateFormat.format(date);
                if (!dirName.equals(nowDirName)) {
                    nowDirName = dirName;
                    path = directory.getAbsolutePath();
                    System.out.println(nowDirName);
                    System.out.println("当前监视目录"+directory.getAbsolutePath()+"\\?\\"+nowDirName);//获取路径
                }*/

            File directory2 = new File(new File("").getAbsolutePath()+"\\");
                String[] files = directory2.list();
                for (String dirName2:files) {
                    String checkPath = directory2.getAbsolutePath()+"\\"+dirName2+"\\";
                    if (new File(checkPath).isDirectory()) {
                        String[] files2 = new File(checkPath).list();
                        for (String dirName3:files2) {
                            //checkPath +=  nowDirName+"\\";
                            String checkPath2 =  checkPath+dirName3+"\\";
                            if (new File(checkPath2).isDirectory()) {
                                String[] files3= new File(checkPath2).list();
                                for (String file:files3) {
                                    System.out.println(checkPath2+file);
                                    if (!new File(checkPath2+file).isDirectory()) {
                                        if (file.split("\\..").length ==2) {
                                            String suffix = file.substring(file.lastIndexOf("."));//后缀名
                                            if (".ppt".equals(suffix)||".pptx".equals(suffix)) {
                                                String fileName = file.substring(0,file.lastIndexOf("."));
                                                String tartName = checkPath2+fileName;
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
                        }


                    }
                }
 //           }
            System.out.println("所有目录转换完成");
        } catch (Exception e) {
            System.out.println(e.getMessage());
/*            File file = new File("errLog.txt");
            PrintStream stream = null;
            e.printStackTrace(stream);
            stream.flush();
            stream.close();*/
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
 /*       // 调用openoffice服务线程
        String command = "C:/Program Files (x86)/OpenOffice 4/program/soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
*//*        String command = sofPath+"soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";*//*
        Process p = Runtime.getRuntime().exec(command);

        // 连接openoffice服务
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(
                "127.0.0.1", 8100);
        connection.connect();
try {
    // 转换word到pdf
    DocumentConverter converter = new OpenOfficeDocumentConverter(
            connection);
    converter.convert(inputFile, outputFile);
    System.out.println("转换完成！");
} catch (Exception e) {
    e.printStackTrace();
    System.out.println("转换失败！");
} finally {

    // 关闭连接
    connection.disconnect();

    // 关闭进程
    p.destroy();
}*/

        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
        String  OPEN_OFFICE_HOME="D:\\lxzProject\\ppttopdf\\LibreOffice 5\\";
        int OPEN_OFFICE_PORT = 8100;
        ProcessManager processManager = new PureJavaProcessManager();
        OfficeManager officeManager;
        try {
            System.out.println("准备启动安装在" + OPEN_OFFICE_HOME + "目录下的openoffice服务....");
            configuration.setOfficeHome(OPEN_OFFICE_HOME);//设置OpenOffice.org安装目录
            configuration.setPortNumbers(OPEN_OFFICE_PORT); //设置转换端口，默认为8100
            configuration.setTaskExecutionTimeout(1000 * 60 * 5L);//设置任务执行超时为5分钟
            configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);//设置任务队列超时为24小时
            configuration.setProcessManager(processManager);

            officeManager = configuration.buildOfficeManager();
            officeManager.start();    //启动服务
            System.out.println("office转换服务启动成功!");
            long startTime = System.currentTimeMillis();
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
            converter.convert(new File(srcPath),new File(desPath));
            System.out.println("转换完成.耗时" +( (System.currentTimeMillis() - startTime) / 60.0)+ "秒");
            System.out.println("关闭office转换服务....");

            if (officeManager != null) {
                officeManager.stop();
            }

            System.out.println("关闭office转换成功!");
            System.out.println("运行结束");
        } catch (Exception ce) {
            System.out.println("office转换服务启动失败!详细信息:" + ce);
        }
    }

}
