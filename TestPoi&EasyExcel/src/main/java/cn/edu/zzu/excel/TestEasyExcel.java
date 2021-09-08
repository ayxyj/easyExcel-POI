package cn.edu.zzu.excel;

import cn.edu.zzu.entity.excel.DataDemo;
import cn.edu.zzu.listener.DataDemoListener;
import com.alibaba.excel.EasyExcel;
import org.apache.poi.ss.formula.functions.T;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class TestEasyExcel {
    private static final Logger logger = LoggerFactory.getLogger(TestEasyExcel.class);

    public static void main(String[] args) {
        logger.info("开始！");
        try {
           // testWrite();
            testRead();
        } catch (IOException e) {
            e.printStackTrace();
        }
        logger.info("结束！");
    }

    //数据
    public static List<DataDemo> data() {
        ArrayList<DataDemo> dataDemos = new ArrayList<>();
        for (double i = 0; i < 10; i++) {
            dataDemos.add(new DataDemo("字符串" + i, new Date(), i));
        }
        return dataDemos;
    }

    private static final String PATH = "D:\\Java\\idea_workspace\\TestPoi&EasyExcel\\";

    public static void testWrite() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "easyExcel.xlsx");
        EasyExcel.write(fileOutputStream, DataDemo.class).sheet().doWrite(data());
    }

    public static void testRead() throws IOException{
        FileInputStream fileInputStream = new FileInputStream(PATH + "easyExcel.xlsx");
        EasyExcel.read(fileInputStream , DataDemo.class , new DataDemoListener()).sheet().doRead();
    }
}
