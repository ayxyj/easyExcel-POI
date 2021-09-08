package cn.edu.zzu.listener;

import cn.edu.zzu.dao.ExcelDao;
import cn.edu.zzu.entity.excel.DataDemo;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import lombok.Data;
import lombok.extern.java.Log;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

// 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
public class DataDemoListener extends AnalysisEventListener<DataDemo> {


    private static final Logger log = LoggerFactory.getLogger(DataDemoListener.class);

    private static final int BATCH_COUNT = 5;

    private static final List<DataDemo> list = new ArrayList<DataDemo>();

    //持久化
    private ExcelDao excelDao ;

    public DataDemoListener() {
        excelDao = new ExcelDao();
    }

    public DataDemoListener(ExcelDao excelDao) {
        this.excelDao = excelDao;
    }

    /**
     *
     * @param dataDemo
     * @param analysisContext
     */
    @Override
    public void invoke(DataDemo dataDemo, AnalysisContext analysisContext) {
        System.out.println(JSON.toJSONString(dataDemo));
        list.add(dataDemo);
        if (list.size() > BATCH_COUNT){
            excelDao.saveDate(list);
            list.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        //处理剩下的数据持久化
        excelDao.saveDate(list);
        log.info("存储结束！");
    }
}
