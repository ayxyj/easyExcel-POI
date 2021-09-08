package cn.edu.zzu.dao;

import cn.edu.zzu.entity.excel.DataDemo;

import java.util.List;

public class ExcelDao {
    public void saveDate(List<DataDemo> data) {
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert,所有数据一次性插入
        System.out.println("存入数据库 " + data.size() + " 条：" + data);
    }
}
