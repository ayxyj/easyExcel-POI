package cn.edu.zzu.entity.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class DataDemo {
    @ExcelProperty("字符串")
    private String string;
    @ExcelProperty("日期类型")
    private Date date;
    @ExcelProperty("数值型")
    private Double doubleData;
}
