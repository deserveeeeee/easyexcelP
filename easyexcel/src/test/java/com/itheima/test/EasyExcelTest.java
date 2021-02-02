package com.itheima.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.itheima.domain.FillData;
import com.itheima.domain.Student;
import com.itheima.listener.StudentListener;
import org.junit.Test;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * @Author Vsunks.v
 * @Date 2020/3/16 19:32
 * @Description:
 */
public class EasyExcelTest {


    /**
     * 案例练习：报表
     */
    @Test
    public void test07() {
        // 加载模板
        InputStream templateIs = this.getClass().getClassLoader().getResourceAsStream(
                "report_template.xlsx");

        // 指定输入位置
        String targetFilePath = "报表.xlsx";

        // 构建workbook对象使用模板
        ExcelWriter workBook =
                EasyExcel.write(targetFilePath, FillData.class).withTemplate(templateIs).build();

        // 工作表对象
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
/*
        // 组合写入 需要准备水平填充
        FillConfig fillConfig =
                FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();*/

// ****** 准备数据 *******
        // 日期
        HashMap<String, String> dateMap = new HashMap<String, String>();
        dateMap.put("date", "2020-03-16");

        // 总会员数
        HashMap<String, String> totalCountMap = new HashMap<String, String>();
        dateMap.put("totalCount", "1000");

        // 新增员数
        HashMap<String, String> increaseCountMap = new HashMap<String, String>();
        dateMap.put("increaseCount", "100");

        // 本周新增会员数
        HashMap<String, String> increaseCountWeekMap = new HashMap<String, String>();
        dateMap.put("increaseCountWeek", "50");

        // 本月新增会员数
        HashMap<String, String> increaseCountMonthMap = new HashMap<String, String>();
        dateMap.put("increaseCountMonth", "100");


        // 新增会员数据
        List<Student> students = initData();
        // **** 准备数据结束****


        // 写入统计数据
        workBook.fill(dateMap, writeSheet);
        workBook.fill(totalCountMap, writeSheet);
        workBook.fill(increaseCountMap, writeSheet);
        workBook.fill(increaseCountWeekMap, writeSheet);
        workBook.fill(increaseCountMonthMap, writeSheet);


        // 填充 多组数据
        workBook.fill(students, writeSheet);


        // 记得关闭流！！！！！！
        workBook.finish();

    }




    /**
     * 水平填充
     */
    @Test
    public void test06() {
        // 加载模板
        InputStream templateIs = this.getClass().getClassLoader().getResourceAsStream(
                "fill_data_template4.xlsx");

        // 指定输入位置
        String targetFilePath = "模板写入-水平填充数据.xlsx";

        // 构建workbook对象使用模板
        ExcelWriter workBook =
                EasyExcel.write(targetFilePath, FillData.class).withTemplate(templateIs).build();

        // 工作表对象
        WriteSheet writeSheet = EasyExcel.writerSheet().build();

        // 组合写入 需要准备水平填充
        FillConfig fillConfig =
                FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();


        // 准备数据
        List<FillData> fillData = initFillData();



        // 填充 多组数据
        workBook.fill(fillData,fillConfig, writeSheet);


        // 记得关闭流！！！！！！
        workBook.finish();

    }


    /**
     * 组合数据填充，因为组合填充需要多次写入，就不能写入一次以后就关闭流，所以不要使用doFill，而是使用fill
     */
    @Test
    public void test05() {
        // 加载模板
        InputStream templateIs = this.getClass().getClassLoader().getResourceAsStream(
                "fill_data_template3.xlsx");

        // 指定输入位置
        String targetFilePath = "模板写入-组合数据.xlsx";

        // 构建workbook对象使用模板
        ExcelWriter workBook =
                EasyExcel.write(targetFilePath, FillData.class).withTemplate(templateIs).build();

        // 工作表对象
        WriteSheet writeSheet = EasyExcel.writerSheet().build();

        // 组合写入 需要准备换行
        FillConfig fillConfig = FillConfig.builder().forceNewRow(true).build();


        // 准备数据
        List<FillData> fillData = initFillData();

        HashMap<String, String> map = new HashMap<String, String>();
        map.put("date", "2020-03-16");
        map.put("total", "1000");


        // 填充 多组数据
        workBook.fill(fillData,fillConfig, writeSheet);


        // 填充 单组数据
        workBook.fill(map, writeSheet);

        // 记得关闭流！！！！！！
        workBook.finish();

    }



    /**
     * 多组数据填充
     */
    @Test
    public void test04() {
        // 加载模板
        InputStream templateIs = this.getClass().getClassLoader().getResourceAsStream(
                "fill_data_template2.xlsx");

        // 指定输入位置
        String targetFilePath = "模板写入-多组数据.xlsx";

        // 构建workbook对象使用模板
        ExcelWriterBuilder writeWorkBook = EasyExcel.write(targetFilePath, FillData.class).withTemplate(templateIs);

        // 准备数据
        List<FillData> fillData = initFillData();

        // 构建sheet & 填充
        writeWorkBook.sheet().doFill(fillData);

    }

    /**
     * 单组数据填充
     */
    @Test
    public void test03() {
        // 加载模板
        InputStream templateIs = this.getClass().getClassLoader().getResourceAsStream("fill_data_template1.xlsx");

        // 指定输入位置
        String targetFilePath = "模板写入-单组数据.xlsx";

        // 构建workbook对象使用模板
        ExcelWriterBuilder writeWorkBook = EasyExcel.write(targetFilePath, FillData.class).withTemplate(templateIs);

        // 准备数据
        FillData fillData = new FillData();
        fillData.setAge(10);
        fillData.setName("杭州黑马");

        // 构建sheet & 填充
        writeWorkBook.sheet().doFill(fillData);

    }


    /**
     * 需求：单实体导出
     * 导出多个学生对象到Excel表格
     * 包含如下列：姓名、性别、出生日期
     * 模板详见：杭州黑马在线202003班学员信息.xlsx
     * <p>
     * <p>
     * 写不需要监听器
     */
    @Test
    public void test02() {
        // workbook
        ExcelWriterBuilder writeWorkBook = EasyExcel.write("杭州黑马在线202003班学员信息表write.xlsx", Student.class);

        // sheet
        ExcelWriterSheetBuilder sheet = writeWorkBook.sheet();

        // 初始化数据
        List<Student> students = initData();

        // 写
        sheet.doWrite(students);
    }


    /**
     * 需求：单实体导入
     * 导入Excel学员信息到系统。
     * 包含如下列：姓名、性别、出生日期
     * 模板详见：杭州黑马在线202003班学员信息.xls
     */
    @Test
    public void test01() {

        /**
         * 创建一个工作簿对象
         *
         * @param pathName
         *            要读取的文件路径
         * @param head
         *            每一行内容封装成的对象的类型
         * @param readListener
         *            读的监听器
         * @return 工作簿对象
         */
        // 工作簿，对应的是一个excel文件
        ExcelReaderBuilder readWorkBook = EasyExcel.read("杭州黑马在线202003班学员信息表.xlsx", Student.class, new StudentListener());
        // 工作表，对应一个sheet。 一个工作簿里面可以含有多个工作表
        ExcelReaderSheetBuilder sheet = readWorkBook.sheet();
        //读
        sheet.doRead();

    }

    /**
     * 初始好数据
     *
     * @return
     */
    private static List<Student> initData() {
        ArrayList<Student> students = new ArrayList<Student>();
        Student data = new Student();
        for (int i = 0; i < 10; i++) {
            data.setName("杭州黑马学号0" + i);
            data.setBirthday(new Date());
            data.setGender("男");
            students.add(data);
        }
        return students;
    }

    private static List<FillData> initFillData() {
        ArrayList<FillData> fillDatas = new ArrayList<FillData>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            fillData.setName("杭州黑马0" + i);
            fillData.setAge(10 + i);
            fillDatas.add(fillData);
        }
        return fillDatas;
    }

}
