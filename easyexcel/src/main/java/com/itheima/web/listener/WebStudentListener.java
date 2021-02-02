package com.itheima.web.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.itheima.domain.Student;
import com.itheima.service.StudentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import java.util.ArrayList;

/**
 * @Author Vsunks.v
 * @Date 2020/3/16 20:37
 * @Description:
 */
@Component
@Scope("prototype")
public class WebStudentListener extends AnalysisEventListener<Student> {


    @Autowired
    StudentService studentService;


    private final Integer STUDENT_COUNT = 5;
    ArrayList<Student> students = new ArrayList<Student>();

    /**
     * 读一行，调用一次
     *
     * @param student
     * @param context
     */
    public void invoke(Student student, AnalysisContext context) {
        students.add(student);

        if (students.size()%STUDENT_COUNT == 0) {
            studentService.save(students);
            students.clear();
        }
    }

    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
