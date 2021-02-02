package com.itheima.service.impl;

import com.itheima.domain.Student;
import com.itheima.service.StudentService;
import org.springframework.stereotype.Service;

import java.util.ArrayList;

/**
 * @Author Vsunks.v
 * @Date 2020/3/16 20:41
 * @Description:
 */
@Service
public class StudentServiceImpl implements StudentService {
    public void save(ArrayList<Student> students) {
        System.out.println("students = " + students);
    }
}
