package com.zmj.rightools.excel.entity;

import lombok.Data;

import java.util.Date;

/**
 * @author zhangmingjian
 * @date 2023/5/1
 */
@Data
public class Personnel {
    
    private Integer id;
    
    private String name;
    
    private Date birthday;
    
    private Integer gender;
    
    private String country;
    
    private String industry;
    
    private String comment;
}
