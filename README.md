# 一个基于hutool的导入导出工具

[![Maven Central](https://maven-badges.herokuapp.com/maven-central/io.github.creekmoon/excel-utils-boot-starter/badge.svg)](https://mvnrepository.com/artifact/io.github.creekmoon/excel-utils-boot-starter)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)

前言:

公司最初基于POI封装了的Excel的工具类. 但后来因为种种原因需要更新这个玩意.

easy-excel它基于注解的形式对于动态表头并不友好

看了一圈网上都没什么现成的工具, 于是只能就自己手撸一个了.

## 开发目标

* 易于使用
* 符合直觉
* 降低心智负担
* (好玩)


## 使用示例
> 构建表头 --> 写入数据 --> 响应数据


```java 
//   导出简单表格
//   | 用户名  |  全名  |
//   ------------------
//   |  张三A  | 张三1  | 
//   |  张三B  | 张三2  |
//   |  张三C  | 张三3  |
@GetMapping(value = "/exportExcel")
public void exportExcel(HttpServletRequest request,HttpServletResponse response){
    
        ArrayList<Student> result=createStudentList(60_000);  
        
        ExcelExport.create("excelName",Student.class)            
        .addTitle("用户名",Student::getUserName)            
        .addTitle("全名",Student::getFullName)   //构建表头
        .write(result)                           //写入数据                            
        .response(response);                     //响应数据
 }
```

```java
//导出多级表头
//   |     用户名       |     年龄   |
//   | 用户名  |  全名  |            |
//   -------------------------------
//   |  张三A  | 张三1  |     20     |
//   |  张三B  | 张三2  |     20     |     
//   |  张三C  | 张三3  |     20     |
@GetMapping(value = "/exportExcel")
public void exportExcel(HttpServletRequest request,HttpServletResponse response){
        
        ArrayList<Student> result=createStudentList(60_000);  
        
        ExcelExport.create("excelName",Student.class)            
        .addTitle("基础信息::用户名",Student::getUserName)          
        .addTitle("基础信息::全名",Student::getFullName)
        .addTitle("年龄",Student::getAge)
        .write(result)                                        
        .response(response);                                
        }
```

```java
// 多Sheet页支持
@GetMapping(value = "/exportExcel")
public void exportExcel(HttpServletRequest request,HttpServletResponse response){

        ArrayList<Student> students = createStudentList(60_000);
        ArrayList<Teacher> teachers = createTeacherList(60_000);
        
        ExcelExport.create()
        .switchSheet("学生数据", Student.class)
        .addTitle("基础信息::用户名",Student::getUserName)           
        .addTitle("基础信息::全名",Student::getFullName)
        .addTitle("年龄",Student::getAge)
        .write(students)
        
        .switchSheet("老师数据", Teacher.class)
        .addTitle("基础信息::用户名",Teacher::getUserName)
        .addTitle("基础信息::全名",Teacher::getFullName)
        .addTitle("年龄",Teacher::getAge)
        .write(teachers)                                      
        .response(response);                               
     }
```
```java
// 导入数据
@PostMapping(value = "/importExcel")
public void importExcel(MultipartFile file,HttpServletRequest request,HttpServletResponse response){
        ExcelImport.create(file,Student::new)
        .addConvert("用户名",Student::setUserName)
        .addConvert("全名",Student::setFullName)
        .addConvert("年龄",IntegerConverter::parse,Student::setAge)
        .read(student->{
                System.out.println(student);
        })
        .response(response);
  }
```

## 快速开始

### 1.引入依赖
注意,从 2.0.0 版本开始, 是以Spring Boot 3.0 (JDK17)的版本进行开发.

```xml
<dependency>
    <groupId>cn.creekmoon</groupId>
  <artifactId>excel-utils-boot-starter</artifactId>
  <version>2.0.0</version>
</dependency>
```


### 2.使用@EnableExcelUtils注解启用工具类

放在Spring Boot启动类上

```java 
@EnableExcelUtils()
@SpringBootApplication()
public class ExampleApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExampleApplication.class, args);
    }
}
```

<details>
<summary>参数说明</summary>
##### 注解可选参数:

customExceptions: 自定义异常(已检查的异常) 用于告诉组件这是安全的异常.

importMaxParallel: 最大导入并发数量 这个参数可以控制同时进行多少个导入工作.防止OOM.

tempFileLifeMinutes: 临时文件寿命 后台维护一个线程进行定时清理临时文件. 默认是五分钟

```java
@EnableExcelUtils(customExceptions = {MyNewException.class}, importMaxParallel = 4, tempFileLifeMinutes = 5)
```

</details>




## 更多例子

请参考源码里的example


## 样式使用

### addTitle方法重载说明

`addTitle`提供三种重载方式，支持灵活的样式配置：

```java
// 1. 基础版本：仅添加标题和取值函数
addTitle(String titleName, Function<R, Object> valueFunction)

// 2. 简化样式版本：单个样式+条件判断
addTitle(String titleName, Function<R, Object> valueFunction, 
         ExcelCellStyle cellStyle, Predicate<R> cellStyleCondition)

// 3. 完整版本：支持多个条件样式
addTitle(String titleName, Function<R, Object> valueFunction, 
         ConditionCellStyle<R>... conditionStyle)
```

### 内置样式常量

工具提供三种预定义样式：

```java
ExcelCellStyle.LIGHT_ORANGE  // 浅橙色背景
ExcelCellStyle.PALE_BLUE     // 浅蓝色背景  
ExcelCellStyle.LIGHT_GREEN   // 浅绿色背景
```

### 样式使用示例

```java
@GetMapping(value = "/exportExcelWithStyle")
public void exportExcelWithStyle(HttpServletResponse response){
    ArrayList<Student> students = createStudentList(1000);
    
    ExcelExport.create("学生成绩表", Student.class)
        // 使用内置样式：年龄大于18的行显示浅绿色
        .addTitle("姓名", Student::getName, 
                  LIGHT_GREEN, student -> student.getAge() > 18)
        
        // 自定义样式：成绩低于60分的显示红色背景
        .addTitle("成绩", Student::getScore,
                  new ExcelCellStyle(style -> {
                      style.setFillForegroundColor(IndexedColors.RED.getIndex());
                      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                  }), student -> student.getScore() < 60)
        
        // 多条件样式：同时支持多种样式条件
        .addTitle("等级", Student::getGrade,
                  of(LIGHT_GREEN, s -> "A".equals(s.getGrade())),
                  of(PALE_BLUE, s -> "B".equals(s.getGrade())),
                  of(LIGHT_ORANGE, s -> "C".equals(s.getGrade())))
        
        .write(students)
        .response(response);
}
```

### 关键属性说明

- **titleName**: 表头名称，支持`::`分隔符创建多级表头
- **valueFunction**: Lambda表达式，用于从对象中提取单元格值  
- **ExcelCellStyle**: 样式定义类，可自定义背景色、字体、边框等
- **Predicate<R>**: 条件判断函数，返回true时应用对应样式
- **ConditionCellStyle**: 条件样式包装类，将样式和条件绑定

### 自定义样式

```java
// 创建自定义样式
ExcelCellStyle customStyle = new ExcelCellStyle((workbook, style) -> {
    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    // 设置数字格式
    style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
});

// 应用到标题
.addTitle("金额", Student::getAmount, customStyle, s -> s.getAmount() > 1000)
```

