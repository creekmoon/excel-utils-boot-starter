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

## 引入依赖

```xml

<dependency>
    <groupId>io.github.creekmoon</groupId>
  <artifactId>excel-utils-boot-starter</artifactId>
  <version>1.2.5</version>
</dependency>
```

## 快速开始

### 1.使用@EnableExcelUtils注解启用工具类

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
##### 可选参数:

customExceptions: 自定义异常,若导入过程中抛出这些异常, 则会以其msg作为导入的结果.

importMaxParallel: 最大导入并发数量 这个参数可以控制同时进行多少个导入工作.防止OOM.

tempFileLifeMinutes: 临时文件寿命 导出时会保存一份临时文件在本地. 后台维护一个线程进行定时清理. 默认是五分钟

```java
@EnableExcelUtils(customExceptions = {MyNewException.class}, importMaxParallel = 4, tempFileLifeMinutes = 5)
```

</details>

### 2.导出简单表格

如下例子, 导出学生表的Excel,

```java
@GetMapping(value = "/exportExcel")
public void exportExcel(HttpServletRequest request,HttpServletResponse response){

        ArrayList<Student> result=createStudentList(60_000);   //模拟一批List数据
        ExcelExport.create("excelName",Student.class)            //声明一次导出, 导出对象是Student.class
        .addTitle("用户名",Student::getUserName)            //定义Excel的模板
        .addTitle("全名",Student::getFullName)
        .write(result)                                        //往里面写入数据 ,可以写多次
        .response(response);                                //响应给前端
        }
```

如上几行代码, 已经完成了一个导出的响应.

<details>
<summary>参数说明</summary>
- 使用 **ExcelExport.create()** 创建初始化本次导出, 第一个参数是返回给前端的文件名称, 第二个参数本次导出目标class.
- 使用 **addTitle()** 增加标题, 第一个参数是标题的名称, 第二个参数是一个lambda表达式, 定义如何从class实例中获取数据
- 使用 **write()** 写入数据
</details>

### 3.导出多级表头

只需要改写title, 将父表头**以::分隔即可**. 例如 "**用户名**","**全名**"的父标题是"**基础信息**"

```java
@GetMapping(value = "/exportExcel")
public void exportExcel(HttpServletRequest request,HttpServletResponse response){

        ArrayList<Student> result=createStudentList(60_000);   //模拟一批List数据
        ExcelExport.create("excelName",Student.class)            //声明一次导出, 导出对象是Student.class
        .addTitle("基础信息::用户名",Student::getUserName)            //定义Excel的模板
        .addTitle("基础信息::全名",Student::getFullName)
        .addTitle("年龄",Student::getAge)
        .write(result)                                        //往里面写入数据 ,可以写多次
        .response(response);                                //响应给前端
        }
```

如上几行代码, 已经完成了一个导出的响应.

### 4.导入表格


```java
@PostMapping(value = "/importExcel")
public void importExcelBySax(MultipartFile file,HttpServletRequest request,HttpServletResponse response){
        ExcelImport.create(file,Student::new)
        .addConvert("用户名",Student::setUserName)
        .addConvert("全名",Student::setFullName)
        .addConvert("年龄",IntegerConverter::parse,Student::setAge)
        .addConvert("邮箱",Student::setEmail)
        .addConvert("生日",DateConverter::parse,Student::setBirthday)
        .addConvert("过期时间",LocalDateTimeConverter::parse,Student::setExpTime)
        .read(student->{
        System.out.println(student);
        })
        .response(response);
        }
```

如上几行代码, 已经完成了一个导入功能

<details>
<summary>参数说明</summary>
如下例子, 导入学生表的Excel,

- 使用 **ExcelImport.create()** 创建初始化本次导入, 第一个参数是前端传入的MultipartFile对象, 第二个参数本次解析的目标class.
- 使用 **addConvert()** 增加转换器,每个转换器对应一个表头. 转换器有三个参数, 第一个参数是表头名称,
  第二个参数是解析器,在需要校验或者转换数据类型时可以选用, 第三个参数对应目标的Setter方法
- 使用 **read()** 读取数据
- 使用 **response()** 响应给前端结果

</details>

## 更多例子

请参考源码里的example

