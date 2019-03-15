# excelhelper-方便进行Excel读写的工具
---  
## 项目简介：
excelhelper是对Apache POI库的封装，主要用于Excel文件的读写，方便快捷的将Excel数据读取映射到Java POJO对象上或反之。   
Apache POI库提供了对EXcel文件读写的能力，但是在工作中，我们常常需要开发Excel文件的导入导出这样的功能，如果直接使用POI库，那么开发人员势必需要编写不少的代码进行Excel文件的读取并映射到相应的Java POJO对象上；反之，在进行数据导出到Excel文件中时，也面临着如何从POJO对象中取数据并写入到Excel上的问题，以及相应的Excel样式处理。虽然这一系列的工作难度不高，但十分繁琐，因此本项目的excelhelper目的就是在于简化这一切，方便开发者进行Excel文件的读写操作。   
## 提供的能力：  
+ 读取Excel文件并将数据映射到相应的POJO对象的相应字段上
+ 支持自定义表头，并提供多种默认样式
+ 将POJO对象的各字段的值写入到对应的Excel文件单元格中
## 库依赖：
+ org.apache.poi：3.17版本
+ com.alibaba.fastjson：1.2.54版本