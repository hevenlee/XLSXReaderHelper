# XLSXReaderHelper
读取大数据量的excel2007文件

需要导入的包：
org.apache.xerces.parsers.SAXParser

Maven依赖:

groupId:xerces

artifactId:xercesImpl

version:2.11.0

使用方法：
1、继承该类，并实现fetchRows()方法，该方法会在读取完一行后，返回读取结果。
2、调用fetchOneSheet()方法进行读取操作。
