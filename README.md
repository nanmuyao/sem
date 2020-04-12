# sem
1 generate app
pyinstaller -F handle.py stat_service.py


sem
#TODO
1.分析多个源文件
2.对于csv文件自动转成XLSX格式
3.exclude 要推荐一些删除的词
4.尝试处理原始csv格式(每次转格式太费劲了)

2019_05_24
1.文件夹下所有csv自动转换成xlsx格式
2.把文件夹下所有的xlsx文件全部进行分析

2019_05-23
支持价格的导出

2019-05-04
2.完成

2019-05-04
第一个版本可以分析sem导出文件完成

2020-03-28
加一些提示，优化log

2020-04-12
直接处理csv格式文件，现在写入状态有些问题，因为filedname会重复，pending