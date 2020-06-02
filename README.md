

## TBM结果处理工具

### 使用说明

#### 环境依赖

- 第三方处理库安装 openpyxl
- Python2.6 及以上环境

```python
# 安装 openpyxl
pip install openpyxl==2.4.8
```

#### 目录结构介绍

```python
│  tbmtools-2@0601.py    #脚本处理程序
│
├─data                   #存放映射关系及表头配置 文件件
│      data.json        #存放映射关系
│      title.conf       #存放导出文件表头信息的配置，字体大小，背景，颜色等
│
├─reports               #导出文件的默认路径
│      -1CentOS_7_4_MOA操作系统#000001_10.14.31.12.xlsx  
│      CentOS_7_4_MOA操作系统#000001_10.14.31.12.xlsx
│      MySql_5_1_南基SMAP数据库#000001_172.16.112.74.xlsx
│
└─templates            # 存放映射关系的data.json数据来自这里的模板
        Apache2.2.xlsx
        apache2.4.xlsx
        centos6.xlsx
        centos7.xlsx
        mysql.xlsx
        mysql5.xlsx
        Oracle11g.xlsx
        redhat6.xlsx
        soliras.xlsx
        tomcat7.0.xlsx
        tomcat8.0.xlsx
        tomcat8.5.xlsx
        tongweb.xlsx
        windows2012.xlsx
```


#### 问题1 模板更新了，如何进行模板更新映射关系？

- 参数 -m ，根据templates中的模板进行重新生成data.json

```python
# 生成 映射关系
python tbmtools.py -m
```

#### 问题2 如何运行程序进行处理文件？

```python

python tbmtools.py -i 存放待处理的excel文件夹路径

```


#### 问题3 输出路径是否可以指定？

**输出路径可以指定,使用 -o 参数指定导出路径即可**

```python
python tbmtools.py -i inputpath  -o outputpath

```


#### 其他问题

```cmd
> python tbmtools.py -h
usage: tbmtools-2@0601.py [-h] [-m] [-t [TEMPLATES]] [-o [OUTPUT]]
                          [-i [INPUT]]

TBM Data fill tool! Help:wang_di@topsec.com.cn Usage: python tbmtools.py -h

optional arguments:
  -h, --help            show this help message and exit
  -m, --make            根据模板文件重新生成映射关系文件 data.json
  -t [TEMPLATES], -templates [TEMPLATES]
                        指定生成data.json映射关系的xlsx模板路径,data.json中的映射关系由该目录中的文件生成,默
                        认路径为./templates
  -o [OUTPUT], --output [OUTPUT]
                        指定导出文件的路径，默认值为 ./reports/
  -i [INPUT], --input [INPUT]
                        指定需要处理的多个文件路径

```