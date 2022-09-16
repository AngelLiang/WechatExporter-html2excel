# WechatExporter-html2excel

基于 WechatExporter 导出的微信聊天记录，HTML转换Excel工具

使用 https://github.com/BlueMatthew/WechatExporter 工具导出HTML格式的聊天记录，然后转换为Excel格式，便于分析

## 环境

- Windows 10
- Python 38+

## 使用说明


1、由于导出的HTML是动态页面，需要下拉滚动才能加载后面的数据，所以第一步需要先把导出的HTML另存为完整的HTML文件，然后放到 input_data 文件夹里

2、执行下面代码，即可在output文件夹里看到转换后的Excel文件

    poetry shell
    python main.py
    

## 格式说明

转换后的Excel格式如下

| <时间>              | <昵称> | <聊天内容> |
| ------------------- | ------ | ---------- |
| 2021-06-18 20:38:03 | 昵称   | 聊天内容   |

## 注意事项

- 自己的发言（聊天框右侧）暂时不会导出
- 图片无法导出到Excel

## changelog

- 20220916：初稿
