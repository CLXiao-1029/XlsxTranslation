# XlsxTranslation
该软件是由 [`Xlsx2Lua`](https://github.com/CLXiao-1029/Xlsx2Lua) 的实时翻译功能拆分而来的。
软件中提供了翻译转表功能，建议将软件放在同 `Xlsx2Lua` 软件的目录下。
软件会根据传入的参数 `--` 来决定是否读取配置，如果配置不存在时会自动创建。


# 支持库引用
| 支持库名 | 版本 | 项目地址 |
| ------- | ---- | ------- |
| `EPPLus` | 6.2.1 | https://epplussoftware.com/|
| `.Net` | 7.0.201 | https://dotnet.microsoft.com/en-us/download/dotnet/7.0|

# TranslateHelper 使用教程


读取配置XlsxTranslation.cfg
```bat
:: 读取配置XlsxTranslation.cfg
".\XlsxTranslation.exe" --
pause
```

根据传参进行配置
```bat
::程序名
::设置 Excel 目录                                         ..\Config
::设置 百度翻译API APPID                                  20230205001551232
::设置 百度翻译API SecretKey                              dkpFIs1HsuAIRBWPQ7ky
::设置 要读取的文件名                                      TranslateMain.xlsx
::设置 单条线程翻译的数量，默认500个                         500
::设置 单条线程启动的间隔，默认500毫秒                       100
::设置 线程执行一次准备翻译的间隔，默认500毫秒                500
::设置 线程执行一次翻译目标的间隔，默认2000毫秒               1000
::设置 显示耗时时间                                        true
::设置 显示日志的等级                                       0:All, 1:Warning+Error, 2:Error
::设置 是否输出日志文件                                     true
::设置 进度提示                                            true

REM 开始导表
".\XlsxTranslation.exe" ..\Config 20230205001551232 dkpFIs1HsuAIRBWPQ7ky TranslateMain.xlsx 500 100 500 1000 true 0 true true

pause
```

具体程序参数说明：

| 编号 |    类型    |          默认值           | 是否必填 | 释义                |
|:--:|:--------:|:----------------------:|:----:|-------------------|
| 1  | `string` |                        |  是   | 设置 Excel文件 目录     |
| 2  | `string` |  `20230205001551232`   |  是   | 百度翻译API APPID     |
| 3  | `string` | `dkpFIs1HsuAIRBWPQ7ky` |  是   | 百度翻译API SecretKey |
| 4  | `string` |  `TranslateMain.xlsx`  |  否   | 要读取/保存的文件名        |
| 5  |  `int`   |         `500`          |  否   | 分割翻译键的数量          |
| 6  |  `int`   |         `100`          |  否   | 启动线程的间隔(毫秒)       |
| 7  |  `int`   |         `1000`         |  否   | 线程执行一次Key的间隔(毫秒)  |
| 8  |  `int`   |         `2000`         |  否   | 线程执行一次Val的间隔(毫秒)  |
| 9  |  `int`   |         `2000`         |  否   | 显示耗时时间            |
| 10 |  `int`   |          `0`           |  否   | 日志显示等级            |
| 11 |  `bool`  |        `false`         |  否   | 是否输出日志文件          |
| 12 |  `bool`  |         `true`         |  否   | 是否开启翻译进度提示        |
---

# 程序配置文件说明
```json
{
  "ConfigPath": "..\\Config",//配置表路径。
  "MultilingualName": "TranslateMain.xlsx",//读取和输出文件名。
  "SplitKeyCount": 500,//单条线程翻译的数量，默认500个
  "ThreadInterval": 100,//单条线程启动的间隔，默认500毫秒
  "ThreadKeyInterval": 500,//线程执行一次准备翻译的间隔，默认500毫秒
  "ThreadValInterval": 1000,//线程执行一次翻译目标的间隔，默认2000毫秒
  "ShowLogLevel": 0,// 开启显示日志的等级。0：显示所有日志，1：只显示警告和报错，2：只显示报错。
  "OutputLogFile": true,// 是否输出日志文件？在导表结束后会在程序执行目录生成log文件，文件名：程序名+当前时间，示例：XlsxTranslation-20231219151955.log
  "TipProgress": true,//是否开启翻译进度提示
  "ShowTimelapse":true,//显示时间流逝
  "LanguageData": {
    "AppId": "20230205001551232",//百度翻译API APPID
    "SecretKey": "dkpFIs1HsuAIRBWPQ7ky",//百度翻译API SecretKey
    "Translations": [
      "en"
    ],// 翻译文件的语种
    "TranslationNames": [
      "英语"
    ]// 翻译文件的语种中文释义
  }
}
```

# Excel表结构
#### 结构说明
1. 读取第一个工作簿（sheet）中的第一列作为翻译源，后几列作为目标源
2. 翻译表的文件名以`Translate`开头，默认的翻译表名`TranslateMain`

#### 示例文件：

`TestExample\总表模式\Config\TranslateMain.xlsx` [打开](https://github.com/CLXiao-1029/XlsxTranslation/blob/master/TestExample/%E5%88%86%E8%A1%A8%E6%A8%A1%E5%BC%8F/Config/TranslateMain.xlsx)

`TestExample\分表模式\Config\TranslateMain.xlsx` [打开](https://github.com/CLXiao-1029/XlsxTranslation/blob/master/TestExample/%E6%80%BB%E8%A1%A8%E6%A8%A1%E5%BC%8F/Config/TranslateMain.xlsx)

`TestExample\翻译演示\Config\TranslateMain.xlsx` [打开](https://github.com/CLXiao-1029/XlsxTranslation/blob/1.0.0/TestExample/%E7%BF%BB%E8%AF%91%E6%BC%94%E7%A4%BA/Config/TranslateMain.xlsx)


|    多语言Key    |  中文   | 英语 | 繁体中文 |
|:------------:|:-----:|:--:|:----:|
|     key      |  zh   | en | cht  |
|              |       |    |
| Items_name_1 |  金币   |
| Items_desc_1 | 通用货币1 |
| Items_name_2 |  钻石   |
| Items_desc_2 | 通用货币2 |
