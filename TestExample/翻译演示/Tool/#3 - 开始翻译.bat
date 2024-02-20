::程序名
::设置 Excel 目录											..\Config
::设置 百度翻译API APPID									20230205001551232
::设置 百度翻译API SecretKey								dkpFIs1HsuAIRBWPQ7ky
::设置 要读取的文件名，默认TranslateMain.xlsx				TranslateMain.xlsx
::设置 单条线程翻译的数量，默认500个						500
::设置 单条线程启动的间隔，默认500毫秒						100
::设置 线程执行一次准备翻译的间隔，默认500毫秒		 		500
::设置 线程执行一次翻译目标的间隔，默认2000毫秒				1000
::设置 显示耗时时间											true
::设置 显示日志的等级										0:All, 1:Warning+Error, 2:Error
::设置 是否输出日志文件										true
::设置 进度提示												true

REM 开始导表
".\XlsxTranslation.exe" ..\Config 20230205001551232 dkpFIs1HsuAIRBWPQ7ky TranslateMain.xlsx 500 100 500 1000 true 0 true true

pause