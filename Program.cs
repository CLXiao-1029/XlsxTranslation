using OfficeOpenXml;

namespace XlsxTranslation
{
    internal class Program
    {
        private static bool _isClose = false;
        
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += CurrentDomain_ProcessExit;
            
            Console.OutputEncoding = System.Text.Encoding.UTF8;
#if NET5_0_OR_GREATER
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //指明非商业应用

            try
            {
                // args = new[] { "--" };
                InitArgs(args);
                
                DateTime dateTimeAll = DateTime.Now;
                // 获取导表后的文件
                AppConfig.InitTranslation();
                // 获取TranslateMain相关附表 Translate*
                AppConfig.InitFinalTranslate();
                
                Logger.LogWarning($"分析翻译文件耗时：{(DateTime.Now - dateTimeAll).TotalMilliseconds} 毫秒");
                
                // 开始翻译
                AppConfig.ToTranslate();
                // 写出文件
                AppConfig.AsSaveXlsxData();
                _isClose = true;
                
                AppConfig.SaveLogFile();
            }
            catch (Exception ex)
            {
                Logger.LogError($"意外关闭，已经翻译的部分文件: TranslateTempMain . . .");
                AppConfig.AsSaveXlsxData("TranslateTempMain");
                Logger.LogException(ex);
                _isClose = true;
                AppConfig.SaveLogFile();
            }
        }

        static void InitArgs(string[] args)
        {
            AppConfig.InitData();
            string val = args[0];
            // 初始化配置
            AppConfig.LoadConfigurationData();
            if (val.StartsWith("--"))
            {
                Logger.LogInfo($"读取XlsxTranslation.cfg");
            }
            else
            {
                if (args.Length < 1)
                {
                    Logger.LogErrorAndExit("未输入Excel表格所在目录");
                    return;
                }
                // 刷新配置数据
                AppConfig.RefreshAppData(args);
                // 保存最新的配置数据
                AppConfig.SaveConfigurationData();
            }
            // 初始化翻译配置
            TranslateHelper.Init();
        }

        private static void CurrentDomain_ProcessExit(object? sender, EventArgs e)
        {
            if (!_isClose)
            {
                Logger.LogError($"意外关闭，已经翻译的部分文件: TranslateTempMain . . .");
                AppConfig.AsSaveXlsxData("TranslateTempMain");
                AppConfig.SaveLogFile();
            }
        }
    }
}