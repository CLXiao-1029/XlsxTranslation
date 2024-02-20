using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XlsxTranslation;

/// <summary>
/// 翻译工具的配置结构
/// </summary>
public class LanguageData
{
    public string AppId { get; set; }
    public string SecretKey { get; set; }
    public string From { get; set; }
    public string[] Translations { get; set; }
    public string[] TranslationNames { get; set; }

    public LanguageData()
    {
        AppId = "";
        SecretKey = "";
        From = "zh";
        Translations = new string[] { "en" };
        TranslationNames = new string[] { "英语" };
    }
}

public struct AppDataStruct
{
    /// <summary>
    /// 配置表路径
    /// </summary>
    public string? ConfigPath { get; set; }

    /// <summary>
    /// 多语言文件名
    /// </summary>
    public string? MultilingualName { get; set; }

    /// <summary>
    /// 分割翻译键的数量，默认500一组
    /// </summary>
    public int SplitKeyCount { get; set; }

    /// <summary>
    /// 线程间隔时间
    /// </summary>
    public int ThreadInterval { get; set; }

    /// <summary>
    /// 线程执行一次Key的间隔
    /// </summary>
    public int ThreadKeyInterval { get; set; }

    /// <summary>
    /// 线程执行一次Val的间隔
    /// </summary>
    public int ThreadValInterval { get; set; }

    /// <summary>
    /// 显示时间流逝
    /// </summary>
    public bool ShowTimelapse { get; set; }

    /// <summary>
    /// 显示日志等级
    /// </summary>
    public long ShowLogLevel { get; set; }

    /// <summary>
    /// 是否输出日志文件
    /// </summary>
    public bool OutputLogFile { get; set; }

    /// <summary>
    /// 进度提示
    /// </summary>
    public bool TipProgress { get; set; }

    public LanguageData LanguageData { get; set; }

    public AppDataStruct()
    {
        Default();
    }

    private void Default()
    {
        MultilingualName = "TranslateMain.xlsx";
        ConfigPath = "";
        SplitKeyCount = 500;
        ThreadInterval = 500;
        ThreadKeyInterval = 500;
        ThreadValInterval = 2000;
        ShowTimelapse = true;
        TipProgress = true;
        ShowLogLevel = 1;
        OutputLogFile = false;
        LanguageData = new LanguageData()
        {
            AppId = "20230205001551232",
            SecretKey = "dkpFIs1HsuAIRBWPQ7ky"
        };
    }
}

/// <summary>
/// 需要翻译的多语言数据结构
/// </summary>
public struct MultilingualData
{
    public string Key { get; set; }
    public Dictionary<string, object> Values { get; set; }

    public void AddValue(string key, object obj)
    {
        if (!Values.ContainsKey(key))
            Values.Add(key, obj);
        else
            ReplaceValue(key, obj);
    }

    public object? GetValue(string key)
    {
        if (Values.TryGetValue(key, out var data))
        {
            return data;
        }

        return null;
    }

    public void ReplaceValue(string key, object obj)
    {
        Values[key] = obj;
    }

    public bool Compare(string key, object target)
    {
        if (Values.TryGetValue(key, out var data))
        {
            return data == target;
        }

        return false;
    }
}

public static class AppConfig
{
    // 用于缩进lua table的字符串
    private static readonly string LUA_TABLE_INDENTATION_STRING = "\t";
    private static readonly string ConfigName = "XlsxTranslation.cfg";
    private static readonly string AppDirectory = AppDomain.CurrentDomain.BaseDirectory;

    /// <summary>
    /// 需要翻译的多语言配置数据
    /// 主键是中文描述
    /// 值对应的当前行中所有需要翻译的多语言语种和结果
    /// </summary>
    public static Dictionary<string, MultilingualData> MulLanguageDataDic;

    /// <summary>
    /// 缓存数据
    /// </summary>
    public static Dictionary<string, List<string>> MultilingualDataTmp;

    /// <summary>
    /// 存放翻译表第一列和第二列数据
    /// 主键是第一列数据，多语言的key
    /// 值是第二列数据，key对应的中文描述
    /// </summary>
    public static Dictionary<object, object> KeyValueData;

    public static AppDataStruct AppData;

    /// <summary>
    /// 翻译表的有效起始行
    /// </summary>
    public static int XlsxStartingRow = 4;
    
    private static float _maxCnt = 0;

    public static bool AppIdIsNull()
    {
        bool isNull = string.IsNullOrEmpty(AppData.LanguageData.AppId) ||
                      string.IsNullOrWhiteSpace(AppData.LanguageData.AppId);
        return isNull;
    }

    public static bool CellIsNull(object? data)
    {
        bool isNull = data == null;
        if (!isNull)
        {
            if (ExcelErrorValue.Values.IsErrorValue(data))
            {
                isNull = true;
            }
        }

        return isNull;
    }

    #region 配置数据

    public static void InitData()
    {
        MulLanguageDataDic = new Dictionary<string, MultilingualData>();
        MultilingualDataTmp = new Dictionary<string, List<string>>();
        KeyValueData = new Dictionary<object, object>();
        AppData = new AppDataStruct();
    }

    /// <summary>
    /// 加载配置数据
    /// </summary>
    public static void LoadConfigurationData()
    {
        string fileName = Path.Combine(AppDirectory, ConfigName);
        if (!File.Exists(fileName))
        {
            FileUtils.SaveJson(fileName, AppData);
        }

        string var = File.ReadAllText(fileName);
        AppData = JsonSerializer.Deserialize<AppDataStruct>(var);

        // 初始化AppID
        if (AppIdIsNull())
        {
            AppData.LanguageData = new LanguageData()
            {
                AppId = "20230205001551232",
                SecretKey = "dkpFIs1HsuAIRBWPQ7ky"
            };
        }
    }

    /// <summary>
    /// 刷新配置数据
    /// </summary>
    /// <param name="args"></param>
    public static void RefreshAppData(string[] args)
    {
        // 路径
        AppData.ConfigPath = args[0];

        if (args.Length > 1)
        {
            AppData.LanguageData.AppId = args[1];
        }

        if (args.Length > 2)
        {
            AppData.LanguageData.SecretKey = args[2];
        }

        // 文件名
        if (args.Length > 3)
        {
            AppData.MultilingualName = args[3];
        }

        // 分割翻译键的数量
        if (args.Length > 4)
        {
            AppData.SplitKeyCount = int.Parse(args[4]);
        }

        if (args.Length > 5)
        {
            AppData.ThreadInterval = int.Parse(args[5]);
        }

        if (args.Length > 6)
        {
            AppData.ThreadKeyInterval = int.Parse(args[6]);
        }

        if (args.Length > 7)
        {
            AppData.ThreadValInterval = int.Parse(args[7]);
        }

        if (args.Length > 8)
        {
            AppData.ShowTimelapse = bool.Parse(args[8]);
        }

        if (args.Length > 9)
        {
            AppData.ShowLogLevel = int.Parse(args[9]);
        }

        if (args.Length > 10)
        {
            AppData.OutputLogFile = bool.Parse(args[10]);
        }

        if (args.Length > 11)
        {
            AppData.TipProgress = bool.Parse(args[11]);
        }
    }

    public static void SaveConfigurationData()
    {
        string cfgPath = Path.Combine(AppDirectory, ConfigName);
        if (File.Exists(cfgPath))
        {
            File.Delete(cfgPath);
        }

        FileUtils.SaveJson(cfgPath, AppData);
        Logger.LogWarning($"保存配置文件：{cfgPath}");
    }

    #endregion

    public static void SaveLogFile()
    {
        if (AppData.OutputLogFile)
        {
            string logPath = Path.Combine(AppDirectory, "log",
                $"{Assembly.GetExecutingAssembly().GetName().Name}-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
            FileUtils.Save(logPath, Logger.LogToString());
            Logger.LogWarning(logPath);
        }
    }

    /// <summary>
    /// 获取文件名
    /// </summary>
    /// <returns></returns>
    public static string GetFileName()
    {
        string fileName = AppData.MultilingualName ?? "TranslateMain.xlsx";

        bool isNull = string.IsNullOrEmpty(fileName) ||
                      string.IsNullOrWhiteSpace(fileName);
        if (!isNull)
        {
            var fileExtension = Path.GetExtension(fileName);
            if (string.IsNullOrEmpty(fileExtension))
            {
                fileName = $"{fileName}.xlsx";
            }
        }

        return fileName;
    }

    /// <summary>
    /// 获取文件完整路径
    /// </summary>
    /// <returns></returns>
    public static string GetFileFullPath(string? fileName = null)
    {
        fileName = fileName ?? GetFileName();
        var fileExtension = Path.GetExtension(fileName);
        if (string.IsNullOrEmpty(fileExtension))
        {
            fileName = $"{fileName}.xlsx";
        }
        string path = Path.Combine(AppData.ConfigPath, fileName);
        return path;
    }

    #region 初始化翻译数据

    /// <summary>
    /// 初始化翻译数据
    /// </summary>
    public static void InitTranslation()
    {
        MulLanguageDataDic.Clear();
        MultilingualDataTmp.Clear();
        KeyValueData.Clear();
        string path = GetFileFullPath();
        Logger.LogInfo(path);
        Logger.LogInfo(GetFileName());
        if (File.Exists(path))
        {
            LoadXlsxData(path);
        }
    }

    /// <summary>
    /// 初始化最终的翻译文件
    /// </summary>
    public static void InitFinalTranslate()
    {
        DirectoryInfo root = new DirectoryInfo(AppData.ConfigPath);
        foreach (FileInfo file in root.GetFiles())
        {
            if (file.Name.StartsWith("Translate") && !file.Name.Equals(AppData.MultilingualName))
            {
                LoadXlsxData(file.FullName, true);
            }
        }
    }

    private static void LoadXlsxData(string filePath, bool isFinal = false)
    {
        using (FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            string fileName = GetFileName();
            ExcelPackage package = new ExcelPackage(fileStream);
            // 获取第一个工作簿数据
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            // 打印工作簿名字
            Logger.LogInfo($"成功读取[{fileName}]表的【{worksheet.Name}】工作簿");
            // 获取有效行数
            int rows = worksheet.Dimension.Rows;
            // 获取有效列数
            int columns = worksheet.Dimension.Columns;

            // 遍历行
            for (int rowIdx = XlsxStartingRow; rowIdx < rows + 1; ++rowIdx)
            {
                // 获取前两列的数据
                var rowData1 = worksheet.Cells[rowIdx, 1].Value;
                var mainRowData = worksheet.Cells[rowIdx, 2].Value;
                // 加入主键数据，正常情况下不会出现两个一样的key
                if (KeyValueData.ContainsKey(rowData1))
                {
                    // 更新主键数据
                    KeyValueData[rowData1] = mainRowData;
                }
                else
                {
                    KeyValueData.TryAdd(rowData1, mainRowData);
                }

                // 如果中文描述为空则跳过这一行
                bool isNull = CellIsNull(mainRowData);
                if (isNull)
                {
                    // 跳过这一列的空行
                    Logger.LogError($"表名：{worksheet.Name} ,第{rowIdx}行1列单元格值为空或错误：{rowData1}");
                    continue;
                }

                string mainKey = mainRowData.ToString() ?? "";
                Logger.LogWarning($"检查：{mainKey}");
                if (string.IsNullOrEmpty(mainKey))
                {
                    // 跳过这一列的空行
                    Logger.LogError($"表名：{worksheet.Name} ,第{rowIdx}行1列单元格值为空或错误：{rowData1}");
                    continue;
                }

                List<string> values = new List<string>();

                #region 最终版数据对比

                if (isFinal)
                {
                    if (MultilingualDataTmp.ContainsKey(mainKey))
                    {
                        List<string> dataTmp = MultilingualDataTmp[mainKey];
                        // 从第二列开始遍历
                        for (int col = 2; col < columns + 1; ++col)
                        {
                            // 获取当前列的语种
                            var langColData = worksheet.Cells[2, col].Value;
                            isNull = CellIsNull(langColData);
                            if (isNull)
                            {
                                // 跳过空的错误的
                                Logger.LogError($"已跳过第2行第{col}列，单元格为空或错误：{langColData}");
                                continue;
                            }

                            //如果之前的查找中存在要翻译的缓存key，并且最终版的当前列不是空，就移除
                            string mulKey = langColData.ToString() ?? "";
                            if (string.IsNullOrEmpty(mulKey))
                            {
                                // 跳过空的错误的
                                Logger.LogError($"已跳过第2行第{col}列，单元格为空或错误：{langColData}");
                                continue;
                            }

                            // 删除已经存在的
                            if (dataTmp.Contains(mulKey))
                            {
                                MultilingualDataTmp[mainKey].Remove(mulKey);
                            }
                        }

                        //如果已经移除完毕了则清理当前key
                        if (MultilingualDataTmp[mainKey].Count == 0)
                        {
                            MultilingualDataTmp.Remove(mainKey);
                        }
                    }
                }

                #endregion

                if (MulLanguageDataDic.TryGetValue(mainKey,out MultilingualData multilingualData))
                {
                    // 从第二列开始遍历
                    for (int colIdx = 2; colIdx < columns + 1; ++colIdx)
                    {
                        // 获取当前列的语种
                        var langColData = worksheet.Cells[2, colIdx].Value;
                        isNull = CellIsNull(langColData);
                        if (isNull)
                        {
                            // 跳过空的错误的
                            Logger.LogError($"已跳过第2行第{colIdx}列，单元格为空或错误：{langColData}");
                            continue;
                        }

                        string mulKey = langColData.ToString() ?? "";
                        if (string.IsNullOrEmpty(mulKey))
                        {
                            // 跳过空的错误的
                            Logger.LogError($"已跳过第2行第{colIdx}列，单元格为空或错误：{langColData}");
                            continue;
                        }
                        Logger.LogWarning($"检查：{mainKey} 语种:{mulKey} 第{colIdx}列 target:{langColData} isFinal:{isFinal}");

                        //如果当前行列不存在翻译，则加入翻译列表
                        if (!isNull)
                        {
                            // 如果
                            if (multilingualData.Compare(mulKey, langColData))
                            {
                                continue;
                            }
                            if (isFinal)//如果俩个值不相等并且是最终版本，则以最终版本为准
                            {
                                multilingualData.ReplaceValue(mulKey, langColData);
                            }
                        }
                    }
                }
                else
                {
                    //存放多语言数据
                    MultilingualData rowMulData = new MultilingualData()
                    {
                        Key = mainKey,
                        Values = new Dictionary<string, object>()
                    };
                    // 从第二列开始遍历
                    for (int colIdx = 2; colIdx < columns + 1; ++colIdx)
                    {
                        // 获取当前列的语种
                        var langColData = worksheet.Cells[2, colIdx].Value;
                        isNull = CellIsNull(langColData);
                        if (isNull)
                        {
                            // 跳过空的错误的
                            Logger.LogError($"已跳过第2行第{colIdx}列，单元格为空或错误：{langColData}");
                            continue;
                        }

                        string mulKey = langColData.ToString() ?? "";
                        if (string.IsNullOrEmpty(mulKey))
                        {
                            // 跳过空的错误的
                            Logger.LogError($"已跳过第2行第{colIdx}列，单元格为空或错误：{langColData}");
                            continue;
                        }
                        var colData = worksheet.Cells[rowIdx, colIdx].Value;
                        isNull = CellIsNull(colData);
                        //如果当前行当前列不存在翻译，则加入翻译列表
                        if (isNull)
                        {
                            Logger.LogWarning($"{rowMulData.Key}不存在目标语言[{mulKey}]");
                            values.Add(mulKey);
                        }
                        else//首次需要加入多语言数据
                        {
                            rowMulData.AddValue(mulKey, colData);
                        }
                    }

                    MultilingualDataTmp.Add(rowMulData.Key, values);
                    MulLanguageDataDic.Add(rowMulData.Key, rowMulData);
                }
            }
        }
    }

    public static void ToTranslate()
    {
        _maxCnt = MultilingualDataTmp.Count;
        DateTime dateTime = DateTime.Now;
        if (_maxCnt > 0)
        {
            string[] keys = MultilingualDataTmp.Keys.ToArray();
            int itemCnt = (int)MathF.Floor(keys.Length / AppData.SplitKeyCount);

            if (itemCnt > 1)
            {
                // 分割
                List<string[]> list = SplitArray(keys, AppData.SplitKeyCount);

                Logger.LogWarning($"申请线程数：{list.Count}");

                foreach (string[] subKey in list)
                {
                    //创建线程，命名为thread
                    Thread thread = new Thread(new ParameterizedThreadStart(MultilingualTranslate)); //记得要在new 创建的时候带上需要执行的方法名
                    thread.IsBackground = true; //这个是把我们这个线程放到后台线程里面执行
                    thread.Start(subKey); //Start就代表开始执行线程
                    Thread.Sleep(AppData.ThreadInterval);
                }
            }
            else
            {
                MultilingualTranslate(keys);
            }
        }
        
        while (MultilingualDataTmp.Count > 0) { }
        Logger.LogWarning($"翻译耗时：{(DateTime.Now - dateTime).TotalMilliseconds}");
        Logger.LogInfo($"导出翻译文件{AppData.MultilingualName} . . .");
    }

    private static void MultilingualTranslate(object keys)
    {
        string[] subKeys = (string[])keys;
        for (int i = subKeys.Length - 1; i >= 0; i--)
        {
            var key = subKeys[i];
            var values = MultilingualDataTmp[key];
            for (int j = values.Count - 1; j >= 0; j--)
            {
                var to = values[j];
                TranslateParse(key, to);
                Thread.Sleep(AppData.ThreadValInterval);
            }

            MultilingualDataTmp.Remove(key);
            Thread.Sleep(AppData.ThreadKeyInterval);
        }
    }

    /// <summary>
    /// 翻译转换
    /// </summary>
    /// <param name="key"></param>
    /// <param name="to"></param>
    private static void TranslateParse(string key, string to)
    {
        var json = TranslateHelper.Get(key, to);
        var jsonNode = JsonNode.Parse(json);
        if (jsonNode != null)
        {
            var errorCode = jsonNode["error_code"];
            if (errorCode != null)
            {
                var errorMsg = jsonNode["error_msg"];
                var err = TranslateHelper.GetErrorCode(errorCode.ToString(), errorMsg);
                Logger.LogError($"翻译内容：{key}，目标语言：{to} {err}");
            }
            else
            {
                var result = jsonNode["trans_result"];
                var dst = result?[0]?["dst"];
                var info = AppData.TipProgress ? $"进度：{_maxCnt - MultilingualDataTmp.Count} / {_maxCnt} {((_maxCnt - MultilingualDataTmp.Count) / _maxCnt) * 100}%":"";
                Logger.LogInfo($"翻译内容：{key}，目标语言：{to} {info}");
                //尝试获取对应翻译内容
                if (MulLanguageDataDic.TryGetValue(key, out var multilingualData))
                {
                    multilingualData.AddValue(to, dst);
                }
                else //如果不存在则添加新翻译项
                {
                    MultilingualData data = new MultilingualData()
                    {
                        Key = key,
                        Values = new Dictionary<string, object>()
                    };
                    data.Values.Add(to, dst);
                    MulLanguageDataDic.Add(key, data);
                }
            }
        }
        else
        {
            Logger.LogError($"错误，无法转换Json数据：{json}");
        }
    }
    
    private static List<T[]> SplitArray<T>(T[] ary, int subSize)
    {
        int count = ary.Length % subSize == 0 ? ary.Length / subSize : ary.Length / subSize + 1;
        List<T[]> subAryList = new List<T[]>();
        for (int i = 0; i < count; i++)
        {
            int index = i * subSize;
            T[] subAry = ary.Skip(index).Take(subSize).ToArray();
            subAryList.Add(subAry);
        }

        return subAryList;
    }
    #endregion

    #region 写出Xlsx配置

    public static void AsSaveXlsxData(string? fileName = null)
    {
        // fileName = fileName ?? GetFileName();
        string path = GetFileFullPath(fileName);
        if (File.Exists(path))
        {
            File.Delete(path);
        }

        using (FileStream fileStream = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            ExcelPackage package = new ExcelPackage(fileStream);
            
            if (package.Workbook.Worksheets.Count == 0)
            {
                package.Workbook.Worksheets.Add("翻译表");
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
            // 添加表头
            worksheet.Cells[1, 1].Value = "多语言Key";
            worksheet.Cells[2, 1].Value = "key";
            worksheet.Cells[1, 2].Value = "中文";
            worksheet.Cells[2, 2].Value = "zh";
            var languageData = AppData.LanguageData;
            for (int i = 0; i < languageData.TranslationNames.Length; i++)
            {
                worksheet.Cells[1, i + 3].Value = languageData.TranslationNames[i];
                worksheet.Cells[2, i + 3].Value = languageData.Translations[i];
            }
            Logger.LogInfo($"准备写出Excel Sheet【{worksheet.Name}】");
            // 插入前两列
            var keyValueKeys = KeyValueData.Keys.ToArray();
            for (int i = 0; i < keyValueKeys.Length; i++)
            {
                var key = keyValueKeys[i];
                var value = KeyValueData[key];
                worksheet.Cells[i + XlsxStartingRow, 1].Value = key;
                worksheet.Cells[i + XlsxStartingRow, 2].Value = value;
            }

            // 插入翻译后的内容
            var keys = MulLanguageDataDic.Keys.ToArray();
            var transLen = languageData.Translations.Length;
            for (int i = 0; i < keys.Length; i++)
            {
                string key = keys[i];
                MultilingualData data = MulLanguageDataDic[key];

                Logger.LogInfo($"准备写出Excel Sheet：{worksheet.Name} - {key}");
                for (int j = 0; j < transLen; j++)
                {
                    var transKey = languageData.Translations[j];
                    if (data.Values.TryGetValue(transKey,out var value))
                    {
                        worksheet.Cells[i + XlsxStartingRow, j + 3].Value = value;
                    }
                }
            }
            
            // 绘制表头
            for (int rowIdx = 1; rowIdx < XlsxStartingRow; ++rowIdx)
            {
                worksheet.Row(rowIdx).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(rowIdx).Style.Fill.BackgroundColor.SetColor(255, 112, 173, 71);
            }
            
            // 保存操作
            package.Save();
        }
    }

    #endregion
}