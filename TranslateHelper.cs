
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace XlsxTranslation;

using static AppConfig;
internal class TranslateHelper
{
    // 翻译AI的APP ID
    static string APPID = "";
    static string APPKEY = "";

    public static void Init()
    {
        APPID = AppData.LanguageData.AppId;
        APPKEY = AppData.LanguageData.SecretKey;
    }

    /// <summary>
    /// 获取翻译结果
    /// </summary>
    /// <param name="content">要翻译的原文</param>
    /// <param name="to">目标语言</param>
    /// <param name="from">源语言</param>
    public static string Get(string content, string to, string from = "zh")
    {
        Random rd = new Random();
        string salt = rd.Next(100000).ToString();
        string sign = EncryptString(APPID + content + salt + APPKEY);
        string url = "http://api.fanyi.baidu.com/api/trans/vip/translate?";
        url += "q=" + HttpUtility.UrlEncode(content);
        url += "&from=" + from;
        url += "&to=" + to;
        url += "&appid=" + APPID;
        url += "&salt=" + salt;
        url += "&sign=" + sign;
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
        request.Method = "GET";
        request.ContentType = "text/html;charset=UTF-8";
        request.UserAgent = null;
        request.Timeout = 6000;
        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
        Stream myResponseStream = response.GetResponseStream();
        StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
        string retString = myStreamReader.ReadToEnd();
        myStreamReader.Close();
        myResponseStream.Close();
        return retString;
    }

    // 计算MD5值
    private static string EncryptString(string str)
    {
        MD5 md5 = MD5.Create();
        // 将字符串转换成字节数组
        byte[] byteOld = Encoding.UTF8.GetBytes(str);
        // 调用加密方法
        byte[] byteNew = md5.ComputeHash(byteOld);
        // 将加密结果转换为字符串
        StringBuilder sb = new StringBuilder();
        foreach (byte b in byteNew)
        {
            // 将字节转换成16进制表示的字符串，
            sb.Append(b.ToString("x2"));
        }

        // 返回加密的字符串
        return sb.ToString();
    }

    public static string GetErrorCode(string errorCode, object msg)
    {
        string msgZh = "";
        if (errorCode == "52000")
            msgZh = "成功";
        if (errorCode == "52001")
            msgZh = "请求超时，请重试";
        if (errorCode == "52002")
            msgZh = "系统错误，请重试";
        if (errorCode == "52003")
            msgZh = "未授权用户，请检查appid是否正确或者服务是否开通";
        if (errorCode == "54000")
            msgZh = "必填参数为空，请检查是否少传参数 ";
        if (errorCode == "54001")
            msgZh = "签名错误，请检查您的签名生成方法 ";
        if (errorCode == "54003")
            msgZh = "访问频率受限，请降低您的调用频率，或进行身份认证后切换为高级版/尊享版 https://fanyi-api.baidu.com/api/trans/product/desktop";
        if (errorCode == "54004")
            msgZh = "账户余额不足，请前往管理控制台为账户充值 https://fanyi-api.baidu.com/api/trans/product/desktop";
        if (errorCode == "54005")
            msgZh = "长query请求频繁，请降低长query的发送频率，3s后再试 ";
        if (errorCode == "58000")
            msgZh = "客户端IP非法，检查个人资料里填写的IP地址是否正确，可前往开发者信息-基本信息修改 ";
        if (errorCode == "58001")
            msgZh = "译文语言方向不支持，检查译文语言是否在语言列表里";
        if (errorCode == "58002")
            msgZh = "服务当前已关闭，请前往管理控制台开启服务 https://fanyi-api.baidu.com/api/trans/product/desktop";
        if (errorCode == "90107")
            msgZh = "认证未通过或未生效 请前往我的认证查看认证进度 https://api.fanyi.baidu.com/myIdentify";
        return $"错误码:{errorCode} 错误信息:{msg} 错误码含义:{msgZh}";
    }
}