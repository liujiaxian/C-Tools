using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Security.Cryptography;
using System.Text;

/// <summary>
///MD5Endcoding 的摘要说明
/// </summary>
public static class MD5Endcoding
{
	public static string Pwd(string msg)
	{
		//
		//TODO: 在此处添加构造函数逻辑
		//
        MD5 md5 = MD5.Create();
        byte[] buffer = System.Text.Encoding.UTF8.GetBytes(msg);
        byte[] Buffer = md5.ComputeHash(buffer);
        StringBuilder bs = new StringBuilder();
        for (int i = 0; i < Buffer.Length; i++)
        {
            bs.Append(Buffer[i].ToString("x2"));
        }
        return bs.ToString();
	}
}