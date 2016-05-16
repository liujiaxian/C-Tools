using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;

public class WebCommon
{
    public static void Show(string msg)
    {
        HttpContext.Current.Response.Write("<script type='text/javascript'>alert('" + msg + "')</script>");
    }
    public static void ShowUrl(string msg, string url)
    {
        HttpContext.Current.Response.Write("<script type='text/javascript'>alert('" + msg + "');window.location='" + url + "'</script>");
    }
    public static void ReLoad(string msg)
    {
        HttpContext.Current.Response.Write("<script type='text/javascript'>alert('" + msg + "');window.location.reload();</script>");
    }
    public static void Url(string url)
    {
        HttpContext.Current.Response.Write("<script type='text/javascript'>window.location='" + url + "'</script>");
    }
    public static string TitleLength(string strlength, int length)
    {
        if (strlength != null && strlength.Length > length)
        {
            strlength = strlength.Substring(0, length) + "......";
        }
        return strlength;
    }

    public static string ContentLength(String strlength, int length)
    {
        if (strlength != null && strlength.Length > length)
        {
            strlength = strlength.Substring(0, length) + "......";
        }
        return strlength;
    }
    public static bool CheckLogin()
    {
        if (HttpContext.Current.Session["admin"] == null)
        {
            HttpContext.Current.Response.Redirect("/ysw_admin/login.aspx");
            return false;
        }
        else
        {
            return true;
        }
    }
}