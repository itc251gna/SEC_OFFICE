<%@ WebHandler Language="C#" Class="SecOfficeMedicoBridge" %>

using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public class SecOfficeMedicoBridge : IHttpHandler
{
    private const string ConnectionStringName = "ConnectionString2";

    private static readonly HashSet<string> AllowedRemoteAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "127.0.0.1",
        "::1",
        "10.4.51.232"
    };

    public bool IsReusable
    {
        get { return true; }
    }

    public void ProcessRequest(HttpContext context)
    {
        context.Response.ContentType = "application/json; charset=utf-8";
        context.Response.Cache.SetCacheability(HttpCacheability.NoCache);

        string remoteAddress = NormalizeRemoteAddress(context.Request.UserHostAddress);
        if (!AllowedRemoteAddresses.Contains(remoteAddress))
        {
            context.Response.StatusCode = 403;
            WriteError(context, "Forbidden remote address.");
            return;
        }

        string kind = (context.Request.QueryString["kind"] ?? "patients").Trim().ToLowerInvariant();
        string sql = QueryForKind(kind);
        if (sql == null)
        {
            context.Response.StatusCode = 400;
            WriteError(context, "Unknown MEDICO query kind.");
            return;
        }

        try
        {
            List<Dictionary<string, object>> rows = ExecuteRows(sql);
            context.Response.Write("{\"status\":\"OK\",\"kind\":");
            context.Response.Write(JsonString(kind));
            context.Response.Write(",\"rows\":");
            context.Response.Write(RowsToJson(rows));
            context.Response.Write("}");
        }
        catch (Exception ex)
        {
            context.Response.StatusCode = 500;
            WriteError(context, ex.Message);
        }
    }

    private static string NormalizeRemoteAddress(string value)
    {
        string address = (value ?? "").Trim();
        if (address.StartsWith("::ffff:", StringComparison.OrdinalIgnoreCase))
        {
            return address.Substring(7);
        }
        return address;
    }

    private static string QueryForKind(string kind)
    {
        if (kind == "patients")
        {
            return @"SELECT ""VIP"", ""PCD"", ""CHR"", ""NAME"", ""UNIT_NAME"", ""STREET"", ""CITY"", ""CFE"",
       ""CFE_TYPE01"", ""HM_EIS"", ""WDS"", ""DEP"", ""ROOM"", ""PAT""
FROM ""HMERKATNOSIL""";
        }

        if (kind == "movement")
        {
            return @"SELECT ""PAT"", ""VIP"", ""PCD"", ""CHR"", ""NAME"", ""PERSCODE"", ""UNIT_NAME"", ""STREET"",
       ""ZIP"", ""CITY"", ""PHONE"", ""CFE"", ""CFE_TYPE01"", ""HM_EIS"", ""AMKA_DOC"",
       ""WDS"", ""DEP"", ""ROOM""
FROM ""HMERKATNOSIL""";
        }

        if (kind == "no_exit")
        {
            return @"SELECT DISTINCT X1100PAT.PAT, X1100PAT.NAMECHR, X1100PAT.DISDCALC, X1280DIA.DEP,
       X8001DEB.NAME, X1000PER.VIP, KEN_KDATA.KCODE
FROM X1000PER, X1100PAT, X1280DIA, X1150COG, X8001DEB, KEN_KDATA
WHERE (X1000PER.PER = X1100PAT.PER)
  AND (X1280DIA.PAT = X1100PAT.PAT)
  AND (X1150COG.PAT = X1100PAT.PAT)
  AND (X8001DEB.DEB = X1150COG.DEB)
  AND (X1100PAT.PAT = KEN_KDATA.EPISODE(+))
  AND (X1100PAT.DISD = '31-12-2099')
  AND (X1100PAT.TYP = 'S')
  AND (X1100PAT.DISDCALC IS NOT NULL)
  AND (X1280DIA.DIT = 'ΕΞΙ')
ORDER BY X1100PAT.NAMECHR, X1100PAT.DISDCALC";
        }

        return null;
    }

    private static List<Dictionary<string, object>> ExecuteRows(string sql)
    {
        ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings[ConnectionStringName];
        if (settings == null)
        {
            throw new InvalidOperationException("Missing connection string " + ConnectionStringName + ".");
        }

        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
        SqlDataSource source = new SqlDataSource();
        source.ConnectionString = settings.ConnectionString;
        source.ProviderName = settings.ProviderName;
        source.SelectCommand = sql;

        IEnumerable selected = source.Select(DataSourceSelectArguments.Empty);
        DataView view = selected as DataView;
        if (view == null)
        {
            throw new InvalidOperationException("MEDICO provider returned an unsupported result type.");
        }

        foreach (DataRowView viewRow in view)
        {
            Dictionary<string, object> row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in view.Table.Columns)
            {
                object value = viewRow[column.ColumnName];
                row[column.ColumnName] = value == DBNull.Value ? null : value;
            }
            rows.Add(row);
        }
        return rows;
    }

    private static void WriteError(HttpContext context, string message)
    {
        context.Response.Write("{\"status\":\"ERROR\",\"message\":");
        context.Response.Write(JsonString(message));
        context.Response.Write("}");
    }

    private static string RowsToJson(List<Dictionary<string, object>> rows)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("[");
        for (int r = 0; r < rows.Count; r++)
        {
            if (r > 0)
            {
                sb.Append(",");
            }
            sb.Append("{");
            int c = 0;
            foreach (KeyValuePair<string, object> item in rows[r])
            {
                if (c > 0)
                {
                    sb.Append(",");
                }
                sb.Append(JsonString(item.Key));
                sb.Append(":");
                sb.Append(JsonValue(item.Value));
                c++;
            }
            sb.Append("}");
        }
        sb.Append("]");
        return sb.ToString();
    }

    private static string JsonValue(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return "null";
        }

        if (value is DateTime)
        {
            DateTime dt = (DateTime)value;
            return JsonString(dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
        }

        if (value is bool)
        {
            return ((bool)value) ? "true" : "false";
        }

        if (value is byte[])
        {
            return JsonString(Convert.ToBase64String((byte[])value));
        }

        if (value is byte || value is sbyte || value is short || value is ushort ||
            value is int || value is uint || value is long || value is ulong ||
            value is float || value is double || value is decimal)
        {
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        return JsonString(Convert.ToString(value, CultureInfo.InvariantCulture));
    }

    private static string JsonString(string value)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("\"");
        string text = value ?? "";
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            switch (ch)
            {
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '"':
                    sb.Append("\\\"");
                    break;
                case '\b':
                    sb.Append("\\b");
                    break;
                case '\f':
                    sb.Append("\\f");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (ch < 32)
                    {
                        sb.Append("\\u");
                        sb.Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                    break;
            }
        }
        sb.Append("\"");
        return sb.ToString();
    }
}
