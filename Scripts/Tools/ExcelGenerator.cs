using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEngine;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Attributes;

/// <summary>
/// 根据数据类型生成Excel存储文件的工具类
/// </summary>
public static class ExcelGenerator
{
    public static void GenerateFile(Type type, bool overwrite = false)
    {
        // 检查文件是否存在
        string filePath = $"{Application.streamingAssetsPath}/{type.Name}s.xlsx";
        if (!overwrite)
        {
            int copyCount = 1;
            while (File.Exists(filePath))
                filePath = $"{Application.streamingAssetsPath}/{type.Name}s - 副本{copyCount++}.xlsx";
            if (copyCount > 1)
                Debug.LogWarning($"类型{type.Name}Excel配置表已存在，已创建副本");
        }

        // 设置样式
        var config = new OpenXmlConfiguration()
        {
            TableStyles = TableStyles.None,
            DynamicColumns = new DynamicExcelColumn[type.GetProperties().Length + 1]
        };

        // 准备数据
        var value = new List<Dictionary<string, string>>();
        for (int i = 0; i < 3; i++)
            value.Add(new Dictionary<string, string>());
        int pCount = 0;
        int ignoreCount = 0;
        foreach (var property in type.GetProperties())
        {
            // 忽略某些列
            if (property.GetCustomAttribute<ExcelIgnoreAttribute>() != null)
            {
                config.DynamicColumns[pCount++] = new DynamicExcelColumn(property.Name) { Ignore = true };
                ++ignoreCount;
                continue;
            }
            // 填充空行
            value[0].Add(property.Name, "");
            // 写入类型
            value[1].Add(property.Name, GetShortName(property.PropertyType));
            // 写入列名
            var name = property.GetCustomAttribute<ExcelColumnNameAttribute>();
            if  (name != null)
                value[2].Add(property.Name, name.ExcelColumnName);
            else
                value[2].Add(property.Name, property.Name);
            // 设置列宽
            config.DynamicColumns[pCount++] = new DynamicExcelColumn(property.Name)
            {
                Width = GetWidth(property),
                Index = pCount - ignoreCount - 1,
                Name = property.Name,
            };
        }

        // 添加注释
        const string comment = "__Comment";
        config.DynamicColumns[pCount++] = new DynamicExcelColumn(comment)
        {
            Width = 60,
            Index = pCount - ignoreCount - 1,
            Name = comment,
        };
        value[0].Add(comment, "第一行可以用来给属性写注释，这一列可以用来给每条数据写备注");
        value[1].Add(comment, "第二行是属性数据类型，不会被读取，可以修改");
        value[2].Add(comment, "第三行是属性名称，会被读取，不可更改");

        MiniExcel.SaveAs(filePath, value, false, type.Name, ExcelType.XLSX, config, true);
    }
    public static void GenerateFile<T>(bool overwrite = false) where T : class, new()
    { 
        GenerateFile(typeof(T), overwrite); 
    }
    private static string GetShortName(Type type)
    {
        if (type.Name == typeof(int).Name)
            return "int";
        if (type.Name == typeof(float).Name)
            return "float";
        if (type.Name == typeof(string).Name)
            return "string";
        if (type.Name == typeof(bool).Name)
            return "bool";
        return type.Name;
    }
    private static double GetWidth(PropertyInfo property)
    {
        var width = property.GetCustomAttribute<ExcelColumnWidthAttribute>();
        if (width != null)
            return width.ExcelColumnWidth;

        var nameAttr = property.GetCustomAttribute<ExcelColumnNameAttribute>();
        string name = nameAttr == null ? property.Name : nameAttr.ExcelColumnName;
        int ret = 2;
        foreach (char c in name)
            ret += c < 128 ? 1 : 2;
        return Mathf.Max(8, ret);
    }
}
