using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEngine;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Attributes;

/// <summary>
/// ����������������Excel�洢�ļ��Ĺ�����
/// </summary>
public static class ExcelGenerator
{
    public static void GenerateFile(Type type, bool overwrite = false)
    {
        // ����ļ��Ƿ����
        string filePath = $"{Application.streamingAssetsPath}/{type.Name}s.xlsx";
        if (!overwrite)
        {
            int copyCount = 1;
            while (File.Exists(filePath))
                filePath = $"{Application.streamingAssetsPath}/{type.Name}s - ����{copyCount++}.xlsx";
            if (copyCount > 1)
                Debug.LogWarning($"����{type.Name}Excel���ñ��Ѵ��ڣ��Ѵ�������");
        }

        // ������ʽ
        var config = new OpenXmlConfiguration()
        {
            TableStyles = TableStyles.None,
            DynamicColumns = new DynamicExcelColumn[type.GetProperties().Length + 1]
        };

        // ׼������
        var value = new List<Dictionary<string, string>>();
        for (int i = 0; i < 3; i++)
            value.Add(new Dictionary<string, string>());
        int pCount = 0;
        int ignoreCount = 0;
        foreach (var property in type.GetProperties())
        {
            // ����ĳЩ��
            if (property.GetCustomAttribute<ExcelIgnoreAttribute>() != null)
            {
                config.DynamicColumns[pCount++] = new DynamicExcelColumn(property.Name) { Ignore = true };
                ++ignoreCount;
                continue;
            }
            // ������
            value[0].Add(property.Name, "");
            // д������
            value[1].Add(property.Name, GetShortName(property.PropertyType));
            // д������
            var name = property.GetCustomAttribute<ExcelColumnNameAttribute>();
            if  (name != null)
                value[2].Add(property.Name, name.ExcelColumnName);
            else
                value[2].Add(property.Name, property.Name);
            // �����п�
            config.DynamicColumns[pCount++] = new DynamicExcelColumn(property.Name)
            {
                Width = GetWidth(property),
                Index = pCount - ignoreCount - 1,
                Name = property.Name,
            };
        }

        // ���ע��
        const string comment = "__Comment";
        config.DynamicColumns[pCount++] = new DynamicExcelColumn(comment)
        {
            Width = 60,
            Index = pCount - ignoreCount - 1,
            Name = comment,
        };
        value[0].Add(comment, "��һ�п�������������дע�ͣ���һ�п���������ÿ������д��ע");
        value[1].Add(comment, "�ڶ����������������ͣ����ᱻ��ȡ�������޸�");
        value[2].Add(comment, "���������������ƣ��ᱻ��ȡ�����ɸ���");

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
