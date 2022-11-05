using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using MiniExcelLibs;

public class ExcelData
{
    public int ID { get; set; }
    public string Name { get; set; }
    public float Speed { get; set; }
    [MiniExcelLibs.Attributes.ExcelColumnName("我将带头冲锋！！！")]
    public string AHead { get; set; }
    [MiniExcelLibs.Attributes.ExcelIgnore]
    public int Dummy { get; set; }
    public long AVeeeeeeeryLongProperty { get; set; }
    public short S { get; set; }
    [MiniExcelLibs.Attributes.ExcelColumnWidth(15)]
    public int Middle { get; set; } 
}

public class ExcelTest : MonoBehaviour
{
    protected void Start()
    {
        ExcelGenerator.GenerateFile<ExcelData>(true);
    }

    protected void Update()
    {
        
    }
}
