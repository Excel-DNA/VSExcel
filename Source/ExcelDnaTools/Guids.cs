// Guids.cs
// MUST match guids.h
using System;

namespace ExcelDna.ExcelDnaTools
{
    static class GuidList
    {
        public const string guidExcelDnaToolsPkgString = "10e82a35-4493-43be-b6d3-228399509924";
        public const string guidExcelDnaToolsCmdSetString = "ed07e70d-6403-4bc3-aa0a-50c726b6442b";
        public const string guidToolWindowPersistanceString = "f4ef482c-71f3-4638-88bd-65eedaa4665e";

        public static readonly Guid guidExcelDnaToolsCmdSet = new Guid(guidExcelDnaToolsCmdSetString);
    };
}