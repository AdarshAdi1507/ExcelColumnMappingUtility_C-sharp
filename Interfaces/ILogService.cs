using System;
using System.Collections.Generic;

namespace ExcelProcessor.Interfaces
{
    public interface ILogService
    {
        void LogInformation(string message);
        void LogError(string message, Exception ex = null);
    }
}