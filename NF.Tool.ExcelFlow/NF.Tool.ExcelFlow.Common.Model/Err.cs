using System.IO;
using System.Runtime.CompilerServices;

namespace NF.Tool.ExcelFlow.Common.Model;

public static class Err
{
    public static string MessageWithInfo(
        string message,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0,
        [CallerMemberName] string memberName = ""
    )
    {
        string errMessage = $"[==> {Path.GetFileName(filePath),-20} - {lineNumber} - {memberName}\n{message}";
#if DEBUG
        // Debug.Fail(errMessage);
#endif // DEBUG

        return errMessage;
    }
}
