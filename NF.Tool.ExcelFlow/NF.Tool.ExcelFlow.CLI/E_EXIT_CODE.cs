namespace NF.Tool.ExcelFlow.CLI;

public enum E_EXIT_CODE
{
    NONE = 0,
    FAIL_BAKE_MODEL = 101,
    //FAIL_BAKE_DB = 102,
    FAIL_BAKE_CODE = 103,
    FAIL_ASSEMBLY = 104,
    FAIL_COLLECT_ROWS = 105,
    CommandAppException = 106,
    COMMAND_APP_EXCEPTION = 107,
    EXCEL_FLOW_EXCEPTION = 108,
    UNHANDLE = 109,
    FAIL_GET_CONFIG = 110,
}
