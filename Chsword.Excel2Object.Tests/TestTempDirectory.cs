using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     把临时目录改指到本进程自己的目录。流式导出会把刷出内存的行写进临时目录下的 <c>poifiles</c>，
///     而 NPOI 只在首次使用时解析一次该目录，各目标框架的测试进程又是并行跑的：共用系统临时目录，
///     「导出结束后临时文件已删除」这类断言就会数到别的进程头上。
/// </summary>
[TestClass]
public static class TestTempDirectory
{
    private static string? _dir;

    [AssemblyInitialize]
    public static void Redirect(TestContext context)
    {
        _dir = Path.Combine(Path.GetTempPath(), "e2o-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
        foreach (var name in new[] {"TMPDIR", "TMP", "TEMP"}) Environment.SetEnvironmentVariable(name, _dir);
    }

    [AssemblyCleanup]
    public static void Remove()
    {
        if (_dir == null) return;

        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
            // 清理失败不必连累测试结论
        }
    }
}
