using System.Collections.Concurrent;
using System.Text;

namespace Chsword.Excel2Object.Internal;

/// <summary>
/// 简单的对象池管理器，用于重用常用对象以减少GC压力
/// </summary>
internal static class ObjectPoolManager
{
    private static readonly ConcurrentQueue<StringBuilder> _stringBuilderPool = new();
    private static readonly ConcurrentQueue<List<string>> _stringListPool = new();
    private static readonly ConcurrentQueue<Dictionary<string, object>> _dictionaryPool = new();
    
    private const int MaxPoolSize = 100;

    /// <summary>
    /// 获取StringBuilder实例
    /// </summary>
    public static StringBuilder GetStringBuilder()
    {
        if (_stringBuilderPool.TryDequeue(out var sb))
        {
            sb.Clear();
            return sb;
        }
        return new StringBuilder();
    }

    /// <summary>
    /// 归还StringBuilder实例
    /// </summary>
    public static void ReturnStringBuilder(StringBuilder sb)
    {
        if (sb != null && _stringBuilderPool.Count < MaxPoolSize)
        {
            sb.Clear();
            _stringBuilderPool.Enqueue(sb);
        }
    }

    /// <summary>
    /// 获取字符串列表实例
    /// </summary>
    public static List<string> GetStringList()
    {
        if (_stringListPool.TryDequeue(out var list))
        {
            list.Clear();
            return list;
        }
        return new List<string>();
    }

    /// <summary>
    /// 归还字符串列表实例
    /// </summary>
    public static void ReturnStringList(List<string> list)
    {
        if (list != null && _stringListPool.Count < MaxPoolSize)
        {
            list.Clear();
            _stringListPool.Enqueue(list);
        }
    }

    /// <summary>
    /// 获取字典实例
    /// </summary>
    public static Dictionary<string, object> GetDictionary()
    {
        if (_dictionaryPool.TryDequeue(out var dict))
        {
            dict.Clear();
            return dict;
        }
        return new Dictionary<string, object>();
    }

    /// <summary>
    /// 归还字典实例
    /// </summary>
    public static void ReturnDictionary(Dictionary<string, object> dict)
    {
        if (dict != null && _dictionaryPool.Count < MaxPoolSize)
        {
            dict.Clear();
            _dictionaryPool.Enqueue(dict);
        }
    }

    /// <summary>
    /// 使用StringBuilder的便捷方法
    /// </summary>
    public static string UsingStringBuilder(Action<StringBuilder> action)
    {
        var sb = GetStringBuilder();
        try
        {
            action(sb);
            return sb.ToString();
        }
        finally
        {
            ReturnStringBuilder(sb);
        }
    }

    /// <summary>
    /// 使用字符串列表的便捷方法
    /// </summary>
    public static T UsingStringList<T>(Func<List<string>, T> func)
    {
        var list = GetStringList();
        try
        {
            return func(list);
        }
        finally
        {
            ReturnStringList(list);
        }
    }
}

/// <summary>
/// 内存监控工具
/// </summary>
internal static class MemoryMonitor
{
    /// <summary>
    /// 监控内存使用情况
    /// </summary>
    public static MemoryUsageInfo GetMemoryUsage()
    {
        var beforeGC = GC.GetTotalMemory(false);
        var afterGC = GC.GetTotalMemory(true);
        
        return new MemoryUsageInfo
        {
            BeforeGC = beforeGC,
            AfterGC = afterGC,
            CollectedBytes = beforeGC - afterGC,
            Gen0Collections = GC.CollectionCount(0),
            Gen1Collections = GC.CollectionCount(1),
            Gen2Collections = GC.CollectionCount(2)
        };
    }

    /// <summary>
    /// 执行带内存监控的操作
    /// </summary>
    public static T MonitorMemory<T>(Func<T> operation, Action<MemoryUsageInfo>? onComplete = null)
    {
        var initialMemory = GC.GetTotalMemory(false);
        var initialGen0 = GC.CollectionCount(0);
        var initialGen1 = GC.CollectionCount(1);
        var initialGen2 = GC.CollectionCount(2);

        var result = operation();

        var finalMemory = GC.GetTotalMemory(false);
        var info = new MemoryUsageInfo
        {
            BeforeGC = initialMemory,
            AfterGC = finalMemory,
            CollectedBytes = Math.Max(0, finalMemory - initialMemory),
            Gen0Collections = GC.CollectionCount(0) - initialGen0,
            Gen1Collections = GC.CollectionCount(1) - initialGen1,
            Gen2Collections = GC.CollectionCount(2) - initialGen2
        };

        onComplete?.Invoke(info);
        return result;
    }

    /// <summary>
    /// 内存使用信息
    /// </summary>
    public class MemoryUsageInfo
    {
        public long BeforeGC { get; set; }
        public long AfterGC { get; set; }
        public long CollectedBytes { get; set; }
        public int Gen0Collections { get; set; }
        public int Gen1Collections { get; set; }
        public int Gen2Collections { get; set; }

        public double MemoryUsageMB => AfterGC / 1024.0 / 1024.0;
        public double CollectedMB => CollectedBytes / 1024.0 / 1024.0;
    }
}
