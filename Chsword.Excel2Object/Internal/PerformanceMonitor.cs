using System.Diagnostics;
using System.Collections.Concurrent;

namespace Chsword.Excel2Object.Internal;

/// <summary>
/// 性能监控工具，用于跟踪和分析Excel操作的性能指标
/// </summary>
internal static class PerformanceMonitor
{
    private static readonly ConcurrentDictionary<string, PerformanceMetrics> _metrics = new();
    private static readonly object _lockObject = new();

    /// <summary>
    /// 监控方法执行性能
    /// </summary>
    public static T Monitor<T>(string operationName, Func<T> operation)
    {
        var stopwatch = Stopwatch.StartNew();
        var initialMemory = GC.GetTotalMemory(false);
        var exception = (Exception?)null;

        try
        {
            var result = operation();
            return result;
        }
        catch (Exception ex)
        {
            exception = ex;
            throw;
        }
        finally
        {
            stopwatch.Stop();
            var finalMemory = GC.GetTotalMemory(false);
            
            RecordMetrics(operationName, stopwatch.ElapsedMilliseconds, 
                finalMemory - initialMemory, exception == null);
        }
    }

    /// <summary>
    /// 监控异步方法执行性能
    /// </summary>
    public static async Task<T> MonitorAsync<T>(string operationName, Func<Task<T>> operation)
    {
        var stopwatch = Stopwatch.StartNew();
        var initialMemory = GC.GetTotalMemory(false);
        var exception = (Exception?)null;

        try
        {
            var result = await operation();
            return result;
        }
        catch (Exception ex)
        {
            exception = ex;
            throw;
        }
        finally
        {
            stopwatch.Stop();
            var finalMemory = GC.GetTotalMemory(false);
            
            RecordMetrics(operationName, stopwatch.ElapsedMilliseconds, 
                finalMemory - initialMemory, exception == null);
        }
    }

    /// <summary>
    /// 创建性能监控作用域
    /// </summary>
    public static IDisposable CreateScope(string operationName)
    {
        return new PerformanceScope(operationName);
    }

    /// <summary>
    /// 记录性能指标
    /// </summary>
    private static void RecordMetrics(string operationName, long elapsedMs, long memoryDelta, bool success)
    {
        _metrics.AddOrUpdate(operationName, 
            new PerformanceMetrics(operationName, elapsedMs, memoryDelta, success),
            (key, existing) => existing.AddMeasurement(elapsedMs, memoryDelta, success));
    }

    /// <summary>
    /// 获取指定操作的性能指标
    /// </summary>
    public static PerformanceMetrics? GetMetrics(string operationName)
    {
        return _metrics.TryGetValue(operationName, out var metrics) ? metrics : null;
    }

    /// <summary>
    /// 获取所有性能指标
    /// </summary>
    public static IReadOnlyDictionary<string, PerformanceMetrics> GetAllMetrics()
    {
        return _metrics.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
    }

    /// <summary>
    /// 清除所有性能指标
    /// </summary>
    public static void ClearMetrics()
    {
        _metrics.Clear();
    }

    /// <summary>
    /// 生成性能报告
    /// </summary>
    public static string GenerateReport()
    {
        var sb = ObjectPoolManager.GetStringBuilder();
        try
        {
            sb.AppendLine("=== Excel2Object 性能报告 ===");
            sb.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            foreach (var kvp in _metrics.OrderBy(x => x.Key))
            {
                var metrics = kvp.Value;
                sb.AppendLine($"操作: {metrics.OperationName}");
                sb.AppendLine($"  总调用次数: {metrics.TotalCalls}");
                sb.AppendLine($"  成功次数: {metrics.SuccessfulCalls}");
                sb.AppendLine($"  失败次数: {metrics.FailedCalls}");
                sb.AppendLine($"  成功率: {metrics.SuccessRate:P2}");
                sb.AppendLine($"  平均执行时间: {metrics.AverageExecutionTime:F2} ms");
                sb.AppendLine($"  最短执行时间: {metrics.MinExecutionTime} ms");
                sb.AppendLine($"  最长执行时间: {metrics.MaxExecutionTime} ms");
                sb.AppendLine($"  平均内存变化: {metrics.AverageMemoryDelta / 1024.0:F2} KB");
                sb.AppendLine($"  总内存分配: {metrics.TotalMemoryAllocated / 1024.0 / 1024.0:F2} MB");
                sb.AppendLine();
            }

            return sb.ToString();
        }
        finally
        {
            ObjectPoolManager.ReturnStringBuilder(sb);
        }
    }

    /// <summary>
    /// 性能监控作用域
    /// </summary>
    private class PerformanceScope : IDisposable
    {
        private readonly string _operationName;
        private readonly Stopwatch _stopwatch;
        private readonly long _initialMemory;
        private bool _disposed;

        public PerformanceScope(string operationName)
        {
            _operationName = operationName;
            _initialMemory = GC.GetTotalMemory(false);
            _stopwatch = Stopwatch.StartNew();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _stopwatch.Stop();
                var finalMemory = GC.GetTotalMemory(false);
                RecordMetrics(_operationName, _stopwatch.ElapsedMilliseconds, 
                    finalMemory - _initialMemory, true);
                _disposed = true;
            }
        }
    }
}

/// <summary>
/// 性能指标数据结构
/// </summary>
internal class PerformanceMetrics
{
    private readonly object _lock = new();
    private long _totalExecutionTime;
    private long _totalMemoryDelta;

    public PerformanceMetrics(string operationName, long executionTime, long memoryDelta, bool success)
    {
        OperationName = operationName;
        TotalCalls = 1;
        SuccessfulCalls = success ? 1 : 0;
        FailedCalls = success ? 0 : 1;
        _totalExecutionTime = executionTime;
        _totalMemoryDelta = memoryDelta;
        MinExecutionTime = executionTime;
        MaxExecutionTime = executionTime;
        TotalMemoryAllocated = Math.Max(0, memoryDelta);
    }

    public string OperationName { get; }
    public int TotalCalls { get; private set; }
    public int SuccessfulCalls { get; private set; }
    public int FailedCalls { get; private set; }
    public long MinExecutionTime { get; private set; }
    public long MaxExecutionTime { get; private set; }
    public long TotalMemoryAllocated { get; private set; }

    public double SuccessRate => TotalCalls > 0 ? (double)SuccessfulCalls / TotalCalls : 0;
    public double AverageExecutionTime => TotalCalls > 0 ? (double)_totalExecutionTime / TotalCalls : 0;
    public double AverageMemoryDelta => TotalCalls > 0 ? (double)_totalMemoryDelta / TotalCalls : 0;

    public PerformanceMetrics AddMeasurement(long executionTime, long memoryDelta, bool success)
    {
        lock (_lock)
        {
            TotalCalls++;
            if (success)
                SuccessfulCalls++;
            else
                FailedCalls++;

            _totalExecutionTime += executionTime;
            _totalMemoryDelta += memoryDelta;
            
            MinExecutionTime = Math.Min(MinExecutionTime, executionTime);
            MaxExecutionTime = Math.Max(MaxExecutionTime, executionTime);
            TotalMemoryAllocated += Math.Max(0, memoryDelta);

            return this;
        }
    }
}
