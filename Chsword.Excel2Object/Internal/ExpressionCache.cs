using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace Chsword.Excel2Object.Internal;

/// <summary>
/// 表达式缓存机制，用于优化重复表达式转换的性能
/// </summary>
internal static class ExpressionCache
{
    private static readonly ConcurrentDictionary<ExpressionCacheKey, string> _cache = new();
    private static readonly object _lockObject = new();
    private const int MaxCacheSize = 1000;

    /// <summary>
    /// 获取或添加表达式转换结果到缓存
    /// </summary>
    public static string GetOrAdd(Expression expression, string[] columns, int rowIndex, Func<string> factory)
    {
        var key = new ExpressionCacheKey(expression.ToString(), columns, rowIndex);
        
        // 尝试从缓存获取
        if (_cache.TryGetValue(key, out var cachedResult))
        {
            return cachedResult;
        }

        // 生成新结果并添加到缓存
        var result = factory();
        
        // 检查缓存大小，避免内存溢出
        if (_cache.Count >= MaxCacheSize)
        {
            ClearOldestEntries();
        }
        
        _cache.TryAdd(key, result);
        return result;
    }

    /// <summary>
    /// 清除缓存
    /// </summary>
    public static void Clear()
    {
        _cache.Clear();
    }

    /// <summary>
    /// 获取缓存统计信息
    /// </summary>
    public static CacheStatistics GetStatistics()
    {
        return new CacheStatistics
        {
            TotalEntries = _cache.Count,
            MaxSize = MaxCacheSize
        };
    }

    /// <summary>
    /// 清除最旧的缓存条目（简单的LRU策略）
    /// </summary>
    private static void ClearOldestEntries()
    {
        lock (_lockObject)
        {
            if (_cache.Count >= MaxCacheSize)
            {
                // 清除一半的缓存条目
                var entriesToRemove = _cache.Count / 2;
                var keysToRemove = _cache.Keys.Take(entriesToRemove).ToList();
                
                foreach (var key in keysToRemove)
                {
                    _cache.TryRemove(key, out _);
                }
            }
        }
    }

    /// <summary>
    /// 表达式缓存键
    /// </summary>
    private readonly struct ExpressionCacheKey : IEquatable<ExpressionCacheKey>
    {
        private readonly string _expressionString;
        private readonly string _columnsHash;
        private readonly int _rowIndex;
        private readonly int _hashCode;

        public ExpressionCacheKey(string expressionString, string[] columns, int rowIndex)
        {
            _expressionString = expressionString;
            _columnsHash = string.Join(",", columns);
            _rowIndex = rowIndex;
            
            // 使用简单的哈希算法兼容低版本.NET
            _hashCode = CombineHashCodes(_expressionString?.GetHashCode() ?? 0, 
                                       _columnsHash?.GetHashCode() ?? 0, 
                                       _rowIndex.GetHashCode());
        }

        private static int CombineHashCodes(int h1, int h2, int h3)
        {
            // 简单的哈希组合算法
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + h1;
                hash = hash * 23 + h2;
                hash = hash * 23 + h3;
                return hash;
            }
        }

        public bool Equals(ExpressionCacheKey other)
        {
            return _expressionString == other._expressionString &&
                   _columnsHash == other._columnsHash &&
                   _rowIndex == other._rowIndex;
        }

        public override bool Equals(object? obj)
        {
            return obj is ExpressionCacheKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }
    }

    /// <summary>
    /// 缓存统计信息
    /// </summary>
    public class CacheStatistics
    {
        public int TotalEntries { get; set; }
        public int MaxSize { get; set; }
        public double UsagePercentage => MaxSize > 0 ? (double)TotalEntries / MaxSize * 100 : 0;
    }
}
