namespace Chsword.Excel2Object.Internal;

/// <summary>
/// 扩展方法来简化批处理操作
/// </summary>
internal static class EnumerableExtensions
{
    /// <summary>
    /// 将集合分块处理
    /// </summary>
    public static IEnumerable<IEnumerable<T>> Chunk<T>(this IEnumerable<T> source, int size)
    {
        var chunk = new List<T>(size);
        foreach (var item in source)
        {
            chunk.Add(item);
            if (chunk.Count == size)
            {
                yield return chunk;
                chunk = new List<T>(size);
            }
        }
        
        if (chunk.Count > 0)
        {
            yield return chunk;
        }
    }

    /// <summary>
    /// 并行映射转换
    /// </summary>
    public static async Task<IEnumerable<TResult>> ParallelSelectAsync<TSource, TResult>(
        this IEnumerable<TSource> source,
        Func<TSource, TResult> selector,
        int maxDegreeOfParallelism = -1,
        CancellationToken cancellationToken = default)
    {
        return await ParallelProcessor.ProcessInParallelAsync(
            source, 
            selector, 
            maxDegreeOfParallelism: maxDegreeOfParallelism == -1 ? Environment.ProcessorCount : maxDegreeOfParallelism,
            cancellationToken: cancellationToken);
    }
}
