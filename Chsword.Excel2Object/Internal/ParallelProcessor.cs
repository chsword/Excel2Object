using System.Collections.Concurrent;

namespace Chsword.Excel2Object.Internal;

/// <summary>
/// 并行处理优化器，提供高效的批量数据处理能力
/// </summary>
internal static class ParallelProcessor
{
    private const int DefaultBatchSize = 1000;
    private static readonly int DefaultMaxDegreeOfParallelism = Environment.ProcessorCount;

    /// <summary>
    /// 并行处理大量数据
    /// </summary>
    public static async Task<IEnumerable<TResult>> ProcessInParallelAsync<TSource, TResult>(
        IEnumerable<TSource> source,
        Func<TSource, TResult> processor,
        int batchSize = DefaultBatchSize,
        int maxDegreeOfParallelism = -1,
        CancellationToken cancellationToken = default)
    {
        if (maxDegreeOfParallelism == -1)
            maxDegreeOfParallelism = DefaultMaxDegreeOfParallelism;

        return await Task.Run(() =>
        {
            var results = new ConcurrentBag<TResult>();
            var batches = ChunkInternal(source, batchSize);

            Parallel.ForEach(batches, new ParallelOptions
            {
                MaxDegreeOfParallelism = maxDegreeOfParallelism,
                CancellationToken = cancellationToken
            }, batch =>
            {
                foreach (var item in batch)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var result = processor(item);
                    results.Add(result);
                }
            });

            return results.AsEnumerable();
        }, cancellationToken);
    }

    /// <summary>
    /// 并行处理集合到集合的转换
    /// </summary>
    public static async Task<IEnumerable<TResult>> ProcessCollectionInParallelAsync<TSource, TResult>(
        IEnumerable<TSource> source,
        Func<IEnumerable<TSource>, IEnumerable<TResult>> processor,
        int batchSize = DefaultBatchSize,
        int maxDegreeOfParallelism = -1,
        CancellationToken cancellationToken = default)
    {
        if (maxDegreeOfParallelism == -1)
            maxDegreeOfParallelism = DefaultMaxDegreeOfParallelism;

        return await Task.Run(() =>
        {
            var results = new ConcurrentBag<TResult>();
            var batches = ChunkInternal(source, batchSize);

            Parallel.ForEach(batches, new ParallelOptions
            {
                MaxDegreeOfParallelism = maxDegreeOfParallelism,
                CancellationToken = cancellationToken
            }, batch =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                var batchResults = processor(batch);
                foreach (var result in batchResults)
                {
                    results.Add(result);
                }
            });

            return results.AsEnumerable();
        }, cancellationToken);
    }

    /// <summary>
    /// 将集合分块处理
    /// </summary>
    private static IEnumerable<IEnumerable<T>> ChunkInternal<T>(IEnumerable<T> source, int size)
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
}

/// <summary>
/// 扩展方法来简化并行处理操作
/// </summary>
internal static class ParallelExtensions
{
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
