using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Tests that allocate a lot, or that measure how much the process allocates, run one at a time.
/// </summary>
/// <remarks>
/// xunit runs test classes in parallel in one process, and a heap measurement is of the whole process: a test
/// that builds a few hundred megabytes of hostile XML beside one that asserts a writer did not grow the heap by
/// four makes the second fail for a reason that has nothing to do with the writer. Classes that share this
/// collection never overlap.
/// </remarks>
[CollectionDefinition(HeapCollection.Name)]
public sealed class HeapCollection
{
    internal const string Name = "Heap measurement";
}
