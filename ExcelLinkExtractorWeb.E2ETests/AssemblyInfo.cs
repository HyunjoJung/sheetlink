using NUnit.Framework;

// Disable parallel test execution to avoid rate limiting and server overload
[assembly: Parallelizable(ParallelScope.None)]
[assembly: LevelOfParallelism(1)]
