using System;

namespace purify.Data;

public interface IDocProcessor : IDisposable
{
    void ProcessFile(OneFileOptions opts);
    string TargetExt { get; }
}