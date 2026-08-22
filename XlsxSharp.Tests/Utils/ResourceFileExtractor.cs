using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// A loader of resources from an assembly.
/// </summary>
public sealed class ResourceFileExtractor
{
    /// <summary>Assembly used to load resources.</summary>
    private readonly Assembly _assembly;

    /// <summary>A prefix of loadable resources names in the assembly.</summary>
    private readonly string _resourcePathPrefix;

    /// <param name="assembly">Assembly that contains the resources.</param>
    /// <param name="resourcePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
    public ResourceFileExtractor(Assembly assembly, string resourcePath)
    {
        this._assembly = assembly ?? Assembly.GetCallingAssembly();
        this._resourcePathPrefix = this._assembly.GetName().Name + resourcePath;
    }

    public IEnumerable<string> GetFileNames(Func<string, bool> predicate)
    {
        foreach (string resourceName in this._assembly.GetManifestResourceNames())
        {
            if (resourceName.StartsWith(this._resourcePathPrefix) && predicate(resourceName))
            {
                yield return resourceName[this._resourcePathPrefix.Length..];
            }
        }
    }

    /// <summary>
    /// Read file in current assembly by specific file name
    /// </summary>
    public Stream ReadFileFromResourceToStream(string fileName)
    {
        string resourceFileName = this._resourcePathPrefix + fileName;
        Stream? stream = this._assembly.GetManifestResourceStream(resourceFileName);
        if (stream is null)
        {
            throw new ArgumentException(
                "Can't find resource file " + resourceFileName,
                nameof(fileName)
            );
        }

        return stream;
    }
}
