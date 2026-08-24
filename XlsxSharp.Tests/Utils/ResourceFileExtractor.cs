using Assembly = System.Reflection.Assembly;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// A loader of resources from an assembly.
/// </summary>
public sealed class ResourceFileExtractor
{
    /// <summary>Assembly used to load resources.</summary>
    private readonly Assembly assembly;

    /// <summary>A prefix of loadable resources names in the assembly.</summary>
    private readonly string resourcePathPrefix;

    /// <param name="assembly">Assembly that contains the resources.</param>
    /// <param name="resourcePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
    public ResourceFileExtractor(Assembly assembly, string resourcePath)
    {
        this.assembly = assembly ?? Assembly.GetCallingAssembly();
        this.resourcePathPrefix = this.assembly.GetName().Name + resourcePath;
    }

    public IEnumerable<string> GetFileNames(Func<string, bool> predicate)
    {
        foreach (string resourceName in this.assembly.GetManifestResourceNames())
        {
            if (resourceName.StartsWith(this.resourcePathPrefix) && predicate(resourceName))
            {
                yield return resourceName[this.resourcePathPrefix.Length..];
            }
        }
    }

    /// <summary>
    /// Read file in current assembly by specific file name
    /// </summary>
    public Stream ReadFileFromResourceToStream(string fileName)
    {
        string resourceFileName = this.resourcePathPrefix + fileName;
        Stream? stream = this.assembly.GetManifestResourceStream(resourceFileName);
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
