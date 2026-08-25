using XlsxSharp.Excel.Caching;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Caching;

public class BaseRepositoryTests
{
    [Test]
    public void DifferentEntitiesWithSameKeyStoredOnce()
    {
        // Arrange
        int key = 12345;
        SampleEntity entity1 = new(key);
        SampleEntity entity2 = new(key);
        SampleRepository sampleRepository = CreateSampleRepository();

        // Act
        SampleEntity? storedEntity1 = sampleRepository.Store(ref key, entity1);
        SampleEntity? storedEntity2 = sampleRepository.Store(ref key, entity2);

        // Assert
        ClassicAssert.AreSame(entity1, storedEntity1);
        ClassicAssert.AreSame(entity1, storedEntity2);
        ClassicAssert.AreNotSame(entity2, storedEntity2);
    }

    [Test]
    public void NonUsedReferencesAreGCed()
    {
#if !DEBUG
        // Arrange
        int key = 12345;
        SampleRepository sampleRepository = CreateSampleRepository();

        // Act
        // In net8, JIT could make a hidden temporary variable for created object that would prevent
        // GC collection. Therefore, make the reference in another method, so the hidden variable
        // doesn't get inlined. https://github.com/dotnet/runtime/issues/63568#issuecomment-1008602069
        WeakReference storedEntityRef1 = AddEntityToRepository(sampleRepository, ref key);

        int count = 0;
        do
        {
            System.Threading.Thread.Sleep(50);
            System.GC.Collect();
            count++;
        } while (storedEntityRef1.IsAlive && count < 10);

        // Assert
        if (count == 10)
        {
            ClassicAssert.Fail("storedEntityRef1 was not GCed");
        }

        ClassicAssert.IsFalse(sampleRepository.Any());

        return;

        static System.WeakReference AddEntityToRepository(SampleRepository repository, ref int key)
        {
            return new System.WeakReference(repository.Store(ref key, new SampleEntity(key)));
        }
#endif
    }

    [Test]
    public void NonUsedReferencesAreGCed2()
    {
#if !DEBUG
        // Arrange
        int countUnique = 30;
        int repeatCount = 1000;
        SampleEntity[] entities = new SampleEntity[countUnique * repeatCount];
        for (int i = 0; i < countUnique; i++)
        {
            for (int j = 0; j < repeatCount; j++)
            {
                entities[i * repeatCount + j] = new SampleEntity(i);
            }
        }

        SampleRepository sampleRepository = CreateSampleRepository();

        // Act
        Parallel.ForEach(
            entities,
            new ParallelOptions { MaxDegreeOfParallelism = 8 },
            e =>
            {
                int key = e.Key;
                sampleRepository.Store(ref key, e);
            }
        );

        System.Threading.Thread.Sleep(50);
        System.GC.Collect();
        List<SampleEntity> storedEntries = [.. sampleRepository];

        // Assert
        ClassicAssert.AreEqual(0, storedEntries.Count);
#endif
    }

    [Test]
    public void ConcurrentAddingCausesNoDuplication()
    {
        // Arrange
        int countUnique = 30;
        int repeatCount = 1000;
        SampleEntity[] entities = new SampleEntity[countUnique * repeatCount];
        for (int i = 0; i < countUnique; i++)
        {
            for (int j = 0; j < repeatCount; j++)
            {
                entities[i * repeatCount + j] = new SampleEntity(i);
            }
        }

        SampleRepository sampleRepository = CreateSampleRepository();

        // Act
        Parallel.ForEach(
            entities,
            new ParallelOptions { MaxDegreeOfParallelism = 8 },
            e =>
            {
                int key = e.Key;
                sampleRepository.Store(ref key, e);
            }
        );
        List<SampleEntity> storedEntries = [.. sampleRepository];

        // Assert
        ClassicAssert.AreEqual(countUnique, storedEntries.Count);
        ClassicAssert.NotNull(entities); // To protect them from GC
    }

    [Test]
    public void ReplaceKeyInRepository()
    {
        // Arrange
        int key1 = 12345;
        int key2 = 54321;
        SampleEntity entity = new(key1);
        SampleRepository sampleRepository = CreateSampleRepository();
        SampleEntity? storedEntity1 = sampleRepository.Store(ref key1, entity);

        // Act
        sampleRepository.Replace(ref key1, ref key2);
        bool containsOld = sampleRepository.ContainsKey(ref key1, out SampleEntity? _);
        bool containsNew = sampleRepository.ContainsKey(ref key2, out SampleEntity? _);
        SampleEntity storedEntity2 = sampleRepository.GetOrCreate(ref key2);

        // Assert
        ClassicAssert.IsFalse(containsOld);
        ClassicAssert.IsTrue(containsNew);
        ClassicAssert.AreSame(entity, storedEntity1);
        ClassicAssert.AreSame(entity, storedEntity2);
    }

    [Test]
    public void ConcurrentReplaceKeyInRepository()
    {
        EditableRepository sampleRepository = new();
        int[] keys = [.. Enumerable.Range(0, 1000)];
        keys.ForEach(key => sampleRepository.GetOrCreate(ref key));

        Parallel.ForEach(
            keys,
            key =>
            {
                int modifiedKey = key + 2000;
                EditableEntity? val1 = sampleRepository.Replace(ref key, ref modifiedKey);
                val1.Key = key + 2000;
                EditableEntity val2 = sampleRepository.GetOrCreate(ref modifiedKey);
                ClassicAssert.AreSame(val1, val2);
            }
        );
    }

    [Test]
    public void ReplaceNonExistingKeyInRepository()
    {
        int key1 = 100;
        int key2 = 200;
        int key3 = 300;
        SampleEntity entity = new(key1);
        SampleRepository sampleRepository = CreateSampleRepository();
        sampleRepository.Store(ref key1, entity);

        sampleRepository.Replace(ref key2, ref key3);
        List<SampleEntity> all = [.. sampleRepository];

        ClassicAssert.AreEqual(1, all.Count);
        ClassicAssert.AreSame(entity, all.First());
    }

    private static SampleRepository CreateSampleRepository() => new();

    /// <summary>
    /// Class under testing
    /// </summary>
    internal class SampleRepository : XLRepositoryBase<int, SampleEntity>
    {
        public SampleRepository()
            : base(key => new SampleEntity(key)) { }
    }

    public class SampleEntity
    {
        public int Key { get; private set; }

        public SampleEntity(int key) => this.Key = key;
    }

    /// <summary>
    /// Class under testing
    /// </summary>
    internal class EditableRepository : XLRepositoryBase<int, EditableEntity>
    {
        public EditableRepository()
            : base(key => new EditableEntity(key)) { }
    }

    public class EditableEntity
    {
        public int Key { get; set; }

        public EditableEntity(int key) => this.Key = key;
    }
}
