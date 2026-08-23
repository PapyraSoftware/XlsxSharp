using System;
using System.Linq;
using System.Reflection;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Extensions;

public class ReflectionExtensionTests
{
    private class TestClass
    {
        static TestClass() { }

        public static int StaticProperty { get; set; }
        public static int StaticField = 0;

        public static event EventHandler<EventArgs> StaticEvent
        {
            add => _ = value;
            remove => _ = value;
        }

        public static void StaticMethod() { }

        public const int Const = 100;

        public int InstanceProperty { get; set; }
        public int InstanceField = 0;

        public event EventHandler<EventArgs> InstanceEvent
        {
            add => _ = value;
            remove => _ = value;
        }

        public void InstanceMethod() { }
    }

    [Test]
    [Arguments(nameof(TestClass.StaticProperty), true)]
    [Arguments(nameof(TestClass.StaticField), true)]
    [Arguments(nameof(TestClass.StaticEvent), true)]
    [Arguments(nameof(TestClass.StaticMethod), true)]
    [Arguments(nameof(TestClass.Const), true)]
    [Arguments(nameof(TestClass.InstanceProperty), false)]
    [Arguments(nameof(TestClass.InstanceField), false)]
    [Arguments(nameof(TestClass.InstanceEvent), false)]
    [Arguments(nameof(TestClass.InstanceMethod), false)]
    public void IsStatic(string memberName, bool expectedIsStatic)
    {
        MemberInfo member = typeof(TestClass).GetMember(memberName).Single();
        ClassicAssert.AreEqual(expectedIsStatic, member.IsStatic());
    }

    [Test]
    [Arguments(BindingFlags.Static | BindingFlags.NonPublic, true)]
    [Arguments(BindingFlags.Instance | BindingFlags.Public, false)]
    public void ConstructorIsStatic(BindingFlags flag, bool expectedIsStatic)
    {
        ConstructorInfo[] constructors = typeof(TestClass).GetConstructors(flag);
        ClassicAssert.AreEqual(expectedIsStatic, constructors.Single().IsStatic());
    }
}
