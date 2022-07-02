using ClosedXML.Extensions;
using NUnit.Framework;
using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Tests.Extensions
{
    public class ReflectionExtensionTests
    {
        private class TestClass
        {
            public static int StaticProperty { get; set; }
            public static int StaticField = default;
#pragma warning disable CS0067
            public static event EventHandler<EventArgs> StaticEvent;
#pragma warning restore CS0067
            public static void StaticMethod()
            {
            }

            public const int Const = 100;

            public TestClass()
            {
            }

            public int InstanceProperty { get; set; }
            public int InstanceField = default;

#pragma warning disable CS0067
            public event EventHandler<EventArgs> InstanceEvent;
#pragma warning restore CS0067
            public void InstanceMethod()
            {
            }
        }

        [TestCase(nameof(TestClass.StaticProperty), true)]
        [TestCase(nameof(TestClass.StaticField), true)]
        [TestCase(nameof(TestClass.StaticEvent), true)]
        [TestCase(nameof(TestClass.StaticMethod), true)]
        [TestCase(nameof(TestClass.Const), true)]
        [TestCase(nameof(TestClass.InstanceProperty), false)]
        [TestCase(nameof(TestClass.InstanceField), false)]
        [TestCase(nameof(TestClass.InstanceEvent), false)]
        [TestCase(nameof(TestClass.InstanceMethod), false)]
        public void IsStatic(string memberName, bool expectedIsStatic)
        {
            var member = typeof(TestClass).GetMember(memberName).Single();
            Assert.AreEqual(expectedIsStatic, member.IsStatic());
        }

        [TestCase(BindingFlags.Static | BindingFlags.NonPublic, true)]
        [TestCase(BindingFlags.Instance | BindingFlags.Public, false)]
        public void ConstructorIsStatic(BindingFlags flag, bool expectedIsStatic)
        {
            var constructors = typeof(TestClass).GetConstructors(flag);
            Assert.AreEqual(expectedIsStatic, constructors.Single().IsStatic());
        }
    }
}
