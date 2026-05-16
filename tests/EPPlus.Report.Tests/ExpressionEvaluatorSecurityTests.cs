using System;
using System.Collections.Generic;
using EPPlus.Report.Evaluation;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class ExpressionEvaluatorSecurityTests
    {
        public class TestAddress
        {
            public string City { get; set; } = "Paris";
            public string ZipCode { get; set; } = "75001";
        }

        public class TestPerson
        {
            public string Name { get; set; } = "Alice";
            public int Age { get; set; } = 30;
            public TestAddress Address { get; set; } = new TestAddress();
        }

        public class TestWriteOnly
        {
            private string _value;
            public string Name { get; set; } = "Test";
            public string WriteOnlyProperty { set => _value = value; }
        }

        [Fact]
        public void Evaluate_WriteOnlyProperty_ThrowsPropertyNotFoundException()
        {
            var evaluator = new ExpressionEvaluator();
            var obj = new TestWriteOnly();

            var ex = Assert.Throws<PropertyNotFoundException>(() => evaluator.Evaluate("WriteOnlyProperty", obj));
            Assert.Contains("WriteOnlyProperty", ex.Message);
            Assert.Contains("write-only", ex.Message);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_PropertyAllowed_ReturnsValue()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson { Name = "Bob" };

            var result = evaluator.Evaluate("Name", person);

            Assert.Equal("Bob", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_PropertyNotAllowed_ThrowsTemplateExpressionNotAllowedException()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson();

            var ex = Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("Age", person));
            Assert.Contains("Age", ex.Message);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_NestedPropertyAllowed_ReturnsValue()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Address.City" }
            };
            var person = new TestPerson { Address = new TestAddress { City = "Lyon" } };

            var result = evaluator.Evaluate("Address.City", person);

            Assert.Equal("Lyon", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_NestedPropertyNotAllowed_ThrowsTemplateExpressionNotAllowedException()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson();

            var ex = Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("Address.City", person));
            Assert.Contains("Address.City", ex.Message);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesEmptySet_BehaviorUnchanged()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string>()
            };
            var person = new TestPerson { Name = "Alice" };

            var result = evaluator.Evaluate("Name", person);

            Assert.Equal("Alice", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesNull_BehaviorUnchanged()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = null
            };
            var person = new TestPerson { Name = "Alice" };

            var result = evaluator.Evaluate("Name", person);

            Assert.Equal("Alice", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_WithFunction_Allowed_ReturnsTransformedValue()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson { Name = "alice" };

            var result = evaluator.Evaluate("Name", person, "Upper");

            Assert.Equal("ALICE", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_WithFunction_NotAllowed_ThrowsTemplateExpressionNotAllowedException()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson();

            var ex = Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("Age", person, "Upper"));
            Assert.Contains("Age", ex.Message);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_WhitespaceExpression_TrimmedAndChecked()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson { Name = "Alice" };

            var result = evaluator.Evaluate("  Name  ", person);

            Assert.Equal("Alice", result);
        }

        [Fact]
        public void Evaluate_AllowedPropertiesSet_WhitespaceExpressionNotAllowed_ThrowsTemplateExpressionNotAllowedException()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Name" }
            };
            var person = new TestPerson();

            var ex = Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("  Age  ", person));
            Assert.Contains("Age", ex.Message);
        }
    }
}
