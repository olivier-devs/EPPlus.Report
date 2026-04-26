using System;
using EPPlus.Report.Evaluation;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class SimpleExpressionTests
    {
        public class TestAddress
        {
            public string City { get; set; } = "Unknown";
        }

        public class TestPerson
        {
            public string Name { get; set; } = "John";
            public TestAddress Address { get; set; } = new TestAddress();
        }

        [Fact]
        public void Evaluate_SimpleProperty_ReturnsValue()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "Alice" };
            
            var result = evaluator.Evaluate("Name", person);
            
            Assert.Equal("Alice", result);
        }

        [Fact]
        public void Evaluate_NestedProperty_ReturnsValue()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson 
            { 
                Name = "Alice",
                Address = new TestAddress { City = "Paris" }
            };
            
            var result = evaluator.Evaluate("Address.City", person);
            
            Assert.Equal("Paris", result);
        }

        [Fact]
        public void Evaluate_NullContext_ThrowsArgumentNullException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentNullException>(() => evaluator.Evaluate("Name", null));
        }

        [Fact]
        public void Evaluate_EmptyExpression_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentException>(() => evaluator.Evaluate("", new TestPerson()));
        }

        [Fact]
        public void Evaluate_PropertyNotFound_ThrowsPropertyNotFoundException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<PropertyNotFoundException>(() => evaluator.Evaluate("NonExistent", new TestPerson()));
        }

        [Fact]
        public void Evaluate_NullIntermediateProperty_ReturnsNull()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "Alice", Address = null };
            
            var result = evaluator.Evaluate("Address.City", person);
            
            Assert.Null(result);
        }

        [Fact]
        public void Evaluate_WhitespaceExpression_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentException>(() => evaluator.Evaluate("   ", new TestPerson()));
        }

        [Fact]
        public void Evaluate_WithUpperFunction_ReturnsUpperCase()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "alice" };
            
            var result = evaluator.Evaluate("Name", person, "Upper");
            
            Assert.Equal("ALICE", result);
        }

        [Fact]
        public void Evaluate_WithLowerFunction_ReturnsLowerCase()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "ALICE" };
            
            var result = evaluator.Evaluate("Name", person, "Lower");
            
            Assert.Equal("alice", result);
        }

        [Fact]
        public void Evaluate_WithTrimFunction_ReturnsTrimmed()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "  alice  " };
            
            var result = evaluator.Evaluate("Name", person, "Trim");
            
            Assert.Equal("alice", result);
        }

        [Fact]
        public void Evaluate_WithEmptyFunctionName_ReturnsRawValue()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "Alice" };
            
            var result = evaluator.Evaluate("Name", person, "");
            
            Assert.Equal("Alice", result);
        }

        [Fact]
        public void Evaluate_WithUnknownFunction_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            var person = new TestPerson { Name = "Alice" };
            
            Assert.Throws<ArgumentException>(() => evaluator.Evaluate("Name", person, "Unknown"));
        }

        [Fact]
        public void RegisterFunction_AddsCustomFunction()
        {
            var evaluator = new ExpressionEvaluator();
            evaluator.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
            var person = new TestPerson { Name = "A" };
            
            var result = evaluator.Evaluate("Name", person, "Double");
            
            Assert.Equal("AA", result);
        }

        [Fact]
        public void RegisterFunction_NullName_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentException>(() => evaluator.RegisterFunction(null, x => x));
        }

        [Fact]
        public void RegisterFunction_EmptyName_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentException>(() => evaluator.RegisterFunction("", x => x));
        }

        [Fact]
        public void RegisterFunction_WhitespaceName_ThrowsArgumentException()
        {
            var evaluator = new ExpressionEvaluator();
            
            Assert.Throws<ArgumentException>(() => evaluator.RegisterFunction("   ", x => x));
        }
    }
}
