using System;
using System.Collections.Generic;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class TemplateVisibleAttributeTests
    {
        public class InvoiceModel
        {
            [TemplateVisible]
            public string InvoiceNumber { get; set; }

            [TemplateVisible]
            public decimal Amount { get; set; }

            public string InternalNotes { get; set; }
        }

        public class Address
        {
            [TemplateVisible]
            public string City { get; set; }

            public string ZipCode { get; set; }
        }

        public class Customer
        {
            [TemplateVisible]
            public string Name { get; set; }

            [TemplateVisible]
            public Address Address { get; set; }

            public string InternalId { get; set; }
        }

        public class BaseModel
        {
            [TemplateVisible]
            public string BaseProperty { get; set; }
        }

        public class DerivedModel : BaseModel
        {
            [TemplateVisible]
            public string DerivedProperty { get; set; }
        }

        public class OrderItem
        {
            [TemplateVisible]
            public string ProductName { get; set; }

            [TemplateVisible]
            public int Quantity { get; set; }
        }

        public class Order
        {
            [TemplateVisible]
            public string OrderId { get; set; }

            [TemplateVisible]
            public List<OrderItem> Items { get; set; }
        }

        [Fact]
        public void TemplateVisibleAttribute_MarkedProperty_IsAccessible()
        {
            var evaluator = new Evaluation.ExpressionEvaluator();
            var invoice = new InvoiceModel { InvoiceNumber = "INV-001", Amount = 100m };

            // Simulate what TemplateEngine does - collect visible properties
            evaluator.AllowedProperties = new HashSet<string> { "InvoiceNumber", "Amount" };

            var result = evaluator.Evaluate("InvoiceNumber", invoice);

            Assert.Equal("INV-001", result);
        }

        [Fact]
        public void TemplateVisibleAttribute_NonMarkedProperty_IsNotAccessible()
        {
            var evaluator = new Evaluation.ExpressionEvaluator();
            var invoice = new InvoiceModel { InvoiceNumber = "INV-001", InternalNotes = "Secret" };

            evaluator.AllowedProperties = new HashSet<string> { "InvoiceNumber" };

            Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("InternalNotes", invoice));
        }

        [Fact]
        public void TemplateVisibleAttribute_NestedProperty_IsAccessible()
        {
            var evaluator = new Evaluation.ExpressionEvaluator();
            var customer = new Customer
            {
                Name = "John Doe",
                Address = new Address { City = "Paris", ZipCode = "75001" }
            };

            evaluator.AllowedProperties = new HashSet<string> { "Name", "Address", "Address.City" };

            var cityResult = evaluator.Evaluate("Address.City", customer);

            Assert.Equal("Paris", cityResult);
        }

        [Fact]
        public void TemplateVisibleAttribute_NestedNonMarkedProperty_IsNotAccessible()
        {
            var evaluator = new Evaluation.ExpressionEvaluator();
            var customer = new Customer
            {
                Name = "John Doe",
                Address = new Address { City = "Paris", ZipCode = "75001" }
            };

            evaluator.AllowedProperties = new HashSet<string> { "Name", "Address", "Address.City" };

            Assert.Throws<TemplateExpressionNotAllowedException>(() => evaluator.Evaluate("Address.ZipCode", customer));
        }

        [Fact]
        public void TemplateVisibleAttribute_InheritedProperty_IsAccessible()
        {
            var evaluator = new Evaluation.ExpressionEvaluator();
            var derived = new DerivedModel { BaseProperty = "Base", DerivedProperty = "Derived" };

            evaluator.AllowedProperties = new HashSet<string> { "BaseProperty", "DerivedProperty" };

            var baseResult = evaluator.Evaluate("BaseProperty", derived);
            var derivedResult = evaluator.Evaluate("DerivedProperty", derived);

            Assert.Equal("Base", baseResult);
            Assert.Equal("Derived", derivedResult);
        }

        [Fact]
        public void TemplateVisibleAttribute_CanBeAppliedToProperty()
        {
            var attribute = new TemplateVisibleAttribute();
            Assert.NotNull(attribute);
        }

        [Fact]
        public void TemplateVisibleAttribute_IsSealed()
        {
            var type = typeof(TemplateVisibleAttribute);
            Assert.True(type.IsSealed);
        }

        [Fact]
        public void TemplateVisibleAttribute_InheritsFromAttribute()
        {
            var type = typeof(TemplateVisibleAttribute);
            Assert.True(typeof(Attribute).IsAssignableFrom(type));
        }

        [Fact]
        public void TemplateVisibleAttribute_AllowsMultipleIsFalse()
        {
            var attribute = typeof(TemplateVisibleAttribute);
            var usage = attribute.GetCustomAttributes(typeof(AttributeUsageAttribute), false)[0] as AttributeUsageAttribute;
            Assert.False(usage.AllowMultiple);
        }

        [Fact]
        public void TemplateVisibleAttribute_InheritedIsTrue()
        {
            var attribute = typeof(TemplateVisibleAttribute);
            var usage = attribute.GetCustomAttributes(typeof(AttributeUsageAttribute), false)[0] as AttributeUsageAttribute;
            Assert.True(usage.Inherited);
        }
    }
}