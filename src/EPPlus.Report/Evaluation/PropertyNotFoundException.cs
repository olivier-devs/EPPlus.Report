using System;

namespace EPPlus.Report.Evaluation
{
    public class PropertyNotFoundException : ArgumentException
    {
        public PropertyNotFoundException(string message) : base(message)
        {
        }
    }
}
