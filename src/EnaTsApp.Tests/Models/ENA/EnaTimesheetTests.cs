using System.Collections.Generic;
using Xunit;
using F23.StringSimilarity;
using Com.Ena.Timesheet.Ena;

namespace Com.Ena.Timesheet.Tests.Models.ENA
{
    public class EnaTimesheetTests
    {
        [Fact]
        public void BestMatch_NullInput_ReturnsEmptyString()
        {
            // Arrange
            var clientTaskSet = new HashSet<string> { "valid-task" };

            // Act
            var result = EnaTimesheet.BestMatch(null, clientTaskSet);

            // Assert
            Assert.Empty(result);
        }

        [Fact]
        public void BestMatch_EmptyInput_ReturnsEmptyString()
        {
            // Arrange
            var clientTaskSet = new HashSet<string> { "valid-task" };

            // Act
            var result = EnaTimesheet.BestMatch("", clientTaskSet);

            // Assert
            Assert.Empty(result);
        }

        [Fact]
        public void BestMatch_EmptyClientSet_ReturnsEmptyString()
        {
            // Arrange
            var clientTaskSet = new HashSet<string>();
            var input = "test-task";

            // Act
            var result = EnaTimesheet.BestMatch(input, clientTaskSet);

            // Assert
            Assert.Empty(result);
        }

        [Fact]
        public void BestMatch_ExactMatch_ReturnsSameString()
        {
            // Arrange
            var clientTaskSet = new HashSet<string> { "task1", "task2", "task3" };
            var input = "task2";

            // Act
            var result = EnaTimesheet.BestMatch(input, clientTaskSet);

            // Assert
            Assert.Equal(input, result);
        }

        [Fact]
        public void BestMatch_SimilarMatch_ReturnsBestMatch()
        {
            // Arrange
            var clientTaskSet = new HashSet<string> { 
                "task1-abc", 
                "task2-def", 
                "task3-ghi" 
            };
            var input = "task1-abd"; // Similar to "task1-abc"

            // Act
            var result = EnaTimesheet.BestMatch(input, clientTaskSet);

            // Assert
            Assert.Equal("task1-abc", result);
        }
    }
}
