using System.IO;

namespace EnaTsApp.Tests.TestHelpers
{
    public static class TestFileHelper
    {
        public static string GetTestFilePath(string relativePath)
        {
            var testDataPath = Path.Combine(
                Path.GetDirectoryName(typeof(TestFileHelper).Assembly.Location)!,
                "..", "..", "..", "..", "..", "src", "EnaTsApp.Tests", "TestData", relativePath);
            return testDataPath;
        }
    }
}
