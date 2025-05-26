using System.IO;

namespace EnaTsApp.Tests.TestHelpers
{
    public static class TestFileHelper
    {
        public static string GetTestFilePath(string relativePath)
        {
            var solutionDir = Path.GetFullPath(Path.Combine(
                Path.GetDirectoryName(typeof(TestFileHelper).Assembly.Location)!,
                "..", "..", "..", ".."));

            var testDataPath = Path.Combine(solutionDir, "src", "EnaTsApp.Tests", "TestData", relativePath);
            return testDataPath;
        }
    }
}
