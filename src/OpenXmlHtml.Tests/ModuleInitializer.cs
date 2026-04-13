public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageSharp.Initialize(ssimThreshold: 0.99);
        VerifierSettings.InitializePlugins();
        VerifierSettings.UniqueForRuntime();
        VerifyOpenXmlConverter.Initialize();
    }
}
