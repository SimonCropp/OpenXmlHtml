public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
#if NET10_0_OR_GREATER
        VerifyImageSharp.Initialize(ssimThreshold: 0.99);
#endif
        VerifierSettings.InitializePlugins();
        VerifierSettings.UniqueForRuntime();
        VerifyOpenXmlConverter.Initialize();
    }
}
