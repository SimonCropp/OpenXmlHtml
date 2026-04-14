public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
#if NET
        VerifyImageSharp.Initialize(ssimThreshold: 0.999);
#endif
        VerifierSettings.InitializePlugins();
        VerifierSettings.UniqueForRuntime();
        VerifyOpenXmlConverter.Initialize();
    }
}
